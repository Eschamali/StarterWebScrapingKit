Attribute VB_Name = "CDPCoreWebSocketHelpers"
'***************************************************************************************************
'             名前付きパイプの仕組みを利用したCDP-WebSocket連携機能を提供します。
'       Excel ←NamedPipe→ PowerShell ←WebSocket→ Chromium といった連携確保まで担います。
' 本来、このツールは`remote-debugging-pipe`特化で構築してる都合上、`remote-debugging-port`の直接統合は堅牢性を落としてしまう問題がありました。
' しかし、`PowerShell × NamedPipe`という仕組みにより、拡張機能として提供することで、コア部分をいじらず、実質対応可能となりました。
' 予め、所定の`PowerShell`を起動しておくことで、下記のようなニッチな目的にも対応できます
' ・AndroidスマートフォンのChromiumブラウザの自動制御
' ・WebView2向けの環境変数`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`に、`--remote-debugging-port=9222`を付与した自動制御
' ・`inspect/#remote-debugging`による今、目の前のブラウザから、自動制御
'
' 【パイプ役割】PowerShell が NamedPipe サーバー（待受け）、Excel がクライアント（接続側）です。
'***************************************************************************************************
Option Explicit
Option Private Module



'***************************************************************************************************
'                                   ■■■ WindowsAPI定義 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long 'プロセス内環境変数用API
Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long



'***************************************************************************************************
'                               ■■■ 変数/定数宣言 ■■■
'***************************************************************************************************
'名前付きパイプ名である接頭辞
Private Const PIPE_Landmark As String = "\\.\pipe\"

'名前付きパイプ関連の設定定数
Private Const GENERIC_READ            As Long = &H80000000  '名前付きパイプの読み取り
Private Const GENERIC_WRITE           As Long = &H40000000  '名前付きパイプへ書き込み
Private Const OPEN_EXISTING           As Long = 3           '開設済み名前付きパイプへの接続
Private Const FILE_ATTRIBUTE_NORMAL   As Long = &H80
Private Const INVALID_HANDLE_VALUE  As LongPtr = -1 '失敗サイン

'デフォルト識別名称
Private Const DefaultName   As String = "ChromiumWebSocket"



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：PowerShell が待受けしている名前付きパイプに Excel から接続します
'---------------------------------------------------------------------------------------------------
'* 注意事項：先に`StartConnectWebSocketForChromium.ps1`を実行して待受けさせてください。待受けが無いと エラー番号:2 を返します
'***************************************************************************************************
Sub ManualSetup()
    '識別名称を設定する
    Dim UseName As String
    With ShSetting01_StartBrowser
        'UseName = .Range(.UseRangeName(2, "Demo_WebSocketViaNamedPipe.FirstStep")).value  '設定セルから、ユーザ名を取得する場合
        UseName = DefaultName   'こちらで用意された`PowerShell`の名称で
    End With

    '名前付きパイプへクライアント接続
    Dim ResultCode As Long: ResultCode = CDPCoreWebSocketHelpers.ConnectNamePipe(UseName)

    'エラーチェック
    Dim ErrorDetail As New WinApiError  'エラーコードから、詳細を取得するやつ
    If ResultCode Then
        MsgBox "名前付きパイプへの接続に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & ErrorDetail.GetMessage(ResultCode, "kernel32"), vbCritical, "ErrorCode: " & ResultCode
    Else
        MsgBox "名前付きパイプへの接続が完了しました。", vbInformation, "Success"
    End If
End Sub

'***************************************************************************************************
'* 機能　　：予めセルに埋め込んだ`StartConnectWebSocketForChromium.ps1`のコードを実行させつつ、接続処理も同時に行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：・セキュリティリスクを伴うため、初回利用時のPowerShellコード配置は、使用者自身で手動でセルに置いてください。
'              必要に応じて、配置したセルに名前を書くと良いでしょう
'
'            ・使用環境によっては、ウイルス誤判定になり、強制終了されます。その場合は、前項の手動起動版で我慢してください。
'***************************************************************************************************
Sub AutoSetup()
    '設定
    Const TargetRangeNameWithPS         As String = "A1"    '`StartWebSocket.ps1`のコードが置いてあるセル名
    Const ConnectChromiumWebSocketWS    As String = ""      '決まった接続名があれば
    Const TIMEOUT                       As Long = 20        '接続までのタイムアウト値(単位：秒)


    ' 1. セルからコードを取得
    Dim psCode As String
    psCode = Sheet1.Range(TargetRangeNameWithPS).value  '新規シート作成直後なら「Sheet1」オブジェクトでアクセス可能
    If Trim(psCode) = "" Then MsgBox "PowerShell コードが埋め込まれてないようです。", vbCritical: Exit Sub

    ' 2. 置換でうまい具合に設定パラメータを適用
    ' 2-1.WebSocketURL
    If ConnectChromiumWebSocketWS <> "" Then psCode = Replace(psCode, "[string]$wsUrl    = """"", "[string]$wsUrl    = """ & ConnectChromiumWebSocketWS & Chr(34), , 1)

    ' 2-2.パイプ名
    Dim UseNamePipe As String
    With ShSetting01_StartBrowser
        UseNamePipe = .Range(.UseRangeName(2, "Demo_WebSocketViaNamedPipe.AutoSetup")).value  '設定セルから、ユーザ名を取得
    End With
    If UseNamePipe <> DefaultName Then psCode = Replace(psCode, "[string]$pipeName = """ & DefaultName & Chr(34), "[string]$pipeName = """ & UseNamePipe & Chr(34), , 1)

    ' 3. そのコードを環境変数に一時登録
    Const EnvironmentName As String = "VBA_PS_CODE_WEBSOCKET"
    SetEnvironmentVariableW StrPtr(EnvironmentName), StrPtr(psCode)

    ' 4. コマンド組み立て。PowerShellへの命令を「環境変数を実行せよ」という極短メッセージにする
    ' -NoProfile: 高速化
    ' -WindowStyle Hidden: ウィンドウを隠す
    ' -EncodedCommand: 特殊文字や改行を安全に渡す
    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ""Invoke-Expression $env:" & EnvironmentName & """"

    ' 5. 非同期実行し、環境変数をクリアさせる
    Shell cmd, vbHide
    SetEnvironmentVariableW StrPtr(EnvironmentName), 0

    ' 6.名前付きパイプへクライアント接続
    Dim ResultCode As Long
    Dim timerStart As Single: timerStart = Timer
    Application.StatusBar = TIMEOUT & "s 以内に接続処理をしてください..."
    Do
        ResultCode = CDPCoreWebSocketHelpers.ConnectNamePipe(UseNamePipe)
        sleep2 500
        DoEvents
    Loop While ResultCode And (Timer - timerStart) <= TIMEOUT
    Application.StatusBar = False

    ' 7.エラーチェック
    Dim ErrorDetail As New WinApiError  'エラーコードから、詳細を取得するやつ
    If ResultCode Then
        MsgBox "名前付きパイプへの接続に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & ErrorDetail.GetMessage(ResultCode, "kernel32"), vbCritical, "ErrorCode: " & ResultCode
    Else
        MsgBox "名前付きパイプへの接続が完了しました。", vbInformation, "Success"
    End If
End Sub



'***************************************************************************************************
'                           ■■■テンプレートプロシージャ ■■■
'***************************************************************************************************
'* 注意事項：・事前に、パイプへ接続（ConnectNamePipe）を済ませること
'            ・専用の`PowerShell`（StartConnectWebSocketForChromium.ps1）が起動中であること
'            ・WebSocket経由の場合は常に`.reattach`始まりとなります
'***************************************************************************************************
Sub WebSocketによる冒険の始まり()
    Dim WebSocketCDP As New CDPContext

    '識別名称を設定する
    Dim UseName As String
    With ShSetting01_StartBrowser
        'UseName = .Range(.UseRangeName(2, "Demo_WebSocketViaNamedPipe.FirstStep")).value  '設定セルから、ユーザ名を取得する場合
        UseName = DefaultName   'こちらで用意された`PowerShell`の名称で
    End With

    '1. まずは、既存のTargetIDに接続できるか？
    If Not WebSocketCDP.reattach(UseName) Then
        '既存のTargetIDが消えちゃったので、別タブへの再接続フェーズへ
        Debug.Print "既存の`targetID`への再接続に失敗。新しいタブか、今開いている直近のタブに再接続して、そこから処理を再開します。"

        '2. 未接続のタブに接続
        'Set WebSocketCDP = WebSocketCDP.InheritanceCDPBrowser.getTab(setMain:=True)
        Set WebSocketCDP = WebSocketCDP.InheritanceCDPBrowser.newTab(setMain:=True)     '新しいタブでもOK
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
    End If

    '↓ 3．再接続できたので、ここから、あなたのイメージをコードに落とし込む ↓
    '例）ページ遷移してみる
    WebSocketCDP.navigate "https://kfk.neocraftstudio.com/en"



    'ブラウザを正常に閉じる
    WebSocketCDP.InheritanceCDPBrowser.quit   '実行と共に、名前付きパイプのハンドルもクリーンします
End Sub



'***************************************************************************************************
'                                 ■■■ 接続のコア部分 ■■■
'***************************************************************************************************
'* 機能　　：PowerShell が作成済みの名前付きパイプにクライアントとして接続し、接続情報等を記録します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：エラーコード　 ※0＝成功
'* 引数　　：UserName       接続名称
'---------------------------------------------------------------------------------------------------
'* 注意事項：・PowerShell 側でまず、名前付きパイプを生成しないと接続に失敗します
'            ・ツール側の仕様上、この`UserName`は設定シートの`ユーザーデータフォルダ名`と共存します。
'            　スクレイピングする際は、この設定値に紐づけられたパイプ値としてやり取りが行われます
'***************************************************************************************************
Private Function ConnectNamePipe(UserName As String) As Long
    Dim hNamePipe As LongPtr

    '1. Excelに記録中のパイプがあったらCloseする
    If deserialize(UserName) Then CloseHandle hNamePipe

    '2. 名前付きパイプのフルパス作成
    Dim UsePipeName As String: UsePipeName = PIPE_Landmark & UserName

    '3. そのフルパスの名前付きパイプへ接続
    hNamePipe = CreateFile(UsePipeName, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    '4. 成功判定
    If hNamePipe = INVALID_HANDLE_VALUE Then ConnectNamePipe = Err.LastDllError: Exit Function

    '5. パイプハンドルを記録するプロシージャへ
    serialize UserName, hNamePipe
End Function

'***************************************************************************************************
'* 機能　　：WebView2のデバッグポートを開く際に使います
'***************************************************************************************************
Sub WebView2のクイックデバッグ切り替え(Optional port As Long = 9222)
    Const EnvironmentName As String = "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"


    If port > 0 Then
        SetEnvironmentVariableW StrPtr(EnvironmentName), StrPtr("--remote-debugging-port=" & port)
        Debug.Print "WebView2のデバッグポートを開けました: " & port
    Else
        SetEnvironmentVariableW StrPtr(EnvironmentName), 0
        Debug.Print "WebView2のデバッグポートを閉じました"
    End If
End Sub



'***************************************************************************************************
'                            ■■■ 設定情報読み書き処理 ■■■
'***************************************************************************************************
'* 機能　　：接続情報をExcelテーブルに残します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   接続名称
'            hNamePipe  `CreateFile`で得たハンドル値
'---------------------------------------------------------------------------------------------------
'* 注意事項：`remote-debugging-pipe`と共存する都合上、一部は仮値にしてます
'***************************************************************************************************
Private Sub serialize(UserName As String, hNamePipe As LongPtr)
    '------------------ 1. パイプ情報の記録準備 ------------------
    '※主要となる情報以外は一旦、一律0とし、必要なデータを`Dictionary`に詰める
    Dim tmp As New Dictionary
    tmp.Add "hStdOutRd", 0
    tmp.Add "hStderrOutRd", 0
    tmp.Add "hStdInWr", 0
    tmp.Add "hCDPOutRd", hNamePipe
    tmp.Add "hCDPInWr", hNamePipe
    tmp.Add "dwProcessId", 0

    'Excelテーブルに、名前付きパイプハンドル情報を記録する
    Set ShSetting01_StartBrowser.serializeToTable(UserName, 1, "CDPCoreWebSocketHelpers.serialize") = tmp


    '------------------ 2. タブ情報の記録準備 ------------------
    '※主要となる情報以外は一旦、一律空欄とし、必要なデータを`Dictionary`に詰める
    tmp.RemoveAll
    tmp.Add "BiDi-context", vbNullString
    tmp.Add "sessionID", vbNullString
    tmp.Add "targetID", vbNullString

    'Excelテーブルに、タブ情報欄を確保する
    Set ShSetting01_StartBrowser.serializeToTable(UserName, 2, "CDPCoreWebSocketHelpers.serialize") = tmp

End Sub

'***************************************************************************************************
'* 機能　　：接続情報をExcelテーブルから取得し、クラス内変数に適用します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：ハンドル値　※0で失敗
'* 引数　　：UserName   接続名称
'---------------------------------------------------------------------------------------------------
'* 詳細説明：リアタッチする際にこれを呼び出します
'* 注意事項：`remote-debugging-pipe`と共存する都合上、一部しか取得しません
'***************************************************************************************************
Private Function deserialize(UserName As String) As Long
    'Excelテーブルから、`UserName`に紐づくパイプハンドル情報を読み込む
    Dim PipeInfo As Dictionary
    Set PipeInfo = ShSetting01_StartBrowser.deserializeFromTable(UserName, 1, "CDPCoreWebSocketHelpers.deserialize")

    '取得に失敗した場合は、接続情報が失ってるので、抜ける
    If PipeInfo Is Nothing Then Exit Function

    '返却
    deserialize = PipeInfo("hCDPOutRd")
End Function
