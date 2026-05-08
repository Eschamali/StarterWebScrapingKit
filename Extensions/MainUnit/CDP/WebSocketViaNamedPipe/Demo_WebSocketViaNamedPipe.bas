Attribute VB_Name = "Demo_WebSocketViaNamedPipe"
'***************************************************************************************************
'               名前付きパイプの仕組みを利用したWebSocket連携機能のDemoコードです
'       Excel ←NamedPipe→ PowerShell ←WebSocket→ Chromium といった連携を前提とします
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI定義 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long 'プロセス内環境変数用API
Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)



'***************************************************************************************************
'                                   ■■■ Demo用に定義 ■■■
'***************************************************************************************************
Private Const DefaultName   As String = "ChromiumWebSocket" 'デフォルト識別名称



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
    Dim WebSocketMode As New WebSocketViaNamedPipe
    Dim ResultCode As Long: ResultCode = WebSocketMode.ConnectNamePipe(UseName)

    'エラーチェック
    Dim ErrorDetail As New WinApiError  'エラーコードから、詳細を取得するやつ
    If ResultCode Then
        MsgBox "名前付きパイプへの接続に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & ErrorDetail.GetMessage(ResultCode, "kernel32"), vbCritical, "ErrorCode: " & ResultCode
    Else
        MsgBox "名前付きパイプへの接続が完了しました。", vbInformation, "Success"
    End If
End Sub

'***************************************************************************************************
'* 機能　　：予めセルに埋め込んだ`StartWebSocket.ps1`のコードを実行させつつ、接続処理も同時に行います
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
    Const Timeout                       As Long = 20        '接続までのタイムアウト値(単位：秒)


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
    Dim WebSocketMode As New WebSocketViaNamedPipe
    Dim ResultCode As Long
    Dim timerStart As Single: timerStart = Timer
    Application.StatusBar = Timeout & "s 以内に接続処理をしてください..."
    Do
        ResultCode = WebSocketMode.ConnectNamePipe(UseNamePipe)
        sleep2 500
        DoEvents
    Loop While ResultCode And (Timer - timerStart) <= Timeout
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
    Dim WebSocketCDP As New CDPBrowser

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
        WebSocketCDP.getTab setMain:=True
        'WebSocketCDP.newTab setMain:=True     '新しいタブでもOK
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
    End If

    '↓ 3．再接続できたので、ここから、あなたのイメージをコードに落とし込む ↓
    '例）ページ遷移してみる




    'ブラウザを正常に閉じる
    WebSocketCDP.quit   '実行と共に、名前付きパイプのハンドルもクリーンします
End Sub
