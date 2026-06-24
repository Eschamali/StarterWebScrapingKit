Attribute VB_Name = "Demo_WebSocketViaAnonymousPipe"
'***************************************************************************************************
'               匿名パイプの仕組みを利用したCDP-WebSocket連携機能のDemoコードです
'       Excel ←匿名Pipe→ PowerShell ←WebSocket→ Chromium といった連携を前提とします
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI定義 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：予めセルに埋め込んだ`StartConnectWebSocketForChromium.ps1`のコードを実行させつつ、接続処理も同時に行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：・セキュリティリスクを伴うため、初回利用時のPowerShellコード配置は、使用者自身で手動でセルに置いてください。
'              必要に応じて、配置したセルに名前を書くと良いでしょう
'
'            ・標準入力によるPowerShellコード注入により、従来の`Shell`起動型よりはウイルスセキュリティ誤検知しづらい仕組みですが万が一、
'            　誤検知された場合は、従来の名前付きパイプ版にて手動起動でお願いします
'***************************************************************************************************
Sub AutoSetup()
    '設定
    Const TargetRangeNameWithPS         As String = "A1"    '`StartWebSocket.ps1`のコードが置いてあるセル名
    Const ConnectChromiumWebSocketWS    As String = ""      '決まった接続wsがあれば指定
    Const ShowConsoleWindow             As Boolean = False  'デバック時は`True`にすることで、標準出力/エラーをコンソール画面で確認できます


    ' 1. セルからコードを取得
    ' 「& { ... };exit 0」といった記法をすることで、一括処理を行いつつ、完走時のプロセス残存しないようにします
    Dim psCode As String
    psCode = "& {" & Sheet1.Range(TargetRangeNameWithPS).value & "};exit 0"    '新規シート作成直後なら「Sheet1」オブジェクトでアクセス可能
    If Trim(psCode) = "" Then MsgBox "PowerShell コードが埋め込まれてないようです。", vbCritical: Exit Sub

    ' ----------------------- 2. 環境変数でうまい具合に設定パラメータを継承できるようにします -----------------------
    Dim RunPowerShell As New PowerShellViaStdPipe

    ' 2-1. WebSocketURL
    RunPowerShell.RegisterAndSetEnv("fromVBA_wsUrl") = ConnectChromiumWebSocketWS

    ' 2-2. 接続先checkのポート指定
    RunPowerShell.RegisterAndSetEnv("fromVBA_port") = 9222

    ' 2-3. CDP制御用ポートハンドルを継承させる
    RunPowerShell.RegisterAndSetEnv("fromVBA_hCDPOutWr") = RunPowerShell.CreatePipeForCDPOutWr
    RunPowerShell.RegisterAndSetEnv("fromVBA_hCDPInRd") = RunPowerShell.CreatePipeForCDPInRd
    ' -----------------------------------------------------------------------------------------

    ' ----------------------- 3. Excel内で使われてるWebView2 を起動させる -----------------------
    WebView2のクイックデバッグ切り替え
    Application.CommandBars.ExecuteMso "Help"
    ' -----------------------------------------------------------------------------------------

    ' 4. PowerShell を起動
    RunPowerShell.UseStdOut = False                     'パイプバッファオーバー対策により、VBAへの標準出力を無効化
    Dim ResultCode As Long
    If ShowConsoleWindow Then
        RunPowerShell.UseStdOuterr = False              'console上でエラーを見れるように、VBAへの標準出力エラーを無効化
        ResultCode = RunPowerShell.Init(asNormal, 0)    '標準出力もconsole上も見れます
    Else
        ResultCode = RunPowerShell.Init(asNormal, CREATE_NO_WINDOW)
    End If

    If ResultCode Then
        Dim ErrorDetail As New WinApiError
        MsgBox "PowerShell の起動に失敗しました。" & vbCrLf & vbCrLf & "<原因>" & vbCrLf & ErrorDetail.GetMessage(ResultCode, "kernel32"), vbCritical, "ErrorCode: " & ResultCode
        Exit Sub
    End If

    '5. 一括挿入で起動させる準備を行う
    If ShowConsoleWindow Then
        'コンソールで表示させるのでせっかくなので、UTF-8で見せるように工夫する
        Dim CharConv As New CharacterCodeConversion
        Dim utf8Bytes() As Byte
        utf8Bytes = CharConv.BytesFromString(psCode)

        Dim b64conv As New WebCrypto
        Dim b64str As String: b64str = b64conv.Encode(utf8Bytes, edfBase64, efNoFolding)

        RunPowerShell.writeProcSTD "chcp 65001" & vbCrLf
        RunPowerShell.writeProcSTD "$c = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('" & b64str & "'));" & _
                  "Invoke-Expression $c" & vbCrLf
    Else
        '非表示なので、そのままで実行させる
        RunPowerShell.writeProcSTD psCode & vbCrLf & vbCrLf
    End If

    '6. 指定の識別名称で、接続パイプハンドル情報を記録
    Dim UseName As String
    With ShSetting01_StartBrowser
        UseName = .Range(.UseRangeName(2, "Demo_WebSocketViaAnonymousPipe.AutoSetup")).value  '設定セルから、ユーザ名を取得します
    End With

    RunPowerShell.serialize UseName

    '7. 後始末
    RunPowerShell.CleanAllRegisteredEnv
    MsgBox "CDP通信PowerShellとの連携準備が完了しました", vbInformation
End Sub

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
'                           ■■■テンプレートプロシージャ ■■■
'***************************************************************************************************
'* 注意事項：・事前に、`AutoSetup`を済ませること
'            ・WebSocket経由の場合は常に`.reattach`始まりとなります
'***************************************************************************************************
Sub WebSocketによる冒険の始まり()
    Dim WebSocketCDP As New CDPContext

    '識別名称を設定する
    Dim UseName As String
    With ShSetting01_StartBrowser
        UseName = .Range(.UseRangeName(2, "Demo_WebSocketViaNamedPipe.FirstStep")).value  '設定セルから、ユーザ名を取得します
    End With

    '1. まずは、既存のTargetIDに接続できるか？
    If Not WebSocketCDP.reattach(UseName) Then
        '既存のTargetIDが消えちゃったので、別タブへの再接続フェーズへ
        Debug.Print "既存の`targetID`への再接続に失敗。新しいタブか、今開いている直近のタブに再接続して、そこから処理を再開します。"

        '2. 未接続のタブに接続
        'Set WebSocketCDP = WebSocketCDP.InheritanceCDPBrowser.getTab(setMain:=True)
        Set WebSocketCDP = WebSocketCDP.InheritanceCDPBrowser.newTab(setMain:=True)  '新しいタブでもOK
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
    End If

    '↓ 3．再接続できたので、ここから、あなたのイメージをコードに落とし込む ↓
    '例）ページ遷移してみる




    'ブラウザを正常に閉じる
    WebSocketCDP.InheritanceCDPBrowser.quit   '実行と共に、名前付きパイプのハンドルもクリーンします
End Sub
