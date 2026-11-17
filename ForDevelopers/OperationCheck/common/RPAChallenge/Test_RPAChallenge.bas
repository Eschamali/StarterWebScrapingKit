Attribute VB_Name = "Test_RPAChallenge"
'==============================================================================================================
'             かの有名な`RPA Challenge`での動作確認を行います。CDP/WebDriverBiDi 両対応です
'--------------------------------------------------------------------------------------------------------------
'      やり取りを最小限に抑えるため、1回のJavaScript実行に収める形で取ります。事前に答え用のJavaScriptを埋め込んでください
'                                   色んなスタート方法を載せておきます
'==============================================================================================================
Option Explicit



'対象のテストサイト
Const RPAChallengeURL   As String = "https://rpachallenge.com/"
Const JSのあるセル番地  As String = "D43"   '※アクティブシートを対象とします。



'***************************************************************************************************
'                                  ■■■ CDP Challenge ■■■
'***************************************************************************************************
'* 機能　　：新規起動として、ブラウザを立ち上げ、CDPモードとしてチャレンジを開始します
'***************************************************************************************************
Sub CDPForFirstStart()
    '1. 初期URLとして、`RPA Challenge`を起動
    Dim RPAChallenge As CDPContext
    Set RPAChallenge = ShSetting01_StartBrowser.StartCDPModeContext(RPAChallengeURL)

    '2. 事前にセルに埋め込んだJavaScriptを流し込む
    Debug.Print "New CDP-Session: " & RPAChallenge.jsEval(Range(JSのあるセル番地).value)

    '3. 少し待機して終了
    CDPHelpers.Sleep 3
    RPAChallenge.ThisCDPBrowser.CloseBrowser
End Sub

'***************************************************************************************************
'* 機能　　：起動済みのブラウザに対しての、チャレンジです
'---------------------------------------------------------------------------------------------------
'* 注意事項：・起動済みの性質上、WebSocketのみです
'            ・ここでは、固定ポート:9222として扱います
'***************************************************************************************************
Sub CDPForReStart1()
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. CDP制御用のWebSocketへ接続
    Dim CDPws As CDPCoreViaWebSocket: Set CDPws = New CDPCoreViaWebSocket
    CDPws.AutoConnectBrowserCDP UserName
    ReattachOptions.Set_WebSocketMode(CDPws) = UserName

    '3. ブラウザに接続
    Dim br As CDPBrowser: Set br = New CDPBrowser
    br.reattach ReattachOptions

    '4. 新規タブを開く
    Dim RPAChallenge As CDPContext
    Set RPAChallenge = br.newTab(setMain:=True)

    '5．そのタブで、指定のページに遷移
    RPAChallenge.navigate RPAChallengeURL

    '6. 事前にセルに埋め込んだJavaScriptを流し込む
    Debug.Print "Exist CDP-Session: " & RPAChallenge.jsEval(Range(JSのあるセル番地).value)

    '7. 少し待機して終了
    CDPHelpers.Sleep 3
    RPAChallenge.ThisCDPBrowser.CloseBrowser
End Sub

'***************************************************************************************************
'* 機能　　：「DevToolsActivePort」ファイルモードで起動したブラウザに対するチャレンジです
'---------------------------------------------------------------------------------------------------
'* 注意事項：このDemoは、普段のブラウザを制御するDemoのため、「edge://inspect/#remote-debugging」にて事前準備が必要です
'***************************************************************************************************
Sub CDPForReStart2()
    '1. 普段使いのユーザー名を指定
    Dim UserName As String
    UserName = "User Data"

    '2. CDP制御用のWebSocketへ接続
    Dim CDPws As CDPCoreViaWebSocket: Set CDPws = New CDPCoreViaWebSocket
    BrowserCapabilities.Set_UserDataDirName = UserName
    CDPws.AutoConnectDevToolsActivePort BrowserCapabilities
    ReattachOptions.Set_WebSocketMode(CDPws) = UserName

    '3. ブラウザに接続
    Dim br As CDPBrowser: Set br = New CDPBrowser
    br.reattach ReattachOptions

    '4. 新規タブを開く
    Dim RPAChallenge As CDPContext
    Set RPAChallenge = br.newTab(setMain:=True)

    '5．そのタブで、指定のページに遷移
    RPAChallenge.navigate RPAChallengeURL

    '6. 事前にセルに埋め込んだJavaScriptを流し込む
    Debug.Print "Everyday use CDP-Session: " & RPAChallenge.jsEval(Range(JSのあるセル番地).value)

    '7. 切断
    RPAChallenge.ThisCDPBrowser.DisconnectCDP
End Sub

'***************************************************************************************************
'* 機能　　：WebView2に対するチャレンジです
'***************************************************************************************************
Sub CDPForWebView2Start()
    With WebView2Form
        '1. WebView2プロセスを起動
        .StartCDPModeWebView2
        
        '2. モードレスで表示
        .show False

        '3．そのタブで、指定のページに遷移
        .ThisCDPContext.navigate RPAChallengeURL

        '4. 事前にセルに埋め込んだJavaScriptを流し込む
        Debug.Print "WebView2 CDP-Session: " & .ThisCDPContext.jsEval(Range(JSのあるセル番地).value)

        '5. 少し待機して、非表示
        CDPHelpers.Sleep 3
        .hide
    End With

    '6. 後始末
    Unload WebView2Form
End Sub



'***************************************************************************************************
'                               ■■■ WebDriverBiDi Challenge ■■■
'***************************************************************************************************
'* 機能　　：新規起動として、ブラウザを立ち上げ、BiDiモードとしてチャレンジを開始します
'***************************************************************************************************
Sub BiDiForFirstStart()
    '1. 初期URLとして、`RPA Challenge`を起動
    Dim RPAChallenge As WebDriverBiDiContext
    Set RPAChallenge = ShSetting01_StartBrowser.StartBiDiModeContext(RPAChallengeURL)

    '2. 事前にセルに埋め込んだJavaScriptを流し込む
    Debug.Print "New BiDi-Session: " & RPAChallenge.jsEval(Range(JSのあるセル番地).value)

    '3. 少し待機して終了
    CDPHelpers.Sleep 3
    RPAChallenge.ThisWebDriverBiDiMode.CloseBrowser
End Sub

'***************************************************************************************************
'* 機能　　：起動済みのブラウザに対しての、チャレンジです
'---------------------------------------------------------------------------------------------------
'* 注意事項：・起動済みの性質上、WebSocketのみです
'            ・ここでは、固定ポート:9222として扱います
'***************************************************************************************************
Sub BiDiForReStart1()
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. 一旦、CDP制御用のWebSocketへ接続
    Dim CDPws As CDPCoreViaWebSocket: Set CDPws = New CDPCoreViaWebSocket
    CDPws.AutoConnectBrowserCDP UserName
    ReattachOptions.Set_WebSocketMode(CDPws) = UserName

    '3. BiDi用のブラウザに接続
    Dim br As WebDriverBiDiMode: Set br = New WebDriverBiDiMode
    br.reattach ReattachOptions

    '4. 新規タブを開く
    Dim RPAChallenge As WebDriverBiDiContext
    Set RPAChallenge = br.newTab(setMain:=True)

    '5．そのタブで、指定のページに遷移
    RPAChallenge.navigate RPAChallengeURL

    '6. 事前にセルに埋め込んだJavaScriptを流し込む
    Debug.Print "Exist BiDi-Session: " & RPAChallenge.jsEval(Range(JSのあるセル番地).value)

    '7. 少し待機して終了
    CDPHelpers.Sleep 3
    RPAChallenge.ThisWebDriverBiDiMode.CloseBrowser
End Sub

'***************************************************************************************************
'* 機能　　：「DevToolsActivePort」ファイルモードで起動したブラウザに対するチャレンジです
'---------------------------------------------------------------------------------------------------
'* 注意事項：このDemoは、普段のブラウザを制御するDemoのため、「edge://inspect/#remote-debugging」にて事前準備が必要です
'***************************************************************************************************
Sub BiDiForReStart2()
    '1. 普段使いのユーザー名を指定
    Dim UserName As String
    UserName = "User Data"

    '2. 一旦、CDP制御用のWebSocketへ接続
    Dim CDPws As CDPCoreViaWebSocket: Set CDPws = New CDPCoreViaWebSocket
    BrowserCapabilities.Set_UserDataDirName = UserName
    CDPws.AutoConnectDevToolsActivePort BrowserCapabilities
    ReattachOptions.Set_WebSocketMode(CDPws) = UserName

    '3. ブラウザに接続
    Dim br As WebDriverBiDiMode: Set br = New WebDriverBiDiMode
    br.reattach ReattachOptions

    '4. 新規タブを開く
    Dim RPAChallenge As WebDriverBiDiContext
    Set RPAChallenge = br.newTab(setMain:=True)

    '5．そのタブで、指定のページに遷移
    RPAChallenge.navigate RPAChallengeURL

    '6. 事前にセルに埋め込んだJavaScriptを流し込む
    Debug.Print "Everyday use BiDi-Session: " & RPAChallenge.jsEval(Range(JSのあるセル番地).value)

    '7. 切断
    RPAChallenge.ThisWebDriverBiDiMode.sessionEnd
End Sub

'***************************************************************************************************
'* 機能　　：起動済みFirefoxブラウザに対するチャレンジです
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に、ex-firefoxブランチの変更を反映する必要があります
'***************************************************************************************************
Sub BiDiForFirefoxStart()
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. Firefox用のWebDriverBiDiServerに接続
    Dim BiDiWS As CDPCoreViaWebSocket: Set BiDiWS = New CDPCoreViaWebSocket
    BiDiWS.ConnectCDP UserName, "/session", 59035

    '3. WebDriverBiDi、直やり取りモードとしてセット
    ReattachOptions.Set_DirectWebDriverBiDi = True
    ReattachOptions.Set_WebSocketMode(BiDiWS) = UserName

    '4. Firefoxブラウザに接続
    Dim br As WebDriverBiDiMode: Set br = New WebDriverBiDiMode
    br.reattach ReattachOptions

    '5. 新規タブを開く
    Dim RPAChallenge As WebDriverBiDiContext
    Set RPAChallenge = br.newTab(setMain:=True)

    '6．そのタブで、指定のページに遷移
    RPAChallenge.navigate RPAChallengeURL

    '7. 事前にセルに埋め込んだJavaScriptを流し込む
    Debug.Print "Firefox BiDi-Session: " & RPAChallenge.jsEval(Range(JSのあるセル番地).value)

    '8. 切断
    RPAChallenge.ThisWebDriverBiDiMode.sessionEnd
End Sub
