Attribute VB_Name = "Demo_WebView2"
'***************************************************************************************************
'   WebView2(CDPCoreViaWebView2)関連のデモ集です
'***************************************************************************************************
Option Explicit
Option Private Module



'***************************************************************************************************
'* 機能　　：WebView2を起動します。基本的な呼び出し型です
'---------------------------------------------------------------------------------------------------
'* 注意事項：`ICoreWebView2Settings`等の一部設定は、ページ遷移前のみ有効です
'***************************************************************************************************
Sub ExcelのユーザーフォームにWebView2を埋め込む()
    '1. UserForm側のWebView2の初期化を済ませる
    With WebView2Form
        If Not .StartCDPModeWebView2 Then Debug.Print "WebView2の初期化に失敗しました。WebView2Loader.dllが見つからない、" & _
                                                        "またはEnvironment/Controllerの生成に失敗した可能性があります。": Exit Sub

        '2. 事前設定を施す(任意)
        .ThisWebView2.DevToolsEnabled = False
        .ThisWebView2.ContextMenuEnabled = False

        '3. ページ遷移
        .ThisCDPContext.navigate "https://github.com/Eschamali/StarterWebScrapingKit"

        '4. フォームを表示
        '※UserFormを閉じるまで、ブロッキングされます
        .show
    End With
End Sub



'***************************************************************************************************
'                                ■■■ 拡張機能デモ ■■■
'***************************************************************************************************
'* 機能　　：WebView2にて拡張機能をインストール/アンインストールする際のDemoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：拡張機能系のみ、CDPコマンドでは出来ないため、WebView2側で用意されてる拡張機能用APIでやるためのDemoとなります
'***************************************************************************************************
Sub 拡張機能インストールアンインストール()
    Dim インストールパス As String
    With Application.FileDialog(4)  'msoFileDialogFolderPicker
        .Title = "拡張機能の基となる`manifest.json`を含むフォルダを選択してください"
        .InitialFileName = Environ("LOCALAPPDATA")    '初期位置

        If .show = -1 Then インストールパス = .SelectedItems(1) Else Exit Sub
    End With


    With WebView2Form
        '1. 拡張機能を有効にする
        .ThisWebView2.EnvironmentOptions.AreBrowserExtensionsEnabled = True

        '2. WebView2を起動
        If Not .StartCDPModeWebView2 Then Debug.Print "WebView2の初期化に失敗しました。WebView2Loader.dllが見つからない、" & _
                                                        "またはEnvironment/Controllerの生成に失敗した可能性があります。": Exit Sub

        '3. ページ遷移
        .ThisCDPContext.navigate "https://github.com/Eschamali/StarterWebScrapingKit"

        '4. フォームを表示
        .show vbModeless

        '5. 拡張機能インストール
        Dim InstallID As String
        InstallID = .ThisWebView2.AddBrowserExtension(インストールパス)
        If LenB(InstallID) = 0 Then MsgBox "拡張機能のインストールに失敗しました", vbCritical, "WebView2": Unload WebView2Form: Exit Sub
        MsgBox "拡張機能のインストールに成功しました。OKを押すとアンインストールします", vbInformation, "exID: " & InstallID
        
        '6. アンインストール
        If Not .ThisWebView2.RemoveBrowserExtension(InstallID) Then MsgBox "拡張機能のアンインストールに失敗しました", vbCritical, "WebView2": Unload WebView2Form: Exit Sub
        MsgBox "拡張機能のアンインストールに成功しました。", vbInformation
    End With

    '7. Demo終了
    Unload WebView2Form
End Sub
