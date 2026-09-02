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
