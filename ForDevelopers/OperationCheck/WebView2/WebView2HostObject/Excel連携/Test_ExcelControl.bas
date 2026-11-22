Attribute VB_Name = "Test_ExcelControl"
Option Explicit



Sub Excel連携Demo()
    Const ProjectFolderPath As String = ""
    
    With WebView2Form
        '1. WebView2プロセスを起動
        .StartCDPModeWebView2

        '2. 遷移前の事前設定を施す
        .ThisWebView2.DevToolsEnabled = False       'DevToolsウィンドウ起動禁止
        .ThisWebView2.ContextMenuEnabled = False    '右クリック禁止
        .ThisWebView2.StatusBarEnabled = False      '左下に出るステータステキストを表示しない
        
        '3. ★ブリッジクラスを生成して、"vba" という名前で WebView2 に公開！
        Dim bridge As clsWebBridge: Set bridge = New clsWebBridge
        .ThisWebView2.AddHostObjectToScript "vba", bridge

        '4. 指定のhtmlファイルを読み込む
        Dim htmlStr As New CharacterCodeConversion
        Dim params As New Dictionary
        params.Add "frameId", .ThisCDPContext.CurrentTargetID
        params.Add "html", htmlStr.BytesToString(htmlStr.BytesFromSavedFile(ProjectFolderPath & "\ForDevelopers\OperationCheck\WebView2\WebView2HostObject\Excel連携", "UserFormWebView2UI.html"))
        .ThisCDPContext.ExecuteCDP "Page.setDocumentContent", params

        '5. フォームを表示
        .show vbModeless
    End With
End Sub
