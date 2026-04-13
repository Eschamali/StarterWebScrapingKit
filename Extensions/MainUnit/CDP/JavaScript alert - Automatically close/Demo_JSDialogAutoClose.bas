Attribute VB_Name = "Demo_JSDialogAutoClose"
'==============================================================================
' Demo - exCDP_JSDialogAutoClose（JavaScript ダイアログ自動クローズ）
'
' 前提: 本モジュールとexCDP_JSDialogAutoClose.cls を VBA プロジェクトに取り込み、
'       Demo_CDP の 設定シートからのCDP起動 が使えること。
'==============================================================================
Option Explicit



Sub TestAlertWithExpansion()
    '設定シートに基づくブラウザ立ち上げ。`selenium`の独自テストページに遷移します
    Dim Demo_alerts As CDPBrowser: Set Demo_alerts = 設定シートからのCDP起動("https://www.selenium.dev/selenium/web/alerts.html")

    '拡張機能を追加（prompt の入力値は 入力文字内容 と同一にし、Debug.Assert と整合させる）
    Dim testEX As New exCDP_JSDialogAutoClose
    testEX.Init Demo_alerts

    '必要な変数を用意
    Dim paramsCDP As New Scripting.Dictionary
    Dim resCDP As Scripting.Dictionary
    Dim searchId As String
    Dim nodeId As Long
    Dim x As Double, y As Double
    
    'テキスト入力用のAlertに入力させる文字列の指定（拡張の DefaultPromptText と同じ値）
    Dim 入力文字内容 As String
    入力文字内容 = "VBAから入力したテスト文字列です！" & WorksheetFunction.Unichar(129418)

    testEX.DefaultAccept = True
    testEX.DefaultPromptText = 入力文字内容

    With Demo_alerts
        ' --- 1. 必要なドメインを有効化 ---
        .invokeMethod ("DOM.enable")
        

        ' --- 2. DOMツリーを同期させ、ID割り振りを行う ---
        paramsCDP.RemoveAll
        paramsCDP.Add "depth", 0        '返却時のDOM情報は不要なので、0にしておく
        paramsCDP.Add "pierce", True    'Shadow DOMの中まで貫通させる
        .invokeMethod "DOM.getDocument", paramsCDP
        ' これでブラウザ内の全ノードにIDが割り振られます

        Dim i As Long
        For i = 1 To 3
            Dim TargetXpath As String
            Select Case i
                Case 1: TargetXpath = "//*[@id='alert']"
                Case 2: TargetXpath = "//*[@id='empty-alert']"
                Case 3: TargetXpath = "//*[@id='prompt']"
            End Select

            ' --- 3. XPathで検索 (Shadow DOMの貫通も可) ---
            paramsCDP.RemoveAll
            paramsCDP.Add "query", TargetXpath  '先頭のリンクを対象に
            Set resCDP = .invokeMethod("DOM.performSearch", paramsCDP)
            searchId = resCDP("searchId")
    
    
            ' --- 4. nodeIdを取得 ---
            paramsCDP.RemoveAll
            paramsCDP.Add "searchId", searchId
            paramsCDP.Add "fromIndex", 0   '先頭の件数から
            paramsCDP.Add "toIndex", 1     '1件分のみ
            Set resCDP = .invokeMethod("DOM.getSearchResults", paramsCDP)
            nodeId = resCDP("nodeIds")(1)  '配列の先頭を取得
    
    
            ' --- 5. nodeId を objectId に変換 ---
            paramsCDP.RemoveAll
            paramsCDP.Add "nodeId", nodeId
            Set resCDP = .invokeMethod("DOM.resolveNode", paramsCDP)


            ' --- 6. あえて、同期でコマンド実行(Jsのクリック処理) ---
            'この瞬間、JavaScriptの`alert`関数が発動されますが、先頭に記述した拡張機能によるイベントキャッチで、
            '同期モードにもかかわらず、JavaScriptアラートが自動で閉じられ、処理が続行されます
            .jsEval "function() { this.click(); }", CStr(resCDP("object")("objectId"))
        Next


        ' --- 7. ブラウザを閉じる ---
        Dim Htmlの表示内容 As String: Htmlの表示内容 = .getElementByXPath("//*[@id='text']/p").innerText
        Debug.Print "htmlの出力文字列：" & Htmlの表示内容
        Debug.Assert Htmlの表示内容 = 入力文字内容

        Set testEX = Nothing      '拡張機能をOFF
        .quit
    End With
End Sub
