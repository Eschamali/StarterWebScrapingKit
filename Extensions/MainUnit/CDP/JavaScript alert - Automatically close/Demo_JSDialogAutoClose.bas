Attribute VB_Name = "Demo_JSDialogAutoClose"
'==============================================================================
' Demo - exCDP_JSDialogAutoClose（JavaScript ダイアログ自動クローズ）
'
' 前提: 本モジュールとexCDP_JSDialogAutoClose.cls を VBA プロジェクトに取り込み、
'       Demo_CDP の StartCDPModeContext が使えること。
'==============================================================================
Option Explicit



Sub TestAlertWithExpansion()
    '1. 設定シートに基づくブラウザ立ち上げ。`selenium`の独自テストページに遷移します
    Dim Demo_alerts As CDPContext: Set Demo_alerts = ShSetting01_StartBrowser.StartCDPModeContext("https://www.selenium.dev/selenium/web/alerts.html")

    '2. 拡張機能を追加（prompt の入力値は 入力文字内容 と同一にし、Debug.Assert と整合させる）
    Dim testEX As New exCDP_JSDialogAutoClose
    testEX.Init Demo_alerts

    '3. テキスト入力用のAlertに入力させる文字列の指定（拡張の DefaultPromptText と同じ値）
    Dim 入力文字内容 As String
    入力文字内容 = "VBAから入力したテスト文字列です！" & WorksheetFunction.Unichar(129418)
    testEX.DefaultAccept = True
    testEX.DefaultPromptText = 入力文字内容

    '4. 3つの要素に対してクリックして、JSアラート発動
    With Demo_alerts
        Dim i As Long
        For i = 1 To 3
            Dim TargetID As String
            Select Case i
                Case 1: TargetID = "alert"
                Case 2: TargetID = "empty-alert"
                Case 3: TargetID = "prompt"
            End Select

            '5. ID要素を特定し、clickするだけ！
            Demo_alerts.getElementByID(TargetID).SimpleClick

            '6. Openログを見る
            Debug.Print testEX.ViewLastJavascriptDialogOpening
        Next

        '7. 一致チェック
        Dim Htmlの表示内容 As String: Htmlの表示内容 = .getElementByXPath("//*[@id='text']/p").innerText
        Debug.Print "htmlの出力文字列：" & Htmlの表示内容
        Debug.Assert Htmlの表示内容 = 入力文字内容

        ' --- 8. ブラウザを閉じる ---
        Set testEX = Nothing      '拡張機能をOFF
        .ThisCDPBrowser.quit
    End With
End Sub
