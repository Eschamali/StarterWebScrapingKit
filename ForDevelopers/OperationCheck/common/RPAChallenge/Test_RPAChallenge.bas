Attribute VB_Name = "Test_RPAChallenge"
'==============================================================================================================
'                           かの有名な`RPA Challenge`の動作確認を行います
'--------------------------------------------------------------------------------------------------------------
'               やり取りを最小限に抑えるため、1回のJavaScript実行に収める形で取ります
'                       なお、ブラウザは既に起動中の状態であるとします。
'==============================================================================================================
Option Explicit



'対象のテストサイト
Const RPAChallengeURL   As String = "https://rpachallenge.com/"
Const JSのあるセル番地  As String = "D43"


Sub ForCDP()
    Dim RPAChallenge As New CDPBrowser

    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    With ShSetting01_StartBrowser
        UserName = .Range(.UseRangeName(2, "Test_RPAChallenge.ForCDP")).value
    End With

    '2. まずは、既存のTargetIDに接続できるか？
    If Not RPAChallenge.reattach(UserName) Then
        '既存のTargetIDが消えちゃったので、別タブへの再接続フェーズへ
        Debug.Print "既存の`targetID`への再接続に失敗。新しいタブに再接続して、そこから処理を再開します。"

        '2-1. タブを作成してそこを確立
        RPAChallenge.newTab setMain:=True
        'RPAChallenge.getTab setMain:=True
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
    End If

    '3．再接続できたので、指定のページに遷移
    RPAChallenge.navigate RPAChallengeURL

    '4. Run!
    Debug.Print RPAChallenge.jsEval(ShSetting01_StartBrowser.Range(JSのあるセル番地).value)

End Sub

Sub ForBiDi()
    '1. 設定セルから、ユーザ名を取得
    With ShSetting01_StartBrowser
        Dim UserName As String
        UserName = .Range(.UseRangeName(2, "Test_RPAChallenge.ForBiDi")).value
    End With

    '2. リアタッチとして起動
    Dim RPAChallenge As New WebDriverBiDiCore
    Dim ResultReattach As Boolean
    ResultReattach = RPAChallenge.reattach(UserName, ReBoot:=False) '`BiDi-CDP Mapper`タブを閉じちゃった場合は、`ReBoot:=True`にしてください

    If Not (ResultReattach) Then Debug.Print "Failed to reattach. `demoReattachmentPart1`を始動しましたか？": Exit Sub

    '3. 現在のコンテキストIDを取得する
    Dim paramsBiDi As Dictionary, resultBiDi As Dictionary
    Set resultBiDi = RPAChallenge.invokeMethod("browsingContext.getTree")
    Dim targetContext As String
    If Not (resultBiDi Is Nothing) Then
        '※ここでエラーが起こる場合、ブラウザのタブを何個か開いてみて下さい。大抵は、2,3個程度追加で開けば、行けると思います。
        targetContext = resultBiDi("contexts")(1)("context")     '一旦は、先頭タブで　※本来はURLcheckとかがいると思うが、低レベル制御の都合上、妥協
    End If

    '4. 指定のページに遷移
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "context", targetContext
    paramsBiDi.Add "url", RPAChallengeURL
    paramsBiDi.Add "wait", "complete"
    RPAChallenge.invokeMethod "browsingContext.navigate", paramsBiDi

    '5. Run!
    Dim paramsBiDi_target As New Dictionary
    paramsBiDi_target.Add "context", targetContext

    paramsBiDi.RemoveAll
    paramsBiDi.Add "expression", ShSetting01_StartBrowser.Range(JSのあるセル番地).value
    paramsBiDi.Add "target", paramsBiDi_target
    paramsBiDi.Add "awaitPromise", True

    Dim JSConv As New WebJsonConverter
    Debug.Print JSConv.ConvertToJson(RPAChallenge.invokeMethod("script.evaluate", paramsBiDi))

End Sub
