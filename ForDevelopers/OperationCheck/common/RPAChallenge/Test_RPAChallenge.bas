Attribute VB_Name = "Test_RPAChallenge"
'==============================================================================================================
'                         かの有名な`RPA Challenge`での動作確認を行います
'--------------------------------------------------------------------------------------------------------------
'               やり取りを最小限に抑えるため、1回のJavaScript実行に収める形で取ります
'                       なお、ブラウザは既に起動中の状態であるとします。
'==============================================================================================================
Option Explicit



'対象のテストサイト
Const RPAChallengeURL   As String = "https://rpachallenge.com/"
Const JSのあるセル番地  As String = "D43"


Sub ForCDP(hoge As CDPCoreViaWebSocket)
    Dim RPAChallenge As New CDPContext

    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Test_RPAChallenge.ForCDP")

    '2. まずは、既存のTargetIDに接続できるか？
    If Not RPAChallenge.reattach(UserName, WebSocketMode:=hoge) Then
        '既存のTargetIDが消えちゃったので、別タブへの再接続フェーズへ
        Debug.Print "既存の`targetID`への再接続に失敗。新しいタブに再接続して、そこから処理を再開します。"

        '2-1. 新しいタブを作成してそこを確立
        Set RPAChallenge = RPAChallenge.InheritanceCDPBrowser.newTab(setMain:=True)
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
    End If

    '3．再接続できたので、指定のページに遷移
    RPAChallenge.navigate RPAChallengeURL

    '4. Run!
    Debug.Print RPAChallenge.jsEval(ShSetting01_StartBrowser.Range(JSのあるセル番地).value)

End Sub

Sub ForBiDi(hoge As CDPCoreViaWebSocket)
    Dim RPAChallenge As New WebDriverBiDiContext

    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Test_RPAChallenge.ForCDP")

    '2. まずは、既存のContextIDに接続できるか？
    If Not RPAChallenge.reattach(UserName, WebSocketMode:=hoge) Then
        '既存のTargetIDが消えちゃったので、別タブへの再接続フェーズへ
        Debug.Print "既存の`targetID`への再接続に失敗。新しいタブに再接続して、そこから処理を再開します。"

        '2-1. 新しいタブを作成してそこを確立
        Set RPAChallenge = RPAChallenge.InheritanceWebDriverBiDiMode.newTab(setMain:=True)
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
    End If

    '3. 指定のページに遷移
    RPAChallenge.navigate RPAChallengeURL

    '4. Run!
    Dim paramsBiDi_target As New Dictionary, paramsBiDi As New Dictionary
    paramsBiDi_target.Add "context", RPAChallenge.context

    paramsBiDi.Add "expression", ShSetting01_StartBrowser.Range(JSのあるセル番地).value
    paramsBiDi.Add "target", paramsBiDi_target
    paramsBiDi.Add "awaitPromise", True

    RPAChallenge.InheritanceWebDriverBiDiMode.ExecuteBiDi "script.evaluate", paramsBiDi

End Sub
