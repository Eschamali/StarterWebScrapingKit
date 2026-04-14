Attribute VB_Name = "Demo_LocalAI"
Option Explicit





Sub AIによる冒険の始まり()
    '設定シートに基づくブラウザ立ち上げ
    Dim AI As CDPBrowser: Set AI = 設定シートからのCDP起動

    '↓ここから、あなたのイメージをコードに落とし込む↓
    AI.navigate "edge://version"



    Dim testAi As New exCDP_LocalAI
    Dim AIObjectID As String
    testAi.Init AI
    AIObjectID = testAi.createSession


    Debug.Print testAi.runAI(AIObjectID)
    Stop

    'ブラウザを正常に閉じる
    AI.quit
End Sub

Dim objectId
