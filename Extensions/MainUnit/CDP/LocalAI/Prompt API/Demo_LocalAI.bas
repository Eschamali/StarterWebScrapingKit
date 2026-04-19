Attribute VB_Name = "Demo_LocalAI"
'===================================================================================================
'                                       Prompt API Demo
'---------------------------------------------------------------------------------------------------
' URL   :
'       https://learn.microsoft.com/ja-jp/microsoft-edge/web-platform/prompt-api
'       https://developer.chrome.com/docs/ai/prompt-api?hl=ja
' Notes :
'       Prompt API が JavaScriptで実行できる仕様を利用したブラウザ内ローカルAIのやり取りDemoです
'       現状、このAPIが使えるURLに制限があるため、`edge://version`といった固有ページでやり取りをします
'===================================================================================================
Option Explicit



Private Const ThisClassName As String = "Demo_LocalAI"



'***************************************************************************************************
'                               ■■■ Prompt API - 準備編 ■■■
'***************************************************************************************************
'* 機能　　：AIモデルデータのDL処理を行います
'---------------------------------------------------------------------------------------------------
'* 詳細説明：起動から、AIモデルデータの保存まで担います
'* 注意事項：既にAIモデルデータを保存中であっても、このプロシージャから実行してください。
'            以降のプロシージャは、リアタッチから始めるためです
'***************************************************************************************************
Sub PromptAPIの準備()
    Const FromProcedureName As String = ThisClassName & ".PromptAPIの準備"


    '設定シートに基づくブラウザ立ち上げ
    Dim ReadyAI As CDPBrowser: Set ReadyAI = 設定シートからのCDP起動

    '`PromptAPI`は、動作URL場所が限られるため、ブラウザ固有の専用ページに遷移させる
    '※オフラインでも遷移/利用できる「バージョン情報」にひとまず設定
    ReadyAI.navigate "edge://version"

    '拡張機能クラスに継承させる
    Dim PromptAPI  As New LocalAI_PromptAPI
    PromptAPI.Init ReadyAI

    '1. API が有効かどうかを確認
    If PromptAPI.IsPromptApiAvailable Then
        ReadyAI.printMsg info_, "お使いの環境では、Prompt APIが利用可能です！", FromProcedureName
    Else
        '「https://learn.microsoft.com/ja-jp/microsoft-edge/web-platform/prompt-api」を参考に、有効化してください
        ReadyAI.printMsg WARN_, "お使いの環境では、Prompt APIは、利用できません。バージョン自体が非対応か、専用のFLAGSがEnableになってません", FromProcedureName
        MsgBox "お使いの環境では、Prompt APIは、利用できません。バージョン自体が非対応か、専用のFLAGSがEnableになってません", vbCritical
        ReadyAI.quit
        Exit Sub
    End If

    '2. モデルの使用状況をcheck
    Dim ModeAvailability As String: ModeAvailability = PromptAPI.CheckAvailability
    Debug.Print ModeAvailability

    '3. 状況に応じた分岐
    Dim Continue As Long
    Select Case ModeAvailability
        Case "unavailable"
            MsgBox "お使いの環境で使えるAIモデルがありません。", vbCritical
            ReadyAI.quit
            Exit Sub

        Case "downloadable", "downloading"
            Continue = MsgBox("この機能を初めて利用するには、AIモデルデータのDLが必要です。" & vbCrLf & "DLを開始してもよろしいでしょうか？", vbExclamation + vbYesNo, "通信が発生します")

            If Continue = vbYes Then
                Debug.Print "js実行結果: " & PromptAPI.ModelDownloadProgress
            Else
                ReadyAI.quit
                Exit Sub
            End If

        Case "available":
            MsgBox "既にAIモデルデータがDLされています。" & vbCrLf & "次項のプロシージャを実行してください。", vbInformation, "Ready!"
            Exit Sub
    End Select

    '非同期イベントを発火させ、進捗値を表示
    Dim AIデータ進捗値 As Double
    Do
        ReadyAI.TakeEvents
        DoEvents
        AIデータ進捗値 = PromptAPI.DLProgressValue

        Debug.Print "AIモデルデータをダウンロード中... " & AIデータ進捗値 & "%"
        ReadyAI.sleep

    Loop Until AIデータ進捗値 >= 100

    '4. 完了！
    MsgBox "AIモデルデータのDLが完了しました！" & vbCrLf & "次項のプロシージャを実行してください。", vbInformation, "Finish!"

End Sub



'***************************************************************************************************
'                          ■■■ Prompt API - 一度切りトーク編 ■■■
'***************************************************************************************************
Sub PromptAPI即席チャット()
    Const チャット内容 As String = "こんにちは！あなたは今、Excel VBAから操作されています。自己紹介をしてください。"


    Dim RunAI As New CDPBrowser

    '設定セルから、ユーザ名を取得
    Dim UserName As String
    With ShSetting01_StartBrowser
        UserName = .Range(.UseRangeName(2, "Demo_LocalAI.PromptAPI即席チャット")).value
    End With

    '1. 既存のTargetIDに接続できるか？
    If Not RunAI.reattach(UserName) Then
        '既存のTargetIDじゃないと使えないので終わり
        MsgBox "PromptAPI が利用できるタブの検出に失敗しました。" & vbCrLf & "`PromptAPIの準備`プロシージャから、やり直して下さい。", vbCritical

        RunAI.quit
        Exit Sub
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
        RunAI.TimeOutSecond = 60    'AI応答なので、少し余裕を持たせる
    End If

    '2. 拡張機能クラスに継承させる
    Dim PromptAPI  As New LocalAI_PromptAPI
    PromptAPI.Init RunAI

    '3. 結果をイミディエイトウィンドウに表示
    Debug.Print "--- AIからの回答 ---"
    Debug.Print PromptAPI.instantSession(チャット内容)

End Sub

Sub PromptAPI即席Streamingチャット()
    Const チャット内容 As String = "Excelの歴史を少し述べてください"


    Dim RunAI As New CDPBrowser

    '設定セルから、ユーザ名を取得
    Dim UserName As String
    With ShSetting01_StartBrowser
        UserName = .Range(.UseRangeName(2, "Demo_LocalAI.PromptAPI即席Streamingチャット")).value
    End With

    '1. 既存のTargetIDに接続できるか？
    If Not RunAI.reattach(UserName) Then
        '既存のTargetIDじゃないと使えないので終わり
        MsgBox "PromptAPI が利用できるタブの検出に失敗しました。" & vbCrLf & "`PromptAPIの準備`プロシージャから、やり直して下さい。", vbCritical

        RunAI.quit
        Exit Sub
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
        RunAI.TimeOutSecond = 60    'AI応答なので、少し余裕を持たせる
    End If

    '2. 拡張機能クラスに継承させる
    Dim PromptAPI  As New LocalAI_PromptAPI
    PromptAPI.Init RunAI

    '3. 結果をイミディエイトウィンドウに表示
    Debug.Print "--- AIからのストリーミング回答 ---"
    PromptAPI.instantStreamingSession チャット内容

    '----パターン1----
    'リアルタイム重視
    Dim StreamingData As String
    Do
        DoEvents
        RunAI.TakeEvents
        StreamingData = PromptAPI.StreamingTopTake

        If StrPtr(StreamingData) Then Debug.Print StreamingData;
    Loop Until StreamingData = Chr(30)
    '-----------------

    '----パターン2----
    '後から高速表示
'    Dim StreamingData
'    Do
'        DoEvents
'        RunAI.TakeEvents
'    Loop Until PromptAPI.StreamingEOFExist
'
'    For Each StreamingData In PromptAPI.StreamingColTake
'        Debug.Print StreamingData;
'    Next
    '-----------------

    Debug.Print vbCrLf & "--- AIからのストリーミング回答終了 ---"

End Sub


'***************************************************************************************************
'                          ■■■ Prompt API - 継続トーク編 ■■■
'***************************************************************************************************
Sub PromptAPIセッション保持設定()
    Dim RunAI As New CDPBrowser

    '設定セルから、ユーザ名を取得
    Dim UserName As String
    With ShSetting01_StartBrowser
        UserName = .Range(.UseRangeName(2, "Demo_LocalAI.PromptAPIセッション保持設定")).value
    End With

    '1. 既存のTargetIDに接続できるか？
    If Not RunAI.reattach(UserName) Then
        '既存のTargetIDじゃないと使えないので終わり
        MsgBox "PromptAPI が利用できるタブの検出に失敗しました。" & vbCrLf & "`PromptAPIの準備`プロシージャから、やり直して下さい。", vbCritical

        RunAI.quit
        Exit Sub
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
    End If

    '2. 拡張機能クラスに継承させる
    Dim PromptAPI  As New LocalAI_PromptAPI
    PromptAPI.Init RunAI

    '3. システムプロンプトとN ショット プロンプトの設定
    With PromptAPI
        'システムプロンプト
        .setSystemPrompt = "Classify the following product reviews as either OK or Not OK."

        'N ショット プロンプト
        '※"user"→"assistant"→"user"→... としてどんどん追加されます
        .addInitialPrompt = "Great shoes! I was surprised at how comfortable these boots are for the price. They fit well and are very lightweight."
        .addInitialPrompt = "OK"
        .addInitialPrompt = "Terrible product. The manufacturer must be completely incompetent."
        .addInitialPrompt = "Not OK"
        .addInitialPrompt = "Could be better. Nice quality overall, but for the price I was expecting something more waterproof"
        .addInitialPrompt = "OK"

        '確認用
        Debug.Print .addInitialPrompt
    End With

    '4. 上記を基にセッションを作成
    PromptAPI.createSession
    Debug.Print PromptAPI.objectidPromptAPI

    '5. セッションを保持
    RunAI.KeepSession = True
End Sub

Sub PromptAPIセッション保持からトーク()
    Const チャット内容  As String = "It was a color that didn't suit me."
    Const AISessionID   As String = "-5944820599456107219.1.1"

    Dim RunAI As New CDPBrowser

    '設定セルから、ユーザ名を取得
    Dim UserName As String
    With ShSetting01_StartBrowser
        UserName = .Range(.UseRangeName(2, "Demo_LocalAI.PromptAPIセッション保持設定")).value
    End With

    '1. 既存のTargetIDに接続できるか？
    If Not RunAI.reattach(UserName, existing_) Then
        '既存のTargetIDじゃないと使えないので終わり
        MsgBox "PromptAPI が利用できるタブの検出に失敗しました。" & vbCrLf & "`PromptAPIの準備`プロシージャから、やり直して下さい。", vbCritical

        RunAI.quit
        Exit Sub
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
    End If

    '2. 拡張機能クラスに継承させ、生成したSessionIDを付与
    Dim PromptAPI  As New LocalAI_PromptAPI
    PromptAPI.Init RunAI
    PromptAPI.objectidPromptAPI = AISessionID

    '3. 結果をイミディエイトウィンドウに表示
    Debug.Print "--- AIからの回答 ---"
    Debug.Print PromptAPI.advancePrompt(チャット内容)

End Sub
