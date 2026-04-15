Attribute VB_Name = "Demo_LocalAI"
Option Explicit



Private Const ThisClassName As String = "Demo_LocalAI"



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
    '※オフラインでも遷移できる「バージョン情報」にひとまず設定
    ReadyAI.navigate "edge://version"

    '拡張機能クラスに継承させる
    Dim PromptAPI  As New exCDP_LocalAI
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
                Debug.Print PromptAPI.ModelDownloadProgress
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
