Attribute VB_Name = "Test_BiDiAlertRace"
'===================================================================================================
' WebDriverBiDi: 受信ずれ耐久テスト（alert）
'
' 狙い:
' - `jsEval(RunAsyncBiDi:=True)`(script.evaluate) の戻りIDを後で `TakeResultBiDi` で回収する流れに対し、
'   userPromptOpened / handleUserPrompt / 同期の`jsEval` を混在させて、受信ずれを起こしやすくする。
'
' 下記のテスト用のイベントクラス処理ファイルのインポートをしてください
' ・WebDriverBiDiEventtest.cls
'
' 実行:
' - Run_BiDiAlertRaceStress
'===================================================================================================
Option Explicit

Private passCount As Long
Private failCount As Long

' StarterWebScrapingKit ルートを設定してください
Private Const WORKSPACE_PATH As String = ""

Private Function EOk() As String
    EOk = WorksheetFunction.Unichar(9989)
End Function

Public Sub Run_BiDiAlertRaceStress(Optional ByVal iterations As Long = 50)
    '---- JavaScriptアラートの自動処理を無効化 ---
    Dim caps As New Dictionary
    Dim alwaysMatch As New Dictionary
    alwaysMatch.Add "unhandledPromptBehavior", "ignore"
    caps.Add "capabilities", New Dictionary
    caps("capabilities").Add "alwaysMatch", alwaysMatch
    '-------------------------------------------

    '設定シートに基づき、BiDiモードで起動
    Dim t As WebDriverBiDiContext
    Set t = ShSetting01_StartBrowser.StartBiDiModeContext("file:///" & Replace(WORKSPACE_PATH & "\ForDevelopers\OperationCheck\WebDriverBiDi\TestHtml\Test_BiDiAlertRace\Test_BiDiAlertRace.html", "\", "/"), sessionCapabilitiesRequest:=caps)

    Dim wd As WebDriverBiDiMode
    Set wd = t.ThisWebDriverBiDiMode

    'alertの自動処理は、拡張クラス側の`ExEvent`で行う
    Dim DialogAutoClose As New WebDriverBiDiEventtest
    DialogAutoClose.Init t

    passCount = 0
    failCount = 0

    PrintHeader "BiDi alert race stress test 開始"

    Dim paramsBiDi As Dictionary
    Dim targetContext As String
    targetContext = t.context
    If targetContext = "" Then
        Debug.Print "  FAIL | targetContext 取得失敗"
        wd.CloseBrowser
        Exit Sub
    End If

    ' alertイベントを購読
    Set paramsBiDi = New Dictionary
    Dim eventsArray As New Collection
    eventsArray.Add "browsingContext.userPromptOpened"
    paramsBiDi.Add "events", eventsArray
    wd.ExecuteBiDi "session.subscribe", paramsBiDi

    Dim i As Long
    For i = 1 To iterations
        On Error GoTo IterErr

'        PrintSection "iter=" & CStr(i)

        ' イベント蓄積を毎回クリア
        Set wd.BiDiEvents = New Dictionary

        ' 1) alertボタンclick を async 実行（これのID回収が主目的）
        Dim asyncClickId As Long
        asyncClickId = EvalAsync(t, "document.getElementById('btn-bidi-alert').click()", False)

        ' 2) さらに async を1本追加し混雑させる
        Dim asyncDummyId As Long
        asyncDummyId = EvalAsync(t, "document.title", False)

        ' 3) イベント処理設定
        DialogAutoClose.DefaultAccept = True
        DialogAutoClose.DefaultPromptText = "BIDI_INPUT_" & CStr(i)

        ' 4) 同期コマンドを挟んで受信ループ中の競合を誘発
        Dim dummySync As Variant
        dummySync = EvalSyncValue(t, "document.body.clientWidth")

        ' 5) ページ状態の整合確認（alertが閉じた後に更新される値）
        Dim lastMsg As String, statusTxt As String
        lastMsg = CStr(EvalSyncValue(t, "window.__bidi_last_alert"))
        statusTxt = CStr(EvalSyncValue(t, "document.getElementById('status').textContent"))

        If InStr(1, lastMsg, "BIDI_RACE_", vbTextCompare) = 0 Then
            failCount = failCount + 1
            Debug.Print "  FAIL | lastMsg 異常: " & lastMsg
            Exit For
        ElseIf InStr(1, statusTxt, "_closed", vbTextCompare) = 0 Then
            failCount = failCount + 1
            Debug.Print "  FAIL | status 異常: " & statusTxt
            Exit For
        Else
            passCount = passCount + 1
'            Debug.Print "  " & EOk() & " PASS | dialog closed: " & lastMsg
        End If

        ' 6) 非同期ID回収（ここがズレ検知ポイント）
        Dim box1 As String, box2 As String
        Dim ok1 As Boolean, ok2 As Boolean
        ok1 = TryTakeResultBiDi(wd, asyncClickId, box1, 2#)
        ok2 = TryTakeResultBiDi(wd, asyncDummyId, box2, 2#)

        If Not ok1 Then
            failCount = failCount + 1
            Debug.Print "  FAIL | asyncClickId 回収失敗 ID=" & CStr(asyncClickId)
            Exit For
        ElseIf Not ok2 Then
            failCount = failCount + 1
            Debug.Print "  FAIL | asyncDummyId 回収失敗 ID=" & CStr(asyncDummyId)
            Exit For
        Else
            passCount = passCount + 1
'            Debug.Print "  " & EOk() & " PASS | async result IDs recovered"
        End If

        CDPHelpers.Sleep 0.03
        GoTo IterNext

IterErr:
        failCount = failCount + 1
        Debug.Print "  FAIL | iter=" & CStr(i) & " Err: " & Err.Description
        Exit For

IterNext:
        On Error GoTo 0
    Next i

    PrintHeader "テスト完了: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)
    wd.CloseBrowser
End Sub

'非同期で`script.evaluate`を送り、コマンドIDだけを返す
Private Function EvalAsync(t As WebDriverBiDiContext, expr As String, Optional awaitPromise As Boolean = False) As Long
    EvalAsync = t.jsEval(expr, awaitPromise:=awaitPromise, RunAsyncBiDi:=True)
End Function

'同期で`script.evaluate`を送り、値を返す
Private Function EvalSyncValue(t As WebDriverBiDiContext, expr As String, Optional awaitPromise As Boolean = True) As Variant
    EvalSyncValue = t.jsEval(expr, awaitPromise:=awaitPromise)
End Function

Private Function WaitForPromptEvent(wd As WebDriverBiDiMode, evName As String, timeoutSec As Double) As Boolean
    Dim t0 As Double: t0 = Timer
    Do
        wd.TakeEvents
        If Not (wd.BiDiEvents Is Nothing) Then
            If wd.BiDiEvents.Exists("EventMethods") Then
                If wd.BiDiEvents("EventMethods").Exists(evName) Then
                    WaitForPromptEvent = True
                    Exit Function
                End If
            End If
        End If
        CDPHelpers.Sleep 0.02
        DoEvents
        If Timer - t0 > timeoutSec Then Exit Do
    Loop
    WaitForPromptEvent = False
End Function

'非同期コマンドの結果(Json文字列)を、タイムアウトまで回収を試みる
Private Function TryTakeResultBiDi(wd As WebDriverBiDiMode, ByVal cmdId As Long, ByRef receiveBox As String, ByVal timeoutSec As Double) As Boolean
    Dim t0 As Double: t0 = Timer
    Do
        receiveBox = wd.TakeResultBiDi(cmdId)
        If StrPtr(receiveBox) Then
            TryTakeResultBiDi = True
            Exit Function
        End If
        wd.TakeEvents
        CDPHelpers.Sleep 0.02
        DoEvents
        If Timer - t0 > timeoutSec Then Exit Do
    Loop
    TryTakeResultBiDi = False
End Function

Private Sub PrintHeader(msg As String)
    Debug.Print ""
    Debug.Print String(70, "=")
    Debug.Print "  " & msg
    Debug.Print String(70, "=")
End Sub

Private Sub PrintSection(msg As String)
    Debug.Print ""
    Debug.Print "  ── " & msg & " ──"
End Sub
