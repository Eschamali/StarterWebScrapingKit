Attribute VB_Name = "Test_BiDiAlertRace"
'===================================================================================================
' WebDriverBiDi: 受信ずれ耐久テスト（alert）
'
' 狙い:
' - invokeMethodAsync("script.evaluate") の戻りIDを後で ResultBiDiForAsync で回収する流れに対し、
'   userPromptOpened / handleUserPrompt / 同期invokeMethod を混在させて、受信ずれを起こしやすくする。
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

    Dim wd As WebDriverBiDiCore
    Set wd = 設定シートからのBiDi起動("file:///" & Replace(WORKSPACE_PATH & "\ForDevelopers\OperationCheck\WebDriverBiDi\TestHtml\Test_BiDiAlertRace\Test_BiDiAlertRace.html", "\", "/"), sessionCapabilitiesRequest:=caps)
    Dim DialogAutoClose As New WebDriverBiDiEventtest
    DialogAutoClose.Init wd
    
    passCount = 0
    failCount = 0

    PrintHeader "BiDi alert race stress test 開始"

    Dim paramsBiDi As Dictionary, resultBiDi As Dictionary
    Dim targetContext As String

    Set resultBiDi = wd.invokeMethod("browsingContext.getTree")
    If Not (resultBiDi Is Nothing) Then
        targetContext = resultBiDi("contexts")(1)("context")
    End If
    If targetContext = "" Then
        Debug.Print "  FAIL | targetContext 取得失敗"
        wd.quit
        Exit Sub
    End If

    ' alertイベントを購読
    Set paramsBiDi = New Dictionary
    Dim eventsArray As New Collection
    eventsArray.Add "browsingContext.userPromptOpened"
    paramsBiDi.Add "events", eventsArray
    wd.invokeMethod "session.subscribe", paramsBiDi

    Dim i As Long
    For i = 1 To iterations
        On Error GoTo IterErr

'        PrintSection "iter=" & CStr(i)

        ' イベント蓄積を毎回クリア
        Set wd.BiDiEvents = New Dictionary

        ' 1) alertボタンclick を async 実行（これのID回収が主目的）
        Dim asyncClickId As Long
        asyncClickId = EvalAsync(wd, targetContext, "document.getElementById('btn-bidi-alert').click()", False)

        ' 2) さらに async を1本追加し混雑させる
        Dim asyncDummyId As Long
        asyncDummyId = EvalAsync(wd, targetContext, "document.title", False)

        ' 3) イベント処理設定
        DialogAutoClose.DefaultAccept = True
        DialogAutoClose.DefaultPromptText = "BIDI_INPUT_" & CStr(i)
        DialogAutoClose.targetContext = targetContext

        ' 4) 同期コマンドを挟んで受信ループ中の競合を誘発
        Dim dummySync As Variant
        dummySync = EvalSyncValue(wd, targetContext, "document.body.clientWidth")

        ' 5) ページ状態の整合確認（alertが閉じた後に更新される値）
        Dim lastMsg As String, statusTxt As String
        lastMsg = CStr(EvalSyncValue(wd, targetContext, "window.__bidi_last_alert"))
        statusTxt = CStr(EvalSyncValue(wd, targetContext, "document.getElementById('status').textContent"))

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
        Dim box1 As Dictionary, box2 As Dictionary
        Dim ok1 As Boolean, ok2 As Boolean
        ok1 = TryResultBiDiForAsync(wd, asyncClickId, box1, 2#)
        ok2 = TryResultBiDiForAsync(wd, asyncDummyId, box2, 2#)

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

        wd.sleep 0.03
        GoTo IterNext

IterErr:
        failCount = failCount + 1
        Debug.Print "  FAIL | iter=" & CStr(i) & " Err: " & Err.Description
        Exit For

IterNext:
        On Error GoTo 0
    Next i

    PrintHeader "テスト完了: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)
    wd.quit
End Sub

Private Function EvalAsync(wd As WebDriverBiDiCore, targetContext As String, expr As String, Optional awaitPromise As Boolean = False) As Long
    Dim paramsBiDi As New Dictionary
    Dim Target As New Dictionary
    Target.Add "context", targetContext

    paramsBiDi.Add "expression", expr
    paramsBiDi.Add "target", Target
    paramsBiDi.Add "awaitPromise", awaitPromise

    EvalAsync = wd.invokeMethodAsync("script.evaluate", paramsBiDi)
End Function

Private Function EvalSyncValue(wd As WebDriverBiDiCore, targetContext As String, expr As String, Optional awaitPromise As Boolean = True) As Variant
    Dim paramsBiDi As New Dictionary
    Dim Target As New Dictionary
    Dim resultBiDi As Dictionary

    Target.Add "context", targetContext
    paramsBiDi.Add "expression", expr
    paramsBiDi.Add "target", Target
    paramsBiDi.Add "awaitPromise", awaitPromise

    Set resultBiDi = wd.invokeMethod("script.evaluate", paramsBiDi)

    If resultBiDi Is Nothing Then Exit Function
    If Not resultBiDi.Exists("result") Then Exit Function
    If Not resultBiDi("result").Exists("value") Then Exit Function
    EvalSyncValue = resultBiDi("result")("value")
End Function

Private Function WaitForPromptEvent(wd As WebDriverBiDiCore, evName As String, timeoutSec As Double) As Boolean
    Dim t0 As Double: t0 = Timer
    Do
        wd.TakeEvents
        If Not (wd.BiDiEvents Is Nothing) Then
            If wd.BiDiEvents("EventMethods").Exists(evName) Then
                WaitForPromptEvent = True
                Exit Function
            End If
        End If
        wd.sleep 0.02
        DoEvents
        If Timer - t0 > timeoutSec Then Exit Do
    Loop
    WaitForPromptEvent = False
End Function

Private Function TryResultBiDiForAsync(wd As WebDriverBiDiCore, ByVal cmdId As Long, ByRef receiveBox As Dictionary, ByVal timeoutSec As Double) As Boolean
    Dim t0 As Double: t0 = Timer
    Do
        Set receiveBox = wd.ResultBiDiForAsync(cmdId)
        If Not receiveBox Is Nothing Then
            TryResultBiDiForAsync = True
            Exit Function
        End If
        wd.TakeEvents
        wd.sleep 0.02
        DoEvents
        If Timer - t0 > timeoutSec Then Exit Do
    Loop
    TryResultBiDiForAsync = False
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
