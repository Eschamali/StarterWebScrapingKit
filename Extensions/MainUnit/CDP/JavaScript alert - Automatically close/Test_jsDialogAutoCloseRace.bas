Attribute VB_Name = "Test_jsDialogAutoCloseRace"
'===================================================================================================
' JSダイアログ自動close拡張を使った「競合ズレ」再現ストレステスト
'
' 目的:
' - RunAsyncCDP=true の click (async command result) と
'   Page.handleJavaScriptDialog (extension がイベント内で同期 invokeMethod) の
'   受信・回収の競合を起こしやすい順序で回す
' - ResultCDPForAsync(asyncClickID) が取り出せず失敗するケースを探す
'
' 実行:
' - VBA: Run_JSDialogAutoClose_RaceTest を実行
'===================================================================================================
'===================================================================================================
' JSダイアログ自動close拡張を使った「競合ズレ」再現ストレステスト
'
' 目的:
' - RunAsyncCDP=true の click (async command result) と
'   Page.handleJavaScriptDialog (extension がイベント内で同期 invokeMethod) の
'   受信・回収の競合を起こしやすい順序で回す
' - ResultCDPForAsync(asyncClickID) が取り出せず失敗するケースを探す
'
' 実行:
' - VBA: Run_JSDialogAutoClose_RaceTest を実行
'===================================================================================================
Option Explicit

Private passCount As Long
Private failCount As Long

'ワークスペースパス（StarterWebScrapingKit ルート）
Private Const WORKSPACE_PATH As String = ""

Private Function EOk() As String
    EOk = WorksheetFunction.Unichar(9989)
End Function

Public Sub Run_JSDialogAutoClose_RaceTest(Optional ByVal iterations As Long = 30)
    Dim br As CDPContext: Set br = 設定シートからのCDP起動ForTab

    br.navigate "file:///" & Replace(WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_jsDialogAutoCloseRace\Test_jsDialogAutoCloseRace.html", "\", "/")
    br.wait

    passCount = 0: failCount = 0

    PrintHeader "JSDialogAutoClose 競合ズレ再現テスト 開始"

    '拡張機能（イベント駆動の自動 close）
    Dim ext As New exCDP_JSDialogAutoClose
    ext.Init br
    ext.DefaultAccept = True
    ext.DefaultPromptText = "N/A"

    br.ExecuteCDP "Page.enable", Nothing

    Dim iter As Long
    For iter = 1 To iterations
'        PrintSection "iter=" & iter

        On Error GoTo Iter_Err

        'button の objectId
        Dim btnOid As Variant
        btnOid = br.jsEval("document.getElementById('btn')", returnByValue:=False, dbgMsg:=False)

        If VarType(btnOid) <> vbString Or Len(btnOid) = 0 Then Err.Raise 5, , "btn objectId が取得できませんでした"

        'RunAsyncCDP=true のクリック（async command result を取り出したい）
        Dim asyncClickId As Variant
        asyncClickId = br.jsEval("function(){ this.click(); }", CStr(btnOid), RunAsyncCDP:=True, dbgMsg:=False)
        If Not IsNumeric(asyncClickId) Or CLng(asyncClickId) <= 0 Then Err.Raise 6, , "asyncClickId が不正です: " & CStr(asyncClickId)

        'さらに async を1個足して、受信バッチの混雑を増やす
        Dim asyncTitleId As Variant
        asyncTitleId = br.jsEval("document.title", RunAsyncCDP:=True, dbgMsg:=False)
        If Not IsNumeric(asyncTitleId) Or CLng(asyncTitleId) <= 0 Then Err.Raise 7, , "asyncTitleId が不正です: " & CStr(asyncTitleId)

        '同期 jsEval を1回挟んで、SendMessage(同期待ち) の受信処理中に
        '拡張のイベント→同期 invokeMethod が入り込む確率を上げる
        Dim dummy As Variant
        dummy = br.jsEval("document.body.clientHeight", returnByValue:=True, dbgMsg:=False)

        'alert が閉じた後に状態が更新される想定
        Dim expectMsg As String
        expectMsg = "RACE_ALERT_" & iter

        Dim statusTxt As String
        statusTxt = CStr(br.jsEval("document.getElementById('status').textContent", returnByValue:=True, dbgMsg:=False))

        Dim lastMsg As String
        lastMsg = CStr(br.jsEval("window.__lastAlertMsg", returnByValue:=True, dbgMsg:=False))

        If InStr(1, statusTxt, expectMsg, vbTextCompare) = 0 Then
            failCount = failCount + 1
            Debug.Print "  FAIL | statusTxt=" & statusTxt
        ElseIf lastMsg <> expectMsg Then
            failCount = failCount + 1
            Debug.Print "  FAIL | lastMsg=" & lastMsg & " expect=" & expectMsg
        Else
            passCount = passCount + 1
'            Debug.Print "  " & EOk() & " PASS | dialog closed ok: " & expectMsg
        End If

        'ここが「競合ズレ」検出ポイント:
        ' - asyncClickId の結果が AccumulatedAsyncResults に残っているか？
        Dim boxClick As Scripting.Dictionary
        Dim boxTitle As Scripting.Dictionary
        Dim okClick As Boolean, okTitle As Boolean

        okClick = TryResultCDPForAsync(br, CLng(asyncClickId), boxClick, 2#)
        okTitle = TryResultCDPForAsync(br, CLng(asyncTitleId), boxTitle, 2#)

        If Not okClick Then
            failCount = failCount + 1
            Debug.Print "  FAIL | ResultCDPForAsync(asyncClickId) 取得できず: id=" & CStr(asyncClickId)
            Debug.Print "  statusTxt=" & statusTxt & " lastMsg=" & lastMsg
            Exit For
        ElseIf Not okTitle Then
            failCount = failCount + 1
            Debug.Print "  FAIL | ResultCDPForAsync(asyncTitleId) 取得できず: id=" & CStr(asyncTitleId)
            Exit For
        Else
'            Debug.Print "  " & EOk() & " PASS | async results recovered ok"
        End If

        'br.sleep 0.05
        GoTo Iter_Next

Iter_Err:
        failCount = failCount + 1
        Debug.Print "  FAIL | iter=" & iter & " Err: " & Err.Description
        Exit For

Iter_Next:
        On Error GoTo 0
    Next iter

    PrintHeader "テスト完了: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)
    Set ext = Nothing
    br.InheritanceCDPBrowser.quit
End Sub

Private Function TryResultCDPForAsync(br As CDPContext, ByVal cmdId As Long, ByRef box As Scripting.Dictionary, ByVal timeoutSec As Double) As Boolean
    Dim t0 As Double: t0 = Timer

    Do
        On Error Resume Next
        Set box = br.InheritanceCDPBrowser.jsConverter.ParseJson(br.ResultCDPFromWithEvents(cmdId))
        
        On Error GoTo 0

        If Not box Is Nothing Then
            TryResultCDPForAsync = True
            Exit Function
        End If

        '取りこぼしを拾うために短くポンプ
        br.InheritanceCDPBrowser.TakeEvents
        DoEvents

        If (Timer - t0) > timeoutSec Then Exit Do
    Loop

    TryResultCDPForAsync = False
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
