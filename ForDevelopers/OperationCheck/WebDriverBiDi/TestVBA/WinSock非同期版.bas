Attribute VB_Name = "WinSockîÒìØä˙î≈"
Option Explicit

'================================================================================
' Main10 Completion-Signal Gate - ExecuteBiDiAsync version
'
' - script.evaluate is submitted with ExecuteBiDiAsync and immediately returns an ID.
' - Application.OnTime periodically calls TakeEvents once and checks that ID.
' - No VBA-side receive wait loop is used, so Excel returns to the user between polls.
' - The JavaScript Promise itself performs:
'       arm MutationObserver -> click -> wait for rewrite -> wait stable 800 ms
'================================================================================

Private Const MAIN10_POLL_SECONDS As Long = 1

Private Enum Main10AsyncPhase
    Main10PhaseIdle = 0
    Main10PhaseWait2015 = 1
    Main10PhaseWait2014 = 2
End Enum

Private g_Main10Browser As WebDriverBiDiContext
Private g_Main10CommandId As Long
Private g_Main10Phase As Main10AsyncPhase
Private g_Main10Active As Boolean
Private g_Main10PumpRunning As Boolean
Private g_Main10NextPollAt As Date

'***************************************************************************************************
'* Main entry point.
'*
'* This procedure returns after the first asynchronous command has been submitted and the first
'* poll has been scheduled. Do not quit the browser here; completion/quit is handled by the pump.
'***************************************************************************************************
Public Sub BiDiÇ…ÇÊÇÈñ`åØÇÃénÇ‹ÇË_îÒìØä˙î≈()
    Const PROC As String = "Demo_WebDriverBiDi_Main10_Async.BiDiÇ…ÇÊÇÈñ`åØÇÃénÇ‹ÇË_îÒìØä˙î≈"

    If g_Main10Active Then
        Err.Raise vbObjectError + 2200, PROC, _
            "Main10 asynchronous processing is already running."
    End If

    On Error GoTo ErrorHandler

    Set g_Main10Browser = ShSetting01_StartBrowser.StartBiDiModeContext

    '-----WebSocketÉãÅ[Ég-----
'    Dim UserName As String
'    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_CDP.SetupWebSocketMode")
'
'    Dim WebSocketBiDi As New CDPCoreViaWebSocket
'    WebSocketBiDi.ConnectCDP UserName, "/devtools/browser/1ee505aa-2eaf-4b0e-874b-c9a8ae154442"
'
'    Dim b As New WebDriverBiDiMode
'    b.reattach UserName, WebSocketMode:=WebSocketBiDi
'    Set g_Main10Browser = b.newTab(setMain:=True)
    '-------------------------

    'Navigation remains the existing synchronous helper.
    'The Main10 click-and-wait operations below use ExecuteBiDiAsync.
    g_Main10Browser.navigate _
        "https://www.scrapethissite.com/pages/ajax-javascript/#2010"

    g_Main10Active = True
    g_Main10Phase = Main10PhaseWait2015

    g_Main10CommandId = Main10_ArmContentSignalAndClickAsync( _
        g_Main10Browser, _
        "//*[@id='table-body']", _
        "//section[@id='oscars']//a[@id='2015']")

    Debug.Print "Main10 async submitted: 2015 / command ID=" & g_Main10CommandId

    Main10_SchedulePump
    Exit Sub

ErrorHandler:
    Dim savedNumber As Long
    Dim savedSource As String
    Dim savedDescription As String

    savedNumber = Err.Number
    savedSource = Err.Source
    savedDescription = Err.Description

    Main10_ReleaseBrowser
    Err.Raise savedNumber, savedSource, savedDescription
End Sub

'***************************************************************************************************
'* Submit one Completion-Signal Gate operation asynchronously.
'*
'* Return value: BiDi command ID.
'* The browser-side Promise remains pending until the DOM rewrite is observed and stable.
'***************************************************************************************************
Public Function Main10_ArmContentSignalAndClickAsync( _
    ByVal browser As WebDriverBiDiContext, _
    ByVal signalXPath As String, _
    ByVal clickXPath As String, _
    Optional ByVal searchTimeoutMs As Long = 10000, _
    Optional ByVal minStableMs As Long = 800, _
    Optional ByVal completionTimeoutMs As Long = 15000) As Long

    Const PROC As String = _
        "Demo_WebDriverBiDi_Main10_Async.Main10_ArmContentSignalAndClickAsync"

    If browser Is Nothing Then
        Err.Raise 91, PROC, "browser is Nothing."
    End If

    If Len(Trim$(signalXPath)) = 0 Then
        Err.Raise 5, PROC, "signalXPath is empty."
    End If

    If Len(Trim$(clickXPath)) = 0 Then
        Err.Raise 5, PROC, "clickXPath is empty."
    End If

    If searchTimeoutMs <= 0 Then searchTimeoutMs = 10000
    If minStableMs <= 0 Then minStableMs = 800

    If completionTimeoutMs <= minStableMs Then
        Err.Raise 5, PROC, _
            "completionTimeoutMs must be greater than minStableMs."
    End If

    'RunAsyncJavScript
    Main10_ArmContentSignalAndClickAsync = browser.jsEval(BuildMain10AsyncExpression( _
                                                            signalXPath, _
                                                            clickXPath, _
                                                            searchTimeoutMs, _
                                                            minStableMs, _
                                                            completionTimeoutMs), awaitPromise:=True, RunAsyncBiDi:=True)

    If Main10_ArmContentSignalAndClickAsync <= 0 Then
        Err.Raise vbObjectError + 2201, PROC, _
            "ExecuteBiDiAsync did not return a valid command ID."
    End If
End Function

'***************************************************************************************************
'* Non-blocking result check.
'*
'* - Calls TakeEvents only once.
'* - Returns False if the command result has not arrived yet.
'* - Returns True after validating a successful result.
'* - Raises an error for BiDi errors, JavaScript exceptions, or an unexpected result.
'***************************************************************************************************
Public Function Main10_TryTakeAsyncCompletion( _
    ByVal browser As WebDriverBiDiContext, _
    ByVal commandId As Long) As Boolean

    Const PROC As String = _
        "Demo_WebDriverBiDi_Main10_Async.Main10_TryTakeAsyncCompletion"

    If browser Is Nothing Then
        Err.Raise 91, PROC, "browser is Nothing."
    End If

    If commandId <= 0 Then
        Err.Raise 5, PROC, "commandId is invalid."
    End If

    'Drain currently available WinSock data without entering a response-wait loop.
    browser.ThisWebDriverBiDiMode.TakeEvents True

    Dim rawJson As String
    rawJson = browser.ThisWebDriverBiDiMode.TakeResultBiDi(commandId)

    'No result yet. The caller can return to Excel and check again later.
    If StrPtr(rawJson) = 0 Then Exit Function

    ValidateMain10AsyncResponse rawJson
    Main10_TryTakeAsyncCompletion = True
End Function

'***************************************************************************************************
'* Cooperative asynchronous pump.
'*
'* Application.OnTime calls this public no-argument procedure. Each invocation performs only one
'* non-blocking receive/check pass, then returns to Excel or schedules the next pass.
'***************************************************************************************************
Public Sub Main10_AsyncPump()
    If Not g_Main10Active Then Exit Sub
    If g_Main10PumpRunning Then Exit Sub

    g_Main10PumpRunning = True
    g_Main10NextPollAt = 0

    On Error GoTo ErrorHandler

    If Not Main10_TryTakeAsyncCompletion( _
        g_Main10Browser, _
        g_Main10CommandId) Then

        g_Main10PumpRunning = False
        Main10_SchedulePump
        Exit Sub
    End If

    Select Case g_Main10Phase
        Case Main10PhaseWait2015
            Debug.Print "Main10 async completed: 2015 / command ID=" & _
                        g_Main10CommandId

            g_Main10Phase = Main10PhaseWait2014
            g_Main10CommandId = Main10_ArmContentSignalAndClickAsync( _
                g_Main10Browser, _
                "//*[@id='table-body']", _
                "//section[@id='oscars']//a[@id='2014']")

            Debug.Print "Main10 async submitted: 2014 / command ID=" & _
                        g_Main10CommandId

            g_Main10PumpRunning = False
            Main10_SchedulePump

        Case Main10PhaseWait2014
            Debug.Print "Main10 async completed: 2014 / command ID=" & _
                        g_Main10CommandId

            g_Main10PumpRunning = False
            Main10_CompleteAndClose

        Case Else
            Err.Raise vbObjectError + 2202, _
                "Demo_WebDriverBiDi_Main10_Async.Main10_AsyncPump", _
                "Unknown asynchronous phase."
    End Select

    Exit Sub

ErrorHandler:
    Dim savedNumber As Long
    Dim savedSource As String
    Dim savedDescription As String

    savedNumber = Err.Number
    savedSource = Err.Source
    savedDescription = Err.Description

    g_Main10PumpRunning = False
    Main10_CancelScheduledPump
    Main10_ReleaseBrowser

    MsgBox _
        "Main10 asynchronous processing failed." & vbCrLf & vbCrLf & _
        "Error " & savedNumber & vbCrLf & _
        savedSource & vbCrLf & _
        savedDescription, _
        vbCritical, _
        "Main10 Async"
End Sub

'***************************************************************************************************
'* Explicit cancellation entry point.
'***************************************************************************************************
Public Sub Main10_AsyncCancel()
    Main10_CancelScheduledPump
    Main10_ReleaseBrowser
    Debug.Print "Main10 asynchronous processing was cancelled."
End Sub

Private Sub Main10_CompleteAndClose()
    Debug.Print "Main10 asynchronous processing completed successfully."
    Main10_CancelScheduledPump
    Main10_ReleaseBrowser
End Sub

Private Sub Main10_SchedulePump()
    If Not g_Main10Active Then Exit Sub

    g_Main10NextPollAt = Now + TimeSerial(0, 0, MAIN10_POLL_SECONDS)

    Application.OnTime _
        EarliestTime:=g_Main10NextPollAt, _
        Procedure:=Main10_PumpProcedureName, _
        Schedule:=True
End Sub

Private Sub Main10_CancelScheduledPump()
    On Error Resume Next

    If g_Main10NextPollAt <> 0 Then
        Application.OnTime _
            EarliestTime:=g_Main10NextPollAt, _
            Procedure:=Main10_PumpProcedureName, _
            Schedule:=False
    End If

    g_Main10NextPollAt = 0
    On Error GoTo 0
End Sub

Private Function Main10_PumpProcedureName() As String
    Main10_PumpProcedureName = _
        "'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!Main10_AsyncPump"
End Function

Private Sub Main10_ReleaseBrowser()
    On Error Resume Next

    If Not g_Main10Browser Is Nothing Then
        g_Main10Browser.ThisWebDriverBiDiMode.quit
    End If

    Set g_Main10Browser = Nothing
    g_Main10CommandId = 0
    g_Main10Phase = Main10PhaseIdle
    g_Main10Active = False
    g_Main10PumpRunning = False
    g_Main10NextPollAt = 0

    On Error GoTo 0
End Sub

Private Sub ValidateMain10AsyncResponse(ByRef rawJson As String)
    Const PROC As String = _
        "Demo_WebDriverBiDi_Main10_Async.ValidateMain10AsyncResponse"

    Dim envelope As BiDiCDPJson
    Set envelope = BiDiCDPJson.Parse(rawJson)

    If envelope Is Nothing Then
        Err.Raise vbObjectError + 2210, PROC, _
            "The asynchronous BiDi response could not be parsed."
    End If

    'Top-level WebDriver BiDi command error.
    If envelope.ExistsKey("error") Then
        Dim bidiError As String
        bidiError = envelope.StringKey("error")

        If envelope.ExistsKey("message") Then
            If Len(bidiError) > 0 Then bidiError = bidiError & ": "
            bidiError = bidiError & envelope.StringKey("message")
        End If

        If Len(bidiError) = 0 Then bidiError = "Unknown BiDi command error."

        Err.Raise vbObjectError + 2211, PROC, bidiError
    End If

    If Not envelope.ExistsKey("result") Then
        Err.Raise vbObjectError + 2212, PROC, _
            "The asynchronous BiDi response has no result member."
    End If

    Dim evaluateResult As BiDiCDPJson
    Set evaluateResult = envelope.NodeKey("result")

    If evaluateResult Is Nothing Then
        Err.Raise vbObjectError + 2213, PROC, _
            "script.evaluate returned no result object."
    End If

    'script.evaluate can return type=success or type=exception.
    If LCase$(evaluateResult.StringKey("type")) <> "success" Then
        Dim detail As String
        Dim exceptionDetails As BiDiCDPJson

        If evaluateResult.ExistsKey("exceptionDetails") Then
            Set exceptionDetails = evaluateResult.NodeKey("exceptionDetails")

            If Not exceptionDetails Is Nothing Then
                detail = exceptionDetails.StringKey("text")
            End If
        End If

        If Len(detail) = 0 Then detail = "Unknown JavaScript exception."

        Err.Raise vbObjectError + 2214, PROC, detail
    End If

    Dim remoteValue As BiDiCDPJson
    Set remoteValue = evaluateResult.NodeKey("result")

    If remoteValue Is Nothing Then
        Err.Raise vbObjectError + 2215, PROC, _
            "script.evaluate returned no remote value."
    End If

    If remoteValue.StringKey("value") <> "OK" Then
        Err.Raise vbObjectError + 2216, PROC, _
            "Unexpected completion result: " & remoteValue.StringKey("value")
    End If
End Sub

Private Function BuildMain10AsyncExpression( _
    ByVal signalXPath As String, _
    ByVal clickXPath As String, _
    ByVal searchTimeoutMs As Long, _
    ByVal minStableMs As Long, _
    ByVal completionTimeoutMs As Long) As String

    Dim js As String

    js = "(async()=>{"
    js = js & "const signalXPath=" & JsStringLiteralForKit(signalXPath) & ";"
    js = js & "const clickXPath=" & JsStringLiteralForKit(clickXPath) & ";"
    js = js & "const searchTimeout=" & CStr(searchTimeoutMs) & ";"
    js = js & "const stableMs=" & CStr(minStableMs) & ";"
    js = js & "const completionTimeout=" & CStr(completionTimeoutMs) & ";"

    js = js & "const byXPath=(xp)=>{"
    js = js & "try{return document.evaluate(xp,document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;}"
    js = js & "catch(e){throw new Error('Invalid XPath: '+xp+' / '+e.message);}};"

    js = js & "const waitForXPath=async(xp,timeout)=>{"
    js = js & "const started=performance.now();"
    js = js & "for(;;){const node=byXPath(xp);if(node)return node;"
    js = js & "if(performance.now()-started>=timeout)throw new Error('Click target not found: '+xp);"
    js = js & "await new Promise(r=>setTimeout(r,50));}};"

    js = js & "const clickNode=await waitForXPath(clickXPath,searchTimeout);"
    js = js & "const signalNode=byXPath(signalXPath);"
    js = js & "if(!signalNode)throw new Error('ArmContentSignal target not found at arm time: '+signalXPath);"

    js = js & "return await new Promise((resolve,reject)=>{"
    js = js & "let observer=null,quietTimer=0,deadlineTimer=0,signalSeen=false,done=false;"
    js = js & "const cleanup=()=>{if(observer)observer.disconnect();clearTimeout(quietTimer);clearTimeout(deadlineTimer);};"
    js = js & "const fail=(message)=>{if(done)return;done=true;cleanup();reject(new Error(message));};"
    js = js & "const scheduleStable=()=>{clearTimeout(quietTimer);quietTimer=setTimeout(()=>{"
    js = js & "if(done||!signalSeen)return;done=true;cleanup();resolve('OK');},stableMs);};"

    js = js & "observer=new MutationObserver(records=>{"
    js = js & "if(records.length===0)return;signalSeen=true;scheduleStable();});"
    js = js & "observer.observe(signalNode,{childList:true,characterData:true,subtree:true});"
    js = js & "deadlineTimer=setTimeout(()=>fail('Timed out waiting for content rewrite: '+signalXPath),completionTimeout);"

    js = js & "try{"
    js = js & "if(!clickNode.isConnected)throw new Error('Click target became detached before click.');"
    js = js & "clickNode.scrollIntoView({block:'center',inline:'center'});"
    js = js & "try{clickNode.focus({preventScroll:true});}catch(_){try{clickNode.focus();}catch(__){}}"
    js = js & "if(typeof clickNode.click!=='function')throw new Error('Target does not support click().');"
    js = js & "clickNode.click();"
    js = js & "}catch(e){fail('Click failed: '+e.message);}});"
    js = js & "})()"

    BuildMain10AsyncExpression = js
End Function

Private Function JsStringLiteralForKit(ByVal value As String) As String
    Dim escaped As String

    escaped = Replace(value, "\", "\\")
    escaped = Replace(escaped, """", "\""")
    escaped = Replace(escaped, vbCr, "\r")
    escaped = Replace(escaped, vbLf, "\n")
    escaped = Replace(escaped, vbTab, "\t")

    JsStringLiteralForKit = """" & escaped & """"
End Function
