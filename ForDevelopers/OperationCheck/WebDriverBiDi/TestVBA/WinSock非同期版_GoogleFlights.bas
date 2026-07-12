Attribute VB_Name = "WinSock非同期版_GoogleFlights"
Option Explicit

'================================================================================
' Google Flights (Sapporo -> Paris / One way) - Kit ExecuteBiDiAsync version 4
'
' Architecture:
'   - Every BiDi command is submitted with ExecuteBiDiAsync.
'   - Application.OnTime invokes a short receive pump; no VBA response-wait loop.
'   - Google Flights' Wiz input replacement is handled inside browser-side async JS.
'   - Calendar clicks are separate async commands and re-resolve the live DOM.
'   - XPath searches inspect all matches and select the first visible node.
'   - Immediate-window logs use [hh:mm:ss.fff] wall-clock timestamps.
'   - Calendar/Search completion is gated by BiDi network.responseCompleted events:
'         GetCalendarPicker
'         GetShoppingResults
'
' Entry point:
'   BiDiによる冒険の始まり_GoogleFlights非同期版
'
' Cancellation:
'   GoogleFlights_AsyncCancel
'
' Notes:
'   - The URL includes hl=en because the XPath locators use English aria-labels.
'   - This module intentionally does not use generic network-idle consensus.
'     Specific DOM and network completion signals are used instead, so background
'     Google telemetry does not prevent completion.
'================================================================================

Private Const GF_POLL_SECONDS As Long = 1
Private Const GF_PHASE_TIMEOUT_SECONDS As Long = 90
Private Const GF_URL As String = "https://www.google.com/travel/flights?hl=en"

Private Enum GoogleFlightsAsyncPhase
    GFPhaseIdle = 0
    GFPhaseWaitSubscribe = 1
    GFPhaseWaitNavigate = 2
    GFPhaseWaitTripType = 3
    GFPhaseWaitDeparture = 4
    GFPhaseWaitDestination = 5
    GFPhaseWaitCalendarOpen = 6
    GFPhaseWaitDateSelect = 7
    GFPhaseWaitCalendarDone = 8
    GFPhaseWaitSearch = 9
End Enum

Private g_GF_Browser As WebDriverBiDiContext
Private g_GF_CommandId As Long
Private g_GF_Phase As GoogleFlightsAsyncPhase
Private g_GF_Active As Boolean
Private g_GF_PumpRunning As Boolean
Private g_GF_NextPollAt As Date
Private g_GF_PhaseStartedAt As Date
Private g_GF_CommandCompleted As Boolean
Private g_GF_NetworkPattern As String
Private g_GF_NetworkMatched As Boolean

'***************************************************************************************************
'* Starts the fully asynchronous Google Flights sequence.
'*
'* This procedure returns after session.subscribe has been submitted. Browser work continues via
'* Application.OnTime + TakeEvents + TakeResultBiDi.
'***************************************************************************************************
Public Sub BiDiによる冒険の始まり_GoogleFlights非同期版()
    Const PROC As String = _
        "Demo_GoogleFlights_Async.BiDiによる冒険の始まり_GoogleFlights非同期版"

    If g_GF_Active Then
        Err.Raise vbObjectError + 2400, PROC, _
            "Google Flights asynchronous processing is already running."
    End If

    On Error GoTo ErrorHandler

    Set g_GF_Browser = 設定シートからのBiDi起動ForTab

    'Enable event accumulation before subscribing, so no responseCompleted event is discarded.
    Set g_GF_Browser.InheritanceWebDriverBiDiMode.BiDiEvents = New Dictionary

    g_GF_Active = True
    g_GF_Phase = GFPhaseWaitSubscribe

    Dim subscribeParams As Dictionary
    Set subscribeParams = New Dictionary

    Dim eventsArray As Collection
    Set eventsArray = New Collection
    eventsArray.Add "network.responseCompleted"

    subscribeParams.Add "events", eventsArray

    'Limit high-volume network events to the controlled tab.
    Dim contextsArray As Collection
    Set contextsArray = New Collection
    contextsArray.Add g_GF_Browser.context
    subscribeParams.Add "contexts", contextsArray

    g_GF_CommandId = _
        g_GF_Browser.InheritanceWebDriverBiDiMode.ExecuteBiDiAsync( _
            "session.subscribe", _
            subscribeParams, _
            True)

    GF_BeginPhase GFPhaseWaitSubscribe
    GF_Log "Google Flights async submitted: session.subscribe / ID=" & g_GF_CommandId

    GF_SchedulePump
    Exit Sub

ErrorHandler:
    Dim savedNumber As Long
    Dim savedSource As String
    Dim savedDescription As String

    savedNumber = Err.Number
    savedSource = Err.Source
    savedDescription = Err.Description

    GF_ReleaseBrowser
    Err.Raise savedNumber, savedSource, savedDescription
End Sub

'***************************************************************************************************
'* Cooperative asynchronous receive/state pump.
'*
'* Each invocation:
'*   1. drains currently available messages once,
'*   2. records a matching network signal if one arrived,
'*   3. checks the current command ID once,
'*   4. advances one state or returns control to Excel.
'***************************************************************************************************
Public Sub GoogleFlights_AsyncPump()
    If Not g_GF_Active Then Exit Sub
    If g_GF_PumpRunning Then Exit Sub

    g_GF_PumpRunning = True
    g_GF_NextPollAt = 0

    On Error GoTo ErrorHandler

    'One non-blocking WinSock receive/drain pass.
    g_GF_Browser.InheritanceWebDriverBiDiMode.TakeEvents True

    If Len(g_GF_NetworkPattern) > 0 Then
        If Not g_GF_NetworkMatched Then
            g_GF_NetworkMatched = GF_HasNetworkPattern(g_GF_NetworkPattern)
            If g_GF_NetworkMatched Then
                GF_Log "Google Flights network gate matched: " & g_GF_NetworkPattern
            End If
        End If
    End If

    If Not g_GF_CommandCompleted Then
        Dim rawJson As String
        rawJson = _
            g_GF_Browser.InheritanceWebDriverBiDiMode.TakeResultBiDi( _
                g_GF_CommandId)

        If StrPtr(rawJson) Then
            GF_ValidateCommandResponse rawJson, _
                (g_GF_Phase >= GFPhaseWaitTripType)
            g_GF_CommandCompleted = True
        End If
    End If

    If GF_PhaseTimedOut Then
        Err.Raise vbObjectError + 2401, _
            "Demo_GoogleFlights_Async.GoogleFlights_AsyncPump", _
            "Timed out in phase: " & GF_PhaseName(g_GF_Phase) & _
            IIf(Len(g_GF_NetworkPattern) > 0 And Not g_GF_NetworkMatched, _
                " / missing network signal: " & g_GF_NetworkPattern, _
                vbNullString)
    End If

    'The current step is not done until both its command and optional network gate are complete.
    If Not g_GF_CommandCompleted Then
        g_GF_PumpRunning = False
        GF_SchedulePump
        Exit Sub
    End If

    If Len(g_GF_NetworkPattern) > 0 And Not g_GF_NetworkMatched Then
        g_GF_PumpRunning = False
        GF_SchedulePump
        Exit Sub
    End If

    Select Case g_GF_Phase
        Case GFPhaseWaitSubscribe
            GF_Log "Google Flights async completed: session.subscribe"
            GF_SubmitNavigate

        Case GFPhaseWaitNavigate
            GF_Log "Google Flights async completed: navigation"
            GF_SubmitScript GFPhaseWaitTripType, _
                GF_BuildTripTypeExpression

        Case GFPhaseWaitTripType
            GF_Log "Google Flights async completed: One way"
            GF_SubmitScript GFPhaseWaitDeparture, _
                GF_BuildDepartureExpression

        Case GFPhaseWaitDeparture
            GF_Log "Google Flights async completed: Sapporo"
            GF_SubmitScript GFPhaseWaitDestination, _
                GF_BuildDestinationExpression

        Case GFPhaseWaitDestination
            GF_Log "Google Flights async completed: Paris"
            GF_ResetNetworkGate "GetCalendarPicker"
            GF_SubmitScript GFPhaseWaitCalendarOpen, _
                GF_BuildCalendarOpenExpression

        Case GFPhaseWaitCalendarOpen
            GF_Log "Google Flights async completed: calendar opened"
            GF_SubmitScript GFPhaseWaitDateSelect, _
                GF_BuildDateSelectExpression

        Case GFPhaseWaitDateSelect
            GF_Log "Google Flights async completed: departure date selected"
            GF_SubmitScript GFPhaseWaitCalendarDone, _
                GF_BuildCalendarDoneExpression

        Case GFPhaseWaitCalendarDone
            GF_Log "Google Flights async completed: calendar Done"
            GF_ResetNetworkGate "GetShoppingResults"
            GF_SubmitScript GFPhaseWaitSearch, _
                GF_BuildSearchExpression

        Case GFPhaseWaitSearch
            GF_Log "Google Flights asynchronous processing completed successfully."
            g_GF_PumpRunning = False
            GF_CompleteAndClose
            MsgBox "Google Flights Test Completed", _
                   vbInformation, _
                   "Google Flights Async"
            Exit Sub

        Case Else
            Err.Raise vbObjectError + 2402, _
                "Demo_GoogleFlights_Async.GoogleFlights_AsyncPump", _
                "Unknown Google Flights asynchronous phase."
    End Select

    g_GF_PumpRunning = False
    GF_SchedulePump
    Exit Sub

ErrorHandler:
    Dim savedNumber As Long
    Dim savedSource As String
    Dim savedDescription As String

    savedNumber = Err.Number
    savedSource = Err.Source
    savedDescription = Err.Description

    Dim failedPhase As GoogleFlightsAsyncPhase
    failedPhase = g_GF_Phase

    GF_Log "Google Flights asynchronous processing failed: " & _
           "phase=" & GF_PhaseName(failedPhase) & _
           " / Error " & CStr(savedNumber) & _
           " / " & savedSource & _
           " / " & savedDescription

    g_GF_PumpRunning = False
    GF_CancelScheduledPump
    GF_ReleaseBrowser

    MsgBox _
        "Google Flights asynchronous processing failed." & vbCrLf & vbCrLf & _
        "Phase: " & GF_PhaseName(failedPhase) & vbCrLf & _
        "Error " & savedNumber & vbCrLf & _
        savedSource & vbCrLf & _
        savedDescription, _
        vbCritical, _
        "Google Flights Async"
End Sub

'***************************************************************************************************
'* Explicit user cancellation entry point.
'***************************************************************************************************
Public Sub GoogleFlights_AsyncCancel()
    GF_CancelScheduledPump
    GF_ReleaseBrowser
    GF_Log "Google Flights asynchronous processing was cancelled."
End Sub

Private Sub GF_SubmitNavigate()
    Dim navParams As Dictionary
    Set navParams = New Dictionary

    navParams.Add "url", GF_URL
    navParams.Add "wait", "complete"

    'Context.ExecuteBiDiAsync adds the current context at the correct top level for navigate.
    g_GF_CommandId = _
        g_GF_Browser.ExecuteBiDiAsync( _
            "browsingContext.navigate", _
            navParams, _
            True)

    GF_BeginPhase GFPhaseWaitNavigate
    GF_Log "Google Flights async submitted: navigation / ID=" & g_GF_CommandId
End Sub

Private Sub GF_SubmitScript( _
    ByVal nextPhase As GoogleFlightsAsyncPhase, _
    ByVal expression As String)

    Const PROC As String = "Demo_GoogleFlights_Async.GF_SubmitScript"

    Dim paramsBiDi As Dictionary
    Set paramsBiDi = New Dictionary

    paramsBiDi.Add "expression", expression

    Dim target As Dictionary
    Set target = New Dictionary
    target.Add "context", g_GF_Browser.context

    paramsBiDi.Add "target", target
    paramsBiDi.Add "awaitPromise", True

    'Use WebDriverBiDiMode directly because script.evaluate requires context in target.context.
    g_GF_CommandId = _
        g_GF_Browser.InheritanceWebDriverBiDiMode.ExecuteBiDiAsync( _
            "script.evaluate", _
            paramsBiDi, _
            True)

    If g_GF_CommandId <= 0 Then
        Err.Raise vbObjectError + 2403, PROC, _
            "ExecuteBiDiAsync did not return a valid command ID."
    End If

    GF_BeginPhase nextPhase
    GF_Log "Google Flights async submitted: " & _
           GF_PhaseName(nextPhase) & _
           " / ID=" & g_GF_CommandId
End Sub

'***************************************************************************************************
'* Writes a consistently timestamped entry to the VBA Immediate window.
'*
'* Timer supplies the fractional second. Its actual precision depends on Windows/VBA timer
'* resolution, but the output is normalized to three decimal places for log comparison.
'***************************************************************************************************
Private Sub GF_Log(ByVal message As String)
    Debug.Print "[" & GF_LogTimestamp() & "] " & message
End Sub

Private Function GF_LogTimestamp() As String
    Dim currentTime As Date
    Dim timerValue As Double
    Dim milliseconds As Long

    currentTime = Now
    timerValue = Timer
    milliseconds = CLng(Fix((timerValue - Fix(timerValue)) * 1000#))

    GF_LogTimestamp = _
        Format$(currentTime, "hh:nn:ss") & "." & _
        Format$(milliseconds, "000")
End Function

Private Sub GF_BeginPhase(ByVal phase As GoogleFlightsAsyncPhase)
    g_GF_Phase = phase
    g_GF_PhaseStartedAt = Now
    g_GF_CommandCompleted = False

    'Network gate is retained only when explicitly armed immediately before this phase.
    If phase <> GFPhaseWaitCalendarOpen And phase <> GFPhaseWaitSearch Then
        g_GF_NetworkPattern = vbNullString
        g_GF_NetworkMatched = False
    End If
End Sub

Private Sub GF_ResetNetworkGate(ByVal urlPattern As String)
    g_GF_NetworkPattern = urlPattern
    g_GF_NetworkMatched = False

    'One-shot event history: old matching requests must not satisfy the new action.
    Set g_GF_Browser.InheritanceWebDriverBiDiMode.BiDiEvents = New Dictionary
End Sub

Private Function GF_HasNetworkPattern(ByVal urlPattern As String) As Boolean
    On Error GoTo SafeExit

    Dim eventsRoot As Dictionary
    Set eventsRoot = _
        g_GF_Browser.InheritanceWebDriverBiDiMode.BiDiEvents

    If eventsRoot Is Nothing Then Exit Function
    If Not eventsRoot.Exists("EventMethods") Then Exit Function

    Dim eventMethods As Dictionary
    Set eventMethods = eventsRoot("EventMethods")

    Const EVENT_NAME As String = "network.responseCompleted"
    If Not eventMethods.Exists(EVENT_NAME) Then Exit Function

    Dim eventItem As Variant
    Dim serialized As String

    For Each eventItem In eventMethods(EVENT_NAME)
        serialized = WebJsonConverter.serialize(eventItem)

        If InStr(1, serialized, urlPattern, vbTextCompare) > 0 Then
            GF_HasNetworkPattern = True
            Exit Function
        End If
    Next eventItem

SafeExit:
End Function

Private Function GF_PhaseTimedOut() As Boolean
    If g_GF_PhaseStartedAt = 0 Then Exit Function

    GF_PhaseTimedOut = _
        (DateDiff("s", g_GF_PhaseStartedAt, Now) > _
         GF_PHASE_TIMEOUT_SECONDS)
End Function

Private Function GF_PhaseName( _
    ByVal phase As GoogleFlightsAsyncPhase) As String

    Select Case phase
        Case GFPhaseIdle: GF_PhaseName = "Idle"
        Case GFPhaseWaitSubscribe: GF_PhaseName = "Subscribe"
        Case GFPhaseWaitNavigate: GF_PhaseName = "Navigate"
        Case GFPhaseWaitTripType: GF_PhaseName = "Trip type"
        Case GFPhaseWaitDeparture: GF_PhaseName = "Departure city"
        Case GFPhaseWaitDestination: GF_PhaseName = "Destination city"
        Case GFPhaseWaitCalendarOpen: GF_PhaseName = "Calendar open"
        Case GFPhaseWaitDateSelect: GF_PhaseName = "Date select"
        Case GFPhaseWaitCalendarDone: GF_PhaseName = "Calendar Done"
        Case GFPhaseWaitSearch: GF_PhaseName = "Search"
        Case Else: GF_PhaseName = "Unknown"
    End Select
End Function

Private Sub GF_SchedulePump()
    If Not g_GF_Active Then Exit Sub

    g_GF_NextPollAt = _
        Now + TimeSerial(0, 0, GF_POLL_SECONDS)

    Application.OnTime _
        EarliestTime:=g_GF_NextPollAt, _
        Procedure:=GF_PumpProcedureName, _
        Schedule:=True
End Sub

Private Sub GF_CancelScheduledPump()
    On Error Resume Next

    If g_GF_NextPollAt <> 0 Then
        Application.OnTime _
            EarliestTime:=g_GF_NextPollAt, _
            Procedure:=GF_PumpProcedureName, _
            Schedule:=False
    End If

    g_GF_NextPollAt = 0
    On Error GoTo 0
End Sub

Private Function GF_PumpProcedureName() As String
    GF_PumpProcedureName = _
        "'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!GoogleFlights_AsyncPump"
End Function

Private Sub GF_CompleteAndClose()
    GF_CancelScheduledPump
    GF_ReleaseBrowser
End Sub

Private Sub GF_ReleaseBrowser()
    On Error Resume Next

    If Not g_GF_Browser Is Nothing Then
        Set g_GF_Browser.InheritanceWebDriverBiDiMode.BiDiEvents = Nothing
        g_GF_Browser.InheritanceWebDriverBiDiMode.quit
    End If

    Set g_GF_Browser = Nothing
    g_GF_CommandId = 0
    g_GF_Phase = GFPhaseIdle
    g_GF_Active = False
    g_GF_PumpRunning = False
    g_GF_NextPollAt = 0
    g_GF_PhaseStartedAt = 0
    g_GF_CommandCompleted = False
    g_GF_NetworkPattern = vbNullString
    g_GF_NetworkMatched = False

    On Error GoTo 0
End Sub

'***************************************************************************************************
'* Validates both ordinary BiDi command responses and script.evaluate responses.
'***************************************************************************************************
Private Sub GF_ValidateCommandResponse( _
    ByRef rawJson As String, _
    ByVal isScriptEvaluate As Boolean)

    Const PROC As String = _
        "Demo_GoogleFlights_Async.GF_ValidateCommandResponse"

    Dim envelope As BiDiCDPJson
    Set envelope = BiDiCDPJson.Parse(rawJson)

    If envelope Is Nothing Then
        Err.Raise vbObjectError + 2410, PROC, _
            "The asynchronous BiDi response could not be parsed."
    End If

    If envelope.ExistsKey("error") Then
        Dim bidiError As String
        bidiError = envelope.StringKey("error")

        If envelope.ExistsKey("message") Then
            If Len(bidiError) > 0 Then bidiError = bidiError & ": "
            bidiError = bidiError & envelope.StringKey("message")
        End If

        If Len(bidiError) = 0 Then bidiError = "Unknown BiDi command error."

        Err.Raise vbObjectError + 2411, PROC, bidiError
    End If

    If Not envelope.ExistsKey("result") Then
        Err.Raise vbObjectError + 2412, PROC, _
            "The asynchronous BiDi response has no result member."
    End If

    If Not isScriptEvaluate Then Exit Sub

    Dim evaluateResult As BiDiCDPJson
    Set evaluateResult = envelope.NodeKey("result")

    If evaluateResult Is Nothing Then
        Err.Raise vbObjectError + 2413, PROC, _
            "script.evaluate returned no result object."
    End If

    If LCase$(evaluateResult.StringKey("type")) <> "success" Then
        Dim detail As String
        Dim exceptionDetails As BiDiCDPJson

        If evaluateResult.ExistsKey("exceptionDetails") Then
            Set exceptionDetails = _
                evaluateResult.NodeKey("exceptionDetails")

            If Not exceptionDetails Is Nothing Then
                detail = exceptionDetails.StringKey("text")
            End If
        End If

        If Len(detail) = 0 Then detail = "Unknown JavaScript exception."

        Err.Raise vbObjectError + 2414, PROC, detail
    End If

    Dim remoteValue As BiDiCDPJson
    Set remoteValue = evaluateResult.NodeKey("result")

    If remoteValue Is Nothing Then
        Err.Raise vbObjectError + 2415, PROC, _
            "script.evaluate returned no remote value."
    End If

    Dim valueText As String
    valueText = remoteValue.StringKey("value")

    If Left$(valueText, 3) <> "OK:" Then
        Err.Raise vbObjectError + 2416, PROC, _
            "Unexpected script completion result: " & valueText
    End If
End Sub

'================================================================================
' Browser-side JavaScript expressions
'================================================================================

Private Function GF_BuildTripTypeExpression() As String
    Dim js As String

    js = GF_JsPrelude
    js = js & "const combo=await waitXPath(" & _
              GF_JsString("(//div[@role='combobox'])[1]") & _
              ",15000,true);"
    js = js & "clickNode(combo);"
    js = js & "const opt=await waitXPath(" & _
              GF_JsString("//*[@role='option' and contains(., 'One way')]") & _
              ",15000,true);"
    js = js & "clickNode(opt);"
    js = js & "await sleep(300);"
    js = js & "return 'OK:TRIP_TYPE';"
    js = js & "})()"

    GF_BuildTripTypeExpression = js
End Function

Private Function GF_BuildDepartureExpression() As String
    Dim js As String

    js = GF_JsPrelude
    js = js & "await inputKeys(" & _
              GF_JsString("(//input[contains(@aria-label, 'Where from')])[last()]") & _
              ",'Sapporo',15000);"
    js = js & "const suggestion=await waitXPath(" & _
              GF_JsString("//*[@role='listbox' and not(@aria-hidden='true')]//li[@role='option' and contains(@aria-label, 'Sapporo')][1]") & _
              ",15000,true);"
    js = js & "clickNode(suggestion);"
    js = js & "await waitXPath(" & _
              GF_JsString("(//input[contains(@aria-label, 'Where to')])[last()]") & _
              ",15000,true);"
    js = js & "return 'OK:DEPARTURE';"
    js = js & "})()"

    GF_BuildDepartureExpression = js
End Function

Private Function GF_BuildDestinationExpression() As String
    Dim js As String

    js = GF_JsPrelude
    js = js & "await inputKeys(" & _
              GF_JsString("(//input[contains(@aria-label, 'Where to')])[last()]") & _
              ",'Paris',15000);"
    js = js & "const suggestion=await waitXPath(" & _
              GF_JsString("//*[@role='listbox' and not(@aria-hidden='true')]//li[@role='option' and contains(@aria-label, 'Paris')][1]") & _
              ",15000,true);"
    js = js & "clickNode(suggestion);"
    js = js & "await waitXPath(" & _
              GF_JsString("//input[@aria-label='Departure']") & _
              ",15000,true);"
    js = js & "return 'OK:DESTINATION';"
    js = js & "})()"

    GF_BuildDestinationExpression = js
End Function

Private Function GF_BuildCalendarOpenExpression() As String
    Dim js As String

    js = GF_JsPrelude
    js = js & "const dep=await waitXPath(" & _
              GF_JsString("//input[@aria-label='Departure']") & _
              ",15000,true);"
    js = js & "clickNode(dep);"

    'Fare cells are the DOM-side completion signal paired with GetCalendarPicker.
    js = js & "await waitXPath(" & _
              GF_JsString("//div[@data-gs]") & _
              ",30000,true);"
    js = js & "return 'OK:CALENDAR_OPEN';"
    js = js & "})()"

    GF_BuildCalendarOpenExpression = js
End Function

Private Function GF_BuildDateSelectExpression() As String
    Dim js As String

    js = GF_JsPrelude
    js = js & "const dateButton=await waitXPath(" & _
              GF_JsString("(//div[@role='gridcell' and @aria-hidden='false'])[8]//div[@role='button']") & _
              ",30000,true);"
    js = js & "clickNode(dateButton);"

    'Keep this click in its own BiDi command. A short browser-side settle gives
    'Google Flights time to apply the date before the next command resolves Done
    'again from the current DOM.
    js = js & "await sleep(800);"
    js = js & "return 'OK:DATE_SELECT';"
    js = js & "})()"

    GF_BuildDateSelectExpression = js
End Function

Private Function GF_BuildCalendarDoneExpression() As String
    Dim js As String

    js = GF_JsPrelude

    'Use exactly the same XPath as the original BiDi Main08. waitXPath now
    'examines every XPath match and returns the first visible node.
    js = js & "const doneXPath=" & _
              GF_JsString("//button[contains(., 'Done')]") & ";"
    js = js & "const doneButton=await waitXPath(doneXPath,15000,true);"
    js = js & "clickNode(doneButton);"

    'Confirm that the close action was accepted before the Search phase.
    js = js & "await waitXPathAbsent(doneXPath,15000,true);"
    js = js & "return 'OK:CALENDAR_DONE';"
    js = js & "})()"

    GF_BuildCalendarDoneExpression = js
End Function

Private Function GF_BuildSearchExpression() As String
    Dim js As String

    js = GF_JsPrelude
    js = js & "const searchButton=await waitXPath(" & _
              GF_JsString("//button[@aria-label='Search']") & _
              ",15000,true);"
    js = js & "clickNode(searchButton);"
    js = js & "return 'OK:SEARCH';"
    js = js & "})()"

    GF_BuildSearchExpression = js
End Function

'***************************************************************************************************
'* Shared browser-side helper library.
'*
'* inputKeys is a direct adaptation of the BiDi edition's Google Flights input_keys logic:
'*   Phase 0: click and lock onto the final stable activeElement after Wiz replacements
'*   Phase 1: clear the existing value
'*   Phase 2: per-character execCommand('insertText') with activeElement retracking
'*   Phase 3: validate the final value length
'***************************************************************************************************
Private Function GF_JsPrelude() As String
    Dim js As String

    js = "(async()=>{"
    js = js & "const sleep=ms=>new Promise(r=>setTimeout(r,ms));"

    js = js & "const visible=e=>{if(!e||!e.isConnected)return false;" & _
              "const s=getComputedStyle(e),r=e.getBoundingClientRect();" & _
              "return s.display!=='none'&&s.visibility!=='hidden'&&" & _
              "parseFloat(s.opacity||'1')>0&&" & _
              "(r.width>0||r.height>0||e.getClientRects().length>0);};"

    'Search all XPath matches. A single-node XPath result can repeatedly return a
    'hidden stale SPA node while a later matching node is visibly active.
    js = js & "const byXPath=(xp,needVisible)=>{try{" & _
              "const r=document.evaluate(xp,document,null," & _
              "XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,null);" & _
              "for(let i=0;i<r.snapshotLength;i++){const n=r.snapshotItem(i);" & _
              "if(!needVisible||visible(n))return n;}return null;" & _
              "}catch(e){throw new Error('Invalid XPath: '+xp+' / '+e.message);}};"

    js = js & "const waitXPath=async(xp,timeout,needVisible)=>{" & _
              "const started=performance.now();for(;;){" & _
              "const e=byXPath(xp,needVisible);if(e)return e;" & _
              "if(performance.now()-started>=timeout)throw new Error(" & _
              "'XPath timeout: '+xp);await sleep(50);}};"

    js = js & "const waitXPathAbsent=async(xp,timeout,visibleOnly)=>{" & _
              "const started=performance.now();for(;;){" & _
              "if(!byXPath(xp,visibleOnly))return;" & _
              "if(performance.now()-started>=timeout)throw new Error(" & _
              "'XPath remained present: '+xp);await sleep(50);}};"

    js = js & "const clickNode=e=>{" & _
              "if(!e||!e.isConnected)throw new Error('Click target is detached.');" & _
              "e.scrollIntoView({block:'center',inline:'center'});" & _
              "try{e.focus({preventScroll:true});}catch(_){try{e.focus();}catch(__){}}" & _
              "if(typeof e.click!=='function')throw new Error('Target has no click().');" & _
              "e.click();};"

    js = js & "const inputKeys=async(xp,v,timeout)=>{"

    'Phase 0: activate the trigger and lock onto the final stable Wiz input.
    js = js & "let e=await waitXPath(xp,timeout,true),o=e;"
    js = js & "e.scrollIntoView({block:'center',inline:'center'});e.click();e.focus();"
    js = js & "let prev=null,stb=0;"
    js = js & "for(let t=Date.now();Date.now()-t<3000;){let a=document.activeElement;"
    js = js & "if(a&&a.tagName==='INPUT'&&(a.type==='text'||a.type===''||!a.type)&&"
    js = js & "typeof a.selectionStart==='number'){"
    js = js & "if(a===prev){stb++;if(stb>=2){e=a;break;}}else{prev=a;stb=1;}}"
    js = js & "else{prev=null;stb=0;}await sleep(15);}"
    js = js & "if(e===o&&document.activeElement!==e)e.focus();"

    'Phase 1: clear the current value.
    js = js & "(e.select?e.select():document.execCommand('selectAll'));"
    js = js & "await sleep(80);document.execCommand('delete');await sleep(80);"
    js = js & "if(e.value){if(e.select)e.select();document.execCommand('forwardDelete');"
    js = js & "await sleep(80);}"
    js = js & "if(e.value){e.value='';e.dispatchEvent(new Event('input',{bubbles:true}));"
    js = js & "await sleep(80);}await sleep(40);"

    'Phase 2: insert one character at a time while tracking activeElement replacement.
    js = js & "for(let c of v){let s=(typeof e.selectionStart==='number'?e.selectionStart:0);"
    js = js & "document.execCommand('insertText',false,c);"
    js = js & "for(let k=0;k<80;k++){let a=document.activeElement;"
    js = js & "if(a&&a!==e&&a.tagName==='INPUT'&&(a.type==='text'||!a.type))e=a;"
    js = js & "if(document.activeElement===e&&typeof e.selectionStart==='number'&&"
    js = js & "e.selectionStart>=s+1)break;await sleep(10);}}"

    'Phase 3: validate and return.
    js = js & "if(v.length>0&&(!e.value||e.value.length<v.length))"
    js = js & "throw new Error('Wiz Input Validation Failed: '+e.value);return e;};"

    GF_JsPrelude = js
End Function

Private Function GF_JsString(ByVal value As String) As String
    Dim escaped As String

    escaped = Replace(value, "\", "\\")
    escaped = Replace(escaped, """", "\""")
    escaped = Replace(escaped, vbCr, "\r")
    escaped = Replace(escaped, vbLf, "\n")
    escaped = Replace(escaped, vbTab, "\t")

    GF_JsString = """" & escaped & """"
End Function


