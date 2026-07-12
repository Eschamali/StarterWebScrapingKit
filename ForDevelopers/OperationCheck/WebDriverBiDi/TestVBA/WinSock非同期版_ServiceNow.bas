Attribute VB_Name = "WinSock非同期版_ServiceNow"
Option Explicit

'================================================================================
' ServiceNow Main07 - Kit ExecuteBiDiAsync / WinSock asynchronous version 1
'
' Original BiDi Main07 sequence:
'   1. Register the consent-banner AutoClicker before navigation.
'   2. Navigate to https://developer.servicenow.com/
'   3. Arm network signal "metadata/application".
'   4. Click "#utility-sign-in button" across Shadow DOM boundaries.
'   5. Enter "aaa" into //input[@id='username'].
'
' Kit implementation:
'   - Every BiDi command is submitted with ExecuteBiDiAsync.
'   - Application.OnTime invokes a short non-blocking receive/state pump.
'   - A preload script forces future closed ShadowRoots to open mode.
'   - A second preload script auto-clicks the TrustArc consent button.
'   - Shadow click completion requires both:
'         script.evaluate command completion
'         network.responseCompleted containing "metadata/application"
'   - Immediate-window logs use [hh:mm:ss.fff] timestamps.
'
' Entry point:
'   BiDiによる冒険の始まり_ServiceNow非同期版
'
' Cancellation:
'   ServiceNow_AsyncCancel
'================================================================================

Private Const SN_POLL_SECONDS As Long = 1
Private Const SN_PHASE_TIMEOUT_SECONDS As Long = 120
Private Const SN_URL As String = "https://developer.servicenow.com/"

Private Enum ServiceNowAsyncPhase
    SNPhaseIdle = 0
    SNPhaseWaitSubscribe = 1
    SNPhaseWaitShadowUnlocker = 2
    SNPhaseWaitConsentAutoClicker = 3
    SNPhaseWaitNavigate = 4
    SNPhaseWaitShadowSignIn = 5
    SNPhaseWaitUsername = 6
End Enum

Private g_SN_Browser As WebDriverBiDiContext
Private g_SN_CommandId As Long
Private g_SN_Phase As ServiceNowAsyncPhase
Private g_SN_Active As Boolean
Private g_SN_PumpRunning As Boolean
Private g_SN_NextPollAt As Date
Private g_SN_PhaseStartedAt As Date
Private g_SN_CommandCompleted As Boolean
Private g_SN_NetworkPattern As String
Private g_SN_NetworkMatched As Boolean

'***************************************************************************************************
'* Starts the fully asynchronous ServiceNow sequence.
'*
'* The procedure returns after session.subscribe is submitted. Browser work then continues through
'* Application.OnTime + TakeEvents + TakeResultBiDi.
'***************************************************************************************************
Public Sub BiDiによる冒険の始まり_ServiceNow非同期版()
    Const PROC As String = _
        "Demo_ServiceNow_Async.BiDiによる冒険の始まり_ServiceNow非同期版"

    If g_SN_Active Then
        Err.Raise vbObjectError + 2500, PROC, _
            "ServiceNow asynchronous processing is already running."
    End If

    On Error GoTo ErrorHandler

    Set g_SN_Browser = 設定シートからのBiDi起動ForTab

    '-----WebSocketルート-----
'    Dim UserName As String
'    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_CDP.SetupWebSocketMode")
'
'    Dim WebSocketBiDi As New CDPCoreViaWebSocket
'    WebSocketBiDi.ConnectCDP UserName, "/devtools/browser/c0b11c73-c215-4515-b90a-4a9c231c6021"
'
'    Dim b As New WebDriverBiDiMode
'    b.reattach UserName, WebSocketMode:=WebSocketBiDi
'    Set g_SN_Browser = b.newTab(setMain:=True)
    '-------------------------

    'Enable event accumulation before subscription.
    Set g_SN_Browser.InheritanceWebDriverBiDiMode.BiDiEvents = New Dictionary

    g_SN_Active = True

    SN_SubmitSubscribe
    SN_SchedulePump
    Exit Sub

ErrorHandler:
    Dim savedNumber As Long
    Dim savedSource As String
    Dim savedDescription As String

    savedNumber = Err.Number
    savedSource = Err.Source
    savedDescription = Err.Description

    SN_ReleaseBrowser
    Err.Raise savedNumber, savedSource, savedDescription
End Sub

'***************************************************************************************************
'* Cooperative asynchronous receive/state pump.
'*
'* Each invocation:
'*   1. drains currently available WinSock messages once,
'*   2. checks a one-shot network gate if armed,
'*   3. checks the current command ID once,
'*   4. advances one state or returns control to Excel.
'***************************************************************************************************
Public Sub ServiceNow_AsyncPump()
    If Not g_SN_Active Then Exit Sub
    If g_SN_PumpRunning Then Exit Sub

    g_SN_PumpRunning = True
    g_SN_NextPollAt = 0

    On Error GoTo ErrorHandler

    'One non-blocking WinSock receive/drain pass.
    g_SN_Browser.InheritanceWebDriverBiDiMode.TakeEvents True

    If Len(g_SN_NetworkPattern) > 0 Then
        If Not g_SN_NetworkMatched Then
            g_SN_NetworkMatched = SN_HasNetworkPattern(g_SN_NetworkPattern)

            If g_SN_NetworkMatched Then
                SN_Log "ServiceNow network gate matched: " & _
                       g_SN_NetworkPattern
            End If
        End If
    End If

    If Not g_SN_CommandCompleted Then
        Dim rawJson As String

        rawJson = _
            g_SN_Browser.InheritanceWebDriverBiDiMode.TakeResultBiDi( _
                g_SN_CommandId)

        If StrPtr(rawJson) Then
            SN_ValidateCommandResponse rawJson, _
                SN_IsScriptEvaluatePhase(g_SN_Phase)

            g_SN_CommandCompleted = True
        End If
    End If

    If SN_PhaseTimedOut Then
        Err.Raise vbObjectError + 2501, _
            "Demo_ServiceNow_Async.ServiceNow_AsyncPump", _
            "Timed out in phase: " & SN_PhaseName(g_SN_Phase) & _
            IIf(Len(g_SN_NetworkPattern) > 0 And _
                Not g_SN_NetworkMatched, _
                " / missing network signal: " & _
                g_SN_NetworkPattern, _
                vbNullString)
    End If

    'The current step is complete only after both the command and any network gate finish.
    If Not g_SN_CommandCompleted Then
        g_SN_PumpRunning = False
        SN_SchedulePump
        Exit Sub
    End If

    If Len(g_SN_NetworkPattern) > 0 And _
       Not g_SN_NetworkMatched Then

        g_SN_PumpRunning = False
        SN_SchedulePump
        Exit Sub
    End If

    Select Case g_SN_Phase
        Case SNPhaseWaitSubscribe
            SN_Log "ServiceNow async completed: session.subscribe"
            SN_SubmitShadowUnlocker

        Case SNPhaseWaitShadowUnlocker
            SN_Log "ServiceNow async completed: Global Shadow DOM Unlocker"
            SN_SubmitConsentAutoClicker

        Case SNPhaseWaitConsentAutoClicker
            SN_Log "ServiceNow async completed: consent AutoClicker registration"
            SN_SubmitNavigate

        Case SNPhaseWaitNavigate
            SN_Log "ServiceNow async completed: navigation"
            SN_ResetNetworkGate "metadata/application"
            SN_SubmitScript SNPhaseWaitShadowSignIn, _
                SN_BuildShadowSignInExpression

        Case SNPhaseWaitShadowSignIn
            SN_Log "ServiceNow async completed: Shadow DOM sign-in click"
            SN_SubmitScript SNPhaseWaitUsername, _
                SN_BuildUsernameExpression

        Case SNPhaseWaitUsername
            SN_Log "ServiceNow async completed: username input"
            SN_Log "ServiceNow asynchronous processing completed successfully."

            g_SN_PumpRunning = False
            SN_CompleteAndClose

            MsgBox "ServiceNow Main07 Test Completed", _
                   vbInformation, _
                   "ServiceNow Async"
            Exit Sub

        Case Else
            Err.Raise vbObjectError + 2502, _
                "Demo_ServiceNow_Async.ServiceNow_AsyncPump", _
                "Unknown ServiceNow asynchronous phase."
    End Select

    g_SN_PumpRunning = False
    SN_SchedulePump
    Exit Sub

ErrorHandler:
    Dim savedNumber As Long
    Dim savedSource As String
    Dim savedDescription As String
    Dim failedPhase As ServiceNowAsyncPhase

    savedNumber = Err.Number
    savedSource = Err.Source
    savedDescription = Err.Description
    failedPhase = g_SN_Phase

    SN_Log "ServiceNow asynchronous processing failed: " & _
           "phase=" & SN_PhaseName(failedPhase) & _
           " / Error " & CStr(savedNumber) & _
           " / " & savedSource & _
           " / " & savedDescription

    g_SN_PumpRunning = False
    SN_CancelScheduledPump
    SN_ReleaseBrowser

    MsgBox _
        "ServiceNow asynchronous processing failed." & vbCrLf & vbCrLf & _
        "Phase: " & SN_PhaseName(failedPhase) & vbCrLf & _
        "Error " & savedNumber & vbCrLf & _
        savedSource & vbCrLf & _
        savedDescription, _
        vbCritical, _
        "ServiceNow Async"
End Sub

'***************************************************************************************************
'* Explicit user cancellation.
'***************************************************************************************************
Public Sub ServiceNow_AsyncCancel()
    SN_CancelScheduledPump
    SN_ReleaseBrowser
    SN_Log "ServiceNow asynchronous processing was cancelled."
End Sub

'================================================================================
' Command submission
'================================================================================

Private Sub SN_SubmitSubscribe()
    Dim subscribeParams As Dictionary
    Set subscribeParams = New Dictionary

    Dim eventsArray As Collection
    Set eventsArray = New Collection
    eventsArray.Add "network.responseCompleted"

    subscribeParams.Add "events", eventsArray

    'Limit the high-volume network event stream to the controlled tab.
    Dim contextsArray As Collection
    Set contextsArray = New Collection
    contextsArray.Add g_SN_Browser.context

    subscribeParams.Add "contexts", contextsArray

    g_SN_CommandId = _
        g_SN_Browser.InheritanceWebDriverBiDiMode.ExecuteBiDiAsync( _
            "session.subscribe", _
            subscribeParams, _
            True)

    SN_BeginPhase SNPhaseWaitSubscribe
    SN_Log "ServiceNow async submitted: session.subscribe / ID=" & _
           g_SN_CommandId
End Sub

Private Sub SN_SubmitShadowUnlocker()
    Dim paramsBiDi As Dictionary
    Set paramsBiDi = New Dictionary

    paramsBiDi.Add "functionDeclaration", _
        SN_BuildShadowUnlockerPreload

    g_SN_CommandId = _
        g_SN_Browser.InheritanceWebDriverBiDiMode.ExecuteBiDiAsync( _
            "script.addPreloadScript", _
            paramsBiDi, _
            True)

    SN_BeginPhase SNPhaseWaitShadowUnlocker
    SN_Log "ServiceNow async submitted: Global Shadow DOM Unlocker / ID=" & _
           g_SN_CommandId
End Sub

Private Sub SN_SubmitConsentAutoClicker()
    Dim paramsBiDi As Dictionary
    Set paramsBiDi = New Dictionary

    paramsBiDi.Add "functionDeclaration", _
        SN_BuildConsentAutoClickerPreload

    g_SN_CommandId = _
        g_SN_Browser.InheritanceWebDriverBiDiMode.ExecuteBiDiAsync( _
            "script.addPreloadScript", _
            paramsBiDi, _
            True)

    SN_BeginPhase SNPhaseWaitConsentAutoClicker
    SN_Log "ServiceNow async submitted: consent AutoClicker / ID=" & _
           g_SN_CommandId
End Sub

Private Sub SN_SubmitNavigate()
    Dim navParams As Dictionary
    Set navParams = New Dictionary

    navParams.Add "url", SN_URL
    navParams.Add "wait", "complete"

    g_SN_CommandId = _
        g_SN_Browser.ExecuteBiDiAsync( _
            "browsingContext.navigate", _
            navParams, _
            True)

    SN_BeginPhase SNPhaseWaitNavigate
    SN_Log "ServiceNow async submitted: navigation / ID=" & _
           g_SN_CommandId
End Sub

Private Sub SN_SubmitScript( _
    ByVal nextPhase As ServiceNowAsyncPhase, _
    ByVal expression As String)

    Const PROC As String = _
        "Demo_ServiceNow_Async.SN_SubmitScript"

    Dim paramsBiDi As Dictionary
    Set paramsBiDi = New Dictionary

    paramsBiDi.Add "expression", expression

    Dim target As Dictionary
    Set target = New Dictionary
    target.Add "context", g_SN_Browser.context

    paramsBiDi.Add "target", target
    paramsBiDi.Add "awaitPromise", True

    'script.evaluate requires context under target.context.
    g_SN_CommandId = _
        g_SN_Browser.InheritanceWebDriverBiDiMode.ExecuteBiDiAsync( _
            "script.evaluate", _
            paramsBiDi, _
            True)

    If g_SN_CommandId <= 0 Then
        Err.Raise vbObjectError + 2503, PROC, _
            "ExecuteBiDiAsync did not return a valid command ID."
    End If

    SN_BeginPhase nextPhase

    SN_Log "ServiceNow async submitted: " & _
           SN_PhaseName(nextPhase) & _
           " / ID=" & g_SN_CommandId
End Sub

'================================================================================
' State, event and lifecycle helpers
'================================================================================

Private Sub SN_BeginPhase(ByVal phase As ServiceNowAsyncPhase)
    g_SN_Phase = phase
    g_SN_PhaseStartedAt = Now
    g_SN_CommandCompleted = False

    'The network gate is retained only for the Shadow sign-in action.
    If phase <> SNPhaseWaitShadowSignIn Then
        g_SN_NetworkPattern = vbNullString
        g_SN_NetworkMatched = False
    End If
End Sub

Private Sub SN_ResetNetworkGate(ByVal urlPattern As String)
    g_SN_NetworkPattern = urlPattern
    g_SN_NetworkMatched = False

    'One-shot history: an old matching request must not satisfy a new action.
    Set g_SN_Browser.InheritanceWebDriverBiDiMode.BiDiEvents = _
        New Dictionary
End Sub

Private Function SN_HasNetworkPattern( _
    ByVal urlPattern As String) As Boolean

    On Error GoTo SafeExit

    Dim eventsRoot As Dictionary
    Set eventsRoot = _
        g_SN_Browser.InheritanceWebDriverBiDiMode.BiDiEvents

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
            SN_HasNetworkPattern = True
            Exit Function
        End If
    Next eventItem

SafeExit:
End Function

Private Function SN_IsScriptEvaluatePhase( _
    ByVal phase As ServiceNowAsyncPhase) As Boolean

    SN_IsScriptEvaluatePhase = _
        (phase = SNPhaseWaitShadowSignIn Or _
         phase = SNPhaseWaitUsername)
End Function

Private Function SN_PhaseTimedOut() As Boolean
    If g_SN_PhaseStartedAt = 0 Then Exit Function

    SN_PhaseTimedOut = _
        (DateDiff("s", g_SN_PhaseStartedAt, Now) > _
         SN_PHASE_TIMEOUT_SECONDS)
End Function

Private Function SN_PhaseName( _
    ByVal phase As ServiceNowAsyncPhase) As String

    Select Case phase
        Case SNPhaseIdle
            SN_PhaseName = "Idle"

        Case SNPhaseWaitSubscribe
            SN_PhaseName = "Subscribe"

        Case SNPhaseWaitShadowUnlocker
            SN_PhaseName = "Shadow Unlocker"

        Case SNPhaseWaitConsentAutoClicker
            SN_PhaseName = "Consent AutoClicker"

        Case SNPhaseWaitNavigate
            SN_PhaseName = "Navigate"

        Case SNPhaseWaitShadowSignIn
            SN_PhaseName = "Shadow sign-in"

        Case SNPhaseWaitUsername
            SN_PhaseName = "Username input"

        Case Else
            SN_PhaseName = "Unknown"
    End Select
End Function

Private Sub SN_SchedulePump()
    If Not g_SN_Active Then Exit Sub

    g_SN_NextPollAt = _
        Now + TimeSerial(0, 0, SN_POLL_SECONDS)

    Application.OnTime _
        EarliestTime:=g_SN_NextPollAt, _
        Procedure:=SN_PumpProcedureName, _
        Schedule:=True
End Sub

Private Sub SN_CancelScheduledPump()
    On Error Resume Next

    If g_SN_NextPollAt <> 0 Then
        Application.OnTime _
            EarliestTime:=g_SN_NextPollAt, _
            Procedure:=SN_PumpProcedureName, _
            Schedule:=False
    End If

    g_SN_NextPollAt = 0
    On Error GoTo 0
End Sub

Private Function SN_PumpProcedureName() As String
    SN_PumpProcedureName = _
        "'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!ServiceNow_AsyncPump"
End Function

Private Sub SN_CompleteAndClose()
    SN_CancelScheduledPump
    SN_ReleaseBrowser
End Sub

Private Sub SN_ReleaseBrowser()
    On Error Resume Next

    If Not g_SN_Browser Is Nothing Then
        Set g_SN_Browser.InheritanceWebDriverBiDiMode.BiDiEvents = _
            Nothing

        g_SN_Browser.InheritanceWebDriverBiDiMode.quit
    End If

    Set g_SN_Browser = Nothing

    g_SN_CommandId = 0
    g_SN_Phase = SNPhaseIdle
    g_SN_Active = False
    g_SN_PumpRunning = False
    g_SN_NextPollAt = 0
    g_SN_PhaseStartedAt = 0
    g_SN_CommandCompleted = False
    g_SN_NetworkPattern = vbNullString
    g_SN_NetworkMatched = False

    On Error GoTo 0
End Sub

'================================================================================
' Response validation
'================================================================================

Private Sub SN_ValidateCommandResponse( _
    ByRef rawJson As String, _
    ByVal isScriptEvaluate As Boolean)

    Const PROC As String = _
        "Demo_ServiceNow_Async.SN_ValidateCommandResponse"

    Dim envelope As BiDiCDPJson
    Set envelope = BiDiCDPJson.Parse(rawJson)

    If envelope Is Nothing Then
        Err.Raise vbObjectError + 2510, PROC, _
            "The asynchronous BiDi response could not be parsed."
    End If

    If envelope.ExistsKey("error") Then
        Dim bidiError As String
        bidiError = envelope.StringKey("error")

        If envelope.ExistsKey("message") Then
            If Len(bidiError) > 0 Then bidiError = bidiError & ": "
            bidiError = bidiError & envelope.StringKey("message")
        End If

        If Len(bidiError) = 0 Then
            bidiError = "Unknown BiDi command error."
        End If

        Err.Raise vbObjectError + 2511, PROC, bidiError
    End If

    If Not envelope.ExistsKey("result") Then
        Err.Raise vbObjectError + 2512, PROC, _
            "The asynchronous BiDi response has no result member."
    End If

    If Not isScriptEvaluate Then Exit Sub

    Dim evaluateResult As BiDiCDPJson
    Set evaluateResult = envelope.NodeKey("result")

    If evaluateResult Is Nothing Then
        Err.Raise vbObjectError + 2513, PROC, _
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

        If Len(detail) = 0 Then
            detail = "Unknown JavaScript exception."
        End If

        Err.Raise vbObjectError + 2514, PROC, detail
    End If

    Dim remoteValue As BiDiCDPJson
    Set remoteValue = evaluateResult.NodeKey("result")

    If remoteValue Is Nothing Then
        Err.Raise vbObjectError + 2515, PROC, _
            "script.evaluate returned no remote value."
    End If

    Dim valueText As String
    valueText = remoteValue.StringKey("value")

    If Left$(valueText, 3) <> "OK:" Then
        Err.Raise vbObjectError + 2516, PROC, _
            "Unexpected script completion result: " & valueText
    End If
End Sub

'================================================================================
' Browser-side preload scripts
'================================================================================

Private Function SN_BuildShadowUnlockerPreload() As String
    Dim js As String

    js = "function(){"
    js = js & "if(Element.prototype._attachShadow)return;"
    js = js & "Element.prototype._attachShadow="
    js = js & "Element.prototype.attachShadow;"
    js = js & "Element.prototype.attachShadow=function(options){"
    js = js & "return this._attachShadow(Object.assign({},"
    js = js & "options||{},{mode:'open'}));};"
    js = js & "}"

    SN_BuildShadowUnlockerPreload = js
End Function

Private Function SN_BuildConsentAutoClickerPreload() As String
    Dim js As String
    Dim consentXPath As String

    consentXPath = "//button[@id='truste-consent-button']"

    js = "function(){"
    js = js & "const x=" & SN_JsString(consentXPath) & ";"
    js = js & "const timeout=30000,start=Date.now();"

    js = js & "const visible=e=>{"
    js = js & "if(!e||!e.isConnected)return false;"
    js = js & "const s=getComputedStyle(e),r=e.getBoundingClientRect();"
    js = js & "return s.display!=='none'&&"
    js = js & "s.visibility!=='hidden'&&"
    js = js & "parseFloat(s.opacity||'1')>0&&"
    js = js & "(r.width>0||r.height>0||"
    js = js & "e.getClientRects().length>0);};"

    js = js & "const find=()=>{try{return document.evaluate("
    js = js & "x,document,null,"
    js = js & "XPathResult.FIRST_ORDERED_NODE_TYPE,null"
    js = js & ").singleNodeValue;}catch(_){return null;}};"

    js = js & "const clickIfReady=()=>{"
    js = js & "const e=find();"
    js = js & "if(e&&visible(e)){e.click();"
    js = js & "console.log('Kit-AutoClicker: Target clicked');"
    js = js & "return true;}return false;};"

    js = js & "const startObserver=()=>{"
    js = js & "if(clickIfReady())return;"
    js = js & "const root=document.documentElement||document;"
    js = js & "const o=new MutationObserver(()=>{"
    js = js & "if(clickIfReady()||Date.now()-start>timeout)"
    js = js & "o.disconnect();});"
    js = js & "o.observe(root,{childList:true,subtree:true,"
    js = js & "attributes:true});};"

    js = js & "if(document.readyState==='loading'){"
    js = js & "document.addEventListener('DOMContentLoaded',"
    js = js & "startObserver,{once:true});"
    js = js & "}else{startObserver();}"
    js = js & "}"

    SN_BuildConsentAutoClickerPreload = js
End Function

'================================================================================
' Browser-side action expressions
'================================================================================

Private Function SN_BuildShadowSignInExpression() As String
    Dim js As String

    js = SN_JsPrelude

    'Equivalent to BiDi Main07:
    '   ExecuteShadowClick "#utility-sign-in button"
    js = js & "const selector='#utility-sign-in button';"

    js = js & "const findDeep=(root,sel)=>{"
    js = js & "let found=root.querySelector(sel);"
    js = js & "if(found)return found;"
    js = js & "for(const node of root.querySelectorAll('*')){"
    js = js & "if(node.shadowRoot){"
    js = js & "found=findDeep(node.shadowRoot,sel);"
    js = js & "if(found)return found;}}return null;};"

    js = js & "const started=performance.now();"
    js = js & "let signIn=null;"

    js = js & "for(;;){"
    js = js & "signIn=findDeep(document,selector);"
    js = js & "if(signIn&&visible(signIn))break;"
    js = js & "if(performance.now()-started>=30000)"
    js = js & "throw new Error("
    js = js & "'ShadowSearchError: '+selector);"
    js = js & "await sleep(100);}"

    js = js & "clickNode(signIn);"
    js = js & "return 'OK:SHADOW_SIGN_IN';"
    js = js & "})()"

    SN_BuildShadowSignInExpression = js
End Function

Private Function SN_BuildUsernameExpression() As String
    Dim js As String

    js = SN_JsPrelude

    'Equivalent to BiDi Main07:
    '   ExecuteInputValueByXPath "//input[@id='username']", "aaa"
    js = js & "const input=await inputKeys("
    js = js & SN_JsString("//input[@id='username']") & ","
    js = js & SN_JsString("aaa") & ",30000);"

    'Explicit final-value validation.
    js = js & "if(input.value!=='aaa')"
    js = js & "throw new Error("
    js = js & "'Username validation failed: '+input.value);"

    js = js & "return 'OK:USERNAME';"
    js = js & "})()"

    SN_BuildUsernameExpression = js
End Function

'***************************************************************************************************
'* Shared browser-side helper library.
'*
'* inputKeys mirrors the BiDi wrapper's default ExecuteInputValueByXPath key-events mode:
'*   Phase 0: click and lock onto a stable active input
'*   Phase 1: clear the current value
'*   Phase 2: per-character insertText with activeElement tracking
'*   Phase 3: validate the final value length
'***************************************************************************************************
Private Function SN_JsPrelude() As String
    Dim js As String

    js = "(async()=>{"
    js = js & "const sleep=ms=>new Promise("
    js = js & "r=>setTimeout(r,ms));"

    js = js & "const visible=e=>{"
    js = js & "if(!e||!e.isConnected)return false;"
    js = js & "const s=getComputedStyle(e),"
    js = js & "r=e.getBoundingClientRect();"
    js = js & "return s.display!=='none'&&"
    js = js & "s.visibility!=='hidden'&&"
    js = js & "parseFloat(s.opacity||'1')>0&&"
    js = js & "(r.width>0||r.height>0||"
    js = js & "e.getClientRects().length>0);};"

    'Search every XPath match and return the first acceptable live node.
    js = js & "const byXPath=(xp,needVisible)=>{try{"
    js = js & "const r=document.evaluate(xp,document,null,"
    js = js & "XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,null);"
    js = js & "for(let i=0;i<r.snapshotLength;i++){"
    js = js & "const n=r.snapshotItem(i);"
    js = js & "if(!needVisible||visible(n))return n;}"
    js = js & "return null;"
    js = js & "}catch(e){throw new Error("
    js = js & "'Invalid XPath: '+xp+' / '+e.message);}};"

    js = js & "const waitXPath=async("
    js = js & "xp,timeout,needVisible)=>{"
    js = js & "const started=performance.now();"
    js = js & "for(;;){"
    js = js & "const e=byXPath(xp,needVisible);"
    js = js & "if(e)return e;"
    js = js & "if(performance.now()-started>=timeout)"
    js = js & "throw new Error('XPath timeout: '+xp);"
    js = js & "await sleep(50);}};"

    js = js & "const clickNode=e=>{"
    js = js & "if(!e||!e.isConnected)"
    js = js & "throw new Error('Click target is detached.');"
    js = js & "e.scrollIntoView({block:'center',"
    js = js & "inline:'center'});"
    js = js & "try{e.focus({preventScroll:true});}"
    js = js & "catch(_){try{e.focus();}catch(__){}}"
    js = js & "if(typeof e.click!=='function')"
    js = js & "throw new Error('Target has no click().');"
    js = js & "e.click();};"

    js = js & "const inputKeys=async(xp,v,timeout)=>{"

    'Phase 0: activate and lock onto a stable final input.
    js = js & "let e=await waitXPath(xp,timeout,true),o=e;"
    js = js & "e.scrollIntoView({block:'center',"
    js = js & "inline:'center'});"
    js = js & "e.click();e.focus();"
    js = js & "let prev=null,stb=0;"
    js = js & "for(let t=Date.now();Date.now()-t<3000;){"
    js = js & "let a=document.activeElement;"
    js = js & "if(a&&a.tagName==='INPUT'&&"
    js = js & "(a.type==='text'||a.type===''||!a.type)&&"
    js = js & "typeof a.selectionStart==='number'){"
    js = js & "if(a===prev){stb++;"
    js = js & "if(stb>=2){e=a;break;}}"
    js = js & "else{prev=a;stb=1;}}"
    js = js & "else{prev=null;stb=0;}"
    js = js & "await sleep(15);}"
    js = js & "if(e===o&&document.activeElement!==e)e.focus();"

    'Phase 1: clear.
    js = js & "(e.select?e.select():"
    js = js & "document.execCommand('selectAll'));"
    js = js & "await sleep(80);"
    js = js & "document.execCommand('delete');"
    js = js & "await sleep(80);"
    js = js & "if(e.value){if(e.select)e.select();"
    js = js & "document.execCommand('forwardDelete');"
    js = js & "await sleep(80);}"
    js = js & "if(e.value){e.value='';"
    js = js & "e.dispatchEvent(new Event("
    js = js & "'input',{bubbles:true}));"
    js = js & "await sleep(80);}"
    js = js & "await sleep(40);"

    'Phase 2: one character at a time, tracking replacement inputs.
    js = js & "for(const c of v){"
    js = js & "let s=(typeof e.selectionStart==='number'?"
    js = js & "e.selectionStart:0);"
    js = js & "document.execCommand('insertText',false,c);"
    js = js & "for(let k=0;k<80;k++){"
    js = js & "let a=document.activeElement;"
    js = js & "if(a&&a!==e&&a.tagName==='INPUT'&&"
    js = js & "(a.type==='text'||!a.type))e=a;"
    js = js & "if(document.activeElement===e&&"
    js = js & "typeof e.selectionStart==='number'&&"
    js = js & "e.selectionStart>=s+1)break;"
    js = js & "await sleep(10);}}"

    'Phase 3: validate.
    js = js & "if(v.length>0&&"
    js = js & "(!e.value||e.value.length<v.length))"
    js = js & "throw new Error("
    js = js & "'Input Validation Failed: '+e.value);"
    js = js & "return e;};"

    SN_JsPrelude = js
End Function

'================================================================================
' Logging and string helpers
'================================================================================

Private Sub SN_Log(ByVal message As String)
    Debug.Print "[" & SN_LogTimestamp() & "] " & message
End Sub

Private Function SN_LogTimestamp() As String
    Dim currentTime As Date
    Dim timerValue As Double
    Dim milliseconds As Long

    currentTime = Now
    timerValue = Timer
    milliseconds = _
        CLng(Fix((timerValue - Fix(timerValue)) * 1000#))

    SN_LogTimestamp = _
        Format$(currentTime, "hh:nn:ss") & "." & _
        Format$(milliseconds, "000")
End Function

Private Function SN_JsString(ByVal value As String) As String
    Dim escaped As String

    escaped = Replace(value, "\", "\\")
    escaped = Replace(escaped, """", "\""")
    escaped = Replace(escaped, vbCr, "\r")
    escaped = Replace(escaped, vbLf, "\n")
    escaped = Replace(escaped, vbTab, "\t")

    SN_JsString = """" & escaped & """"
End Function


