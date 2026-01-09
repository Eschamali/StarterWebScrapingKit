VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
' Automating Chromium-Based Browsers with Chrome Dev Protocol API and VBA
'------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
'       Daniel Polak (SYS, Supporting Your Systems B.V.)
' Last Update :
'       31/12/25 Daniel Polak: - added BringWindowToTop and bringtoforeground
'       24/10/25 Daniel Polak: - sendMessage: added soft re-attach attempt when pipe error
'                              - snapPage: added CDP native screenshot capability which avoids use of external JavaScript library
'                              - getTab: improved tab selection
'                              - quit: avoid immediate hard kills, reduces race conditions, and makes shutdown behavior predictable and robust in non-responsive cases
'                              - searchNull: new faster version
'                              - use constants (ERR_TIMEOUT, ERR_PIPE, ERR_PROTOCOL, TIMEOUT_DEFAULT, TIMEOUT_NAVIGATE) instead of numbers for clarity
'       23/10/25 Daniel Polak: - replaced unnecessary use of FileSystemObject
'                              - added Option Explicit and added all required variable declarations
'                              - quit: delete the one and only cdpid.txt
'                              - Class_Initialize: prevent compilation errors
'                              - cleanUpSessions: solved bug
'                              - added DeleteCDPID so we use a function instead of repeating same code
'                              - 64-bit safety: window handles as LongPtr
'                              - start: added missing space before --homepage
'                              - jsAddLib: fixed bug in check whether JS library is loaded
'                              - getTab: fixed typo timeStart instead of timerStart and fixed bug 60000s instead of 60s
'                              - cleanUpSessions: only terminate automation-processes
'                              - snapPage: fixed bug where the all files in the Downloads folder were deleted instead of just the temporary screenshot
'                              - getBrowserPath: added fallback path when registry lookup fails
'       18/08/23 Long Vh:      - set core to Private i/o Public as it is meant to be Private.
'       06/06/23 Long Vh:      - applied recommendations by MitsOmitsaras for Access & Word compat.
'       26/05/23 Long Vh:      - made some tweaks for compatibility with Access & Word out-of-the-box
'       26/04/23 Long Vh:      - added .setMain for better tab management.
'       22/03/23 Long Vh:      - added many .getElementsBy methods.
'                              - added advanced methods to extend features of CDP apps.
'                              - changed deserialize/serialize methods to read from a temp file.
'                              - added .snapPage method for capturing an entire web screen.
'                              - added interactions with the new element class object.
'                              - many other changes to improve debug handlings.
' References  :
'       Microsoft Scripting Runtime
' Sources     :
'       Main Git: https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA
'       CDP Tutorials: https://github.com/cyrus-and/chrome-remote-interface/issues/282
'       CodeProject: https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA
'       Javascript Tutorials: https://www.w3schools.com/js/
'       Chrome CDP Documentation: https://chromedevtools.github.io/devtools-protocol/
'================================================================================================
Option Explicit

Private Const CDPVer                    As String = "3.0.0"
Private Const ArgRemoteDebuggingName    As String = "remote-debugging-pipe"
'===================================
' Win APIs Declarations
'===================================

Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)
    
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExW" ( _
    ByVal hwndParent As LongPtr, _
    ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As LongPtr, _
    ByVal lpszWindow As LongPtr) As LongPtr
    
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextW" ( _
    ByVal hWnd As LongPtr, _
    ByVal lpString As LongPtr, _
    ByVal nMaxCount As Long) As Long
    
Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nCmdShow As Long) As Boolean

Private Declare PtrSafe Function BringWindowToTop Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

'DJP 20241105 to solve problems with terminating sessions on terminal server
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" ( _
    ) As Long

Private Declare PtrSafe Function ProcessIdToSessionId Lib "kernel32" ( _
    ByVal dwProcessId As Long, _
    ByRef pSessionId As Long) As Long

'===================================
' Constants
'===================================

Public Enum WinState
    asHidden = 0            'Hides the window and activates another window.
    asNormal = 1            'Activates and displays a window. If the window is minimized or maximized, the system restores it to its original
    asMinimized = 2         'Activates the window and displays it as a minimized window.
    asMaximized = 3         'Activates the window and displays it as a maximized window.
    doShowNoActivate = 4    'Displays a window in its most recent size and position. This value is similar to asNormal, except that the window is not activated.
    doShow = 5              'Activates the window and displays it in its current size and position.
    doMinimize = 6          'Minimizes the specified window and activates the next top-level window in the Z order.
    doShowMinNoActivate = 7 'Displays the window as a minimized window. This value is similar to asMinimized, except the window is not activated.
    doShowNA = 8            'Displays the window in its current size and position. This value is similar to doShow, except that the window is not activated.
    doRestore = 9           'Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    doShowDefault = 10      'Sets the show state based on the WS_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
    doForceMin = 11         'Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread.
End Enum

Public Enum ReadyState      'Used for .wait method
    isLoading = 0           'equivalence of the browser's "loading" state
    isInteractive = 1       'equivalence of the browser's "interactive" state
    isComplete = 2          'equivalence of the browser's "complete" state
End Enum

Public Enum TargetgetTargetsType
    Page
    iFrame
End Enum


Private Const brChrome As String = "chrome.exe"
Private Const brEdge As String = "msedge.exe"
Private Const chromeWindowClass As String = "Chrome_WidgetWin_1"    'same window class for Edge

Private Const ERR_TIMEOUT As Long = 900
Private Const ERR_PIPE As Long = 901
Private Const ERR_PROTOCOL As Long = 902
Private Const TIMEOUT_DEFAULT As Single = 30
Private Const TIMEOUT_NAVIGATE As Single = 45

'===================================
' Global Variables
'===================================

Private core As CDPCore                 'Must be initialized for CDP engine
Private lngLastID As Long               'Every message sent to edge has an ID
Private brHWnd As LongPtr               'Browser object window handle
Private PieceData As String             'ストリーム受信で欠けたデータを一時的に蓄積するための変数です
Private brWID As String                 'Useful for windows repositioning
Private brType As String                'Browser type: either 'chrome' or 'edge'
Private brUser As String                'Browser user profile
Private brPath As String                'Browser default installation path
Private brID As String                  'The ID of the current CDPBrowser object. Useful when reading debugs of multiple tabs.
Private brName As String                'The focused tab name / window name of the active CDPBrowser
Private brURL As String                 'The focused tab url of the active CDPBrowser
Private brVer As String                 'The current browser version
Private brReAttach As Boolean           'Remember reAttach preference from start()
Private brCallString As String          'The initial call string used by .start
Private varIDs As Scripting.Dictionary  'Collection of all unique element IDs for CDPElement reference - Added 02/03/2023
Private sessionID As String             'The ID to the current session tab.
Private targetID As String              'The targetID of the current tab for InvokeMethod
Private thisApp As Object               'A dymanic variable used in CDPBrowser initialize procedure to determine if the VBA host is an Excel / Word / Access

Private seeRawSendMsgDbg As Boolean     'set to True if needing to see the raw json msg received from the browser for debugging
Private doPrintDbgMsg As Boolean        'イミディエイトに、デバッグメッセージを出すか
Private doLog As Boolean                'ログファイルにも、デバッグメッセージを残すか
Private logPath As String               'ログファイルの保存フォルダパス

'=============================================================================
' PUBLIC FUNCTIONS
'=============================================================================

Public Sub start(Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional cleanActive As Boolean = True, Optional reAttach As Boolean = True, _
    Optional userProfile As String, Optional addArgs As String)
'--------------------------------------------------------------------------------------------------------------
' This function initiates a new browser instance.
' By default, this function initializes chrome as the default choice.
' To init edge, parse 'edge' to the first parameter (eg. browser.start "edge").
' To open in App Mode, input the starting URL for appURL parameter as app mode cannot be open on a blank tab.
' If cleanActive is False, the method will skip cleaning currently open instances. This will save
' initialization time but if an instance is already running and not closed, pipe error will be triggered.
' If reAttach is True, the method will always try to reattach to the previous automation session if it is
' still active. If not, it will then move on to start a brand new session. If False, it will always open a new
' session regardless whether there is an on going one or not.
' To add additional arguments to the call string, use the addArguments parameter for instance:
' .start addArgs:="--disable-info-bar" (https://peter.sh/experiments/chromium-command-line-switches/)
' Note: not all arguments work the same for Chrome and Edge. Some may work for one but not the others.
' Updated: 22/03/23 Long Vh:      - Enhanced to dynamically retrieve installation path of the browser
'          26/03/23 Long Vh:      - Made some quality-of-life changes
'          25/05/25 Long Vh:      - set default profile to "CDP" so that it can work with the latest chrome update
'          23/10/25 Daniel Polak: - Added missing space before --homepage
'--------------------------------------------------------------------------------------------------------------

    Dim strCall As String
    Dim result As Long
    Dim timeStart As Single

    timeStart = Timer
    brReAttach = reAttach

   'Determine if already open -----------------------------------------------------
    If reAttach And isLive Then
        printParams     'Useful for checking
        Exit Sub
    End If
    printMsg "Reattachment aborted as reAttach=" & reAttach & " & isLive=" & isLive
    printMsg "New Browser Session Started", isHeader:=True
   '-------------------------------------------------------------------------------
                                   
   'Determine default browser parameters -------------------------------------------------------------------------------------
    Select Case Name
        Case "edge": brType = brEdge
        Case "chrome": brType = brChrome
        Case Else: printMsg "Browser initiation fails: only the defined browser types can be accepted", doRaiseError:=True
    End Select
    brPath = getBrowserPath(brType)
    strCall = """" & brPath & """ --" & ArgRemoteDebuggingName
   '--------------------------------------------------------------------------------------------------------------------------

   'Add starting URL to the call string -------------------------------------------------------------------------------------------------------------------
    If appUrl <> "" Then
        If appUrl = "about:blank" Then printMsg "Browser cannot be open under Application Mode when homepage is set to ""about:blank""", doRaiseError:=True
        If InStr(appUrl, "http") = 0 Then appUrl = "http://" & appUrl   'ie. devs can just simply input "google.com" i/o "http://google.com"
        strCall = strCall & " --app=""" & appUrl & """"
    Else
        strCall = strCall & " --homepage ""about:blank"""
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------------------------

   'Start with a choice profile -------------------------------------------------------------------------------------------------------------
   Dim addProfile As String
   If userProfile <> "" Then
        Select Case brType
            Case brEdge: addProfile = " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Microsoft\Edge\" & userProfile & """"
            Case brChrome: addProfile = " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Google\Chrome\" & userProfile & """"
        End Select
        strCall = strCall & addProfile
    End If
    brUser = userProfile
   '-----------------------------------------------------------------------------------------------------------------------------------------
       
   'Prepare for initialization ------------------------------------------------------------------------
    strCall = strCall & " " & addArgs    'Add additional arguments if instructed
    If cleanActive Then cleanUpSessions  'Clean up all running browser instances to prevent pipe error
    setCrashStateNormal                  'Prevent "Restore Session" popup
   '---------------------------------------------------------------------------------------------------

   'Initialize CDPCore class ----------------------------------------------------------------------------------------------------------
    Set core = New CDPCore
    brCallString = strCall
    result = core.Init(brCallString)
    If result <> 0 Then printMsg "Unable to initialize CDPCore", doRaiseError:=True Else printMsg "CDPCore initialized successfully"
    lngLastID = 1
   '-----------------------------------------------------------------------------------------------------------------------------------
       
   'Complete initialization ------------------------------------------------------
    attachToSession     'Find and attach the newly open Session ID
    getSessionInfo      'Store initial session info for printParams debugging
    serialize           'Create a serial string for future attachment
   '------------------------------------------------------------------------------
    
   'Make sure the page is fully loaded for inAppMode ------------------------------------------------------------------------------------
    Dim timerStart As Single, timerOut As Long, strLoc As String
   
   '07/01/25: Long Vh: simplified the check as there are web apps that does not end up loading the target url but points to another url.
   'As a result, the below check simply checks if there is "movement" in the name of the page instead
    If appUrl <> "" Then
        timerStart = Timer: timerOut = 45
        Do
            If Me.Name <> "" Then Exit Do
            If Timer - timerStart > timerOut Then Err.Raise ERR_TIMEOUT, , "Timeout waiting for the browser to load the homepage."
            DoEvents
        Loop
       wait
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------

   'Finally print the session details for debugging ------------------------------
    printParams
    printMsg "Browser initialization completed after " & Timer - timeStart & "s"
   '------------------------------------------------------------------------------

End Sub

Private Sub serialize()
'------------------------------------------------------------------
' This method transcribes Chrome's Json message into a string.
' The serialized string can then be used to automate an active
' session with a known serial.
' Updated: 03/05/23 - Replace sessionID with targetID since it is
'                     constant and stable on every CDP session
'------------------------------------------------------------------

    'パイプ通信のハンドル情報も記録しておく
    Call core.serialize

    'Excelの所定のワークシート上のテーブル:ブラウザセッション情報 に記録する
    With ShSetting01_StartBrowser
        With .ListObjects(.UseTableName(2, "CDPBrowser.serialize")).DataBodyRange
            WorksheetFunction.XLookup("lngLastID", .Columns(1), .Columns(2)).value = lngLastID
            WorksheetFunction.XLookup("targetID", .Columns(1), .Columns(2)).value = targetID
        End With
    End With

    'printMsg "New serial string: " & serialize
End Sub

Public Sub deserialize()
'------------------------------------------------------------------
' This method deciphers a serialized string to obtain the currently
' active session ID of the already open instance. The serial
' must be obtained first from the serialize method. The isLive
' public function checks if reattachment is successful.
' Updated: 03/05/23 - replace sessionID with targetID which is
'                     constant and stable on every CDP session
'------------------------------------------------------------------
           
    On Error GoTo iferr
           
    Set core = New CDPCore
    
    'ハンドル情報の読み込みへ移る
1   Call core.deserialize

    'Excelの所定のワークシート上のテーブルにある情報を読み込む
    With ShSetting01_StartBrowser
        With .ListObjects(.UseTableName(2, "CDPBrowser.deserialize")).DataBodyRange
2           lngLastID = WorksheetFunction.XLookup("lngLastID", .Columns(1), .Columns(2), 0)
3           targetID = WorksheetFunction.XLookup("targetID", .Columns(1), .Columns(2), "")
    
4           sessionID = attachToTab(targetID)
        End With
    End With
    Exit Sub

iferr: 'if deserialize fails, simply move on to starting anew
        
    printMsg "deserialization failed at line " & Erl
        
End Sub

Public Function jsEval(strString As String, Optional dbgMsg As Boolean = True) As Variant
'------------------------------------------------------------
' This function commands the attached browser instance to
' execute a Javascript expression to perform required tasks.
' Note: the original function has a bool retry variable which
' was removed for simplicity.
' Added dbgMsg to read debug line on the go when needed.
'------------------------------------------------------------
    
    Dim results As Scripting.Dictionary
    Dim strMessage As String
    Dim objMessage As Scripting.Dictionary
    
step1:
    
   'Construct message for sending ----------------------------------------------------------
    strMessage = "{""method"":""Runtime.evaluate"",""params"":{""expression"":""1+1;""}}"
    Set objMessage = core.jsConverter.ParseJson(strMessage)
    objMessage.Item("params").Item("expression") = strString & ";"
    strMessage = core.jsConverter.ConvertToJson(objMessage)
   '----------------------------------------------------------------------------------------
        
   'Send and check result -------------------------------------------------------------------------------------------
    Set results = SendMessage(strMessage, dbgMsg:=dbgMsg)
    If results Is Nothing Then printMsg "Unable to contact the Browser instance!", doRaiseError:=True
    If Not results.Exists("error") Then GoTo step2
   '-----------------------------------------------------------------------------------------------------------------
    
   'Handle errors ---------------------------------------------------------------------------------------------------
    Dim errMsg As String, retried As Boolean
    errMsg = results("error")("message")
    printMsg "jsEval Error: " & errMsg
    If errMsg = "'Runtime.evaluate' wasn't found" Then printMsg "Possibly the main session has been abruptly closed."
    If errMsg = "Session with given id not found." And Not retried Then
        printMsg "Attempting to update Session ID " & sessionID
        sessionID = "": sessionID = attachToTab(targetID)
        printMsg "New Session ID obtained: " & sessionID
        retried = True: GoTo step1
    End If
    printMsg "The Browser was unable to evaluate the JS expression: " & strString, doRaiseError:=True
   '-----------------------------------------------------------------------------------------------------------------
    
step2:

   'If the return type has a specific type we can return the result -------------------------------------------
    Dim result As Object
    Set result = results.Item("result").Item("result")
    Select Case True
        Case result.Item("subtype") = "error"
            jsEval = Split(result.Item("description"), vbNewLine)(0)
        Case result.Item("type") = "string", result.Item("type") = "boolean", result.Item("type") = "number"
            jsEval = result.Item("value")
        Case result.Item("subtype") = "node"
            jsEval = result.Item("className") & " (class='" & result.Item("description") & "')"
        Case result.Item("subtype") = "null"
            jsEval = "null"
    End Select
   '-----------------------------------------------------------------------------------------------------------
                
step3:
                
   'Output to Immediate Window if enabled -------------------------------------------------------------------
    Dim strResult As String
    If Not jsEval = "" Then strResult = " | Returned: " & jsEval Else strResult = ""
    strResult = Replace(strResult, vbNewLine, "")
    If Len(jsEval) > 200 Then strResult = Left(strResult, 200) & "(showing 200/" & Len(jsEval) & " chars)"
    If dbgMsg Then printMsg "Script executed: " & strString & strResult
   '---------------------------------------------------------------------------------------------------------
        
End Function

Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
'-----------------------------------------------------------------------
' This sub is a quick wrapper to navigate to an url.
' Updated 23/03/23
'-----------------------------------------------------------------------
    Dim curURL As String
   'Check if already navigated
    curURL = jsEval("window.location.href", False)
    If curURL = strURL Then
        printMsg "Already navigated to """ & strURL & """"
        Exit Sub
    End If

   'Quick formatting for simple URL input
    If strURL <> "about:blank" And InStr(strURL, "http") = 0 Then strURL = "http://" & strURL

   'If not yet then do navigate
    Dim results As Scripting.Dictionary
    Set results = SendMessage("{""method"":""Page.navigate"",""params"":{""url"":""" & strURL & """}}")
    'should also add a maxTimeLimit to Wait
    wait till:=till
    getSessionInfo 'update new session info for printParams
    printMsg "Navigated to: " & strURL
End Sub

Public Sub quit()
'-----------------------------------------------------------------------------------
' Closes the currently attached browser instance
' Original Name: closeBrowser
'-----------------------------------------------------------------------------------
    
    Dim results As Scripting.Dictionary
    Dim sngT As Single

    On Error Resume Next
    Set results = SendMessage("{""method"":""Browser.close"",""params"":{}}")

    sngT = Timer
    Do While isLive And Timer - sngT < 5
        sleep 0.1
        DoEvents
    Loop
    If isLive Then
        printMsg "Browser didn't close gracefully within 5s; invoking cleanUpSessions."
        cleanUpSessions
    End If
    Call deleteCDPID

    sleep 1 'it seems without waiting a bit the browser crashes and the next time it wants to recover from a crash
    printMsg "The browser has been terminated."

End Sub

Public Function isLive() As Boolean
'-----------------------------------------------------------------------------------
' Check if there is ongoing connection with an attached instance
' Original Name: connectionAlive
'-----------------------------------------------------------------------------------

    On Error GoTo iferr
    Dim strLoc As String
    strLoc = jsEval("window.location.href", dbgMsg:=False)
    isLive = True
    Exit Function

iferr:
    isLive = False
End Function


Public Sub wait(Optional till As ReadyState = isComplete, Optional dbgState As Boolean = False)
'--------------------------------------------------------------------------------------------------
' Wait until browser reaches a specific state. Default is until it fully loads ("complete")
' Possible states:  loading / interactive / complete
' Complete state is preferable when certain information is only showed when fully loaded.
' Otherwise, an interactive state is good enough and can potentially save a lot of wait time.
' However, do note that some web apps only fully displays elements at complete readystate.
' To read ReadyState on the go, set the dbgState to True to print the state to the Immediate Window
' Original Name: waitCompletion
' Updated: 26/04/23 - Enhanced 'till' to Enum type for ease of use
'                   - Added exit for targetState = interactive in case CDP misses the state change
'--------------------------------------------------------------------------------------------------
    Dim targetState As String, strState As String
    Select Case till
        Case isLoading: targetState = "loading"
        Case isInteractive: targetState = "interactive"
        Case isComplete: targetState = "complete"
    End Select
    
    '初期値を設定しておく
    '　どうせ最初はほぼ必ずloadingだから、`targetState`が`loading`になってたら、わざわざチェックする必要性はない
    strState = "loading"

    Do Until strState = targetState
        strState = jsEval("document.readyState", dbgMsg:=False)                 'dbgMsg = False to not spam Immediate Window
        If dbgState Then printMsg "Current ReadyState: " & strState             'option to print ReadyState info during iteration
        If targetState = "interactive" And strState = "complete" Then Exit Do   'also exit if better than interactive is satisfied
        sleep 0.1                                                               'reduce sleep will speed up but will cost cpu power
    Loop

    If dbgState Then printMsg "Wait '" & targetState & "' satisfied"
End Sub


Public Sub show(Optional state As WinState = asNormal, Optional xywh As String = "0 0 0 0")
'-----------------------------------------------------------------------------------
' Display the window with various optional state. The default state is "normal".
' Can also use after .attach to bring the window topmost as well. By default at
' start, the browser window's state will be that of the previous session run.
' xywh is used for inputing left top width height parameters:
' eg.1: xywh = "100 200" left = 100 and top = 200, width and height unchanged
' eg.2: xywh = "100 100 300 400" left = top = 100, width = 300 and height = 400
' eg.3: xywh = "0 0 0 0" the method will ignore and not make any resizing
'-----------------------------------------------------------------------------------
' Note: This method targets the window based on its name so it is best applied when
' the browser window is not refreshing a URL which may potentially change the name.
' Updated: 02/03/23 - Changed CInt to Int to prevent Overflow error.
'-----------------------------------------------------------------------------------

    If brHWnd = 0 Then brHWnd = getHandle
    ShowWindow brHWnd, doShow  'Make visible
    ShowWindow brHWnd, state

   'Resize to specific position and size
   'Will not work properly if window is currently maximized
    If xywh <> "0 0 0 0" Then
    
        sleep 0.3      'Sleep a bit for previous Win Api executions to finish, else the resize effect will be accidentally cancelled out
                    
        Dim params As New Scripting.Dictionary
        If brWID = "" Then brWID = invokeMethod("Browser.getWindowForTarget").Item("windowId")
        params("windowId") = Int(brWID)     '02/03/23 CInt() will throw error, Int() is better
                                
        Dim bounds As New Scripting.Dictionary
        Dim x As Long, y As Long, w As Long, h As Long
        On Error Resume Next
        x = CLng(Split(xywh, " ")(0))
        y = CLng(Split(xywh, " ")(1))
        w = CLng(Split(xywh, " ")(2))
        h = CLng(Split(xywh, " ")(3))
        On Error GoTo 0
    
        If x <> 0 Then bounds("left") = x
        If y <> 0 Then bounds("top") = y
        If w <> 0 Then bounds("width") = w
        If h <> 0 Then bounds("height") = h
        Set params("bounds") = bounds
        invokeMethod "Browser.setWindowBounds", params
    
    End If
End Sub


Public Sub hide()
'------------------------------------------------------------------
' This is the equivalent of show "hidden" but this function is
' added to make it easy to access and use the same feature.
'------------------------------------------------------------------
    If brHWnd = 0 Then brHWnd = getHandle
    ShowWindow brHWnd, asHidden
End Sub


Public Function bringToForeground()
'------------------------------------------------------------------
' Brings the browser window to the foreground
' Updated: Long Vh - 06/01/26 should be a function like others
'------------------------------------------------------------------

    If brHWnd = 0 Then brHWnd = getHandle
    If brHWnd <> 0 Then
        ShowWindow brHWnd, doShow
        BringWindowToTop brHWnd
        SetForegroundWindow brHWnd
    End If

End Function

Public Function newTab(Optional Url As String, Optional newWindow As Boolean, Optional setMain As Boolean) As CDPBrowser
'------------------------------------------------------------------
' Create a new blank tab within the same instance.
' url: The initial URL the page will be navigated to. An empty
' string indicates about:blank.
' newWindow: If true will open a separate window instead of a tab.
' In the original CDP document, each new tab is called "Target" and
' the equivalence of this method is Target.createTarget
' This returns the Session ID of the newly created tab.
' Updated: 20/03/23 Return the result as a new Browser object
'          26/04/23 Format url argument to accept simplified input
'------------------------------------------------------------------
    Dim sessionID_ As String
    Dim params As New Scripting.Dictionary
    If Url <> "" Then Url = IIf(Url <> "about:blank" And InStr(Url, "http") = 0, "http://" & Url, Url)
    params("url") = CStr(Url)
    If newWindow Then params("newWindow") = CBool(newWindow)

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.createTarget", params)
   
   'Return the new tab's Session ID
    Dim tabId As String
    tabId = results("targetId")
    sessionID_ = attachToTab(tabId)

   'Switch attachment to the new tab
    printMsg "newTab will now open a new CDP instance"
    Set newTab = New CDPBrowser
    With ShSetting01_StartBrowser
        '`new CDP object`に、既存のパイプ通信情報を渡すため、OFFになっても、情報を渡すようにする
        If Not (.Range(.UseRangeName(5, "CDPBrowser.Class_Initialize")).value) Then newTab.deserialize
    End With
    newTab.id = sessionID_
    If Url <> "" Then newTab.wait

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = sessionID_
    getSessionInfo
    serialize
    printMsg "Tab [" & brName & " | " & targetID & "] is now the main session tab for " & brID
End Function

'***************************************************************************************************
'* 機能　　：既に開いているタブのセッションIDを、タブの名前またはURLに基づいて検索します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - CDPBrowser
'* 引数　　：tabName    タブ名
'            Url        URL
'            setMain    たった今使ってる クラスモジュール - CDPBrowserに、この引数の結果に基づいて、セッションIDを置き換えます
'            SearchType "page","iframe" といった、セッションの種類を指定します
'            doRetry    `True`で、出現まで何回かリトライします。
'---------------------------------------------------------------------------------------------------
'* 更新履歴：01/08/26 Eschamali     - iframeへの切り替えができるように改良
'            06/06/23               - Added doRetry argument
'            03/05/23               - Combined the method with the getTabNew
'            24/10/25 Daniel Polak  - Predictable, precise, and script-friendly tab selection that scales well when there are many similar tabs open
'                                     Exact title and URL-prefix matching avoids accidentally selecting tabs with similar titles/URLs
'            23/10/25 Daniel Polak  - Fixed typo timeStart instead of timerStart and fixed bug 60000s instead of 60s
'---------------------------------------------------------------------------------------------------
'* 機能説明：tabName引数とurl引数は部分的な指定が可能です。SearchType はデフォルトは、"page"にしてます。
'* 注意事項：これらの引数を両方とも省略すると、getTabは最も近い未接続のタブに接続しようとします。
'***************************************************************************************************
Public Function getTab(Optional tabName As String, Optional Url As String, Optional setMain As Boolean, Optional SearchTypeID As TargetgetTargetsType, Optional doRetry As Boolean = False) As CDPBrowser
    Dim results As Scripting.Dictionary
    Dim tabId As String, candidateTabId As String, errMsg As String
    Dim timerStart As Single
    Dim targets As Object, Obj As Object
    Dim tabfound As Boolean
    Dim timeStart As Long
    Dim bestScore As Long, score As Long
    Dim info As String

    If doRetry Then timerStart = Timer
    
    'Typeを決める
    Dim SearchType As String
    Select Case SearchTypeID
        Case TargetgetTargetsType.Page: SearchType = "page"
        Case TargetgetTargetsType.iFrame: SearchType = "iframe"
        Case Else: Err.Raise vbObjectError + 1, "CDPBrowser.getTab", "このパラメーターは、未定義です"
    End Select

searchit:

    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")
    
    'スコア方式を利用して、理想の対象タブ情報を取得します
    Const MaxScore As Long = 100
    bestScore = -1
    candidateTabId = ""
    For Each Obj In targets
        If Obj("type") = "page" Then
            score = -1
            Select Case True
                Case tabName <> ""
                    If Obj.Exists("title") Then
                        If Obj("title") = tabName Then
                            score = 100
                        ElseIf InStr(Obj("title"), tabName) Then
                            score = 50
                        Else
                            errMsg = "with name containing """ & tabName & """"
                        End If
                    End If
                Case Url <> ""
                    If Obj.Exists("url") Then
                        If Left$(Obj("url"), Len(Url)) = Url Then
                            score = 90
                        ElseIf InStr(Obj("url"), Url) Then
                            score = 40
                        Else
                            errMsg = "with url containing """ & Url & """"
                        End If
                    End If
                Case Else
                    If Obj("attached") = False Then
                        score = 80
                    Else
                        score = 30
                        errMsg = "as no new unattached tab could be found"
                    End If
            End Select
            If score > bestScore Then
                bestScore = score
                candidateTabId = Obj("targetId")
            End If
            
            '最大スコアが出てるなら、捜査打ち切り
            If bestScore = MaxScore Then Exit For
        End If
    Next

    If bestScore >= 0 Then
        tabfound = True
        tabId = candidateTabId
    End If

   'Do retry as long as not timeout (default 60s)
    If Not tabfound And doRetry And Timer - timerStart < 60 Then GoTo searchit

   'Returns nothing if not found
   'Nothing is preferred over throwing error
    If Not tabfound Then
        info = IIf(tabName <> "", tabName, IIf(Url <> "", Url, "(no new tab found)"))
        printMsg "Unable to find the tab: " & info
        Exit Function
    Else
        printMsg "Found the tab with ID: " & tabId & ". (確証度：" & bestScore & "%)"
    End If

searchsound:
   Dim sessionID_ As String
   'Return the session object for reference
    sessionID_ = attachToTab(tabId)

   'Return the new CDP object
    Set getTab = New CDPBrowser
    With ShSetting01_StartBrowser
        '`new CDP object`に、既存のパイプ通信情報を渡すため、OFFになっても、情報を渡すようにする
        If Not (.Range(.UseRangeName(5, "CDPBrowser.Class_Initialize")).value) Then getTab.deserialize
    End With

    getTab.id = sessionID_  'will also update the targetID
    sleep 'give some time so that the debugging session can switch to the new tab

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = sessionID_
    getSessionInfo
    serialize
    printMsg "Tab [" & targetID & "] is now the main session tab for " & brID
End Function


Public Sub closeTab()
'------------------------------------------------------------------
' Close the current tab of the session. Requires targetId which can
' be retrieved by using getTargetId. This is the targetId of the
' active tab under focus of the automation pipe.
'------------------------------------------------------------------
    Dim serial As String, mainID As String
    
    With ShSetting01_StartBrowser
        With .ListObjects(.UseTableName(2, "CDPBrowser.closeTab")).DataBodyRange
            'Get the original serialized tab ID
            mainID = WorksheetFunction.XLookup("targetID", .Columns(1), .Columns(2), "")
        End With
    End With

    '操作中のタブ自体を閉じないように仕向ける
    If targetID = mainID Then printMsg "Unable to close this tab " & targetID & " as it is the main session tab", doRaiseError:=True: Exit Sub
           
    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    invokeMethod "Target.closeTarget", params
    printMsg brID & " was closed"
End Sub


Public Function getElementByID(strID As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class from the element ID, similar to the getElementByID of JS.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Get a new unique varID
    Dim varID As String: varID = newVarID

   'Get the object and assign to the varID in the browser
    Dim varJS As String: Dim result As String
    varJS = varID & " = document.getElementById(""" & strID & """)"
    result = jsEval(varJS)

   'Return result
    Set getElementByID = New CDPElement
    getElementByID.Init Me, varIDs, varID, varJS, result
End Function


Public Function getElementByXPath(strXPath As String) As CDPElement
'---------------------------------------------------------------------------------------
' Get an element class from XPath. The element class is init in CDPElement module.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' XPath Syntax help: https://www.w3schools.com/xml/xpath_syntax.asp
' Updated: 02/03/2023
' Examples -----------------------------------------------------------------------------
' eg.1 "//b[text()='T-BILL CURVE']": tag "b" with innertext to be 'T-BILL CURVE'
' eg.2 "//input[@name='username']": tag "input" & attribute "name" to be 'username'
' eg.3 "//a[@a1='abc' and @a2='xyz']": tag "a" with attribute a1 = abc and a2 = xyz
' eg.4 "//div[@id='abc']/div": get first div child of the div with id = abc
' eg.5 "//div[@id='abc'].//div": get first div descendant of the div with id = abc
' eg.6 "div": get first div child from the prescribed contextNode (default = 'document')
'---------------------------------------------------------------------------------------
       
   'Get a new unique varID
    Dim varID As String: varID = newVarID
       
   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")
       
   'Get the object and assign to the varID in the browser
    Dim varJS As String: Dim result As String
    varJS = varID & " = document.evaluate(""" & strXPath & """, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
    result = jsEval(varJS)

   'Return result
    Set getElementByXPath = New CDPElement
    getElementByXPath.Init Me, varIDs, varID, varJS, result
End Function


Public Function getElementsByXPath(strXPath As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by XPath. For further info on the particular XPath
' syntax used for retrieving the elements, refer to Mozilla doc below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
       
   'Get a new unique varID
    Dim varID As String: varID = newVarID & "s"     's to signify a collection varID
       
   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")
       
   'Get the object and assign to the varID in the browser
    Dim result As Variant
    Dim elements As Collection
    result = jsEval(varID & " = document.evaluate(""" & strXPath & """, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)")

   'Return result
    If result = "null" Then Exit Function Else Set elements = New Collection
   'Store element members to the collection
    Dim varJS As String                                 'Store each new member access JS string
    Dim length_ As Long
    length_ = jsEval(varID & ".snapshotLength")
    Dim tElement As CDPElement, N As Long
    For N = 0 To length_ - 1                            'HTML index starts at zero
        varJS = varID & ".snapshotItem(" & N & ")"
        Set tElement = New CDPElement
        tElement.Init Me, varIDs, varJS_:=varJS
        elements.Add tElement
    Next

    Set getElementsByXPath = elements
End Function


Public Function getElementByQuery(strQuery As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class by QuerySelector - similar to that of the same in JS.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation. Refer to the below for some examples.
' https://www.w3schools.com/jsref/met_document_queryselector.asp
' https://stackoverflow.com/a/29937844
' https://www.w3.org/TR/selectors-3/#context
' Updated: 08/11/2023
' Examples ------------------------------------------------------------------------
' eg.1 "#abc": get the first element with id attribute = abc
' eg.2 "tb.xyz": get the first element with tag tb and with class xyz
' eg.3 "a[name='abc']": get the first element with tag a and attribute name = abc
' eg.4 "div > p": get the first element with tag 'p' and has parent with tag 'div'
' eg.5 "[a1='abc'][a2='xyz']": get element with attribute a1 = abc and a2 = 123
' eg.6 "[a1='abc'] [a2='xyz']": get element with attribute a2 = xyz and is a child
' of an element with attribute a1 = abc
' eg.7 "#processType [id*='select_option']": get child element with id text that
' contains "select_option" and is a child of element with id "processType"
' eg.8 "#upload:not([disabled])": get the first element with id=upload and does not
' have attribute 'disabled'.
' eg.9 "#upload:disabled": get the first element with id=upload and has attribute
' 'disabled'.
' eg.10 "#id1 tr:last-child [id*='invoiceno']": first get element with id text
' 'id1', then look for the last child with tag 'tr', then finally look within the
' child's children for the first element with id containing 'invoiceno'.
'----------------------------------------------------------------------------------
    
   'Get a new unique varID
    Dim varID As String: varID = newVarID
   
   'Get the object and assign to the varID in the browser
    Dim varJS As String: Dim result As String
    varJS = varID & " = document.querySelector(""" & strQuery & """)"
    result = jsEval(varJS)
    
   'Return result
    Set getElementByQuery = New CDPElement
    getElementByQuery.Init Me, varIDs, varID, varJS, result
    
End Function

Public Function getElementsByQuery(strQuery As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by QuerySelectorAll. The syntax is as below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
           
   'Get a new unique varID
    Dim varID As String: varID = newVarID & "s"     's to signify a collection varID
           
   'Get the object and assign to the varID in the browser
    Dim result As Variant, elements As Collection
    result = jsEval(varID & " = document.querySelectorAll(""" & strQuery & """)")
    
   'Return result
    If result = "null" Then Exit Function Else Set elements = New Collection
            
   'Store element members to the collection
    Dim tElement As CDPElement
    Dim varJS As String                                 'Store each new member access JS string
    Dim length_ As Long, N As Long
    length_ = jsEval(varID & ".length")
    For N = 0 To length_ - 1                            'HTML index starts at zero
        varJS = varID & ".item(" & N & ")"
        Set tElement = New CDPElement
        tElement.Init Me, varIDs, varJS_:=varJS
        elements.Add tElement
    Next
    
    Set getElementsByQuery = elements
    
End Function


Public Sub jsAddLib(strURL As String)
'----------------------------------------------------------------------------------
' Load the external JS Library to the document for use of specialized JS methods.
' This is very useful when built-in JS commands are not sufficient to execute a
' task easily, such as when taking a screenshot of an entire webpage while this can
' be done relatively easy using external JS libraries. This function shall enables
' massive possibility of extended features.
' https://stackoverflow.com/a/57572107
' Updated: 23/03/2023
'          23/10/2025 Daniel Polak - fixed bug in check whether JS library is loaded
'----------------------------------------------------------------------------------
' Note: the script library should only be added after the page is fully loaded. If
' it is loaded when the HTML document is refereshing, it risks adding the script to
' the old HTML document which will be replaced by the new one causing the script to
' be inexistent and resulting in "ReferenceError" due to its functions not defined.
'----------------------------------------------------------------------------------
    
   'Check if already loaded, if so then no need to add again
    If getElementByQuery("script[src='" & strURL & "']").varResult <> "null" Then
        printMsg "The HTML already contains the JS Library: """ & strURL & """"
        Exit Sub
    End If

   'Add script loading function
   '(Javascript Original)
    '    function loadScript(name) {
    '        var script = document.createElement('script');
    '        script.type = 'text/javascript'; script.src = name;
    '        document.head.appendChild(script);
    '        script.onload = () => {};                              //added to make sure script has been loaded
    '        return 'Script Library fully loaded';
    '    }
    jsEval "function loadScript(name) { var script = document.createElement('script'); script.type = 'text/javascript'; script.src = name; document.head.appendChild(script); script.onload = () => {}; return 'Script library loaded'; }"

   'Load the script URL
    jsEval "loadScript('" & strURL & "')"
    printMsg "Successfully added the JS Library: """ & strURL & """"

End Sub

Public Sub jsAddScript(strLocalScriptPath As String)
'----------------------------------------------------------------------------------
' Load the contents of the local script file directly to the HTML page. This is
' useful when you have a custom local script to add in a certain feature to a web
' application (eg adding a custom popup message).
' Updated: 01/12/2025
'----------------------------------------------------------------------------------
' Note: the library should only be added after the page is fully loaded. If
' it is loaded when the HTML document is refereshing, it risks adding the script to
' the old HTML document which will be replaced by the new one causing the script to
' be inexistent and resulting in "ReferenceError" due to its functions not defined.
'----------------------------------------------------------------------------------

    If Dir(strLocalScriptPath) = "" Then Exit Sub

    Dim fso As New FileSystemObject
    jsEval fso.OpenTextFile(strLocalScriptPath).ReadAll 'Load the script contents
    printMsg "Successfully added the local script: """ & strLocalScriptPath & """"

    Set fso = Nothing

End Sub

Public Sub notify(msg)
'----------------------------------------------------------------------------------
' Send a custom notification message on the bottom right corner of the screen. This
' popup message is nicer than alert() and it is not modal: ie. does not the user to
' interact with it while pausing the automation.
' Credit: Claude AI
' Updated: 01/12/2025
'----------------------------------------------------------------------------------

   'Make sure backslash if any is escaped
    msg = Replace(msg, "\", "\\")

    jsEval "function showStatus(msg, duration = 10000) {" & _
                "let c = document.getElementById('statusContainer');" & _
                "if (!c) {" & _
                    "c = document.createElement('div');" & _
                    "c.id = 'statusContainer';" & _
                    "c.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:9999;';" & _
                    "document.body.appendChild(c);" & _
                "}" & _
                "const m = document.createElement('div');" & _
                "m.style.cssText = 'background:#4CAF50;color:white;padding:16px 24px;margin-top:10px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3);white-space:nowrap;transition:all 0.3s;opacity:0;transform:translateX(400px);font-size:16px;font-family:system-ui,-apple-system,BlinkMacSystemFont,Roboto,sans-serif;font-weight:300;line-height:1.5;';" & _
                "m.textContent = msg;" & _
                "c.appendChild(m);" & _
                "setTimeout(() => { m.style.opacity = '1'; m.style.transform = 'translateX(0)'; }, 10);" & _
                "setTimeout(() => {" & _
                    "m.style.opacity = '0';" & _
                    "m.style.transform = 'translateX(400px)';" & _
                    "setTimeout(() => c.removeChild(m), 300);" & _
                "}, duration);" & _
            "}"

    jsEval "showStatus('" & msg & "')"

End Sub

Public Sub snapPage(fullNameWithPngExtension As String, Optional getFullPage As Boolean = False)
'----------------------------------------------------------------------------------
' Capture a screenshot of the entire HTML page from the top to bottom. The method
' employs html2canvas external library is which loaded via addJsLibrary. To snap
' only a certain element within the page, indicate the same with the targetElement
' argument (default is "document.body" for the entire page).
' Original source: https://stackoverflow.com/a/51478809
' Updated: 14/04/23
'          23/10/25 Daniel Polak - fixed bug where the all files in the Downloads folder were deleted instead of just the temporary screenshot
'          24/10/25 Daniel Polak - added CDP native screenshot capability which avoids use of external JavaScript library
'          07/01/26 Long Vh      - removed the use of html2canvas js lib in favor of using the native CDP method to capture the page. Overhauled the function design.
'----------------------------------------------------------------------------------
' Note : if getFullPage = false, snapPage will only get a screenshot of the current
'        view. If true, it will get all the way from the start to the bottom. See
'        demo.getSnapShot sub procedure for example.
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary
    Dim r As Scripting.Dictionary
    Dim b64 As String, tmpPng As String

   'Verify if the page can be captured with screenshot
    Set r = invokeMethod("Page.getLayoutMetrics")
    If invokeError(r) Then
        printMsg "Unable to perform snapPage on this screen as it has no layout metrics."
        Exit Sub
    End If

   'Perform the snip
    If getFullPage Then params.Add "captureBeyondViewport", True
    Set r = invokeMethod("Page.captureScreenshot", params)

   'Save to the same folder if not indicated
    If Not invokeError(r) Then
        b64 = r("data")
        SaveBase64ToFile b64, fullNameWithPngExtension
        printMsg "Screenshot saved to: " & fullNameWithPngExtension
    End If

End Sub

Public Sub printMsg(strMsg As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
'----------------------------------------------------------------------------------
' A custom printMsg that besides printing to the Immediate Window, it will
' also print to a dated log file. This is useful at user level where
' the Immediate Window is not easily accessible. doRaiseError will also raise an
' error aside from printing the error message to the log file.
' Updated: 07/01/26 Long Vh: the message output is disabled by default to avoid
'                            spamming info. This can be turned on by setting globally
'                            noDbgMsg = False or noLogging = False
'----------------------------------------------------------------------------------
    
    Dim intFile As Integer
    Dim logName As String, strFormattedMsg As String
    
    If isHeader Then strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-") _
    Else: strFormattedMsg = Format(Now, "hh:mm:ss") & " | " & brID & " | " & strMsg

   'Perform msg output to Immediate Window if needed
    If doPrintDbgMsg Then
        Debug.Print strFormattedMsg
    End If

   'Perform msg logging if needed
    If doLog Then
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".cdp.txt"
        If Dir(logPath, vbDirectory) = "" Then MkDir logPath
        intFile = FreeFile
        Open logPath & "\" & logName For Append As #intFile
        Print #intFile, strFormattedMsg
        Close #intFile
    End If

On Error Resume Next 'GCUser99
    If doRaiseError Then Err.Raise ERR_PROTOCOL, Description:=strMsg

End Sub

Public Sub printTabs()
'------------------------------------------------------------------
' Print all current tabs info. Useful for debugging.
'------------------------------------------------------------------
    Dim targets As Object, target_ As Object, i As Long
    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

   'Read info
    printMsg "---------------------------------------------------"
    printMsg "Tabs Information"
    printMsg "---------------------------------------------------"

    i = 1
    For Each target_ In targets
        If target_("type") = "page" Then
            printMsg "#" & i & vbTab & _
                        "type: " & target_("type") & " | " & _
                        "attached: " & target_("attached") & " | " & _
                        "title: " & target_("title") & " | " & _
                        "url: " & target_("url") & " | " & _
                        "target ID: " & target_("targetId")
            i = i + 1
        End If
    Next
End Sub

Public Sub printParams()
'------------------------------------------------------------------
' Print all parameters of the current session for debuggings.
'------------------------------------------------------------------

    printMsg "---------------------------------------------------"
    printMsg brID & " Information"
    printMsg "---------------------------------------------------"

    printMsg "Session ID    | sessionID | " & sessionID
    printMsg "Last ID       | lngLastID | " & lngLastID
    printMsg "Window Handle | brHWnd    | " & brHWnd
    printMsg "Window ID     | brWId     | " & brWID
    printMsg "Browser Type  | brType    | " & brType
    printMsg "Browser Ver   | brVer     | " & brVer
    printMsg "Browser Path  | brPath    | " & brPath
    printMsg "Tab URL       | brUrl     | " & brURL
    printMsg "Tab Title     | brName    | " & brName
    printMsg "Tab ID        | targetID  | " & targetID
    printMsg "Core StdInWr  | hStdInWr  | " & core.hStdInWr
    printMsg "Core StdOutRd | hStdOutRd | " & core.hStdOutRd
    printMsg "Core CDPInWr  | hCDPInWr  | " & core.hCDPInWr
    printMsg "Core CDPOutRd | hCDPOutRd | " & core.hCDPOutRd
    printMsg "CDP Version   | CDPVer    | " & CDPVer
End Sub


Public Function tabCount() As Long
'------------------------------------------------------------------------------
' Print all current tabs info. Useful for debugging.
' Updated : 05/02/25 - changed from printTabs to tabCount to get the count too
'------------------------------------------------------------------------------

    Dim results, targets, target_ As Scripting.Dictionary
    Dim i As Integer

    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

   'Read info
    printMsg "---------------------------------------------------"
    printMsg "Tabs Information"
    printMsg "---------------------------------------------------"

    i = 0
    For Each target_ In targets
        If target_("type") = "page" Then
            printMsg "#" & i & vbTab & _
                        "type: " & target_("type") & " | " & _
                        "attached: " & target_("attached") & " | " & _
                        "title: " & target_("title") & " | " & _
                        "url: " & target_("url") & " | " & _
                        "target ID: " & target_("targetId")
            i = i + 1
        End If
    Next

    tabCount = i

End Function

Public Sub sleep(Optional seconds As Double = 0.5)
'------------------------------------------------------------------
' Custom sleep function. Sleep by 0.5s by default.
' Useful for a quick necessary pause when needed.
'------------------------------------------------------------------

    Const baseUnit As Long = 1000    'ie. millisecs

    sleep2 seconds * baseUnit
    DoEvents
End Sub

Public Sub activate()
'----------------------------------------------------------------------------------
' Switch focus to the current tab being attached to under the current BRID.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    invokeMethod "Target.activateTarget", params
End Sub


Public Sub handleDialog(Accept As Boolean, Optional promptText As Variant)
'----------------------------------------------------------------------------------
' Accepts or dismisses a JavaScript initiated dialog (alert, confirm, prompt, or
' onbeforeunload). accept = True to accept the dialog, else dimiss it. promptText
' is the text to enter into the dialog prompt before accepting, used only if this is a
' prompt dialog.
' Note: this method requires the execution of "pageEnable" before the dialog appears
' so that the dialog event is probably exposed and captured as an event.
' Updated: 23/07/2025 - added as requested by Jane Zhang for her project
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary

    params("accept") = CBool(Accept)
    If Not IsMissing(promptText) Then params("promptText") = CStr(promptText)
    invokeMethod "Page.handleJavaScriptDialog", params

End Sub

Public Sub pageEnable()
'----------------------------------------------------------------------------------
' Enable the page for events listening. Required for handleDialog for instance.
'----------------------------------------------------------------------------------

    invokeMethod "Page.enable"

End Sub

'***************************************************************************************************
'* 機能　　：CDPコマンドの核となる、`method`、およびオプションの`params`を直接指定して、実行します
'---------------------------------------------------------------------------------------------------
'* 返り値　：実行結果のJson-Dictionary
'* 引数　　：methodName         Network.getCookies とか、Browser.getVersion 等々
'            params             オプションパラメーターセット。予め呼び出し側で、Json-Dictionaryとして組み立ててください
'            EventMessages      ブラウザからのイベント情報を取得したい場合に指定します。参照渡しによりここに実行までの過去のイベントが、Json-Dictionaryとして積まれていきます。省略で、過去のイベントを破棄します
'            dbgMsg             実行結果をイミディエイトに出すか指定します
'---------------------------------------------------------------------------------------------------
'* 機能説明：・低レベル操作により、実装されてる全ての CDP－Json が実行可能です
'            ・イベント情報も確認したい場合は第三引数に、New Dictionary 済みの変数を渡すことで、受け取ることが可能です
'
'* 注意事項：・エラーメッセージは、invokeError によってイミディエイトウィンドウに自動的に表示されます
'            ・一部のイベントは、「"method": "XXXX.enable"」といった形式で事前送信(設定)する必要があります
'---------------------------------------------------------------------------------------------------
'* 関連URL ：・出典                     https://github.com/PerditionC/VBAChromeDevProtocol
'            ・CDPドキュメントブック　  https://chromedevtools.github.io/devtools-protocol/tot/Brows
'---------------------------------------------------------------------------------------------------
'* 更新履歴：・25/12/18 - イベント情報を格納/破棄のスイッチングができる仕組みを実装
'            ・25/04/23 - Set this to Public as it is useful
'            ・20/03/23 - Added printMsg msgs & set to Private
'***************************************************************************************************
Public Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional EventMessages As Dictionary, Optional dbgMsg As Boolean = False) As Scripting.Dictionary
    Dim paramStr As String
    Dim results As Scripting.Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = core.jsConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}", EventMessages)

   'Print error if there is
    If invokeError(results) Then printMsg "Failed to invoke method """ & methodName & """", doRaiseError:=True

   'Return result object
    If results.Exists("result") Then Set invokeMethod = results.Item("result")
    If dbgMsg Then printMsg "Invoked """ & methodName & """ with result: " & core.jsConverter.ConvertToJson(results("result"))
End Function



'=============================================================================
' PUBLIC PROPERTIES
'=============================================================================

Property Get Url() As String
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------
    brURL = jsEval("document.URL", False)
    Url = brURL
End Property


Property Let Url(newURL As String)
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------
    navigate newURL
    brURL = newURL
End Property


Property Get Name() As String
'------------------------------------------------------------
' Get the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------
    brName = jsEval("document.title", False)
    Name = brName
End Property

Property Let Name(newName As String)
'------------------------------------------------------------
' Set the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------
    jsEval "document.title = '" & newName & "'", False
    printMsg "Tab name changed to " & newName
    brName = newName
End Property

Property Get id() As String
'------------------------------------------------------------
' Get the session ID of the current browser
' Added 02/03/2023
'------------------------------------------------------------
    id = sessionID
End Property

Property Let id(newID As String)
'------------------------------------------------------------
' Set the session ID of the current browser to a new value
' Added 02/03/2023
'------------------------------------------------------------
    sessionID = newID
    getSessionInfo        'update all session infos
    printMsg brID & " updated Session ID to " & sessionID
End Property

Property Get html() As String
'------------------------------------------------------------
' Get the html text of the entire document
' Added 02/03/2023
'------------------------------------------------------------
    html = jsEval("document.documentElement.innerHTML")
End Property



'***************************************************************************************************
'                              ■■■ プライベートプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP にメッセージを送信し、Chromium ブラウザでメッセージが受信されているかどうかを確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド の実行結果(Json-Dictionary 形式)
'* 引数　　：strMessage         Chromium ブラウザに送るCDP-Jsonコマンド(プレーンテキスト)
'            objAllMessages     Chromium ブラウザから得たブラウザからのイベント情報格納用変数。実質返り値
'            dbgMsg             `True`でエラー時、イミディエイトに表示させます。
'---------------------------------------------------------------------------------------------------
'* 更新履歴：・03/05/23 - Added soft-reset on peeknamepipe error
'            ・03/22/23 - Added error handling for better debugging
'            ・24/10/25 Daniel Polak - Added soft re-attach attempt when pipe error
'***************************************************************************************************
Private Function SendMessage(strMessage As String, Optional objAllMessages As Scripting.Dictionary, Optional dbgMsg As Boolean = True) As Scripting.Dictionary
    Dim intRes As Long
    Dim strRes As String
    Dim lngCurrentId As Long
    Dim results As Scripting.Dictionary       'to store a collection of IDs obtained from CDP
            
    lngCurrentId = lngLastID
    lngLastID = lngLastID + 1       'We increase the global ID counter

RetrySend:

    On Error GoTo ifError
            
   'Before sending a message the message-buffer is emptied
   'All messages that we have received so far cannot be an answer to the message that we will send
   'So they can be safely discarded
    Dim ErrorCode_PeekNamedPipe As Long, ErrorCode_ReadFile As Long
    Dim ErrorDetail As New WinApiError
1   If objAllMessages Is Nothing Then

        'データが尽きるまで、ループ受信処理を行います
        Do
            '受信処理
2           intRes = core.readProcCDP(strRes, ErrorCode_PeekNamedPipe, ErrorCode_ReadFile)

            '返り値チェック
            If intRes > 0 Then
               'データありだが、あくまで破棄処理のため何もしない(強いていうなら、デバッグメッセージを出す)
3               printMsg "Emptying the message buffer..."
4               If seeRawSendMsgDbg Then printMsg "strRes = " & strRes
            Else
                'データなしの場合は、エラーコードチェックを行う
                If ErrorCode_PeekNamedPipe Then
                    'PeekNamedPipe にてエラーがあった場合
5                   printMsg "Error PeekNamedPipe in readProcCDP. ErrorCode:" & ErrorCode_PeekNamedPipe & " - " & ErrorDetail.GetMessage(ErrorCode_PeekNamedPipe, "kernel32")
                    Err.Raise ERR_PIPE, Description:="Error PeekNamedPipe in readProcCDP." & vbCrLf & vbCrLf & "<ErrorCode:" & ErrorCode_PeekNamedPipe & ">" & vbCrLf & ErrorDetail.GetMessage(ErrorCode_PeekNamedPipe, "kernel32")
                ElseIf ErrorCode_ReadFile Then
                    'ReadFile にてエラーがあった場合
6                   printMsg "Error ReadFile in readProcCDP. ErrorCode:" & ErrorCode_ReadFile & " - " & ErrorDetail.GetMessage(ErrorCode_ReadFile, "kernel32")
                    Err.Raise ERR_PIPE, Description:="Error ReadFile in readProcCDP." & vbCrLf & vbCrLf & "<ErrorCode:" & ErrorCode_ReadFile & ">" & vbCrLf & ErrorDetail.GetMessage(ErrorCode_ReadFile, "kernel32")
                Else
                    'データがまだ、ない。つまり、破棄完了
                    printMsg "Message queue destruction completed."
                End If
            End If
        Loop Until intRes < 1
    End If
    
   '--------------------------------------------------
   'Send message to the browser
   '--------------------------------------------------
7   intRes = core.readProcSTD(strRes)                                                            'Sometimes edge writes to stdout so we clear stdout too.
8   strMessage = Left(strMessage, Len(strMessage) - 1)                                           'We add the currentID and sessionID to the message
9   If sessionID <> "" Then strMessage = strMessage & ", ""sessionId"":""" & sessionID & """"
10  strMessage = strMessage & ", ""id"":" & lngCurrentId & "}" & vbNullChar
11  core.writeProc strMessage                                                                    'Send message to the browser instance
    
   '------------------------------------------------
   'Read message returned by browser
   '------------------------------------------------
    Dim timerStart As Double, timerOut As Long
    Dim EventCount As Long
   
    timerStart = Timer                                                                           'Failsafe timer to ensure loop ends when server fails to respond
    timerOut = 30                                                                                'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
    intRes = 1

    '受信処理→Json処理→受信処理→Json処理→…　というロジックでループ処理を行います
    'CDP-Jsonコマンドを送った際のIDが返ってくるまで、ループ受信処理を行いますが、タイムアウトになったら、抜けます
    Dim ResultSendCommand As Boolean
    Do
        '受信処理
12      intRes = core.readProcCDP(strRes, ErrorCode_PeekNamedPipe, ErrorCode_ReadFile)

        '返り値チェック
        If intRes > 0 Then
            'データあり。受信メッセージを活用するため、null文字区切りとしてスマートに分割
13          Dim JsonMessages: JsonMessages = Split(strRes, vbNullChar)
                
            'メモリを解放
14          strRes = vbNullString
                
            Dim i As Long
            For i = LBound(JsonMessages) To UBound(JsonMessages)
15              If seeRawSendMsgDbg Then printMsg "strRes = " & JsonMessages(i)                              'Only enable if needs to see every raw JS message for debugging/development

                '最後の要素か？
16              If i = UBound(JsonMessages) Then
                    '空欄か？
                    If JsonMessages(i) = "" Then
                        '1回の受信で全て正しいJson形式として得れました
                        PieceData = vbNullString
                        printMsg "readProcCDP was able to receive everything !"
                    Else
                        'データが、"欠けた"状態なので蓄積で、キャッシュしておく。頻繁には来ない想定
                        PieceData = PieceData & JsonMessages(i)
                        printMsg "readProcCDP was unable to receive all the data. It accumulates the last element data..."
                    End If
                Else
                    '欠けたデータがあるか？
                    Dim tmp As String
                    If StrPtr(PieceData) Then
                        '結合する
                        tmp = PieceData & JsonMessages(i)
                        
                        'メモリを解放
                        PieceData = vbNullString
                    Else
                        'そのまんま使う
                        tmp = JsonMessages(i)
                    End If
                    
                    '受信データを Json-Dictionary として変換
17                  Set results = core.jsConverter.ParseJson(tmp)

                    'イベント内容を確認するモードなら、格納しておく
18                  If Not objAllMessages Is Nothing Then OnNewEventReceived objAllMessages, results, EventCount

                    '実行コマンド結果であるか？
19                  If results.Exists("id") Then
                        'そのIDは、Excelに記録してるIDと一致してるか？
20                      If results.Item("id") = lngCurrentId Then
                            '結果を返して、受信ループ終了フラグを立てる
21                          Set SendMessage = results
                            ResultSendCommand = True
                            printMsg "The command result was obtained from readProcCDP ! - ID:" & lngCurrentId
                        End If
                    End If
                End If
            Next
        Else
            'データなしの場合は、エラーコードチェックを行う
            If ErrorCode_PeekNamedPipe Then
                'PeekNamedPipe にてエラーがあった場合
22              printMsg "Error PeekNamedPipe in readProcCDP. ErrorCode:" & ErrorCode_PeekNamedPipe & " - " & ErrorDetail.GetMessage(ErrorCode_PeekNamedPipe, "kernel32")
                Err.Raise ERR_PIPE, Description:="Error PeekNamedPipe in readProcCDP." & vbCrLf & vbCrLf & "<ErrorCode:" & ErrorCode_PeekNamedPipe & ">" & vbCrLf & ErrorDetail.GetMessage(ErrorCode_PeekNamedPipe, "kernel32")
            ElseIf ErrorCode_ReadFile Then
                'ReadFile にてエラーがあった場合
23              printMsg "Error ReadFile in readProcCDP. ErrorCode:" & ErrorCode_ReadFile & " - " & ErrorDetail.GetMessage(ErrorCode_ReadFile, "kernel32")
                Err.Raise ERR_PIPE, Description:="Error ReadFile in readProcCDP." & vbCrLf & vbCrLf & "<ErrorCode:" & ErrorCode_ReadFile & ">" & vbCrLf & ErrorDetail.GetMessage(ErrorCode_ReadFile, "kernel32")
            Else
                'データがまだ、ない
                printMsg "The command result from readProcCDP has not yet been received, so reception continues."
            End If
        End If

        DoEvents

        '一定時間経っても来ないなら、ループから抜けます
24      If Timer - timerStart > timerOut Then Err.Raise ERR_TIMEOUT, , "Timeout waiting for browser to return message."

        'CDP-Jsonコマンドの結果が得れてるなら、受信ループ処理を終了します
    Loop Until ResultSendCommand
Exit Function

afterErr:

    printMsg "Attempt to soft re-attach when pipe error because sometimes a simple retry is all it is needed to make it work :)"
    On Error Resume Next
    attachToSession
    GoTo RetrySend

ifError:
                  
    If Not dbgMsg Then Exit Function    'Allow the option to bypass error if intended
                  
   'Record errmsg and print helpful infos
    Dim errorLine As Long, errorDesc As String
    errorLine = Erl
    errorDesc = Err.Description
    printMsg "sendMessage error at line " & errorLine & ": """ & errorDesc & """"
    printMsg "strMessage used by sendMessage: " & strMessage

    Dim blnDidSoftRetry As Boolean, blnDoSoftRetry As Boolean
    blnDoSoftRetry = False
    If Not blnDidSoftRetry Then
        If InStr(errorDesc, "PeekNamedPipe") > 0 And brReAttach Then
            blnDoSoftRetry = True
            blnDidSoftRetry = True
            Err.Clear
        End If
    End If
    If blnDoSoftRetry Then
        Resume afterErr
    End If

    If InStr(errorDesc, "PeekNamedPipe") > 0 Then
        printMsg "Possible reasons: (1) the browser is not actually running (ie. call string failure); (2) the browser window was accidentally closed; (3) the CDP Pipe connection to the browser was cut-off somehow."
        deleteCDPID
    End If

    Err.Raise IIf(InStr(errorDesc, "PeekNamedPipe") > 0, ERR_PIPE, ERR_PROTOCOL), , "sendMessage error at line " & errorLine & ": """ & errorDesc & """"
        
End Function

'***************************************************************************************************
'* 機能　　：イベントの「種類（メソッド名）」ごとの仕切りを入れ、そこにCollectionを入れていきます
'---------------------------------------------------------------------------------------------------
'* 実質返値：Accumulator
'* 引数　　：Accumulator    これまでの格納イベント。
'            eventData      Json-Dictionry済みのブラウザからのイベントデータ
'            Counter        呼ばれた回数。参照渡しにより、カウントは引き継がれます
'---------------------------------------------------------------------------------------------------
'* 機能説明：メソッド名をキーにし、それぞれの中では、時系列が保つように追加していきます。
'            これにより、「特定の"種類"のイベント」を、個別にチェックしやすくなります。
'            時系列データ(Index値)もおまけとして残しておきます。
'***************************************************************************************************
Private Sub OnNewEventReceived(Accumulator As Dictionary, eventData As Dictionary, Counter As Long)
    ' イベント名を取得しますが、イベント(Method)名がない場合は、コマンド結果なので即抜けする
    Dim eventMethod As String
    If eventData.Exists("method") Then eventMethod = eventData("method") Else Exit Sub

    ' インデックスを付与(1始まり)
    ' これにより、全体の時系列データも保持できます
    Counter = Counter + 1
    eventData.Add "__index__", Counter
    

    ' 1. このイベント名の「仕切り板」が、すでにあるか？
    Dim eventList As Collection
    If Not Accumulator.Exists(eventMethod) Then
        ' なければ、新しい「カードの束（Collection）」を作って、仕切り板と一緒に、棚に入れる
        Set eventList = New Collection
        Accumulator.Add eventMethod, eventList
    Else
        ' あれば、その仕切り板の後ろにある、既存の束を取り出す
        Set eventList = Accumulator(eventMethod)
    End If
    
    ' 2. 取り出した束に、今回のイベントの詳細データを、追加する
    eventList.Add eventData
End Sub

Private Function attachToSession() As Integer
'-----------------------------------------------------------------------
' This function must be called after start and before all other methods.
' This function attaches to a session of the browser.
' Updated: 03/05/23 - Added printMsg msgs & removed URL checking req.
'-----------------------------------------------------------------------
' Note: if needed to read the returned message result from sendMessage,
' type ?core.jsConverter.ConvertToJson(objAllMessages) in the ImmWnd
'-----------------------------------------------------------------------
    
    Dim results As Scripting.Dictionary
    Dim results_ As Scripting.Dictionary
    Dim results__ As Scripting.Dictionary
    Dim results___ As Scripting.Dictionary
    
    Dim strKey As Variant
    
    On Error GoTo ifError
    
    Dim objAllMessages As New Scripting.Dictionary
1   Set results = SendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":true}}", objAllMessages)
    printMsg "Now looking for the browser window"
    
   'Read message to confirm the requested url is found
   'Checking results__.Item("url") has been skipped for simplicity since v2.3
   'A new session always has the property attached = false and method = Target.targetCreated
   'This way of verification vastly improves the stability over checking by URL as sometimes
   'there are rogues targets with similar URL but they are not attachable (ie. no Session ID)
    Dim boolFound As Boolean
    For Each results In objAllMessages("Target.targetCreated")
        Set results_ = results.Item("params")
        If results_.Exists("targetInfo") Then
            Set results__ = results_.Item("targetInfo")
            If results__.Item("attached") = False And results__.Item("type") = "page" Then
                boolFound = True
                Exit For
            End If
        End If
    Next
2   If Not boolFound Then printMsg "Unable to find the browser session", doRaiseError:=True _
    Else: printMsg "Now attaching CDP pipes to the browser session"
    
   'Establish connection with the CDP pipes
3   targetID = results__.Item("targetId")
4   Set results___ = SendMessage("{""method"":""Target.attachToTarget"",""params"":{""targetId"":""" & results__.Item("targetId") & """,""flatten"":true}}")
5   sessionID = results___.Item("result").Item("sessionId")
6   Set results___ = SendMessage("{""method"":""Runtime.enable"",""params"":{}}")
7   Set results___ = SendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":false}}")
    
    printMsg "Successfully attached to Session ID " & sessionID
    Exit Function
    
ifError:
                 
   'Record errmsg and print helpful infos
    Dim errorLine As Long, errorDesc As String
    errorLine = Erl
    errorDesc = Err.Description
    printMsg "attachToSession error at line " & errorLine & ": """ & errorDesc & """"
    printMsg "objAllMessages = " & core.jsConverter.ConvertToJson(objAllMessages)
    printMsg "results = " & core.jsConverter.ConvertToJson(results)
    printMsg "results_ = " & core.jsConverter.ConvertToJson(results_)
    printMsg "results__ = " & core.jsConverter.ConvertToJson(results__)
    printMsg "results___ = " & core.jsConverter.ConvertToJson(results___)
    printMsg ".start call string: " & brCallString
    
    printParams
    Err.Raise ERR_PROTOCOL, , "attachToSession error at line " & errorLine & ": """ & errorDesc & """"
    
End Function

Private Function invokeError(result As Scripting.Dictionary) As Boolean
'-----------------------------------------------------------------------------------
' Check if InvokeMethod fails and what the error message returned is.
' Error message will be shown via printMsg
' Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'-----------------------------------------------------------------------------------

    If result Is Nothing Then
        Dim ErrorCode As Long, ErrorMessage As String, ErrorData As Variant
        ErrorCode = -1 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        ErrorData = vbEmpty
        printMsg "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage
        invokeError = True
    ElseIf result.Exists("error") Then
        Dim errorDetails As Scripting.Dictionary
        Set errorDetails = result("error")
        ErrorCode = errorDetails("code")
        ErrorMessage = errorDetails("message")
        printMsg "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage
        If errorDetails.Exists("data") Then
            ErrorData = errorDetails("data")
            printMsg "       " & ErrorData
        Else
            ErrorData = vbEmpty ' vbNullString? is errorDetails("data") a string or object?
        End If
        invokeError = True
    Else
        ErrorCode = 0
        ErrorMessage = vbNullString
        ErrorData = vbEmpty
    End If
End Function


Private Function getHandle() As LongPtr
'-----------------------------------------------------------------------------------
' Added by Long: obtain the window handle of the current browser session
' Updated: 02/03/23 - Add mechanism to work with browser windows that are blank.
'-----------------------------------------------------------------------------------

    Dim bufferString As String
    Dim titleLen As Long
    Dim hWnd As LongPtr

    'Determine current tab title, if error: the tab is a blank tab
    Dim sessionTitle As String, windowTitle As String
    sessionTitle = jsEval("document.getElementsByTagName(""title"")[0].innerHTML", False)
    If InStr(sessionTitle, "TypeError") Then sessionTitle = "about:blank"

    'Obtain handles to the correct browser window
    'Both Edge and Chrome use the same class chromeWindowClass
    hWnd = FindWindowEx(0&, 0&, StrPtr(chromeWindowClass), 0&)        'To return Chrome / Edge window handles
    Do
                
        'Get Window Title based on specified handles
        'Requires string buffering declaration as string
        bufferString = String(1024, " ")
        titleLen = GetWindowText(hWnd, StrPtr(bufferString), Len(bufferString))
        windowTitle = Left(bufferString, titleLen)
    
        'Verify if matching with the sessionTitle and is really an edge/chrome window
        If windowTitle <> "" Then If InStr(windowTitle, sessionTitle) Then Exit Do
    
        hWnd = FindWindowEx(0&, hWnd, StrPtr(chromeWindowClass), 0&)      'Continue with the next handle value
    
    Loop Until hWnd = 0 ' no more matching windows

    getHandle = hWnd
End Function

Private Sub cleanUpSessions()
'-----------------------------------------------------------------------------------------------------------------------------------------
' Provides cleaning before firing up a new session to prevent
' pipe error
' Update: 05/11/24: Daniel Polak - modified to solve problems with closing sessions (that do not belong to this user) on terminal server
'         23/10/25: Daniel Polak - p.name and p.terminate were not correct, replaced by proc.name and proc.terminate (discovered this bug because we are using Option Explicit now). On Error Resume Next was hiding the bug too.
'                                - only terminate automation-processes (where command line contains --remote-debugging-pipe)
'-----------------------------------------------------------------------------------------------------------------------------------------

    '1. 今、Excelが動いてるプロセスIDとそのログインユーザーのセッションIDを取得します
    Dim lProcessId As Long, lSessionId As Long
    lProcessId = GetCurrentProcessId()          'Get the current(Excel) process ID
    ProcessIdToSessionId lProcessId, lSessionId 'Get the Session ID for the current(Excel) process

    'Clean up all running sessions
    Dim proc As Object
    On Error Resume Next

    '2. 前述の2つの情報を基に、起動中のアプリたちを取得する
    For Each proc In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process Where SessionId = " & lSessionId)
        '3. 起動中のブラウザ名であるか確認
        If LCase$(proc.Name) = LCase$(brType) Then
            '4. 自動化プロセスのみを終了します（コマンドラインに --remote-debugging-pipe が含まれている場合）
            If InStr(1, proc.CommandLine, "--" & ArgRemoteDebuggingName, vbTextCompare) > 0 Then proc.Terminate
        End If
        DoEvents
    Next
    On Error GoTo 0
    Err.Number = 0

    'Also clean up all prev CDP
    Call deleteCDPID

   'Clean core params for new session
    lngLastID = 0   'To make a clean slate for new session
    sessionID = ""  'Likewise

    printMsg "cleanUpSessions completed and previous CdpSessionID removed"
End Sub


Private Sub setCrashStateNormal()
'------------------------------------------------------------------
' Set Crashed state to Normal to prevent the "Restore Sessions?"
' popup from appearing every time the browser is started after
' being closed by browser.quit method.
' Requires user data dir profile stored in private var brUser and
' the browser type stored in private variable brType.
'------------------------------------------------------------------

   'Mark 'Normal' on the preference file to prevent Restore Page popup on start
   'This file locates at %localappdata%/Google/User Data/Default/Preferences
   Dim filePath As String, prefFolder As String, prefFile As Object, strContents As String, strNewContents As String, intF As Integer
    If brUser = "Default" Then prefFolder = brUser Else prefFolder = brUser & "\Default"
    Select Case brType
        Case brChrome: filePath = Environ("LOCALAPPDATA") & "\Google\Chrome\" & prefFolder & "\Default\Preferences"
        Case brEdge: filePath = Environ("LOCALAPPDATA") & "\Microsoft\Edge\" & prefFolder & "\Default\Preferences"
    End Select

    If LenB(Dir(filePath)) = 0 Then Exit Sub

    intF = FreeFile
    Open filePath For Binary As #intF
        strContents = Input$(LOF(intF), #intF)
    Close #intF

   'Precise replacement: only replace exit_type value
    strNewContents = Replace(strContents, """exit_type"":""Crashed""", """exit_type"":""Normal""", 1, -1, vbBinaryCompare)

    If strNewContents <> strContents Then
        intF = FreeFile
        Open filePath For Binary As #intF
            Put #intF, 1, strNewContents
        Close #intF
    End If

    printMsg "setCrashStateNormal completed"
End Sub


Private Function attachToTab(tabId As String) As String
'----------------------------------------------------------------------------------
' Attach to a specific tab with its tabId (a.k.a targetId)
' flatten: Enables "flat" access to the session via specifying sessionId attribute
' in the commands. Set to True by default to make attachToTab work.
' Returns: a session ID string that is needed for CDP automation.
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(tabId)
    params("flatten") = True
    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.attachToTarget", params)
    attachToTab = results("sessionId")
End Function


Private Function getTargetID() As String
'----------------------------------------------------------------------------------
' Retrieve the targetId of the current tab (a.k.k "target" in CDP Documentation).
' This targetId is useful for tab automation, such as the .closeTab method.
'----------------------------------------------------------------------------------

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.getTargetInfo")
    getTargetID = results("targetInfo").Item("targetId")
End Function

'***************************************************************************************************
'* 機能　　：指定のプロセスIDが生きているか判定します
'---------------------------------------------------------------------------------------------------
'* 返り値　：True       生きてる
'            False      お亡くなり
'* 引数　　：PID        プロセスID
'***************************************************************************************************
Private Function IsProcessRunning(ByVal pid As Long) As Boolean
    Dim wmi As Object
    Dim col As Object
    
    ' PIDが0以下なら論外
    If pid <= 0 Then
        IsProcessRunning = False
        Debug.Print "無効なプロセスIDです。新規で立ち上げます。"
        With ShSetting01_StartBrowser
            .Range(.UseRangeName(6, "CDPBrowser.IsProcessRunning")).value = True
        End With
        Exit Function
    End If

    On Error Resume Next
    ' WMIサービスに接続
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    ' そのPIDを持つプロセスを探すクエリ
    Set col = wmi.ExecQuery("Select ProcessId From Win32_Process Where ProcessId = " & pid)
    
    ' エラーがなくて、かつ見つかった数が0より大きければ生きている
    If Err.Number = 0 And col.Count > 0 Then
        IsProcessRunning = True
        Debug.Print "既存のデバッグブラウザプロセスを見つけました。再アタッチします。"
    Else
        IsProcessRunning = False
        Debug.Print "既存のデバッグブラウザプロセスが見つかりませんでした。新規で立ち上げます。"
        With ShSetting01_StartBrowser
            .Range(.UseRangeName(6, "CDPBrowser.IsProcessRunning")).value = True
        End With
    End If
    On Error GoTo 0
End Function

Private Function newVarID() As String
'----------------------------------------------------------------------------------
' Create a unique varID for element reference with CDPElement. If the ID is not
' unique, it is possible that new element assignment will override existing ones
' under the same ID by chance if working on a large number of elements.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Init the ID collection
    If varIDs Is Nothing Then Set varIDs = New Scripting.Dictionary
    Dim newID As String
    Do: newID = "varID" & Format(Rnd * 100000, "000000")
    Loop Until Not varIDs.Exists(newID)
    varIDs.Add newID, newID

   'Return the new varID
    newVarID = newID
End Function


Private Function getBrowserPath(strBrowserType As String) As String
'----------------------------------------------------------------------------------
' Obtain the installation path of the browser program. This is useful when
' the installation path is changed by the user or their corporation for security.
' Updated: 22/03/2023
' 20251023 Daniel Polak - added fallback path when registry lookup fails (in a locked-down environment)
'----------------------------------------------------------------------------------
    
    Select Case True
        Case InStr(LCase(strBrowserType), "chrome"): brType = brChrome
        Case InStr(LCase(strBrowserType), "edg"): brType = brEdge         'On edge it is just "Edg"
        Case Else: printMsg "Unable to getBrowserPath for browser type: " & brType, doRaiseError:=True
    End Select

   'Return the path
    On Error Resume Next
    Dim objWSHost As Object, reg As Object, strKeyPath As String
    Set reg = GetObject("winmgmts://./root/default:StdRegProv")
    
    'カスタムパスがあったらそっちを優先
    With ShSetting01_StartBrowser
        If .Range(.UseRangeName(1, "CDPBrowser.getBrowserPath.カスタムパス判定")).value = "" Then
            Set objWSHost = CreateObject("WScript.Shell")
            strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & brType & "\Path"
            getBrowserPath = objWSHost.regread(strKeyPath) & "\" & brType
        Else
            getBrowserPath = .Range(.UseRangeName(1, "CDPBrowser.getBrowserPath.カスタムパス判定")).value
        End If
    End With

    If Err.Number <> 0 Or LenB(getBrowserPath) = 0 Then
        Err.Clear
        Select Case brType
            Case brChrome
                If LenB(Dir(Environ("ProgramFiles") & "\Google\Chrome\Application\" & brType)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles") & "\Google\Chrome\Application\" & brType
                ElseIf LenB(Dir(Environ("ProgramFiles(x86)") & "\Google\Chrome\Application\" & brType)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles(x86)") & "\Google\Chrome\Application\" & brType
                End If
            Case brEdge
                If LenB(Dir(Environ("ProgramFiles") & "\Microsoft\Edge\Application\" & brType)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles") & "\Microsoft\Edge\Application\" & brType
                ElseIf LenB(Dir(Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\" & brType)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\" & brType
                End If
        End Select
        If LenB(getBrowserPath) > 0 Then
            If LenB(Dir(getBrowserPath)) = 0 Then getBrowserPath = ""
        End If
        If LenB(getBrowserPath) = 0 Then printMsg "Unable to locate " & brType & " installation path.", doRaiseError:=True
    End If

End Function

Private Sub getSessionInfo()
    '----------------------------------------------------------------------------------
    ' Obtain session info on the fly for printParams.
    ' Updated: 22/03/2023
    '----------------------------------------------------------------------------------
   'Activate to getHandle correctly
    activate

   'Get infos
    brURL = Url
    brName = Name
    brHWnd = getHandle
    brWID = invokeMethod("Browser.getWindowForTarget").Item("windowId")
    brType = invokeMethod("Browser.getVersion").Item("product")
    brVer = Split(brType, "/")(1)
    brPath = getBrowserPath(brType)
    targetID = getTargetID
End Sub

'***************************************************************************************************
'* 機能　　：デバッグブラウザのハンドル情報を消去します
'***************************************************************************************************
Private Sub deleteCDPID()
    With ShSetting01_StartBrowser
        .ListObjects(.UseTableName(1, "CDPBrowser.deleteCDPID")).DataBodyRange.Columns(2).Clear 'パイプ通信Handle情報
        .ListObjects(.UseTableName(2, "CDPBrowser.deleteCDPID")).DataBodyRange.Columns(2).Clear 'ブラウザセッション情報
    End With
End Sub



Private Sub Class_Initialize()
'----------------------------------------------------------------------------------
' A quick way to quickly check if an automated session is already active and open.
' If so, automatically attempt to reattach to it.
' Updated: 07/06/2023
'----------------------------------------------------------------------------------

    '設定シートの各セルから設定値を取得し、適用
    With ShSetting01_StartBrowser
        'ログに関する設定を行う
        doPrintDbgMsg = .Range(.UseRangeName(7, "Class_Initialize.設定シートからの起動")).value              'イミディエイトに表示させるか？
        seeRawSendMsgDbg = .Range(.UseRangeName(8, "Class_Initialize.設定シートからの起動")).value           'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
    
        Dim LogFolderPath As String: LogFolderPath = .Range(.UseRangeName(9, "Class_Initialize.設定シートからの起動")).value
        If LogFolderPath <> "" Then
            doLog = True            'ログファイルとして保存するか？
            logPath = LogFolderPath
        End If
    End With


   'Determine project host type for compatibility
    Select Case True
        Case InStr(Application.Name, "Excel"): Set thisApp = ThisWorkbook
        'Case InStr(Application.Name, "Word"): Set thisApp = ThisDocument
        'Case InStr(Application.Name, "Access"): Set thisApp = CurrentProject
    End Select

   'Assign an objID for easy debug identification
    brID = "BRID" & Format(Rnd * 1000, "000")
    
    '既存のデバッグブラウザプロセスの使用判定
    With ShSetting01_StartBrowser
        If .Range(.UseRangeName(5, "CDPBrowser.Class_Initialize")).value Then
            'プロセスIDチェック関数へ
            'Falseなら、ここで終り
            With .ListObjects(.UseTableName(1, "CDPBrowser.Class_Initialize")).DataBodyRange
                If Not (IsProcessRunning(WorksheetFunction.XLookup("dwProcessId", .Columns(1), .Columns(2), 0))) Then Exit Sub
            End With
        Else
            '設定側で、そもそも使用しないパターンは、ここで終り
            Exit Sub
        End If
    End With

   'Attempt to reattach and check if successful
    Call deserialize

    If Not isLive Then
        lngLastID = 0   'To make a clean slate for new session
        sessionID = ""  'No need to clean targetID as it will be retrieved anyway
        printMsg "Failed to reattach using the previous session parameters"
        Exit Sub
    End If
            
   'Obtain instance useful info if already live
    printMsg "Browser ID : " & brID & vbNewLine & _
             "Tab ID     : " & targetID, isHeader:=True
    printMsg brID & " successfully attached to Tab ID " & targetID
    getSessionInfo
        
End Sub

Private Sub SaveBase64ToFile(strB64 As String, strPath As String)
'----------------------------------------------------------------------------------
' Decode a base64 string and save to file without external dependencies (late bound).
'----------------------------------------------------------------------------------

    Dim dom As Object, el As Object, Bytes() As Byte, stm As Object

    Set dom = CreateObject("MSXML2.DOMDocument")
    Set el = dom.createElement("tmp")
    el.DataType = "bin.base64"
    el.Text = strB64
    Bytes = el.nodeTypedValue

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write Bytes
    stm.SaveToFile strPath, 2
    stm.Close

End Sub

'===============================================================================================
' Experimental Functions
'===============================================================================================
