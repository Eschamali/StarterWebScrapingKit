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
'       22/01/26 Long Vh:      - some tweaks to the .start method to improve the usage of user-data-dir switch
'                              - enhanced the framework with a new feature called multi-instances operation (see demo example demoMultiProfileOperation)
'                              - removed auto reattach feature from the framework and added .reattach method to suit
'                                the new multi-instances operation feature
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

Private Const CDPVer                    As String = "3.1.0"
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
Private Const ThisClassName As String = "CDPBrowser"                'ログのトレース用

Private Const ERR_TIMEOUT As Long = 900
Private Const ERR_PIPE As Long = 901
Private Const ERR_PROTOCOL As Long = 902
Private Const ERR_DictionaryStructure As Long = 903
Private Const ERR_UnregisteredProtocol As Long = 904
Private Const ERR_PolicyViolation As Long = 905
Private Const TIMEOUT_DEFAULT As Single = 30
Private Const TIMEOUT_NAVIGATE As Single = 45

'ブラウザからのイベント情報を格納する際のDictionaryKey名設定
Private Const EventsDictionaryKeyName01 As String = "TotalEvents"
Private Const EventsDictionaryKeyName02 As String = "EventMethods"

'===================================
' Global Variables
'===================================

Private jsConverter As New WebJsonConverter 'Excelで、JsonデータをDictionary(Collection)に変換できるすごいやつ
Private LogControl As New Logger        'ログレベルの制御
Private core As CDPCore                 'Must be initialized for CDP engine
Private ErrorDetail As New WinApiError  'エラーコードから、詳細を取得するやつ

Private lngLastID As Long               'Every message sent to edge has an ID
Private strBuffer As String             'ストリーム受信で欠けたデータを一時的に蓄積するための変数です
Private brHWnd As LongPtr               'Browser object window handle
Private brWID As String                 'Useful for windows repositioning
Private brType As String                'Browser type: either 'chrome' or 'edge'
Private brUser As String                'Browser user profile
Private brPath As String                'Browser default installation path
Private brID As String                  'The ID of the current CDPBrowser object. Useful when reading debugs of multiple tabs.
Private brName As String                'The focused tab name / window name of the active CDPBrowser
Private brURL As String                 'The focused tab url of the active CDPBrowser
Private brVer As String                 'The current browser version
Private brReAttach As Boolean           'Remember reAttach preference from .reattach method - updated by Long 22/01/2026
Private brCallString As String          'The initial call string used by .start
Private varIDs As Scripting.Dictionary  'Collection of all unique element IDs for CDPElement reference - Added 02/03/2023
Private sessionID As String             'The ID to the current session tab.
Private targetID As String              'The targetID of the current tab for InvokeMethod
Private thisApp As Object               'A dymanic variable used in CDPBrowser initialize procedure to determine if the VBA host is an Excel / Word / Access
Private objAllMessages As Dictionary    'ブラウザからのイベントデータを蓄積するための変数です
Private timerOut As Long                'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
Private seeRawSendMsgDbg As Boolean     'set to True if needing to see the raw json msg received from the browser for debugging
Private doLog As Boolean                'ログファイルにも、デバッグメッセージを残すか
Private logPath As String               'ログファイルの保存フォルダパス
Private errorDetails As Dictionary      'CDP-Jsonコマンド実行時の最後のエラー情報
Private AccumulatedCommandResults As Dictionary '非同期によるコマンド結果情報
Private MaxCommandResults As Long       '非同期によるコマンド結果の最大蓄積数

Private Policy_UserDataDirExists    As Boolean  'ユーザーデータディレクトリが固定化されてるか判定
Private Policy_UserDataDirContents  As String   'ポリシーによって設定されたデータディレクトリパス

'=============================================================================
' PUBLIC FUNCTIONS
'=============================================================================

Public Sub start(Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional cleanActive As Boolean = True, _
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
'          22/01/26 Long Vh:      - Removed automatic reattachment from the start method as this restricts
'                                   the possibility of running multiple CDP browser instances under different
'                                   profiles - which is a very useful feature.
'                                 - Tweaked the addProfile section and the cleanUpSessions to reflect the
'                                   recent Chromium updates and the Edge policy discovery.
'--------------------------------------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".start"


    Dim strCall As String, addProfile As String
    Dim intRes As Long
    Dim timeStart As Single

    timeStart = Timer

   'Determine default browser parameters -------------------------------------------------------------------------------------
   '07/01/26: avoid using --enable-automation tag as it triggers many web apps to perform bot check
    Select Case Name
        Case "edge": brType = brEdge
        Case "chrome": brType = brChrome
        Case Else: printMsg ERROR_, "Browser initiation fails: only the defined browser types can be accepted", FromProcedureName, doRaiseError:=True
    End Select
    brPath = getBrowserPath(brType)
    strCall = """" & brPath & """ --" & ArgRemoteDebuggingName
   '--------------------------------------------------------------------------------------------------------------------------

   'Add starting URL to the call string -------------------------------------------------------------------------------------------------------------------
    If appUrl <> "" Then
        If appUrl = "about:blank" Then printMsg ERROR_, "Browser cannot be open under Application Mode when homepage is set to ""about:blank""", FromProcedureName, doRaiseError:=True
        strCall = strCall & " --app=""" & checkURI(appUrl) & """"
    Else
        strCall = strCall & " --homepage ""about:blank"""
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------------------------

   'Start with the choice profile -------------------------------------------------------------------------------------------------------------
   '22/01/26: I have recently noted that if an organization sets the group policy of Edge of UserDataDir to a specific structure, the argument --user-data-dir
   'will be effectively disabled and blocked to the specified data dir by the policy. In this scenario, CDP session can only load the browser state from the default
   'data dir which already cannot be allowed to run --remote-debugging-pipe from Google's latest Chromium. If Edge's Chromium updates to the latest, this policy will
   'cause the CDP initialization to fail. If Edge's Chromium is still behind the latest Chromium version, CDP init can still be possible with the --profile-directory
   'switch but the requirement is that all running Edge instances that had been started up manually will need to be closed first otherwise the PeekNamedPipe error
   'will result. For Chrome, v136 and later does not honor remote-debugging request if user profile is loaded from the Chrome\User Data\ folder.
   'As a result, here we set it to Chrome\CDP\ folder instead (auto-created if not exists).
    Select Case brType
        Case brEdge: addProfile = IIf(Policy_UserDataDirExists, " --profile-directory=""" & userProfile & """", " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Microsoft\Edge\" & userProfile & """")      'ポリシー状況に応じて、引数を変える
        Case brChrome: addProfile = IIf(Policy_UserDataDirExists, " --profile-directory=""" & userProfile & """", " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Google\Chrome\" & userProfile & """")     'ポリシー状況に応じて、引数を変える
    End Select
    strCall = strCall & addProfile
    brUser = userProfile 'store the user profile name for cpid serialization
   '-----------------------------------------------------------------------------------------------------------------------------------------

   'Clean up previous sessions if they could not be reattached to avoid PeekNamedPipe error --------------------------------------------
   'Note: The error usually occurs when the same User Profile is already being used by another pipe. Hence if this
   'session is unable to connect to the previous pipe (for instance, due to the session application host was abruptly terminated),
   'PeekNamedPipe error will be thrown.
    If cleanActive Then cleanUpSessions onlyCdpSessions:=Not (Policy_UserDataDirExists)  'Note: you can set onlyCdpSessions to True by default if your organization does not force UserDataDir policy
    setCrashStateNormal 'Prevent "Restore Session" popup
   '------------------------------------------------------------------------------------------------------------------------------------

   'Prepare for initialization ------------------------------------------------------------------------
    strCall = strCall & " " & addArgs  'Add additional arguments if instructed
   '---------------------------------------------------------------------------------------------------

   'Initialize CDPCore class ----------------------------------------------------------------------------------------------------------
    Set core = New CDPCore
    brCallString = strCall
    printMsg Debug_, "Starting pipe communication... → " & brCallString, FromProcedureName
    intRes = core.Init(brCallString)  'Serializable is set to True by default for ease of accessibility
    If intRes <> 0 Then printMsg ERROR_, "Unable to initialize CDPCore. ErrorCode:" & intRes & " - " & ErrorDetail.GetMessage(intRes, "kernel32"), FromProcedureName, doRaiseError:=True Else printMsg info_, "CDPCore initialized successfully", FromProcedureName
    lngLastID = 1 'reset the pipe message ID
   '-----------------------------------------------------------------------------------------------------------------------------------

   'Complete initialization ------------------------------------------------------
    attachToSession     'Find and attach the newly open Session ID
    getSessionInfo      'Store initial session info for printParams debugging
    serialize           'Create a serial string for future attachment
   '------------------------------------------------------------------------------

   'Make sure the page is fully loaded for inAppMode ------------------------------------------------------------------------------------
   '07/01/25: Long Vh: simplified the check as there are web apps that does not end up loading the target url but points to another url.
   'As a result, the below check simply checks if there is "movement" in the name of the page instead
    Dim timerStart As Single, strLoc As String
    If appUrl <> "" Then
        timerStart = Timer
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
    printMsg info_, "Browser initialization completed after " & Timer - timeStart & "s", FromProcedureName
   '------------------------------------------------------------------------------

End Sub

Private Sub serialize()
'------------------------------------------------------------------
' This method transcribes Chrome's Json message into a string.
' The serialized string can then be used to automate an active
' session with a known serial.
' Updated: 03/05/23 - Replace sessionID with targetID since it is
'                     constant and stable on every CDP session
'          22/01/26 - Added specification to user profile name to
'                     the cpid file name to allow targetting of
'                     specific profile for reattachment.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".serialize"


    'データ記録準備
    Dim excelSerialize As String
    With WorksheetFunction
        excelSerialize = "{""" & .TextJoin(""";""", False, _
                                    .TextJoin(""",""", False, brUser, "lngLastID", lngLastID), _
                                    .TextJoin(""",""", False, brUser, "sessionID", sessionID), _
                                    .TextJoin(""",""", False, brUser, "targetID", targetID) _
        ) & """}"
    End With

    'パイプ通信のハンドル情報も記録しておき、結果を表示しておく
    printMsg info_, "New serial string: " & core.serialize(brUser) & " - " & excelSerialize, FromProcedureName
 
    'テーブルに、ブラウザセッション情報を記録する
    'Evaluate関数を利用して、上記で作った記述方式で簡易的な2次元配列を作成します
    ShSetting01_StartBrowser.serializeToTable(2, FromProcedureName) = Evaluate(excelSerialize)

End Sub

Private Sub deserialize(UserName As String, detachCurrent As Boolean)
'------------------------------------------------------------------
' This method deciphers a serialized string to obtain the currently
' active session ID of the already open instance. The serial
' must be obtained first from the serialize method. The isLive
' public function checks if reattachment is successful.
' Updated: 03/05/23 - replace sessionID with targetID which is
'                     constant and stable on every CDP session
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".deserialize"


    On Error GoTo iferr

    Dim objSerialize As Scripting.Dictionary
    Set core = New CDPCore

    '`UserName`に紐づくパイプ通信のハンドル情報の読み込みへ移る
1   Call core.deserialize(UserName)

    'Excelテーブルから、`UserName`に紐づくブラウザセッション情報を読み込む
2   Set objSerialize = ShSetting01_StartBrowser.deserializeFromTable(UserName, 2, FromProcedureName)

    'クラス内変数に適用
3   lngLastID = objSerialize.Item("lngLastID")
4   targetID = objSerialize.Item("targetID")

    '旧SessionIDの破棄を試みる
    '※`getTab`,`newTab`からの場合は破棄すると、親子関係Sessionの影響で道ずれ破棄されるので、破棄しないように引数で制御します
    If detachCurrent Then
        Dim ResultDetach As Dictionary, CDPparameters As Dictionary
5       Dim oldSessionID As String: oldSessionID = objSerialize.Item("sessionID")
        Set CDPparameters = New Dictionary

        CDPparameters.Add "sessionId", oldSessionID
        printMsg Debug_, "Discard the old session ID... - " & oldSessionID, FromProcedureName
6       Set ResultDetach = invokeMethod("Target.detachFromTarget", CDPparameters)

        '結果を判定
        '※エラーによるエラーメッセージは、`invokeMethod`が担います
7       If Not (ResultDetach Is Nothing) Then printMsg info_, "The old session ID has been detached. - " & oldSessionID, FromProcedureName

        'SessionIDを再生成し、適用
8       sessionID = attachToTab(targetID)
        printMsg info_, "A new session ID has been assigned. - " & sessionID, FromProcedureName

        '記録しておく
        '※旧SessionIDの破棄をしない場合は、上書きせずそのままにします
9       ShSetting01_StartBrowser.serializeSingleToHandleInfoTable(UserName, Table02RowID.sessionID, FromProcedureName) = sessionID
    Else
        '旧SessionIDを再利用
        sessionID = objSerialize.Item("sessionID")
        printMsg info_, "Reattached by reusing the Session ID. - " & sessionID, FromProcedureName
    End If
    Exit Sub

iferr: 'if deserialize fails, simply move on to starting anew

    printMsg ERROR_, "deserialization failed at line " & Erl & " for serial string " & WorksheetFunction.TextJoin(",", True, "lngLastID:" & lngLastID, "targetID:" & targetID, "sessionID:" & sessionID), FromProcedureName

End Sub

Public Function jsEval(strString As String, Optional dbgMsg As Boolean = True) As Variant
'------------------------------------------------------------
' This function commands the attached browser instance to
' execute a Javascript expression to perform required tasks.
' Note: the original function has a bool retry variable which
' was removed for simplicity.
' Added dbgMsg to read debug line on the go when needed.
'------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".jsEval"


    Dim results As Scripting.Dictionary
    Dim strMessage As String
    Dim objMessage As Scripting.Dictionary

step1:

   'Construct message for sending ----------------------------------------------------------
    strMessage = "{""method"":""Runtime.evaluate"",""params"":{""expression"":""1+1;""}}"
    Set objMessage = jsConverter.ParseJson(strMessage)
    objMessage.Item("params").Item("expression") = strString & ";"
    strMessage = jsConverter.ConvertToJson(objMessage)
   '----------------------------------------------------------------------------------------

   'Send and check result -------------------------------------------------------------------------------------------
    Set results = SendMessage(strMessage, dbgMsg:=dbgMsg)
    If results Is Nothing Then printMsg ERROR_, "Unable to contact the Browser instance!", FromProcedureName, doRaiseError:=True
    If Not results.Exists("error") Then GoTo step2
   '-----------------------------------------------------------------------------------------------------------------

   'Handle errors ---------------------------------------------------------------------------------------------------
    Dim errMsg As String, retried As Boolean
    errMsg = results("error")("message")
    printMsg WARN_, "jsEval Error: " & errMsg, FromProcedureName
    If errMsg = "'Runtime.evaluate' wasn't found" Then printMsg WARN_, "Possibly the main session has been abruptly closed.", FromProcedureName
    If errMsg = "Session with given id not found." And Not retried Then
        printMsg info_, "Attempting to update Session ID " & sessionID, FromProcedureName
        sessionID = "": sessionID = attachToTab(targetID)
        printMsg info_, "New Session ID obtained: " & sessionID, FromProcedureName
        retried = True: GoTo step1
    End If
    printMsg ERROR_, "The Browser was unable to evaluate the JS expression: " & strString, FromProcedureName, doRaiseError:=True
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
    printMsg Debug_, "Script executed: " & strString & strResult, FromProcedureName
   '---------------------------------------------------------------------------------------------------------

End Function

Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
'-----------------------------------------------------------------------
' This sub is a quick wrapper to navigate to an url.
' Updated 23/03/23
'-----------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".navigate"


   'Check if already navigated
    Dim curURL As String
    curURL = jsEval("window.location.href", False)
    If curURL = checkURI(strURL) Then
        printMsg WARN_, "Already navigated to """ & strURL & """", FromProcedureName
        Exit Sub
    End If

   'If not yet then do navigate
    Dim results As Scripting.Dictionary
    Set results = SendMessage("{""method"":""Page.navigate"",""params"":{""url"":""" & strURL & """}}")

    'should also add a maxTimeLimit to Wait
    wait till:=till
    getSessionInfo 'update new session info for printParams
    printMsg info_, "Navigated to: " & strURL, FromProcedureName

End Sub

Public Sub quit()
'-----------------------------------------------------------------------------------
' Closes the currently attached browser instance
' Original Name: closeBrowser
' Update: 24/10/2025: Daniel Polak - avoid immediate hard kills, reduce race conditions, and make shutdown behavior predictable and robust in non-responsive cases
'-----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".quit"


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
        printMsg WARN_, "Browser didn't close gracefully within 5s; invoking cleanUpSessions.", FromProcedureName
        cleanUpSessions
    End If

    'パイプハンドル、ブラウザセッション情報を消去しに行きます
    deleteCDPID

    sleep 1 'it seems without waiting a bit the browser crashes and the next time it wants to recover from a crash
    printMsg info_, "The browser has been terminated.", FromProcedureName

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
    Const FromProcedureName As String = ThisClassName & ".wait"


    Dim strState As String, targetState As String
    Select Case till
        Case isLoading: targetState = "loading"
        Case isInteractive: targetState = "interactive"
        Case isComplete: targetState = "complete"
    End Select

    '初期値を設定しておく
    '最初はほぼ必ずloadingだから、`targetState`が`loading`になってたら、わざわざチェックする必要性はない
    strState = "loading"

    Do Until strState = targetState
        strState = jsEval("document.readyState", dbgMsg:=False)                 'dbgMsg = False to not spam Immediate Window
        printMsg Debug_, "Current ReadyState: " & strState, FromProcedureName   'option to print ReadyState info during iteration
        If targetState = "interactive" And strState = "complete" Then Exit Do   'also exit if better than interactive is satisfied
        sleep 0.1                                                               'reduce sleep will speed up but will cost cpu power
    Loop

    printMsg Debug_, "Wait '" & targetState & "' satisfied", FromProcedureName

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
    ShowWindow brHWnd, doShow 'Make visible
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
'          22/01/26 Added .reattach where appropriate
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".newTab"


    Dim params As New Scripting.Dictionary
    Dim sessionID_ As String

    params("url") = CStr(checkURI(Url))
    If newWindow Then params("newWindow") = CBool(newWindow)

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.createTarget", params)

   'Return the new tab's Session ID
    Dim tabId As String
    tabId = results("targetId")
    sessionID_ = attachToTab(tabId)

   'Switch attachment to the new tab
   '22/01/26: From v3.1, reattachment is no longer automatic so .reattach is needed
    printMsg info_, "newTab will now open a new CDP instance", FromProcedureName
    Set newTab = New CDPBrowser
    newTab.reattach brUser, False   '親子Session関係により、旧SessionIDは破棄しないようにする
    newTab.id = sessionID_          '新しくタブ生成した際の`SessionID`として、改めて割り当て
    If Url <> "" Then newTab.wait

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = sessionID_
    getSessionInfo
    serialize
    printMsg info_, "Tab [" & brName & " | " & targetID & "] is now the main session tab for " & brID, FromProcedureName

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
'* 更新履歴：06/06/23               - Added doRetry argument
'            03/05/23               - Combined the method with the getTabNew
'            24/10/25 Daniel Polak  - Predictable, precise, and script-friendly tab selection that scales well when there are many similar tabs open
'                                     Exact title and URL-prefix matching avoids accidentally selecting tabs with similar titles/URLs
'            23/10/25 Daniel Polak  - Fixed typo timeStart instead of timerStart and fixed bug 60000s instead of 60s
'            01/08/26 Eschamali     - iframeへの切り替えができるように改良
'            22/01/26 Long Vh       - Added .reattach where appropriate
'---------------------------------------------------------------------------------------------------
'* 機能説明：tabName引数とurl引数は部分的な指定(部分一致)が可能です。SearchType はデフォルトは、"page"にしてます。
'* 注意事項：これらの引数を両方とも省略すると、getTabは最も近い未接続のタブに接続しようとします。
'***************************************************************************************************
Public Function getTab(Optional tabName As String, Optional Url As String, Optional setMain As Boolean, Optional SearchTypeID As TargetgetTargetsType, Optional doRetry As Boolean = False) As CDPBrowser
    Const FromProcedureName As String = ThisClassName & ".getTab"


    Dim results As Scripting.Dictionary, Obj As Scripting.Dictionary
    Dim tabId As String, candidateTabId As String, errMsg As String
    Dim targets As Collection
    Dim timerStart As Single
    Dim tabfound As Boolean
    Dim bestScore As Long, score As Long
    Dim info As String

    If doRetry Then timerStart = Timer

    'Typeを決める
    Dim SearchType As String
    Select Case SearchTypeID
        Case TargetgetTargetsType.Page: SearchType = "page"
        Case TargetgetTargetsType.iFrame: SearchType = "iframe"
        Case Else: Err.Raise vbObjectError + 1, FromProcedureName, "このパラメーターは、未定義です"
    End Select

searchit:

    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

    'スコア方式を利用して、理想の対象タブ情報を取得します
    Const MaxScore As Long = 100
    bestScore = -1
    candidateTabId = ""
    For Each Obj In targets
        If Obj("type") = SearchType Then
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
        printMsg WARN_, "Unable to find the tab: " & info, FromProcedureName
        Exit Function
    Else
        printMsg info_, "Found the tab with ID: " & tabId & ". (Certainty：" & bestScore & "%)", FromProcedureName
    End If

searchfound:

    Dim sessionID_ As String
   'Return the session object for reference
    sessionID_ = attachToTab(tabId)

   'Return the new CDP object
   '22/01/26 .reattach brUser is needed from v3.1 onwards to enable multi-instances automation
    Set getTab = New CDPBrowser
    getTab.reattach brUser, False   '親子Session関係により、旧SessionIDは破棄しないようにする
    getTab.id = sessionID_  'will also update the targetID
    sleep 'give some time so that the debugging session can switch to the new tab

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = sessionID_
    getSessionInfo
    serialize
    printMsg info_, "Tab [" & targetID & "] is now the main session tab for " & brID, FromProcedureName

End Function

Public Sub closeTab()
'------------------------------------------------------------------
' Close the current tab of the session. Requires targetId which can
' be retrieved by using getTargetId. This is the targetId of the
' active tab under focus of the automation pipe.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".closeTab"


   'Get the original serialized tab ID
    Dim mainID As String, serial As Dictionary
    Set serial = ShSetting01_StartBrowser.deserializeFromTable(brUser, 2, FromProcedureName)
    mainID = serial.Item("targetID")

    '操作中のタブ自体を閉じないように仕向ける
    If targetID = mainID Then
        printMsg WARN_, "Unable to close this tab " & targetID & " as it is the main session tab", FromProcedureName
        Exit Sub
    End If

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    invokeMethod "Target.closeTarget", params
    printMsg info_, brID & " was closed", FromProcedureName

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
    Const FromProcedureName As String = ThisClassName & ".jsAddLib"


   'Check if already loaded, if so then no need to add again
    If getElementByQuery("script[src='" & strURL & "']").varResult <> "null" Then
        printMsg WARN_, "The HTML already contains the JS Library: """ & strURL & """", FromProcedureName
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
    printMsg info_, "Successfully added the JS Library: """ & strURL & """", FromProcedureName

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
    Const FromProcedureName As String = ThisClassName & ".jsAddScript"


    If Dir(strLocalScriptPath) = "" Then Exit Sub

    Dim fso As New FileSystemObject
    jsEval fso.OpenTextFile(strLocalScriptPath).ReadAll 'Load the script contents
    printMsg info_, "Successfully added the local script: """ & strLocalScriptPath & """", FromProcedureName

    Set fso = Nothing

End Sub

'***************************************************************************************************
'* 機能　　：Excelテーブルにある既存のパイプハンドル、ブラウザセッション情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功可否論理値
'* 引数　　：userProfile    再アタッチしたいユーザー名(user-data-dirに基づく)
'            detachCurrent  True ：Excelに記録してるSessionIDを破棄し、再更新したのに置き換えます(規定値-推奨)
'                           False：SessionIDを再利用します。基本的に非推奨。内部クラスの処理都合で利用することがあります。
'---------------------------------------------------------------------------------------------------
'* 機能説明：Reattach to the target CPID file. This method enables the management of multiple
'            CDP instances at once under different Chrome/Edge user profiles. The benefit of
'            this is that if one instance fails/crashes, it does not affect the other ones.
'            and thus enables the possibility of running multiple CDP automation operations
'            on entirely different tasks simultaneously.
'* 注意事項：`detachCurrent`の`False`は、下記の内部処理にて使用する想定です。
'            親子関係SessionIDを維持するために破棄しないようにしてます。それ以外は基本的に`True`で、SessionIDの更新で`reattach`するようにしてください。
'            ・newTab
'            ・getTab
'***************************************************************************************************
Public Function reattach(userProfile As String, Optional detachCurrent As Boolean = True) As Boolean
    Const FromProcedureName As String = ThisClassName & ".reattach"


    brUser = userProfile
    deserialize brUser, detachCurrent

    If isLive Then
        reattach = True
        printParams
    Else
        reattach = False
        printMsg WARN_, "Failed to reattach to previous CPID.", FromProcedureName
    End If

End Function

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

Public Sub snapPage(FolderPath As String, FileName As String, Optional getFullPage As Boolean = False)
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
    Const FromProcedureName As String = ThisClassName & ".snapPage"


    Dim params As New Scripting.Dictionary
    Dim r As Scripting.Dictionary
    Dim b64 As String, tmpPng As String

   'Verify if the page can be captured with screenshot
    Set r = invokeMethod("Page.getLayoutMetrics")
    If invokeError(r) Then
        printMsg WARN_, "Unable to perform snapPage on this screen as it has no layout metrics.", FromProcedureName
        Exit Sub
    End If

   'Perform the snip
    If getFullPage Then params.Add "captureBeyondViewport", True
    Set r = invokeMethod("Page.captureScreenshot", params)

   'Save to the same folder if not indicated
    If Not invokeError(r) Then
        b64 = r("data")
        SaveBase64ToFile b64, FolderPath, FileName
        printMsg info_, "Screenshot saved to: " & FolderPath & "\" & FileName, FromProcedureName
    End If

End Sub

Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
'----------------------------------------------------------------------------------
' A custom printMsg that besides printing to the Immediate Window, it will
' also print to a dated log file. This is useful at user level where
' the Immediate Window is not easily accessible. doRaiseError will also raise an
' error aside from printing the error message to the log file.
' Updated: 07/01/26 Long Vh: the message output is disabled by default to avoid
'                            spamming info. This can be turned on by setting globally
'                            noDbgMsg = False or noLogging = False
'----------------------------------------------------------------------------------

    Dim logName As String, strFormattedMsg As String

    If isHeader Then strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-") _
    Else strFormattedMsg = " | " & brID & " | " & strMsg

   'Perform msg logging if needed
    If doLog Then
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".cdp.log"
        If Dir(logPath, vbDirectory) = "" Then MkDir logPath
    End If

   'Perform msg output to Immediate Window if needed
    With LogControl
        Select Case LogLevel_
            Case LogLevelName.Trace_: .LogTrace strFormattedMsg, From, logPath, logName
            Case LogLevelName.Debug_: .LogDebug strFormattedMsg, From, logPath, logName
            Case LogLevelName.info_: .LogInfo strFormattedMsg, From, logPath, logName
            Case LogLevelName.WARN_: .LogWarn strFormattedMsg, From, logPath, logName
            Case LogLevelName.ERROR_: .LogError strFormattedMsg, From, logPath, logName, Err.Number
            Case Else: 'ログ出力無効化
        End Select
    End With


    'ブラウザ自身や、パイプ通信自体に問題や、タイムアウトになったと判断された場合は、このツールでの処理(CDP-Json)ができないので、エラーで停止させておく
    If doRaiseError Then Err.Raise ERR_PROTOCOL, From, Description:=strMsg

End Sub

Public Sub printTabs()
'------------------------------------------------------------------
' Print all current tabs info. Useful for debugging.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".printTabs"


    Dim results As Scripting.Dictionary, target_ As Scripting.Dictionary
    Dim i As Long
    Dim targets As Collection
    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

   'Read info
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, "Tabs Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    i = 1
    For Each target_ In targets
        If target_("type") = "page" Then
            printMsg Debug_, "#" & i & vbTab & _
                        "type: " & target_("type") & " | " & _
                        "attached: " & target_("attached") & " | " & _
                        "title: " & target_("title") & " | " & _
                        "url: " & target_("url") & " | " & _
                        "target ID: " & target_("targetId"), FromProcedureName
            i = i + 1
        End If
    Next

End Sub

Public Sub printParams()
'------------------------------------------------------------------
' Print all parameters of the current session for debuggings.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".printParams"


    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, brID & " Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "Session ID        | sessionID    | " & sessionID, FromProcedureName
    printMsg Debug_, "Last ID           | lngLastID    | " & lngLastID, FromProcedureName
    printMsg Debug_, "Window Handle     | brHWnd       | " & brHWnd, FromProcedureName
    printMsg Debug_, "Window ID         | brWId        | " & brWID, FromProcedureName
    printMsg Debug_, "Browser Type      | brType       | " & brType, FromProcedureName
    printMsg Debug_, "Browser Ver       | brVer        | " & brVer, FromProcedureName
    printMsg Debug_, "Browser Path      | brPath       | " & brPath, FromProcedureName
    printMsg Debug_, "Tab URL           | brUrl        | " & brURL, FromProcedureName
    printMsg Debug_, "Tab Title         | brName       | " & brName, FromProcedureName
    printMsg Debug_, "Tab ID            | targetID     | " & targetID, FromProcedureName
    printMsg Debug_, "Core StdInWr      | hStdInWr     | " & core.hStdInWr, FromProcedureName
    printMsg Debug_, "Core StdOutRd     | hStdOutRd    | " & core.hStdOutRd, FromProcedureName
    printMsg Debug_, "Core hStderrOutRd | hStderrOutRd | " & core.hStderrOutRd, FromProcedureName
    printMsg Debug_, "Core CDPInWr      | hCDPInWr     | " & core.hCDPInWr, FromProcedureName
    printMsg Debug_, "Core CDPOutRd     | hCDPOutRd    | " & core.hCDPOutRd, FromProcedureName
    printMsg Debug_, "Core dwProcessId  | dwProcessId  | " & core.dwProcessId, FromProcedureName
    printMsg Debug_, "CDP Version       | CDPVer       | " & CDPVer, FromProcedureName

End Sub

Public Function tabCount() As Long
'------------------------------------------------------------------------------
' Print all current tabs info. Useful for debugging.
' Updated : 05/02/25 - changed from printTabs to tabCount to get the count too
'------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".tabCount"


    Dim results, targets, target_ As Scripting.Dictionary
    Dim i As Integer

    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

   'Read info
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, "Tabs Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    i = 0
    For Each target_ In targets
        If target_("type") = "page" Then
            printMsg Debug_, "#" & i & vbTab & _
                        "type: " & target_("type") & " | " & _
                        "attached: " & target_("attached") & " | " & _
                        "title: " & target_("title") & " | " & _
                        "url: " & target_("url") & " | " & _
                        "target ID: " & target_("targetId"), FromProcedureName
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
'* 機能　　：CDPのMethod名と、parametersを渡して、低レベルにてブラウザ制御をします
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド の実行結果(Json-Dictionary 形式)。失敗時は、`Nothing`となります。
'* 引数　　：methodName             method名。"Network.getCookies"や"DOM.getDocument"というのを記入します
'            params                 Json-Dictionary形式で、KeyとItemを事前に定義したオブジェクト変数を指定します。
'            alwaysBrowserContext   `True` ：`sessionID`の有無問わず、常にブラウザ自身を対象したコマンド実行モードとなります
'                                   `False`：`sessionID`があれば、その`SessionID`に対するコマンド実行。なければ、ブラウザ自身を対象したコマンド実行となります。
'
'            StopError              `True`で、実行失敗時に`Err.Raise`で停止し、以降の処理を止めます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：This function enables advanced method invocation under CDP.
'            For a list of all methods available, see:
'               https://chromedevtools.github.io/devtools-protocol/tot/Browser/
'            Error message will be prompted automatically by invokeError in the Immediate Window.
'            The return object is Json String that has been converted to a dictionary object.
'            Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'
'* 更新履歴：・18/02/26 - SessionIDがあっても、BrowserContext優先にコマンド実行できるように改良。エラーで止めるかのスイッチングを導入
'            ・28/04/23 - Set this to Public as it is useful
'            ・20/03/23 - Added printMsg msgs & set to Private
'
'* 注意事項：・CDPの運営思想の都合上、多数の`Deprecated`,`Experimental`が存在します。
'              これらのタグのmethodを使用する際は、`StopError`を`True`にすることを推奨します。失敗時の以降の処理の予期せぬデータ破損などを防ぐためです。
'
'            ・非同期イベントの確認をする際は、空打ちによる継続確認が必要なのでこの場合は、`StopError`を`False`にしてください。
'              これは、VBAのシングルスレッド制約の中での苦肉の策です。
'
'            ・`StopError`の規定値は、`True`なので、特別な理由がない限りは、省略推奨です
'            ・このメソッドは、コマンド結果が来るまで待機するメソッドです。送るコマンド内容や、サイトの挙動によっては、タイムアウトエラーになる場合があります。
'              その場合は、後述の結果を待たないタイプのメソッド`invokeMethodAsync`をご利用ください。
'***************************************************************************************************
Public Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Const FromProcedureName As String = ThisClassName & ".invokeMethod"


    Dim paramStr As String
    Dim results As Scripting.Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = jsConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}", alwaysBrowserContext)

   'Print error if there is
    If invokeError(results) Then
        'エラーによる`Err.Raise`で止めるか？
        If StopError Then
            'エラーになったmethod名と、結果生responseを表示して停止
            printMsg ERROR_, "Failed to invoke method """ & methodName & """" & vbCrLf & vbCrLf & "=== RawErrorResponse ===" & vbCrLf & jsConverter.ConvertToJson(results), FromProcedureName, doRaiseError:=True
        Else
            'エラーになったmethod名だけログ出して続行
            printMsg WARN_, "Failed to invoke method """ & methodName & """", FromProcedureName
        End If
    End If

   'Return result object
    If results.Exists("result") Then Set invokeMethod = results.Item("result")
    printMsg Debug_, "Invoked """ & methodName & """ with result: " & jsConverter.ConvertToJson(results), FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`invokeMethod`の非同期実行(コマンド結果を待たない)バージョンです
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド 実行時のID値(整理券番号のイメージでOK)
'* 引数　　：※`invokeMethod`と一緒のため、省略
'***************************************************************************************************
Public Function invokeMethodAsync(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Long
    Const FromProcedureName As String = ThisClassName & ".invokeMethodAsync"


    Dim paramStr As String
    Dim results As Scripting.Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = jsConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}", alwaysBrowserContext, Async:=True)

    '非同期の場合は、`WriteFile`のエラーをチェックするだけとなります
    '※1件しかAddしてないので、直Index指定する
    If results.Items(0) Then
        'エラーによる`Err.Raise`で止めるか？
        If StopError Then
            'エラーになったmethod名と、`WriteFile`の実行エラーコードを表示して停止
            printMsg ERROR_, "Failed to invokeAsync-WriteFile method """ & methodName & """" & vbCrLf & vbCrLf & "=== WriteFileErrorDetail ===" & vbCrLf & "ErrorCode:" & results.Items(0) & " - " & ErrorDetail.GetMessage(results.Items(0)), FromProcedureName, doRaiseError:=True
        Else
            'エラーになったmethod名だけログ出して続行
            printMsg WARN_, "Failed to invokeAsync-WriteFile method """ & methodName & """　ErrorCode:" & results.Items(0) & " - " & ErrorDetail.GetMessage(results.Items(0), "kernel32"), FromProcedureName
        End If
    End If

    '実行したIDを返却
    invokeMethodAsync = results.keys(0)
    printMsg Debug_, "AsyncInvoked """ & methodName & """ with WriteFileResult: " & results.Items(0) & " - " & ErrorDetail.GetMessage(results.Items(0), "kernel32"), FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：パイプから、メッセージを取り出し、引数に応じて、非同期で実行したIDの結果やイベントを別途格納します
'---------------------------------------------------------------------------------------------------
'* 引数　　：destruction   `True` :IDの解析やイベントの中身の解析を行わず、破棄します
'                          `False`:IDの解析やイベントの中身の解析を行います
'***************************************************************************************************
Public Sub TakeEvents(Optional destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakeEvents"


    Dim ErrorCode_PeekNamedPipe As Long, ErrorCode_ReadFile As Long
    Dim intRes As Long
    Dim strRes As String
    printMsg Debug_, "Starts retrieving data from the pipe...", FromProcedureName

    '終端null文字があって尚且つ、受信データ0が来るまで、ループ受信を行います
    Do
        '受信処理
        intRes = core.readProcCDP(strRes, ErrorCode_PeekNamedPipe, ErrorCode_ReadFile)
        printMsg Trace_, "ResSize = " & intRes, FromProcedureName

        '返り値チェック
        If intRes > 0 Then
            '破棄モードかの判定
            If destruction Then
                '解析などせず、そのまんま出して破棄します
                '※強いていうなら、破棄するメッセージをそのまま出す
                printMsg Trace_, "Discarding data in pipe...", FromProcedureName
                If seeRawSendMsgDbg Then printMsg Trace_, "Discarded - strRes = " & strRes, FromProcedureName
            Else
                '中身の解析をし、IDやイベントを取り出します
                printMsg Trace_, "Analyzing the data in the pipe...", FromProcedureName
                Call ReceiveFromPipeSplit(strRes)
            End If
        Else
            'データなしの場合は、エラーコードチェックを行う
            If ErrorCode_PeekNamedPipe Then
                'PeekNamedPipe にてエラーがあった場合
                printMsg ERROR_, "Error PeekNamedPipe in readProcCDP. ErrorCode:" & ErrorCode_PeekNamedPipe & " - " & ErrorDetail.GetMessage(ErrorCode_PeekNamedPipe, "kernel32"), FromProcedureName, doRaiseError:=True

            ElseIf ErrorCode_ReadFile Then
                'ReadFile にてエラーがあった場合
                printMsg ERROR_, "Error ReadFile in readProcCDP. ErrorCode:" & ErrorCode_ReadFile & " - " & ErrorDetail.GetMessage(ErrorCode_ReadFile, "kernel32"), FromProcedureName, doRaiseError:=True
            Else
                'データがまだ、ない。つまり、破棄完了
                printMsg Trace_, "Message queue destruction completed.", FromProcedureName
            End If
        End If
    Loop Until InStrRev(strRes, vbNullChar) = Len(strRes) And intRes < 1

    printMsg Debug_, "Finished getting data from the pipe.", FromProcedureName
End Sub



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
    Const FromProcedureName As String = ThisClassName & ".Name"


    jsEval "document.title = '" & newName & "'", False
    printMsg info_, "Tab name changed to " & newName, FromProcedureName
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
    Const FromProcedureName As String = ThisClassName & ".id"


    sessionID = newID
    getSessionInfo        'update all session infos
    printMsg info_, brID & " updated Session ID to " & sessionID, FromProcedureName

End Property

Property Get html() As String
'------------------------------------------------------------
' Get the html text of the entire document
' Added 02/03/2023
'------------------------------------------------------------

    html = jsEval("document.documentElement.innerHTML")

End Property

'***************************************************************************************************
'* 機能　　：ブラウザからのイベント内容に関するプロパティを行います
'---------------------------------------------------------------------------------------------------
'* GET返値 ：呼び出した時点での、ブラウザイベント内容
'* LET引数 ：置き換えるDictionary変数。基本は、New Dictionary を渡してのリセットです
'---------------------------------------------------------------------------------------------------
'* 機能説明：・`Set BrowserEvents` に、`New Dictionary`を渡すと、イベント記録モードになります。
'            ・`Set BrowserEvents` に、`Nothing`を渡すと、イベント破棄モードになります。
'* 注意事項：Set プロパティ使用時、簡易的なDictionary構造のチェックを行います。構造が違うDictionaryを渡された場合、警告メッセージが出て、停止します。
'***************************************************************************************************
Property Get BrowserEvents() As Dictionary
    Set BrowserEvents = objAllMessages
End Property

Property Set BrowserEvents(ObjDic As Dictionary)
    Const FromProcedureName As String = ThisClassName & ".BrowserEvents"


    'SetされたDictionaryの構造をチェックします
    '1. そもそもDictionary オブジェクトが備わっているか？
    If ObjDic Is Nothing Then
        'Nothing 処理。以降は、イベントキャプチャはされません。
        Set objAllMessages = Nothing
        printMsg info_, "Event capture save data has been discarded. Event capture will not be performed.", FromProcedureName

    '2. まっさらの場合
    ElseIf ObjDic.Count = 0 Then
        '初期化処理。これにより、イベントキャプチャを開始します。
        Set objAllMessages = FormatDictionary_BrowserEvents
        printMsg info_, "A new event capture save data has been created. Starts event capture.", FromProcedureName

    '3. 指定のKeyがあるか？
    ElseIf ObjDic.Exists(EventsDictionaryKeyName01) And ObjDic.Exists(EventsDictionaryKeyName02) And ObjDic.Count = 2 Then
        '既存データをセット。これにより、このDictionaryにイベントキャプチャの追記が行われます。
        Set objAllMessages = ObjDic
        printMsg info_, "Existing event capture save data was successfully loaded. Resume event capture with this save data.", FromProcedureName

    '4. 不正なDictionary 形式なのでエラーで止める
    Else
        printMsg ERROR_, "Event capture save data structure is invalid.", FromProcedureName, doRaiseError:=True
    End If
End Property

'***************************************************************************************************
'* 機能　　：CDP-Json コマンド結果を得るためのタイムアウト時間を設定します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：設定中のタイムアウト秒数
'* LET引数 ：置き換えるタイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 機能説明：最初のブラウザ起動時の遷移完了判定や、CDP-Json コマンド結果を得るのに待つ時間を設定します。デフォルトは、10sです。
'***************************************************************************************************
Property Get TimeOutSecond() As Long
    TimeOutSecond = timerOut
End Property

Property Let TimeOutSecond(TimeSec As Long)
    timerOut = TimeSec
End Property

'***************************************************************************************************
'* 機能　　：非同期で実行した、CDP-Json コマンド結果を蓄積させる上限を決めます
'---------------------------------------------------------------------------------------------------
'* GET返値 ：設定中の上限数
'* LET引数 ：置き換える上限数
'---------------------------------------------------------------------------------------------------
'* 機能説明：`invokeMethodAsync`による結果をあとで、取り出す際の格納上限件数を来ます。デフォルトは、10件です。
'***************************************************************************************************
Property Get ResultBoxCDPForAsync() As Long
    ResultBoxCDPForAsync = MaxCommandResults
End Property

Property Let ResultBoxCDPForAsync(Number As Long)
    MaxCommandResults = Number
End Property

'***************************************************************************************************
'* 機能　　：同期モードでの CDP-Json コマンド結果のうち、最後のエラー情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後のエラー情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」のエラー情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます。
'            ・適切な`invokeMethod`を実行したつもりなのに返り値が、`Nothing`となるのが、エラーの目印となるので、それでこのエラー情報を見に行ってください。
'            ・`invokeMethod`の`StopError`が`Flase`だと実質機能しません。
'            ・非同期によるあとで取り出したパターンとは、連動しません。
'***************************************************************************************************
Property Get LastCDPJsonError() As Dictionary
    Set LastCDPJsonError = errorDetails
End Property

'***************************************************************************************************
'* 機能　　：非同期で実行されたIDを取り出します
'---------------------------------------------------------------------------------------------------
'* 引数　　：CommandID  非同期で実行した際の返り値(整理券番号)
'            ErrorExist エラー有無(参照渡しにより、実質返り値となります)
'
'* 返値    ：そのIDの実行結果のJson-Dictionary
'---------------------------------------------------------------------------------------------------
'* 注意事項：・事前に、`SendMessage` or `TakeEvents` が実行されてる必要があります
'            ・見つかったIDは、取り出した後、Dictionaryから自動で削除されます
'            ・ここで取り出した結果は、`LastCDPJsonError`には、反映しません
'***************************************************************************************************
Property Get ResultCDPForAsync(CommandID As Long, ByRef ErrorExist As Boolean) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".ResultCDPByAsync"


    '1. 存在するか？
    If AccumulatedCommandResults.Exists(CommandID) Then
        printMsg Debug_, "CommandID - " & CommandID & ",Raw-Json:" & AccumulatedCommandResults(CommandID), FromProcedureName

        '2. Jsonパーサーを施す
        Dim tmp As Dictionary: Set tmp = jsConverter.ParseJson(AccumulatedCommandResults(CommandID))

        '3. `result`オブジェクトがあるか？
        'ない場合は、そのまま返却
        If tmp.Exists("result") Then
            '正常に終了していると伝える
            ErrorExist = False
            printMsg Debug_, "This command ID(" & CommandID & ") has been completed successfully.", FromProcedureName

            Set ResultCDPForAsync = tmp("result")
        Else
            'エラーありとして伝える
            ErrorExist = True
            printMsg WARN_, "This command ID(" & CommandID & ") is in error. - Raw:" & AccumulatedCommandResults(CommandID), FromProcedureName

            Set ResultCDPForAsync = tmp
        End If

        '4. 取り出したので消しておく
        AccumulatedCommandResults.Remove CommandID
    Else
        ErrorExist = True
        printMsg WARN_, "The command ID(" & CommandID & ") passed is not running asynchronously or has been deleted.", FromProcedureName

        '2. カスタムエラーメッセージを作成
        Dim ErrorCode As Long: ErrorCode = -2
        Dim ErrorMessage As String: ErrorMessage = "The specified ID:" & CommandID & " does not exist."

        '3. カスタムエラーメッセージJsonで返却
        Set ResultCDPForAsync = jsConverter.ParseJson("{""error"":{""code"":" & ErrorCode & ",""message"":""" & ErrorMessage & """}}")
    End If
End Property



'***************************************************************************************************
'                              ■■■ プライベートプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP にメッセージを送信し、Chromium ブラウザでメッセージが受信されているかどうかを確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：Async が、`False` → CDP-Jsonコマンド の実行結果(Json-Dictionary 形式)
'                      `True`  → 実行したコマンドIDと、WriteFile結果コードのペアーDictionary
'
'* 引数　　：strMessage             Chromium ブラウザに送るCDP-Jsonコマンド(プレーンテキスト)
'            alwaysBrowserContext   `True` ：`sessionID`の有無問わず、常にブラウザ自身を対象したコマンド実行モードとなります
'                                   `False`：`sessionID`があれば、その`SessionID`に対するコマンド実行。なければ、ブラウザ自身を対象したコマンド実行となります。
'
'            InternalUseMessages    非同期イベントをクラス内部で後程、使用したいときに、`New Dictionry`として、渡して有効化します。
'            Async                  `True` ：コマンドの結果を待ちません。Do Loopフェーズには入りません。
'                                   `False`：コマンドの結果を待ちます。Do Loopフェーズに入り、そのIDが来るまで待ちます。この場合、非同期に実行した結果のID回収もついでに行います
'
'            dbgMsg                 Allow the option to bypass error if intended
'---------------------------------------------------------------------------------------------------
'* 更新履歴：・03/05/23 - Added soft-reset on peeknamepipe error
'            ・03/22/23 - Added error handling for better debugging
'            ・24/10/25 Daniel Polak - Added soft re-attach attempt when pipe error
'---------------------------------------------------------------------------------------------------
'* 注意事項：引数-Async の値によって、Dictionaryの構造が変わる特殊なメソッドです。別途公開メソッドで、役割分担しております
'***************************************************************************************************
Private Function SendMessage(strMessage As String, Optional alwaysBrowserContext As Boolean, Optional InternalUseMessages As Scripting.Dictionary, Optional Async As Boolean, Optional dbgMsg As Boolean = True) As Scripting.Dictionary
    Const FromProcedureName As String = ThisClassName & ".SendMessage"


    Dim intRes As Long
    Dim strRes As String
    Dim lngCurrentId As Long
    Dim results As Scripting.Dictionary       'to store a collection of IDs obtained from CDP
    Dim ErrorCode_PeekNamedPipe As Long, ErrorCode_ReadFile As Long

    lngCurrentId = lngLastID
    lngLastID = lngLastID + 1       'We increase the global ID counter

RetrySend:

    On Error GoTo ifError

   '--------------------------------------------------
   'ブラウザ自体の標準出力、標準出力エラー の確認
   '--------------------------------------------------
1   intRes = core.readProcSTD(strRes, False)                                                     'Sometimes edge writes to stdout so we clear stdout too.
    If intRes > 0 Then printMsg Debug_, "stdout = " & Replace(strRes, vbCrLf, ""), FromProcedureName

2   intRes = core.readProcSTD(strRes, True)                                                      'Sometimes edge writes to stderrout so we clear stderrout too.
    If intRes > 0 Then printMsg ERROR_, "stderrout = " & Replace(strRes, vbCrLf, ""), FromProcedureName

   '--------------------------------------------------
   'Send message to the browser
   '--------------------------------------------------
3   strMessage = Left(strMessage, Len(strMessage) - 1)                                           'We add the currentID and sessionID to the message
4   If sessionID <> "" And Not (alwaysBrowserContext) Then strMessage = strMessage & ", ""sessionId"":""" & sessionID & """"
5   strMessage = strMessage & ", ""id"":" & lngCurrentId & "}" & vbNullChar

    printMsg Debug_, "sendCDP-JsonCommand = " & strMessage, FromProcedureName
    
6   Dim WriteResult As Long: WriteResult = core.writeProc(strMessage)                                                                  'Send message to the browser instance
    printMsg Debug_, "Result-WriteFile:" & WriteResult, FromProcedureName

    '非同期(結果はいらない)モードなら、整理券(ID)を渡してここで抜けます
    If Async Then
        Set SendMessage = New Dictionary

        '実行したコマンドIDと、WriteFileの結果コードを渡して、抜けます
        SendMessage.Add lngCurrentId, WriteResult
        Exit Function
    End If

   '------------------------------------------------
   'Read message returned by browser
   '------------------------------------------------
    Dim timerStart As Double: timerStart = Timer                                                 'Failsafe timer to ensure loop ends when server fails to respond

    '受信処理→Json処理→受信処理→Json処理→…　というロジックでループ処理を行います
    'CDP-Jsonコマンドを送った際のIDが返ってくるまで、ループ受信処理を行いますが、タイムアウトになったら、抜けます
    Dim ResultSendCommand As Boolean                                                             'We read from edge and process messages until we receive a message with our ID
    printMsg Trace_, "Starting Processes incoming messages from the browser...", FromProcedureName
    Do
        '受信処理
7       intRes = core.readProcCDP(strRes, ErrorCode_PeekNamedPipe, ErrorCode_ReadFile)
        printMsg Trace_, "ResSize = " & intRes, FromProcedureName

        '返り値チェック
        If intRes > 0 Then
            'データあり。受信メッセージを活用するため、担当プロシージャへ遷移
8           Set SendMessage = ReceiveFromPipeSplit(strRes, InternalUseMessages, lngCurrentId)

            'オブジェクトありなら、直近に送ったコマンド結果であることを確認できたので、受信ループ終了フラグを立てる
9           If Not (SendMessage Is Nothing) Then ResultSendCommand = True
        Else
            'データなしの場合は念のため、エラーコードチェックを行う
            If ErrorCode_PeekNamedPipe Then
                'PeekNamedPipe にてエラーがあった場合
10              printMsg ERROR_, "Error PeekNamedPipe in readProcCDP. ErrorCode:" & ErrorCode_PeekNamedPipe & " - " & ErrorDetail.GetMessage(ErrorCode_PeekNamedPipe, "kernel32"), FromProcedureName, doRaiseError:=True
            ElseIf ErrorCode_ReadFile Then
                'ReadFile にてエラーがあった場合
11              printMsg ERROR_, "Error ReadFile in readProcCDP. ErrorCode:" & ErrorCode_ReadFile & " - " & ErrorDetail.GetMessage(ErrorCode_ReadFile, "kernel32"), FromProcedureName, doRaiseError:=True
            Else
                'データがまだ、ない
12              printMsg Trace_, "I used `readProcCDP` but it returned 0 Bytes.", FromProcedureName
            End If
        End If

        'いくつかのログ表示
        If ResultSendCommand = False Then
            '1. コマンド結果が、まだ来ない
            printMsg Trace_, "The command result from readProcCDP has not yet been received, so reception continues.", FromProcedureName

        ElseIf ResultSendCommand And StrPtr(strBuffer) <> 0 Then
            '2. コマンド結果は得れても、欠損Jsonがまだある場合は、不整合を防ぐため、ループ受信継続の旨を通知しておく
            printMsg Trace_, "The CDP-Json command result ID:" & lngCurrentId & " has been confirmed, but the receiving loop will continue until there is missing Json data or until it is completed.", FromProcedureName
        End If


        '一定時間経ってもコマンド結果が来ないなら、ループから抜けます
13      If Timer - timerStart > timerOut Then Err.Raise ERR_TIMEOUT, FromProcedureName, "Timeout waiting for browser to return message." Else printMsg Trace_, Timer - timerStart & " seconds have passed...", FromProcedureName
        DoEvents

        'CDP-Jsonコマンドの結果が得れておりなおかつ、欠けてるデータがなければ、受信ループを終了します
    Loop Until ResultSendCommand And StrPtr(strBuffer) = 0

    printMsg Trace_, "Command result ID:" & lngCurrentId & " was returned from CDP-Json, so the receiving loop will end.", FromProcedureName
    Exit Function

afterErr:

    printMsg WARN_, "Attempt to soft re-attach when pipe error because sometimes a simple retry is all it is needed to make it work :)", FromProcedureName
    On Error Resume Next
    attachToSession
    GoTo RetrySend

ifError:

    If Not dbgMsg Then Exit Function    'Allow the option to bypass error if intended

   'Record errmsg and print helpful infos
    Dim errorLine As Long, errorDesc As String
    errorLine = Erl
    errorDesc = Err.Description
    printMsg ERROR_, "sendMessage error at line " & errorLine & ": """ & errorDesc & """", FromProcedureName
    printMsg ERROR_, "strMessage used by sendMessage: " & strMessage, FromProcedureName

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
        printMsg ERROR_, "Possible reasons: (1) the browser is not actually running (ie. call string failure); (2) the browser window was accidentally closed; (3) the CDP Pipe connection to the browser was cut-off somehow.", FromProcedureName
        deleteCDPID
    End If

    Err.Raise IIf(InStr(errorDesc, "PeekNamedPipe") > 0, ERR_PIPE, ERR_PROTOCOL), FromProcedureName, "sendMessage error at line " & errorLine & ": """ & errorDesc & """"

End Function

'***************************************************************************************************
'* 機能　　：1度の`ReadFile`で取り出した長文を、1件1件に分割し、状況に応じて内部で、イベントキャプチャしたり、コマンド結果の確認を行います
'---------------------------------------------------------------------------------------------------
'* 返値　　：`checkCommandID`に対するCDP-Jsonコマンド結果Dictionary形式
'* 引数　　：ReceiveRawString       パイプから得たそのまんまの文字列
'            InternalUseMessages    非同期イベントをクラス内部で後程、使用したいときに、`New Dictionry`として、渡して有効化します。
'            checkCommandID         確認するコマンドID
'---------------------------------------------------------------------------------------------------
'* 機能説明：`checkCommandID`に合わないIDが来ても`AccumulatedCommandResults`に蓄積されます
'***************************************************************************************************
Private Function ReceiveFromPipeSplit(ByRef ReceiveRawString As String, Optional ByRef InternalUseMessages As Dictionary, Optional ByRef checkCommandID As Long = -1) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".ReceiveFromPipeSplit"


    '受信メッセージを活用するため、null文字区切りとしてスマートに分割
1   Dim JsonMessages: JsonMessages = Split(ReceiveRawString, vbNullChar)
    printMsg Trace_, UBound(JsonMessages) + 1 & " messages received.", FromProcedureName

    'メモリを解放
    If seeRawSendMsgDbg Then printMsg Trace_, "Raw - strRes = " & ReceiveRawString, FromProcedureName
2   ReceiveRawString = vbNullString     '参照渡しにより、元の変数もクリアできる

    '分割した数分、処理
    Dim i As Long
    For i = LBound(JsonMessages) To UBound(JsonMessages)
3       If seeRawSendMsgDbg Then printMsg Debug_, "Split(" & i & ") - strRes = " & JsonMessages(i), FromProcedureName  'Only enable if needs to see every raw JS message for debugging/development

        '最後の要素か？
4       If i = UBound(JsonMessages) Then
            '空欄か？
            If JsonMessages(i) = "" Then
                '1回の受信で全て正しいJson形式として得れました
                printMsg Debug_, "readProcCDP was able to receive everything!", FromProcedureName
            Else
                'データが、"欠けた"状態なので蓄積で、キャッシュしておく。頻繁には来ない想定
                printMsg Trace_, "The Split(" & i & ") data is missing JSON data. Accumulate until completion.", FromProcedureName
                printMsg Trace_, "Before-strBuffer = " & strBuffer, FromProcedureName
5               strBuffer = strBuffer & JsonMessages(i)
                printMsg Trace_, "After-strBuffer = " & strBuffer, FromProcedureName

                printMsg Debug_, "readProcCDP was unable to receive all the data. It accumulates the last element data...", FromProcedureName
            End If
        Else
            '1. 欠けたデータがあるか？
            Dim tmp As String
            If StrPtr(strBuffer) Then
                '空データじゃない場合は、結合する
                If JsonMessages(i) = "" Then
                    tmp = strBuffer
6                   printMsg Trace_, "Input data has arrived, but since it is an empty string, the currently stored data will be treated as JSON data.", FromProcedureName
                Else
7                   tmp = strBuffer & JsonMessages(i)
                    printMsg Trace_, "The merge completes the JSON data.", FromProcedureName
                End If

                '最後の要素ではないため、Jsonデータが完成されたとし、蓄積中のデータをクリアする
8               strBuffer = vbNullString
            Else
                'そのまんま使う
9               tmp = JsonMessages(i)
                printMsg Trace_, "There is currently no data being accumulated, so we will use it as is.", FromProcedureName
            End If

            '2. `vbnullchar`が連続で来た時に備えて、空文字はそもそも受け付けないようにする
            If seeRawSendMsgDbg Then printMsg Trace_, "Check the JSON structure → " & tmp, FromProcedureName
            If tmp <> "" Then
                '受信データのチェックを行うプロシージャへ遷移します。
10              If BrowserReceivedDataCheck(tmp, InternalUseMessages, checkCommandID) Then
                    '指定したコマンド結果であることを確認できたので、コマンド結果を返す
11                  Set ReceiveFromPipeSplit = jsConverter.ParseJson(tmp)    'JsonPurserして、コマンド結果を返却
                    printMsg Debug_, "The command result was obtained from readProcCDP! - ID:" & checkCommandID, FromProcedureName
                End If
            Else
                printMsg Debug_, "`vbnullchar` seems to be received with multiple characters.", FromProcedureName
            End If
        End If

        DoEvents
    Next
End Function

'***************************************************************************************************
'* 機能　　：イベントの「種類（メソッド名）」ごとの仕切りを入れ、そこにCollectionを入れていきます
'---------------------------------------------------------------------------------------------------
'* 実質返値：Accumulator
'* 引数　　：Accumulator    これまでの格納イベント。
'            eventData      Json-Dictionry済みのブラウザからのイベントデータ
'---------------------------------------------------------------------------------------------------
'* 機能説明：メソッド名をキーにし、それぞれの中では、時系列が保つように追加していきます。
'            これにより、「特定の"種類"のイベント」を、個別にチェックしやすくなります。
'            時系列データ(Index値)と合計キャプチャ数もおまけとして残しておきます。
'* 注意事項；このプロシージャを呼び出す前に、イベントデータであるかのチェックを済ませてください
'***************************************************************************************************
Private Sub OnNewEventReceived(Accumulator As Dictionary, eventData As Dictionary)
    Const FromProcedureName As String = ThisClassName & ".OnNewEventReceived"


    'イベント名を取得
    Dim EventMethodName As String: EventMethodName = eventData("method")

    ' 合計イベントキャプチャ数をカウントする
    Accumulator(EventsDictionaryKeyName01) = Accumulator(EventsDictionaryKeyName01) + 1
    printMsg Trace_, "Now Total of accumulated events :" & Accumulator(EventsDictionaryKeyName01), FromProcedureName

    ' インデックスも付与(1始まり)
    ' これにより、全体の時系列データも保持できます
    eventData.Add "__index__", Accumulator(EventsDictionaryKeyName01)


    '格納開始
    ' 1.イベントメソッドデータがあるDictionaryに遷移
    Dim EventsDictionary As Dictionary: Set EventsDictionary = Accumulator(EventsDictionaryKeyName02)

    ' 2. このイベント名の`Collection`が、すでにあるか？
    Dim eventList As Collection
    If EventsDictionary.Exists(EventMethodName) Then
        ' あれば、そのコレクションデータをセット
        Set eventList = EventsDictionary(EventMethodName)
        printMsg Debug_, "Stored event data from the browser. - " & EventMethodName, FromProcedureName
    Else
        ' なければ、新しい`Collection`をセット
        Set eventList = New Collection
        EventsDictionary.Add EventMethodName, eventList

        printMsg Debug_, "A new type of event. - " & EventMethodName, FromProcedureName
    End If

    ' 3. セットされたコレクションに、今回のイベントの詳細データを、追加する
    eventList.Add eventData
End Sub

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックします
'---------------------------------------------------------------------------------------------------
'* 返値    ：TRUE :受信ループを止める合図。
'               ・`CDPJsonString`に、ExcelからのCDP-Jsonコマンド結果がある
'               ・コマンド構文エラーによる、ID項目なし
'
'            FALSE:受信ループを継続する合図。
'               ・`CDPJsonString`は、ブラウザからのイベント情報である
'
'* 引数　　：returnCDPJson          1件分のJson形式文字列のブラウザからのイベントやCDP-Jsonコマンド結果
'            InternalUseMessages    クラス内プロシージャで、イベント内容を扱う用として蓄積させる参照変数
'            checkCommandID         確認するコマンドID
'---------------------------------------------------------------------------------------------------
'* 詳細説明：イベントデータ or CDP-Jsonコマンド結果 or 構文エラー の確認をし、ブラウザイベントの場合はついでに、イベント情報を格納します。
'* 注意事項；CDP-Jsonコマンド結果IDの不一致(シングルスレッド配下なら多分ないはず...)の場合は、不整合により不安定につながるため、エラーで停止します
'***************************************************************************************************
Private Function BrowserReceivedDataCheck(CDPJsonString As String, InternalUseMessages As Dictionary, checkCommandID As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".BrowserReceivedDataCheck"


    '======================================= チェック開始 =======================================
    Dim ResultCommandID As Long, AsynchronousEventName As String
    Dim JsonDictionaryObj As Dictionary
    Dim FlagValue As Long

    Const resultCommand As Long = 2 ^ 0 'IDのコマンド結果を検知出来た
    Const mismatchID    As Long = 2 ^ 1 'クラス内変数に記録してるIDとは違うコマンド結果が検知された。恐らく以前、非同期で実行したときの結果と思われる
    Const browserEvent  As Long = 2 ^ 2 '非同期イベントである
    Const EventCapture  As Long = 2 ^ 3 '非同期イベントキャプチャモードである
    Dim mask As Long: mask = resultCommand + mismatchID + browserEvent + EventCapture


    '--------- 1. ブラウザからのイベントか？ ---------
    If GetContentsMethod(CDPJsonString, AsynchronousEventName) Then
        FlagValue = FlagValue Or browserEvent   '非同期イベントとして、フラグ付け

        '--------- 2. いずれかのDictionary変数が、非同期イベントキャプチャモードか？ ---------
        If (objAllMessages Is Nothing) And (InternalUseMessages Is Nothing) Then
            '両方とも`Nothing`なら、わざわざ中身を解析する意味がないので、パーサーは使わない
            printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName
        Else
            'イベントキャプチャモードなので、JsonPurserを使用する
             printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
            Set JsonDictionaryObj = jsConverter.ParseJson(CDPJsonString)

            '--------- 3. `SessionID`項目があるか？ ---------
            If JsonDictionaryObj.Exists("sessionId") Then
                printMsg Trace_, "This event is a PageContext.", FromProcedureName

                '--------- 4. その`SessionID`は、クラス内変数に記録してる値と一致してるか？ ---------
                If JsonDictionaryObj("sessionId") = sessionID Then
                    'ページとしてのイベント情報です
                    FlagValue = FlagValue Or EventCapture   'イベントキャプチャモードとして、フラグ付け
                    printMsg Trace_, "An event for the current session ID was detected. - SessionId:" & sessionID, FromProcedureName
                Else
                    '旧セッションが、うまく破棄できなかった時に備えて、二重格納を抑制するため、`EventCapture`フラグは付与しません。
                    printMsg WARN_, "The event information does not match the session ID recorded in Excel. An old session remains, so please discard it with `Target.detachFromTarget`. - oldSessionId:" & JsonDictionaryObj("sessionId"), FromProcedureName
                End If
            Else
                'ブラウザ自身としてのイベント情報です
                FlagValue = FlagValue Or EventCapture       'イベントキャプチャモードとして、フラグ付け
                printMsg Trace_, "This event is a BrowserContext.", FromProcedureName
            End If
        End If

    Else
        FlagValue = FlagValue Or resultCommand  'IDのコマンド結果を検知出来たとして、フラグ付け

        '--------- 2. `id`項目があるか？ ---------
        If GetContentsID(CDPJsonString, ResultCommandID) Then
            '--------- 3. そのIDは、クラス内変数に記録してるIDと一致してるか？ ---------
            If ResultCommandID = checkCommandID Then
                '正しいコマンド結果が返ったので次項に、受信ループ終了フラグを返すようにする
                printMsg Trace_, "It was confirmed that the execution command ID matches the one recorded in Excel. ID:" & checkCommandID, FromProcedureName
            Else
                '多分、非同期で実行したコマンド結果と思われる
                'VBAによるシングルスレッドの制約上、上記以外のズレはありえないとおもう...
                FlagValue = FlagValue Or mismatchID '現在のコマンド結果IDと返ったコマンド結果IDが異なる場合のフラグ付け
                printMsg info_, "The command result (ID:" & ResultCommandID & ") was returned, but it does not match the command ID (" & checkCommandID & ") recorded in Excel. This is probably an asynchronous result. The results will be accumulated.", FromProcedureName

                'コマンド結果ID用受信Boxへ格納しておく
                '※パーサー処理負荷を抑えるため、Json文字列まんまとします
                AccumulatedCommandResults.Add ResultCommandID, CDPJsonString

                '格納件数チェックし、上限を超えたら、最古データを削除します
                If AccumulatedCommandResults.Count > MaxCommandResults Then
                    ' 一番古いキーを取得して削除
                    AccumulatedCommandResults.Remove AccumulatedCommandResults.keys(0)
                    printMsg info_, "The accumulation limit has exceeded " & MaxCommandResults & ", so the oldest has been deleted.", FromProcedureName
                End If
            End If
        Else
            '`method`も`id`項目もない場合、CDP-Jsonコマンドの構文エラーとして返った可能性があるとし、受信ループ終了とする
            printMsg WARN_, "Both the `id` and `method` fields are missing. This may be a syntax error.", FromProcedureName
        End If
    End If


    '======================================= このプロシージャから抜けるか判定 =======================================
    Select Case (FlagValue And mask)
        '送ったコマンドの結果は来たので、受信ループを止めて、抜ける
        Case resultCommand
            BrowserReceivedDataCheck = True
            Exit Function

        'コマンド結果だがID不一致。以前、非同期で実行したときの結果と思われるので、まだ受信ループは止めずに、抜ける
        Case resultCommand + mismatchID
            BrowserReceivedDataCheck = False
            Exit Function

        '受信ループは止めないが、イベントキャプチャしないモード(格納用のDictionaryがどちらもNothing or 古いSessionID)なので、次項の格納処理はせず、抜ける
        Case browserEvent
            BrowserReceivedDataCheck = False
            Exit Function

        'イベントキャプチャモードにより、Dictionaryオブジェクト格納フェーズへ行けるようにしておく。なお、受信ループは止めません
        Case browserEvent + EventCapture
            BrowserReceivedDataCheck = False

        '未知の組み合わせ
        Case Else
            printMsg ERROR_, "An unknown flag was passed. (" & (FlagValue And mask) & ")", FromProcedureName, doRaiseError:=True

    End Select


    '======================================= イベント情報を格納するフェーズへ =======================================
    'パブリック(呼び出し側使用)用
    If objAllMessages Is Nothing Then
        printMsg Trace_, "No event information is required on the module side.", FromProcedureName
    Else
        printMsg Trace_, "The module is requesting the accumulation of event information. Saving...", FromProcedureName
        OnNewEventReceived objAllMessages, JsonDictionaryObj
    End If

    'プライベート(クラス内変数)用
    If InternalUseMessages Is Nothing Then
        printMsg Trace_, "No event information needed inside the class.", FromProcedureName
    Else
        printMsg Trace_, "The class is requesting the accumulation of event information. Saving...", FromProcedureName
        OnNewEventReceived InternalUseMessages, JsonDictionaryObj
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`id`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：JsonString         CDPのレスポンス(Json文字列そのまんま)
'            idValueContents    コマンド結果のID値を格納する変数(参照渡しによって返します)
'***************************************************************************************************
Private Function GetContentsID(jsonString As String, ByRef idValueContents As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".GetContentsID"


    '検索設定
    Const SearchKeyName     As String = "id"    '検索文字列
    Const StartSearchPos    As Long = 3         'あるべき位置


    '1. 文字列を探索
    Dim StartPos As Long
    StartPos = InStr(StartSearchPos, jsonString, SearchKeyName)

    '2. 所定の位置にない場合は、存在しないとし、抜ける
    '下記の例に基づき、3文字目にない場合は、コマンド結果ではないとして、扱います。
    '・{"id":27,"result":{"result":...
    '・{"id":57,"error":{"code":...
    If StartPos = StartSearchPos Then
        GetContentsID = True

        '3. 開始位置をうまい具合にずらす
        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""":") '「id":」分、ずらすイメージ

        '4. ID値の終点を求める
        Dim EndPos As Long
        EndPos = InStr(StartPos, jsonString, ",")

        '5. ID値を取り出す
        idValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, "This response is the command result. ID:" & idValueContents, FromProcedureName
    Else
        GetContentsID = False
        printMsg Trace_, "This response does not have an `id`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`method`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：JsonString             CDPのレスポンス(Json文字列そのまんま)
'            methodValueContents    非同期イベントのメソッド名(参照渡しによって返します)
'***************************************************************************************************
Private Function GetContentsMethod(jsonString As String, ByRef methodValueContents As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".GetContentsMethod"


    '検索設定
    Const SearchKeyName     As String = "method"    '検索文字列
    Const StartSearchPos    As Long = 3             'あるべき位置


    '1. 文字列を探索
    Dim StartPos As Long
    StartPos = InStr(StartSearchPos, jsonString, SearchKeyName)

    '2. 所定の位置にない場合は、存在しないとし、抜ける
    '下記の例に基づき、3文字目にない場合は、非同期イベントではないとして、扱います。
    '・{"method":"Network.dataReceived","params":{"...
    If StartPos = StartSearchPos Then
        GetContentsMethod = True

        '3. 開始位置をうまい具合にずらす
        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""":""") '「method":"」分、ずらすイメージ

        '4. method名の終点を求める
        Dim EndPos As Long
        EndPos = InStr(StartPos, jsonString, """,")

        '5. method名を取り出す
        methodValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, "This response is an asynchronous event: " & methodValueContents, FromProcedureName
    Else
        GetContentsMethod = False
        printMsg Trace_, "This response does not have an `method`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザイベント情報を格納するためのフォーマットDictionaryを作成します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：フォーマット済みDictionary
'***************************************************************************************************
Private Function FormatDictionary_BrowserEvents() As Dictionary
    '新規オブジェクトセット
    Set FormatDictionary_BrowserEvents = New Dictionary

    '初期構造を作成します
    FormatDictionary_BrowserEvents.Add EventsDictionaryKeyName01, 0                 '合計イベントキャプチャ数
    FormatDictionary_BrowserEvents.Add EventsDictionaryKeyName02, New Dictionary    'イベントキャプチャDictionaryデータ
End Function

Private Function attachToSession() As Integer
'-----------------------------------------------------------------------
' This function must be called after start and before all other methods.
' This function attaches to a session of the browser.
' Updated: 03/05/23 - Added printMsg msgs & removed URL checking req.
'-----------------------------------------------------------------------
' Note: if needed to read the returned message result from sendMessage,
' type ?jsConverter.ConvertToJson(objAllMessages) in the ImmWnd
'-----------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".attachToSession"


    Dim results As Scripting.Dictionary
    Dim results_ As Scripting.Dictionary
    Dim results__ As Scripting.Dictionary
    Dim results___ As Scripting.Dictionary

    Dim strKey As Variant

    On Error GoTo ifError

    '`attachToSession`関数専用、イベントキャプチャDictionaryを作成します
    Dim EventMessageForAttachToSession As Dictionary: Set EventMessageForAttachToSession = FormatDictionary_BrowserEvents

    '実行
1   Set results = SendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":true}}", True, EventMessageForAttachToSession)
    printMsg info_, "Now looking for the browser window", FromProcedureName

   'Read message to confirm the requested url is found
   'Checking results__.Item("url") has been skipped for simplicity since v2.3
   'A new session always has the property attached = false and method = Target.targetCreated
   'This way of verification vastly improves the stability over checking by URL as sometimes
   'there are rogues targets with similar URL but they are not attachable (ie. no Session ID)
    Dim boolFound As Boolean
    For Each results In EventMessageForAttachToSession(EventsDictionaryKeyName02)("Target.targetCreated")
        Set results_ = results.Item("params")
        If results_.Exists("targetInfo") Then
            Set results__ = results_.Item("targetInfo")
            If results__.Item("attached") = False And results__.Item("type") = "page" Then
                boolFound = True
                Exit For
            End If
        End If
    Next
2   If Not boolFound Then printMsg ERROR_, "Unable to find the browser session", FromProcedureName, doRaiseError:=True _
    Else printMsg info_, "Now attaching CDP pipes to the browser session", FromProcedureName

   'Establish connection with the CDP pipes
3   targetID = results__.Item("targetId")
4   Set results___ = SendMessage("{""method"":""Target.attachToTarget"",""params"":{""targetId"":""" & results__.Item("targetId") & """,""flatten"":true}}")
5   sessionID = results___.Item("result").Item("sessionId")
6   Set results___ = SendMessage("{""method"":""Runtime.enable"",""params"":{}}")
7   Set results___ = SendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":false}}")

    printMsg info_, "Successfully attached to Session ID " & sessionID, FromProcedureName
    Exit Function

ifError:

   'Record errmsg and print helpful infos
    Dim errorLine As Long, errorDesc As String
    errorLine = Erl
    errorDesc = Err.Description
    printMsg ERROR_, "attachToSession error at line " & errorLine & ": """ & errorDesc & """", FromProcedureName
    printMsg Debug_, "EventMessageForAttachToSession = " & jsConverter.ConvertToJson(objAllMessages), FromProcedureName
    printMsg Debug_, "results = " & jsConverter.ConvertToJson(results), FromProcedureName
    printMsg Debug_, "results_ = " & jsConverter.ConvertToJson(results_), FromProcedureName
    printMsg Debug_, "results__ = " & jsConverter.ConvertToJson(results__), FromProcedureName
    printMsg Debug_, "results___ = " & jsConverter.ConvertToJson(results___), FromProcedureName
    printMsg Debug_, ".start call string: " & brCallString, FromProcedureName

    printParams
    Err.Raise ERR_PROTOCOL, , "attachToSession error at line " & errorLine & ": """ & errorDesc & """"

End Function

Private Function searchNull() As Long
'------------------------------------------------------------
' CDP messages received from chrome are null-terminated
' Updated: 25/10/25: Daniel Polak - new faster version
'------------------------------------------------------------

    Dim lngPos As Long

    If Len(strBuffer) = 0 Then Exit Function
    lngPos = InStr(1, strBuffer, vbNullChar, vbBinaryCompare)
    If lngPos > 0 Then
        searchNull = lngPos
    Else
        searchNull = 0
    End If

End Function

Private Function invokeError(result As Scripting.Dictionary) As Boolean
'-----------------------------------------------------------------------------------
' Check if InvokeMethod fails and what the error message returned is.
' Error message will be shown via printMsg
' Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'-----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".invokeError"


    Dim ErrorCode As Long, ErrorMessage As String, ErrorData As Variant
    If result Is Nothing Then
        ErrorCode = -1 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        ErrorData = vbEmpty
        printMsg WARN_, "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage, FromProcedureName
        invokeError = True
        Set errorDetails = jsConverter.ParseJson("{""code"":" & ErrorCode & ",""message"":""" & ErrorMessage & """}")     'カスタムエラーメッセージJsonで返却
    ElseIf result.Exists("error") Then
        Set errorDetails = result("error")
        ErrorCode = errorDetails("code")
        ErrorMessage = errorDetails("message")
        printMsg WARN_, "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage, FromProcedureName
        If errorDetails.Exists("data") Then
            ErrorData = errorDetails("data")
            printMsg WARN_, "       " & ErrorData, FromProcedureName
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

Private Sub cleanUpSessions(Optional onlyCdpSessions As Boolean = True)
'-----------------------------------------------------------------------------------------------------------------------------------------
' Provides cleaning before firing up a new session to prevent
' pipe error
' Update: 05/11/24: Daniel Polak - modified to solve problems with closing sessions (that do not belong to this user) on terminal server
'         23/10/25: Daniel Polak - p.name and p.terminate were not correct, replaced by proc.name and proc.terminate (discovered this bug because we are using Option Explicit now). On Error Resume Next was hiding the bug too.
'                                - only terminate automation-processes (where command line contains --remote-debugging-pipe)
'-----------------------------------------------------------------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".cleanUpSessions"


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
            '4. Killタイプ判定
            If Not onlyCdpSessions Then
                '5. デバッグブラウザ(CdpSessions)であるか問わず、通常ブラウザもKillします
                proc.Terminate

            ElseIf InStr(1, proc.CommandLine, "--" & ArgRemoteDebuggingName) > 0 And InStr(1, proc.CommandLine, brUser) > 0 Then
                '5. デバッグブラウザ(CdpSessions)であるかつ、対象ユーザー名なら、Killします
                 proc.Terminate
            End If
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

    printMsg info_, "cleanUpSessions completed and previous CdpSessionID removed", FromProcedureName

End Sub

Private Sub setCrashStateNormal()
'------------------------------------------------------------------
' Set Crashed state to Normal to prevent the "Restore Sessions?"
' popup from appearing every time the browser is started after
' being closed by browser.quit method.
' Requires user data dir profile stored in private var brUser and
' the browser type stored in private variable brType.
' Update: 22/01/26 - Tweaked to work well with the new multi-instance
'                    structure.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".setCrashStateNormal"


    Dim filePath_Preferences As String
    Dim strContents As String, strNewContents As String
    Dim intF As Integer

    '状況に応じたクラッシュパス情報を指定する
    '`UserDataDir`のポリシーがあったらそっちを優先
    If Policy_UserDataDirExists Then
        'Use this switch instead if your organization disables UserDataDir with policy
        filePath_Preferences = Policy_UserDataDirContents & "\" & brUser & "\Preferences"
    Else
        'user-data-dir
        Select Case brType
            Case brChrome: filePath_Preferences = Environ("LOCALAPPDATA") & "\Google\Chrome\" & brUser & "\Default\Preferences"
            Case brEdge: filePath_Preferences = Environ("LOCALAPPDATA") & "\Microsoft\Edge\" & brUser & "\Default\Preferences"
        End Select
    End If


    'クラッシュ情報がない場合は、ここで抜ける
    If LenB(Dir(filePath_Preferences)) = 0 Then Exit Sub

    intF = FreeFile
    Open filePath_Preferences For Binary As #intF
        strContents = Input$(LOF(intF), #intF)
    Close #intF

   'Precise replacement: only replace exit_type value
    strNewContents = Replace(strContents, """exit_type"":""Crashed""", """exit_type"":""Normal""", 1, -1, vbBinaryCompare)

    If strNewContents <> strContents Then
        intF = FreeFile
        Open filePath_Preferences For Binary As #intF
            Put #intF, 1, strNewContents
        Close #intF
    End If

    printMsg info_, "setCrashStateNormal completed", FromProcedureName

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
    Const FromProcedureName As String = ThisClassName & ".getBrowserPath"


    Select Case True
        Case InStr(LCase(strBrowserType), "chrome"): brType = brChrome
        Case InStr(LCase(strBrowserType), "edg"): brType = brEdge         'On edge it is just "Edg"
        Case Else: printMsg ERROR_, "Unable to getBrowserPath for browser type: " & brType, FromProcedureName, doRaiseError:=True
    End Select

   'Return the path
    On Error Resume Next
    Dim objWSHost As Object, reg As Object, strKeyPath As String
    Set reg = GetObject("winmgmts://./root/default:StdRegProv")

    'カスタムパスの指定があったらそっちを優先
    With ShSetting01_StartBrowser
        If .Range(.UseRangeName(1, FromProcedureName)).value = "" Then
            '1. レジストリで、ブラウザインストールパスを特定
            Set objWSHost = CreateObject("WScript.Shell")
            strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & brType & "\Path"
            getBrowserPath = objWSHost.regread(strKeyPath) & "\" & brType
        Else
            '1. カスタムパスを指定(ポータブル版ブラウザなど)
            getBrowserPath = .Range(.UseRangeName(1, FromProcedureName)).value
        End If
    End With

    'レジストリ操作にてエラー or バスが空の場合は、次の捜索へ
    If Err.Number <> 0 Or LenB(getBrowserPath) = 0 Then
        Err.Clear

        '2. デフォルトインストールパスチェック
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

        '2段階の捜索でもパスが見つからない場合は、ブラウザが入ってない or カスタムパスインストール or ポータブル版ブラウザ or Canaryブラウザの可能性
        On Error GoTo 0
        If LenB(getBrowserPath) = 0 Then printMsg ERROR_, "Unable to locate " & brType & " installation path.", FromProcedureName, doRaiseError:=True
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

Private Sub deleteCDPID()
'----------------------------------------------------------------------------------
' By      : Eschamali
' 機能    ：デバッグブラウザのパイプハンドル/セッション情報を消去します
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".deleteCDPID"


    With ShSetting01_StartBrowser
        .DeleteCDPToTable(1, FromProcedureName) = brUser  'パイプ通信Handle情報
        .DeleteCDPToTable(2, FromProcedureName) = brUser  'ブラウザセッション情報
    End With

End Sub

Private Sub SaveBase64ToFile(strB64 As String, FolderPath As String, FileName As String)
'----------------------------------------------------------------------------------
' Decode a base64 string and save to file without external dependencies (late bound).
'----------------------------------------------------------------------------------

    Dim dom As Object, el As Object, Bytes() As Byte, CharConv As CharacterCodeConversion

    Set dom = CreateObject("MSXML2.DOMDocument")
    Set el = dom.createElement("tmp")
    el.DataType = "bin.base64"
    el.Text = strB64
    Bytes = el.nodeTypedValue

    Set CharConv = New CharacterCodeConversion
    CharConv.BytesToSaveFile Bytes, FolderPath, FileName

End Sub

'***************************************************************************************************
'* 機能　　：URIのプロトコルチェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数の内容そのまんま
'* 引数　　：URI   ブラウザのURLに渡す予定の文字列
'---------------------------------------------------------------------------------------------------
'* 注意事項：・セキュリティ重視により、ホワイトリスト外のプロトコルが来たら、エラーで止まります
'            ・既に開いてるURIに関して(`getTab`メソッド)はチェック対象外とします
'***************************************************************************************************
Private Function checkURI(URI As String) As String
    Const FromProcedureName As String = ThisClassName & ".checkURI"


    '0. 空文字の場合は、チェックしません
    printMsg Debug_, "Checking the protocol of the URI...", FromProcedureName
    If URI = "" Then
        printMsg Debug_, "Checking is skipped because the URI is empty.", FromProcedureName

        checkURI = URI
        Exit Function
    End If

    '1. ホワイトリストのプロトコルを取得
    With ShSetting01_StartBrowser
        Dim tmp, WhiteList
        tmp = .ListObjects(.UseTableName(3, FromProcedureName)).DataBodyRange.Columns(Table03ColumnID.接頭辞)

        '配列か？
        If IsArray(tmp) Then
            'そのまんま
            WhiteList = tmp
        Else
            '配列じゃない場合は配列にしておく
            ReDim WhiteList(1 To 1, 1 To 1)
            WhiteList(1, 1) = tmp
        End If

        '解放
        tmp = vbNullString
    End With

    '2. 先頭始まりかチェック
    Dim i As Long
    For i = 1 To UBound(WhiteList)
        '先頭始まりなら、引数の内容そのまんまを返して抜ける
        If InStr(1, URI, WhiteList(i, 1)) = 1 Then
            printMsg Debug_, "Protocol check passed：" & WhiteList(i, 1), FromProcedureName

            checkURI = URI
            Exit Function
        End If
    Next

    '3. プロトコル名をチェック
    Dim ProtocolName, ErrorMes As String
    ProtocolName = Split(URI, ":")
    If UBound(ProtocolName) = 0 Then
        ErrorMes = "URLの形式が正しくありません。" & vbCrLf & "安全のため省略せず、`https://`や`file://`のような、プロトコル名を、URLの先頭に付けてください。" & vbCrLf & vbCrLf & "[問題のURL]" & vbCrLf & URI
    Else
        ErrorMes = "プロトコル`" & ProtocolName(0) & "`は、ホワイトリストに登録されていません。" & vbCrLf & "意図しないアプリケーションの起動などを防ぐため、処理を中断しました。" & vbCrLf & vbCrLf & "【解決策】" & vbCrLf & "このプロトコルへのアクセスが、本当に必要である場合は、シート「ブラウザ起動設定」のテーブル「ProtocolWhiteList」に、`" & ProtocolName(0) & "`を追加してから、もう一度、お試しください。"
    End If

    '4. エラーで中断
    printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True

End Function

'***************************************************************************************************
'* 機能　　：リモート デバッグに関するポリシーチェックを行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：BrowserExeName        "msedge.exe" とかを指定します。本家ライブラリに沿った形です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・企業によっては、ポリシー設定でリモートデバッグできないようにしています。
'　　　　　　・設定されていると、起動はできたのに、なんか操作できないといった混乱が発生するため、このチェックを設けています
'            ・現在は下記の確認に対応してます
'               ・RemoteDebuggingAllowed
'               ・UserDataDir
'***************************************************************************************************
Private Sub IsRemoteDebuggingBlockedByPolicy(BrowserExeName As String)
    Const FromProcedureName As String = ThisClassName & ".IsRemoteDebuggingBlockedByPolicy"


    'レジストリ操作に必要な定数/オブジェクトの定義
    Dim PolicyCheck As Object: Set PolicyCheck = GetObject("winmgmts:\\.\root\default:StdRegProv")
    Const HKEY_LOCAL_MACHINE    As Long = &H80000002           'このPCの、"すべて"のユーザーに影響する、強力な設定
    Const HKEY_CURRENT_USER     As Long = &H80000001           '"今、ログインしているユーザー"だけに、影響する設定
    Const wbemNoErr             As Long = 0                    '存在するときは、`0`を返します。　https://learn.microsoft.com/ja-jp/windows/win32/api/wbemdisp/ne-wbemdisp-wbemerrorenum#constants
    Const DisableRemote         As Long = 0                    '`0`が、リモート デバッグを使用することはできません。

    'ブラウザ名に応じて、ベースとなるキーパスを設定
    Dim RegPathBrowser As String
    Select Case BrowserExeName
        Case brEdge: RegPathBrowser = "\Microsoft\Edge"
        Case brChrome: RegPathBrowser = "\Google\Chrome"     'https://issues.chromium.org/issues/42323066
        Case Else: printMsg ERROR_, "Browser initiation fails: only the defined browser types can be accepted", FromProcedureName, doRaiseError:=True
    End Select

    'ポリシー情報のあるパスを設定します
    Dim BaseKeyPath As String: BaseKeyPath = "SOFTWARE\Policies" & RegPathBrowser


    '必要な変数を用意
    Dim checkValueName  As String   '値の名前
    Dim ValueExists     As Long     '存在有無結果
    Dim ValueContents   As Variant  '値の中身

    '"HKLM"と"HKCU"を確認する
    Dim i As Long, SetRegistryTree As Long
    Dim ErrorMes As String
    For i = 0 To 1
        'レジストリツリーを設定
        '最初は、PC単位、次は、現在のログインユーザー単位で確認します
        SetRegistryTree = IIf(i, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE)


        ' ------------ 1. RemoteDebuggingAllowed ------------
        'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/remotedebuggingallowed
        '※これに引っかかると残念ですが、CDP制御はできません
        checkValueName = "RemoteDebuggingAllowed"

        '1-1. 32bit値としてチェックする
        ValueExists = PolicyCheck.GetDWORDValue(SetRegistryTree, BaseKeyPath, checkValueName, ValueContents)
        printMsg Debug_, IIf(i, "HKCU-", "HKLM-") & "Policy：" & checkValueName & " existence status - " & ValueExists, FromProcedureName

        '1-2. そのデータ名があるか？
        If ValueExists = wbemNoErr Then
            printMsg Debug_, IIf(i, "HKCU-", "HKLM-") & "Policy：" & checkValueName & " contents - " & ValueContents, FromProcedureName

            '1-3. リモートできない設定値になってるか？
            If CLng(ValueContents) = DisableRemote Then
                '通達
                ErrorMes = IIf(i, "現在のログインPCアカウントでは、ブラウザのリモートデバッグが、ブロックされています。" & vbCrLf & vbCrLf & "CDP機能は、利用できません。" & vbCrLf & "これが個人のPCアカウントの場合は、レジストリ操作にて`1`にして下さい。", _
                                  "会社のセキュリティポリシーにより、ブラウザのリモートデバッグが、ブロックされています。" & vbCrLf & vbCrLf & "CDP機能は、利用できません。" & vbCrLf & "情報システム部門に、お問い合わせください。")

                printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True
                Exit For
            End If
        End If


        '------------ 2. UserDataDir ------------
        'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/userdatadir
        '※これに引っかかるとCDP制御はできますが、デバッグブラウザとは関係ない普通のブラウザも毎回、落とさないといけない制限となります
        checkValueName = "UserDataDir"

        '2-1. 文字列として、チェックする
        ValueExists = PolicyCheck.GetStringValue(SetRegistryTree, BaseKeyPath, checkValueName, ValueContents)
        printMsg Debug_, IIf(i, "HKCU-", "HKLM-") & "Policy：" & checkValueName & " existence status - " & ValueExists, FromProcedureName

        '2-2. そのデータ名があるか？
        If ValueExists = wbemNoErr Then
            printMsg Debug_, IIf(i, "HKCU-", "HKLM-") & "Policy：" & checkValueName & " contents - " & ValueContents, FromProcedureName

            'フラグ立て
            '※中身が空の場合は、`User Data`配下になります。
            Policy_UserDataDirExists = True
            Policy_UserDataDirContents = IIf(CStr(ValueContents) = "", Environ("LOCALAPPDATA") & RegPathBrowser & "\User Data", CStr(ValueContents))

            '通達
            ErrorMes = IIf(i, "現在のログインPCアカウントでは、ブラウザデータの配置場所が`" & Policy_UserDataDirContents & "`に固定されています。起動中のブラウザを全て落とします。これが個人のPCアカウントであり、この仕様を回避する場合は、レジストリ操作にて削除して下さい。", _
                              "会社のセキュリティポリシーにより、ブラウザデータの配置場所が`" & Policy_UserDataDirContents & "`に固定されています。起動中のブラウザを全て落とします。")

            printMsg WARN_, ErrorMes, FromProcedureName
            Exit For
        End If
    Next
End Sub



Private Sub Class_Initialize()
'----------------------------------------------------------------------------------
' A quick way to quickly check if an automated session is already active and open.
' If so, automatically attempt to reattach to it.
' Updated: 07/06/23
'          22/10/25 Daniel Polak - prevent compilation errors because depending on whether we are running in Excel, Word or Access only one of ThisWorkbook, ActiveDocument and CurrentProject is available
'          22/01/26 Long Vh - removed the auto deserialization feature to running multiple instances of browser profiles.
'----------------------------------------------------------------------------------

    'Assign an objID for easy debug identification
    brID = "BRID" & Format(Rnd * 1000, "000")


    '------------------------- クラス内変数の一部を初期化 -------------------------
    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        '1. イミディエイトに関する設定を行う
        seeRawSendMsgDbg = .Range(.UseRangeName(8, "CDPBrowser.Class_Initialize")).value           'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
    
        '2. パスが設定されてたら、ログファイル保存を有効化
        Dim LogFolderPath As String: LogFolderPath = .Range(.UseRangeName(9, "CDPBrowser.Class_Initialize")).value
        If LogFolderPath <> "" Then
            doLog = True                'ログファイルとして保存するか？
            logPath = LogFolderPath     '保存先のフォルダパスを設定
        End If

        '3. いくつかのポリシーチェックを行います。
        'この確認を使用することで、原因調査の時短につながる可能性があります。
        '※主要なブラウザのみチェックします
        IsRemoteDebuggingBlockedByPolicy IIf(.Range(.UseRangeName(4, "CDPBrowser.Class_Initialize")).value, brChrome, brEdge)
    End With

    'タイムアウト設定(デフォルトで、10s)
    timerOut = 10

    '非同期でのコマンド結果を蓄積させる件数
    Set AccumulatedCommandResults = New Dictionary
    MaxCommandResults = 10
    '------------------------------------------------------------------------------


    'Determine project host type for compatibility
    '※このフォークツールは、Excel特化に改良してるため、Excel以外は動作しません
    Select Case True
        Case InStr(Application.Name, "Excel"): Set thisApp = ThisWorkbook
        'Case InStr(Application.Name, "Word"): Set thisApp = ThisDocument
        'Case InStr(Application.Name, "Access"): Set thisApp = CurrentProject
    End Select

End Sub

'===============================================================================================
' Experimental Functions
'===============================================================================================
