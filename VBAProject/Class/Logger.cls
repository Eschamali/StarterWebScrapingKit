VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' VBA-Log v0.1.0
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Log
'
' Logging for VBA
'
' @module Logger
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' Levels
' 0: Off
' 1: Trace/All
' 2: Debug
' 3: Info
' 4: Warn
' 5: Error

''
' Callback to call on log
' Signature: Level As Long, Message As String, From As String
'
' @example
' ```VB.net
' Public Sub ImmediateLog(Level As Long, Message As String, From As String)
'   Dim LevelValue As String
'   Select Case Level
'   Case 1
'     LevelValue = "Trace"
'   Case 2
'     LevelValue = "Debug"
'   Case 3
'     LevelValue = "Info"
'   Case 4
'     LevelValue = "WARN"
'   Case 5
'     LevelValue = "ERROR"
'   End Select
'
'   Debug.Print LevelValue & " - " & IIf(From <> "", From & ": ", "") & Message
' End Sub
'
' Logger.LogCallback = "ImmediateLog"
'
' Logger.LogWarn "watch out..."
' ' -> WARNING - watch out...
'
' Logger.LogError "uh oh...", "Module.Sub"
' ' -> ERROR - Module.Sub: uh oh...
'
' ' Can also attach multiple logging callbacks
'
' Public Sub FileLog(Level As Long, Message As String, From As String)
'   ' ...
' End Sub
'
' Logger.LogCallback = Array("ImmediateLog", "FileLog")
' ```
' @property LogCallback
' @type String|Array
''
Option Explicit



Public LogCallback As Variant

''
' Turn logging off (0) or only log messages that are >= threshold
'
' LogThreshold = 4 -> Level >= 4 -> Warn and Error
' LogThreshold = 0 -> Off
'
' @property LogThreshold
' @type Long
' @default 0
''
Private LogThreshold As Long

''
' @property LogEnabled
' @type Boolean
' @default False
''
Public Property Get LogLevel() As Long
    LogLevel = LogThreshold
End Property

Public Property Let LogLevel(value As Long)
    LogThreshold = value
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' @method Log
' @param {Long} Level
' @param {String} Message
' @param {String} [From = ""]
''
Private Sub Log(Level As Long, message As String, Optional From As String)
    If LogThreshold = 0 Or Level < LogThreshold Then
        Exit Sub
    End If

    If Not VBA.IsEmpty(LogCallback) Then
        Select Case VBA.VarType(LogCallback)
        Case VBA.vbString
            Application.Run CStr(LogCallback), Level, message, From
        Case VBA.vbArray To VBA.vbArray + VBA.vbByte
            Dim log_i As Long
            For log_i = LBound(LogCallback) To UBound(LogCallback)
                Application.Run CStr(LogCallback(log_i)), Level, message, From
            Next log_i
        End Select
    Else
        '
        ' For applications that don't have Application.Run (e.g. Access)
        ' comment out the above Application.Run code to fix compilation issues
        ' and manually insert the logging functionality here
        '
        Dim log_LevelValue As String
        Select Case Level
        Case 1
            log_LevelValue = "Trace"
        Case 2
            log_LevelValue = "Debug"
        Case 3
            log_LevelValue = "Info "
        Case 4
            log_LevelValue = "WARN "
        Case 5
            log_LevelValue = "ERROR"
        End Select
    
        Debug.Print "[" & Format(Now(), "yyyy/mm/dd hh:mm:ss") & "] " & log_LevelValue & " " & IIf(From <> "", "{" & From & "} ", "") & message
    End If
End Sub

''
' @method LogTrace
' @param {String} Message
' @param {String} From
''
Public Sub LogTrace(message As String, Optional From As String)
    Log 1, message, From
End Sub

''
' @method LogDebug
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogDebug(message As String, Optional From As String)
    Log 2, message, From
End Sub

''
' @method LogInfo
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogInfo(message As String, Optional From As String)
    Log 3, message, From
End Sub

''
' @method LogWarning
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogWarn(message As String, Optional From As String)
    Log 4, message, From
End Sub

''
' @method LogError
' @param {String} Message
' @param {String} [From = ""]
' @param {Long} [ErrNumber = 0]
''
Public Sub LogError(message As String, Optional From As String, Optional ErrNumber As Long = 0)
    Dim log_ErrorValue As String
    If ErrNumber <> 0 Then
        log_ErrorValue = ErrNumber
    
        ' For object errors, extract from vbObjectError and get Hex value
        If ErrNumber < 0 Then
            log_ErrorValue = log_ErrorValue & " (" & (ErrNumber - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(ErrNumber)) & ")"
        End If
        
        log_ErrorValue = log_ErrorValue & ", "
    End If

    Log 5, log_ErrorValue & message, From
End Sub



'初期化時に、条件付きコンパイル引数 を基に、初期ログレベルを決める
Private Sub Class_Initialize()
    #If LOG_LEVEL = 4 Then
        LogThreshold = 4
    
    #ElseIf LOG_LEVEL = 3 Then
        LogThreshold = 3
    
    #ElseIf LOG_LEVEL = 2 Then
        LogThreshold = 2
        
    #ElseIf LOG_LEVEL = 1 Then
        LogThreshold = 1
    
    #Else
        LogThreshold = 0

    #End If
End Sub
