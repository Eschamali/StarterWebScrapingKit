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



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)    'ミリ秒単位を扱うためのAPI(時差は、ローカルPCに準拠)
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)   'ミリ秒単位を扱うためのAPI(UTC準拠)

'構造体宣言
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'ログレベル設定用ユーザー定義
Public Enum LogLevelName
    Trace_
    Debug_
    info_
    WARN_
    ERROR_
End Enum



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
Private UTCMode As Boolean
Private doPrintDbgMsg As Boolean        'イミディエイトに、デバッグメッセージを出すか

Private SaveLogObj As New CharacterCodeConversion

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
' Private Methods
' ============================================= '

''
' @method Log
' @param {Long} Level
' @param {String} Message
' @param {String} [From = ""]
''
Private Sub Log(Level As Long, message As String, Optional From As String, Optional SaveLogFolderPath As String, Optional LogFileName As String)
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
        Dim log_LevelValue As String * 5
        Select Case Level
        Case 1
            log_LevelValue = "TRACE"
        Case 2
            log_LevelValue = "DEBUG"
        Case 3
            log_LevelValue = "INFO"
        Case 4
            log_LevelValue = "WARN"
        Case 5
            log_LevelValue = "ERROR"
        End Select

        '文字列生成
        Dim fromFixedLength As String * 50: fromFixedLength = From  '固定長文字列に変換して、ログの見た目を整える
        Dim LogString As String: LogString = "[" & GetNowFormat(UTCMode) & "] " & log_LevelValue & " | " & fromFixedLength & " " & message

        'イミディエイトへ表示
        If doPrintDbgMsg Then Debug.Print LogString

        'ファイルパスが指定されてたら、ログファイルへ保存
        If SaveLogFolderPath <> "" And LogFileName <> "" Then
            With SaveLogObj
                .BytesToSaveFile .BytesFromString(LogString & vbCrLf), SaveLogFolderPath, LogFileName, True
            End With
        End If
    End If
End Sub

'***************************************************************************************************
'* 機能　　：現在日時をミリ秒単位で取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`yyyy/mm/dd hh:mm:ss.fff`という書式
'* 引数　　：UseUTC
'               `True`　UTC基準で取得
'               `False` 時差込みで取得(ローカルPCの設定に依存)
'---------------------------------------------------------------------------------------------------
'* 注意事項：関数自体やPC環境ごとにミリ秒の精度に誤差があるため、実際の時刻とはどうしても数ミリ秒から数10ミリ秒程度でズレてしまうことは認識しておく必要があります。
'***************************************************************************************************
Private Function GetNowFormat(Optional UseUTC As Boolean) As String
    Dim t As SYSTEMTIME

    ' 現在日時取得
    If UseUTC Then GetSystemTime t Else GetLocalTime t

    ' `yyyy/mm/dd hh:mm:ss.fff`に整形
    With t
        GetNowFormat = Format(.wYear, "0000") & "/" & _
                       Format(.wMonth, "00") & "/" & _
                       Format(.wDay, "00") & " " & _
                       Format(.wHour, "00") & ":" & _
                       Format(.wMinute, "00") & ":" & _
                       Format(.wSecond, "00") & "." & _
                       Format(.wMilliseconds, "000")
    End With
End Function



' ============================================= '
' Public Methods
' ============================================= '

''
' @method LogTrace
' @param {String} Message
' @param {String} From
''
Public Sub LogTrace(message As String, Optional From As String, Optional FolderPath As String, Optional FileName As String)
    Log 1, message, From, FolderPath, FileName
End Sub

''
' @method LogDebug
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogDebug(message As String, Optional From As String, Optional FolderPath As String, Optional FileName As String)
    Log 2, message, From, FolderPath, FileName
End Sub

''
' @method LogInfo
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogInfo(message As String, Optional From As String, Optional FolderPath As String, Optional FileName As String)
    Log 3, message, From, FolderPath, FileName
End Sub

''
' @method LogWarning
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogWarn(message As String, Optional From As String, Optional FolderPath As String, Optional FileName As String)
    Log 4, message, From, FolderPath, FileName
End Sub

''
' @method LogError
' @param {String} Message
' @param {String} [From = ""]
' @param {Long} [ErrNumber = 0]
''
Public Sub LogError(message As String, Optional From As String, Optional FolderPath As String, Optional FileName As String, Optional ErrNumber As Long = 0)
    Dim log_ErrorValue As String
    If ErrNumber <> 0 Then
        log_ErrorValue = ErrNumber
    
        ' For object errors, extract from vbObjectError and get Hex value
        If ErrNumber < 0 Then
            log_ErrorValue = log_ErrorValue & " (" & (ErrNumber - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(ErrNumber)) & ")"
        End If
        
        log_ErrorValue = " | ErrorCode: " & log_ErrorValue
    End If

    Log 5, message & log_ErrorValue, From, FolderPath, FileName
End Sub



Private Sub Class_Initialize()
    '------------------------- クラス内変数の一部を初期化 -------------------------
    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        doPrintDbgMsg = .Range(.UseRangeName(7, "Logger.Class_Initialize")).value   'イミディエイトに表示させるか？
        LogThreshold = .Range(.UseRangeName(5, "Logger.Class_Initialize")).value    'ログレベルの設定
    End With
    '------------------------------------------------------------------------------
End Sub
