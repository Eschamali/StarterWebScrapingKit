Attribute VB_Name = "CDPHelpers"
'==============================================================================================================
'                       汎用的に使うCDP/BiDi関連の便利プロシージャ/固定設定値です
'                        特定のClassでしか使わない物はここに配置してはいけません
'==============================================================================================================
Option Explicit
Option Private Module



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
'----- 待機関連 -----
Private Declare PtrSafe Sub sleep2 Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long    'タイマー用
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long         '周波数取得用



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'JSの`document.readyState`の状態一式
Public Enum ReadyState      'Used for .wait method
    isLoading = 0           'equivalence of the browser's "loading" state
    isInteractive = 1       'equivalence of the browser's "interactive" state
    isComplete = 2          'equivalence of the browser's "complete" state
End Enum

'バッファー設定関連
Public Const InitialBuffer             As Long = 2 ^ 20        'CDPやり取りPipe/テキスト変数/ADODB.Stream 初期バッファー上限
Public Const RunDoEventsCount          As Long = 2 ^ 10        '長いループ中に`DoEvents`を挟む間隔値

'ブラウザからの非同期イベント情報を格納する際のDictionaryKey名設定
Public Const EventsDictionaryKeyName01 As String = "TotalEvents"
Public Const EventsDictionaryKeyName02 As String = "EventMethods"

'その他
Public Const LimitCommandID    As Long = 2000000000             'CDP/BiDiコマンド送信時のID上限値
Public Const chromeWindowClass As String = "Chrome_WidgetWin_1" 'same window class for Edge
Private Const ThisModuleName   As String = "CDPHelpers"         'トレース用



'***************************************************************************************************
'                                   ■■■ 各種変数 ■■■
'***************************************************************************************************
Private m_Frequency As Currency '実行マシンでの周波数記録用
Private LogControl As New Logger   'ログレベルの制御



'***************************************************************************************************
'                       ■■■ Enum → 文字列 変換プロシージャ ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] ブラウザ種別をexe名で返します
'---------------------------------------------------------------------------------------------------
Public Function EnumToStringBrowserList_exeName(param As BrowserList) As String
    Select Case param
        Case BrowserList.RunChrome: EnumToStringBrowserList_exeName = "chrome.exe"
        Case BrowserList.RunEdge:   EnumToStringBrowserList_exeName = "msedge.exe"
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ブラウザ種別を相対パス名で返します
'---------------------------------------------------------------------------------------------------
Public Function EnumToStringBrowserList_RegPath(param As BrowserList) As String
    Select Case param
        Case BrowserList.RunChrome: EnumToStringBrowserList_RegPath = "\Google\Chrome"
        Case BrowserList.RunEdge:   EnumToStringBrowserList_RegPath = "\Microsoft\Edge"
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 待機の種類を文字列で返します
'---------------------------------------------------------------------------------------------------
Public Function EnumToStringReadyState(param As ReadyState) As String
    Select Case param
        Case ReadyState.isLoading:      EnumToStringReadyState = "loading"
        Case ReadyState.isInteractive:  EnumToStringReadyState = "interactive"
        Case ReadyState.isComplete:     EnumToStringReadyState = "complete"
    End Select
End Function



'***************************************************************************************************
'                                     ■■■ 待機系 ■■■
'***************************************************************************************************
'* 機能　　：シンプルな待機です
'---------------------------------------------------------------------------------------------------
'* 引数　　：seconds    何秒間止めるか？
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Custom sleep function. Sleep by 0.5s by default.
'* 注意事項：Useful for a quick necessary pause when needed.
'***************************************************************************************************
Public Sub Sleep(Optional seconds As Double = 0.5)
    Const baseUnit As Long = 1000    'ie. millisecs


    sleep2 seconds * baseUnit
    DoEvents
End Sub

'***************************************************************************************************
'* 機能　　：PC起動時からの「経過時間（ミリ秒）」を正確に返します（Timer関数の神互換）
'---------------------------------------------------------------------------------------------------
'* 返り値　：PC起動時からの「経過時間（ミリ秒）」
'---------------------------------------------------------------------------------------------------
'* 詳細説明：これを使う前に、`InitQueryPerformanceFrequency`を呼び出すこと
'***************************************************************************************************
Public Function TimerCounter() As Double
    '1. 今の「振動回数」を取得する
    Dim currentCount As Currency
    Call QueryPerformanceCounter(currentCount)

    '2. 割り算して「秒」にし、1000倍して「ミリ秒（Double型）」にして返す！
    TimerCounter = (CDbl(currentCount) / CDbl(m_Frequency)) * 1000#
End Function

'***************************************************************************************************
'* 機能　　：実行マシンの「1秒間の振動数（周波数）」を取得して記憶します
'***************************************************************************************************
Public Sub InitQueryPerformanceFrequency()
    Call QueryPerformanceFrequency(m_Frequency)
End Sub



'***************************************************************************************************
'                                     ■■■ ログ系 ■■■
'***************************************************************************************************
'* 機能　　：Immediate Window と、任意のフォルダへのログファイル出力を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：LogLevel_          ログレベル
'            strMsg            本文
'            From              呼び出し元（トレース用）
'            LogFileExtension  ログファイルの拡張子（各 Class の定数を渡す）
'            LogID             オブジェクト識別ID。空なら本文のみ
'            isHeader          True で区切り線付きヘッダー表示
'            doRaiseError      True でログ出力後にエラーを発生させる
'            SaveLogFolderPath ログファイルの保存フォルダ。空ならファイル出力しない
'            RaiseErrorNumber  doRaiseError 時のエラー番号。0 なら `CDPCustomErrorCodes.Protocol`
'---------------------------------------------------------------------------------------------------
'* 注意事項：保存先パスと拡張子は各 Class 側で保持し、この引数へ渡してください
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, LogFileExtension As String, _
    Optional LogID As String, _
    Optional isHeader As Boolean = False, _
    Optional doRaiseError As Boolean = False, _
    Optional SaveLogFolderPath As String, _
    Optional RaiseErrorNumber As Long)

    Dim logName As String, strFormattedMsg As String

    If isHeader Then
        strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-")
    Else
        strFormattedMsg = " | " & LogID & " | " & strMsg
    End If

    If LenB(SaveLogFolderPath) Then
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & LogFileExtension
        If LenB(Dir(SaveLogFolderPath, vbDirectory)) = 0 Then MkDir SaveLogFolderPath
    End If

    With LogControl
        Select Case LogLevel_
            Case LogLevelName.Trace_: .LogTrace strFormattedMsg, From, SaveLogFolderPath, logName
            Case LogLevelName.Debug_: .LogDebug strFormattedMsg, From, SaveLogFolderPath, logName
            Case LogLevelName.info_: .LogInfo strFormattedMsg, From, SaveLogFolderPath, logName
            Case LogLevelName.WARN_: .LogWarn strFormattedMsg, From, SaveLogFolderPath, logName
            Case LogLevelName.ERROR_: .LogError strFormattedMsg, From, SaveLogFolderPath, logName, Err.Number
            Case Else: 'ログ出力無効化
        End Select
    End With

    If doRaiseError Then
        If RaiseErrorNumber = 0 Then RaiseErrorNumber = CDPCustomErrorCodes.Protocol
        Err.Raise RaiseErrorNumber, From, Description:=strMsg
    End If
End Sub
