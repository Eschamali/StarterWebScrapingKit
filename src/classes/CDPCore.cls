VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
' Core APIs to generate Pipes and read CDP message
' This class is used by Browser class (CDPBrowser)
'--------------------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       24/10/25 Daniel Polak: added Option Explicit
'       29/11/22 Long Vh: removed the serializability check + class_terminate to keep handles constantly
'                         open for subsequent deserializations if needed.
'       03/06/22 Long Vh: added extensive commentations + refactored codes for HSBC usages
' Notes       :
'       This class should only be initiated internally by the CDPBrowser class and should not need to be
'       initiated directly at any times.
' References  :
'       Nil
' Sources     :
'       Creating Pipe: https://stackoverflow.com/questions/62172551/error-with-createpipe-in-vba-office-64bit
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
'----- 全ての始まり -----
Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessW" ( _
    ByVal lpApplicationName As LongPtr, _
    ByVal lpCommandLine As LongPtr, _
    lpProcessAttributes As Any, _
    lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Any, _
    ByVal lpCurrentDirectory As LongPtr, _
    lpStartupInfo As Any, _
    lpProcessInformation As Any) As Long

'----- パイプ通信関連 -----
Private Declare PtrSafe Function CreatePipe Lib "kernel32" ( _
    phReadPipe As LongPtr, _
    phWritePipe As LongPtr, _
    lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long

Private Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As Long

Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As LongPtr, _
    lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long) As Long

Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    lpOverlapped As Long) As Long



'***************************************************************************************************
'                                   ■■■ 各種構造体 ■■■
'***************************************************************************************************
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

' This is the structure to pass more than 3 fds to a child process (https://github.com/libuv/libuv/blob/v1.x/src/win/process-stdio.c)
' The fields crt_flags and os_handle must lie contigously in memory i.e. should not be aligned to byte boundaries
' You cannot define a packed struct in VBA, that's why we need to have multiple structs

Private Type STDIO_BUFFER
    number_of_fds As Long
    crt_flags(0 To 4) As Byte
    os_handle(0 To 4) As LongPtr
End Type

#If Win64 Then

    Private Type STDIO_BUFFER2
        number_of_fds As Long
        raw_bytes(0 To 44) As Byte
    End Type

#Else

    Private Type STDIO_BUFFER2
        number_of_fds As Long
        raw_bytes(0 To 24) As Byte
    End Type

#End If



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'ウィンドウ表示設定値一式 https://learn.microsoft.com/ja-jp/windows/win32/api/winuser/nf-winuser-showwindow
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

'このツール独自のエラーコード
Public Enum CDPCustomErrorCodes
    TIMEOUT = 900   '時間内に、CDPから応答がありませんでした
    PIPE = 901      'Pipe通信周り(`PeekNamedPipe`など)でエラー
    Protocol = 902  'CDPコマンド自体のエラー
End Enum

'プロセス起動関連設定
Private Const STARTF_USESTDHANDLES  As Long = &H100&
Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const STARTF_USESHOWWINDOW  As Long = &H1&

'バッファー設定関連
Private Const InitialBuffer             As Long = 2 ^ 20        'CDPやり取りPipe/テキスト変数/ADODB.Stream 初期バッファー上限　※`CDPCoreViaWebSocket.cls - InitialBuffer`と揃えること
Private Const dwAllocationGranularity   As Long = 2 ^ 16        'メモリアロケーション粒度。本来ならば、「GetSystemInfo - dwAllocationGranularity」で正確に取るべきだが、これのために何十行の宣言等を書くのが面倒なので、定数で妥協。過去30年変わってないので。バッファー拡張判定用に使います。なお、この補正はあくまでも、Pipe経路に限ります。
Private Const RunDoEventsCount          As Long = 2 ^ 10        '長いループ中に`DoEvents`を挟む回数
Private Const MAX_STR_LEN               As Long = 2 ^ 30 - 1    'VBA文字列の限界

'その他
Private Const ThisClassName             As String = "CDPCore"
Private Const LimitCommandID            As Long = 2000000000    'CDPコマンド送信時のID上限値



'***************************************************************************************************
'                                   ■■■ イベント定義 ■■■
'***************************************************************************************************
'CDPレスポンスの種類に応じた発信機
Public Event CDPBrowserEvent(methodName As String, RawJson As String)                   'ブラウザに対する非同期イベント
Public Event CDPBrowserID(id As Long, RawJson As String)                                'ブラウザに対するCDPコマンド結果
Public Event CDPContextEvent(methodName As String, RawJson As String, sessionID As String)  'タブに対する非同期イベント
Public Event CDPContextID(id As Long, RawJson As String, sessionID As String)               'タブに対するCDPコマンド結果

'コマンドIDリセットサイン
Public Event ResetCommandID()



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'各種クラスオブジェクト
Private Utf8Converter           As CharacterCodeConversion  'UTF-8を使用した日本語通信を有効化するやつ
Private LogControl              As New Logger               'ログレベルの制御
Private WithEvents viaWebSocket As CDPCoreViaWebSocket      'WebSocket経由でCDP制御したい場合のオブジェクト用
Attribute viaWebSocket.VB_VarHelpID = -1
Private WithEvents viaWebView2  As CDPCoreViaWebView2       'WebView2経由でCDP制御したい場合のオブジェクト用
Attribute viaWebView2.VB_VarHelpID = -1

'バッファー管理
Private Type ResponseBufferCDP
    strBuffer       As String   'ブラウザから送られたCDP情報をどんどん蓄積する変数
    length          As Long     '現在のバッファ全体の長さ
    StartCursor     As Long     '取り出しの開始位置
    AddCursor       As Long     '次に蓄積し始める位置
    EndCursor       As Long     '有効予定なデータ末尾位置(※バッファー拡張判定用)
    ExistsNewRes    As Boolean  '新規メッセージ有無
End Type
Private responseCDP     As ResponseBufferCDP
Private CDPPipeStream   As Object   'Pipeバイトストリームオブジェクト

'ログ設定状況
Private Type StateLog
    seeRawSendMsgDbg    As Boolean  'set to True if needing to see the raw json msg received from the browser for debugging
    doLog               As Boolean  'ログファイルにも、デバッグメッセージを残すか
    logPath             As String   'ログファイルの保存フォルダパス
End Type
Private currentLog  As StateLog
Private LogID       As String

'パイプ関連のエラー番号
Private Type ErrorCode
    PeekNamedPipe   As Long
    ReadFile        As Long
    WriteFile       As Long
End Type
Private LastPipeError   As ErrorCode

'各種ハンドル情報
Private hStdOutWr       As LongPtr
Private hStdOutRd       As LongPtr
Private hStderrOutWr    As LongPtr
Private hStderrOutRd    As LongPtr
Private hStdInWr        As LongPtr
Private hStdInRd        As LongPtr
Private hCDPOutWr       As LongPtr
Private hCDPOutRd       As LongPtr
Private hCDPInWr        As LongPtr
Private hCDPInRd        As LongPtr
Private hProcess        As LongPtr

'起動ウィンドウ情報
Private setShowWindow   As Integer
Private dwProcessId     As LongPtr

'コマンドIDとSessionIDの紐づけ管理
Private DictionarySessionID As Dictionary

'その他
Private m_Frequency         As Currency '実行マシンでの周波数記録用
Private m_NullTerminator(0) As Byte     'バイト配列で送るヌル文字用



'***************************************************************************************************
'                             ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：Chromiumを`remote-debugging-pipe`として起動し、パイプ接続を開始します
'---------------------------------------------------------------------------------------------------
'* 返り値　：実行結果　※0で成功です
'* 引数　　：strExec            コマンドライン
'            ShowWindowMode     初期ウィンドウ表示タイプ
'***************************************************************************************************
Public Function InitCDP(strExec As String, ShowWindowMode As WinState) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim result As Long
    Dim pipes As STDIO_BUFFER
    Dim pipes2 As STDIO_BUFFER2

   'First we create all 5 pipes
   'We start with stdout of the edge process
    sa.nLength = LenB(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    result = CreatePipe(hStdOutRd, hStdOutWr, sa, 0)
    If result = 0 Then
        InitCDP = Err.LastDllError     'エラーコードで返す
        Exit Function
    End If

    'stderr
    result = CreatePipe(hStderrOutRd, hStderrOutWr, sa, 0)
    If result = 0 Then
        InitCDP = Err.LastDllError     'エラーコードで返す
        Exit Function
    End If

   'Then stdin
    result = CreatePipe(hStdInRd, hStdInWr, sa, 0)
    If result = 0 Then
        InitCDP = Err.LastDllError     'エラーコードで返す
        Exit Function
    End If

   'Then the out pipe for the CDP Protocol
    result = CreatePipe(hCDPOutRd, hCDPOutWr, sa, InitialBuffer)    '※下限は、2 ^ 16
    If result = 0 Then
        InitCDP = Err.LastDllError     'エラーコードで返す
        Exit Function
    End If

   'And finally the in pipe
    result = CreatePipe(hCDPInRd, hCDPInWr, sa, InitialBuffer)      '※下限は、2 ^ 16
    If result = 0 Then
        InitCDP = Err.LastDllError     'エラーコードで返す
        Exit Function
    End If

   'Then we fill the special structure for passing arbitrary pipes (i.e. fds) to a process
    pipes.number_of_fds = 5
    pipes.os_handle(0) = hStdInRd
    pipes.os_handle(1) = hStdOutWr
    pipes.os_handle(2) = hStderrOutWr
    pipes.os_handle(3) = hCDPInRd
    pipes.os_handle(4) = hCDPOutWr
    pipes.crt_flags(0) = 9
    pipes.crt_flags(1) = 9
    pipes.crt_flags(2) = 9
    pipes.crt_flags(3) = 9
    pipes.crt_flags(4) = 9

   'pipes2 is filled by copying memory from pipes
    pipes2.number_of_fds = pipes.number_of_fds

    '※MoveMemory 排除: バイトループで crt_flags + os_handle を raw_bytes に詰め、STDIO_BUFFER の パッキング問題を解消させる
    '　→https://web.archive.org/web/20250524152550/https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA
    Dim iMem As Long
    ' crt_flags: 5 bytes
    For iMem = 0 To 4
        pipes2.raw_bytes(iMem) = pipes.crt_flags(iMem)
    Next iMem
    ' os_handle: 5 × LenB(LongPtr) bytes
    Dim handleBytes() As Byte
    ReDim handleBytes(0 To LenB(pipes.os_handle(0)) - 1)
    Dim hIdx As Long, bIdx As Long
    Dim destOffset As Long: destOffset = 5
    For hIdx = 0 To 4
        Dim hVal As LongPtr: hVal = pipes.os_handle(hIdx)
#If Win64 Then
        ' 8 bytes per handle (LongLong)
        Dim rawHandle As LongLong: rawHandle = hVal
        pipes2.raw_bytes(destOffset) = CByte(rawHandle And &HFF^)
        pipes2.raw_bytes(destOffset + 1) = CByte((rawHandle \ &H100^) And &HFF^)
        pipes2.raw_bytes(destOffset + 2) = CByte((rawHandle \ &H10000^) And &HFF^)
        pipes2.raw_bytes(destOffset + 3) = CByte((rawHandle \ &H1000000^) And &HFF^)
        pipes2.raw_bytes(destOffset + 4) = CByte((rawHandle \ &H100000000^) And &HFF^)
        pipes2.raw_bytes(destOffset + 5) = CByte((rawHandle \ &H10000000000^) And &HFF^)
        pipes2.raw_bytes(destOffset + 6) = CByte((rawHandle \ &H1000000000000^) And &HFF^)
        pipes2.raw_bytes(destOffset + 7) = CByte((rawHandle \ &H100000000000000^) And &HFF^)
        destOffset = destOffset + 8
#Else
        ' 4 bytes per handle (Long)
        Dim rawHandleL As Long: rawHandleL = CLng(hVal)
        pipes2.raw_bytes(destOffset) = CByte(rawHandleL And &HFF&)
        pipes2.raw_bytes(destOffset + 1) = CByte((rawHandleL \ &H100&) And &HFF&)
        pipes2.raw_bytes(destOffset + 2) = CByte((rawHandleL \ &H10000) And &HFF&)
        pipes2.raw_bytes(destOffset + 3) = CByte((rawHandleL \ &H1000000) And &HFF&)
        destOffset = destOffset + 4
#End If
    Next hIdx

   'Define new instance parameters and start it
    With start
        .cb = LenB(start)
        .dwFlags = STARTF_USESHOWWINDOW
        .hStdOutput = hStdOutWr
        .hStdInput = hStdInRd
        .hStdError = hStderrOutWr
        .wShowWindow = setShowWindow
        .cbReserved2 = LenB(pipes2)
        .lpReserved2 = VarPtr(pipes2)
    End With

    result = CreateProcess(0&, StrPtr(strExec), sa, sa, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    '起動アプリハンドルを取得
    hProcess = proc.hProcess
    dwProcessId = proc.dwProcessId

    '実行結果チェック
    If result = 0 Then
        'エラーコードを返す
        InitCDP = Err.LastDllError

        '失敗なので、自身で使用するはずだったハンドルも削除
        CloseHandleBrowser
    Else
        '成功は、0として返却
        InitCDP = 0
    End If

   'We close the sides of the handles that we dont need anymore
    CloseHandle hStdOutWr
    CloseHandle hStdInRd
    CloseHandle hCDPOutWr
    CloseHandle hCDPInRd
    CloseHandle hStderrOutWr

End Function



'***************************************************************************************************
'                              ■■■ Pipe関連操作 ■■■
'***************************************************************************************************
'* 機能　　：パイプハンドル情報をExcelのテーブルに記録します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   記録する識別名称
'---------------------------------------------------------------------------------------------------
'* 注意事項：ここでの、Dictionary.Add の Item 引数は `(hoge)` と括弧必須。64bit Dictionary が LongPtr を
'            参照渡しと誤解して 稀にクラッシュする不具合を回避するための強制値渡しです。
'***************************************************************************************************
Public Sub serialize(UserName As String)
    '0. データ記録準備
    Dim tmp As New Dictionary

    '1. 必要なデータを`Dictionary`に詰める
    tmp.Add "hStdOutRd", (hStdOutRd)
    tmp.Add "hStderrOutRd", (hStderrOutRd)
    tmp.Add "hStdInWr", (hStdInWr)
    tmp.Add "hCDPOutRd", (hCDPOutRd)
    tmp.Add "hCDPInWr", (hCDPInWr)
    tmp.Add "hProcess", (hProcess)
    tmp.Add "dwProcessId", (dwProcessId)

    '2. Excelのテーブルへ記録する
    Set ShSetting01_StartBrowser.TableBrowserHandle(UserName, "CDPCore.serialize") = tmp

End Sub

'***************************************************************************************************
'* 機能　　：Excelのテーブルに記録されてるパイプハンドル情報を読み込みます
'---------------------------------------------------------------------------------------------------
'* 返り値  ：設定の読み取りができたかの論理値
'* 引数　　：UserName   読み取る識別名称
'***************************************************************************************************
Public Function deserialize(UserName As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".deserialize"


    'Excelテーブルから、`UserName`に紐づくパイプハンドル情報を読み込む
    Dim PipeInfo As Dictionary
    Set PipeInfo = ShSetting01_StartBrowser.TableBrowserHandle(UserName, FromProcedureName)

    '取得に失敗した場合は、Excelのテーブルに記録されてないので抜ける
    If PipeInfo Is Nothing Then printMsg WARN_, "Pipe information for '" & UserName & "' is not recorded in the table.", FromProcedureName: Exit Function

    'クラス内変数に適用
    hStdOutRd = PipeInfo("hStdOutRd")
    hStderrOutRd = PipeInfo("hStderrOutRd")
    hStdInWr = PipeInfo("hStdInWr")
    hCDPOutRd = PipeInfo("hCDPOutRd")
    hCDPInWr = PipeInfo("hCDPInWr")
    hProcess = PipeInfo("hProcess")
    dwProcessId = PipeInfo("dwProcessId")

    '成功サインを返す
    deserialize = True
End Function

'***************************************************************************************************
'* 機能　　：Pipe/WebSocketが生きてるか確認します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：ハンドル生存可否
'***************************************************************************************************
Public Function isAvailability() As Boolean
    Const FromProcedureName As String = ThisClassName & ".isAvailability"


    '各種判定
    Select Case False
        Case viaWebSocket Is Nothing
            'WebSocket生存判定
            isAvailability = viaWebSocket.isAvailability
    
            'ステータスを最新にしておく
            LastPipeError.PeekNamedPipe = viaWebSocket.IsErrorStylePeekNamedPipe

        Case viaWebView2 Is Nothing
            'WebView2生存判定
            isAvailability = viaWebView2.isAvailability

        Case Else
            'Pipe生存判定
            Dim tmp As Long
            If PeekNamedPipe(hCDPOutRd, ByVal 0&, 0&, ByVal 0&, tmp, ByVal 0&) <> 0 Then isAvailability = True
    
            'ステータスを最新にしておく
            LastPipeError.PeekNamedPipe = Err.LastDllError
            printMsg Debug_, WinApiError.GetMessage(LastPipeError.PeekNamedPipe, "kernel32"), FromProcedureName

    End Select
End Function

'***************************************************************************************************
'* 機能　　：各種ハンドル情報をクリーニングします
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に、`deserialize`をした状態にすること
'***************************************************************************************************
Public Sub CloseHandleBrowser()
    'WebSocket/WebView2モードならそっち側のClose処理を行う
    Select Case False
        Case viaWebSocket Is Nothing
            viaWebSocket.DisconnectCDP

        Case viaWebView2 Is Nothing
            viaWebView2.DisconnectCDP

        Case Else
            CloseHandle hStdOutRd
            CloseHandle hStdInWr
            CloseHandle hCDPOutRd
            CloseHandle hCDPInWr
            CloseHandle hStderrOutRd
            CloseHandle hProcess
    End Select
End Sub



'***************************************************************************************************
'                              ■■■ Pipeの読み書き ■■■
'***************************************************************************************************
'* 機能　　：標準出力エラーと標準出力を読み取ります
'---------------------------------------------------------------------------------------------------
'* 返り値　：基本は、読み取ったバイト数を返しますが、
'               -1：データなし or `PeekNamedPipe`でエラー
'               -2：`ReadFile`でエラー
'
'* 引数　　：strData            標準出力エラーあるいは、標準出力内容を格納したい変数。実質返り値です
'            ErrorMode          `True`  標準出力エラーで読み取り
'                               `False` 標準出力で読み取り
'---------------------------------------------------------------------------------------------------
'* 機能説明：読み取りは非ブロッキングです。データがなくても、次に進めれます。
'* 注意事項：・ロジック自体は、前述の`readProcCDP`と一緒ですが、そこまで重要視しないため、簡略化してます
'            ・実際のエラーコードも特に確認しません
'***************************************************************************************************
Private Function readProcSTD(ByRef strData As String, ErrorMode As Boolean) As Long
    '使用するパイプ通信判定
    Dim UsehStdRd As LongPtr: UsehStdRd = IIf(ErrorMode, hStderrOutRd, hStdOutRd)

    '有効なパイプハンドルでない場合は、即抜け
    If UsehStdRd = 0 Then Exit Function

    '処理が成功すると、戻り値は 0 以外になります。
    Dim lPeekData As Long
    PeekNamedPipe UsehStdRd, ByVal 0&, 0&, ByVal 0&, lPeekData, ByVal 0&

    'データがあったら実際に、データを入手する処理へ
    If lPeekData > 0 Then
        Dim Buffer() As Byte
        Dim bSuccess As Long
        ReDim Buffer(lPeekData - 1)
        bSuccess = ReadFile(UsehStdRd, ByVal VarPtr(Buffer(0)), lPeekData, 0&, 0&)

        If bSuccess = 1 Then
            'バイト配列をUTF-8に変換するか？
            If Utf8Converter Is Nothing Then
                'そのまんま返却
                strData = StrConv(Buffer, vbUnicode)
            Else
                '変換した上で、返却
                strData = Utf8Converter.BytesToString(Buffer)
            End If

            '取得サイズを返却
            readProcSTD = lPeekData
        Else
            readProcSTD = -2
        End If

    Else
        readProcSTD = -1
    End If

End Function

'***************************************************************************************************
'* 機能　　：標準出力、標準エラー出力を確認します
'***************************************************************************************************
Private Sub viewStdoutAndStderr()
    Const FromProcedureName As String = ThisClassName & ".viewStdoutAndStderr"
    Const DividerMark       As String = vbLf  '標準出力、標準エラー出力は、改行コード区切りであることを確認


    Dim strRes As String, intRes As Long

    'Sometimes edge writes to stdout so we clear stdout too.
    intRes = readProcSTD(strRes, False)
    If intRes > 0 Then
        If InStr(1, strRes, vbCrLf) Then strRes = Replace(strRes, vbCrLf, DividerMark)  'Windows用の改行コードがあったら一旦、Linux用改行コードに置換
        printMsg Debug_, "stdout = " & Replace(strRes, DividerMark, Chr(30)), FromProcedureName     '改行区切りだと、ログが見ずらくなるので専用の制御文字として置換
    End If

    'Sometimes edge writes to stderrout so we clear stderrout too.
    intRes = readProcSTD(strRes, True)
    If intRes > 0 Then
        If InStr(1, strRes, vbCrLf) Then strRes = Replace(strRes, vbCrLf, DividerMark)  'Windows用の改行コードがあったら一旦、Linux用改行コードに置換
        printMsg ERROR_, "stderrout = " & Replace(strRes, DividerMark, Chr(30)), FromProcedureName  '改行区切りだと、ログが見ずらくなるので専用の制御文字として置換
    End If
End Sub

'***************************************************************************************************
'* 機能　　：CDP-Jsonコマンド結果や、ブラウザからのイベント情報をストリート受信して蓄積させます
'---------------------------------------------------------------------------------------------------
'* 返り値　：読み取ったバイト数を返します
'* 引数　　：EndOfData      終端null文字か？`True`なら、「欠けなし」という意味になります
'            destruction    ストリーム蓄積させるか？`True`で、蓄積させません
'---------------------------------------------------------------------------------------------------
'* 機能説明：読み取りは非ブロッキングです。データがなくても、次に進めれます。各種エラーコードは、`Get property`で確認してください。
'* 注意事項：このメソッドでは直接のメッセージ受け取りは行いません
'            `EndOfData`を初期化せずに、使いまわす場合、次のような仕様になります
'            ・`PeekNamedPipe`にて、受信データ量が0あるいは、エラー：引数に渡した`EndOfData`の内容を保持します
'            ・`PeekNamedPipe`にて、受信データ量が1以上　　　　　　：終端null文字判定結果を`EndOfData`に更新します
'***************************************************************************************************
Private Function readProcCDP(ByRef EndOfData As Boolean, destruction As Boolean) As Long
    Dim lPeekData As Long
    Dim lngRes As Long

    '処理が成功すると、戻り値は 0 以外になります。
    lngRes = PeekNamedPipe(hCDPOutRd, ByVal 0&, 0&, ByVal 0&, lPeekData, ByVal 0&)

    'API関数:PeekNamedPipe が失敗した場合は、0 を返すため、エラーコードを記録
    If lngRes = 0 Then LastPipeError.PeekNamedPipe = Err.LastDllError: Exit Function

    'データがあったら実際に、データを入手する処理へ
    If lPeekData > 0 Then
        Dim bSuccess As Long
        ReDim Buffer(lPeekData - 1) As Byte
        bSuccess = ReadFile(hCDPOutRd, ByVal VarPtr(Buffer(0)), lPeekData, 0&, 0&)

        '0以外なら成功です
        If bSuccess Then
            '破棄しない場合は、ストリームに流し込む ※バッファー拡張は、`ADODB.Stream`が担います
            If Not destruction Then
                CDPPipeStream.Write Buffer  '`ADODB.Stream`に追加
                StrBufferCheck lPeekData    'バッファー拡張判定を行う
            End If

            '終端null文字があるかチェック
            EndOfData = (Buffer(lPeekData - 1) = 0)

            '取得サイズを返却
            readProcCDP = lPeekData
        Else
            'API関数:ReadFile が失敗した場合は、0 を返すため、エラーコードを記録
            LastPipeError.ReadFile = Err.LastDllError
        End If
    End If
End Function

'***************************************************************************************************
'* 機能　　：Pipe経由で、ブラウザへ命令を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：strData    完成したCDP-Json文字列コマンド
'***************************************************************************************************
Private Sub writeProcCDP(strData As String)
    Dim lngWritten As Long
    Dim ResultCode As Long

    'UTF-8バイト配列を使用するか？
    If Utf8Converter Is Nothing Then
        'VBAの「Unicode」文字列そのまんまで、コマンドを送ります
        ResultCode = WriteFile(hCDPInWr, ByVal strData, Len(strData), lngWritten, ByVal 0&)
    Else
        'VBAの「Unicode」文字列をUTF-8でエンコードしたバイト配列を取得します
        Dim BinData() As Byte: BinData = Utf8Converter.BytesFromString(strData)

        'そのバイト配列として、コマンドを送ります
        ResultCode = WriteFile(hCDPInWr, ByVal VarPtr(BinData(0)), UBound(BinData) + 1, lngWritten, ByVal 0&)
    End If

    '終端目印`vbnullchar`相当も送る ※前回の`WriteFile`が成功してる場合
    If ResultCode Then ResultCode = WriteFile(hCDPInWr, ByVal VarPtr(m_NullTerminator(0)), 1, lngWritten, ByVal 0&)

    '失敗時は、実際のエラーコードを記録
    If ResultCode = 0 Then LastPipeError.WriteFile = Err.LastDllError
End Sub



'***************************************************************************************************
'                              ■■■ WebSocketの読み書き ■■■
'***************************************************************************************************
'* 機能　　：WinHttpWebSocketReceive から得たデータを`--remote-debugging-pipe`処理ロジックで扱えるように蓄積します
'---------------------------------------------------------------------------------------------------
'* 機能説明：CDPCoreViaWebSocket-GetAsyncMessage を行うことで内部で`RaiseEvent`が起こり、ここのプロシージャが呼ばれます
'* 注意事項：受信サイズ0あるいは、`TakeWebSocket`にて破棄モードで指定した場合は、呼ばれません
'***************************************************************************************************
Private Sub viaWebSocket_WebSocketReceive(res() As Byte, resSize As Long, EOS As Boolean)
    Const FromProcedureName As String = ThisClassName & ".viaWebSocket_WebSocketReceive"


    'ストリームに流し込む ※バッファー拡張は、`ADODB.Stream`が担います
    CDPPipeStream.Write res
    printMsg Trace_, "Storing received messages...", FromProcedureName

    'メッセージの終わりであれば、その目印も追記
    If EOS Then CDPPipeStream.Write m_NullTerminator

    '取得サイズをcursor管理へ反映し、バッファー拡張判定を行う
    StrBufferCheck resSize + EOS * -1
    printMsg Trace_, "ResSize = " & resSize & " , EOS: " & EOS, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：WebSocket経由で、ブラウザへ命令を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：strData    完成したCDP-Json文字列コマンド
'***************************************************************************************************
Private Sub writeWebSocketCDP(strData As String)
    'WebSocketオブジェクト経由で送信処理を委ねる
    LastPipeError.WriteFile = viaWebSocket.SendCommandCDP(strData)
End Sub



'***************************************************************************************************
'                              ■■■ WebView2の読み書き ■■■
'***************************************************************************************************
'* 機能　　：`CDPCoreViaWebView2 - SendCommandCDP`から得たデータを、Pipe/WebSocketと同じ蓄積先へ
'            直接書き込みます
'---------------------------------------------------------------------------------------------------
'* 機能説明：WebView2はCOMコールバック経由で完成JSON文字列(欠けなし・フラグメントなし)を届けるため、
'            `viaWebSocket_WebSocketReceive`のようなバイト列蓄積とストリーム変換は不要で、
'            直接`responseCDP.strBuffer`へ書き込みます
'* 注意事項：・メッセージの取り出し自体(`TakeCDPMessage`/`searchNull`)はPipe/WebSocket共通のロジックを
'            　そのまま流用しており、`vbNullChar`をメッセージ区切りとして検索する前提になっているため、
'            　必ず末尾に終端目印の`vbNullChar`を1文字付け足してから書き込むこと
'***************************************************************************************************
Private Sub viaWebView2_CDPMessageReceived(RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".viaWebView2_CDPMessageReceived"


    '1. バッファー拡張判定を行う
    StrBufferCheck Len(RawJson) + 1 '`+1`は、`vbnullchar`補正

    '2. まずは、本体CDPJsonメッセージを追加
    AddCDPMessage RawJson

    '3. `vbnullchar`で〆
    AddCDPMessage vbNullChar
    printMsg Trace_, "Storing received messages (WebView2)...", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：WebView2経由で、ブラウザへ命令を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：strData    完成したCDP-Json文字列コマンド
'***************************************************************************************************
Private Sub writeWebView2CDP(strData As String)
    'WebView2オブジェクト経由で送信処理を委ねる
    LastPipeError.WriteFile = viaWebView2.SendCommandCDP(strData)
End Sub



'***************************************************************************************************
'                               ■■■ CDPコマンド実行前の下準備 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザへ命令する前の組み立てを行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：実行コマンドID
'* 引数　　：dicMessage Chromium ブラウザに送るCDP-Jsonコマンド(Dictionary型)。最上位要素に、`method`、`params`があるようにしてください。
'            SessionID  タブへの命令の場合は指定します
'---------------------------------------------------------------------------------------------------
'* 機能説明：`id`と`sessionId`の組み立てと紐づけ処理を行います
'***************************************************************************************************
Public Function ReadyRunCDP(dicMessage As Dictionary, Optional sessionID As String) As Long
    Const FromProcedureName As String = ThisClassName & ".ReadyRunCDP"


    '1. 現在のコマンドIDを発行  ※内部でカウントアップを行います
    ReadyRunCDP = IssuanceCommandID

    '2. `sessionId`check
    If StrPtr(sessionID) Then
        '2-1. `sessionId`付与
        dicMessage.Add "sessionId", sessionID

        '2-2. commandIDと`sessionId`の紐づけを行う
        '※`sessionId`付きコマンド実行時にて、エラーが起こると、`sessionId`なしエラーとして返る理不尽仕様に対応するため
        DictionarySessionID.Add ReadyRunCDP, sessionID
        printMsg Trace_, "commandID:" & ReadyRunCDP & " corresponds to SessionID:" & sessionID, FromProcedureName
    Else
        printMsg Trace_, "commandID:" & ReadyRunCDP & " does not have a SessionID.", FromProcedureName
    End If

    '3.`id`を付与
    dicMessage.Add "id", ReadyRunCDP

    '4. 実際に実行
    Dim strMessage As String
'    strMessage = BiDiCDPJson.StringifyValue(dicMessage)     '理想はこっちにして、1つに完結したいﾋﾟｴﾝ
    strMessage = WebJsonConverter.serialize(dicMessage)     'Dictionary→Json文字列化のみ`VBA-FastJSON`を使用。こっちのほうがパフォーマンスが良いため
    printMsg Debug_, "sendCDP - JsonCommand = " & strMessage, FromProcedureName

    '5. 通信モードに応じてPipe/WebSocket/WebView2分岐
    Select Case False
        Case viaWebSocket Is Nothing:  writeWebSocketCDP strMessage
        Case viaWebView2 Is Nothing:    writeWebView2CDP strMessage
        Case Else:                          writeProcCDP strMessage
    End Select
End Function

'***************************************************************************************************
'* 機能　　：CDP実行の際のコマンドIDを発行します
'---------------------------------------------------------------------------------------------------
'* 返り値　：発行されたcommandID
'---------------------------------------------------------------------------------------------------
'* 注意事項：20億を超えると、1にリセットされます
'***************************************************************************************************
Private Function IssuanceCommandID() As Long
    Const FromProcedureName As String = ThisClassName & ".IssuanceCommandID"


    '1. 上限チェック
    Static LastCommandID As Long
    If LastCommandID + 1 > LimitCommandID Then
        '1-1. リクエスト数の上限に達したら、0にリセットしておく
        LastCommandID = 0
        printMsg info_, "Reset to `1` to prevent internal counter overflow.", FromProcedureName

        '1-2. リセットした旨を伝える
        RaiseEvent ResetCommandID
    End If

    '2. カウントアップ
    LastCommandID = LastCommandID + 1   'We increase the global ID counter

    '3. 返却
    IssuanceCommandID = LastCommandID
End Function



'***************************************************************************************************
'                         ■■■ CDPレスポンス、蓄積/取得管理処理 ■■■
'***************************************************************************************************
'* 機能　　：適切な文字列変数バッファーが確保されてるかチェックし、確保されてない場合は、拡張を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：ResSize    `ReadFile/recv etc`から得たデータサイズそのまんま
'***************************************************************************************************
Private Sub StrBufferCheck(resSize As Long)
    Const FromProcedureName As String = ThisClassName & ".StrBufferCheck"


    '1. 新規メッセージがない場合は即抜け
    If resSize <= 0 Then printMsg WARN_, "No message was found, so the buffer check will be skipped.", FromProcedureName: Exit Sub

    '2. バッファーcheckへ
    With responseCDP
        'そもそもVBAで扱える文字列の上限を超える
        If .EndCursor + resSize > MAX_STR_LEN Then
            printMsg ERROR_, "The VBA string limit (1,073,741,823 characters) has been exceeded.", FromProcedureName, doRaiseError:=True

        '今のバッファーサイズを越しそうになったら、拡張する(高負荷ポイント)
        ElseIf .EndCursor + resSize > .length Then
            '2-1. 拡張サイズを決める
            Dim AddLength  As Long
            If .length > (MAX_STR_LEN \ 2) Then
                '次に2倍すると上限を超えるので、最後のバインバインは上限ピッタリに着地！
                AddLength = MAX_STR_LEN - .length
                printMsg WARN_, "This is the final expansion. Further expansions are not possible. Add buffer size: " & AddLength, FromProcedureName
            Else
                'まずは、「バインバイン（2倍）」で設定
                AddLength = .length

                'バインバインでも足りないかcheck
                If .EndCursor + resSize > .length + AddLength Then
                    '加算
                    AddLength = (.EndCursor + resSize) - .length
                    printMsg info_, "Expand the buffer. Add buffer size: " & AddLength, FromProcedureName
                Else
                    'バインバインのままで
                    printMsg info_, "Expand the buffer(x2). Add buffer size: " & AddLength, FromProcedureName
                End If
            End If

            '2-2. 実際に拡張
            .strBuffer = .strBuffer & String$(AddLength, vbNullChar)

            '2-3. バッファーサイズを同期
            .length = .length + AddLength

            '2-4. ログ化
            printMsg Debug_, "The buffer has been expanded: " & .length, FromProcedureName
        End If

        '3. 有効予定なCDPメッセージカーソル末尾位置を更新
        .EndCursor = .EndCursor + resSize
    End With
End Sub

'***************************************************************************************************
'* 機能　　：ストリームからテキスト変換したのを`responseCDP.strBuffer`に反映して、メッセージ取り出しの準備を行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：これを呼び出す際は`ADODB.Stream`の有効な`Stream`範囲の末尾が`vbnullchar`になってることを前提とします
'***************************************************************************************************
Private Sub ReadyCDPMessage()
    Const FromProcedureName As String = ThisClassName & ".ReadyCDPMessage"


    With CDPPipeStream
        '1. CDPからのメッセージがまだない場合は即抜け
        If .Position = 0 Then printMsg Debug_, "The stored messages are empty.", FromProcedureName: Exit Sub

        '2. 取り出すバイト数を確定
        .SetEOS
        .Position = 0   'CDPメッセージを取り出すため、先頭に巻き戻す

        '3. UTF-16変数に格納する準備
        Dim tmp As String
        If Utf8Converter Is Nothing Then
            'そのまんまUTF-16変数に格納
            AddCDPMessage StrConv((.Read), vbUnicode)
        Else
            'UTF-8に変換した上で、UTF-16変数に格納
            AddCDPMessage Utf8Converter.BytesToString((.Read))
        End If

        '4 .追加出来たので、カーソル位置をリセットし、次の新着メッセージバイト蓄積に備える
        .Position = 0
    End With
End Sub

'***************************************************************************************************
'* 機能　　：バッファーにCDPメッセージを追加します
'---------------------------------------------------------------------------------------------------
'* 引数    ：TextJsonCDP    CDPからのテキストJsonメッセージ
'---------------------------------------------------------------------------------------------------
'* 詳細説明：内部でcursor管理を行い、`Mid$`による高速な追記を行います
'* 注意事項：元がバイト配列の場合は事前に、テキスト変換が必要です
'***************************************************************************************************
Private Sub AddCDPMessage(TextJsonCDP As String)
    With responseCDP
        '1. UTF-16変数に格納
        Mid$(.strBuffer, .AddCursor + 1) = TextJsonCDP

        '2. 次の蓄積位置を記録
        .AddCursor = .AddCursor + Len(TextJsonCDP)    '追加した文字数分をシフト
        .EndCursor = .AddCursor                       'テキスト変換による差分補正

        '3. 文字列メッセージとして取り出せるので、`True`に更新
        .ExistsNewRes = True
    End With
End Sub

'***************************************************************************************************
'* 機能　　：呼び出す度に1件1件、CDPメッセージを取り出します
'***************************************************************************************************
Private Function TakeCDPMessage() As String
    Const FromProcedureName As String = ThisClassName & ".TakeCDPMessage"


    With responseCDP
        '1. 次に蓄積し始める位置が先頭の場合は即抜け
        If .AddCursor = 0 Then printMsg WARN_, "There are no new responses.", FromProcedureName: Exit Function

        '2. 今の位置から`vbnullchar`の位置を探索
        Dim EndPos As Long
        EndPos = searchNull(.strBuffer, .StartCursor + 1)

        '3. 見つからないあるいは、`.AddCursor`を超える場合は、メッセージ取り出しできないので、抜ける
        If EndPos = 0 Or EndPos > .AddCursor Then printMsg WARN_, "There are new messages, but some are still missing.", FromProcedureName: Exit Function

        '4. その目印を元に、1件分返却
        TakeCDPMessage = Mid$(.strBuffer, .StartCursor + 1, EndPos - .StartCursor - 1)

        '5. 次の検索開始位置を更新
        If .AddCursor = EndPos Then
            '次に蓄積し始める位置にたどり着いちゃったら、カーソル情報をリセットしておく
            .EndCursor = .EndCursor - .AddCursor
            .AddCursor = 0
            .StartCursor = 0
            .ExistsNewRes = False   '有効な文字列メッセージは全部取り出せたので`False`に更新

            printMsg Debug_, "All accumulated CDPmessages have been processed. Resetting cursor information.", FromProcedureName
        Else
            '次の取り出しに備えて更新
            .StartCursor = EndPos
        End If
    End With
End Function

Private Function searchNull(checkString As String, StartPos As Long) As Long
'------------------------------------------------------------
' CDP messages received from chrome are null-terminated
' Updated: 25/10/25: Daniel Polak - new faster version
'------------------------------------------------------------

    Dim lngPos As Long

    If LenB(checkString) = 0 Then Exit Function
    lngPos = InStr(StartPos, checkString, vbNullChar, vbBinaryCompare)
    If lngPos > 0 Then
        searchNull = lngPos
    Else
        searchNull = 0
    End If

End Function



'***************************************************************************************************
'                             ■■■ 各種CDPメッセージ取り出し ■■■
'***************************************************************************************************
'* 機能　　：パイプから、データ欠損がなくなるまで、ループして取り出していきます
'---------------------------------------------------------------------------------------------------
'* 引数　　：StopPipeError  `True` ：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあったら停止します
'                           `False`：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあっても停止しません
'                                    別途、`LastError○○`で後で確認可能です
'
'            destruction    `True` :ストリームに蓄積しません。破棄用にどうぞ
'                           `False`:ストリームに蓄積し、ID、イベント解析に備えます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・不整合を防ぐため、終端ヌル文字(vbnullchar)が来るまで、受信を繰り返すようにしてます
'            ・また、カーネルバッファを1刻も早く「空」にするための判定も含めています
'***************************************************************************************************
Private Sub TakePipe(StopPipeError As Boolean, destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakePipe"


    Dim intRes As Long
    printMsg Trace_, "Starts retrieving data from the pipe...", FromProcedureName

    '終端null文字があって尚且つ、受信データ0が来るまで、ループ受信を行います
    Dim LoopCounter As Long
    Dim EOvbNullChar As Boolean: EOvbNullChar = True  '初期化
    Do
        '一定の回数ごとに`DoEvents`を呼ぶ。※初回ループは、必ず呼ぶ
        If LoopCounter Mod RunDoEventsCount = 0 Then DoEvents
        LoopCounter = LoopCounter + 1


        'std系の中身を出力
        viewStdoutAndStderr

        '受信処理
        intRes = readProcCDP(EOvbNullChar, destruction) '※内部で、`ADODB.Stream`へバイト配列が蓄積されます
        printMsg Trace_, "ResSize = " & intRes, FromProcedureName

        '返り値チェック
        If intRes > 0 Then
            '破棄モードかの判定 ※ログメッセージ分岐のみ
            If destruction Then
                '解析などせず、そのまんま出して破棄します
                printMsg Trace_, "Discarding data in pipe...", FromProcedureName
            Else
                '中身の解析に備えて、蓄積準備
                printMsg Trace_, "Storing received messages...", FromProcedureName
            End If
        Else
            'データなしの場合は、エラーコードチェックを行う
            If LastPipeError.PeekNamedPipe Then
                'PeekNamedPipe にてエラーがあった場合
                printMsg ERROR_, "Error PeekNamedPipe in readProcCDP. ErrorCode:" & LastPipeError.PeekNamedPipe & " - " & WinApiError.GetMessage(LastPipeError.PeekNamedPipe, "kernel32"), FromProcedureName, doRaiseError:=StopPipeError
                Exit Sub

            ElseIf LastPipeError.ReadFile Then
                'ReadFile にてエラーがあった場合
                printMsg ERROR_, "Error ReadFile in readProcCDP. ErrorCode:" & LastPipeError.ReadFile & " - " & WinApiError.GetMessage(LastPipeError.ReadFile, "kernel32"), FromProcedureName, doRaiseError:=StopPipeError
                Exit Sub

            Else
                'パイプ内は空
                printMsg Trace_, "The ReadFile is currently empty.", FromProcedureName

            End If
        End If
    Loop Until EOvbNullChar And intRes < 1
    printMsg Trace_, "Finished getting data from the pipe.", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：WebSocketから、データ欠損がなくなるまで、ループして取り出していきます
'---------------------------------------------------------------------------------------------------
'* 引数　　：StopWebSocketError `True` ：WebSocketに、エラーがあったら停止します
'                               `False`：WebSocketに、エラーがあっても停止しません
'                                        別途、`LastError○○`で後で確認可能です。統合の都合上、エラー変数名はPipe準拠です
'
'            destruction        `True` :ストリームに蓄積しません。破棄用にどうぞ
'                               `False`:ストリームに蓄積し、ID、イベント解析に備えます
'---------------------------------------------------------------------------------------------------
'* 注意事項：・不整合を防ぐため、`isFIN:=True`が来るまで、受信を繰り返すようにしてます
'            ・また、カーネルバッファを1刻も早く「空」にするための判定も含めています
'***************************************************************************************************
Private Sub TakeWebSocket(StopWebSocketError As Boolean, destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakeWebSocket"


    Dim intRes As Long
    printMsg Trace_, "Starts retrieving data from the WebSocket...", FromProcedureName

    '終了サインかつ、受信データなしになるまで、ループ受信を行います
    Dim LoopCounter As Long
    Dim isFIN As Boolean: isFIN = True  '初期化
    Do
        '一定の回数ごとに`DoEvents`を呼ぶ。※初回ループは、必ず呼ぶ
        If LoopCounter Mod RunDoEventsCount = 0 Then DoEvents
        LoopCounter = LoopCounter + 1


        '受信処理
        intRes = viaWebSocket.GetAsyncMessage(isFIN, destruction)   '※内部で`WithEvents`が走り、そこで`ADODB.Stream`へバイト配列が蓄積されます

        '返り値チェック
        If intRes > 0 Then
            '破棄モードかの判定 ※ログメッセージ分岐のみ
            If destruction Then
                '解析などせず、そのまんま出して破棄します
                printMsg Trace_, "Discarding data in WebSocket...", FromProcedureName
            Else
                '中身の解析に備えて、蓄積準備
                printMsg Trace_, "Storing received messages...", FromProcedureName
            End If
        Else
            'データなしの場合は、エラーコードチェックを行う
            If viaWebSocket.IsErrorStylePeekNamedPipe Then
                'ioctlsocket にてエラーがあった場合
                LastPipeError.PeekNamedPipe = viaWebSocket.IsErrorStylePeekNamedPipe    'エラーコードを同期させておく
                printMsg ERROR_, "Error ioctlsocket in GetAsyncMessage. ErrorCode:" & viaWebSocket.IsErrorStylePeekNamedPipe & " - " & WinApiError.GetMessage(viaWebSocket.IsErrorStylePeekNamedPipe, "ws2_32"), FromProcedureName, doRaiseError:=StopWebSocketError
                Exit Sub

            ElseIf viaWebSocket.IsErrorStyleReadFile Then
                'recv にてエラーがあった場合
                LastPipeError.ReadFile = viaWebSocket.IsErrorStyleReadFile              'エラーコードを同期させておく
                printMsg ERROR_, "Error recv in GetAsyncMessage. ErrorCode:" & viaWebSocket.IsErrorStyleReadFile & " - " & WinApiError.GetMessage(viaWebSocket.IsErrorStyleReadFile, "ws2_32"), FromProcedureName, doRaiseError:=StopWebSocketError
                Exit Sub

            ElseIf viaWebSocket.LastBufferType = WS_CLOSE_MESSAGE Then
                '切断要求が来た場合
                printMsg ERROR_, "Disconnected because a disconnection request was received from the Chromium browser.", FromProcedureName, doRaiseError:=StopWebSocketError
                Exit Sub
            Else
                'パイプ内は空あるいは、WebSocketフレーム単位としてまだ届いていない
                printMsg Trace_, "The WebSocket is currently empty.", FromProcedureName

            End If
        End If
    Loop Until isFIN And intRes < 1

End Sub

'***************************************************************************************************
'* 機能　　：WebView2からの、保留中のコールバックを処理させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：StopApiError   (このモードでは未使用。Pipe/WebSocket版とのシグネチャ統一のためだけに存在)
'            destruction    `True` :今回`DoEvents`中に届いたメッセージも蓄積せず、破棄します
'                           `False`:蓄積し、ID、イベント解析に備えます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：WebView2のコールバックは`DoEvents`実行中に(同期的に)発火する。`PeekNamedPipe`の
'            ような「残量チェック」の手段がないため、カーソルの変化がなかったら即抜けするロジックにする
'***************************************************************************************************
Private Sub TakeWebView2(StopApiError As Boolean, destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakeWebView2"


    Dim BeforeEndCursor As Long
    Do
        '1. 内部でWebView2のコールバックを呼ぶ試行を行う。溜まってる場合、そこに仕込んだ`RaiseEvent`で、CDPメッセージを蓄積させる処理にワープします
        '※`DoEvents`自体がイベント発火のきっかけになることから、「一定回数ごとに`DoEvents`」といったことはしません
        DoEvents

        '2. カーソルの変化がない場合は、ループから抜ける
        If BeforeEndCursor = responseCDP.EndCursor Then Exit Do

        '3. 前回値を記録
        BeforeEndCursor = responseCDP.EndCursor
    Loop

    If destruction Then
        '`DoEvents`中に届いてしまった分を明示的に破棄する
        With responseCDP
            .AddCursor = 0
            .StartCursor = 0
            .EndCursor = 0
            .ExistsNewRes = False
        End With
        printMsg Trace_, "Discarding data in WebView2...", FromProcedureName
    End If
End Sub

'***************************************************************************************************
'* 機能　　：CDPからバイト配列を一括で受け取り、テキストに変換します
'---------------------------------------------------------------------------------------------------
'* 引数　　：StopApiError   `True` ：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあったら停止します
'                           `False`：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあっても停止しません
'                                    別途、`LastError○○`で後で確認可能です。統合の都合上、エラー変数名はPipe準拠です
'
'            destruction    `True` :ストリームに蓄積しません。破棄用にどうぞ。`responseCDP.ExistsNewRes`は更新されません。
'                           `False`:ストリームに蓄積し、ID、イベント解析に備えます。新規データがあれば、`responseCDP.ExistsNewRes`が`True`になります。
'***************************************************************************************************
Public Sub AnalyzeCDP(Optional StopApiError As Boolean, Optional destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".AnalyzeCDP"


    '1. 通信モードに応じてPipe/WebSocket/WebView2分岐し、受信メッセージバイト配列の取り出し
    Select Case False
        Case viaWebSocket Is Nothing:  TakeWebSocket StopApiError, destruction
        Case viaWebView2 Is Nothing:    TakeWebView2 StopApiError, destruction
        Case Else:                          TakePipe StopApiError, destruction
    End Select
    
    '2. `ReadFile` or `recv` で得たバイト配列をテキストに変換させ、メッセージ処理を準備させる
    'WebView2モードは、`viaWebView2_CDPMessageReceived`が`responseCDP.strBuffer`へ直接書き込む
    'ため、`CDPPipeStream`経由のテキスト変換(`ReadyCDPMessage`)は不要かつ不整合の元になるので
    '呼ばない
    If Not destruction Then
        If viaWebView2 Is Nothing Then Call ReadyCDPMessage
    End If

    '3. ログ化
    If destruction Then
        printMsg Trace_, "The disposal of the accumulated CDP data has been completed.", FromProcedureName
    Else
        printMsg Trace_, "Preparation for CDPmessage analysis is complete.", FromProcedureName
    End If
End Sub



'***************************************************************************************************
'                               ■■■ レスポンス捌き ■■■
'***************************************************************************************************
'* 機能　　：`responseCDP.strBuffer`内から、1件分のメッセージ解析を行い、イベント発火まで行います
'---------------------------------------------------------------------------------------------------
'* 詳細説明：イベントデータ or CDP-Jsonコマンド結果 or 構文エラー の確認をし、`RaiseEvent`で1件分、発信します
'* 注意事項：このメソッドの場合、新着受信の`AnalyzeCDP`、新着確認の`NewResCDP`は、別途呼び出し側で用意/制御する必要があります。
'            特定イベント検出時の「即時割り込み&残りのイベント処理ストップ」等といった上級者向けの手動メソッドです。
'***************************************************************************************************
Public Sub TakeEvent()
    Const FromProcedureName As String = ThisClassName & ".TakeEvent"


    '1. ログによるレスの確認
    Dim tmp As String: tmp = TakeCDPMessage
    If currentLog.seeRawSendMsgDbg Then printMsg Debug_, "Result - CDP          : " & tmp, FromProcedureName  'Only enable if needs to see every raw JS message for debugging/development

    '2. 受信データを`RaiseEvent`するプロシージャへ委託
    Call BrowserReceivedDataCheck(tmp)
    printMsg Trace_, "I triggered one `RaiseEvent`.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：`responseCDP.strBuffer`内にある全メッセージ解析を行い、イベント発火まで行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：StopApiError   `True` ：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあったら停止します
'                           `False`：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあっても停止しません
'                                    別途、`LastError○○`で後で確認可能です
'
'            destruction    `True` :メッセージ解析をせずに溜まったデータの取り出しのみ行います
'                           `False`:ID、イベント解析を行いつつ、溜まったデータの取り出しも並行して行います
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・イベントデータ or CDP-Jsonコマンド結果 or 構文エラー の確認をし、各種`RaiseEvent`で全件分、発信します
'            ・新規メッセージ取り出し&チェックから全件イベント発火まで一通り、行えます。基本はこちらを利用します
'* 注意事項：このメソッドは、届いている分だけを「スナップショット」として吸い上げ、Do While ループで`RaiseEvent`する全自動方式を採用してます
'***************************************************************************************************
Public Sub TakeEvents(Optional StopApiError As Boolean = True, Optional destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakeEvents"


    '1. CDPからバイト配列を受け取り、テキストへ変換
    AnalyzeCDP StopApiError, destruction

    '2. 新規メッセージがあれば、ループで全てを取り出し終わるまで、`RaiseEvent`します
    Dim ResCounter As Long
    Do While responseCDP.ExistsNewRes
        '3-1. 一定の回数ごとに`DoEvents`を呼ぶ。※初回ループは、必ず呼ぶ
        If ResCounter Mod RunDoEventsCount = 0 Then DoEvents
        ResCounter = ResCounter + 1

        '3-2. 1個ずつ取り出して`RaiseEvent`するプロシージャへ
        printMsg Debug_, "Executing the " & ResCounter & " `RaiseEvent`...", FromProcedureName
        TakeEvent
    Loop

    '3. CDPメッセージ処理件数をログ化
    If ResCounter Then
        printMsg Debug_, "Executed " & ResCounter & " `RaiseEvent` calls.", FromProcedureName
    Else
        printMsg Trace_, "There were no messages received this time.", FromProcedureName
    End If
End Sub

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックし、各クラスへイベント発信をします
'---------------------------------------------------------------------------------------------------
'* 引数　　：CDPJsonString          1件分のJson形式文字列のブラウザからのイベントやCDP-Jsonコマンド結果
'---------------------------------------------------------------------------------------------------
'* 詳細説明：イベントデータ or CDP-Jsonコマンド結果 or 構文エラー の確認をし、各種`RaiseEvent`で発信します
'* 注意事項：構文エラーの場合、不整合により不安定につながるため、エラーで停止します
'***************************************************************************************************
Private Sub BrowserReceivedDataCheck(CDPJsonString As String)
    Const FromProcedureName As String = ThisClassName & ".BrowserReceivedDataCheck"


    Dim ResultCommandID As Long, AsynchronousEventName As String, sessionID As String
    Dim CDPNode As BiDiCDPJson: Set CDPNode = BiDiCDPJson.Parse(CDPJsonString)

    '--------- 1. ブラウザからのイベントか？ ---------
    If CDPNode.ExistsKey("method") Then
        '非同期イベント名を取得
        AsynchronousEventName = CDPNode.StringKey("method")

        '--------- 2. タブに対する非同期イベントか？ ---------
        If CDPNode.ExistsKey("sessionId") Then
            'タブに対する非同期イベントとして発信
            sessionID = CDPNode.StringKey("sessionId")
            RaiseEvent CDPContextEvent(AsynchronousEventName, CDPJsonString, sessionID)
        Else
            'ブラウザに対する非同期イベントとして発信
            RaiseEvent CDPBrowserEvent(AsynchronousEventName, CDPJsonString)
        End If
    Else
        '--------- 2. `id`項目があるか？ ---------
        If CDPNode.ExistsKey("id") Then
            'id値を取得
            ResultCommandID = CDPNode.NumberKey("id")

            '--------- 3. タブに対するCDP-Jsonコマンド結果か？ ---------
            If DictionarySessionID.Exists(ResultCommandID) Then    '`id`項目での`sessionId`checkは、送信時に登録した`Dictionary`checkに委ねる
                'タブに対するコマンド結果として発信
                sessionID = DictionarySessionID(ResultCommandID)
                RaiseEvent CDPContextID(ResultCommandID, CDPJsonString, sessionID)

                '発信したので削除しておく
                DictionarySessionID.Remove (ResultCommandID)
            Else
                'ブラウザに対するコマンド結果として発信
                RaiseEvent CDPBrowserID(ResultCommandID, CDPJsonString)
            End If
        Else
            '`method`も`id`項目もない場合、CDP-Jsonコマンドの構文エラーとして返った可能性があるとし、不整合を防ぐため、エラーで止める
            printMsg ERROR_, "A syntax error was detected when executing with ExecuteCDP/ExecuteCDPAsync. Processing was interrupted to prevent inconsistencies." & vbCrLf & "> " & CDPJsonString, FromProcedureName, doRaiseError:=True
        End If
    End If
End Sub



'***************************************************************************************************
'                               ■■■ プロパティ機能 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザ側に対する各種パイプハンドル取得用プロパティ
'***************************************************************************************************
Property Get ConnectionPipe_hStdOutRd() As LongPtr      '標準出力 ※`--remote-debugging-pipe`による起動のみ
    ConnectionPipe_hStdOutRd = hStdOutRd
End Property

Property Get ConnectionPipe_hStderrOutRd() As LongPtr   '標準エラー出力 ※`--remote-debugging-pipe`による起動のみ
    ConnectionPipe_hStderrOutRd = hStderrOutRd
End Property

Property Get ConnectionPipe_hStdInWr() As LongPtr       '標準入力 ※`--remote-debugging-pipe`による起動のみ
    ConnectionPipe_hStdInWr = hStdInWr
End Property

Property Get ConnectionPipe_hCDPOutRd() As LongPtr      'CDP用標準出力(CDP結果や非同期イベントを取得)
    ConnectionPipe_hCDPOutRd = hCDPOutRd
End Property

Property Get ConnectionPipe_hCDPInWr() As LongPtr       'CDP用標準入力(CDPコマンド送信)
    ConnectionPipe_hCDPInWr = hCDPInWr
End Property

Property Get Connection_hProcess() As LongPtr           '`--remote-debugging-pipe`で起動した際のhProcess
    Connection_hProcess = hProcess
End Property

Property Get Connection_dwProcessId() As LongPtr        '`--remote-debugging-pipe`で起動した際のProcessID
    Connection_dwProcessId = dwProcessId
End Property

'***************************************************************************************************
'* 機能　　：最後の○○系取得用プロパティ
'***************************************************************************************************
Property Get LastErrorPeekNamedPipe() As Long           '`PeekNamedPipe`使用時のエラーコード
    LastErrorPeekNamedPipe = LastPipeError.PeekNamedPipe
End Property

Property Get LastErrorReadFile() As Long                '`ReadFile`使用時のエラーコード
    LastErrorReadFile = LastPipeError.ReadFile
End Property

Property Get LastErrorWriteFile() As Long               '`WriteFile`使用時のエラーコード
    LastErrorWriteFile = LastPipeError.WriteFile
End Property

'***************************************************************************************************
'* 機能　　：ログ設定値、取得用プロパティ
'***************************************************************************************************
Property Get ShowLogCDPresponse() As Boolean
    ShowLogCDPresponse = currentLog.seeRawSendMsgDbg
End Property

'***************************************************************************************************
'* 機能　　：新規CDPレスポンスがあるかどうか？
'***************************************************************************************************
Property Get NewResCDP() As Boolean
    NewResCDP = responseCDP.ExistsNewRes
End Property

'***************************************************************************************************
'* 機能　　：WebSocketモードでのメッセージ取得/解析に切り替えます
'***************************************************************************************************
Property Set WebSocketMode(ws As CDPCoreViaWebSocket)
    Set viaWebSocket = ws
End Property

'***************************************************************************************************
'* 機能　　：WebView2モードでのメッセージ取得/解析に切り替えます
'***************************************************************************************************
Property Set WebView2Mode(wv As CDPCoreViaWebView2)
    Set viaWebView2 = wv
End Property



'***************************************************************************************************
'                               ■■■ その他ヘルパーメソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, LogID As String, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
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
    Else strFormattedMsg = " | " & LogID & " | " & strMsg

   'Perform msg logging if needed
    If currentLog.doLog Then
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".cdp.log"
        If LenB(Dir(currentLog.logPath, vbDirectory)) = 0 Then MkDir currentLog.logPath
    End If

   'Perform msg output to Immediate Window if needed
    With LogControl
        Select Case LogLevel_
            Case LogLevelName.Trace_: .LogTrace strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.Debug_: .LogDebug strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.info_:   .LogInfo strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.WARN_:   .LogWarn strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.ERROR_: .LogError strFormattedMsg, From, currentLog.logPath, logName, Err.Number
            Case Else: 'ログ出力無効化
        End Select
    End With


    '主に、CDPコマンドのエラー時に止まります
    If doRaiseError Then Err.Raise CDPCustomErrorCodes.Protocol, From, Description:=strMsg

End Sub

'***************************************************************************************************
'* 機能　　：呼び出し時点での`CDPPipeStream`の中身をファイルに保存します
'---------------------------------------------------------------------------------------------------
'* 引数    ：FolderPath 保存フォルダ名
'            LogName    保存ログ名
'---------------------------------------------------------------------------------------------------
'* 詳細説明：デバック用としてお使い下さい
'* 注意事項：`ReadyCDPMessage`を行うたびに上書きする点に注意。Public的には、`TakeEvents`で上書きします
'***************************************************************************************************
Public Sub SaveCDPPipeStream(FolderPath As String, FileName As String)
    With CDPPipeStream
        '1, 今の位置を記録
        Dim tmpPos As Long: tmpPos = .Position
        .SetEOS

        '2. 保存
        .Position = 0
        .SaveToFile FolderPath & "\" & FileName, 2

        '3. 元に戻す
        .Position = tmpPos
    End With
End Sub



'***************************************************************************************************
'                               ■■■ 初期化/後始末処理 ■■■
'***************************************************************************************************
'* 機能　　：各種設定を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    '------------------------- クラス内変数の一部を初期化 -------------------------
    '1. 初期バッファー設定
    '  文字列バッファーの初期化
    With responseCDP
        .length = InitialBuffer + dwAllocationGranularity   'メモリアロケーション粒度を考慮
        .strBuffer = String$(.length, vbNullChar)
    End With

    '  バイト配列バッファーの初期化
    ReDim binBuffer(InitialBuffer - 1) As Byte
    Set CDPPipeStream = CreateObject("ADODB.Stream")
    With CDPPipeStream
        .Type = 1           ' adTypeBinary
        .Open
        .Write binBuffer    ' 疑似的に大きめのバッファーを確保
        .Position = 0       ' カーソルだけ先頭に戻す
        .SetEOS             ' 終端（EOS）としてマークする
    End With

    '2. 設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        '2-1. イミディエイトに関する設定を行う
        currentLog.seeRawSendMsgDbg = .UseRangeID(8, FromProcedureName) 'ブラウザから受信した生のJSONメッセージもイミディエイトに表示させるか？

        '2-2. パスが設定されてたら、ログファイル保存を有効化
        Dim logFolderPath As String: logFolderPath = .UseRangeID(9, FromProcedureName)
        If LenB(logFolderPath) > 0 Then
            currentLog.doLog = True                'ログファイルとして保存するか？
            currentLog.logPath = logFolderPath     '保存先のフォルダパスを設定
        End If

        '2-3. Assign an objID for easy debug identification
        LogID = "COID" & Format(Rnd * 1000, "000")

        '2-4. UTF8変換モードを設定するか
        If .UseRangeID(10, FromProcedureName) Then Set Utf8Converter = New CharacterCodeConversion

        '2-5. 初期起動ウィンドウ状態を設定
        setShowWindow = .UseRangeID(12, FromProcedureName)
    End With

    '3. コマンドIDに紐づくSessionIDの管理を初期化
    Set DictionarySessionID = New Dictionary

    '4. 実行マシンの「1秒間の振動数（周波数）」を取得して、bas側で記憶する
    CDPHelpers.InitQueryPerformanceFrequency
    '------------------------------------------------------------------------------
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    CDPPipeStream.Close
    Set CDPPipeStream = Nothing ' 役目を終えたので、Windowsにメモリを返してあげる
End Sub



#If DEBUG_MODE Then
'***************************************************************************************************
'                                 ■■■ Test Only Area ■■■
'***************************************************************************************************
' 単体テストをしやすくするための専用テストメソッド。
' 事前に「条件付きコンパイル引数」にて、`DEBUG_MODE = 1`をすることで有効化される
'***************************************************************************************************
'* 機能　　：`StrBufferCheck`のバッファーチェック
'---------------------------------------------------------------------------------------------------
'* 返り値  ：有効サイズ
'* 引数　　：仮想resSize
'---------------------------------------------------------------------------------------------------
'* 前提条件：`CreatePipe`の第四引数と辻褄を合わせる仕組みにしてます
'***************************************************************************************************
Public Function Test_StrBufferCheck(Optional resSize As Long = InitialBuffer + dwAllocationGranularity) As Long
    Const FromProcedureName As String = ThisClassName & ".Test_StrBufferCheck"


    'テスト前提条件を満たしてるかチェック
    '・負になっていないこと
    '・`CreatePipe`の第四引数のバッファー値(メモリアロケーション粒度込み)が上限圏内になってること
    If resSize < 0 Or resSize > InitialBuffer + dwAllocationGranularity Then Err.Raise vbObjectError, , "テスト前提条件を満たしてません"

    '渡してみる
    StrBufferCheck resSize

    '現在の有効エリアカーソルを返す
    printMsg info_, "`Len(responseCDP.strBuffer)`: " & Len(responseCDP.strBuffer), FromProcedureName
    printMsg info_, "`responseCDP.length`        : " & responseCDP.length, FromProcedureName
    Test_StrBufferCheck = responseCDP.EndCursor
End Function

Public Sub Test_SendEvent()
    'テスト用の結果を格納
    RaiseEvent CDPContextID(1000, "これは、タブに対する結果です", "")
    RaiseEvent CDPBrowserID(1001, "これは、ブラウザに対する結果です")
End Sub

Public Function Test_RunCountUPcommandID() As Long
    Test_RunCountUPcommandID = IssuanceCommandID
End Function
#End If



'***************************************************************************************************
'                              ■■■ HISTORICAL ARCHIVE ■■■
'***************************************************************************************************
' かつて、VBAのJSONパース制限と戦うため、極限の位置計算によって
' 巨大データから必要項目を1ミリ秒で撃ち抜いていた伝説のロマンコード。
' ゼロコピーパーサー(vbacollective/json)の導入に伴い、ここに栄誉を称え封印する。
' 2026/06/28 退役
'
' 消すのがなんかもったいないので、残しておきます。※`Github`に残るだろ、というクレームは受け付けません
'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`id`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：JsonString         CDPのレスポンス(Json文字列そのまんま)
'            idValueContents    コマンド結果のID値を格納する変数(参照渡しによって返します)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`{"id":27,"result":{"result":...`という感じで返るので、`InStr`による探索で抜き出します
'***************************************************************************************************
'Private Function GetContentsID(jsonString As String, ByRef idValueContents As Long) As Boolean
'    Const FromProcedureName As String = ThisClassName & ".GetContentsID"
'
'
'    '検索設定
'    Const SearchKeyName     As String = "id"    '検索文字列
'    Const StartSearchPos    As Long = 3         'あるべき位置
'
'
'    '1. 文字列を探索
'    Dim StartPos As Long
'    StartPos = InStr(StartSearchPos, jsonString, SearchKeyName)
'
'    '2. 所定の位置にない場合は、存在しないとし、抜ける
'    '下記の例に基づき、3文字目にない場合は、コマンド結果ではないとして、扱います。
'    '・{"id":27,"result":{"result":...
'    '・{"id":57,"error":{"code":...
'    If StartPos = StartSearchPos Then
'        GetContentsID = True
'
'        '3. 開始位置をうまい具合にずらす
'        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""":") '「id":」分、ずらすイメージ
'
'        '4. ID値の終点を求める
'        Dim EndPos As Long
'        EndPos = InStr(StartPos, jsonString, ",")
'
'        '5. ID値を取り出す
'        idValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
'        printMsg Trace_, "This response is the command result. ID:" & idValueContents, FromProcedureName
'    Else
'        GetContentsID = False
'        printMsg Trace_, "This response does not have an `id`.", FromProcedureName
'    End If
'End Function
'
''***************************************************************************************************
''* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`method`チェックを行います
''---------------------------------------------------------------------------------------------------
''* 返り値  ：存在有無論理値
''* 引数　　：JsonString             CDPのレスポンス(Json文字列そのまんま)
''            methodValueContents    非同期イベントのメソッド名(参照渡しによって返します)
''---------------------------------------------------------------------------------------------------
''* 詳細説明：`{"method":"Network.dataReceived","params":{"...`という感じで返るので、`InStr`による探索で抜き出します
''***************************************************************************************************
'Private Function GetContentsMethod(jsonString As String, ByRef methodValueContents As String) As Boolean
'    Const FromProcedureName As String = ThisClassName & ".GetContentsMethod"
'
'
'    '検索設定
'    Const SearchKeyName     As String = "method"    '検索文字列
'    Const StartSearchPos    As Long = 3             'あるべき位置
'
'
'    '1. 文字列を探索
'    Dim StartPos As Long
'    StartPos = InStr(StartSearchPos, jsonString, SearchKeyName)
'
'    '2. 所定の位置にない場合は、存在しないとし、抜ける
'    '下記の例に基づき、3文字目にない場合は、非同期イベントではないとして、扱います。
'    '・{"method":"Network.dataReceived","params":{"...
'    If StartPos = StartSearchPos Then
'        GetContentsMethod = True
'
'        '3. 開始位置をうまい具合にずらす
'        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""":""") '「method":"」分、ずらすイメージ
'
'        '4. method名の終点を求める
'        Dim EndPos As Long
'        EndPos = InStr(StartPos, jsonString, """,")
'
'        '5. method名を取り出す
'        methodValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
'        printMsg Trace_, "This response is an asynchronous event: " & methodValueContents, FromProcedureName
'    Else
'        GetContentsMethod = False
'        printMsg Trace_, "This response does not have an `method`.", FromProcedureName
'    End If
'End Function
'
''***************************************************************************************************
''* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`sessionId`チェックを行います
''---------------------------------------------------------------------------------------------------
''* 返り値  ：存在有無論理値
''* 引数　　：JsonString         CDPのレスポンス(Json文字列そのまんま)
''            idValueContents    sessionId値を格納する変数(参照渡しによって返します)
''---------------------------------------------------------------------------------------------------
''* 詳細説明：`...},"sessionId":"0DB67B7FD7C79501E6FBA9752A68F35A"}`という感じで返るので、後方から抜き出します
''***************************************************************************************************
'Private Function GetContentsSessionId(jsonString As String, ByRef sessionIdValueContents As String) As Boolean
'    Const FromProcedureName As String = ThisClassName & ".GetContentsSessionId"
'
'
'    '検索設定
'    Const SearchKeyName     As String = "sessionId" '検索文字列
'    Const LenSessionID      As Long = 32            '現時点での、`sessionId`文字数
'
'    '検索開始位置を特定
'    Dim StartSearchPos As Long
'    StartSearchPos = Len(jsonString) - Len(SearchKeyName) - Len(""":""") - LenSessionID - Len("""}")   '「"sessionId":"0DB67B7FD7C79501E6FBA9752A68F35A"}」分、後ろに設定するイメージ
'
'    '負になってたら即抜け
'    If StartSearchPos < 0 Then
'        GetContentsSessionId = False
'        printMsg Trace_, "This response is the short command result for the browser.", FromProcedureName
'        Exit Function
'    End If
'
'    '1. 文字列検索
'    Dim StartPos As Long
'    StartPos = InStr(StartSearchPos, jsonString, Chr(34) & SearchKeyName & """:""")   '「"sessionId":"」で検索をかけるようにする
'
'    '2. 所定の位置にない場合は、存在しないとし、抜ける
'    If StartPos = StartSearchPos Then
'        GetContentsSessionId = True
'
'        '3. sessionId値を取り出す
'        sessionIdValueContents = Mid(jsonString, Len(jsonString) - LenSessionID - Len("""}") + 1, LenSessionID)  '「0DB67B7FD7C79501E6FBA9752A68F35A"}」分、後ろに設定するイメージ
'        printMsg Trace_, "This response is the command result for the tab. sessionId:" & sessionIdValueContents, FromProcedureName
'    Else
'        GetContentsSessionId = False
'        printMsg Trace_, "This response is the command result for the browser.", FromProcedureName
'    End If
'End Function
