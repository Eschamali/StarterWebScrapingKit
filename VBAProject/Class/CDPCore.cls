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

'----- 待機関連 -----
Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long    'タイマー用
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long         '周波数取得用



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
'ウィンドウ表示設定値一式
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

'その他
Private Const ThisClassName     As String = "CDPCore"
Private Const InitialAllocationStrBuffer As Long = 2 ^ 8    '初期バッファー上限
Private Const LimitCommandID    As Long = 2000000000        'CDPコマンド送信時のID上限値



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
'クラスオブジェクト
Private Utf8Converter  As CharacterCodeConversion  'UTF-8を使用した日本語通信を有効化するやつ
Private LogControl     As New Logger               'ログレベルの制御

'定義側からも呼び出せるチェーンクラスオブジェクト
Public ErrorDetail     As New WinApiError          'エラーコードから、詳細を取得するやつ

 'バッファー管理(パイプ通信の仕様による欠けデータの蓄積管理)
Private Type StateFragmented
    strBuffer()     As String   'ストリーム受信で欠けたデータを一時的に蓄積するための配列です
    Index           As Long     '埋まってる位置
    expansionCount  As Long     '拡張確保回数
End Type
Private FragmentedDate As StateFragmented

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

'パイプ情報
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
Private LastCommandID_Pipe  As Long     '現在のコマンドID
Private m_Frequency         As Currency '実行マシンでの周波数記録用



'***************************************************************************************************
'                             ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：Chromiumを`remote-debugging-pipe`として起動し、パイプ接続を開始します
'---------------------------------------------------------------------------------------------------
'* 返り値　：実行結果　※0で成功です
'* 引数　　：strExec            コマンドライン
'            ShowWindowMode     初期ウィンドウ表示タイプ
'***************************************************************************************************
Public Function Init(strExec As String, ShowWindowMode As WinState) As Long
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
        Init = Err.LastDllError     'エラーコードで返す
        Exit Function
    End If

    'stderr
    result = CreatePipe(hStderrOutRd, hStderrOutWr, sa, 0)
    If result = 0 Then
        Init = Err.LastDllError     'エラーコードで返す
        Exit Function
    End If

   'Then stdin
    result = CreatePipe(hStdInRd, hStdInWr, sa, 0)
    If result = 0 Then
        Init = Err.LastDllError     'エラーコードで返す
        Exit Function
    End If

   'Then the out pipe for the CDP Protocol
    result = CreatePipe(hCDPOutRd, hCDPOutWr, sa, 2 ^ 20)   '※下限は、2 ^ 16
    If result = 0 Then
        Init = Err.LastDllError     'エラーコードで返す
        Exit Function
    End If

   'And finally the in pipe
    result = CreatePipe(hCDPInRd, hCDPInWr, sa, 2 ^ 20)     '※下限は、2 ^ 16
    If result = 0 Then
        Init = Err.LastDllError     'エラーコードで返す
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
        .wShowWindow = ShowWindowMode
        .cbReserved2 = LenB(pipes2)
        .lpReserved2 = VarPtr(pipes2)
    End With

    result = CreateProcess(0&, StrPtr(strExec), sa, sa, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    '実行結果チェック
    If result = 0 Then
        'エラーコードを返す
        Init = Err.LastDllError

        '失敗なので、自身で使用するはずだったハンドルも削除
        ClosePipeCDP
    Else
        '成功は、0として返却
        Init = 0
    End If

   'We close the sides of the handles that we dont need anymore
    CloseHandle hStdOutWr
    CloseHandle hStdInRd
    CloseHandle hCDPOutWr
    CloseHandle hCDPInRd
    CloseHandle hStderrOutWr

    hProcess = proc.hProcess
    dwProcessId = proc.dwProcessId

End Function



'***************************************************************************************************
'                              ■■■ Pipe関連操作 ■■■
'***************************************************************************************************
'* 機能　　：パイプハンドル情報をExcelのテーブルに記録します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   記録する識別名称
'***************************************************************************************************
Public Sub serialize(UserName As String)
    '0. データ記録準備
    Dim tmp As New Dictionary

    '1. 必要なデータを`Dictionary`に詰める
    tmp.Add "hStdOutRd", hStdOutRd
    tmp.Add "hStderrOutRd", hStderrOutRd
    tmp.Add "hStdInWr", hStdInWr
    tmp.Add "hCDPOutRd", hCDPOutRd
    tmp.Add "hCDPInWr", hCDPInWr
    tmp.Add "dwProcessId", dwProcessId

    '2. Excelのテーブルへ記録する
    Set ShSetting01_StartBrowser.serializeToTable(UserName, 1, "CDPCore.serialize") = tmp

End Sub

'***************************************************************************************************
'* 機能　　：Excelのテーブルに記録されてるパイプハンドル情報を読み込みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   読み取る識別名称
'***************************************************************************************************
Public Function deserialize(UserName As String) As Boolean
    'Excelテーブルから、`UserName`に紐づくパイプハンドル情報を読み込む
    Dim PipeInfo As Dictionary
    Set PipeInfo = ShSetting01_StartBrowser.deserializeFromTable(UserName, 1, "CDPCore.deserialize")

    '取得に失敗した場合は、Excelのテーブルに記録されてないので抜ける
    If PipeInfo Is Nothing Then Exit Function

    'クラス内変数に適用
    hStdOutRd = PipeInfo("hStdOutRd")
    hStderrOutRd = PipeInfo("hStderrOutRd")
    hStdInWr = PipeInfo("hStdInWr")
    hCDPOutRd = PipeInfo("hCDPOutRd")
    hCDPInWr = PipeInfo("hCDPInWr")
    dwProcessId = PipeInfo("dwProcessId")

    '成功サインを返す
    deserialize = True
End Function

'***************************************************************************************************
'* 機能　　：パイプハンドル情報をクリーニングします
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に、`deserialize`をした状態にすること
'***************************************************************************************************
Public Sub ClosePipeCDP()
    CloseHandle hStdOutRd
    CloseHandle hStdInWr
    CloseHandle hCDPOutRd
    CloseHandle hCDPInWr
    CloseHandle hStderrOutRd
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
    Const DividerMark       As String = vbCrLf  '標準出力、標準エラー出力は、改行コード区切りであることを確認


    Dim strRes As String, intRes As Long

    'Sometimes edge writes to stdout so we clear stdout too.
    intRes = readProcSTD(strRes, False)
    If intRes > 0 Then printMsg Debug_, LogID, "stdout = " & Replace(strRes, DividerMark, Chr(30)), FromProcedureName     '改行区切りだと、ログが見ずらくなるので専用の制御文字として置換

    'Sometimes edge writes to stderrout so we clear stderrout too.
    intRes = readProcSTD(strRes, True)
    If intRes > 0 Then printMsg ERROR_, LogID, "stderrout = " & Replace(strRes, DividerMark, Chr(30)), FromProcedureName  '改行区切りだと、ログが見ずらくなるので専用の制御文字として置換
End Sub

'***************************************************************************************************
'* 機能　　：CDP-Jsonコマンド結果や、ブラウザからのイベント情報の読み取りを試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：読み取ったバイト数を返します。
'* 引数　　：strData            CDP-Json の実行結果、ブラウザからのイベント情報を格納したい変数。実質返り値です
'---------------------------------------------------------------------------------------------------
'* 機能説明：読み取りは非ブロッキングです。データがなくても、次に進めれます。各種エラーコードは、`Get property`で確認してください。
'* 注意事項：`strData`を初期化せずに、使いまわす場合、次のような仕様になります
'            ・`PeekNamedPipe`にて、受信データ量が0あるいは、エラー：最後に受け取った`ReadFile`の内容を保持します
'            ・`PeekNamedPipe`にて、受信データ量が1以上　　　　　　：その受信した`ReadFile`の内容に置き換わります
'***************************************************************************************************
Private Function readProcCDP(ByRef strData As String) As Long
    Dim lPeekData As Long
    Dim lngRes As Long

    '処理が成功すると、戻り値は 0 以外になります。
    lngRes = PeekNamedPipe(hCDPOutRd, ByVal 0&, 0&, ByVal 0&, lPeekData, ByVal 0&)

    'API関数:PeekNamedPipe が失敗した場合は、0 を返すため、エラーコードを記録
    If lngRes = 0 Then LastPipeError.PeekNamedPipe = Err.LastDllError: Exit Function

    'データがあったら実際に、データを入手する処理へ
    If lPeekData > 0 Then
        Dim Buffer() As Byte
        Dim bSuccess As Long
        ReDim Buffer(lPeekData - 1)
        bSuccess = ReadFile(hCDPOutRd, ByVal VarPtr(Buffer(0)), lPeekData, 0&, 0&)

        '0以外なら成功です
        If bSuccess Then
            'バイト配列をUTF-8に変換するか？
            If Utf8Converter Is Nothing Then
                'そのまんま返却
                strData = StrConv(Buffer, vbUnicode)
            Else
                '変換した上で、返却
                strData = Utf8Converter.BytesToString(Buffer)
            End If

            '取得サイズを返却
            readProcCDP = lPeekData
        Else
            'API関数:ReadFile が失敗した場合は、0 を返すため、エラーコードを記録
            LastPipeError.ReadFile = Err.LastDllError
        End If
    End If
End Function

'***************************************************************************************************
'* 機能　　：実際にブラウザへ命令を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：strData    完成したCDP-Jsonコマンド
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

    '失敗時は、実際のエラーコードを記録
    If ResultCode = 0 Then LastPipeError.WriteFile = Err.LastDllError
End Sub



'***************************************************************************************************
'                               ■■■ CDPコマンド実行前の下準備 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザへ命令する前の組み立てを行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：`WriteFile`の実行結果  ※0で成功です
'* 引数　　：strMessage CDP-Jsonコマンド
'            SessionID  タブへの命令の場合は指定します
'---------------------------------------------------------------------------------------------------
'* 機能説明：`id`と`sessionId`の組み立てと紐づけ処理を行います
'***************************************************************************************************
Public Sub ReadyRunCDP(strMessage As String, Optional sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".ReadyRunCDP"


    '1. 現在のコマンドIDを発行  ※内部でカウントアップを行います
    Dim currentCommandID As Long: currentCommandID = IssuanceCommandID

    '2. 末尾付与しやすいように、末尾1文字削減
    strMessage = Left(strMessage, Len(strMessage) - 1)

    '3. `sessionId`check
    If StrPtr(sessionID) Then
        '3-1. 末尾に`sessionId`付与を行う
        strMessage = strMessage & ", ""sessionId"":""" & sessionID & """"

        '3-2. commandIDと`sessionId`の紐づけを行う
        DictionarySessionID.Add currentCommandID, sessionID
        printMsg Trace_, LogID, "commandID:" & currentCommandID & " corresponds to SessionID:" & sessionID, FromProcedureName
    Else
        printMsg Trace_, LogID, "commandID:" & currentCommandID & " does not have a SessionID.", FromProcedureName
    End If

    '4. 末尾に`id`,ヌル文字を付与
    strMessage = strMessage & ", ""id"":" & currentCommandID & "}" & vbNullChar

    '5. 実際に実行
    printMsg Debug_, LogID, "sendCDP - JsonCommand = " & strMessage, FromProcedureName
    writeProcCDP strMessage
End Sub

'***************************************************************************************************
'* 機能　　：CDP実行の際のコマンドIDを発行します
'---------------------------------------------------------------------------------------------------
'* 返り値　：発行されたcommandID
'---------------------------------------------------------------------------------------------------
'* 注意事項：20億を超えると、1にリセットされます
'***************************************************************************************************
Private Function IssuanceCommandID() As Long
    Const FromProcedureName As String = ThisClassName & ".IssuanceCommandID"


    If LastCommandID_Pipe + 1 > LimitCommandID Then
        '1. リクエスト数の上限に達したら、1にリセットしておく
        LastCommandID_Pipe = 1
        printMsg info_, LogID, "Reset to `1` to prevent internal counter overflow.", FromProcedureName

        '2. リセットした旨を伝える
        RaiseEvent ResetCommandID
    Else
        '1. カウントアップ
        LastCommandID_Pipe = LastCommandID_Pipe + 1   'We increase the global ID counter
    End If

    '2. 返却
    IssuanceCommandID = LastCommandID_Pipe
End Function



'***************************************************************************************************
'                               ■■■ 断片データの結合処理 ■■■
'***************************************************************************************************
'* 機能　　：`ReadFile`から得たデータが欠けた際のバッファー管理を担うヘルパープライベートプロパティです
'---------------------------------------------------------------------------------------------------
'* LET引数 ：pieceData  欠けたデータ
'* GET返値 ：集めた`pieceData`を結合した物
'---------------------------------------------------------------------------------------------------
'* 機能説明：断片データ（pieceData）を高速結合するための専用バッファプロパティ。
'  　　　　　【Let】データを配列の空き部屋（Index）に順次格納。（溢れた時のみバッファ拡張）
'  　　　　　【Get】格納された全データを Join し、vbNullChar を目印に1件分の完全なJSONを抽出。
'  　　　　　　　　 抽出後は Index を 0 に戻すことで、実質的なメモリクリアを負荷ゼロで行います。
'
'* 注意事項：・`pieceData`に入れるのは必ず欠けたデータであること
'            ・GET時、`vbnullchar`がどこかに存在するとし、先頭から1件分取り出せる目印として機能していること
'***************************************************************************************************
Private Property Let FragmentedCDP(pieceData As String)
    Const FromProcedureName As String = ThisClassName & ".FragmentedCDP - Let"


    With FragmentedDate
        '1. 上限を越したら、拡張する(高負荷ポイント)
        If .Index >= InitialAllocationStrBuffer * .expansionCount Then
            .expansionCount = .expansionCount + 1
            ReDim Preserve .strBuffer(InitialAllocationStrBuffer * .expansionCount - 1)
            printMsg Debug_, LogID, "The buffer has been expanded to " & InitialAllocationStrBuffer * .expansionCount, FromProcedureName
        End If

        '2. 配列に入れる
        .strBuffer(.Index) = pieceData
        printMsg Trace_, LogID, "Data has been placed in array " & .Index, FromProcedureName

        '3. カウント
        .Index = .Index + 1
    End With
End Property

Private Property Get FragmentedCDP() As String
    Const FromProcedureName As String = ThisClassName & ".FragmentedCDP - Get"


    '1. 確保中の配列を全部結合する
    printMsg Trace_, LogID, "Now we'll begin the buffer merging process...", FromProcedureName
    Dim tmp As String: tmp = Join(FragmentedDate.strBuffer, "")
    printMsg Trace_, LogID, "concatenated Raw-string: " & tmp, FromProcedureName

    '2. `vbnullchar`がある位置を先頭から探す
    Dim EndPos As Long: EndPos = searchNull(tmp)

    '3. その目印を元に、1件分返却
    FragmentedCDP = Mid(tmp, 1, EndPos - 1)
    printMsg Debug_, LogID, "Trim complete. Result: " & FragmentedCDP, FromProcedureName

    '4. インデックスリセット
    FragmentedDate.Index = 0
    printMsg Trace_, LogID, "Buffer Index reset.", FromProcedureName

End Property

Private Function searchNull(checkString As String) As Long
'------------------------------------------------------------
' CDP messages received from chrome are null-terminated
' Updated: 25/10/25: Daniel Polak - new faster version
'------------------------------------------------------------

    Dim lngPos As Long

    If Len(checkString) = 0 Then Exit Function
    lngPos = InStr(1, checkString, vbNullChar, vbBinaryCompare)
    If lngPos > 0 Then
        searchNull = lngPos
    Else
        searchNull = 0
    End If

End Function



'***************************************************************************************************
'                               ■■■ レスポンス捌き ■■■
'***************************************************************************************************
'* 機能　　：パイプから、データ欠損がなくなるまで、ループして取り出していきます
'---------------------------------------------------------------------------------------------------
'* 引数　　：StopPipeError  `True` ：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあったら停止します
'                           `False`：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあっても停止しません
'                                    別途、`LastError○○`で後で確認可能です
'
'            destruction    `True` :IDの解析やイベントの中身の解析を行わず、破棄します。解析しない分、多少パフォーマンスが向上します
'                           `False`:IDの解析やイベントの中身の解析を行います
'
'***************************************************************************************************
Public Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakeEvents"


    Dim intRes As Long
    Dim strRes As String
    printMsg Trace_, LogID, "Starts retrieving data from the pipe...", FromProcedureName

    '終端null文字があって尚且つ、受信データ0が来るまで、ループ受信を行います
    Dim vbNullCharTermination As Boolean: vbNullCharTermination = True  '初期化
    Do
        'std系の中身を出力
        viewStdoutAndStderr

        '受信処理
        intRes = readProcCDP(strRes)
        printMsg Trace_, LogID, "ResSize = " & intRes, FromProcedureName

        '返り値チェック
        If intRes > 0 Then
            '破棄モードかの判定
            If destruction Then
                '解析などせず、そのまんま出して破棄します
                '※強いていうなら、破棄するメッセージをそのまま出す
                printMsg Trace_, LogID, "Discarding data in pipe...", FromProcedureName
                If currentLog.seeRawSendMsgDbg Then printMsg Trace_, LogID, "Discarded - strRes = " & strRes, FromProcedureName

                '欠けチェック
                vbNullCharTermination = (InStrRev(strRes, vbNullChar) = Len(strRes))
            Else
                '中身の解析をし、IDや非同期イベントを取り出します
                printMsg Trace_, LogID, "Analyzing the data in the pipe...", FromProcedureName
                Call ReceiveFromPipeSplit(strRes)

                '欠けチェック
                vbNullCharTermination = (FragmentedDate.Index = 0)
            End If
        Else
            'データなしの場合は、エラーコードチェックを行う
            If LastPipeError.PeekNamedPipe Then
                'PeekNamedPipe にてエラーがあった場合
                printMsg ERROR_, LogID, "Error PeekNamedPipe in readProcCDP. ErrorCode:" & LastPipeError.PeekNamedPipe & " - " & ErrorDetail.GetMessage(LastPipeError.PeekNamedPipe, "kernel32"), FromProcedureName, doRaiseError:=StopPipeError
                Exit Sub

            ElseIf LastPipeError.ReadFile Then
                'ReadFile にてエラーがあった場合
                printMsg ERROR_, LogID, "Error ReadFile in readProcCDP. ErrorCode:" & LastPipeError.ReadFile & " - " & ErrorDetail.GetMessage(LastPipeError.ReadFile, "kernel32"), FromProcedureName, doRaiseError:=StopPipeError
                Exit Sub

            Else
                'パイプ内は空
                printMsg Trace_, LogID, "The message queue is currently empty.", FromProcedureName

            End If
        End If

        '応答なし防止
        DoEvents
    Loop Until vbNullCharTermination And intRes < 1

    printMsg Trace_, LogID, "Finished getting data from the pipe.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：1度の`ReadFile`で取り出した長文を、1件1件に分割し、必要に応じて、断片データを1件分に復元するようにします
'---------------------------------------------------------------------------------------------------
'* 引数　　：ReceiveRawString       パイプから得たそのまんまの文字列
'***************************************************************************************************
Private Function ReceiveFromPipeSplit(ReceiveRawString As String)
    Const FromProcedureName As String = ThisClassName & ".ReceiveFromPipeSplit"
    Const DividerMark       As String = vbNullChar  'CDPの受信データは、null文字区切りであることを確認


    '1. 受信メッセージを活用するため、null文字区切りとしてスマートに分割
    Dim JsonMessages: JsonMessages = Split(ReceiveRawString, DividerMark)

    '2. 各要素数を確認
    Dim FirstIndex  As Long: FirstIndex = LBound(JsonMessages)  'TIP：Split後の配列は、0始まりIndexです
    Dim LastIndex   As Long:  LastIndex = UBound(JsonMessages)

    '3. いくつかログ化
    '3-1. `ReadFile`で得た内容を確認
    If currentLog.seeRawSendMsgDbg Then printMsg Trace_, LogID, "Raw - ReadFileRes = " & ReceiveRawString, FromProcedureName

    '3-2. 最初と最後の中身を確認
    If FirstIndex = LastIndex Then
        printMsg Trace_, LogID, "SingleResCDP  - contents: " & JsonMessages(FirstIndex), FromProcedureName
    Else
        printMsg Trace_, LogID, "FirstIndexCDP - contents: " & JsonMessages(FirstIndex), FromProcedureName
        printMsg Trace_, LogID, "LastIndexCDP  - contents: " & JsonMessages(LastIndex), FromProcedureName
    End If

    '------------------------- 4. 最初の要素に対して特別な処理を施す -------------------------
    '※最初と最後の要素数が同じの場合は、欠けてる状態なのでパス
    If FirstIndex < LastIndex Then
        '4-1. 欠けたデータがあるか？
        If FragmentedDate.Index > 0 Then
            '4-2. 空データか？
            If JsonMessages(FirstIndex) = "" Then
                '4-2-1. 終端目印を付与
                FragmentedCDP = DividerMark
                printMsg Trace_, LogID, "Input data has arrived, but since it is an empty string, the currently stored data will be treated as JSON data.", FromProcedureName
            Else
                '4-2-1. 最後の断片データを挿入し、終端目印も付与
                FragmentedCDP = JsonMessages(FirstIndex)
                FragmentedCDP = DividerMark
                printMsg Trace_, LogID, "The merge completes the JSON data.", FromProcedureName
            End If

            '4-3. 集まったデータを取り出し、最初の要素に上書きし、以降のFor文で活用できるようにする
            '※内部で、Jsonデータが完成されたとし、蓄積中のデータも実質クリアされます
            JsonMessages(FirstIndex) = FragmentedCDP
        Else
            printMsg Trace_, LogID, "There is no missing data being stored.", FromProcedureName
        End If
    End If
    '-----------------------------------------------------------------------------------------

    '------------------------- 5. 最後の要素に対して特別な処理を施す -------------------------
    '5-1. 空欄か？
    If JsonMessages(LastIndex) = "" Then
        '5-2. 今回の受信で全て正しいJson形式として得れました
        printMsg Debug_, LogID, LastIndex & " messages received.", FromProcedureName
        printMsg Debug_, LogID, "readProcCDP was able to receive everything!", FromProcedureName
    Else
        '5-2. データが、"欠けた"状態なので蓄積で、キャッシュしておく。頻繁には来ない想定
        printMsg Debug_, LogID, LastIndex + 1 & " messages received.", FromProcedureName
        printMsg Trace_, LogID, "The Split(" & LastIndex & ") data is missing JSON data. Accumulate until completion.", FromProcedureName
        FragmentedCDP = JsonMessages(LastIndex)

        printMsg Debug_, LogID, "readProcCDP was unable to receive all the data. It accumulates the last element data...", FromProcedureName
    End If
    '-----------------------------------------------------------------------------------------

    '6. 最後の要素以外をループで回して、ブラウザからのメッセージを確認する
    Dim i As Long
    For i = LBound(JsonMessages) To UBound(JsonMessages) - 1
        'レスの確認
        If currentLog.seeRawSendMsgDbg Then printMsg Debug_, LogID, "Split(" & i & ") - strRes = " & JsonMessages(i), FromProcedureName 'Only enable if needs to see every raw JS message for debugging/development

        '受信データのチェックを行うプロシージャへ遷移します。
        Call BrowserReceivedDataCheck(CStr(JsonMessages(i)))
    Next
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックし、各クラスへイベント発信をします
'---------------------------------------------------------------------------------------------------
'* 引数　　：CDPJsonString          1件分のJson形式文字列のブラウザからのイベントやCDP-Jsonコマンド結果
'---------------------------------------------------------------------------------------------------
'* 詳細説明：イベントデータ or CDP-Jsonコマンド結果 or 構文エラー の確認をし、各種`RaiseEvent`で発信します
'* 注意事項；構文エラーの場合、不整合により不安定につながるため、エラーで停止します
'***************************************************************************************************
Private Sub BrowserReceivedDataCheck(CDPJsonString As String)
    Const FromProcedureName As String = ThisClassName & ".BrowserReceivedDataCheck"


    Dim ResultCommandID As Long, AsynchronousEventName As String, sessionID As String

    '--------- 1. ブラウザからのイベントか？ ---------
    If GetContentsMethod(CDPJsonString, AsynchronousEventName) Then

        '--------- 2. タブに対する非同期イベントか？ ---------
        If GetContentsSessionId(CDPJsonString, sessionID) Then
            'タブに対する非同期イベントとして発信
            RaiseEvent CDPContextEvent(AsynchronousEventName, CDPJsonString, sessionID)
        Else
            'ブラウザに対する非同期イベントとして発信
            RaiseEvent CDPBrowserEvent(AsynchronousEventName, CDPJsonString)
        End If
    Else
        '--------- 2. `id`項目があるか？ ---------
        If GetContentsID(CDPJsonString, ResultCommandID) Then
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
            printMsg ERROR_, LogID, "A syntax error was detected when executing with ExecuteCDP/ExecuteCDPAsync. Processing was interrupted to prevent inconsistencies." & vbCrLf & "> " & CDPJsonString, FromProcedureName, doRaiseError:=True
        End If
    End If
End Sub

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`id`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：JsonString         CDPのレスポンス(Json文字列そのまんま)
'            idValueContents    コマンド結果のID値を格納する変数(参照渡しによって返します)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`{"id":27,"result":{"result":...`という感じで返るので、`InStr`による探索で抜き出します
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
        printMsg Trace_, LogID, "This response is the command result. ID:" & idValueContents, FromProcedureName
    Else
        GetContentsID = False
        printMsg Trace_, LogID, "This response does not have an `id`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`method`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：JsonString             CDPのレスポンス(Json文字列そのまんま)
'            methodValueContents    非同期イベントのメソッド名(参照渡しによって返します)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`{"method":"Network.dataReceived","params":{"...`という感じで返るので、`InStr`による探索で抜き出します
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
        printMsg Trace_, LogID, "This response is an asynchronous event: " & methodValueContents, FromProcedureName
    Else
        GetContentsMethod = False
        printMsg Trace_, LogID, "This response does not have an `method`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`sessionId`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：JsonString         CDPのレスポンス(Json文字列そのまんま)
'            idValueContents    sessionId値を格納する変数(参照渡しによって返します)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`...},"sessionId":"0DB67B7FD7C79501E6FBA9752A68F35A"}`という感じで返るので、後方から抜き出します
'***************************************************************************************************
Private Function GetContentsSessionId(jsonString As String, ByRef sessionIdValueContents As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".GetContentsSessionId"


    '検索設定
    Const SearchKeyName     As String = "sessionId" '検索文字列
    Const LenSessionID      As Long = 32            '現時点での、`sessionId`文字数

    '検索開始位置を特定
    Dim StartSearchPos As Long
    StartSearchPos = Len(jsonString) - Len(SearchKeyName) - Len(""":""") - LenSessionID - Len("""}")   '「"sessionId":"0DB67B7FD7C79501E6FBA9752A68F35A"}」分、後ろに設定するイメージ

    '負になってたら即抜け
    If StartSearchPos < 0 Then
        GetContentsSessionId = False
        printMsg Trace_, LogID, "This response is the short command result for the browser.", FromProcedureName
        Exit Function
    End If

    '1. 文字列検索
    Dim StartPos As Long
    StartPos = InStr(StartSearchPos, jsonString, Chr(34) & SearchKeyName & """:""")   '「"sessionId":"」で検索をかけるようにする

    '2. 所定の位置にない場合は、存在しないとし、抜ける
    If StartPos = StartSearchPos Then
        GetContentsSessionId = True

        '3. sessionId値を取り出す
        sessionIdValueContents = Mid(jsonString, Len(jsonString) - LenSessionID - Len("""}") + 1, LenSessionID)  '「0DB67B7FD7C79501E6FBA9752A68F35A"}」分、後ろに設定するイメージ
        printMsg Trace_, LogID, "This response is the command result for the tab. sessionId:" & sessionIdValueContents, FromProcedureName
    Else
        GetContentsSessionId = False
        printMsg Trace_, LogID, "This response is the command result for the browser.", FromProcedureName
    End If
End Function



'***************************************************************************************************
'                               ■■■ プロパティ機能 ■■■
'***************************************************************************************************
'* 機能　　：コマンドを送る(WriteFile)前に、VBAの「Unicode」文字列をUTF-8でエンコードしたバイト配列として、コマンドを送るかを設定します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：UTF-8 の変換が作動してるかの論理値
'* LET引数 ：・`True`   VBAの「Unicode」文字列をUTF-8でエンコードしたバイト配列として、コマンドを送ります
'            ・`False`  VBAの「Unicode」文字列そのまんまで、コマンドを送ります
'---------------------------------------------------------------------------------------------------
'* 詳細説明：日本語での送信が必要の場合は`True`で設定しないと、途中で動作が止まることがあるため、このプロパティを設けます
'***************************************************************************************************
Property Let UseConverterUtf8FromString(Flag As Boolean)
    If Flag Then Set Utf8Converter = New CharacterCodeConversion Else Set Utf8Converter = Nothing
End Property

Property Get UseConverterUtf8FromString() As Boolean
    UseConverterUtf8FromString = Not (Utf8Converter Is Nothing)
End Property

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

Property Get ConnectionProcessId() As LongPtr           '`--remote-debugging-pipe`で起動した際のProcessID
    ConnectionProcessId = dwProcessId
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

Property Get LastCommandID() As Long                     '最後に実行したcommandID値
    LastCommandID = LastCommandID_Pipe
End Property

'***************************************************************************************************
'* 機能　　：ログ設定値、取得用プロパティ
'***************************************************************************************************
Property Get ShowLogCDPresponse() As Boolean
    ShowLogCDPresponse = currentLog.seeRawSendMsgDbg
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
        If Dir(currentLog.logPath, vbDirectory) = "" Then MkDir currentLog.logPath
    End If

   'Perform msg output to Immediate Window if needed
    With LogControl
        Select Case LogLevel_
            Case LogLevelName.Trace_: .LogTrace strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.Debug_: .LogDebug strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.info_: .LogInfo strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.WARN_: .LogWarn strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.ERROR_: .LogError strFormattedMsg, From, currentLog.logPath, logName, Err.Number
            Case Else: 'ログ出力無効化
        End Select
    End With


    '主に、CDPコマンドのエラー時に止まります
    If doRaiseError Then Err.Raise CDPCustomErrorCodes.Protocol, From, Description:=strMsg

End Sub

Public Sub sleep(Optional seconds As Double = 0.5)
'------------------------------------------------------------------
' Custom sleep function. Sleep by 0.5s by default.
' Useful for a quick necessary pause when needed.
'------------------------------------------------------------------

    Const baseUnit As Long = 1000    'ie. millisecs

    sleep2 seconds * baseUnit
    DoEvents

End Sub

'***************************************************************************************************
'* 機能　　：PC起動時からの「経過時間（ミリ秒）」を正確に返します（Timer関数の神互換）
'---------------------------------------------------------------------------------------------------
'* 返り値　：PC起動時からの「経過時間（ミリ秒）」
'***************************************************************************************************
Public Function GetTicksMs() As Double
    '1. 今の「振動回数」を取得する
    Dim currentCount As Currency
    Call QueryPerformanceCounter(currentCount)

    '2. 割り算して「秒」にし、1000倍して「ミリ秒（Double型）」にして返す！
    GetTicksMs = (CDbl(currentCount) / CDbl(m_Frequency)) * 1000#
End Function



'***************************************************************************************************
'                               ■■■ 初期化処理 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    '------------------------- クラス内変数の一部を初期化 -------------------------
    '1. Json分断時の一時キュー関連の初期設定
    With FragmentedDate
        .expansionCount = 1
        ReDim .strBuffer(InitialAllocationStrBuffer * .expansionCount - 1)
        .Index = 0
    End With

    '2. 設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        '2-1. イミディエイトに関する設定を行う
        currentLog.seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value           'ブラウザから受信した生のJSONメッセージもイミディエイトに表示させるか？

        '2-2. パスが設定されてたら、ログファイル保存を有効化
        Dim LogFolderPath As String: LogFolderPath = .Range(.UseRangeName(9, FromProcedureName)).value
        If LogFolderPath <> "" Then
            currentLog.doLog = True                'ログファイルとして保存するか？
            currentLog.logPath = LogFolderPath     '保存先のフォルダパスを設定
        End If

        '2-3. Assign an objID for easy debug identification
        LogID = "COID" & Format(Rnd * 1000, "000")
    End With

    '3. コマンドIDに紐づくSessionIDの管理を初期化
    Set DictionarySessionID = New Dictionary

    '4. 実行マシンの「1秒間の振動数（周波数）」を取得して記憶する
    Call QueryPerformanceFrequency(m_Frequency)
    '------------------------------------------------------------------------------
End Sub
