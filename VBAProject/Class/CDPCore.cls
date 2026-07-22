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

'その他
Private Const ThisClassName     As String = "CDPCore"
Private Const InitialBuffer     As Long = 2 ^ 20        'CDPやり取りPipe/テキスト変数/ADODB.Stream 初期バッファー上限
Private Const RunDoEventsCount  As Long = 2 ^ 10        '長いループ中に`DoEvents`を挟む回数
Private Const LimitCommandID    As Long = 2000000000    'CDPコマンド送信時のID上限値



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

'バッファー管理
Private Type ResponseBufferCDP
    strBuffer       As String   'ブラウザから送られたCDP情報をどんどん蓄積する変数
    length          As Long     '現在のバッファ全体の長さ
    StartCursor     As Long     '取り出しの開始位置
    AddCursor       As Long     '次に蓄積し始める位置
    EndCursor       As Long     '有効なデータ末尾位置
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
Private LastCommandID_Pipe  As Long     '現在のコマンドID
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
        .wShowWindow = ShowWindowMode
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
'* 引数　　：UserName   読み取る識別名称
'***************************************************************************************************
Public Function deserialize(UserName As String) As Boolean
    'Excelテーブルから、`UserName`に紐づくパイプハンドル情報を読み込む
    Dim PipeInfo As Dictionary
    Set PipeInfo = ShSetting01_StartBrowser.TableBrowserHandle(UserName, "CDPCore.deserialize")

    '取得に失敗した場合は、Excelのテーブルに記録されてないので抜ける
    If PipeInfo Is Nothing Then Exit Function

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
'* 機能　　：各種ハンドル情報をクリーニングします
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に、`deserialize`をした状態にすること
'***************************************************************************************************
Public Sub CloseHandleBrowser()
    'WebSocketモードならそっち側のClose処理を行う
    If viaWebSocket Is Nothing Then
        CloseHandle hStdOutRd
        CloseHandle hStdInWr
        CloseHandle hCDPOutRd
        CloseHandle hCDPInWr
        CloseHandle hStderrOutRd
        CloseHandle hProcess
    Else
        viaWebSocket.DisconnectCDP
    End If
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
        printMsg Debug_, LogID, "stdout = " & Replace(strRes, DividerMark, Chr(30)), FromProcedureName     '改行区切りだと、ログが見ずらくなるので専用の制御文字として置換
    End If

    'Sometimes edge writes to stderrout so we clear stderrout too.
    intRes = readProcSTD(strRes, True)
    If intRes > 0 Then
        If InStr(1, strRes, vbCrLf) Then strRes = Replace(strRes, vbCrLf, DividerMark)  'Windows用の改行コードがあったら一旦、Linux用改行コードに置換
        printMsg ERROR_, LogID, "stderrout = " & Replace(strRes, DividerMark, Chr(30)), FromProcedureName  '改行区切りだと、ログが見ずらくなるので専用の制御文字として置換
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
            'ストリームに流し込む ※バッファー拡張は、`ADODB.Stream`が担います
            If Not destruction Then CDPPipeStream.Write Buffer

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
    printMsg Trace_, LogID, "Storing received messages...", FromProcedureName

    'メッセージの終わりであれば、その目印も追記
    If EOS Then CDPPipeStream.Write m_NullTerminator

    '取得サイズをcursor管理へ反映
    StrBufferCheck resSize + EOS * -1
    printMsg Trace_, LogID, "ResSize = " & resSize & " , EOS: " & EOS, FromProcedureName
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
'                               ■■■ CDPコマンド実行前の下準備 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザへ命令する前の組み立てを行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：`WriteFile`の実行結果  ※0で成功です
'* 引数　　：dicMessage Chromium ブラウザに送るCDP-Jsonコマンド(Dictionary型)。最上位要素に、`method`、`params`があるようにしてください。
'            SessionID  タブへの命令の場合は指定します
'---------------------------------------------------------------------------------------------------
'* 機能説明：`id`と`sessionId`の組み立てと紐づけ処理を行います
'***************************************************************************************************
Public Sub ReadyRunCDP(dicMessage As Scripting.Dictionary, Optional sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".ReadyRunCDP"


    '1. 現在のコマンドIDを発行  ※内部でカウントアップを行います
    Dim currentCommandID As Long: currentCommandID = IssuanceCommandID

    '2. `sessionId`check
    If StrPtr(sessionID) Then
        '2-1. `sessionId`付与
        dicMessage.Add "sessionId", sessionID

        '2-2. commandIDと`sessionId`の紐づけを行う
        DictionarySessionID.Add currentCommandID, sessionID
        printMsg Trace_, LogID, "commandID:" & currentCommandID & " corresponds to SessionID:" & sessionID, FromProcedureName
    Else
        printMsg Trace_, LogID, "commandID:" & currentCommandID & " does not have a SessionID.", FromProcedureName
    End If

    '3.`id`を付与
    dicMessage.Add "id", currentCommandID

    '4. 実際に実行
    Dim strMessage As String
'    strMessage = BiDiCDPJson.StringifyValue(dicMessage)     '理想はこっちにして、1つに完結したいﾋﾟｴﾝ
    strMessage = WebJsonConverter.serialize(dicMessage)     'Dictionary→Json文字列化のみ`VBA-FastJSON`を使用。こっちのほうがパフォーマンスが良いため
    printMsg Debug_, LogID, "sendCDP - JsonCommand = " & strMessage, FromProcedureName

    '5. 通信モードに応じてPipe/WebSocket分岐
    If viaWebSocket Is Nothing Then writeProcCDP strMessage Else writeWebSocketCDP strMessage
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
'                         ■■■ CDPレスポンス、蓄積/取得管理処理 ■■■
'***************************************************************************************************
'* 機能　　：適切な文字列変数バッファーが確保されてるかチェックします
'---------------------------------------------------------------------------------------------------
'* 引数    ：ResSize    `ReadFile`から得たデータサイズそのまんま
'***************************************************************************************************
Private Sub StrBufferCheck(resSize As Long)
    Const FromProcedureName As String = ThisClassName & ".StrBufferCheck"


    '新規メッセージがない場合は即抜け
    If resSize <= 0 Then printMsg WARN_, LogID, "No message was found, so the buffer check will be skipped.", FromProcedureName: Exit Sub
    With responseCDP
        '1. 上限を越しそうになったら、拡張する(高負荷ポイント)
        '※オーバフローしたらこのツールでは対応できません
        If .EndCursor + resSize > .length Then
            .strBuffer = .strBuffer & String$(.length, vbNullChar)
            .length = .length * 2
            printMsg Debug_, LogID, "The buffer has been expanded: " & .length, FromProcedureName
        End If

        '2. 有効予定なCDPメッセージカーソル末尾位置を更新
        .EndCursor = .EndCursor + resSize

        '3. 新規メッセージフラグを付与
        .ExistsNewRes = True
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
        '1. 取り出すバイト数を確定
        .SetEOS

        '2. CDPメッセージを取り出すため、先頭に巻き戻す
        '※既に0の場合(メッセージが空)は即抜け
        If .Position Then .Position = 0 Else printMsg Debug_, LogID, "The stored messages are empty.", FromProcedureName: Exit Sub

        '3. UTF-16変数に格納する準備
        If Utf8Converter Is Nothing Then
            'そのまんまUTF-16変数に格納
            Mid$(responseCDP.strBuffer, responseCDP.AddCursor + 1) = StrConv((.Read), vbUnicode)
        Else
            'UTF-8に変換した上で、UTF-16変数に格納
            Mid$(responseCDP.strBuffer, responseCDP.AddCursor + 1) = Utf8Converter.BytesToString((.Read))
        End If

        '4. 次の蓄積位置を記録
        responseCDP.AddCursor = responseCDP.EndCursor

        '5 .追加出来たので、カーソル位置をリセットし、次の新着メッセージ蓄積に備える
        .Position = 0
    End With
End Sub

'***************************************************************************************************
'* 機能　　：呼び出す度に1件1件、CDPメッセージを取り出します
'***************************************************************************************************
Private Function TakeCDPMessage() As String
    Const FromProcedureName As String = ThisClassName & ".TakeCDPMessage"


    With responseCDP
        '1. 次に蓄積し始める位置が先頭の場合は即抜け
        If .AddCursor = 0 Then printMsg WARN_, LogID, "There are no new responses.", FromProcedureName: Exit Function

        '2. 今の位置から`vbnullchar`の位置を探索
        Dim EndPos As Long
        EndPos = searchNull(.strBuffer, .StartCursor + 1)

        '3. 見つからないあるいは、`.AddCursor`を超える場合は欠け判定とする
        If EndPos = 0 Or EndPos > .AddCursor Then printMsg WARN_, LogID, "There are new messages, but some are still missing.", FromProcedureName: Exit Function

        '4. その目印を元に、1件分返却
        TakeCDPMessage = Mid$(.strBuffer, .StartCursor + 1, EndPos - .StartCursor - 1)

        '5. 次の検索開始位置を更新
        If .AddCursor = EndPos Then
            '次に蓄積し始める位置にたどり着いちゃったら、カーソル情報をリセットしておく
            .EndCursor = .EndCursor - .AddCursor
            .AddCursor = 0
            .StartCursor = 0
            .ExistsNewRes = False

            printMsg Debug_, LogID, "All accumulated CDPmessages have been processed. Resetting cursor information.", FromProcedureName
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
'* 注意事項：不整合を防ぐため、終端ヌル文字が来るまで、受信を繰り返すようにしてください
'***************************************************************************************************
Private Sub TakePipe(StopPipeError As Boolean, destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakePipe"


    Dim intRes As Long
    printMsg Trace_, LogID, "Starts retrieving data from the pipe...", FromProcedureName

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
        intRes = readProcCDP(EOvbNullChar, destruction)
        printMsg Trace_, LogID, "ResSize = " & intRes, FromProcedureName

        '返り値チェック
        If intRes > 0 Then
            '破棄モードかの判定
            If destruction Then
                '解析などせず、そのまんま出して破棄します
                printMsg Trace_, LogID, "Discarding data in pipe...", FromProcedureName
            Else
                '中身の解析に備えて、蓄積準備
                StrBufferCheck intRes
                printMsg Trace_, LogID, "Storing received messages...", FromProcedureName
            End If
        Else
            'データなしの場合は、エラーコードチェックを行う
            If LastPipeError.PeekNamedPipe Then
                'PeekNamedPipe にてエラーがあった場合
                printMsg ERROR_, LogID, "Error PeekNamedPipe in readProcCDP. ErrorCode:" & LastPipeError.PeekNamedPipe & " - " & WinApiError.GetMessage(LastPipeError.PeekNamedPipe, "kernel32"), FromProcedureName, doRaiseError:=StopPipeError
                Exit Sub

            ElseIf LastPipeError.ReadFile Then
                'ReadFile にてエラーがあった場合
                printMsg ERROR_, LogID, "Error ReadFile in readProcCDP. ErrorCode:" & LastPipeError.ReadFile & " - " & WinApiError.GetMessage(LastPipeError.ReadFile, "kernel32"), FromProcedureName, doRaiseError:=StopPipeError
                Exit Sub

            Else
                'パイプ内は空
                printMsg Trace_, LogID, "The ReadFile is currently empty.", FromProcedureName

            End If
        End If
    Loop Until EOvbNullChar And intRes < 1
    printMsg Trace_, LogID, "Finished getting data from the pipe.", FromProcedureName

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
'***************************************************************************************************
Private Sub TakeWebSocket(StopWebSocketError As Boolean, destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakeWebSocket"


    Dim intRes As Long
    printMsg Trace_, LogID, "Starts retrieving data from the WebSocket...", FromProcedureName

    '終了サインかつ、受信データなしになるまで、ループ受信を行います
    Dim LoopCounter As Long
    Dim isFIN As Boolean: isFIN = True  '初期化
    Do
        '一定の回数ごとに`DoEvents`を呼ぶ。※初回ループは、必ず呼ぶ
        If LoopCounter Mod RunDoEventsCount = 0 Then DoEvents
        LoopCounter = LoopCounter + 1


        '受信処理
        intRes = viaWebSocket.GetAsyncMessage(isFIN, destruction)

        '返り値チェック
        If intRes > 0 Then
            '破棄モードかの判定
            If destruction Then
                '解析などせず、そのまんま出して破棄します
                printMsg Trace_, LogID, "Discarding data in WebSocket...", FromProcedureName
            Else
                '中身の解析に備えて、`WithEvents`側で蓄積します
                printMsg Trace_, LogID, "Storing received messages...", FromProcedureName
            End If
        Else
            'データなしの場合は、エラーコードチェックを行う
            If viaWebSocket.IsErrorStylePeekNamedPipe Then
                'ioctlsocket にてエラーがあった場合
                printMsg ERROR_, LogID, "Error ioctlsocket in GetAsyncMessage. ErrorCode:" & viaWebSocket.IsErrorStylePeekNamedPipe & " - " & WinApiError.GetMessage(viaWebSocket.IsErrorStylePeekNamedPipe, "ws2_32"), FromProcedureName, doRaiseError:=StopWebSocketError
                Exit Sub

            ElseIf viaWebSocket.IsErrorStyleReadFile Then
                'recv にてエラーがあった場合
                printMsg ERROR_, LogID, "Error recv in GetAsyncMessage. ErrorCode:" & viaWebSocket.IsErrorStyleReadFile & " - " & WinApiError.GetMessage(viaWebSocket.IsErrorStyleReadFile, "ws2_32"), FromProcedureName, doRaiseError:=StopWebSocketError
                Exit Sub

            ElseIf viaWebSocket.LastBufferType = WS_CLOSE_MESSAGE Then
                '切断要求が来た場合
                printMsg ERROR_, LogID, "Disconnected because a disconnection request was received from the Chromium browser.", FromProcedureName, doRaiseError:=StopWebSocketError
                Exit Sub
            Else
                'パイプ内は空あるいは、WebSocketフレーム単位としてまだ届いていない
                printMsg Trace_, LogID, "The WebSocket is currently empty.", FromProcedureName

            End If
        End If
    Loop Until isFIN And intRes < 1

End Sub



'***************************************************************************************************
'                               ■■■ レスポンス捌き ■■■
'***************************************************************************************************
'* 機能　　：メッセージ解析を行い、イベント発火まで行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：StopPipeError  `True` ：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあったら停止します
'                           `False`：`PeekNamedPipe` OR `ReadFile`いずれかに、エラーがあっても停止しません
'                                    別途、`LastError○○`で後で確認可能です
'
'            destruction    `True` :メッセージ解析をせずに溜まったデータの取り出しのみ行います
'                           `False`:ID、イベント解析を行いつつ、溜まったデータの取り出しも並行して行います
'***************************************************************************************************
Public Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakeEvents"


    '1. 通信モードに応じてPipe/WebSocket分岐し、受信メッセージの取り出し
    If viaWebSocket Is Nothing Then TakePipe StopPipeError, destruction Else TakeWebSocket StopPipeError, destruction

    '2. 破棄モード or 新規受信なし ならここで抜ける
    If destruction Or Not responseCDP.ExistsNewRes Then Exit Sub

    '3. `ReadFile` or `recv` で得たバイト配列をテキストに変換させ、メッセージ処理を準備
    Call ReadyCDPMessage
    printMsg Trace_, LogID, "Preparation for CDPmessage analysis is complete.", FromProcedureName

    '4. 解析を始めます
    Dim ResCounter As Long
    If responseCDP.ExistsNewRes Then printMsg Trace_, LogID, "Analyzing the data in the responseCDP.strBuffer...", FromProcedureName
    Do While responseCDP.ExistsNewRes
        '4-1. カウント
        ResCounter = ResCounter + 1

        '4-2. レスの確認
        Dim tmp As String: tmp = TakeCDPMessage
        If StrPtr(tmp) Then
            If currentLog.seeRawSendMsgDbg Then printMsg Debug_, LogID, "strRes(" & ResCounter & ") = " & tmp, FromProcedureName  'Only enable if needs to see every raw JS message for debugging/development

            '受信データのチェックを行うプロシージャへ遷移します。
            Call BrowserReceivedDataCheck(tmp)
        End If

        '4-3. 受信データのチェックの間に来たPipe/WebSocketも取り出しておき、リアルタイム性を確保
        If viaWebSocket Is Nothing Then TakePipe StopPipeError, False Else TakeWebSocket StopPipeError, False

        '4-4. バイト配列をテキストに変換させ、メッセージ処理を準備
        Call ReadyCDPMessage
    Loop

    '5. CDPメッセージ処理件数をログ化
    If ResCounter Then
        printMsg Debug_, LogID, "Executed " & ResCounter & " `RaiseEvent` calls.", FromProcedureName
    Else
        printMsg Trace_, LogID, "There were no messages received this time.", FromProcedureName
    End If
End Sub

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
            printMsg ERROR_, LogID, "A syntax error was detected when executing with ExecuteCDP/ExecuteCDPAsync. Processing was interrupted to prevent inconsistencies." & vbCrLf & "> " & CDPJsonString, FromProcedureName, doRaiseError:=True
        End If
    End If
End Sub



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
'* 機能　　：WebSocketモードでのメッセージ取得/解析に切り替えます
'***************************************************************************************************
Property Set WebSocketMode(ws As CDPCoreViaWebSocket)
    Set viaWebSocket = ws
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
        Dim tmp As Long: tmp = .Position
        .SetEOS

        '2. 保存
        .Position = 0
        .SaveToFile FolderPath & "\" & FileName, 2

        '3. 元に戻す
        .Position = tmp
    End With
End Sub



'***************************************************************************************************
'                               ■■■ 初期化/後始末処理 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    '------------------------- クラス内変数の一部を初期化 -------------------------
    '1. 初期バッファー設定
    '  文字列バッファーの初期化
    With responseCDP
        .length = InitialBuffer
        .strBuffer = String$(.length, vbNullChar)
    End With

    '  バイト配列バッファーの初期化
    ReDim binBuffer(InitialBuffer - 1) As Byte
    Set CDPPipeStream = CreateObject("ADODB.Stream")
    With CDPPipeStream
        .Open
        .Type = 1           ' adTypeBinary
        .Write binBuffer    ' 疑似的に大きめのバッファーを確保
        .Position = 0       ' カーソルだけ先頭に戻す
        .SetEOS             ' 終端（EOS）としてマークする
    End With

    '2. 設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        '2-1. イミディエイトに関する設定を行う
        currentLog.seeRawSendMsgDbg = .UseRangeID(8, FromProcedureName) 'ブラウザから受信した生のJSONメッセージもイミディエイトに表示させるか？

        '2-2. パスが設定されてたら、ログファイル保存を有効化
        Dim LogFolderPath As String: LogFolderPath = .UseRangeID(9, FromProcedureName)
        If LenB(LogFolderPath) > 0 Then
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

Private Sub Class_Terminate()
    On Error Resume Next
    CDPPipeStream.Close
    Set CDPPipeStream = Nothing ' 役目を終えたので、Windowsにメモリを返してあげる
End Sub



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
'        printMsg Trace_, LogID, "This response is the command result. ID:" & idValueContents, FromProcedureName
'    Else
'        GetContentsID = False
'        printMsg Trace_, LogID, "This response does not have an `id`.", FromProcedureName
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
'        printMsg Trace_, LogID, "This response is an asynchronous event: " & methodValueContents, FromProcedureName
'    Else
'        GetContentsMethod = False
'        printMsg Trace_, LogID, "This response does not have an `method`.", FromProcedureName
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
'        printMsg Trace_, LogID, "This response is the short command result for the browser.", FromProcedureName
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
'        printMsg Trace_, LogID, "This response is the command result for the tab. sessionId:" & sessionIdValueContents, FromProcedureName
'    Else
'        GetContentsSessionId = False
'        printMsg Trace_, LogID, "This response is the command result for the browser.", FromProcedureName
'    End If
'End Function
