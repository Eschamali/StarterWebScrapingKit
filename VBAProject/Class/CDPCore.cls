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

'===================================
' Win APIs Declarations
'===================================

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
 
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal length As LongPtr)
 
'===================================
' Types & Structures Declarations
'===================================
 
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
 
'===================================
' Constants
'===================================
 
Private Const STARTF_USESTDHANDLES = &H100&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_RESTORE = 9

'===================================
' Shared Variables
'===================================
 
Private hStdOutWr As LongPtr
Public hStdOutRd As LongPtr
Private hStderrOutWr As LongPtr
Public hStderrOutRd As LongPtr
Public hStdInWr As LongPtr
Private hStdInRd As LongPtr
Private hCDPOutWr As LongPtr
Public hCDPOutRd As LongPtr
Public hCDPInWr As LongPtr
Private hCDPInRd As LongPtr
Private hProcess As LongPtr
Public dwProcessId As LongPtr 'プロセスID
Private Utf8Converter As CharacterCodeConversion



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



'===================================
' Class Functions
'===================================
 
Public Function serialize(UserName As String) As String
'-------------------------------------------------------------------
' Create serialized Json string from the current browser instance.
'-------------------------------------------------------------------

    'データ記録準備
    Dim tmp As String
    With WorksheetFunction
        tmp = "{""" & .TextJoin(""";""", False, _
                            .TextJoin(""",""", False, UserName, "hStdOutRd", hStdOutRd), _
                            .TextJoin(""",""", False, UserName, "hStderrOutRd", hStderrOutRd), _
                            .TextJoin(""",""", False, UserName, "hStdInWr", hStdInWr), _
                            .TextJoin(""",""", False, UserName, "hCDPOutRd", hCDPOutRd), _
                            .TextJoin(""",""", False, UserName, "hCDPInWr", hCDPInWr), _
                            .TextJoin(""",""", False, UserName, "dwProcessId", dwProcessId) _
        ) & """}"
    End With

    'Excelテーブルに、パイプハンドル情報を記録する
    'Evaluate関数を利用して、上記で作った記述方式で簡易的な2次元配列を作成します
    ShSetting01_StartBrowser.serializeToTable(1, "CDPCore.serialize") = Evaluate(tmp)

    '作成情報の返却
    serialize = tmp

End Function

Public Sub deserialize(UserName As String)
'--------------------------------------------------------------------
' Decipher serialized Json string and assign values to CDP variables.
'--------------------------------------------------------------------

    'Excelテーブルから、`UserName`に紐づくパイプハンドル情報を読み込む
    Dim PipeInfo As Dictionary
    Set PipeInfo = ShSetting01_StartBrowser.deserializeFromTable(UserName, 1, "CDPCore.deserialize")

    'クラス内変数に適用
    hStdOutRd = PipeInfo("hStdOutRd")
    hStderrOutRd = PipeInfo("hStderrOutRd")
    hStdInWr = PipeInfo("hStdInWr")
    hCDPOutRd = PipeInfo("hCDPOutRd")
    hCDPInWr = PipeInfo("hCDPInWr")
    dwProcessId = PipeInfo("dwProcessId")

End Sub

Public Function Init(strExec As String) As Long
'-------------------------------------------------------------------
' Initialize new browser instance via pipe creation.
'-------------------------------------------------------------------
 
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim result As Long
    Dim pipes As STDIO_BUFFER
    Dim pipes2 As STDIO_BUFFER2
 
   'First we create all 4 pipes
   'We start with stdout of the edge process
    sa.nLength = LenB(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    result = CreatePipe(hStdOutRd, hStdOutWr, sa, 0)
    If result = 0 Then
        Init = -2
        Exit Function
    End If

    'stderr
    result = CreatePipe(hStderrOutRd, hStderrOutWr, sa, 0)
    If result = 0 Then
        Init = -2
        Exit Function
    End If

   'Then stdin
    result = CreatePipe(hStdInRd, hStdInWr, sa, 0)
    If result = 0 Then
        Init = -2
        Exit Function
    End If
    
   'Then the out pipe for the CDP Protocol
    result = CreatePipe(hCDPOutRd, hCDPOutWr, sa, 2 ^ 20)
    If result = 0 Then
        Init = -2
        Exit Function
    End If
    
   'And finally the in pipe
    result = CreatePipe(hCDPInRd, hCDPInWr, sa, 0)
    If result = 0 Then
        Init = -2
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
    
    MoveMemory pipes2.raw_bytes(0), pipes.crt_flags(0), 5
    MoveMemory pipes2.raw_bytes(5), pipes.os_handle(0), UBound(pipes2.raw_bytes) - 4
    
   'Define new instance parameters and start it
    With start
        .cb = LenB(start)
        .dwFlags = STARTF_USESHOWWINDOW
        .hStdOutput = hStdOutWr
        .hStdInput = hStdInRd
        .hStdError = hStderrOutWr
        .wShowWindow = SW_SHOWMINIMIZED 'minimize the window at start
        .cbReserved2 = LenB(pipes2)
        .lpReserved2 = VarPtr(pipes2)
    End With

    result = CreateProcess(0&, StrPtr(strExec), sa, sa, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    '実行結果チェック
    If result = 0 Then
        'エラーコードを返す
        Init = Err.LastDllError

        '失敗なので、自身で使用するはずだったハンドルも削除
        CloseHandle hStdOutRd
        CloseHandle hStdInWr
        CloseHandle hCDPOutRd
        CloseHandle hCDPInWr
        CloseHandle hStderrOutRd
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
'* 機能　　：CDP-Jsonコマンド結果や、ブラウザからのイベント情報の読み取りを試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：読み取ったバイト数を返します。
'* 引数　　：strData            CDP-Json の実行結果、ブラウザからのイベント情報を格納したい変数。実質返り値です
'            Err_PeekNamedPipe  PeekNamedPipe にてエラーが起こった際のエラーコードを格納します。実質返り値です
'            Err_ReadFile       ReadFile にてエラーが起こった際のエラーコードを格納します。実質返り値です
'---------------------------------------------------------------------------------------------------
'* 機能説明：読み取りは非ブロッキングです。データがなくても、次に進めれます。各種エラーコードは、第2引数以降で確認して下さい。
'***************************************************************************************************
Public Function readProcCDP(ByRef strData As String, ByRef Err_PeekNamedPipe As Long, ByRef Err_ReadFile As Long) As Long
    'エラーコードの初期化
    Err_PeekNamedPipe = 0
    Err_ReadFile = 0

    Dim lPeekData As Long
    Dim lngRes As Long
    
    '処理が成功すると、戻り値は 0 以外になります。
    lngRes = PeekNamedPipe(hCDPOutRd, ByVal 0&, 0&, ByVal 0&, lPeekData, ByVal 0&)

    'API関数:PeekNamedPipe が失敗した場合は、0 を返すため、`Err_PeekNamedPipe`にエラーコードを格納
    If lngRes = 0 Then Err_PeekNamedPipe = Err.LastDllError: Exit Function


    'データがあったら実際に、データを入手する処理へ
    If lPeekData > 0 Then
        Dim Buffer As String
        Dim L As Long
        Dim bSuccess As Long
        Buffer = Space$(lPeekData)
        bSuccess = ReadFile(hCDPOutRd, ByVal Buffer, Len(Buffer), L, 0&)

        '0以外なら成功です
        If bSuccess Then
            strData = Buffer
            readProcCDP = Len(strData)
        Else
            'API関数:ReadFile が失敗した場合は、0 を返すため、`Err_ReadFile`にエラーコードを格納
            Err_ReadFile = Err.LastDllError
        End If
    End If
End Function

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
'* 注意事項：ロジック自体は、前述の`readProcCDP`と一緒ですが、そこまで重要視しないため、簡略化してます。
'            実際のエラーコードも特に確認しません。
'***************************************************************************************************
Public Function readProcSTD(ByRef strData As String, ErrorMode As Boolean) As Long
    '使用するパイプ通信判定
    Dim UsehStdRd As LongPtr: UsehStdRd = IIf(ErrorMode, hStderrOutRd, hStdOutRd)

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

Public Function writeProc(ByVal strData As String) As Long
' --------------------------------------------------------------------------------
' This functions sends a CDP message to the Chromium browser.
'---------------------------------------------------------------------------------
 
    Dim lngWritten As Long
    
    'UTF-8バイト配列を使用するか？
    If Utf8Converter Is Nothing Then
        'VBAの「Unicode」文字列そのまんまで、コマンドを送ります
        writeProc = WriteFile(hCDPInWr, ByVal strData, Len(strData), lngWritten, ByVal 0&)
    Else
        'VBAの「Unicode」文字列をUTF-8でエンコードしたバイト配列を取得します
        Dim BinData() As Byte: BinData = Utf8Converter.BytesFromString(strData)

        'そのバイト配列として、コマンドを送ります
        writeProc = WriteFile(hCDPInWr, ByVal VarPtr(BinData(0)), UBound(BinData) + 1, lngWritten, ByVal 0&)
    End If
End Function



'***************************************************************************************************
'* 機能　　：設定シートから、設定値を読み込みます
'---------------------------------------------------------------------------------------------------
'* 機能説明：クラスモジュール構築時、設定値を適用させます
'***************************************************************************************************
Private Sub Class_Initialize()
    '------------------------- クラス内変数の一部を初期化 -------------------------
    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        UseConverterUtf8FromString = .Range(.UseRangeName(10, "CDPCore.Class_Initialize")).value
    End With
End Sub
