VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerShellViaStdPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
' 【概要】
' 本クラスは、Windows APIによる双方向非ブロッキング・パイプラインを用いた
' セキュリティ・バイパス型 PowerShell ホスト・エミュレータです。
'
' 【背景】
' 通常、VBAからShell関数等を用いてPowerShellを起動する際、引数（-Command等）に
' スクリプト内容を付与しますが、これは現代のセキュリティソフト（Norton/CrowdStrike等）が
' 監視する「Command Line Inspection」の格好の標的となり、不当にブロックされます。
'
' 【解決策】
' 本モジュールは、起動時の引数を「手ぶら」の状態にして検閲を回避した上で、
' プロセス開通後に「標準入力(StdIn)」ストリームを介してコードを動的に注入します。
' これは、ユーザーが Terminal.exe から手動でタイピングを行う「正当な対話操作」を
' ソフトウェア・レイヤーで完全に模倣するものであり、セキュリティの文脈(Context)を
' 書き換えることで、本来の実行能力を安全に解放します。
' 本モジュールは、 CDP周りの拡張も入ってますが、PowerShell周りのやり取りコアは整ってるので、他の用途にも転用しやすいかと思います
'
' 【役割】
' 本モジュールは、 Excel ←匿名Pipe→ PowerShell ←WebSocket→ Chromium といった連携確保まで担います。
' 本来、このツールは`remote-debugging-pipe`特化で構築してる都合上、`remote-debugging-port`の直接統合は堅牢性を落としてしまう問題がありました。
' そこで、`PowerShell × 匿名Pipe`という仕組みにより、拡張機能として提供することで、コア部分をいじらず、実質対応可能となりました。
' 予め、所定の`PowerShell`コードをセルに埋め込んで起動しておくことで、下記のようなニッチな目的にも対応できます
' ・AndroidスマートフォンのChromiumブラウザの自動制御
' ・WebView2向けの環境変数`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`に、`--remote-debugging-port=9222`を付与した自動制御
' ・`inspect/#remote-debugging`による今、目の前のブラウザから、自動制御
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                 ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function CreatePipe Lib "kernel32" ( _
    phReadPipe As LongPtr, _
    phWritePipe As LongPtr, _
    lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long

Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, _
    lpThreadAttributes As SECURITY_ATTRIBUTES, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    lpEnvironment As Any, _
    ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As LongPtr, _
    lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long) As Long

Private Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As LongPtr) As Long

Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As LongPtr) As Long

Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long



'***************************************************************************************************
'                            ■■■ WindowsAPIに関する構造体 ■■■
'***************************************************************************************************
#If Win64 And VBA7 Then
    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        padding1 As Long
        lpSecurityDescriptor As LongPtr
        bInheritHandle As Long
        padding2 As Long
    End Type

    Private Type STARTUPINFO
        cb As Long
        padding1 As Long
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
        padding2 As Long
        lpReserved2 As LongPtr
        hStdInput As LongPtr
        hStdOutput As LongPtr
        hStdError As LongPtr
    End Type
#Else
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
#End If

Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

' CreateProcess関数の第6引数 dwCreationFlags 用の列挙型
Public Enum ProcessCreationFlags
    ' 親プロセスがデバッガとして動作し、子プロセスで発生するイベントを監視する
    DEBUG_PROCESS = &H1
    
    ' デバッグ対象をこのプロセスのみに限定する（子プロセスには波及させない）
    DEBUG_ONLY_THIS_PROCESS = &H2
    
    ' プロセスを「停止状態」で作成する。ResumeThreadを呼ぶまで動き出さない。
    ' ※プロセスのメモリを書き換えてから動かしたい場合などに使用。
    CREATE_SUSPENDED = &H4
    
    ' コンソールアプリを「コンソールなし」で起動する。
    ' ※親のコンソールを引き継がず、かつ新しいコンソールも作らない。
    DETACHED_PROCESS = &H8
    
    ' 新しいコンソール画面を作成して起動する。
    CREATE_NEW_CONSOLE = &H10
    
    ' プロセスグループのリーダーとして作成する（Ctrl+Cシグナルの管理用）
    CREATE_NEW_PROCESS_GROUP = &H200
    
    ' 環境変数ブロックがUnicode(UTF-16)であることを示す。
    ' ※VBAからUnicode環境変数を渡す場合に必須。
    CREATE_UNICODE_ENVIRONMENT = &H400
    
    ' 16bit Windowsアプリ（WOW）を個別の仮想マシンで動かす
    CREATE_SEPARATE_WOW_VDM = &H800
    
    ' 子プロセスがジョブオブジェクト（制限の枠組み）から外れることを許可する
    CREATE_BREAKAWAY_FROM_JOB = &H1000000
    
    ' ★【本命】コンソールウィンドウを物理的に作成せずに起動する。
    ' これを指定すると黒い窓は一切出ず、かつGUIのShowDialog等は普通に表示できる。
    CREATE_NO_WINDOW = &H8000000
    
    ' --- 優先度クラス（これらも第6引数に混ぜることができます） ---
    IDLE_PRIORITY_CLASS = &H40            ' 低
    BELOW_NORMAL_PRIORITY_CLASS = &H4000  ' やや低
    NORMAL_PRIORITY_CLASS = &H20          ' 標準
    ABOVE_NORMAL_PRIORITY_CLASS = &H8000  ' やや高
    HIGH_PRIORITY_CLASS = &H80            ' 高
    REALTIME_PRIORITY_CLASS = &H100       ' リアルタイム（極めて危険）
End Enum



'***************************************************************************************************
'                               ■■■ 各種handle ■■■
'***************************************************************************************************
' --- 一般的なハンドル ---
Private hStdOutWr As LongPtr
Private hStdOutRd As LongPtr
Private hStderrOutWr As LongPtr
Private hStderrOutRd As LongPtr
Private hStdInWr As LongPtr
Private hStdInRd As LongPtr
Private hProcess As LongPtr
Private dwProcessId As LongPtr

' --- 拡張用ハンドル ---
Private hCDPOutWr As LongPtr
Private hCDPOutRd As LongPtr
Private hCDPInWr As LongPtr
Private hCDPInRd As LongPtr



'***************************************************************************************************
'                               ■■■ 管理変数 ■■■
'***************************************************************************************************
Private sa              As SECURITY_ATTRIBUTES
Private m_EnvRegistry   As Dictionary ' セットした環境変数の名前を記録するメモ帳

Public UseStdOuterr As Boolean
Public UseStdOut    As Boolean



'***************************************************************************************************
'                               ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const scriptName            As String = "powershell.exe"
'Private Const scriptName            As String = "PowerShell_ISE.exe"
Private Const STARTF_USESTDHANDLES  As Long = &H100
Private Const STARTF_USESHOWWINDOW  As Long = &H1



'***************************************************************************************************
'                             ■■■ 基本となる処理 ■■■
'***************************************************************************************************
'* 機能　　：対話型として、PowerShell を起動します
'---------------------------------------------------------------------------------------------------
'* 返り値　：起動結果   ※0で成功です
'* 引数　　：showMode   ウィンドウの表示モードを指定します
'            startMode  プロセス生成段階での起動モードを指定します。`OR`による複数指定も可能
'---------------------------------------------------------------------------------------------------
'* 機能説明：事前のプロパティ設定で、標準エラー/出力の無効化も可能です
'* 注意事項：追加のパイプハンドルも渡す場合は、それも作ったうえでこれを呼び出すこと
'***************************************************************************************************
Public Function Init(Optional showMode As WinState, Optional startMode As ProcessCreationFlags) As Long
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION

    ' --- 1. 基本セットとなるパイプを作成する ---
    If UseStdOut Then CreatePipe hStdOutRd, hStdOutWr, sa, 0
    CreatePipe hStdInRd, hStdInWr, sa, 0
    If UseStdOuterr Then CreatePipe hStderrOutRd, hStderrOutWr, sa, 0

    ' --- 2. STARTUPINFOの設定 ---
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
        .hStdOutput = hStdOutWr
        .hStdInput = hStdInRd
        .hStdError = hStderrOutWr
        .wShowWindow = showMode
    End With

    ' --- 3. PowerShellの起動 ---
    Dim success As Long
    success = CreateProcess(vbNullString, scriptName, sa, sa, 1, startMode, ByVal 0&, vbNullString, si, pi)

    '実行結果チェック
    If success = 0 Then
        'エラーコードを返す
        Init = Err.LastDllError

        '失敗なので、自身で使用するはずだったハンドルも削除
        ClosePipeCDP
    End If

    'PowerShell側へ渡した物は、破棄しておく
    If UseStdOut Then CloseHandle hStdOutWr
    CloseHandle hStdInRd
    If UseStdOuterr Then CloseHandle hStderrOutWr

    hProcess = pi.hProcess
    dwProcessId = pi.dwProcessId

End Function

'***************************************************************************************************
'* 機能　　：PowerShellからのレスポンス取得を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：読み取ったバイト数を返します。
'* 引数　　：ErrorMode          `True`  標準出力エラーで読み取り
'                               `False` 標準出力で読み取り
'
'            strData            PowerShell からのレスポンス。実質返り値です
'            Err_PeekNamedPipe  PeekNamedPipe にてエラーが起こった際のエラーコードを格納します。実質返り値です
'            Err_ReadFile       ReadFile にてエラーが起こった際のエラーコードを格納します。実質返り値です
'---------------------------------------------------------------------------------------------------
'* 機能説明：読み取りは非ブロッキングです。データがなくても、次に進めれます。各種エラーコードは、第2引数以降で確認して下さい。
'* 注意事項：`strData`を初期化せずに、使いまわす場合、次のような仕様になります
'            ・`PeekNamedPipe`にて、受信データ量が0あるいは、エラー：最後に受け取った`ReadFile`の内容を保持します
'            ・`PeekNamedPipe`にて、受信データ量が1以上　　　　　　：その受信した`ReadFile`の内容に置き換わります
'***************************************************************************************************
Public Function readProcSTD(ErrorMode As Boolean, ByRef strData As String, Optional ByRef Err_PeekNamedPipe As Long, Optional ByRef Err_ReadFile As Long) As Long
    '使用するパイプ通信判定
    Dim UsehStdRd As LongPtr: UsehStdRd = IIf(ErrorMode, hStderrOutRd, hStdOutRd)

    'エラーコードの初期化
    Err_PeekNamedPipe = 0
    Err_ReadFile = 0

    Dim lPeekData As Long
    Dim lngRes As Long
    
    '処理が成功すると、戻り値は 0 以外になります。
    lngRes = PeekNamedPipe(UsehStdRd, ByVal 0&, 0&, ByVal 0&, lPeekData, ByVal 0&)

    'API関数:PeekNamedPipe が失敗した場合は、0 を返すため、`Err_PeekNamedPipe`にエラーコードを格納
    If lngRes = 0 Then Err_PeekNamedPipe = Err.LastDllError: Exit Function


    'データがあったら実際に、データを入手する処理へ
    If lPeekData > 0 Then
        Dim Buffer() As Byte
        Dim bSuccess As Long
        ReDim Buffer(lPeekData - 1)
        bSuccess = ReadFile(UsehStdRd, ByVal VarPtr(Buffer(0)), lPeekData, 0&, 0&)

        '0以外なら成功です
        If bSuccess Then
            'そのまんま返却
            strData = StrConv(Buffer, vbUnicode)

            '取得サイズを返却
            readProcSTD = lPeekData
        Else
            'API関数:ReadFile が失敗した場合は、0 を返すため、`Err_ReadFile`にエラーコードを格納
            Err_ReadFile = Err.LastDllError
        End If
    End If
End Function

'***************************************************************************************************
'* 機能　　：PowerShellへコマンドを送ります
'---------------------------------------------------------------------------------------------------
'* 返り値　：`WriteFile`の実行結果　※0で成功
'* 引数　　：strData    送る内容
'---------------------------------------------------------------------------------------------------
'* 注意事項：・標準出力が有効な状態で一気に大量のコマンド文字列を送ると、バッファオーバーにより、内部でフリーズします。
'            　改行ごとに送るのがベストです。
'            ・PowerShellの特性上、使用環境の規定文字コードバイト配列に変換したうえで送ります
'***************************************************************************************************
Public Function writeProcSTD(ByVal strData As String) As Long

    Dim lngWritten As Long
    Dim ResultCode As Long

    '規定文字コードバイトで、コマンドを送ります
    Dim sjisBuf() As Byte: sjisBuf = StrConv(strData, vbFromUnicode)
    ResultCode = WriteFile(hStdInWr, sjisBuf(0), UBound(sjisBuf) + 1, lngWritten, ByVal 0&)

    '失敗時は、実際のエラーコードで返します
    If ResultCode = 0 Then writeProcSTD = Err.LastDllError
End Function

'***************************************************************************************************
'* 機能　　：Excel側で使う 標準出力/エラー/入力 パイプハンドル情報をクリーニングします
'***************************************************************************************************
Public Sub ClosePipeStd()
    CloseHandle hStdOutRd
    CloseHandle hStdInWr
    CloseHandle hStderrOutRd
End Sub

'***************************************************************************************************
'* 機能　　：PowerShellへ受け渡す際の環境変数を作成します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Name   環境変数名
'            Value  登録値  ※`vbnullstring`で、削除となります
'---------------------------------------------------------------------------------------------------
'* 注意事項：厳密には、起動中のExcel.exe全体から見える環境変数です
'***************************************************************************************************
Property Let RegisterAndSetEnv(ByVal Name As String, ByVal value As String)
    '環境変数にセット
    Call SetEnvironmentVariableW(StrPtr(Name), StrPtr(value))

    '削除の場合は、リストからも消して抜ける
    If StrPtr(value) = 0 Then m_EnvRegistry.Remove (Name): Exit Property

    '「後で消すリスト」に名前を登録しておく
    If m_EnvRegistry.Exists(Name) Then m_EnvRegistry(Name) = value Else m_EnvRegistry.Add Name, value
End Property

'***************************************************************************************************
'* 機能　　：登録したリストを基に、環境変数を削除します
'***************************************************************************************************
Public Sub CleanAllRegisteredEnv()
    Dim varName: varName = m_EnvRegistry.keys

    ' メモ帳に書かれた名前を、上から順に自動で消していく
    Dim i As Long
    For i = LBound(varName) To UBound(varName)
        Call SetEnvironmentVariableW(StrPtr(varName(i)), 0)
    Next

    ' メモ帳を空にする
    Set m_EnvRegistry = New Dictionary
End Sub



'***************************************************************************************************
'                          ■■■ CDP用として拡張した専用領域 ■■■
'***************************************************************************************************
'* 機能　　：PowerShellへCDP送信したときの読み取りパイプハンドルを作成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：PowerShellへ渡すハンドル値 ※0は、失敗です
'---------------------------------------------------------------------------------------------------
'* 注意事項：`Init`の前に呼び出すこと
'***************************************************************************************************
Public Function CreatePipeForCDPInRd() As LongPtr

    'パイプハンドル生成
    CreatePipe hCDPInRd, hCDPInWr, sa, 2 ^ 20

    '返却
    CreatePipeForCDPInRd = hCDPInRd
End Function

'***************************************************************************************************
'* 機能　　：PowerShellからCDP結果等を送信するときの書き込みパイプハンドルを作成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：PowerShellへ渡すハンドル値 ※0は、失敗です
'---------------------------------------------------------------------------------------------------
'* 注意事項：`Init`の前に呼び出すこと
'***************************************************************************************************
Public Function CreatePipeForCDPOutWr() As LongPtr

    'パイプハンドル生成
    CreatePipe hCDPOutRd, hCDPOutWr, sa, 2 ^ 20

    '返却
    CreatePipeForCDPOutWr = hCDPOutWr
End Function

'***************************************************************************************************
'* 機能　　：PowerShellへの接続情報を記録します
'---------------------------------------------------------------------------------------------------
'* 返り値　：任意の利用ユーザー名
'---------------------------------------------------------------------------------------------------
'* 注意事項：`Init`の後に呼び出すこと
'***************************************************************************************************
Public Sub serialize(UserName As String)
    '------------------ 1. パイプ情報の記録準備 ------------------
    Dim tmp As New Dictionary
    tmp.Add "hStdOutRd", hStdOutRd
    tmp.Add "hStderrOutRd", hStderrOutRd
    tmp.Add "hStdInWr", hStdInWr
    tmp.Add "hCDPOutRd", hCDPOutRd
    tmp.Add "hCDPInWr", hCDPInWr
    tmp.Add "dwProcessId", dwProcessId

    'Excelテーブルに、名前付きパイプハンドル情報を記録する
    Set ShSetting01_StartBrowser.serializeToTable(UserName, 1, "PowerShellViaStdPipe.serialize") = tmp


    '------------------ 2. タブ情報の記録準備 ------------------
    '※主要となる情報以外は一旦、一律空欄とし、必要なデータを`Dictionary`に詰める
    tmp.RemoveAll
    tmp.Add "lngLastID", 1
    tmp.Add "sessionID", ""
    tmp.Add "targetID", ""

    'Excelテーブルに、タブ情報欄を確保する
    Set ShSetting01_StartBrowser.serializeToTable(UserName, 2, "PowerShellViaStdPipe.serialize") = tmp
End Sub

'***************************************************************************************************
'* 機能　　：Excel側で使う CDPの入出力 パイプハンドル情報をクリーニングします
'***************************************************************************************************
Public Sub ClosePipeCDP()
    CloseHandle hCDPOutRd
    CloseHandle hCDPInWr
End Sub



'***************************************************************************************************
'                          ■■■ クラス構築/破棄時挙動 ■■■
'***************************************************************************************************
'* 機能　　：いくつかの設定値を適用させます
'***************************************************************************************************
Private Sub Class_Initialize()
    '標準出力/エラーをdefaultで使用
    UseStdOuterr = True
    UseStdOut = True

    'パイプ継承の設定を適用
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1 ' パイプを継承可能にする

    'メモ帳の初期化
    Set m_EnvRegistry = New Dictionary
End Sub

'***************************************************************************************************
'* 機能　　：PowerShellに渡した拡張分のパイプハンドル情報を破棄します
'***************************************************************************************************
Private Sub Class_Terminate()
    CloseHandle hCDPOutWr
    CloseHandle hCDPInRd
End Sub
