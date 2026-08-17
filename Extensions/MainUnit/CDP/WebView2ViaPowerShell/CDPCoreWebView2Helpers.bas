Attribute VB_Name = "CDPCoreWebView2Helpers"
'***************************************************************************************************
'   WebView2ViaPowerShell - VBA用 名前付きパイプ接続ヘルパー
'---------------------------------------------------------------------------------------------------
'   PowerShell(StartWebView2Pipe.ps1)がホストするWebView2への名前付きパイプ接続を確立し、
'   ハンドル情報をExcelテーブルに記録するだけの、薄い「接続専用」モジュールです。
'
'   CDPコマンドの送受信・フレーミング・イベント処理は本モジュールでは一切行いません。
'   既存の`CDPCore.cls`/`CDPBrowser.cls`/`CDPContext.cls`(StarterWebScrapingKit)が、匿名パイプの
'   場合と全く同じように、そのまま読み書きしてくれます(`hCDPOutRd`/`hCDPInWr`はパイプハンドルで
'   ある以上、生成元がCreateProcessかCreateFile(名前付きパイプ)かを問わないため)。
'   本モジュールが担うのは「PowerShellホストを起動し、名前付きパイプへ接続して、そのハンドルを
'   `CDPCore.cls`が読める形式でExcelテーブルへ記録する」という、既存クラスに無い部分だけです。
'
'   【使い方・自動起動版】(`WebSocketViaPowerShell\CDPCoreWebSocketHelpers.bas`のデモと同じ作法)
'       CDPCoreWebView2Helpers.ConnectNamePipe "WebView2CDP"
'
'       Dim c As New CDPContext
'       If Not c.reattach("WebView2CDP", False) Then
'           Set c = c.InheritanceCDPBrowser.newTab(setMain:=True)
'       End If
'       c.navigate "https://example.com"
'
'   【使い方・手動起動版】
'       ExcelがPowerShellを直接`Shell`起動すると、環境によってはアンチウイルス/EDRの
'       「Office→PowerShell」ヒューリスティックに誤検知されることがあります。その場合は、
'       PowerShellを自分で(Windows Terminal等から)先に起動しておき、VBA側は接続のみ行ってください。
'
'       1. 手元のPowerShellコンソールで、以下を実行して待機させる：
'            powershell.exe -sta -NoProfile -ExecutionPolicy Bypass -File "...\WebView2ViaPowerShell\StartWebView2Pipe.ps1" -PipeName "WebView2CDP"
'          「名前付きパイプ サーバー:WebView2CDP を起動しました。Excelからの接続を待機しています...」
'          と表示されたら準備完了です。
'
'       2. VBA側は`Shell`を呼ばず、接続だけ行う(数十秒リトライして待ちます)：
'            CDPCoreWebView2Helpers.ConnectNamePipeManual "WebView2CDP"
'
'            Dim c As New CDPContext
'            If Not c.reattach("WebView2CDP", False) Then
'                Set c = c.InheritanceCDPBrowser.newTab(setMain:=True)
'            End If
'            c.navigate "https://example.com"
'***************************************************************************************************
Option Explicit
Option Private Module



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

Private Declare PtrSafe Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const PIPE_Landmark         As String = "\\.\pipe\"
Private Const GENERIC_READ          As Long = &H80000000
Private Const GENERIC_WRITE         As Long = &H40000000
Private Const OPEN_EXISTING         As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE  As LongPtr = -1
Private Const ScriptRelativePath    As String = "WebView2ViaPowerShell\StartWebView2Pipe.ps1"



'***************************************************************************************************
'                              ■■■ 接続本体 ■■■
'***************************************************************************************************
'* 機能　　：PowerShellホストを（未起動なら）起動し、名前付きパイプへ接続してExcelテーブルに記録します
'---------------------------------------------------------------------------------------------------
'* 返り値　：接続成否
'* 引数　　：UserName           識別名称。パイプ名にもそのまま使用されます(`BrowserHandleInfo`テーブルの主キー)
'            UserDataFolder     WebView2のユーザーデータフォルダ。省略時は`%LOCALAPPDATA%`配下に自動生成
'            TimeoutSeconds     パイプ出現待ちのタイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 詳細説明：既に生きているパイプが`BrowserHandleInfo`テーブルに記録されていれば、何もせず終了します。
'            生存確認は`CDPCore.cls`のpublicメソッド(`deserialize`/`isAvailability`)をそのまま流用し、
'            本モジュール側で`PeekNamedPipe`等を独自に持つことはしません。
'***************************************************************************************************
Public Function ConnectNamePipe(Optional UserName As String = "WebView2CDP", Optional UserDataFolder As String, Optional TimeoutSeconds As Double = 10) As Boolean
    '1. 既存の生きたパイプがあれば、何もせず終了 (CDPCore自身のdeserialize/isAvailabilityを流用)
    If isAlreadyConnected(UserName) Then ConnectNamePipe = True: Exit Function

    '2. WebView2用の作業フォルダを確定
    If LenB(UserDataFolder) = 0 Then UserDataFolder = Environ$("LOCALAPPDATA") & "\WebView2ViaPowerShell\" & UserName

    '3. PowerShell(64bit固定)をホストとして起動 ※SysWOW64の32bit版だと`WebView2Loader.dll`のアーキテクチャ不一致で失敗するため明示指定
    '   ※Excelから直接`Shell`起動すると、環境によってはアンチウイルス/EDRの「Office→PowerShell」
    '     ヒューリスティックに誤検知される場合があります。その場合は`ConnectNamePipeManual`をご利用ください。
    Dim psPath As String: psPath = Environ$("SystemRoot") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
    Dim scriptPath As String: scriptPath = ThisWorkbook.Path & "\" & ScriptRelativePath

    Dim cmd As String
    cmd = """" & psPath & """ -sta -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -PipeName """ & UserName & """ -UserDataFolder """ & UserDataFolder & """"

    Shell cmd, vbNormalNoFocus

    '4. パイプが出現するまでリトライ接続し、ハンドル情報をExcelテーブルに記録
    ConnectNamePipe = TryConnectAndRecord(UserName, TimeoutSeconds)
End Function

'***************************************************************************************************
'* 機能　　：PowerShellを自分で(Windows Terminal等から)先に起動しておいた場合の、接続専用の入り口です
'---------------------------------------------------------------------------------------------------
'* 返り値　：接続成否
'* 引数　　：UserName           識別名称。手動起動したPowerShellの`-PipeName`引数と一致させてください
'            TimeoutSeconds     パイプ出現待ちのタイムアウト秒数(既定60秒。手動起動の時間差を考慮して長め)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`ConnectNamePipe`と異なり、`Shell`によるPowerShell起動は一切行いません。
'            `WebSocketViaPowerShell\CDPCoreWebSocketHelpers.ConnectNamePipe`と同じ位置づけの関数です。
'* 前提　　：あらかじめ、手元のPowerShellコンソールから下記を実行し、待受け状態にしておいてください：
'              powershell.exe -sta -NoProfile -ExecutionPolicy Bypass -File "...\StartWebView2Pipe.ps1" -PipeName "WebView2CDP"
'***************************************************************************************************
Public Function ConnectNamePipeManual(Optional UserName As String = "WebView2CDP", Optional TimeoutSeconds As Double = 60) As Boolean
    '1. 既存の生きたパイプがあれば、何もせず終了
    If isAlreadyConnected(UserName) Then ConnectNamePipeManual = True: Exit Function

    '2. パイプが出現するまでリトライ接続し、ハンドル情報をExcelテーブルに記録
    ConnectNamePipeManual = TryConnectAndRecord(UserName, TimeoutSeconds)
End Function

'***************************************************************************************************
'* 機能　　：`BrowserHandleInfo`テーブルに記録済みの、生きたパイプが既にあるか確認します
'---------------------------------------------------------------------------------------------------
'* 注意事項：死んでる記録が見つかった場合は、ここで後片付け(CloseHandle)しておきます
'***************************************************************************************************
Private Function isAlreadyConnected(UserName As String) As Boolean
    Dim probe As New CDPCore
    If probe.deserialize(UserName) Then
        If probe.isAvailability Then isAlreadyConnected = True: Exit Function
        probe.CloseHandleBrowser   '死んでいたので、後片付け
    End If
End Function

'***************************************************************************************************
'* 機能　　：名前付きパイプが出現するまでリトライ接続し、成功したらハンドル情報をExcelテーブルに記録します
'---------------------------------------------------------------------------------------------------
'* 返り値　：接続成否
'***************************************************************************************************
Private Function TryConnectAndRecord(UserName As String, TimeoutSeconds As Double) As Boolean
    Dim hNamePipe As LongPtr
    Dim StartTime As Double: StartTime = Timer
    Do
        hNamePipe = CreateFile(PIPE_Landmark & UserName, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
        If hNamePipe <> INVALID_HANDLE_VALUE Then Exit Do
        SleepAPI 200
        DoEvents
    Loop While Timer - StartTime < TimeoutSeconds

    If hNamePipe = INVALID_HANDLE_VALUE Then Exit Function

    'ハンドル情報(PIDは手動起動の場合、VBA側からは分からないため`0`)を、`CDPCore.cls`が読める形式で記録
    serialize UserName, hNamePipe, 0
    TryConnectAndRecord = True
End Function

'***************************************************************************************************
'* 機能　　：パイプハンドル情報を、`CDPCore.cls`と同一スキーマでExcelテーブルに記録します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`hStdOutRd`等、匿名パイプ専用の項目は`0`で埋めます(`WebSocketViaPowerShell\CDPCoreWebSocketHelpers.serialize`と同じ考え方)。
'            `dwProcessId`にはPowerShellのPIDを保存しておき、`CDPCore.Connection_dwProcessId`経由で
'            参照できるようにします(異常終了時の後片付け等に利用可能)。
'***************************************************************************************************
Private Sub serialize(UserName As String, hNamePipe As LongPtr, psProcessId As Long)
    '------------------ 1. パイプ情報の記録準備 ------------------
    '※主要となる情報以外は一旦、一律0とし、必要なデータを`Dictionary`に詰める
    Dim tmp As New Dictionary
    tmp.Add "hStdOutRd", 0
    tmp.Add "hStderrOutRd", 0
    tmp.Add "hStdInWr", 0
    tmp.Add "hCDPOutRd", (hNamePipe)
    tmp.Add "hCDPInWr", (hNamePipe)
    tmp.Add "hProcess", 0
    tmp.Add "dwProcessId", psProcessId

    'Excelテーブルに、名前付きパイプハンドル情報を記録する
    Set ShSetting01_StartBrowser.TableBrowserHandle(UserName, "CDPCoreWebView2Helpers.serialize") = tmp


    '------------------ 2. タブ情報の記録準備 ------------------
    '※一律空欄とし、必要なデータ枠を`Dictionary`に詰める
    tmp.RemoveAll
    tmp.Add "BiDi-context", vbNullString
    tmp.Add "sessionID", vbNullString
    tmp.Add "targetID", vbNullString

    'Excelテーブルに、タブ情報欄を確保する
    Set ShSetting01_StartBrowser.TableBrowserContext(UserName, "CDPCoreWebView2Helpers.serialize") = tmp

End Sub
