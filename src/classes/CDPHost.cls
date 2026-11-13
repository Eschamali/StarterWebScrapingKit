VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'                             ローカルPC上のブラウザに関する機能群です
'                            Chromium系のみ取り扱います。※WebView2を除く
'==============================================================================================================
'【ポリシーに関する補足】
'  企業PCでは、ブラウザのポリシー設定によって CDP 制御が制限されることがあります。
'  このクラスは起動前にレジストリからポリシーを読み、分かる範囲で事前に警告・停止します。
'
'  ■ 関係するポリシー
'    ・RemoteDebuggingAllowed               0 だと、デバッグポート/パイプ自体が開かない（接続不可）
'    ・DeveloperToolsAvailability           接続後、どのターゲットを attach できるかを決める
'                                             未設定/0:強制インストール拡張のみ禁止 / 1:全許可 / 2:全禁止
'                                             ※禁止されたターゲットは Target.attachToTarget が「Not allowed」で失敗
'    ・DeveloperToolsAvailabilityAllowlist  (Edge148～) URL単位の許可リスト。Blocklist より優先
'    ・DeveloperToolsAvailabilityBlocklist  (Edge148～) URL単位の禁止リスト
'                                             ※どちらにも一致しない URL は DeveloperToolsAvailability に従う
'    ・UserDataDir                          データ保存先の固定。デフォルトパス固定だとデバッグ不可
'
'  ■ どの値が採用されるか（「拒否優先」ではない）
'    同じポリシーが複数箇所にある場合、ポリシーごとに勝者を1つ決め、その値だけが使われる
'      1. 必須(Mandatory) ＞ 推奨(Recommended)  ※上記ポリシーはすべて必須のみなので今は、気にしなくて良い
'      2. HKLM ＞ HKCU（HKLM に値があれば、HKCU は無視される）
'         例）HKLM=1(許可) / HKCU=0(拒否) → 1 が採用され、制御できる
'      3. リスト系(Allowlist/Blocklist)も合算されず、勝った側のリストがそのまま使われる
'
'  ■ このクラスの判定範囲
'    ・読むのは HKLM / HKCU の SOFTWARE\Policies\<ブラウザ> のみ（GPO・Intune 由来はここに出る）
'    ・クラウド配信のポリシー(Edge管理サービス等)は考慮しない
'    ・そのため「レジストリに無い」＝「制限なし」とは限らない
'      → 最終的な判定はブラウザ側の応答（「Not allowed」等）で行う
'    ・正確な適用状況は edge://policy（chrome://policy）で確認できる。競合も表示されるのでなにかおかしい場合はまず、ここを見ることを推奨します
'
'  ■ WebView2 について
'    WebView2 はブラウザ版 Edge のポリシー対象外のため、上記の制限を受けないことがある
'
'  ■ 制約について
'    初期ブラウザ起動時のみこのClassを利用します。`ReattachOptions`による起動済みブラウザからの制御の場合は対象外です。
'
'  参考：https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/remotedebuggingallowed
'        https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailability
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
'ブラウザパス、ポリシー確認用API
Private Declare PtrSafe Function RegGetValueW Lib "Advapi32" ( _
    ByVal hKey As LongPtr, _
    ByVal lpSubKey As LongPtr, _
    ByVal lpValue As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef pdwType As Long, _
    ByVal pvData As LongPtr, _
    ByRef pcbData As Long _
) As Long

'リスト系ポリシー(サブキー配下の"1","2"...)の列挙用API
Private Declare PtrSafe Function RegOpenKeyExW Lib "Advapi32" ( _
    ByVal hKey As LongPtr, _
    ByVal lpSubKey As LongPtr, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    ByRef phkResult As LongPtr _
) As Long

Private Declare PtrSafe Function RegEnumValueW Lib "Advapi32" ( _
    ByVal hKey As LongPtr, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As LongPtr, _
    ByRef lpcchValueName As Long, _
    ByVal lpReserved As LongPtr, _
    ByRef lpType As Long, _
    ByVal lpData As LongPtr, _
    ByRef lpcbData As Long _
) As Long

Private Declare PtrSafe Function RegCloseKey Lib "Advapi32" ( _
    ByVal hKey As LongPtr _
) As Long

'DJP 20241105 to solve problems with terminating sessions on terminal server
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" ( _
    ) As Long

Private Declare PtrSafe Function ProcessIdToSessionId Lib "kernel32" ( _
    ByVal dwProcessId As Long, _
    ByRef pSessionId As Long) As Long

Private Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As LongPtr



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'起動する Chromium 系ブラウザ
Private Enum BrowserList
    '未設定用
    NoPreset = -1

    ' Edge 系
    RunEdge_Stable
    RunEdge_Beta
    RunEdge_Dev
    RunEdge_Canary

    ' Chrome 系
    RunChrome_Stable
    RunChrome_Beta
    RunChrome_Dev
    RunChrome_Canary
End Enum

'===============レジストリ関連の定数===============
' ルートキー
Private Const HKEY_LOCAL_MACHINE    As Long = &H80000002    'このPCの、"すべて"のユーザー
Private Const HKEY_CURRENT_USER     As Long = &H80000001    '"今、ログインしているユーザー"だけ
' 戻り値の型（RegGetValueW の dwFlags に指定する制限）
Private Const RRF_RT_REG_SZ         As Long = &H2           ' 文字列 (REG_SZ) に限定
Private Const RRF_RT_REG_DWORD      As Long = &H10          ' 32ビット数値 (REG_DWORD) に限定
' 型サイズ
Private Const SIZE_REG_DWORD        As Long = 4             ' DWORD は 4バイト (32bit) と決まっている
' APIの戻り値
Private Const RRF_SUCCESS           As Long = 0             '成功サイン
Private Const ERROR_MORE_DATA       As Long = 234           'バッファー不足（必要サイズが返される）
' レジストリキーのアクセス権（RegOpenKeyExW 用）
Private Const KEY_READ              As Long = &H20019       '読み取り専用
' 値の型（RegEnumValueW で返される種別）
Private Const REG_SZ                As Long = 1             '文字列
Private Const REG_EXPAND_SZ         As Long = 2             '環境変数を含む文字列
' 値の名前の最大文字数（終端の NULL を含む）
Private Const MAX_VALUE_NAME        As Long = 16384
'==================================================

'OpenProcess 用アクセス権（TerminateProcess / 終了待ち / 生存確認に必要な最小セット）
Private Const PROCESS_TERMINATE                 As Long = &H1&
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
Private Const SYNCHRONIZE                       As Long = &H100000

'スプラッシュ起動画面のタイトル部分 ※`--app`で起動時のみ
Private Const SplashTitle               As String = "Starting CDP control from VBA..."

'その他
Private Const ThisClassName             As String = "CDPHost"               'ログのトレース用
Private Const WebSocketInfoFileName     As String = "DevToolsActivePort"    'アクティブなデバックポート情報が記載されてるファイル名
Private Const DebugModeSign             As String = "remote-debugging"      'デバッグブラウザサイン引数



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'---------------- ポリシー状況一覧 ----------------
Private Type StatePolicy
    Contents    As Variant  '中身の値。※存在しない場合は、Emptyとします。
    isHKLM      As Boolean  'マシン単位としてのポリシーか？
End Type

Private Type CheckPolicy
    UserDataDir                         As StatePolicy  'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/userdatadir
    RemoteDebuggingAllowed              As StatePolicy  'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/remotedebuggingallowed
    DeveloperToolsAvailability          As StatePolicy  'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailability
    DeveloperToolsAvailabilityAllowlist As StatePolicy  'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailabilityallowlist
    DeveloperToolsAvailabilityBlocklist As StatePolicy  'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailabilityblocklist
End Type
Private currentPolicy As CheckPolicy
Private checkedPolicy As Boolean
'--------------------------------------------------

'取得するデバックブラウザ設定
Private UserBrowserMode         As BrowserList
Private m_CustomChromiumPath    As String  'カスタムブラウザのexeフルパス。空ならプリセットに従う
Private DebugUserName           As String
Private PipeMode                As Boolean

'ログ識別ID
Private m_LogID As String



'***************************************************************************************************
'                       ■■■ Enum → 文字列 変換プロシージャ ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] ブラウザ種別をexe名で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringBrowserList_exeName(param As BrowserList) As String
    Select Case param
        Case RunChrome_Stable, RunChrome_Beta, RunChrome_Dev, RunChrome_Canary: EnumToStringBrowserList_exeName = "chrome.exe"
        Case RunEdge_Stable, RunEdge_Beta, RunEdge_Dev, RunEdge_Canary:         EnumToStringBrowserList_exeName = "msedge.exe"
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ポリシー情報のあるブラウザ種別を相対レジストリキーパス名で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringBrowserList_RegPath(param As BrowserList) As String
    Select Case param
        Case RunChrome_Stable, RunChrome_Beta, RunChrome_Dev, RunChrome_Canary: EnumToStringBrowserList_RegPath = "\Google\Chrome"
        Case RunEdge_Stable, RunEdge_Beta, RunEdge_Dev, RunEdge_Canary:         EnumToStringBrowserList_RegPath = "\Microsoft\Edge"
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 開発レベル×ブラウザ種別を相対フォルダパス名で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringBrowserList_BrowserPath(param As BrowserList) As String
    Select Case param
        'Chome
        Case RunChrome_Stable:  EnumToStringBrowserList_BrowserPath = "\Google\Chrome"
        Case RunChrome_Beta:    EnumToStringBrowserList_BrowserPath = "\Google\Chrome Beta"
        Case RunChrome_Dev:     EnumToStringBrowserList_BrowserPath = "\Google\Chrome Dev"
        Case RunChrome_Canary:  EnumToStringBrowserList_BrowserPath = "\Google\Chrome SxS"

        'Edge
        Case RunEdge_Stable:    EnumToStringBrowserList_BrowserPath = "\Microsoft\Edge"
        Case RunEdge_Beta:      EnumToStringBrowserList_BrowserPath = "\Microsoft\Edge Beta"
        Case RunEdge_Dev:       EnumToStringBrowserList_BrowserPath = "\Microsoft\Edge Dev"
        Case RunEdge_Canary:    EnumToStringBrowserList_BrowserPath = "\Microsoft\Edge SxS"

        'カスタムパス系
        Case Else:              EnumToStringBrowserList_BrowserPath = "\Microsoft\Office\" & Application.Version & "\Wef\Chromium"  'Officeのワークスペースフォルダを流用
    End Select
End Function



'***************************************************************************************************
'                                   ■■■ 起動前準備 ■■■
'***************************************************************************************************
'* 機能　　：このClassが機能する必要な情報を渡します
'---------------------------------------------------------------------------------------------------
'* 引数　　：caps       ブラウザ起動時のオプション設定オブジェクト
'            UsePipe    Port経由で制御する場合は`False`
'***************************************************************************************************
Friend Sub Init(caps As BrowserCapabilities, UsePipe As Boolean)
    '必要な情報のみ、取り出す
    With caps
        DebugUserName = .Set_UserDataDirName
        m_CustomChromiumPath = .Set_CustomChromiumPath
        UserBrowserMode = .Set_UseBrowserPreset
    End With

    '固有情報
    PipeMode = UsePipe
End Sub

'***************************************************************************************************
'* 機能　　：新しいデバッグブラウザセッションを開始する前にクリーンアップを行います
'---------------------------------------------------------------------------------------------------
'* 機能説明：・基本は対象ブラウザの`remote-debugging-pipe/port`プロセスのみを終了します
'            ・コマンドライン引数に加えて、ターミナルサーバー上で (このユーザーに属さない) セッションを閉じる際の問題も解決しています
'***************************************************************************************************
Public Sub cleanUpSessions()
    Const FromProcedureName As String = ThisClassName & ".cleanUpSessions"


    '1. 今、Excelが動いてるプロセスIDとそのログインユーザーのセッションIDを取得します
    Dim lProcessId As Long, lSessionId As Long
    lProcessId = GetCurrentProcessId()          'Get the current(Excel) process ID
    ProcessIdToSessionId lProcessId, lSessionId 'Get the Session ID for the current(Excel) process

    ' 2. exe名を取得
    Dim exeName As String, tmp
    tmp = Split(getBrowserPath, "\")
    exeName = tmp(UBound(tmp))

    ' 3. 起動中のプロセスを巡る
    Dim proc As Object
    On Error Resume Next
    For Each proc In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process Where SessionId = " & lSessionId)
        '4. 起動中のブラウザ名であるか確認
        If LCase$(proc.Name) = LCase$(exeName) Then
            '5. Killタイプ判定
            If Not IsEmpty(currentPolicy.UserDataDir.Contents) Then
                '6. デバッグブラウザ(CdpSessions)であるか問わず、通常ブラウザもKillします
                printMsg Debug_, "The `UserDataDir` policy is in effect, so all running browser instances will be terminated..." & proc.CommandLine, FromProcedureName
                proc.Terminate

            ElseIf InStr(1, proc.CommandLine, DebugModeSign) > 0 And InStr(1, proc.CommandLine, GenerateArgUserDataDir) > 0 Then
                '6. デバッグブラウザ(CdpSessions)であるかつ、対象ユーザー名なら、Killします
                printMsg Debug_, "Detect debug browser: " & proc.CommandLine, FromProcedureName
                printMsg Debug_, "I'm going to kill that browser.", FromProcedureName
                proc.Terminate
            End If
        End If
        DoEvents
    Next
    On Error GoTo 0
    Err.Clear

    printMsg info_, "cleanUpSessions completed and previous CdpSessionID removed", FromProcedureName

End Sub

Public Sub setCrashStateNormal()
'------------------------------------------------------------------
' Set Crashed state to Normal to prevent the "Restore Sessions?"
' popup from appearing every time the browser is started after
' being closed by browser.quit method.
' Requires user data dir profile stored in private var DebugUserName and
' the browser type stored in private variable br.Type.
' Update: 22/01/26 - Tweaked to work well with the new multi-instance
'                    structure.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".setCrashStateNormal"


    Dim filePath_Preferences As String
    Dim strContents As String, strNewContents As String
    Dim intF As Integer

    '状況に応じたクラッシュパス情報を指定する
    '`UserDataDir`のポリシーがあったらそっちを優先
    If Not IsEmpty(currentPolicy.UserDataDir.Contents) Then
        'Use this switch instead if your organization disables UserDataDir with policy
        filePath_Preferences = currentPolicy.UserDataDir.Contents & "\" & DebugUserName & "\Preferences"
    Else
        'user-data-dir
        filePath_Preferences = getUserDataDirPath & "\Default\Preferences"
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

'***************************************************************************************************
'* 機能　　：利用するブラウザパスを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：ブラウザパス
'---------------------------------------------------------------------------------------------------
'* 注意事項：カスタムパスルートの場合は、exe名を含むフルパスを前提とします
'***************************************************************************************************
Property Get getBrowserPath() As String
    Const FromProcedureName As String = ThisClassName & ".getBrowserPath"


    '1. カスタムパスの指定があったらそっちを最優先
    If LenB(m_CustomChromiumPath) Then
        'そのカスタムパスとして返却(ポータブル版ブラウザ等)
        getBrowserPath = m_CustomChromiumPath

        '以降の処理は不要なので、抜ける
        Exit Property
    End If

    '2. ブラウザexe名を取得
    Dim BrowserExeName As String
    BrowserExeName = EnumToStringBrowserList_exeName(UserBrowserMode)

    '3. インストール先の相対パスを取得
    Dim subPath As String
    subPath = EnumToStringBrowserList_BrowserPath(UserBrowserMode)

    '4. 開発レベルに応じたブラウザ親パス生成
    Dim envLocal As String: envLocal = Environ$("LOCALAPPDATA")
    Dim envProg As String:  envProg = Environ$("ProgramW6432")      '64bit の Excel でも 32bit の Excelでも同じパス(C:\Program Files)を返すように
    Dim envX86 As String:   envX86 = Environ$("ProgramFiles(x86)")

    If LenB(envProg) = 0 Then envProg = Environ$("ProgramFiles")    '取れなかったときに備える(Windows 7 より前には無い変数)

    Select Case UserBrowserMode
        Case RunChrome_Stable, RunEdge_Stable, RunChrome_Beta, RunEdge_Beta, RunChrome_Dev, RunEdge_Dev
            '4-1. まずは「x86」側で確認
            '※基本はこれ
            getBrowserPath = envX86 & subPath & "\Application\" & BrowserExeName
            If LenB(Dir(getBrowserPath)) Then Exit Property  '見つかったらそこで抜ける

            '4-2. つぎに無印側(x64)で確認
            '※最近のChromeはこれ
            getBrowserPath = envProg & subPath & "\Application\" & BrowserExeName
            If LenB(Dir(getBrowserPath)) Then Exit Property  '見つかったらそこで抜ける

            '4-3. 最後に、ユーザー単位
            '※UAC で『いいえ』／管理者権限なし／Edge のポリシー
            getBrowserPath = envLocal & subPath & "\Application\" & BrowserExeName

        Case RunChrome_Canary, RunEdge_Canary
            '4-1. 常に、ユーザー単位
            getBrowserPath = envLocal & subPath & "\Application\" & BrowserExeName
    End Select
End Property

'***************************************************************************************************
'* 機能　　：ブラウザが使用するユーザーパスを返します
'---------------------------------------------------------------------------------------------------
'* 返り値　："user-data-dir"のパス部分
'---------------------------------------------------------------------------------------------------
'* 詳細説明：以下の用途で役立ちます
'            ・`Preferences`内のクラッシュ情報をクリアする
'            ・`--remote-debugging-port=0`で起動した際の、`DevToolsActivePort`を拝見して、自動で空いてるポートに繋げる仕組み
'***************************************************************************************************
Property Get getUserDataDirPath() As String
    If Not IsEmpty(currentPolicy.UserDataDir.Contents) Then
        '`UserDataDir`ポリシーがある場合はそっちを優先に返却
        getUserDataDirPath = currentPolicy.UserDataDir.Contents
    ElseIf InStr(1, DebugUserName, ":") = 2 Then
        '絶対パス指定の場合は、そのまんま返却
        getUserDataDirPath = DebugUserName
    Else
        'ブラウザタイプに応じた標準パス先で返却
        getUserDataDirPath = Environ("LOCALAPPDATA") & EnumToStringBrowserList_BrowserPath(UserBrowserMode) & "\" & DebugUserName
    End If
End Property

'***************************************************************************************************
'* 機能　　：プロセスIDから、操作可能なプロセスハンドル(`hProcess`)を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功時は `hProcess`。失敗時は `0`
'* 引数　　：ProcessId      `Shell` 関数や `CreateProcess` が返すプロセスID
'---------------------------------------------------------------------------------------------------
'* 詳細説明：VBA の `Shell` は PID しか返さないため、`TerminateProcess` 用のハンドルはこちらで開きます
'* 注意事項：取得したハンドルは呼び出し側が `CloseHandle` すること
'***************************************************************************************************
Public Function OpenProcessHandle(ProcessId As LongPtr) As LongPtr
    Const FromProcedureName As String = ThisClassName & ".OpenProcessHandle"


    If ProcessId = 0 Then Exit Function

    OpenProcessHandle = OpenProcess(PROCESS_TERMINATE Or PROCESS_QUERY_LIMITED_INFORMATION Or SYNCHRONIZE, 0, CLng(ProcessId))
    If OpenProcessHandle = 0 Then
        Dim errNo As Long: errNo = Err.LastDllError
        printMsg WARN_, "OpenProcess failed. PID:" & ProcessId & " ErrorCode:" & errNo & " - " & WinApiError.GetMessage(errNo, "kernel32"), FromProcedureName
    Else
        printMsg Debug_, "OpenProcess succeeded. PID:" & ProcessId & " hProcess:" & OpenProcessHandle, FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザフルパスから、バージョン情報を確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：ブラウザバージョン
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`FileSystemObject`を利用したスマートなバージョン確認手段を用いります。
'            起動せずに確認できるメリットがあります。
'***************************************************************************************************
Property Get getBrowserVersion() As String
    getBrowserVersion = CreateObject("Scripting.FileSystemObject").GetFileVersion(getBrowserPath)
End Property



'***************************************************************************************************
'                            ■■■ DevToolsActivePort 関連 ■■■
'***************************************************************************************************
'* 機能　　：ユーザーデータフォルダ直下にある「DevToolsActivePort」ファイルを削除します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：古いWebSocket情報を見ないようにする対策です
'* 注意事項：Chromium起動前に実行してください
'***************************************************************************************************
Public Sub DeleteFileDevToolsActivePort()
    Dim DevToolsActivePortFilePath As String: DevToolsActivePortFilePath = getUserDataDirPath & "\" & WebSocketInfoFileName
    If LenB(Dir(DevToolsActivePortFilePath)) Then Kill DevToolsActivePortFilePath
End Sub

'***************************************************************************************************
'* 機能　　：ユーザーデータフォルダ直下にある「DevToolsActivePort」ファイルが出現するまで待機します
'---------------------------------------------------------------------------------------------------
'* 返り値　：指定Tickまで待機後、あったら`True`となります
'* 引数　　：MaxTickWait    待機Tick数(デフォルトは、10sです。)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：指定Tickまで待機します
'* 注意事項：・必要に応じて、前述のプロシージャで削除を行います
'            ・Chromium起動後に実行して下さい
'***************************************************************************************************
Public Function AutoWaitExistsFileDevToolsActivePort(Optional MaxTickWait As Long = 100) As Boolean
    Const FromProcedureName As String = ThisClassName & ".AutoWaitExistsFileDevToolsActivePort"


    '1. すである場合は、ここの段階で抜ける
    Dim DevToolsActivePortFilePath As String: DevToolsActivePortFilePath = getUserDataDirPath & "\" & WebSocketInfoFileName
    If LenB(Dir(DevToolsActivePortFilePath)) Then AutoWaitExistsFileDevToolsActivePort = True: Exit Function

    Dim TickCount As Long
    Do
        '2. 少し待つ
        printMsg Debug_, "Waiting for the DevToolsActivePort file to appear...", FromProcedureName
        TickCount = TickCount + 1
        CDPHelpers.Sleep 0.1

        '3. 存在判定
        If LenB(Dir(DevToolsActivePortFilePath)) Then AutoWaitExistsFileDevToolsActivePort = True: Exit Do

    Loop Until TickCount = MaxTickWait
End Function

'***************************************************************************************************
'* 機能　　：`DevToolsActivePort`ファイルから、WebSocket接続情報を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：下記の構成となる1次元配列
'            要素0：ポート番号
'            要素1：WebSocketパス
'---------------------------------------------------------------------------------------------------
'* 注意事項：ここでは、ファイル存在checkはしません。前述のプロシージャ側で行って下さい
'***************************************************************************************************
Public Function AnalysisDevToolsActivePort() As Variant
    '1. ファイルOpen
    Dim DevToolsActivePortFilePath As String: DevToolsActivePortFilePath = getUserDataDirPath & "\" & WebSocketInfoFileName
    Dim intFree As Integer
    intFree = FreeFile
    Open DevToolsActivePortFilePath For Input As #intFree

    '2. 中身を全部読む
    Dim tmp
    tmp = Input(LOF(intFree), #intFree)

    '3. 閉じる
    Close #intFree

    '4. 分割
    AnalysisDevToolsActivePort = Split(tmp, vbLf)
End Function



'***************************************************************************************************
'                                 ■■■ 起動引数生成アシスト ■■■
'***************************************************************************************************
'* 機能　　：デバッグ引数を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`--remote-debugging-...`として返却
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`InitInfo`の第三引数に応じて変わります
'***************************************************************************************************
Public Function GenerateArgRemoteDebugging() As String
    GenerateArgRemoteDebugging = "--" & DebugModeSign & IIf(PipeMode, "-pipe", "-port=0")
End Function

'***************************************************************************************************
'* 機能　　：`user-data-dir`引数生成を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       生成フォルダ名
'            BrowserType    ブラウザタイプ
'
'* 返り値　：`--user-data-dir="..."`として返却
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ブラウザに合わせたパスと同じ場所に配置します
'* 注意事項：デバッグ時はこれが必須です。https://developer.chrome.com/blog/remote-debugging-port
'***************************************************************************************************
Private Function GenerateArgUserDataDir() As String
    Const ArgName As String = "user-data-dir"


    GenerateArgUserDataDir = "--" & ArgName & "=""" & getUserDataDirPath & """"
End Function

'***************************************************************************************************
'* 機能　　：`profile-directory`引数生成を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       生成フォルダ名
'            BrowserType    ブラウザタイプ
'
'* 返り値　：`--profile-directory="..."`として返却
'---------------------------------------------------------------------------------------------------
'* 注意事項：`UserDataDir`ポリシーに引っかかる場合は、これを代わりに指定します
'***************************************************************************************************
Private Function GenerateArgProfileDirectory() As String
    Const ArgName As String = "profile-directory"


    GenerateArgProfileDirectory = "--" & ArgName & "=""" & DebugUserName & """"
End Function

'***************************************************************************************************
'* 機能　　：`UserDataDir`ポリシーに応じて、最適なユーザーデータ保存先引数を提供します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`GenerateArgUserDataDir` or `GenerateArgUserDataDir`のいずれか
'---------------------------------------------------------------------------------------------------
'* 注意事項：`CheckRemoteDebuggingBlockedByPolicy`を事前に処理させておくこと
'***************************************************************************************************
Public Function GenerateArgUseUserData() As String
    GenerateArgUseUserData = IIf(IsEmpty(currentPolicy.UserDataDir.Contents), GenerateArgUserDataDir, GenerateArgProfileDirectory)
End Function

'***************************************************************************************************
'* 機能　　：初期起動時の遷移URL引数生成を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：appUrl             初期遷移URL
'            SplashScreenMode   スプラッシュ画面として返却
'* 返り値　：`--app="..."` or `--homepage "about:blank"`として返却
'---------------------------------------------------------------------------------------------------
'* 詳細説明：引数の状況に応じて下記のようになります
'            ・appUrl未指定 ：常に`--homepage "about:blank"`として返却
'            ・URL指定かつ`SplashScreenMode:=True` ：`data:text/html,`を駆使した起動スプラッシュ画面からスタートさせるhtmlコードを渡します。いきなりタブ制御から始める場合は、レースコンディション対策済みのこちらを使ってください
'            ・URL指定かつ`SplashScreenMode:=False`：そのURLへ直行します。ブラウザから起動する場合は、こちらを使ってください
'***************************************************************************************************
Public Function GenerateArgApp(appUrl As String, SplashScreenMode As Boolean) As String
    '状況に応じた分岐
    If LenB(appUrl) Then
        If SplashScreenMode Then
            'かっこいい起動スプラッシュ画面付きで始めるhtmlコードとして渡します
            '- 黒 or 白画面(Windowsのテーマ依存)
            '- どこに遷移するかの案内
            '- リングマーキープログレスバー
            GenerateArgApp = "--app=""" & GetSplashHtmlText(appUrl) & Chr(34)
        Else
            'URLに直行
            GenerateArgApp = "--app=""" & appUrl & Chr(34)
        End If
    Else
        'アドレスバー付きの空ブラウザとして起動
        '※`app`の感覚で、`=`を付けると失敗します
        GenerateArgApp = "--homepage """ & EmptyPageName & Chr(34)
    End If
End Function

'***************************************************************************************************
'* 機能　　：スプラッシュ起動画面を表す簡易的な`data:text/html`形式を返します
'---------------------------------------------------------------------------------------------------
'* 引数　　：appUrl     初期遷移URL
'* 返り値　：`data:text/html`形式
'---------------------------------------------------------------------------------------------------
'* 詳細説明：以下の機能を備えます
'            ・黒 or 白画面(Windowsのテーマ依存)
'            ・どこに遷移するかの案内
'            ・リングマーキープログレスバー
'***************************************************************************************************
Property Get GetSplashHtmlText(appUrl As String) As String
    '1. appUrl の中の「?」を「%3F」に置き換えて、外側Chromeのクエリ判定を回避する！
    Dim safeUrl As String
    safeUrl = Replace(appUrl, "?", "%3F")
    safeUrl = Replace(safeUrl, "#", "%23") ' ついでにハッシュ記号も安全化

    '2. 実際に使うURLを返却
    GetSplashHtmlText = "data:text/html,<title>" & SplashTitle & "</title>" & _
                                    "<style>" & _
                                    ":root{--bg:white;--col:rgb(34,34,34);--rb:rgba(0,0,0,0.1)}" & _
                                    "@media(prefers-color-scheme:dark){:root{--bg:rgb(34,34,34);--col:white;--rb:rgba(255,255,255,0.2)}}" & _
                                    "@keyframes s{to{transform:rotate(360deg)}}" & _
                                    "body{background:var(--bg);color:var(--col);display:flex;flex-direction:column;justify-content:center;align-items:center;height:90vh;font-family:sans-serif;margin:0;}" & _
                                    ".ring{width:40px;height:40px;border:4px solid var(--rb);border-top:4px solid dodgerblue;border-radius:999px;animation:s 1s linear infinite;margin-bottom:15px;}" & _
                                    "</style>" & _
                                    "<body><div class='ring'></div><h2>navigating... `" & safeUrl & "`</h2></body>"
End Property



'***************************************************************************************************
'                                   ■■■ ポリシー情報 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] CDP系ポリシーの取り出し処理系
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：リモート デバッグが出来るかの確認を行います
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・企業によっては、ポリシー設定でリモートデバッグ自体が、できないようにしています。
'　　　　　　・設定されていると、起動はできたのに、なんか操作できないといった混乱が発生するため、このチェックを設けています
'            ・ここのプロシージャでは特に、入口チェックをし、そもそもデバックできるかのチェックをし、事前に警告/停止します。具体的には下記です
'               ・RemoteDebuggingAllowed
'               ・UserDataDir
'* 注意事項：・メジャーなブラウザ(Edge,Chrome)のみ対応します
'            ・他のポリシー(DeveloperToolsAvailability等)は、別プロシージャにてチェックします。あくまでも、入口に関わるチェックのみです
'***************************************************************************************************
Public Sub CheckRemoteDebuggingBlockedByPolicy()
    Const FromProcedureName As String = ThisClassName & ".CheckRemoteDebuggingBlockedByPolicy"


    '1. メジャーなブラウザ以外は即抜け
    If UserBrowserMode = NoPreset Then printMsg WARN_, "The policy check will be skipped because a browser that does not support the policy has been detected.", FromProcedureName: Exit Sub
    checkedPolicy = True

    '2. 入口系のデバックポリシー状況を取得していく
    '※(RemoteDebuggingAllowed + UserDataDir)×(HKLM + HKCU)の4パターン
    Dim RemoteDebuggingAllowed_HKLM, RemoteDebuggingAllowed_HKCU
    RemoteDebuggingAllowed_HKLM = CheckPolicy_RemoteDebuggingAllowed(False)
    RemoteDebuggingAllowed_HKCU = CheckPolicy_RemoteDebuggingAllowed(True)

    Dim UserDataDir_HKLM, UserDataDir_HKCU
    UserDataDir_HKLM = CheckPolicy_UserDataDir(False)
    UserDataDir_HKCU = CheckPolicy_UserDataDir(True)

    '3. 優先順位を考慮した記録をする
    printMsg info_, "Checking policy... RemoteDebuggingAllowed", FromProcedureName
    AchieveResultsPolicy currentPolicy.RemoteDebuggingAllowed, RemoteDebuggingAllowed_HKLM, RemoteDebuggingAllowed_HKCU

    printMsg info_, "Checking policy... UserDataDir", FromProcedureName
    AchieveResultsPolicy currentPolicy.UserDataDir, UserDataDir_HKLM, UserDataDir_HKCU

    '4. グローバル変数を基に、制御可否判定を行う
    '4-1. リモートデバックそのものが封じられてるか？
    Dim ErrorMes As String
    If Not IsEmpty(currentPolicy.RemoteDebuggingAllowed.Contents) Then
        '拒否値か？
        If currentPolicy.RemoteDebuggingAllowed.Contents = 0 Then
            'ポリシー場所に応じたエラーメッセージにし、止める
            ErrorMes = IIf(currentPolicy.RemoteDebuggingAllowed.isHKLM, "会社のセキュリティポリシーにより、ブラウザのリモートデバッグが、ブロックされています。" & vbCrLf & vbCrLf & "CDP機能は、利用できません。" & vbCrLf & "情報システム部門に、お問い合わせするか、WebView2ルートでの制御をお試しください。", _
                                                                        "現在のログインPCアカウントでは、ブラウザのリモートデバッグが、ブロックされています。" & vbCrLf & vbCrLf & "CDP機能は、利用できません。" & vbCrLf & "これが個人のPCアカウントの場合は、レジストリ操作にて`1`にして下さい。")

            printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True
            Exit Sub
        End If
    End If

    '4-2. ブラウザ標準の`UserDataDir`になってしまってないか？
    If Not IsEmpty(currentPolicy.UserDataDir.Contents) Then
        '実質、拒否値か？
        If currentPolicy.UserDataDir.Contents = "" Then
            'ポリシー場所に応じたエラーメッセージにし、止める
            ErrorMes = IIf(currentPolicy.UserDataDir.isHKLM, "会社のセキュリティポリシーにより、ブラウザデータの配置場所がデフォルトパスに固定されています。" & vbCrLf & vbCrLf & "Chromiumの仕様上、デフォルトパスのデバッグはできません。" & vbCrLf & "情報システム部門、お問い合わせするか、WebView2ルートでの制御をお試しください。", _
                                                             "現在のログインPCアカウントでは、ブラウザデータの配置場所がデフォルトパスに固定されています。" & vbCrLf & vbCrLf & "Chromiumの仕様上、デフォルトパスのデバッグはできません。" & vbCrLf & "これが個人のPCアカウントの場合は、レジストリ操作にて削除するか、何かパスを指定して下さい。")

            printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True
            Exit Sub
        Else
            'デバックがし辛くなるだけなので、警告のみに留める
            ErrorMes = IIf(currentPolicy.UserDataDir.isHKLM, "会社のセキュリティポリシーにより、ブラウザデータの配置場所が`" & currentPolicy.UserDataDir.Contents & "`に固定されています。起動中のブラウザを全て落とします。", _
                                                             "現在のログインPCアカウントでは、ブラウザデータの配置場所が`" & currentPolicy.UserDataDir.Contents & "`に固定されています。起動中のブラウザを全て落とします。これが個人のPCアカウントであり、この仕様を回避する場合は、レジストリ操作にて削除して下さい。")

            printMsg WARN_, ErrorMes, FromProcedureName
        End If
    End If

    '5. デバックに支障が出る系の情報取り出しを行う
    CheckRemoteDebuggingAvailabilityByPolicy
End Sub

'***************************************************************************************************
'* 機能　　：リモート デバッグに支障をきたす系の情報を確認します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・企業によっては、ポリシー設定で一部のURLに対してリモートデバッグが、できないようにしています。
'　　　　　　・設定されていると、デバックはできるのに、ブロック先のURLに行くと、謎の「Not allowed」というあいまいエラーで???になるので、予めブロックリストを取ってあとで、失敗時にログに出すような仕組みを取ってます
'
'* 注意事項：・メジャーなブラウザ(Edge,Chrome)のみ対応します
'            ・現時点では、「CDPBrowser-attachToTab」の判定側でのみ使用します。
'            ・値の取り出しは別途、Getpropertyで取ってください
'***************************************************************************************************
Private Sub CheckRemoteDebuggingAvailabilityByPolicy()
    Const FromProcedureName As String = ThisClassName & ".CheckRemoteDebuggingAvailabilityByPolicy"


    '1. 動的系のデバックポリシー状況を取得していく
    Dim DeveloperToolsAvailability_HKLM, DeveloperToolsAvailability_HKCU
    DeveloperToolsAvailability_HKLM = CheckPolicy_DeveloperToolsAvailability(False)
    DeveloperToolsAvailability_HKCU = CheckPolicy_DeveloperToolsAvailability(True)

    Dim DeveloperToolsAvailabilityBlocklist_HKLM, DeveloperToolsAvailabilityBlocklist_HKCU
    DeveloperToolsAvailabilityBlocklist_HKLM = CheckPolicy_DeveloperToolsAvailabilityList(False, False)
    DeveloperToolsAvailabilityBlocklist_HKCU = CheckPolicy_DeveloperToolsAvailabilityList(True, False)

    Dim DeveloperToolsAvailabilityAllowlist_HKLM, DeveloperToolsAvailabilityAllowlist_HKCU
    DeveloperToolsAvailabilityAllowlist_HKLM = CheckPolicy_DeveloperToolsAvailabilityList(False, True)
    DeveloperToolsAvailabilityAllowlist_HKCU = CheckPolicy_DeveloperToolsAvailabilityList(True, True)

    '2. 優先順位を考慮した記録をする
    printMsg info_, "Checking policy... DeveloperToolsAvailability", FromProcedureName
    AchieveResultsPolicy currentPolicy.DeveloperToolsAvailability, DeveloperToolsAvailability_HKLM, DeveloperToolsAvailability_HKCU

    printMsg info_, "Checking policy... DeveloperToolsAvailabilityBlocklist", FromProcedureName
    AchieveResultsPolicy currentPolicy.DeveloperToolsAvailabilityBlocklist, DeveloperToolsAvailabilityBlocklist_HKLM, DeveloperToolsAvailabilityBlocklist_HKCU

    printMsg info_, "Checking policy... DeveloperToolsAvailabilityAllowlist", FromProcedureName
    AchieveResultsPolicy currentPolicy.DeveloperToolsAvailabilityAllowlist, DeveloperToolsAvailabilityAllowlist_HKLM, DeveloperToolsAvailabilityAllowlist_HKCU
End Sub

'***************************************************************************************************
'* 機能　　：優先順位を考慮した最終判定を、グローバル変数に残します
'---------------------------------------------------------------------------------------------------
'* 引数　　：LastResult グローバル変数に残すものを指定
'           arg_HKLM    事前に取ったポリシーHKLM値
'           arg_HKCU    事前に取ったポリシーHKCU値
'---------------------------------------------------------------------------------------------------
'* 詳細説明：適用優先は、HKLM > HKCU。HKLM に値があれば HKCU は無視される。拒否優先ではない
'***************************************************************************************************
Private Sub AchieveResultsPolicy(ByRef LastResult As StatePolicy, arg_HKLM As Variant, arg_HKCU As Variant)
    Const FromProcedureName As String = ThisClassName & ".AchieveResultsPolicy"


    '1. `HKLM`側に有効な値があるか？
    If Not IsEmpty(arg_HKLM) Then
        '1-1. `HKLM`側を記録
        LastResult.isHKLM = True
        LastResult.Contents = arg_HKLM
        printMsg info_, "I have reviewed HKLM's policy. value: " & PolicyValueToString(arg_HKLM), FromProcedureName

        '1-2. `HKCU`にもあったら、無視した旨を伝える
        If Not IsEmpty(arg_HKCU) Then printMsg WARN_, "I have also checked HKCU's policy(" & PolicyValueToString(arg_HKCU) & "), but I will prioritize HKLM.", FromProcedureName

    '2. `HKCU`側に有効な値があるか？
    ElseIf Not IsEmpty(arg_HKCU) Then
        '2-1. `HKCU`側を記録
        LastResult.isHKLM = False
        LastResult.Contents = arg_HKCU
        printMsg info_, "I have reviewed HKCU's policy. value: " & PolicyValueToString(arg_HKCU), FromProcedureName
    Else
        '3. ない旨を伝える
        LastResult.isHKLM = False
        LastResult.Contents = Empty
        printMsg info_, "There is no policy.", FromProcedureName
    End If
End Sub

'***************************************************************************************************
'* 機能　　：ポリシー値を、ログ表示用の文字列にします
'---------------------------------------------------------------------------------------------------
'* 引数　　：PolicyValue    ポリシー値（数値/文字列/文字列配列）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：リスト系ポリシー(Allowlist/Blocklist)は配列で持つため、そのまま`&`で連結すると型エラーになります。その対策です
'***************************************************************************************************
Private Function PolicyValueToString(ByVal PolicyValue As Variant) As String
    If IsArray(PolicyValue) Then
        PolicyValueToString = "[" & Join(PolicyValue, ", ") & "]"
    Else
        PolicyValueToString = CStr(PolicyValue)
    End If
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] CDP系ポリシーの各種取り出し処理系
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：`RemoteDebuggingAllowed`の確認を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：useHKCU            True ：HKEY_CURRENT_USER　側を確認
'                               False：HKEY_LOCAL_MACHINE 側を確認
'
'* 返り値　:0以外:リモート デバッグを使用できる
'           0    :リモート デバッグを使用できません
'           Empty:存在しない(※IsEmptyで判定可)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/remotedebuggingallowed
'***************************************************************************************************
Private Function CheckPolicy_RemoteDebuggingAllowed(useHKCU As Boolean) As Variant
    Const checkValueName As String = "RemoteDebuggingAllowed"


    '1. 必要な変数を用意
    Dim KeyExists       As Long     '存在有無結果
    Dim ValueContents   As Long     '値の中身
    Dim ValueSize       As Long     '値の大きさ
    Dim TargetBrowser   As String: TargetBrowser = "SOFTWARE\Policies" & EnumToStringBrowserList_RegPath(UserBrowserMode)

    '2. ルートキーをセット
    Dim RootTree As Long
    RootTree = IIf(useHKCU, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE)

    '3. 32bit値としてチェックする
    ValueSize = SIZE_REG_DWORD
    KeyExists = RegGetValueW(RootTree, StrPtr(TargetBrowser), StrPtr(checkValueName), RRF_RT_REG_DWORD, 0, VarPtr(ValueContents), ValueSize)

    '4. そのキー名があるか？
    If KeyExists = RRF_SUCCESS Then CheckPolicy_RemoteDebuggingAllowed = ValueContents
End Function

'***************************************************************************************************
'* 機能　　：`UserDataDir`の確認を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：useHKCU            True ：HKEY_CURRENT_USER　側を確認
'                               False：HKEY_LOCAL_MACHINE 側を確認
'
'* 返り値　：パスが返ったらそれが固定の保存場所です。なお、空文字パターンもあり得ます。
''           Empty:存在しない(※IsEmptyで判定可)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/userdatadir
'***************************************************************************************************
Private Function CheckPolicy_UserDataDir(useHKCU As Boolean) As Variant
    Const checkValueName As String = "UserDataDir"


    '1. 必要な変数を用意
    Dim KeyExists       As Long     '存在有無結果
    Dim ValueSize       As Long     '値の大きさ
    Dim TargetBrowser   As String: TargetBrowser = "SOFTWARE\Policies" & EnumToStringBrowserList_RegPath(UserBrowserMode)

    '2. ルートキーをセット
    Dim RootTree As Long
    RootTree = IIf(useHKCU, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE)

    '3. 文字列なのでまずは文字数を取得
    KeyExists = RegGetValueW(RootTree, StrPtr(TargetBrowser), StrPtr(checkValueName), RRF_RT_REG_SZ, 0, 0, ValueSize)

    '4. そのキー名があるか？
    If KeyExists = RRF_SUCCESS Then
        '5. 得た文字数を基に、バッファーを確保しつつ、末尾の`vbNullChar`も落とすように調整する
        Dim tmp As String
        tmp = String(ValueSize / 2 - 2, vbNullChar)

        '6. 返却
        RegGetValueW RootTree, StrPtr(TargetBrowser), StrPtr(checkValueName), RRF_RT_REG_SZ, 0, StrPtr(tmp), ValueSize
        CheckPolicy_UserDataDir = tmp
    End If
End Function

'***************************************************************************************************
'* 機能　　：`DeveloperToolsAvailability`の確認を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：useHKCU            True ：HKEY_CURRENT_USER　側を確認
'                               False：HKEY_LOCAL_MACHINE 側を確認
'
'* 返り値　：0    :ポリシーで強制インストールされた拡張機能のみ、DevTools禁止（未設定時の既定と同じ）
'           1    :DevToolsをすべて許可
'           2    :DevToolsをすべて禁止
'           Empty:存在しない(※IsEmptyで判定可)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailability
'***************************************************************************************************
Private Function CheckPolicy_DeveloperToolsAvailability(useHKCU As Boolean) As Variant
    Const checkValueName As String = "DeveloperToolsAvailability"


    '1. 必要な変数を用意
    Dim KeyExists       As Long     '存在有無結果
    Dim ValueContents   As Long     '値の中身
    Dim ValueSize       As Long     '値の大きさ
    Dim TargetBrowser   As String: TargetBrowser = "SOFTWARE\Policies" & EnumToStringBrowserList_RegPath(UserBrowserMode)

    '2. ルートキーをセット
    Dim RootTree As Long
    RootTree = IIf(useHKCU, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE)

    '3. 32bit値としてチェックする
    ValueSize = SIZE_REG_DWORD
    KeyExists = RegGetValueW(RootTree, StrPtr(TargetBrowser), StrPtr(checkValueName), RRF_RT_REG_DWORD, 0, VarPtr(ValueContents), ValueSize)

    '4. そのキー名があるか？
    If KeyExists = RRF_SUCCESS Then CheckPolicy_DeveloperToolsAvailability = ValueContents
End Function

'***************************************************************************************************
'* 機能　　：`DeveloperToolsAvailabilityAllowlist/Blocklist`の確認を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：useHKCU            True ：HKEY_CURRENT_USER　側を確認
'                               False：HKEY_LOCAL_MACHINE 側を確認
'           useAllowlist       True ：`DeveloperToolsAvailabilityAllowlist` を確認
'                               False：`DeveloperToolsAvailabilityBlocklist` を確認
'
'* 返り値　：URLパターンの文字列配列(0始まり)
'           Empty:存在しない、または有効な値が1件もない(※IsEmptyで判定可)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・リスト系ポリシーは、サブキー配下に"1","2","3"...という名前の REG_SZ として保存されます
'            ・Chromium(registry_dict.cc)は、名前が数字の値を「すべて」拾い、数字以外の名前は無視します
'              番号が飛んでいても(1,2,5など)拾われるため、連番で読むのではなく全件を列挙しています
'            ・並び順は RegEnumValueW の返却順で、番号順の保証はありません（ログ表示用途のため気にしない）
'            ・同名の REG_SZ に JSON 文字列で書く形式も Chromium は受け付けますが、GPO/Intune は使わないため対象外です
'            https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailabilityallowlist
'            https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailabilityblocklist
'***************************************************************************************************
Private Function CheckPolicy_DeveloperToolsAvailabilityList(useHKCU As Boolean, useAllowlist As Boolean) As Variant
    '1. 必要な変数を用意
    Dim TargetKey As String: TargetKey = "SOFTWARE\Policies" & EnumToStringBrowserList_RegPath(UserBrowserMode) & "\DeveloperToolsAvailability" & IIf(useAllowlist, "Allowlist", "Blocklist")

    '2. ルートキーをセット
    Dim RootTree As Long
    RootTree = IIf(useHKCU, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE)

    '3. サブキーを開く。開けない(=存在しない)なら、ポリシーなしとして抜ける
    Dim hKey As LongPtr
    If RegOpenKeyExW(RootTree, StrPtr(TargetKey), 0, KEY_READ, hKey) <> RRF_SUCCESS Then Exit Function

    '4. 値を全件列挙する
    Dim Items       As Collection: Set Items = New Collection
    Dim NameBuf     As String: NameBuf = String$(MAX_VALUE_NAME, vbNullChar)   '名前は最大長で確保しておくので、名前起因のバッファー不足は起きない
    Dim DataBuf     As String: DataBuf = String$(1024, vbNullChar)             'URLパターン用。足りなければ、後述で拡張する
    Dim Index       As Long
    Dim namelen     As Long
    Dim DataSize    As Long
    Dim ValueType   As Long
    Dim result      As Long
    Dim ValueName   As String
    Dim ValueData   As String
    Dim NullPos     As Long

    Do
        '4-1. 入出力サイズを毎回リセット（API側で書き換えられるため）
        namelen = Len(NameBuf)
        DataSize = LenB(DataBuf)

        '4-2. Index番目の値を取得
        result = RegEnumValueW(hKey, Index, StrPtr(NameBuf), namelen, 0, ValueType, StrPtr(DataBuf), DataSize)

        Select Case result
            Case RRF_SUCCESS
                '名前が数字、かつ文字列型のものだけ拾う（Chromiumと同じ扱い）
                ValueName = Left$(NameBuf, namelen)
                If LenB(ValueName) > 0 And Not ValueName Like "*[!0-9]*" Then
                    If ValueType = REG_SZ Or ValueType = REG_EXPAND_SZ Then
                        'DataSize はバイト数。終端の vbNullChar が含まれる場合があるので、最初の vbNullChar より前を採用
                        ValueData = Left$(DataBuf, DataSize \ 2)
                        NullPos = InStr(ValueData, vbNullChar)
                        If NullPos > 0 Then ValueData = Left$(ValueData, NullPos - 1)
                        Items.Add ValueData
                    End If
                End If

                '次へ
                Index = Index + 1

            Case ERROR_MORE_DATA
                'データ用バッファーが足りないので、必要サイズ(DataSize)で確保し直して、同じIndexを再取得
                '※念のため、サイズが増えない場合は無限ループ防止で、その値は飛ばす
                If DataSize > LenB(DataBuf) Then
                    DataBuf = String$(DataSize \ 2 + 1, vbNullChar)
                Else
                    Index = Index + 1
                End If

            Case Else
                'ERROR_NO_MORE_ITEMS(全件取得済み)、またはその他のエラー
                Exit Do
        End Select
    Loop

    '5. サブキーを閉じる
    RegCloseKey hKey

    '6. 1件もなければ、ポリシーなし(Empty)として返す
    If Items.Count = 0 Then Exit Function

    '7. 文字列配列にして返却
    Dim UrlPatterns() As String
    ReDim UrlPatterns(Items.Count - 1)

    Dim i As Long
    For i = 1 To Items.Count
        UrlPatterns(i - 1) = Items(i)
    Next

    CheckPolicy_DeveloperToolsAvailabilityList = UrlPatterns
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] CDP系ポリシーの後で取り出す系
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：`DeveloperToolsAvailabilityAllowlist`を取り出します
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に、`CheckRemoteDebuggingAvailabilityByPolicy`を行うこと
'***************************************************************************************************
Property Get CurrentDeveloperToolsAvailabilityAllowlist() As Variant
    CurrentDeveloperToolsAvailabilityAllowlist = currentPolicy.DeveloperToolsAvailabilityAllowlist.Contents
End Property

'***************************************************************************************************
'* 機能　　：`DeveloperToolsAvailabilityBlocklist`を取り出します
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に、`CheckRemoteDebuggingAvailabilityByPolicy`を行うこと
'***************************************************************************************************
Property Get CurrentDeveloperToolsAvailabilityBlocklist() As Variant
    CurrentDeveloperToolsAvailabilityBlocklist = currentPolicy.DeveloperToolsAvailabilityBlocklist.Contents
End Property

'***************************************************************************************************
'* 機能　　：`DeveloperToolsAvailability`を取り出します
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に、`CheckRemoteDebuggingAvailabilityByPolicy`を行うこと
'***************************************************************************************************
Property Get CurrentDeveloperToolsAvailability() As Variant
    CurrentDeveloperToolsAvailability = currentPolicy.DeveloperToolsAvailability.Contents
End Property

'***************************************************************************************************
'* 機能　　：ポリシーチェックをしたかの判定を取り出します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：プリセット以外ではポリシーチェックはされないのでその目印にどうぞ
'***************************************************************************************************
Property Get isCheckedPolicy() As Boolean
    isCheckedPolicy = checkedPolicy
End Property



'***************************************************************************************************
'                                     ■■■ ログヘルパー ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, currentLog.logPath
End Sub



'***************************************************************************************************
'                                     ■■■ 初期化 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    'ログID
    m_LogID = "CHID" & Format(Rnd * 1000, "000")
End Sub
