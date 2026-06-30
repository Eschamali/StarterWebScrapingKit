VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'                         ブラウザそのものに関するメソッドやパラメータ運用を行います
'                  このClassオブジェクトは、`--user-data-dir`単位で管理する前提とします
'
'   `ExecuteCDP`等で使用できるコマンド一覧は、右記にてアクセスして下さい → https://chromedevtools.github.io/devtools-protocol/
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
'DJP 20241105 to solve problems with terminating sessions on terminal server
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" ( _
    ) As Long

Private Declare PtrSafe Function ProcessIdToSessionId Lib "kernel32" ( _
    ByVal dwProcessId As Long, _
    ByRef pSessionId As Long) As Long

Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long

'ポリシー情報確認用API(ReadOnly)
Private Declare PtrSafe Function RegGetValueW Lib "Advapi32" ( _
    ByVal hKey As LongPtr, _
    ByVal lpSubKey As LongPtr, _
    ByVal lpValue As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef pdwType As Long, _
    ByVal pvData As LongPtr, _
    ByRef pcbData As Long _
) As Long



'***************************************************************************************************
'                                   ■■■ 各種構造体 ■■■
'***************************************************************************************************
'タブの種類。一旦は2種類のみとする。
'全容→https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/devtools_agent_host_impl.cc?ss=chromium&q=f:devtools%20-f:out%20%22::kTypeTab%5B%5D%22
Public Enum TargetgetTargetsType
    Page
    iFrame
End Enum

'EdgeをWebView2っぽく見せる専用引数
Public Enum edgeKioskType
    none            'キオスクモードなしの通常起動
    fullscreen      'キオスクモード デジタル/対話型サイネージ
    publicBrowsing  'キオスクモード パブリック ブラウジング
End Enum



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'このツール固有情報
Private Const CDPVer                    As String = "3.1.0"
Private Const ArgRemoteDebuggingName    As String = "remote-debugging-pipe"

'各種ブラウザ固有の情報
Private Const brChrome          As String = "chrome.exe"
Private Const brEdge            As String = "msedge.exe"
Private Const chromeWindowClass As String = "Chrome_WidgetWin_1"    'same window class for Edge

'レジストリ関連の定数==============================
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
'==================================================

'その他
Private Const TIMEOUT_DEFAULT   As Double = 30              '初期タイムアウト秒数
Private Const ThisClassName     As String = "CDPBrowser"    'ログのトレース用
Private Const MaxStoreResultCDP As Long = 2 ^ 16            'ブラウザからのコマンド結果蓄積数



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'CDP制御用コアクラスオブジェクト
Private WithEvents PipeCore As CDPCore              'Must be initialized for CDP engine
Attribute PipeCore.VB_VarHelpID = -1

'操作中のブラウザ情報一式
Private Type StateBrowser
    Type        As String   'Browser type: either 'chrome' or 'edge'
    User        As String   'Browser user profile
    Path        As String   'Browser default installation path
    LogID       As String   'The ID of the current CDPBrowser object. Useful when reading debugs of multiple tabs.
    Ver         As String   'The current browser version
    CallString  As String   'The initial call string used by .start
    setShowWindow As WinState   '`remote-debugging-pipe`で起動する際のブラウザ初期表示モードの設定値
End Type
Private br As StateBrowser

'CDPコマンドの結果を蓄積する際の管理(`WithEvents`で直接は取らないようにする仕組み作り)
Private Type cacheResultCDP
    response    As Dictionary   '「ID - それに対するCDPコマンドの結果」というDictionary
    Limit       As Long         'その`Dictionary`に詰め込む上限数
End Type
Private ResultCDP_Browser As cacheResultCDP

'---------------- ポリシー状況一覧 ----------------
Private Type StatePolicy
    Exists      As Boolean  '存在有無
    Contents    As Variant  '中身の値
End Type

Private Type CheckPolicy
    UserDataDir As StatePolicy  'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/userdatadir
End Type
Private currentPolicy As CheckPolicy
Private UseChrome     As Boolean
'--------------------------------------------------

'各種エラー情報
Private lastErrorDetailsCDP     As Dictionary  'CDP-Jsonコマンド実行時の最後のエラー情報
Private lastExceptionDetailsJS  As Dictionary  'JavaScript実行時の最後のエラー情報

'その他
Private thisApp     As Object                   'A dymanic variable used in CDPBrowser initialize procedure to determine if the VBA host is an Excel / Word / Access
Private timerOut    As Double                   'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
Private UseConverterUtf8FromString As Boolean   'ブラウザへ送信する際、UTF-8で使用するかしないか
Private TargetInfos As Collection               '非同期イベント`Target.targetCreated`で得た`Target.TargetInfo`情報。基本的に、`CDPContext.cls - attachToSession`専用
Private DefinitionCDPContext As Long            '確保中の`CDPContext.cls` ※適切なメソッドを呼び出した場合に限る



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する所定の起動引数をつけて起動させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：Name           ブラウザ名。※現時点では、Chrome OR Edge のみ
'            appUrl         起動引数：`--app`　につけるURLを指定します
'            userProfile    起動引数：`--user-data-dir`　につけるユーザーディレクトリ名を指定します
'            addArgs        追加の起動引数
'            KioskMode      Edgeのキオスクモードを使う際にセットします
'---------------------------------------------------------------------------------------------------
'* 機能説明
' This function initiates a new browser instance.
' By default, this function initializes chrome as the default choice.
' To init edge, parse 'edge' to the first parameter (eg. browser.start "edge").
' To open in App Mode, input the starting URL for appURL parameter as app mode cannot be open on a blank tab.
' If cleanActive is False, the method will skip cleaning currently open instances. This will save
' initialization time but if an instance is already running and not closed, pipe error will be triggered.
' If reAttach is True, the method will always try to reattach to the previous automation session if it is
' still active. If not, it will then move on to start a brand new session. If False, it will always open a new
' session regardless whether there is an on going one or not.
' To add additional arguments to the call string, use the addArguments parameter for instance:
' .start addArgs:="--disable-info-bar" (https://peter.sh/experiments/chromium-command-line-switches/)
' Note: not all arguments work the same for Chrome and Edge. Some may work for one but not the others.
' Updated: 22/03/23 Long Vh:      - Enhanced to dynamically retrieve installation path of the browser
'          26/03/23 Long Vh:      - Made some quality-of-life changes
'          25/05/25 Long Vh:      - set default profile to "CDP" so that it can work with the latest chrome update
'          23/10/25 Daniel Polak: - Added missing space before --homepage
'          22/01/26 Long Vh:      - Removed automatic reattachment from the start method as this restricts
'                                   the possibility of running multiple CDP browser instances under different
'                                   profiles - which is a very useful feature.
'                                 - Tweaked the addProfile section and the cleanUpSessions to reflect the
'                                   recent Chromium updates and the Edge policy discovery.
'***************************************************************************************************
Public Sub start(Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional userProfile As String, _
    Optional addArgs As String, Optional KioskMode As edgeKioskType)
    Const FromProcedureName As String = ThisClassName & ".start"


    'パイプ関連の初期化処理
    Set PipeCore = New CDPCore
    PipeCore.UseConverterUtf8FromString = UseConverterUtf8FromString '`New`による直後に、文字コード設定を反映

    'いくつかのポリシーチェックを行います。
    IsRemoteDebuggingBlockedByPolicy IIf(UseChrome, brChrome, brEdge)

    Dim strCall As String, addProfile As String
    Dim intRes As Long
    Dim timeStart As Double

    timeStart = TimerCounter

   'Determine default browser parameters -------------------------------------------------------------------------------------
   '07/01/26: avoid using --enable-automation tag as it triggers many web apps to perform bot check
    Select Case Name
        Case "edge": br.Type = brEdge
        Case "chrome": br.Type = brChrome
        Case Else: printMsg ERROR_, "Browser initiation fails: only the defined browser types can be accepted", FromProcedureName, doRaiseError:=True
    End Select
    br.Path = getBrowserPath(br.Type)
    strCall = """" & br.Path & """ --" & ArgRemoteDebuggingName
   '--------------------------------------------------------------------------------------------------------------------------

   'Add starting URL to the call string -------------------------------------------------------------------------------------------------------------------
    If appUrl <> "" Then
        If appUrl = "about:blank" Then printMsg ERROR_, "Browser cannot be open under Application Mode when homepage is set to ""about:blank""", FromProcedureName, doRaiseError:=True
        strCall = strCall & " --app=""" & checkURI(appUrl) & """"
    Else
        strCall = strCall & " --homepage ""about:blank"""
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------------------------

   'Start with the choice profile -------------------------------------------------------------------------------------------------------------
   '22/01/26: I have recently noted that if an organization sets the group policy of Edge of UserDataDir to a specific structure, the argument --user-data-dir
   'will be effectively disabled and blocked to the specified data dir by the policy. In this scenario, CDP session can only load the browser state from the default
   'data dir which already cannot be allowed to run --remote-debugging-pipe from Google's latest Chromium. If Edge's Chromium updates to the latest, this policy will
   'cause the CDP initialization to fail. If Edge's Chromium is still behind the latest Chromium version, CDP init can still be possible with the --profile-directory
   'switch but the requirement is that all running Edge instances that had been started up manually will need to be closed first otherwise the PeekNamedPipe error
   'will result. For Chrome, v136 and later does not honor remote-debugging request if user profile is loaded from the Chrome\User Data\ folder.
   'As a result, here we set it to Chrome\CDP\ folder instead (auto-created if not exists).
    Select Case br.Type
        Case brEdge: addProfile = IIf(currentPolicy.UserDataDir.Exists, " --profile-directory=""" & userProfile & """", " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Microsoft\Edge\" & userProfile & """")      'ポリシー状況に応じて、引数を変える
        Case brChrome: addProfile = IIf(currentPolicy.UserDataDir.Exists, " --profile-directory=""" & userProfile & """", " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Google\Chrome\" & userProfile & """")     'ポリシー状況に応じて、引数を変える
    End Select
    strCall = strCall & addProfile
    br.User = userProfile 'store the user profile name for cpid serialization
   '-----------------------------------------------------------------------------------------------------------------------------------------

    'Edge特有のキオスクモードを設定します -------------------------------------------------------------------------------------------------------------
    '※Chromeは、Edgeほど調整できないため、同じ引数にします
    Select Case KioskMode
        Case edgeKioskType.fullscreen: strCall = strCall & IIf(br.Type = brChrome, " --kiosk", " --kiosk --edge-kiosk-type=fullscreen")
        Case edgeKioskType.publicBrowsing: strCall = strCall & IIf(br.Type = brChrome, " --kiosk", " --kiosk --edge-kiosk-type=public-browsing")
    End Select
    br.setShowWindow = IIf(KioskMode, WinState.doShowMinNoActivate, br.setShowWindow)   'キオスクモード時は、ウィンドウを非アクティブ化し、最小化されたウィンドウとして表示する優先にします。(https://learn.microsoft.com/ja-jp/windows/win32/api/winuser/nf-winuser-showwindow)
   '-----------------------------------------------------------------------------------------------------------------------------------------

   'Clean up previous sessions if they could not be reattached to avoid PeekNamedPipe error --------------------------------------------
   'Note: The error usually occurs when the same User Profile is already being used by another pipe. Hence if this
   'session is unable to connect to the previous pipe (for instance, due to the session application host was abruptly terminated),
   'PeekNamedPipe error will be thrown.
    If PipeCore.deserialize(br.User) Then PipeCore.CloseHandleBrowser        'Excelテーブルにユーザー名に紐づくハンドル情報が取れちゃったら、新規起動なのでClose処理しておく
    cleanUpSessions onlyCdpSessions:=Not (currentPolicy.UserDataDir.Exists)  'Note: you can set onlyCdpSessions to True by default if your organization does not force UserDataDir policy
    setCrashStateNormal 'Prevent "Restore Session" popup
   '------------------------------------------------------------------------------------------------------------------------------------

   'Prepare for initialization ------------------------------------------------------------------------
    strCall = strCall & " " & addArgs  'Add additional arguments if instructed
   '---------------------------------------------------------------------------------------------------

   'Initialize CDPCore class ----------------------------------------------------------------------------------------------------------
    br.CallString = strCall
    printMsg Debug_, "Starting pipe communication... → " & br.CallString, FromProcedureName
    intRes = PipeCore.InitCDP(br.CallString, br.setShowWindow) 'Serializable is set to True by default for ease of accessibility
    If intRes <> 0 Then printMsg ERROR_, "Unable to initialize CDPCore. ErrorCode:" & intRes & " - " & WinApiError.GetMessage(intRes, "kernel32"), FromProcedureName, doRaiseError:=True Else printMsg info_, "CDPCore initialized successfully", FromProcedureName
   '-----------------------------------------------------------------------------------------------------------------------------------

    'ブラウザセッションパイプ情報を記録
    serializeForPipe

    '起動中のブラウザ情報を変数に記録
    getBrowserInfo

    'ブラウザステータス表示
    printParams
    printMsg info_, "Browser initialization completed after " & (TimerCounter - timeStart) / 1000 & "s", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：`targetCreated/targetInfoChanged/targetDestroyed`の3種の非同期イベントを流しつつ、Context のカウントも行います
'---------------------------------------------------------------------------------------------------
'* SET引数 ：論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：・`CDPContext.cls`の生存状況把握に使います
'            ・うまい具合に、`CDPContext.cls`確保中のカウントも行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：`CDPContext.cls`を生成/破棄時に呼ぶ想定です。
'***************************************************************************************************
Property Let EnableDiscoverTargets(Flag As Boolean)
    Const FromProcedureName As String = ThisClassName & ".EnableDiscoverTargets"


    '1. カウント
    DefinitionCDPContext = DefinitionCDPContext + IIf(Flag, 1, -1)

    '2. コマンド準備
    Dim CDPcommand As New Scripting.Dictionary, CDPparams As New Scripting.Dictionary
    CDPcommand.Add "method", "Target.setDiscoverTargets"

    CDPparams.Add "discover", Flag
    CDPcommand.Add "params", CDPparams

    If DefinitionCDPContext = 1 Or DefinitionCDPContext = 0 Then SendMessage CDPcommand, , False

    '3. 下限は、0にしておく
    If DefinitionCDPContext < 0 Then DefinitionCDPContext = 0: printMsg WARN_, "`CDPContext.cls` no longer exists.", FromProcedureName

    '4. ログ表示
    printMsg info_, "Currently, " & DefinitionCDPContext & " `CDPContext.cls` file is being defined.", FromProcedureName
End Property

'***************************************************************************************************
'* 機能　　："Target.targetCreated"による非同期イベント収集を有効/無効化します
'---------------------------------------------------------------------------------------------------
'* SET引数 ：格納させたいCollection変数。`Nothing`で渡せば、収集を無効化します
'---------------------------------------------------------------------------------------------------
'* 機能説明：・タブやiframeが開いた際の非同期イベントを流すかの切り替えです。
'            ・参照渡し方式により、わざわざ`Get`プロパティを介さず、取得可能です
'---------------------------------------------------------------------------------------------------
'* 注意事項：用が済んだら必ず、`Nothing`を渡して、参照カウンタを0にします
'***************************************************************************************************
Property Set RunCollect_TargetCreated(ArgCollect As Collection)
    '引数のポインタCollectionを指定させる
    Set TargetInfos = ArgCollect
End Property

Private Function getBrowserPath(strBrowserType As String) As String
'----------------------------------------------------------------------------------
' Obtain the installation path of the browser program. This is useful when
' the installation path is changed by the user or their corporation for security.
' Updated: 22/03/2023
' 20251023 Daniel Polak - added fallback path when registry lookup fails (in a locked-down environment)
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".getBrowserPath"


    Select Case True
        Case InStr(LCase(strBrowserType), "chrome"): br.Type = brChrome
        Case InStr(LCase(strBrowserType), "edg"): br.Type = brEdge         'On edge it is just "Edg"
        Case Else: printMsg ERROR_, "Unable to getBrowserPath for browser type: " & br.Type, FromProcedureName, doRaiseError:=True
    End Select

    'カスタムパスの指定があったらそっちを優先
    With ShSetting01_StartBrowser
        If LenB(.Range(.UseRangeName(1, FromProcedureName)).value) = 0 Then
            '1. 必要な変数を用意
            Dim strKeyPath      As String   'ベースパス
            Dim ValueSize       As Long     '値の大きさ

            '2. ブラウザのインストールパスのあるレジストリ場所を指定
            strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & br.Type

            '3. レジストリから、ブラウザインストールパスを特定
            '※文字列の場合、仕組み上まず、文字数チェックを行います
            If RegGetValueW(HKEY_LOCAL_MACHINE, StrPtr(strKeyPath), 0, RRF_RT_REG_SZ, 0, 0, ValueSize) = RRF_SUCCESS Then
                '4. 得た文字数を基に改めて、値の中身を取り出す
                getBrowserPath = String(ValueSize / 2, vbNullChar)
                RegGetValueW HKEY_LOCAL_MACHINE, StrPtr(strKeyPath), 0, RRF_RT_REG_SZ, 0, StrPtr(getBrowserPath), ValueSize

                '5. `vbnullchar`以前を取り出す
                getBrowserPath = Split(getBrowserPath, vbNullChar)(0)
            End If
        Else
            '1. カスタムパスを指定(ポータブル版ブラウザなど)
            getBrowserPath = .Range(.UseRangeName(1, FromProcedureName)).value
        End If
    End With

    'パスが空の場合は、次の捜索へ
    If LenB(getBrowserPath) = 0 Then

        '2. デフォルトインストールパスチェック
        Select Case br.Type
            Case brChrome
                If LenB(Dir(Environ("ProgramFiles") & "\Google\Chrome\Application\" & br.Type)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles") & "\Google\Chrome\Application\" & br.Type
                ElseIf LenB(Dir(Environ("ProgramFiles(x86)") & "\Google\Chrome\Application\" & br.Type)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles(x86)") & "\Google\Chrome\Application\" & br.Type
                End If
            Case brEdge
                If LenB(Dir(Environ("ProgramFiles") & "\Microsoft\Edge\Application\" & br.Type)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles") & "\Microsoft\Edge\Application\" & br.Type
                ElseIf LenB(Dir(Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\" & br.Type)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\" & br.Type
                End If
        End Select

        If LenB(getBrowserPath) > 0 Then
            If LenB(Dir(getBrowserPath)) = 0 Then getBrowserPath = ""
        End If


        '2段階の捜索でもパスが見つからない場合は、ブラウザが入ってない or カスタムパスインストール or ポータブル版ブラウザ or Canaryブラウザの可能性
        If LenB(getBrowserPath) = 0 Then printMsg ERROR_, "Unable to locate " & br.Type & " installation path.", FromProcedureName, doRaiseError:=True
    End If

End Function

'***************************************************************************************************
'* 機能　　：新しいデバッグブラウザセッションを開始する前にクリーンアップを行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：実行結果　※0で成功です
'* 引数    ：BrowserPath        起動しようとするブラウザフルパス
'            UserDataArg        `--user-data-dir` or `--profile-directory`といった起動ユーザー引数を適用します
'            onlyCdpSessions    `True` ：`--remote-debugging-pipe`,`--user-data-dir="..."`を含む、CDP制御サインに対してのブラウザをKillします
'                               `False`：起動しようとするブラウザパスのみに対してのブラウザをKillします。ポリシー：UserDataDir の対策です
'---------------------------------------------------------------------------------------------------
'* 機能説明：・基本は対象ブラウザの`remote-debugging-pipe`プロセスのみを終了します
'            ・コマンドライン引数に加えて、ターミナルサーバー上で (このユーザーに属さない) セッションを閉じる際の問題も解決しています
'***************************************************************************************************
Private Sub cleanUpSessions(onlyCdpSessions As Boolean)
    Const FromProcedureName As String = ThisClassName & ".cleanUpSessions"


    '1. 今、Excelが動いてるプロセスIDとそのログインユーザーのセッションIDを取得します
    Dim lProcessId As Long, lSessionId As Long
    lProcessId = GetCurrentProcessId()          'Get the current(Excel) process ID
    ProcessIdToSessionId lProcessId, lSessionId 'Get the Session ID for the current(Excel) process

    ' 2. exe名を取得
    Dim exeName As String, tmp
    tmp = Split(br.Path, "\")
    exeName = tmp(UBound(tmp))

    ' 3. 起動中のプロセスを巡る
    Dim proc As Object
    On Error Resume Next
    For Each proc In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process Where SessionId = " & lSessionId)
        '4. 起動中のブラウザ名であるか確認
        If LCase$(proc.Name) = LCase$(exeName) Then
            '5. Killタイプ判定
            If Not onlyCdpSessions Then
                '6. デバッグブラウザ(CdpSessions)であるか問わず、通常ブラウザもKillします
                proc.Terminate

            ElseIf InStr(1, proc.CommandLine, "--" & ArgRemoteDebuggingName) > 0 And InStr(1, proc.CommandLine, br.User) > 0 Then
                '6. デバッグブラウザ(CdpSessions)であるかつ、対象ユーザー名なら、Killします
                 proc.Terminate
            End If
        End If
        DoEvents
    Next
    On Error GoTo 0
    Err.Clear

    printMsg info_, "cleanUpSessions completed and previous CdpSessionID removed", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：リモート デバッグに関するポリシーチェックを行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：BrowserExeName        "msedge.exe" とかを指定します。本家ライブラリに沿った形です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・企業によっては、ポリシー設定でリモートデバッグできないようにしています。
'　　　　　　・設定されていると、起動はできたのに、なんか操作できないといった混乱が発生するため、このチェックを設けています
'            ・現在は下記の確認に対応してます
'               ・RemoteDebuggingAllowed
'               ・UserDataDir
'***************************************************************************************************
Private Sub IsRemoteDebuggingBlockedByPolicy(BrowserExeName As String)
    Const FromProcedureName As String = ThisClassName & ".IsRemoteDebuggingBlockedByPolicy"


    'レジストリ確認に必要な定数　※ポリシーチェックに使います。
    Const DisableRemote As Long = 0                    '`0`が、リモート デバッグを使用することはできません。

    'ブラウザ名に応じて、ベースとなるキーパスを設定
    Dim RegPathBrowser As String
    Select Case BrowserExeName
        Case brEdge: RegPathBrowser = "\Microsoft\Edge"
        Case brChrome: RegPathBrowser = "\Google\Chrome"     'https://issues.chromium.org/issues/42323066
        Case Else: printMsg ERROR_, "Browser initiation fails: only the defined browser types can be accepted", FromProcedureName, doRaiseError:=True
    End Select

    'ポリシー情報のあるパスを設定します
    Dim BaseKeyPath As String: BaseKeyPath = "SOFTWARE\Policies" & RegPathBrowser


    '必要な変数を用意
    Dim checkValueName          As String   '値の名前
    Dim ValueExists             As Long     '存在有無結果
    Dim ValueContents_long      As Long     '値の中身(数字のみ)
    Dim ValueContents_String    As String   '値の中身(文字列のみ)
    Dim ValueSize               As Long     '値の大きさ

    '"HKLM"と"HKCU"を確認する
    Dim i As Long, SetRegistryTree As Long
    Dim ErrorMes As String
    For i = 0 To 1
        'レジストリツリーを設定
        '最初は、PC単位、次は、現在のログインユーザー単位で確認します
        SetRegistryTree = IIf(i, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE)


        ' ------------ 1. RemoteDebuggingAllowed ------------
        'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/remotedebuggingallowed
        '※これに引っかかると残念ですが、CDP制御はできません
        checkValueName = "RemoteDebuggingAllowed"

        '1-1. 32bit値としてチェックする
        ValueSize = SIZE_REG_DWORD
        ValueExists = RegGetValueW(SetRegistryTree, StrPtr(BaseKeyPath), StrPtr(checkValueName), RRF_RT_REG_DWORD, 0, VarPtr(ValueContents_long), ValueSize)
        printMsg Debug_, IIf(i, "HKCU-", "HKLM-") & "Policy：" & checkValueName & " existence status - " & ValueExists, FromProcedureName

        '1-2. そのデータ名があるか？
        If ValueExists = RRF_SUCCESS Then
            printMsg Debug_, IIf(i, "HKCU-", "HKLM-") & "Policy：" & checkValueName & " contents - " & ValueContents_long, FromProcedureName

            '1-3. リモートできない設定値になってるか？
            If ValueContents_long = DisableRemote Then
                '通達
                ErrorMes = IIf(i, "現在のログインPCアカウントでは、ブラウザのリモートデバッグが、ブロックされています。" & vbCrLf & vbCrLf & "CDP機能は、利用できません。" & vbCrLf & "これが個人のPCアカウントの場合は、レジストリ操作にて`1`にして下さい。", _
                                  "会社のセキュリティポリシーにより、ブラウザのリモートデバッグが、ブロックされています。" & vbCrLf & vbCrLf & "CDP機能は、利用できません。" & vbCrLf & "情報システム部門に、お問い合わせください。")

                printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True
                Exit For
            End If
        End If


        '------------ 2. UserDataDir ------------
        'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/userdatadir
        '※これに引っかかるとCDP制御はできますが、デバッグブラウザとは関係ない普通のブラウザも毎回、落とさないといけない制限となります
        checkValueName = "UserDataDir"

        '2-1. 文字列として、チェックする
        '※文字列の場合、仕組み上まず、文字数チェックを行います
        ValueSize = 0   'リセット
        ValueExists = RegGetValueW(SetRegistryTree, StrPtr(BaseKeyPath), StrPtr(checkValueName), RRF_RT_REG_SZ, 0, 0, ValueSize)
        printMsg Debug_, IIf(i, "HKCU-", "HKLM-") & "Policy：" & checkValueName & " existence status - " & ValueExists, FromProcedureName

        '2-2. そのデータ名があるか？
        If ValueExists = RRF_SUCCESS Then
            printMsg Debug_, IIf(i, "HKCU-", "HKLM-") & "Policy：" & checkValueName & " contents - " & ValueContents_String, FromProcedureName

            'フラグ立て
            currentPolicy.UserDataDir.Exists = True

            '得た文字数を基に改めて、値の中身を取り出す
            ValueContents_String = String(ValueSize / 2, vbNullChar)
            RegGetValueW SetRegistryTree, StrPtr(BaseKeyPath), StrPtr(checkValueName), RRF_RT_REG_SZ, 0, StrPtr(ValueContents_String), ValueSize

            '中身が空の場合は、`User Data`配下として返します
            '※厳密には、`vbnullchar`が、先頭にある時
            currentPolicy.UserDataDir.Contents = IIf(InStr(1, ValueContents_String, vbNullChar) = 1, Environ("LOCALAPPDATA") & RegPathBrowser & "\User Data", Split(ValueContents_String, vbNullChar)(0))

            '通達
            ErrorMes = IIf(i, "現在のログインPCアカウントでは、ブラウザデータの配置場所が`" & currentPolicy.UserDataDir.Contents & "`に固定されています。起動中のブラウザを全て落とします。これが個人のPCアカウントであり、この仕様を回避する場合は、レジストリ操作にて削除して下さい。", _
                              "会社のセキュリティポリシーにより、ブラウザデータの配置場所が`" & currentPolicy.UserDataDir.Contents & "`に固定されています。起動中のブラウザを全て落とします。")

            printMsg WARN_, ErrorMes, FromProcedureName
            Exit For
        End If
    Next
End Sub



'***************************************************************************************************
'                               ■■■ 再接続処理系 ■■■
'***************************************************************************************************
'* 機能　　：Excelテーブルにある既存のパイプハンドル情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：パイプ生存論理値
'* 引数　　：userProfile        再アタッチしたいユーザー名(user-data-dirに基づく)
'---------------------------------------------------------------------------------------------------
'* 注意事項：パイプが生きてない場合は、このメソッドから再開できません
'***************************************************************************************************
Public Function reattach(userProfile As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".reattach"


    '1. 識別名称を設定
    br.User = userProfile

    '2. パイプハンドル制御するための初期化
    Set PipeCore = New CDPCore
    PipeCore.UseConverterUtf8FromString = UseConverterUtf8FromString    '`New`による直後に、文字コード設定を反映

    '3. 1を基に、Excelテーブルからハンドル情報を取得
    deserializeForPipe

    '4. バージョン確認CDPコマンドを打ってパイプ生存確認
    If isLiveBrowser Then
        reattach = True

        'ステータス更新
        getBrowserInfo
        printParams
    Else
        reattach = False
        printMsg ERROR_, "Failed to reattach to previous Pipe handle. user-data-dir=" & userProfile, FromProcedureName
    End If

End Function

'***************************************************************************************************
'* 機能　　：Excelのテーブルから、ブラウザへの接続パイプ情報を取得します
'***************************************************************************************************
Private Sub deserializeForPipe()
    Const FromProcedureName As String = ThisClassName & ".deserializeForPipe"


    '`UserName`に紐づくパイプ通信のハンドル情報の読み込みへ移る
    '※載ってない物はエラーとする
    If Not PipeCore.deserialize(br.User) Then printMsg ERROR_, Chr(34) & br.User & """ cannot reattach because there is no pipe handle information in the Excel table.", FromProcedureName, doRaiseError:=True
End Sub

'***************************************************************************************************
'* 機能　　：パイプハンドルの生存確認を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：パイプ生存論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：バージョン確認コマンドを送り、それでパイプ生存チェックを行います
'***************************************************************************************************
Public Function isLiveBrowser() As Boolean
    '1. バージョン確認CDPを送る
    '※生存確認なので、内部でのエラーは無視する
    Dim CDPcommand As New Scripting.Dictionary
    CDPcommand.Add "method", "Browser.getVersion"
    SendMessage CDPcommand, dbgMsg:=False

    '2. いずれもパイプエラーコードがなければ、`True`とします
    If PipeCore.LastErrorPeekNamedPipe = 0 And PipeCore.LastErrorReadFile = 0 And PipeCore.LastErrorWriteFile = 0 Then isLiveBrowser = True
End Function



'***************************************************************************************************
'                           ■■■ 各種接続情報記録/削除設定 ■■■
'***************************************************************************************************
'* 機能　　：Excelのテーブルに、ブラウザへの接続パイプ情報を記録します
'***************************************************************************************************
Private Sub serializeForPipe()
    PipeCore.serialize br.User
End Sub

'***************************************************************************************************
'* 機能　　：Excelのテーブルに、メインタブの接続情報を記録します
'---------------------------------------------------------------------------------------------------
'* 引数　　：sessionID_mainTab      メインタブでのSessionId
'            targetID_mainTab       メインタブでのtargetId
'---------------------------------------------------------------------------------------------------
'* 機能説明：
' This method transcribes Chrome's Json message into a string.
' The serialized string can then be used to automate an active
' session with a known serial.
' Updated: 03/05/23 - Replace sessionID with targetID since it is
'                     constant and stable on every CDP session
'          22/01/26 - Added specification to user profile name to
'                     the cpid file name to allow targetting of
'                     specific profile for reattachment.
'***************************************************************************************************
Public Sub serializeForMainTab(sessionID_mainTab As String, targetID_mainTab As String)
    Const FromProcedureName As String = ThisClassName & ".serialize"


    '1. 必要なデータを`Dictionary`に詰める
    Dim BrowserInfo As New Dictionary
    BrowserInfo.Add "BiDi-context", vbNullString  '`WebDriverBiDi`用に枠だけ確保
    BrowserInfo.Add "sessionID", sessionID_mainTab
    BrowserInfo.Add "targetID", targetID_mainTab

    '2. Excelのテーブルへ記録する
    Set ShSetting01_StartBrowser.serializeToTable(br.User, 2, FromProcedureName) = BrowserInfo
End Sub

'***************************************************************************************************
'* 機能　　：Excelのテーブルから、ブラウザへの接続パイプとタブ情報を削除します
'***************************************************************************************************
Private Sub deleteCDPID()
    Const FromProcedureName As String = ThisClassName & ".deleteCDPID"


    With ShSetting01_StartBrowser
        .DeleteCDPToTable(1, FromProcedureName) = br.User  'パイプ通信Handle情報
        .DeleteCDPToTable(2, FromProcedureName) = br.User  'ブラウザセッション情報
    End With
End Sub



'***************************************************************************************************
'                              ■■■ タブ操作関連 ■■■
'***************************************************************************************************
Public Function newTab(Optional Url As String, Optional newWindow As Boolean, Optional setMain As Boolean, Optional isHidden As Boolean, Optional browserContextId As String, Optional isBackground As Boolean) As CDPContext
'------------------------------------------------------------------
' Create a new blank tab within the same instance.
' url: The initial URL the page will be navigated to. An empty
' string indicates about:blank.
' newWindow: If true will open a separate window instead of a tab.
' setMain: If true will set the new tab as the main session tab.
' isBackground: If true, creates the new target in background.
' In the original CDP document, each new tab is called "Target" and
' the equivalence of this method is Target.createTarget
' This returns the Session ID of the newly created tab.
' Updated: 20/03/23 Return the result as a new Browser object
'          26/04/23 Format url argument to accept simplified input
'          22/01/26 Added .reattach where appropriate
'          07/03/26 Added hidden argument
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".newTab"


    Dim params As New Scripting.Dictionary
    Dim sessionID_ As String

    params("url") = checkURI(Url)
    If newWindow Then params("newWindow") = newWindow
    If isHidden Then params("hidden") = isHidden
    If StrPtr(browserContextId) Then params("browserContextId") = browserContextId
    If isBackground Then params("background") = isBackground

    Dim results As BiDiCDPJson
    Set results = ExecuteCDP("Target.createTarget", params)

   'Return the new tab's Session ID
    Dim tabId As String
    tabId = results.StringKey("targetId")
    sessionID_ = attachToTab(tabId)

   'Switch attachment to the new tab
   '22/01/26: From v3.1, reattachment is no longer automatic so .reattach is needed
    printMsg info_, "newTab will now open a new CDP instance", FromProcedureName


    '------ 専用のクラスとして返却する準備 ------
    EnableDiscoverTargets = True    '`CDPContext.cls`を生成するので、カウントアップ
    Set newTab = New CDPContext
    With newTab
        '1. このブラウザのパイプハンドル情報を渡す
        Set .InheritanceCDPCore = PipeCore

        '2. このブラウザ情報を渡す
        Set .InheritanceCDPBrowser = Me

        '3. 先ほど生成した`sessionId`を渡す
        .CurrentSessionID = sessionID_

        '4. Urlが設定されてる場合は、読み込み完了まで待機する
        If LenB(Url) Then .wait

        '5. ステータスの更新
        .getSessionInfo

        '6. メイン化する場合は、Excelに記録しておく
        If setMain Then serializeForMainTab .CurrentSessionID, .CurrentTargetID
    End With
End Function

'***************************************************************************************************
'* 機能　　：既に開いているタブのセッションIDを、タブの名前またはURLに基づいて検索します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - CDPBrowser
'* 引数　　：tabName        タブ名
'            Url            URL
'            setMain        `True`で、Excelのタブ情報テーブルに上書きします。タブのリアタッチに利用します
'            SearchType     "page","iframe" といった、セッションの種類を指定します
'            doRetrySecond  指定秒以内に出るまでリトライを繰り返します。0で、1回のみです
'---------------------------------------------------------------------------------------------------
'* 更新履歴：06/06/23               - Added doRetry argument
'            03/05/23               - Combined the method with the getTabNew
'            24/10/25 Daniel Polak  - Predictable, precise, and script-friendly tab selection that scales well when there are many similar tabs open
'                                     Exact title and URL-prefix matching avoids accidentally selecting tabs with similar titles/URLs
'            23/10/25 Daniel Polak  - Fixed typo timeStart instead of timerStart and fixed bug 60000s instead of 60s
'            01/08/26 Eschamali     - iframeへの切り替えができるように改良
'            22/01/26 Long Vh       - Added .reattach where appropriate
'---------------------------------------------------------------------------------------------------
'* 機能説明：tabName引数とurl引数は部分的な指定(部分一致)が可能です。SearchType はデフォルトは、"page"にしてます。
'* 注意事項：これらの引数を両方とも省略すると、getTabは最も近い未接続のタブに接続しようとします。
'***************************************************************************************************
Public Function getTab(Optional tabName As String, Optional Url As String, Optional setMain As Boolean, Optional SearchTypeID As TargetgetTargetsType, Optional doRetrySecond As Double) As CDPContext
    Const FromProcedureName As String = ThisClassName & ".getTab"


    Dim NodeToken As Long
    Dim tabId As String, candidateTabId As String, errMsg As String
    Dim targets As BiDiCDPJson
    Dim timerStart As Double
    Dim tabfound As Boolean
    Dim bestScore As Long, score As Long
    Dim info As String

    If doRetrySecond > 0 Then timerStart = TimerCounter

    'Typeを決める
    Dim SearchType As String
    Select Case SearchTypeID
        Case TargetgetTargetsType.Page: SearchType = "page"
        Case TargetgetTargetsType.iFrame: SearchType = "iframe"
        Case Else: Err.Raise vbObjectError + 1, FromProcedureName, "This parameter is undefined."
    End Select

searchit:

    Set targets = ExecuteCDP("Target.getTargets").NodeKey("targetInfos")
    NodeToken = targets.FirstChildToken  ' 配列の先頭トークン

    'スコア方式を利用して、理想の対象タブ情報を取得します
    Const MaxScore As Long = 100
    bestScore = -1
    candidateTabId = ""
    Do While NodeToken > 0
        If targets.TokenString(NodeToken, "type") = SearchType Then
            score = -1
            Select Case True
                Case LenB(tabName)
                    If targets.TokenString(NodeToken, "title") = tabName Then
                        score = 100
                    ElseIf InStr(targets.TokenString(NodeToken, "title"), tabName) Then
                        score = 50
                    Else
                        errMsg = "with name containing """ & tabName & """"
                    End If
                Case LenB(Url)
                    If Left$(targets.TokenString(NodeToken, "url"), Len(Url)) = Url Then
                        score = 90
                    ElseIf InStr(targets.TokenString(NodeToken, "url"), Url) Then
                        score = 40
                    Else
                        errMsg = "with url containing """ & Url & """"
                    End If
                Case Else
                    If targets.TokenBool(NodeToken, "attached") = False Then
                        score = 80
                    Else
                        score = 30
                        errMsg = "as no new unattached tab could be found"
                    End If
            End Select

            If score > bestScore Then
                bestScore = score
                candidateTabId = targets.TokenString(NodeToken, "targetId")
            End If

            '最大スコアが出てるなら、捜査打ち切り
            If bestScore = MaxScore Then Exit Do
        End If

        NodeToken = targets.NextToken(NodeToken) ' 次の要素へ
    Loop

    If bestScore >= 0 Then
        tabfound = True
        tabId = candidateTabId
    End If

   'Do retry as long as not timeout
    If Not tabfound And TimerCounter - timerStart < doRetrySecond * 1000 Then GoTo searchit

   'Returns nothing if not found
   'Nothing is preferred over throwing error
    If Not tabfound Then
        info = IIf(LenB(tabName), tabName, IIf(LenB(Url), Url, "(no new tab found)"))
        printMsg WARN_, "Unable to find the tab: " & info, FromProcedureName
        Exit Function
    Else
        printMsg info_, "Found the tab with ID: " & tabId & ". (Certainty：" & bestScore & "%)", FromProcedureName
    End If

searchfound:

   'Return the session object for reference
    Dim sessionID_ As String
    sessionID_ = attachToTab(tabId)


    '------ 専用のクラスとして返却する準備 ------
    '22/01/26 .reattach br.User is needed from v3.1 onwards to enable multi-instances automation
    EnableDiscoverTargets = True    '`CDPContext.cls`を生成するので、カウントアップ
    Set getTab = New CDPContext
    With getTab
        '1. このブラウザのパイプハンドル情報を渡す
        Set .InheritanceCDPCore = PipeCore

        '2. このブラウザ情報を渡す
        Set .InheritanceCDPBrowser = Me

        '3. 先ほど生成した`sessionId`を渡す
        .CurrentSessionID = sessionID_

        '4. ステータスの更新
        .getSessionInfo

        '5. メイン化する場合は、Excelに記録しておく
        If setMain Then serializeForMainTab .CurrentSessionID, .CurrentTargetID
    End With
End Function

'***************************************************************************************************
'* 機能　　：`targetId`に対する`DevTools`ウィンドウを開きます
'---------------------------------------------------------------------------------------------------
'* 引数　　：targetId   対象のタブID
'            panelId    最初に開きたい項目があれば指定
'---------------------------------------------------------------------------------------------------
'* 機能説明：F12が効かないページでも`DevTools`ウィンドウを開かせます
'* 注意事項：WebView2,Electronの場合は、効かない場合があります
'***************************************************************************************************
Public Sub openDevTools(TrgetID As String, Optional panelId As String)
    'パラメーター準備
    Dim params As New Scripting.Dictionary
    params.Add "targetId", TrgetID
    If StrPtr(panelId) Then params.Add "panelId", panelId

    '実行
    ExecuteCDP "Target.openDevTools", params
End Sub

Public Sub printTabs()
'------------------------------------------------------------------
' Print all current tabs info. Useful for debugging.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".printTabs"


    Dim results As Scripting.Dictionary, target_ As Scripting.Dictionary
    Dim i As Long
    Dim targets As Collection
    Set results = ExecuteCDP("Target.getTargets")
    Set targets = results("targetInfos")

   'Read info
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, "Tabs Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    i = 1
    For Each target_ In targets
        If target_("type") = "page" Then
            printMsg Debug_, "#" & i & vbTab & _
                        "type: " & target_("type") & " | " & _
                        "attached: " & target_("attached") & " | " & _
                        "title: " & target_("title") & " | " & _
                        "url: " & target_("url") & " | " & _
                        "target ID: " & target_("targetId"), FromProcedureName
            i = i + 1
        End If
    Next

End Sub

Public Function tabCount() As Long
'------------------------------------------------------------------------------
' Print all current tabs info. Useful for debugging.
' Updated : 05/02/25 - changed from printTabs to tabCount to get the count too
'------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".tabCount"


    Dim results, targets, target_ As Scripting.Dictionary
    Dim i As Integer

    Set results = ExecuteCDP("Target.getTargets")
    Set targets = results("targetInfos")

   'Read info
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, "Tabs Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    i = 0
    For Each target_ In targets
        If target_("type") = "page" Then
            printMsg Debug_, "#" & i & vbTab & _
                        "type: " & target_("type") & " | " & _
                        "attached: " & target_("attached") & " | " & _
                        "title: " & target_("title") & " | " & _
                        "url: " & target_("url") & " | " & _
                        "target ID: " & target_("targetId"), FromProcedureName
            i = i + 1
        End If
    Next

    tabCount = i

End Function



'***************************************************************************************************
'                              ■■■ CDPcommand操作関連 ■■■
'***************************************************************************************************
'* 機能　　：CDPのMethod名と、parametersを渡して、低レベルにてブラウザ制御をします
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド の実行結果情報。失敗時は、`Nothing`となります。
'* 引数　　：methodName             method名。"Network.getCookies"や"DOM.getDocument"というのを記入します
'            params                 Json-Dictionary形式で、KeyとItemを事前に定義したオブジェクト変数を指定します。
'            StopCDPError           `True`で、実行失敗時に`Err.Raise`で停止し、以降の処理を止めます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：This function enables advanced method invocation under CDP.
'            For a list of all methods available, see:
'               https://chromedevtools.github.io/devtools-protocol/tot/Browser/
'            Error message will be prompted automatically by ExecuteError in the Immediate Window.
'            The return object is Json String that has been converted to a dictionary object.
'            Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'
'* 更新履歴：・18/02/26 - SessionIDがあっても、BrowserContext優先にコマンド実行できるように改良。エラーで止めるかのスイッチングを導入
'            ・28/04/23 - Set this to Public as it is useful
'            ・20/03/23 - Added printMsg msgs & set to Private
'
'* 注意事項：・CDPの運営思想の都合上、多数の`Deprecated`,`Experimental`が存在します。
'              これらのタグのmethodを使用する際は、`StopError`を`True`にすることを推奨します。失敗時の以降の処理の予期せぬデータ破損などを防ぐためです。
'
'            ・`StopError`の規定値は、`True`なので、特別な理由がない限りは、省略推奨です
'            ・このメソッドは、コマンド結果が来るまで待機するメソッドです。送るコマンド内容や、サイトの挙動によっては、タイムアウトエラーになる場合があります。
'              その場合は、後述の結果を待たないタイプのメソッド`ExecuteCDPAsync`をご利用ください。
'***************************************************************************************************
Public Function ExecuteCDP(methodName As String, Optional params As Scripting.Dictionary, Optional StopCDPError As Boolean = True) As BiDiCDPJson
    Const FromProcedureName As String = ThisClassName & ".ExecuteCDP"


    'コマンド組み立て
    Dim CDPcommand As New Scripting.Dictionary
    CDPcommand.Add "method", methodName
    If Not params Is Nothing Then CDPcommand.Add "params", params   '`nothing`の場合は、`params`項目を省略

    '送信して結果を受け取り、Node変換
    Dim ResultStr As String: ResultStr = SendMessage(CDPcommand, dbgMsg:=StopCDPError)
    Dim resultJson As BiDiCDPJson
    Set resultJson = BiDiCDPJson.Parse(ResultStr)

    'Print error if there is
    If ExecuteError(resultJson) Then
        'エラーによる`Err.Raise`で止めるか？
        If StopCDPError Then
            'エラーになったmethod名と、結果生responseを表示して停止
            printMsg ERROR_, "Failed to invoke method """ & methodName & """" & vbCrLf & vbCrLf & "=== RawErrorResponse ===" & vbCrLf & ResultStr, FromProcedureName, doRaiseError:=True
        Else
            'エラーになったmethod名だけログ出して続行
            printMsg WARN_, "Failed to invoke method """ & methodName & """", FromProcedureName
        End If
    Else
        'Return result object
        If resultJson.Exists("result") Then Set ExecuteCDP = resultJson.NodeKey("result")
    End If

    'デバッグ用に、As Is 出力
    If PipeCore.ShowLogCDPresponse Then printMsg Debug_, "ExecuteCDP `" & methodName & "` with result: " & ResultStr, FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`ExecuteCDP`の非同期実行(コマンド結果を待たない)バージョンです
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド 実行時のID値(整理券番号のイメージでOK)
'* 引数　　：※`ExecuteCDP`と一緒のため、省略
'***************************************************************************************************
Public Function ExecuteCDPAsync(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As Long
    Const FromProcedureName As String = ThisClassName & ".ExecuteCDPAsync"


    'コマンド組み立て
    Dim CDPcommand As New Scripting.Dictionary
    CDPcommand.Add "method", methodName
    If Not params Is Nothing Then CDPcommand.Add "params", params   '`nothing`の場合は、`params`項目を省略

    '送信
    SendMessage CDPcommand, True, StopError

    '実行したIDを返却
    ExecuteCDPAsync = PipeCore.LastCommandID
    printMsg Debug_, "Async ExecuteCDP: " & methodName & " - ID: " & ExecuteCDPAsync, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：CDP にメッセージを送信し、Chromium ブラウザでメッセージが受信されているかどうかを確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：Async が、`False` → CDP-Jsonコマンド の実行結果(Json文字列)
'                      `True`  → `vbnullstring`
'
'* 引数　　：DirMessage             Chromium ブラウザに送るCDP-Jsonコマンド(Dictionary型)。最上位要素に、`method`、`params`があるようにしてください。
'
'            Async                  `True` ：コマンドの結果を待ちません。Do Loopフェーズには入りません。
'                                   `False`：コマンドの結果を待ちます。Do Loopフェーズに入り、そのIDが来るまで待ちます。この場合、非同期に実行した結果のID回収もついでに行います
'
'            dbgMsg                 `True` ：`CDPCore.cls - TakeEvents`にて、エラーがあったら止まります
'                                   `False`：`CDPCore.cls - TakeEvents`にて、エラーがあっても止まらず、`SendMessage`は`nothing`として返ります
'---------------------------------------------------------------------------------------------------
'* 更新履歴：・03/05/23 - Added soft-reset on peeknamepipe error
'            ・03/22/23 - Added error handling for better debugging
'            ・24/10/25 Daniel Polak - Added soft re-attach attempt when pipe error
'---------------------------------------------------------------------------------------------------
'* 注意事項：Async:=True の場合は常に`vbnullchar`で、返ります
'***************************************************************************************************
Private Function SendMessage(DirMessage As Scripting.Dictionary, Optional Async As Boolean, Optional dbgMsg As Boolean = True) As String
    Const FromProcedureName As String = ThisClassName & ".SendMessage"


   '--------------------------------------------------
   'Send message to the browser
   '--------------------------------------------------
    printMsg Debug_, "runCDP  - Async       = " & Async, FromProcedureName

    PipeCore.ReadyRunCDP DirMessage                                         'Send message to the browser instance
    Dim currentCommandID As Long: currentCommandID = PipeCore.LastCommandID '`ReadyRunCDP`直後のcommandIDを取得

    'WriteFile にてエラーがあった場合
    If PipeCore.LastErrorWriteFile Then printMsg ERROR_, "Error WriteFile in CDPCore.writeProcCDP. ErrorCode:" & PipeCore.LastErrorWriteFile & " - " & WinApiError.GetMessage(PipeCore.LastErrorWriteFile, "kernel32"), FromProcedureName, doRaiseError:=dbgMsg: Exit Function

    '非同期(結果はいらない)モードなら、ここで抜けます
    If Async Then Exit Function

   '------------------------------------------------
   'Read message returned by browser
   '------------------------------------------------
    'CDP-Jsonコマンドを送った際のIDが返ってくるまで、ループ受信処理を行いますが、タイムアウトになったら、抜けます
    Dim timerStart As Double: timerStart = TimerCounter
    Dim tmp As String
    printMsg Trace_, "Starting Processes incoming messages from the browser...", FromProcedureName
    Do
        '1. 受信処理によるパイプを拝見する
        '※内部にて`DoEvents`が走ります
        TakeEvents dbgMsg

        '2. 受信処理にてエラーがあったら、即抜け
        If PipeCore.LastErrorPeekNamedPipe > 0 Or PipeCore.LastErrorReadFile > 0 Then printMsg ERROR_, "Possible reasons: (1) the browser is not actually running (ie. call string failure); (2) the browser window was accidentally closed; (3) the CDP Pipe connection to the browser was cut-off somehow.", FromProcedureName: Exit Function

        '3. 内部の`RaiseEvent CDPBrowserID`による発火で、Dictionaryへ蓄積されたか？
        tmp = TakeResultCDP(currentCommandID)
        If StrPtr(tmp) Then
            '送ったコマンド結果であることを確認できたので結果を返す
            SendMessage = tmp
            printMsg info_, "Browser have confirmed receipt of command result ID: " & currentCommandID, FromProcedureName

            'ループから抜ける
            Exit Do
        Else
            printMsg Trace_, "Browser have not yet been able to confirm command ID: " & currentCommandID, FromProcedureName
        End If

        '4. 一定時間経ってもコマンド結果が来ないなら、エラーで停止します
        If TimerCounter - timerStart > timerOut Then Err.Raise CDPCustomErrorCodes.TIMEOUT, FromProcedureName, "Timeout waiting for browser to return message." Else printMsg Trace_, (TimerCounter - timerStart) / 1000 & " seconds have passed...", FromProcedureName
    Loop

    printMsg Trace_, "Command result ID:" & currentCommandID & " was returned from CDP-Json, so the receiving loop will end.", FromProcedureName
End Function



'***************************************************************************************************
'                              ■■■ CDPcommand結果関連 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対するCDPレスポンス結果を一旦、Dictionaryへ蓄積させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：id             コマンドID
'            RawJson        コマンド結果の生Json文字列　※As Is. ID付きのままで渡されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：`ID`をKeyに、Dictionaryへ`RawJson`を格納します。
'* 注意事項：・Jsonは、文字列のままで渡されます。
'            ・Dictionaryへの蓄積には、上限があります。上限を超えると、結果履歴(特に、`ExecuteCDPAsync`の結果)が全て削除されます
'***************************************************************************************************
Private Sub PipeCore_CDPBrowserID(id As Long, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPBrowserID"


    '1. 格納件数チェックし、次の蓄積で上限を超えたら、全履歴を削除します
    If ResultCDP_Browser.response.Count + 1 > ResultCDP_Browser.Limit Then
        ResultCDP_Browser.response.RemoveAll
        printMsg WARN_, "The accumulation limit has exceeded " & ResultCDP_Browser.Limit & ", so All result history has been deleted to prevent performance degradation.", FromProcedureName
    End If

    '2. その後、結果を蓄積
    ResultCDP_Browser.response.Add id, RawJson
    printMsg Trace_, "The results of the CDP(ID: " & id & ") execution from the browser have been received. → " & RawJson, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：オーバーフロー対策によるcommandIDがリセットされると呼び出されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：IDが20億に達する場合も、結果履歴が全て削除されます
'***************************************************************************************************
Private Sub PipeCore_ResetCommandID()
    ResultCDP_Browser.response.RemoveAll
End Sub

'***************************************************************************************************
'* 機能　　：`WithEvents`によって蓄積されたコマンドIDを取り出します
'---------------------------------------------------------------------------------------------------
'* 引数　　：CommandID  整理券番号
'* 返値    ：そのIDの実行結果のJson文字列
'---------------------------------------------------------------------------------------------------
'* 注意事項：・事前に、`TakeEvents` が実行されてる必要があります
'            ・見つかったIDは、取り出した後、Dictionaryから自動で削除されます
'            ・結果がない場合は、`vbnullstring`として返ります。`StrPtr`で判定可能です
'            ・エラー判定、JsonDictionay化は別途行って下さい
'***************************************************************************************************
Property Get TakeResultCDP(commandID As Long) As String
    Const FromProcedureName As String = ThisClassName & ".TakeResultCDP"


    '1. 存在するか？
    If ResultCDP_Browser.response.Exists(commandID) Then
        '2. Log化
        Dim tmp As String: tmp = ResultCDP_Browser.response(commandID)
        printMsg info_, "Browser were able to confirm the existence of CDP command ID:" & commandID & "　→　" & tmp, FromProcedureName

        '3. 返却
        TakeResultCDP = tmp

        '4. 取り出したので消しておく
        ResultCDP_Browser.response.Remove commandID
        printMsg Trace_, "CDP-command ID: " & commandID & " has been removed from `ResultCDP_Browser.response`.", FromProcedureName
    Else
        '2. 存在しない旨を伝える
        printMsg Trace_, "This CDP-command ID(" & commandID & ") passed is not running asynchronously or has been deleted.", FromProcedureName
    End If
End Property

'***************************************************************************************************
'* 機能　　：CDP-Jsonコマンド実行結果のエラーチェックを行い、最後のエラー記録も担います
'---------------------------------------------------------------------------------------------------
'* 返り値　：エラー有無論理値
'* 引数　　：result         パース済CDP-Jsonコマンド実行結果
'            ErrorCode      エラーコード　      ※参照渡しにより返されます
'            ErrorMessage   エラーメッセージ    ※参照渡しにより返されます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・Check if ExecuteCDP fails and what the error message returned is.
'            　Error message will be shown via printMsg
'
'            ・または、第2引数以降の変数からでもエラー情報をクイックに受け取れます
'---------------------------------------------------------------------------------------------------
'* Credit  ：https://github.com/PerditionC/VBAChromeDevProtocol
'***************************************************************************************************
Private Function ExecuteError(result As BiDiCDPJson, Optional ErrorCode As Long, Optional ErrorMessage As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ExecuteError"


    Dim ErrorData As Variant

    'データの状態をチェック
    If result.Count = 0 Then
        'タイムアウトもしくは、パイプは生きてるが、応答なしによる結果無し
        ExecuteError = True


        'カスタムなエラー情報として記載
        ErrorCode = -1 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        ErrorData = vbEmpty
        printMsg WARN_, "ExecuteCDP Error: [" & ErrorCode & "] " & ErrorMessage, FromProcedureName

        '最後のエラー情報を記録
        Set lastErrorDetailsCDP = New Dictionary
        lastErrorDetailsCDP.Add "code", ErrorCode
        lastErrorDetailsCDP.Add "message", ErrorMessage

    Else
        If result.Exists("error") Then
            'CDPコマンドエラー
            ExecuteError = True


            '`error`オブジェクトの中身で記録
            Dim errNode As BiDiCDPJson
            Set errNode = result.NodeKey("error")

            '主要なエラー情報を抜き出す
            ErrorCode = errNode.NumberKey("code")
            ErrorMessage = errNode.StringKey("message")
            printMsg WARN_, "ExecuteCDP Error: [" & ErrorCode & "] " & ErrorMessage, FromProcedureName

            '追加情報がある場合は、それもログ化
            If errNode.Exists("data") Then
                ErrorData = errNode("data")
                printMsg WARN_, "       " & ErrorData, FromProcedureName
            Else
                ErrorData = vbEmpty ' vbNullString? is lastErrorDetailsCDP("data") a string or object?
            End If

            '最後のエラー情報を記録
            Set lastErrorDetailsCDP = New Dictionary
            lastErrorDetailsCDP.Add "code", ErrorCode
            lastErrorDetailsCDP.Add "message", ErrorMessage
        Else
            'CDPコマンド正常終了
            ExecuteError = False


            'クリアしておく
            ErrorCode = 0
            ErrorMessage = vbNullString
            ErrorData = vbEmpty
        End If
    End If

End Function

'***************************************************************************************************
'* 機能　　：`ExecuteError`メソッドを通した際の、最後のエラー情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後のエラー情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」のエラー情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます
'            ・適切な`ExecuteCDP`を実行したつもりなのに返り値が、`Nothing`となるのが、エラーの目印となるので、それでこのエラー情報を見に行ってください
'            ・`ExecuteError`メソッドを通さないと機能しません
'            ・`Dictionary`オブジェクトとして得れるので必要に応じて`.RemoveAll`でエラークリアすることも可能です
'***************************************************************************************************
Property Get LastCDPJsonError() As Dictionary
    Set LastCDPJsonError = lastErrorDetailsCDP
End Property

'***************************************************************************************************
'* 機能　　：CDP-Json コマンド結果を蓄積させる上限を決めます
'---------------------------------------------------------------------------------------------------
'* LET引数 ：置き換える上限数
'---------------------------------------------------------------------------------------------------
'* 機能説明：格納上限件数を設定します。デフォルトは、65536件です。
'***************************************************************************************************
Property Let SetLimitCDPResult(Number As Long)
    ResultCDP_Browser.Limit = Number
End Property



'***************************************************************************************************
'                                       ■■■ 後始末系 ■■■
'***************************************************************************************************
Public Sub quit()
'-----------------------------------------------------------------------------------
' Closes the currently attached browser instance
' Original Name: closeBrowser
' Update: 24/10/2025: Daniel Polak - avoid immediate hard kills, reduce race conditions, and make shutdown behavior predictable and robust in non-responsive cases
'-----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".quit"


    Dim sngT As Double

    '1. 終了コマンドを送る
    '※閉じると同時に、Pipeエラーが起こることがあるので、エラー無視しておく
    ExecuteCDP "Browser.close", StopCDPError:=False

    '2. パイプ生存確認
    Dim PipeExists As Boolean
    If PipeCore.LastErrorPeekNamedPipe > 0 Or PipeCore.LastErrorReadFile > 0 Or PipeCore.LastErrorWriteFile > 0 Then
        '2-1. もうパイプは閉じてる
        PipeExists = False
    Else
        '2-1. どうもブラウザは生きてるようなのでコマンドを送って確認する
        sngT = TimerCounter
        Do While PipeExists And TimerCounter - sngT < 5000
            PipeExists = isLiveBrowser

            '2-2. なくなってたら、即抜け
            If Not PipeExists Then Exit Do
            sleep 0.1
        Loop
    End If

    '3. 一定時間たっても閉じない場合は、応答なしとし、プロセス終了をかけて、無理やり閉じる
    If PipeExists Then
        printMsg WARN_, "Browser didn't close gracefully within 5s; invoking cleanUpSessions.", FromProcedureName
        TerminateDebugBrowser
    End If

    '4. パイプハンドル、ブラウザセッション情報を消去しに行きます
    PipeCore.CloseHandleBrowser '各種ハンドルを閉じる
    deleteCDPID                 'テーブルからも削除する

    '5. it seems without waiting a bit the browser crashes and the next time it wants to recover from a crash
    sleep 1
    printMsg info_, "The browser has been terminated.", FromProcedureName

End Sub

Private Sub setCrashStateNormal()
'------------------------------------------------------------------
' Set Crashed state to Normal to prevent the "Restore Sessions?"
' popup from appearing every time the browser is started after
' being closed by browser.quit method.
' Requires user data dir profile stored in private var br.User and
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
    If currentPolicy.UserDataDir.Exists Then
        'Use this switch instead if your organization disables UserDataDir with policy
        filePath_Preferences = currentPolicy.UserDataDir.Contents & "\" & br.User & "\Preferences"
    Else
        'user-data-dir
        Select Case br.Type
            Case brChrome: filePath_Preferences = Environ("LOCALAPPDATA") & "\Google\Chrome\" & br.User & "\Default\Preferences"
            Case brEdge: filePath_Preferences = Environ("LOCALAPPDATA") & "\Microsoft\Edge\" & br.User & "\Default\Preferences"
        End Select
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
'* 機能　　：`CDPCore.InitCDP`で起動したデバッグブラウザプロセスハンドルに対して、プロセス終了をかけます
'---------------------------------------------------------------------------------------------------
'* 注意事項：CDPコマンド`Browser.close`を実行したのに、閉じれないアクシデント時に使用して下さい
'***************************************************************************************************
Private Sub TerminateDebugBrowser()
    TerminateProcess PipeCore.Connection_hProcess, 0
End Sub



'***************************************************************************************************
'                                ■■■ 非同期イベント関連 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する非同期イベントを処理します
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName 非同期イベントメソッド名
'            RawJson    非同期イベントの生Json文字列　※As Is. メソッド名付きのままで渡されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：`Target.targetCreated`に対する情報の蓄積を行います
'***************************************************************************************************
Private Sub PipeCore_CDPBrowserEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPBrowserEvent"


    '1. Select Case でスマートに、method名ごとの独自処理を行う
    Select Case methodName
        '1-1. `attachToSession`が行えるように準備する
        Case "Target.targetCreated"
            '1-2. `TargetInfo`を収集しない場合は、即抜け
            If TargetInfos Is Nothing Then Exit Sub

            '1-3. そのまま、蓄積
            printMsg Debug_, "Collecting `Target.targetCreated`... → " & RawJson, FromProcedureName
            TargetInfos.Add RawJson

        Case Else
            'None

    End Select
End Sub



'***************************************************************************************************
'                                   ■■■ ステータス系 ■■■
'***************************************************************************************************
Public Sub printParams()
'------------------------------------------------------------------
' Print all parameters of the current session for debuggings.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".printParams"


    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, br.LogID & " Browser - Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "Browser Type      | brType       | " & br.Type, FromProcedureName
    printMsg Debug_, "Browser Ver       | brVer        | " & br.Ver, FromProcedureName
    printMsg Debug_, "Browser Path      | brPath       | " & br.Path, FromProcedureName
    printMsg Debug_, "Core StdInWr      | hStdInWr     | " & PipeCore.ConnectionPipe_hStdInWr, FromProcedureName
    printMsg Debug_, "Core StdOutRd     | hStdOutRd    | " & PipeCore.ConnectionPipe_hStdOutRd, FromProcedureName
    printMsg Debug_, "Core hStderrOutRd | hStderrOutRd | " & PipeCore.ConnectionPipe_hStderrOutRd, FromProcedureName
    printMsg Debug_, "Core CDPInWr      | hCDPInWr     | " & PipeCore.ConnectionPipe_hCDPInWr, FromProcedureName
    printMsg Debug_, "Core CDPOutRd     | hCDPOutRd    | " & PipeCore.ConnectionPipe_hCDPOutRd, FromProcedureName
    printMsg Debug_, "Core dwProcessId  | dwProcessId  | " & PipeCore.Connection_dwProcessId, FromProcedureName
    printMsg Debug_, "CDP Version       | CDPVer       | " & CDPVer, FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：`Browser.getVersion`によるブラウザ情報を取得します
'***************************************************************************************************
Private Sub getBrowserInfo()
    br.Type = ExecuteCDP("Browser.getVersion")("product")
    br.Ver = Split(br.Type, "/")(1)
    If PipeCore.Connection_dwProcessId > 0 Then br.Path = getBrowserPath(br.Type)  'プロセスID情報がない場合は、取らない
End Sub



'***************************************************************************************************
'                             ■■■ SessionID発行と破棄 ■■■
'***************************************************************************************************
'* 機能　　：SessionID の発行を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：`TargetID`に対する`sessionId`  ※失敗時は、`vbnullstring`となります
'* 引数　　：tabId  TargetID
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Attach to a specific tab with its tabId (a.k.a targetId)
'            flatten: Enables "flat" access to the session via specifying sessionId attribute
'            in the commands. Set to True by default to make attachToTab work.
'            Returns: a session ID string that is needed for CDP automation.
'***************************************************************************************************
Public Function attachToTab(tabId As String) As String
    '1. コマンド準備
    Dim CDPcommand As New Dictionary
    Dim CDPparams As New Dictionary
    CDPcommand.Add "method", "Target.attachToTarget"

    CDPparams.Add "targetId", tabId
    CDPparams.Add "flatten", True
    CDPcommand.Add "params", CDPparams

    '2. 実行
    Dim tmp As BiDiCDPJson
    Set tmp = BiDiCDPJson.Parse(SendMessage(CDPcommand))

    '2. 結果があれば、返却
    If Not ExecuteError(tmp) Then attachToTab = tmp("result")("sessionId")
End Function

'***************************************************************************************************
'* 機能　　：SessionID の破棄を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：破棄成功論理値
'* 引数　　：SessionID      破棄したいSessionID
'---------------------------------------------------------------------------------------------------
'* 詳細説明：不必要なSessionIDを破棄することで、ブラウザの安定性が向上し、奇妙な不具合も消えるかもしれません
'* 注意事項：・ブラウザ権限で実行するため、メイン化してる`SessionID`も破棄すると、処理が止まるため、使う箇所は慎重に...
'            ・念のため成功論理値を返してますが基本、確認するほどでもないと思います
'***************************************************************************************************
Public Function DiscardSessionID(sessionID As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".DiscardSessionID"


    '1. コマンド準備
    Dim CDPparams As New Scripting.Dictionary
    CDPparams.Add "sessionId", sessionID

    '2. 破棄コマンド実行
    Dim tmp As BiDiCDPJson
    Set tmp = ExecuteCDP("Target.detachFromTarget", CDPparams, False)

    '3. ログとして出しておく
    If tmp Is Nothing Then
        printMsg Debug_, "Failed to destroy SessionID: " & sessionID, FromProcedureName
    Else
        printMsg info_, "Successfully destroyed the SessionID: " & sessionID, FromProcedureName
    End If
End Function



'***************************************************************************************************
'                               ■■■ その他ヘルパーメソッド ■■■
'***************************************************************************************************
'* 機能　　：URIのプロトコルチェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数の内容そのまんま
'* 引数　　：URI   ブラウザのURLに渡す予定の文字列
'---------------------------------------------------------------------------------------------------
'* 注意事項：・セキュリティ重視により、ホワイトリスト外のプロトコルが来たら、エラーで止まります
'            ・既に開いてるURIに関して(`getTab`メソッド)はチェック対象外とします
'***************************************************************************************************
Public Function checkURI(URI As String) As String
    Const FromProcedureName As String = ThisClassName & ".checkURI"


    '0. 空文字の場合は、チェックしません
    printMsg Debug_, "Checking the protocol of the URI...", FromProcedureName
    If LenB(URI) = 0 Then
        printMsg Debug_, "Checking is skipped because the URI is empty.", FromProcedureName

        checkURI = URI
        Exit Function
    End If

    '1. ホワイトリストのプロトコルを取得
    With ShSetting01_StartBrowser
        Dim tmp, WhiteList
        tmp = .ListObjects(.UseTableName(3, FromProcedureName)).DataBodyRange.Columns(Table03ColumnID.接頭辞)

        '配列か？
        If IsArray(tmp) Then
            'そのまんま
            WhiteList = tmp
        Else
            '配列じゃない場合は配列にしておく
            ReDim WhiteList(1 To 1, 1 To 1)
            WhiteList(1, 1) = tmp
        End If

        '解放
        tmp = vbNullString
    End With

    '2. 先頭始まりかチェック
    Dim i As Long
    For i = 1 To UBound(WhiteList)
        '先頭始まりなら、引数の内容そのまんまを返して抜ける
        If InStr(1, URI, WhiteList(i, 1)) = 1 Then
            printMsg Debug_, "Protocol check passed：" & WhiteList(i, 1), FromProcedureName

            checkURI = URI
            Exit Function
        End If
    Next

    '3. プロトコル名をチェック
    Dim ProtocolName, ErrorMes As String
    ProtocolName = Split(URI, ":")
    If UBound(ProtocolName) = 0 Then
        ErrorMes = "URLの形式が正しくありません。" & vbCrLf & "安全のため省略せず、`https://`や`file://`のような、プロトコル名を、URLの先頭に付けてください。" & vbCrLf & vbCrLf & "[問題のURL]" & vbCrLf & URI
    Else
        ErrorMes = "プロトコル`" & ProtocolName(0) & "`は、ホワイトリストに登録されていません。" & vbCrLf & "意図しないアプリケーションの起動などを防ぐため、処理を中断しました。" & vbCrLf & vbCrLf & "【解決策】" & vbCrLf & "このプロトコルへのアクセスが、本当に必要である場合は、シート「ブラウザ起動設定」のテーブル「ProtocolWhiteList」に、`" & ProtocolName(0) & "`を追加してから、もう一度、お試しください。"
    End If

    '4. エラーで中断
    printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True

End Function



'***************************************************************************************************
'                              ■■■ 複製用プロパティ ■■■
'***************************************************************************************************
'* 機能　　：このClass内で定義してる`CDPCore.cls`を受け取ります
'***************************************************************************************************
Property Get InheritanceCDPCore() As CDPCore
    Set InheritanceCDPCore = PipeCore
End Property



'***************************************************************************************************
'                              ■■■ PUBLIC PROPERTIES ■■■
'***************************************************************************************************
'* 機能　　：このClass内で起動してるブラウザパス値を受け取ります
'***************************************************************************************************
Property Get CurrentBrowserPath() As String
    CurrentBrowserPath = br.Path
End Property

'***************************************************************************************************
'* 機能　　：このClass内で起動してるブラウザの`user-data-dir`値を受け取ります
'***************************************************************************************************
Property Get CurrentBrowserUser() As String
    CurrentBrowserUser = br.User
End Property

'***************************************************************************************************
'* 機能　　：このClass内で起動してるブラウザのWindowクラス値を受け取ります
'***************************************************************************************************
Property Get CurrentBrowserWindowClass() As String
    CurrentBrowserWindowClass = chromeWindowClass
End Property

'***************************************************************************************************
'* 機能　　：CDP-Json コマンド結果を得るためのタイムアウト時間を設定します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：設定中のタイムアウト秒数
'* LET引数 ：置き換えるタイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 機能説明：最初のブラウザ起動時の遷移完了判定や、CDP-Json コマンド結果を得るのに待つ時間を設定します。デフォルトは、30sです。
'***************************************************************************************************
Property Get TimeOutSecond() As Double
    TimeOutSecond = timerOut / 1000 'ミリ秒に補正する
End Property

Property Let TimeOutSecond(TimeSec As Double)
    timerOut = TimeSec * 1000       'ミリ秒に補正する
End Property



'***************************************************************************************************
'                               ■■■ CDPCore.cls 直通メソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    PipeCore.printMsg LogLevel_, br.LogID, strMsg, From, isHeader, doRaiseError
End Sub

Public Sub sleep(Optional seconds As Double = 0.5)
    PipeCore.sleep seconds
End Sub

Public Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    PipeCore.TakeEvents StopPipeError, destruction
End Sub

Public Function TimerCounter() As Double
    TimerCounter = PipeCore.GetTicksMs
End Function



'***************************************************************************************************
'                         ■■■ このClassオブジェクト生成時のアクション ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------
' A quick way to quickly check if an automated session is already active and open.
' If so, automatically attempt to reattach to it.
' Updated: 07/06/23
'          22/10/25 Daniel Polak - prevent compilation errors because depending on whether we are running in Excel, Word or Access only one of ThisWorkbook, ActiveDocument and CurrentProject is available
'          22/01/26 Long Vh - removed the auto deserialization feature to running multiple instances of browser profiles.
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'Assign an objID for easy debug identification
    br.LogID = "BRID" & Format(Rnd * 1000, "000")


    '------------------------- クラス内変数の一部を初期化 -------------------------
    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        '1. Edgeウィンドウの表示モードの設定
        br.setShowWindow = .Range(.UseRangeName(12, FromProcedureName)).value

        '2. 文字コード設定
        UseConverterUtf8FromString = .Range(.UseRangeName(10, FromProcedureName)).value

        '3. ブラウザの使用タイプ
        UseChrome = .Range(.UseRangeName(4, FromProcedureName)).value
    End With

    'タイムアウト設定(デフォルトで、30s)
    timerOut = TIMEOUT_DEFAULT * 1000   'ミリ秒に補正

    '`WithEvents`からのコマンド結果を蓄積させる件数設定(デフォルトで、256件)
    Set ResultCDP_Browser.response = New Dictionary
    ResultCDP_Browser.Limit = MaxStoreResultCDP
    '------------------------------------------------------------------------------


    'Determine project host type for compatibility
    '※このフォークツールは、Excel特化に改良してるため、Excel以外は動作しません
    Select Case True
        Case InStr(Application.Name, "Excel"): Set thisApp = ThisWorkbook
        'Case InStr(Application.Name, "Word"): Set thisApp = ThisDocument
        'Case InStr(Application.Name, "Access"): Set thisApp = CurrentProject
    End Select

End Sub
