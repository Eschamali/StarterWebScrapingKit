VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
' WebView2Core.cls
' TLB不要・参照設定不要・外部ファイル不要の 64bit VBA 純粋実装
' DispCallFunc + vtable偽造 による WebView2 COM 直接制御
'
' 動作要件:
'   - 64bit Excel (VBA7)
'   - Microsoft Edge (Chromium) インストール済み
'   - WebView2Callbacks.bas（AddressOf制約のため必須）
'
' 基本フロー:
'   Initialize → OnEnvironmentCreated（CB） → OnControllerCreated（CB）
'   → Ready イベント → Navigate でページ遷移
'
' CDP: CallDevToolsProtocolMethod の vtable index はランタイムでずれるため
'   CDPMethodVTableIndex で上書き可（既定 36）。QueryInterface で「CDP 専用
'   インターフェース」を取る方式は WebView2 に存在しないため不可。将来は
'   ランタイム版取得＋テーブルで自動設定が現実的。
'***

' =====================================================================
' 型定義
' =====================================================================
' COM vtable 偽造オブジェクト（16バイト: vtablePtr(8) のみ）
Private Type VtblObj
    pVTable As LongPtr
End Type
' 4エントリの vtable（QI / AddRef / Release / Invoke）
Private Type VtblData4
    fn(0 To 3) As LongPtr
End Type
' ICoreWebView2EnvironmentOptions + Options2 用 13 エントリ
Private Type VtblData13
    fn(0 To 12) As LongPtr
End Type

Private Type RECT_T
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type



'***************************************************************************************************
'                  ■■■ Win32 API Declares (64bit PtrSafe) ■■■
'***************************************************************************************************
#If VBA7 Then

'Excelが自動で読み込んでくれる仕様
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions _
    Lib "WebView2Loader" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal environmentOptions As LongPtr, _
    ByVal handler As LongPtr) As Long

' OleAut32 ─ COM vtable メソッドを番号（オフセット）で呼ぶ
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal lCc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    prgvt As Any, _
    prgpvarg As Any, _
    pvargResult As Variant) As Long

' Kernel32 ─ メモリ操作・DLLロード・環境変数
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal length As Long)
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" ( _
    ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
' User32 ─ ウィンドウ操作
Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal uFlags As Long) As Long

' User32 ─ メッセージポンプ（COM STA コールバックの配送に必要）
Private Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageW" ( _
    lpMsg As Any, ByVal hWnd As LongPtr, _
    ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, _
    ByVal wRemoveMsg As Long) As Long
Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As Any) As Long
Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (lpMsg As Any) As LongPtr

' User32 ― WebView2 子 HWND 取得
Private Declare PtrSafe Function GetWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal uCmd As Long) As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" ( _
    ByVal hWnd As LongPtr, lpRect As RECT_T) As Long

#End If

' =====================================================================
' 定数
' =====================================================================
Private Const CC_STDCALL  As Long = 4
Private Const PTR_SIZE    As Long = 8   ' 64bit ポインタ = 8バイト
Private Const PM_REMOVE   As Long = 1   ' PeekMessage フラグ

' --- ICoreWebView2Environment vtable オフセット (index × PTR_SIZE) ---
Private Const OFF_ENV_CreateController As Long = 3 * 8   ' = 24  index 3

' --- ICoreWebView2Controller vtable オフセット ---
' 正規 ICoreWebView2Controller vtable (IUnknown 3本 + 以下の順):
'   [3] get_IsVisible  [4] put_IsVisible  [5] get_Bounds   [6] put_Bounds
'   [7] get_ZoomFactor [8] put_ZoomFactor [9] add_ZoomFactorChanged ...
'   [24] Close         [25] get_CoreWebView2
Private Const OFF_CTRL_get_CoreWebView2 As Long = 25 * 8 ' = 200  index 25
Private Const OFF_CTRL_put_IsVisible    As Long = 4 * 8  ' = 32   index 4
Private Const OFF_CTRL_put_Bounds       As Long = 6 * 8  ' = 48   index 6 (x64: RECT 16byte → ポインタ渡し)
Private Const OFF_CTRL_NotifyParentPos  As Long = 23 * 8 ' = 184  index 23 NotifyParentWindowPositionChanged
Private Const OFF_CTRL_Close            As Long = 24 * 8 ' = 192  index 24

' --- ICoreWebView2 vtable オフセット ---
Private Const OFF_WV2_Navigate             As Long = 5 * 8  ' = 40
Private Const OFF_WV2_NavigateToString     As Long = 6 * 8  ' = 48
Private Const OFF_WV2_get_BrowserProcessId As Long = 37 * 8 ' = 296  [propget] UINT32* value

' CallDevToolsProtocolMethod の vtable index。ランタイム/SDK でずれるためプロパティで上書き可（既定 36）
Private m_CDPMethodVTableIndex As Long
Private m_CDPMethodForSessionVTableIndex As Long
Private ResultCDPTimeout As Long

'ベースとなる作業フォルダ名
Const WorkSpaceFolderName As String = "WebView2FromVBA"



' =====================================================================
' プライベート状態
' =====================================================================
' vtable 偽造オブジェクト（環境 / コントローラ / CDP完了ハンドラ）
Private m_EnvHandlerThis  As VtblObj
Private m_EnvHandlerVT    As VtblData4
Private m_CtrlHandlerThis As VtblObj
Private m_CtrlHandlerVT   As VtblData4
Private m_CDPHandlerThis  As VtblObj
Private m_CDPHandlerVT    As VtblData4
Private m_CDPAsyncHandlerThis  As VtblObj
Private m_CDPAsyncHandlerVT    As VtblData4
' ICoreWebView2EnvironmentOptions 偽装（追加ブラウザ引数用）
Private m_EnvOptionsThis  As VtblObj
Private m_EnvOptionsVT    As VtblData13

' COM オブジェクトポインタ
Private m_pEnvironment  As LongPtr   ' ICoreWebView2Environment
Private m_pController   As LongPtr   ' ICoreWebView2Controller
Private m_pWebView2     As LongPtr   ' ICoreWebView2

' 親ウィンドウ
Private m_hWndParent As LongPtr

' リサイズ時に DoTimerResize が GetClientRect するための Frame HWND
Private m_hWndFrame As LongPtr

' 表示領域
Private m_Bounds As RECT_T

' フィードバックループ防止：前回 put_Bounds に渡した値
' put_Bounds → WM_SIZE → UserForm_Resize → ScheduleResize → DoTimerResize → put_Bounds...
' の無限ループを、「同じ bounds ならスキップ」で断ち切る。
Private m_LastAppliedRight  As Long
Private m_LastAppliedBottom As Long

' 初期遷移 URL
Private m_InitialUrl As String

' エラー情報
Private m_LastErrorCode As Long
Private m_LastErrorDesc As String

' 初期化ステート管理
'   0 = 初期 / 1 = Environment取得済み(Controller未作成) / 2 = Ready
Private m_State As Long


Private ErrorDetail As New WinApiError      'エラーコードから、詳細を取得するやつ
Private LogControl  As New Logger           'ログレベルの制御
Private brID                    As String
Private doLog                   As Boolean  'ログファイルにも、デバッグメッセージを残すか
Private logPath                 As String   'ログファイルの保存フォルダパス
Private Const ThisClassName         As String = "WebView2Core"       'ログのトレース用
Private Const ERR_TIMEOUT As Long = 900
Private Const ERR_CoreWebView2 As Long = 901
Private Const ERR_PROTOCOL As Long = 902
Private Const ERR_DictionaryStructure As Long = 903
Private Const ERR_UnregisteredProtocol As Long = 904
Private Const ERR_PolicyViolation As Long = 905
Private Const TIMEOUT_DEFAULT As Single = 30
Private Const TIMEOUT_NAVIGATE As Single = 45



' =====================================================================
' CDP メッセージ キュー / イベントキャプチャ
' =====================================================================
Private m_CDPResultQueue As Collection
Private m_LastCDPResultErrorCode As Long

'----------------------------------------------------------------------
' Async CDPコマンド結果の後回収（invokeMethodAsync 相当）
'   - async を投げるたびに、コールバック識別用の this ポインタをユニーク化
'   - CB_OnCDPMessage で pThis から cmdId を引き当て、結果を保存
'   - WebView2Browser 側から ResultCDPForAsync(cmdId) で回収
'----------------------------------------------------------------------
Private Const MAX_ASYNC_CDP_COMMANDS As Long = 256
Private m_CDPAsyncHandlerThisPool(0 To MAX_ASYNC_CDP_COMMANDS - 1) As VtblObj
Private m_CDPAsyncHandlerUsed(0 To MAX_ASYNC_CDP_COMMANDS - 1) As Boolean

Private m_CDPAsyncPThisHexToInfo As Dictionary 'hex(pThis) -> Array(idx, cmdId)
Private m_CDPAsyncResultJsonByCmdId As Dictionary 'cmdId -> raw result json
Private m_CDPAsyncErrorCodeByCmdId As Dictionary  'cmdId -> ErrorCode

' DevToolsProtocol event receiver（複数購読対応）
Private m_DevToolsProtocolEventSubscriptions As Dictionary ' eventName -> array(receiverPtr, token, handlerThisPtr)
Private m_DevToolsProtocolEventReceiverHexToName As Dictionary ' receiverPtr(hex) -> eventName（pSender 用・フォールバック）
Private m_DevToolsProtocolHandlerThisToEventName As Dictionary ' handlerThis(hex) -> eventName（コールバックの pThis で振り分け）

' ICoreWebView2DevToolsProtocolEventReceivedEventHandler の vtable 偽装（同一 vtable を全ハンドラで共有）
Private m_DTPEventHandlerThis As VtblObj
Private m_DTPEventHandlerVT As VtblData4
Private Const MAX_DTP_EVENT_HANDLERS As Long = 256                              '一度に非同期イベントを購読する数を設定
Private m_DTPEventHandlerThisArray(0 To MAX_DTP_EVENT_HANDLERS - 1) As VtblObj  ' 固定領域（再配置しない）
Private m_DTPEventHandlerThisCount As Long

' ICoreWebView2::GetDevToolsProtocolEventReceiver の vtable index（既定値は推定。必要なら差し替え可）
Private m_DevToolsProtocolEventReceiverVTableIndex As Long

Private Type Guid16Core
    d(0 To 15) As Byte
End Type

' =====================================================================
' Public Events
' =====================================================================
Public Event Ready()
Public Event NavigationCompleted(ByVal isSuccess As Boolean, ByVal webErrorStatus As Long)
Public Event InitializeFailed(ByVal ErrorCode As Long, ByVal Description As String)

' CDP の非同期イベント
'   methodName … ドメイン付きメソッド名（例: Network.responseReceived）
'   eventRawJson … WebView2 が渡す params 部分の JSON 文字列をそのまま（再シリアライズしない）
Public Event CDPEventReceived(methodName As String, eventParamsJson As String)

' 親ウィンドウ HWND を外部から参照するためのプロパティ
Public Property Get ParentHwnd() As LongPtr
    ParentHwnd = m_hWndParent
End Property

' WebView2 ブラウザプロセス ID（ICoreWebView2::get_BrowserProcessId）。未作成時は 0。
Public Property Get dwProcessId() As Long
    If m_pWebView2 = 0 Then
        dwProcessId = 0
        Exit Property
    End If
    Dim pid As Long
    Dim vPid(0 To 1) As LongPtr
    Dim vtII As Integer: vtII = 20  ' VT_I8
    CopyMemory vPid(0), vtII, 2
    vPid(1) = VarPtr(pid)
    Dim argT(0) As Integer: argT(0) = 20
    Dim argP(0) As LongPtr: argP(0) = VarPtr(vPid(0))
    Dim retV As Variant
    DispCallFunc m_pWebView2, OFF_WV2_get_BrowserProcessId, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
    dwProcessId = pid
End Property

Private Function IID_IUnknown_Core() As Guid16Core
    ' {00000000-0000-0000-C000-000000000046}
    Static g As Guid16Core
    Static inited As Boolean
    If Not inited Then
        g.d(6) = &HC0
        g.d(14) = &H46
        inited = True
    End If
    IID_IUnknown_Core = g
End Function

Private Function COM_Release(ByVal pUnk As LongPtr) As Long
    If pUnk = 0 Then Exit Function
    Dim retV As Variant
    Dim hr As Long
    hr = DispCallFunc(pUnk, 2 * PTR_SIZE, CC_STDCALL, vbLong, 0, 0, 0, retV)
    If hr <> 0 Then Exit Function
    COM_Release = CLng(retV)
End Function

Private Function COM_ToIUnknownPtr(ByVal pObj As LongPtr) As LongPtr
    If pObj = 0 Then Exit Function

    Dim iid As Guid16Core
    iid = IID_IUnknown_Core()

    Dim pUnk As LongPtr: pUnk = 0
    Dim vtII As Integer: vtII = 20 'VT_I8
    Dim vRiid(0 To 1) As LongPtr
    Dim vOut(0 To 1) As LongPtr
    CopyMemory vRiid(0), vtII, 2
    vRiid(1) = VarPtr(iid)
    CopyMemory vOut(0), vtII, 2
    vOut(1) = VarPtr(pUnk)

    Dim argT(0 To 1) As Integer
    Dim argP(0 To 1) As LongPtr
    argT(0) = vbLongLong: argP(0) = VarPtr(vRiid(0))
    argT(1) = vbLongLong: argP(1) = VarPtr(vOut(0))

    Dim retV As Variant
    Dim hr As Long
    hr = DispCallFunc(pObj, 0 * PTR_SIZE, CC_STDCALL, vbLong, 2, argT(0), argP(0), retV)
    If hr <> 0 Then Exit Function
    If CLng(retV) <> 0 Then Exit Function

    COM_ToIUnknownPtr = pUnk
End Function



'***************************************************************************************************
'                                   ■■■ 初期化関連 ■■■
'***************************************************************************************************
'* 機能　　：WebView2を起動します
'---------------------------------------------------------------------------------------------------
'* 返値　　：True = 初期化コマンド送信成功（Ready イベント待ち）
'* 引数　　：hWndParent                     WebView2 を埋め込む親ウィンドウハンドル
'            boundsLeft/Top/Width/Height    表示領域（ピクセル）
'            initialUrl                     最初に遷移するURL（省略すると`about:blank`とします）
'            additionalBrowserArguments     追加ブラウザ引数（省略可）。例: "--disable-web-security"
'---------------------------------------------------------------------------------------------------
'* 注意事項：現時点では、`CoreWebView2EnvironmentOptions`による引数直接指定が効かないため、環境変数で間接的に付与しています
'***************************************************************************************************
Public Function Initialize( _
    ByVal hwndParent As LongPtr, _
    ByVal boundsLeft As Long, _
    ByVal boundsTop As Long, _
    ByVal boundsWidth As Long, _
    ByVal boundsHeight As Long, _
    Optional ByVal initialUrl As String = "about:blank", Optional UserDataName As String, Optional additionalBrowserArguments As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".Initialize"


    m_hWndParent = hwndParent
    m_InitialUrl = initialUrl
    m_Bounds.Left = boundsLeft
    m_Bounds.Top = boundsTop
    m_Bounds.Right = boundsLeft + boundsWidth
    m_Bounds.Bottom = boundsTop + boundsHeight

    ' グローバル参照をセット（コールバックから Me へ戻るため）
    Set g_WebView2Core = Me

    ' vtable を構築
    BuildVtables

    ' 環境作成開始（非同期; WV2_EnvCB_Invoke が呼ばれると CB_EnvironmentCreated が動く）
    Dim hr As Long
    Dim userDataPath As String
    
    'Excelが使うWebView2の作業ディレクトリに合わせる
    userDataPath = Environ("LOCALAPPDATA") & "\Microsoft\Office\" & Application.Version & "\" & WorkSpaceFolderName & "\" & UserDataName
    printMsg info_, "WebView2 working folder: " & userDataPath, FromProcedureName

    Dim pOptions As LongPtr
    Dim useOptions As Boolean
    If Len(additionalBrowserArguments) > 0 Then
        g_EnvOptions_AdditionalArgs = additionalBrowserArguments
        printMsg info_, "Additional startup arguments for WebView2: " & g_EnvOptions_AdditionalArgs, FromProcedureName

        BuildEnvOptionsVtable
        pOptions = VarPtr(m_EnvOptionsThis)
        useOptions = True
    Else
        pOptions = 0
        useOptions = False
        
        ' この起動では追加引数を使わないので、前回のフォールバック設定をクリア
        SetEnvironmentVariableW StrPtr("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"), 0
        printMsg info_, "There are no additional startup arguments for WebView2.", FromProcedureName
    End If

    Dim pUserData   As LongPtr: pUserData = StrPtr(userDataPath)
    Dim pHandler    As LongPtr: pHandler = VarPtr(m_EnvHandlerThis)

    hr = CreateCoreWebView2EnvironmentWithOptions(0, pUserData, pOptions, pHandler)

    ' E_INVALIDARG: 偽装 COM がマーシャリング等で拒否される。環境変数でフォールバック。
    ' pOptions=0 のとき、ローダーは WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS を参照する。
    If hr <> 0 And useOptions And hr = &H80070057 Then
        printMsg WARN_, "Failed to launch WebView2 with arguments. Attempting to provide them indirectly via environment variables. ErrorCode: " & hr, FromProcedureName

        SetEnvironmentVariableW StrPtr("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"), StrPtr(additionalBrowserArguments)
        hr = CreateCoreWebView2EnvironmentWithOptions(0, pUserData, 0, pHandler)
        ' この時点で環境作成は開始されたので、他の Excel 内 WebView2 への影響を避けるため即座にクリア
        SetEnvironmentVariableW StrPtr("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"), 0
    Else
        printMsg info_, "WebView2 launched successfully. resultCode: " & hr, FromProcedureName
    End If

    If hr <> 0 Then
        printMsg ERROR_, "WebView2 failed to start. ErrorCode: " & hr, FromProcedureName

        m_LastErrorCode = hr
        If Len(m_LastErrorDesc) = 0 Then m_LastErrorDesc = "CreateCoreWebView2EnvironmentWithOptions failed: 0x" & Hex(hr)
        RaiseEvent InitializeFailed(hr, m_LastErrorDesc)
        Initialize = False
    Else
        Initialize = True
    End If
End Function



'***************************************************************************************************
'                       ■■■ 呼び出し側で実際に使うメソッド関連 ■■■
'***************************************************************************************************
'* 機能　　：CDP メソッド実行。結果 JSON を同期待ちで返す
'---------------------------------------------------------------------------------------------------
'* 返値　　：CDP メソッド実行
'* 引数　　：methodName "Domain.method"（例: "Runtime.evaluate", "Page.navigate"）
'            paramsJson `parameters`部分のJson文字列を指定します。空の場合は"{}"で
'---------------------------------------------------------------------------------------------------
'* 注意事項：・Immediate に [CDP] Invoke が出なければコールバック未呼び出し（vtable 要確認）
'            ・Edge×Pipe×CDP の場合、IDやSessionID等の結果フルセットが届きますが、WebView2の場合、
'              →CDP成功：Resultオブジェクトの中身のみ
'              　CDP失敗：IDやSessionID等の結果フルセット
'              となります｡取り扱いには､ご注意下さい｡
'***************************************************************************************************
Public Function CallDevToolsProtocolMethod(ByVal methodName As String, ByVal paramsJson As String) As String
    Const FromProcedureName As String = ThisClassName & ".CallDevToolsProtocolMethod"


    If m_pWebView2 = 0 Then CallDevToolsProtocolMethod = "": Exit Function
    '同期呼び出しは「次に届くコマンド結果」だけを返す
    '（イベントが混ざって届いても誤判定しないよう、キューで分離します）
    If m_CDPResultQueue Is Nothing Then Set m_CDPResultQueue = New Collection
    m_LastCDPResultErrorCode = 0

    Dim hr As Long
    hr = COM_CallDevToolsProtocolMethod(m_pWebView2, methodName, paramsJson, VarPtr(m_CDPHandlerThis))
    printMsg Trace_, "Execution parameter of `CallDevToolsProtocolMethod`: " & WorksheetFunction.TextJoin(" , ", True, "m_pWebView2-" & m_pWebView2, "methodName-" & methodName, "paramsJson-" & paramsJson, "VarPtr(m_CDPHandlerThis)-" & VarPtr(m_CDPHandlerThis)), FromProcedureName
    printMsg Debug_, "Execution result of `CallDevToolsProtocolMethod`: " & hr, FromProcedureName
    If hr <> 0 Then
        m_LastErrorCode = hr
        CallDevToolsProtocolMethod = "": Exit Function
    End If
    Dim t As Single: t = Timer
    Dim spin As Long
    spin = 0
    printMsg Debug_, "CDP sync wait: start method=" & methodName & " timeoutSec=" & CStr(ResultCDPTimeout), FromProcedureName
    Do
        spin = spin + 1
        If spin = 1 Or (spin Mod 100) = 0 Then
            Dim qCnt As Long
            If m_CDPResultQueue Is Nothing Then
                qCnt = -1
            Else
                qCnt = m_CDPResultQueue.Count
            End If
            printMsg Trace_, "CDP sync wait: spin=" & spin & " method=" & methodName & " elapsedSec=" & Format$(CSng(Timer - t), "0.00") & " queueCount=" & qCnt, FromProcedureName
        End If
        ProcessMessages
        DoEvents

        If Not m_CDPResultQueue Is Nothing Then
            If m_CDPResultQueue.Count > 0 Then
                CallDevToolsProtocolMethod = CStr(m_CDPResultQueue(1))
                m_CDPResultQueue.Remove 1
                If m_LastCDPResultErrorCode <> 0 Then m_LastErrorCode = m_LastCDPResultErrorCode
                printMsg Debug_, "CDP sync wait: done method=" & methodName & " spins=" & spin, FromProcedureName
                Exit Function
            End If
        End If
    Loop While (Timer - t) < ResultCDPTimeout

    printMsg WARN_, "CDP sync wait: TIMEOUT method=" & methodName & " spins=" & spin & " timeoutSec=" & CStr(ResultCDPTimeout), FromProcedureName
    If WV2logView Then Debug.Print "[CDP] timeout: command result not received from queue (try other vtable index 22,36,37,39)"
    CallDevToolsProtocolMethod = ""
End Function

'***************************************************************************************************
'* 機能　　：CDP メソッド実行（結果待ちしない）
'---------------------------------------------------------------------------------------------------
'* 返値　　：COM 呼び出し HRESULT（0=成功）
'* 引数　　：methodName "Domain.method"（例: "Runtime.evaluate", "Page.navigate"）
'            paramsJson `parameters`部分のJson文字列を指定します。空の場合は"{}"で
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'***************************************************************************************************
Public Function CallDevToolsProtocolMethodAsync(ByVal methodName As String, ByVal paramsJson As String, ByVal CommandID As Long) As Long
    Const FromProcedureName As String = ThisClassName & ".CallDevToolsProtocolMethodAsync"
    
    
    If m_pWebView2 = 0 Then
        CallDevToolsProtocolMethodAsync = &H80004005 'E_FAIL
        Exit Function
    End If

    If m_CDPAsyncPThisHexToInfo Is Nothing Then Set m_CDPAsyncPThisHexToInfo = New Dictionary

    '空いている async slot を確保し、コールバックの pThis を cmd ごとにユニーク化する
    Dim idx As Long: idx = -1
    Dim pThis As LongPtr: pThis = 0
    Dim keyHex As String: keyHex = vbNullString
    Dim i As Long
    For i = 0 To MAX_ASYNC_CDP_COMMANDS - 1
        If m_CDPAsyncHandlerUsed(i) = False Then
            idx = i
            m_CDPAsyncHandlerUsed(i) = True
            pThis = VarPtr(m_CDPAsyncHandlerThisPool(i))
            keyHex = PtrToHex(pThis)
            m_CDPAsyncPThisHexToInfo(keyHex) = Array(i, CommandID)
            Exit For
        End If
    Next i

    If idx < 0 Then
        m_LastErrorCode = &H8007000E 'E_OUTOFMEMORY (approx)
        printMsg WARN_, "CallDevToolsProtocolMethodAsync: no free async slot (CommandID=" & CommandID & ")", FromProcedureName
        CallDevToolsProtocolMethodAsync = m_LastErrorCode
        Exit Function
    End If

    Dim hr As Long
    hr = COM_CallDevToolsProtocolMethod(m_pWebView2, methodName, paramsJson, pThis)
    printMsg Trace_, "Execution parameter of `CallDevToolsProtocolMethodAsync`: " & WorksheetFunction.TextJoin(" , ", True, "m_pWebView2-" & m_pWebView2, "methodName-" & methodName, "paramsJson-" & paramsJson, "pThis(async)-" & pThis, "CommandID-" & CommandID, "slotIdx-" & idx), FromProcedureName
    printMsg Debug_, "CallDevToolsProtocolMethodAsync: COM hr=" & hr & " CommandID=" & CommandID & " slotIdx=" & idx & " pThis=" & pThis & " method=" & methodName, FromProcedureName

    If hr <> 0 Then
        m_LastErrorCode = hr
        m_CDPAsyncHandlerUsed(idx) = False
        m_CDPAsyncPThisHexToInfo.Remove keyHex
    End If
    CallDevToolsProtocolMethodAsync = hr
End Function

'***************************************************************************************************
'* 機能　　：sessionID  を指定して、CDP メソッド実行。結果 JSON を同期待ちで返す
'---------------------------------------------------------------------------------------------------
'* 返値　　：CDP メソッド実行
'* 引数　　：sessionID  操作タブ先のsessionID
'            methodName "Domain.method"（例: "Runtime.evaluate", "Page.navigate"）
'            paramsJson `parameters`部分のJson文字列を指定します。空の場合は"{}"で
'---------------------------------------------------------------------------------------------------
'* 注意事項：・Immediate に [CDP] Invoke が出なければコールバック未呼び出し（vtable 要確認）
'            ・Edge×Pipe×CDP の場合、IDやSessionID等の結果フルセットが届きますが、WebView2の場合、
'              →CDP成功：Resultオブジェクトの中身のみ
'              　CDP失敗：IDやSessionID等の結果フルセット
'              となります｡取り扱いには､ご注意下さい｡
'***************************************************************************************************
Public Function CallDevToolsProtocolMethodForSession(ByVal sessionID As String, ByVal methodName As String, ByVal paramsJson As String) As String
    Const FromProcedureName As String = ThisClassName & ".CallDevToolsProtocolMethodForSession"

    If m_pWebView2 = 0 Then CallDevToolsProtocolMethodForSession = "": Exit Function
    
    '同期呼び出しは「次に届くコマンド結果」だけを返す
    '（イベントが混ざって届いても誤判定しないよう、キューで分離します）
    If m_CDPResultQueue Is Nothing Then Set m_CDPResultQueue = New Collection
    m_LastCDPResultErrorCode = 0

    Dim hr As Long
    hr = COM_CallDevToolsProtocolMethodForSession(m_pWebView2, sessionID, methodName, paramsJson, VarPtr(m_CDPHandlerThis))
    If hr <> 0 Then
        m_LastErrorCode = hr
        CallDevToolsProtocolMethodForSession = "": Exit Function
    End If
    Dim t As Single: t = Timer
    Dim spin As Long
    spin = 0
    printMsg Debug_, "CDP sync wait (session): start sessionID=" & sessionID & " method=" & methodName & " timeoutSec=" & CStr(ResultCDPTimeout), FromProcedureName
    Do
        spin = spin + 1
        If spin = 1 Or (spin Mod 100) = 0 Then
            Dim qCntFs As Long
            If m_CDPResultQueue Is Nothing Then
                qCntFs = -1
            Else
                qCntFs = m_CDPResultQueue.Count
            End If
            printMsg Trace_, "CDP sync wait (session): spin=" & spin & " method=" & methodName & " elapsedSec=" & Format$(CSng(Timer - t), "0.00") & " queueCount=" & qCntFs, FromProcedureName
        End If
        ProcessMessages
        DoEvents

        If Not m_CDPResultQueue Is Nothing Then
            If m_CDPResultQueue.Count > 0 Then
                CallDevToolsProtocolMethodForSession = CStr(m_CDPResultQueue(1))
                m_CDPResultQueue.Remove 1
                If m_LastCDPResultErrorCode <> 0 Then m_LastErrorCode = m_LastCDPResultErrorCode
                printMsg Debug_, "CDP sync wait (session): done sessionID=" & sessionID & " method=" & methodName & " spins=" & spin, FromProcedureName
                Exit Function
            End If
        End If
    Loop While (Timer - t) < ResultCDPTimeout

    printMsg WARN_, "CDP sync wait (session): TIMEOUT sessionID=" & sessionID & " method=" & methodName & " spins=" & spin & " timeoutSec=" & CStr(ResultCDPTimeout), FromProcedureName
    If WV2logView Then Debug.Print "[CDP] timeout: CallDevToolsProtocolMethodForSession (vtable index might be incorrect)"
    CallDevToolsProtocolMethodForSession = ""
End Function

'***************************************************************************************************
'* 機能　　：セッション指定で CDP メソッド実行（結果待ちしない）
'---------------------------------------------------------------------------------------------------
'* 返値　　：COM 呼び出し HRESULT（0=成功）
'* 引数　　：sessionID  操作タブ先のsessionID
'            methodName "Domain.method"（例: "Runtime.evaluate", "Page.navigate"）
'            paramsJson `parameters`部分のJson文字列を指定します。空の場合は"{}"で
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'***************************************************************************************************
Public Function CallDevToolsProtocolMethodForSessionAsync(ByVal sessionID As String, ByVal methodName As String, ByVal paramsJson As String, ByVal CommandID As Long) As Long
    Const FromProcedureName As String = ThisClassName & ".CallDevToolsProtocolMethodForSessionAsync"

    If m_pWebView2 = 0 Then
        CallDevToolsProtocolMethodForSessionAsync = &H80004005 'E_FAIL
        Exit Function
    End If

    If m_CDPAsyncPThisHexToInfo Is Nothing Then Set m_CDPAsyncPThisHexToInfo = New Dictionary

    '空いている async slot を確保し、コールバックの pThis を cmd ごとにユニーク化する
    Dim idx As Long: idx = -1
    Dim pThis As LongPtr: pThis = 0
    Dim keyHex As String: keyHex = vbNullString
    Dim i As Long
    For i = 0 To MAX_ASYNC_CDP_COMMANDS - 1
        If m_CDPAsyncHandlerUsed(i) = False Then
            idx = i
            m_CDPAsyncHandlerUsed(i) = True
            pThis = VarPtr(m_CDPAsyncHandlerThisPool(i))
            keyHex = PtrToHex(pThis)
            m_CDPAsyncPThisHexToInfo(keyHex) = Array(i, CommandID)
            Exit For
        End If
    Next i

    If idx < 0 Then
        m_LastErrorCode = &H8007000E 'E_OUTOFMEMORY (approx)
        printMsg WARN_, "CallDevToolsProtocolMethodForSessionAsync: no free async slot (CommandID=" & CommandID & ")", FromProcedureName
        CallDevToolsProtocolMethodForSessionAsync = m_LastErrorCode
        Exit Function
    End If

    Dim hr As Long
    hr = COM_CallDevToolsProtocolMethodForSession(m_pWebView2, sessionID, methodName, paramsJson, pThis)
    printMsg Trace_, "Execution parameter of `CallDevToolsProtocolMethodForSessionAsync`: " & WorksheetFunction.TextJoin(" , ", True, "sessionID-" & sessionID, "methodName-" & methodName, "paramsJson-" & paramsJson, "pThis(async)-" & pThis, "CommandID-" & CommandID, "slotIdx-" & idx), FromProcedureName
    printMsg Debug_, "CallDevToolsProtocolMethodForSessionAsync: COM hr=" & hr & " CommandID=" & CommandID & " slotIdx=" & idx & " pThis=" & pThis & " method=" & methodName, FromProcedureName

    If hr <> 0 Then
        m_LastErrorCode = hr
        m_CDPAsyncHandlerUsed(idx) = False
        m_CDPAsyncPThisHexToInfo.Remove keyHex
    End If
    CallDevToolsProtocolMethodForSessionAsync = hr
End Function

'----------------------------------------------------------------------
' Navigate : 指定 URL へ遷移
'----------------------------------------------------------------------
Public Sub navigate(ByVal Url As String)
    If m_pWebView2 = 0 Then Exit Sub
    Dim hr As Long
    hr = COM_Navigate(m_pWebView2, Url)
End Sub

'----------------------------------------------------------------------
' NavigateToString : HTML 文字列を直接描画
'----------------------------------------------------------------------
Public Sub NavigateToString(ByVal htmlContent As String)
    If m_pWebView2 = 0 Then Exit Sub
    COM_DispCallPtr1 m_pWebView2, OFF_WV2_NavigateToString, StrPtr(htmlContent)
End Sub



'***************************************************************************************************
'                       ■■■ ブラウザからの非同期イベント購読関連 ■■■
'***************************************************************************************************
'* 機能　　：WebView2 の DevToolsProtocol event (非同期イベント)を購読します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：登録成功可否論理値
'* 引数　　：eventName      購読したい非同期イベントmethod名
'---------------------------------------------------------------------------------------------------
'* 注意事項：・メモリ変動によるアクセス違反を防ぐため同時に購読できるmethod数は、「MAX_DTP_EVENT_HANDLERS」に依存します。
'            　デフォルトは、256個まで一括購読できます。この上限は、変更可能ですが、動いてる最中は変更できません。
'
'            ・この購読が働いている状態で、VBAのブレークポイント等で動作を止めると、`ntdll.dll - 0xc0000027`エラーによるクラッシュが発生します。
'              これを回避するには、次項の購読解除処理が必要です。
'***************************************************************************************************
Public Function SubscribeDevToolsProtocolEvent(ByVal eventName As String) As Boolean
    '機能が整ってない場合は、即抜
    If m_pWebView2 = 0 Then SubscribeDevToolsProtocolEvent = False: Exit Function

    '空文字も同様
    If Len(eventName) = 0 Then SubscribeDevToolsProtocolEvent = False: Exit Function

    '初購読は、Dictionaryオブジェクトを作成
    If m_DevToolsProtocolEventSubscriptions Is Nothing Then Set m_DevToolsProtocolEventSubscriptions = New Dictionary

    '既に登録済みなら、即抜
    If m_DevToolsProtocolEventSubscriptions.Exists(eventName) Then SubscribeDevToolsProtocolEvent = True: Exit Function

    '購読の上限に突破した場合は、ここで抜ける
    Dim idx As Long: idx = m_DTPEventHandlerThisCount
    If idx >= MAX_DTP_EVENT_HANDLERS Then SubscribeDevToolsProtocolEvent = False: Exit Function

    ' イベントごとに別の handler this を用意（コールバックの pThis で振り分けするため）し、カウント
    m_DTPEventHandlerThisArray(idx).pVTable = VarPtr(m_DTPEventHandlerVT.fn(0))
    Dim pHandlerThis As LongPtr: pHandlerThis = VarPtr(m_DTPEventHandlerThisArray(idx))
    m_DTPEventHandlerThisCount = m_DTPEventHandlerThisCount + 1

    'receiver にイベントハンドラ登録（このイベント専用の pHandlerThis を渡す）
    Const OFF_RCV_add_DevToolsProtocolEventReceived As Long = 3 * 8 'index 3
    Dim token As LongPtr: token = 0

    'receiver を取得（イベント名ごと）
    Dim pReceiver As LongPtr
    pReceiver = COM_GetDevToolsProtocolEventReceiver(m_pWebView2, eventName)
    If pReceiver = 0 Then
        SubscribeDevToolsProtocolEvent = False
        Exit Function
    End If

    Dim vHandler(0 To 1) As LongPtr
    Dim vToken(0 To 1) As LongPtr
    Dim vtII As Integer: vtII = 20 'VT_I8

    CopyMemory vHandler(0), vtII, 2
    vHandler(1) = pHandlerThis

    CopyMemory vToken(0), vtII, 2
    vToken(1) = VarPtr(token)

    Dim argT(0 To 1) As Integer
    Dim argP(0 To 1) As LongPtr
    argT(0) = vbLongLong: argP(0) = VarPtr(vHandler(0))
    argT(1) = vbLongLong: argP(1) = VarPtr(vToken(0))

    Dim retV As Variant
    Dim hr As Long
    hr = DispCallFunc(pReceiver, OFF_RCV_add_DevToolsProtocolEventReceived, CC_STDCALL, vbLong, 2, argT(0), argP(0), retV)
    If hr <> 0 Then
        SubscribeDevToolsProtocolEvent = False
        Exit Function
    End If

    '複数購読用に保存（receiver, token, handlerThis の3要素）
    Dim subInfo(0 To 2) As LongPtr
    subInfo(0) = pReceiver
    subInfo(1) = token
    subInfo(2) = pHandlerThis
    m_DevToolsProtocolEventSubscriptions.Add eventName, subInfo

    ' pThis -> eventName（コールバックで必ずこちらを参照して振り分け）
    If m_DevToolsProtocolHandlerThisToEventName Is Nothing Then
        Set m_DevToolsProtocolHandlerThisToEventName = New Dictionary
    End If
    m_DevToolsProtocolHandlerThisToEventName(PtrToHex(pHandlerThis)) = eventName

    If m_DevToolsProtocolEventReceiverHexToName Is Nothing Then
        Set m_DevToolsProtocolEventReceiverHexToName = New Dictionary
    End If
    m_DevToolsProtocolEventReceiverHexToName(PtrToHex(pReceiver)) = eventName
    Dim pReceiverUnk As LongPtr
    pReceiverUnk = COM_ToIUnknownPtr(pReceiver)
    If pReceiverUnk <> 0 Then
        m_DevToolsProtocolEventReceiverHexToName(PtrToHex(pReceiverUnk)) = eventName
        COM_Release pReceiverUnk
    End If

    SubscribeDevToolsProtocolEvent = True
End Function

'***************************************************************************************************
'* 機能　　：WebView2 の DevToolsProtocol event (非同期イベント)購読を解除します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：解除成功可否論理値
'* 引数　　：eventName      購読解除したい非同期イベントmethod名　※省略で、全解除します
'***************************************************************************************************
Public Function UnsubscribeDevToolsProtocolEvent(Optional ByVal eventName As String) As Boolean
    On Error GoTo iferr
    If m_DevToolsProtocolEventSubscriptions Is Nothing Then UnsubscribeDevToolsProtocolEvent = True: Exit Function
    If m_DevToolsProtocolEventSubscriptions.Count = 0 Then UnsubscribeDevToolsProtocolEvent = True: Exit Function

    ' ICoreWebView2DevToolsProtocolEventReceiver::remove_DevToolsProtocolEventReceived
    ' vtable: IUnknown(0..2), add(3), remove(4)
    Const OFF_RCV_remove_DevToolsProtocolEventReceived As Long = 4 * 8 'index 4

    '引数なし => 全解除
    If StrPtr(eventName) = 0 Then
        Dim k As Variant
        For Each k In m_DevToolsProtocolEventSubscriptions.keys
            Dim subInfoAll As Variant: subInfoAll = m_DevToolsProtocolEventSubscriptions(k)
            Dim pReceiver As LongPtr: pReceiver = subInfoAll(0)
            Dim token As LongPtr: token = subInfoAll(1)
            Dim pHandlerThisAll As LongPtr: pHandlerThisAll = subInfoAll(2)

            Dim vtII As Integer: vtII = 20 'VT_I8
            Dim vToken(0 To 1) As LongPtr
            CopyMemory vToken(0), vtII, 2
            vToken(1) = token

            Dim argT(0) As Integer
            Dim argP(0) As LongPtr
            argT(0) = vbLongLong: argP(0) = VarPtr(vToken(0))

            Dim retV As Variant
            Dim hr As Long
            hr = DispCallFunc(pReceiver, OFF_RCV_remove_DevToolsProtocolEventReceived, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV)
            If Not m_DevToolsProtocolHandlerThisToEventName Is Nothing Then
                On Error Resume Next
                m_DevToolsProtocolHandlerThisToEventName.Remove PtrToHex(pHandlerThisAll)
                On Error GoTo 0
            End If
            ' hr != 0 でも最後まで処理する
        Next

        Set m_DevToolsProtocolEventSubscriptions = New Dictionary
        If Not m_DevToolsProtocolEventReceiverHexToName Is Nothing Then Set m_DevToolsProtocolEventReceiverHexToName = New Dictionary
        If Not m_DevToolsProtocolHandlerThisToEventName Is Nothing Then Set m_DevToolsProtocolHandlerThisToEventName = New Dictionary
        m_DTPEventHandlerThisCount = 0
        UnsubscribeDevToolsProtocolEvent = True
        Exit Function
    End If

    '指定 event のみ解除
    If Not m_DevToolsProtocolEventSubscriptions.Exists(eventName) Then
        UnsubscribeDevToolsProtocolEvent = True
        Exit Function
    End If

    Dim subInfoOne As Variant: subInfoOne = m_DevToolsProtocolEventSubscriptions(eventName)
    Dim pReceiverOne As LongPtr: pReceiverOne = subInfoOne(0)
    Dim tokenOne As LongPtr: tokenOne = subInfoOne(1)
    Dim pHandlerThisOne As LongPtr: pHandlerThisOne = subInfoOne(2)

    vtII = 20 'VT_I8
    Dim vTokenOne(0 To 1) As LongPtr
    CopyMemory vTokenOne(0), vtII, 2
    vTokenOne(1) = tokenOne

    Dim argTOne(0) As Integer
    Dim argPOne(0) As LongPtr
    argTOne(0) = vbLongLong: argPOne(0) = VarPtr(vTokenOne(0))

    Dim retVOne As Variant
    Dim hrOne As Long
    hrOne = DispCallFunc(pReceiverOne, OFF_RCV_remove_DevToolsProtocolEventReceived, CC_STDCALL, vbLong, 1, argTOne(0), argPOne(0), retVOne)

    m_DevToolsProtocolEventSubscriptions.Remove eventName
    If Not m_DevToolsProtocolHandlerThisToEventName Is Nothing Then
        On Error Resume Next
        m_DevToolsProtocolHandlerThisToEventName.Remove PtrToHex(pHandlerThisOne)
        On Error GoTo iferr
    End If
    If Not m_DevToolsProtocolEventReceiverHexToName Is Nothing Then
        On Error Resume Next
        m_DevToolsProtocolEventReceiverHexToName.Remove PtrToHex(pReceiverOne)
        On Error GoTo iferr
    End If

    UnsubscribeDevToolsProtocolEvent = (hrOne = 0)
    Exit Function

iferr:
    UnsubscribeDevToolsProtocolEvent = False
End Function

'***************************************************************************************************
'* 機能　　：WebView2 の DevToolsProtocol event method名に紐づくキーを生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：そのポインターの16進数表記文字列
'* 引数　　：method名に紐づく発信ポインター値
'***************************************************************************************************
Private Function PtrToHex(ByVal p As LongPtr) As String
    'LongPtr を Dictionary のキーにするための簡易変換
    PtrToHex = "0x" & Hex$(p)
End Function



'***************************************************************************************************
'                               ■■■ ウィンドウ周り関連 ■■■
'***************************************************************************************************
'----------------------------------------------------------------------
' Resize : 表示領域を変更
'----------------------------------------------------------------------
Public Sub Resize(ByVal newLeft As Long, ByVal newTop As Long, _
                  ByVal newWidth As Long, ByVal newHeight As Long)
    m_Bounds.Left = newLeft
    m_Bounds.Top = newTop
    m_Bounds.Right = newLeft + newWidth
    m_Bounds.Bottom = newTop + newHeight
    If m_pController <> 0 Then COM_SetBounds m_pController, m_Bounds
End Sub

'----------------------------------------------------------------------
' NotifyPositionChanged
'   親ウィンドウが移動またはリサイズされたときに呼び出す。
'   WebView2 内部の IME 位置・ドロップダウン位置などを正しく再計算させる。
'   Resize の後に呼び出すことを推奨。
'----------------------------------------------------------------------
Public Sub NotifyPositionChanged()
    If m_pController = 0 Then Exit Sub
    Dim retV As Variant
    DispCallFunc m_pController, OFF_CTRL_NotifyParentPos, CC_STDCALL, vbLong, 0, 0, 0, retV
End Sub

'----------------------------------------------------------------------
' ResizeDirect  ― MoveWindow で WebView2 子 HWND を直接リサイズ
'   EdgeInExcelForm.frm と同じ方式（put_Bounds の代わり）。
'   m_hWndParent（= Frame hWnd）の最初の子ウィンドウが WebView2 HWND。
'----------------------------------------------------------------------
Public Sub ResizeDirect(ByVal pxWidth As Long, ByVal pxHeight As Long)
    If m_hWndParent = 0 Then Exit Sub
    Const GW_CHILD As Long = 5
    Dim hWv2 As LongPtr
    hWv2 = GetWindow(m_hWndParent, GW_CHILD)
    If hWv2 = 0 Then Exit Sub
    MoveWindow hWv2, 0, 0, pxWidth, pxHeight, 1
End Sub

'----------------------------------------------------------------------
' ResizeToParent  ― Frame HWND の GetClientRect 実ピクセルで put_Bounds
'   hWndFrame : UserForm 側で取得したフレッシュな Frame hWnd
'   ★ WM_SIZE ハンドラの内側から呼ぶと EmbeddedBrowserWebView.dll がクラッシュする
'     → 安全な呼び切りは ScheduleResize を使くこと
'----------------------------------------------------------------------
Public Sub ResizeToParent(ByVal hWndFrame As LongPtr)
    If m_pController = 0 Then Exit Sub
    If hWndFrame = 0 Then Exit Sub

    Dim rc As RECT_T
    Dim gcResult As Long
    gcResult = GetClientRect(hWndFrame, rc)
    If WV2logView Then Debug.Print "[WV2] ResizeToParent: hFrame=0x" & Hex(hWndFrame) & _
                ", GetClientRect=" & gcResult & _
                ", rect={" & rc.Left & "," & rc.Top & "," & rc.Right & "," & rc.Bottom & "}"

    If gcResult = 0 Then Exit Sub
    If rc.Right <= 0 Or rc.Bottom <= 0 Then Exit Sub

    m_Bounds.Left = 0
    m_Bounds.Top = 0
    m_Bounds.Right = rc.Right
    m_Bounds.Bottom = rc.Bottom

    Dim argT(0) As Integer: argT(0) = vbLongLong
    Dim pBounds As LongPtr: pBounds = VarPtr(m_Bounds)
    Dim argP(0) As LongPtr: argP(0) = VarPtr(pBounds)
    Dim retV As Variant
    DispCallFunc m_pController, OFF_CTRL_put_Bounds, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
    If WV2logView Then Debug.Print "  put_Bounds({0,0," & rc.Right & "," & rc.Bottom & "}) HRESULT=0x" & Hex(CLng(retV))
End Sub

'----------------------------------------------------------------------
' ScheduleResize  ― UserForm_Resize から呼ぶ。
'   hWndFrame(Frame の HWND)を保存し、1ms タイマーをセットするだけ。
'   GetClientRect は WM_SIZE 完了後のタイマーコールバック（DoTimerResize）で
'   実行するため、このメソッド内では Win32 サイズ取得を行わない。
'----------------------------------------------------------------------
Public Sub ScheduleResize(ByVal hWndFrame As LongPtr)
    If m_pController = 0 Then Exit Sub
    If hWndFrame = 0 Then Exit Sub
    m_hWndFrame = hWndFrame      ' タイマー内で GetClientRect するために保存
    WV2_ScheduleResize
End Sub

'----------------------------------------------------------------------
' DoTimerResize  ― WV2_ResizeTimerProc から呼ばれる。
'   WM_SIZE が完全に処理された後（タイマーコールバック）に GetClientRect を
'   呼び出すことで、遷移中の古い値ではなく確定サイズを取得できる。
'   「前回と同じ bounds」なら put_Bounds をスキップしてフィードバックループを防ぐ。
'----------------------------------------------------------------------
Friend Sub DoTimerResize()
    If m_pController = 0 Then Exit Sub
    If m_hWndFrame = 0 Then Exit Sub

    Dim rc As RECT_T
    If GetClientRect(m_hWndFrame, rc) = 0 Then Exit Sub
    If rc.Right <= 0 Or rc.Bottom <= 0 Then Exit Sub

    ' 「前回と同じ値」なら put_Bounds を呼ばない（フィードバックループ防止）
    '   put_Bounds → WM_SIZE → UserForm_Resize → ScheduleResize → DoTimerResize → ...
    If rc.Right = m_LastAppliedRight And rc.Bottom = m_LastAppliedBottom Then
        If WV2logView Then Debug.Print "[WV2] DoTimerResize: スキップ (bounds 変化なし: {0,0," & rc.Right & "," & rc.Bottom & "})"
        Exit Sub
    End If

    m_Bounds.Left = 0
    m_Bounds.Top = 0
    m_Bounds.Right = rc.Right
    m_Bounds.Bottom = rc.Bottom

    If WV2logView Then Debug.Print "[WV2] DoTimerResize: bounds={0,0," & m_Bounds.Right & "," & m_Bounds.Bottom & "}"
    COM_SetBounds m_pController, m_Bounds

    ' 適用済み値を記録（次回のループ検出用）
    m_LastAppliedRight = rc.Right
    m_LastAppliedBottom = rc.Bottom
End Sub



'***************************************************************************************************
'                   ■■■ ブラウザからのCDPメッセージコールバック処理関連 ■■■
'***************************************************************************************************
'* 機能　　：Excel→WebView2へ送ったCDPコマンド結果が返ります
'---------------------------------------------------------------------------------------------------
'* 引数　　：ErrorCode  実行自体の結果コード
'            MsgJson    CDPコマンドの結果Json文字列
'            pThis      どのポインターとして実行したものか？
'---------------------------------------------------------------------------------------------------
'* 機能説明：CDPコマンド結果を処理するプロシージャとなります
'* 注意事項：・Edge×Pipe×CDP の場合、IDやSessionID等の結果フルセットが届きますが、WebView2の場合、
'              →CDP成功：Resultオブジェクトの中身のみ
'              　CDP失敗：IDやSessionID等の結果フルセット
'              となります｡取り扱いには､ご注意下さい｡
'
'            ・非同期イベントは、`CB_OnDevToolsProtocolEvent`プロシージャで別途処理を設けています。ここのプロシージャはあくまでも、コマンド結果のみ担います。
'***************************************************************************************************
Friend Sub CB_OnCDPMessage(ByVal ErrorCode As Long, ByVal MsgJson As String, ByVal pThis As LongPtr)
    Const FromProcedureName As String = ThisClassName & ".CB_OnCDPMessage"

    '空文字の場合はそもそも、通さないようにする
    If Len(MsgJson) = 0 Then
        printMsg Debug_, "CB_OnCDPMessage: empty MsgJson (ignored), pThis=" & pThis & " CDP-ErrorCode=" & ErrorCode, FromProcedureName
        Exit Sub
    End If
    If m_CDPResultQueue Is Nothing Then Set m_CDPResultQueue = New Collection
    printMsg Trace_, "CDP-Result: " & MsgJson & " , pThis: " & pThis & " , CDP-ErrorCode: " & ErrorCode, FromProcedureName
    printMsg Debug_, "CB_OnCDPMessage: enter pThis=" & pThis & " syncThis=" & VarPtr(m_CDPHandlerThis) & " CDP-ErrorCode=" & ErrorCode & " jsonLen=" & Len(MsgJson), FromProcedureName

    '発信元のポインターをチェックし、同期モードか非同期モードかを確認
    If pThis = VarPtr(m_CDPHandlerThis) Then
        '同期結果(CDPコマンド結果)
        printMsg Debug_, "CDP-Result: " & MsgJson & " , pThis(Sync mode): " & pThis & " , CDP-ErrorCode: " & ErrorCode, FromProcedureName
        m_LastCDPResultErrorCode = ErrorCode
        m_CDPResultQueue.Add MsgJson
        Exit Sub
    End If

    Dim keyHex As String: keyHex = PtrToHex(pThis)

    '非同期結果：pThis(=投げた文脈)から cmdId を復元
    If m_CDPAsyncPThisHexToInfo Is Nothing Then
        printMsg WARN_, "CB_OnCDPMessage: async pThis but m_CDPAsyncPThisHexToInfo is Nothing. pThis=" & pThis & " keyHex=" & keyHex, FromProcedureName
        Exit Sub
    End If

    printMsg Debug_, "CDP-Result: " & MsgJson & " , pThis(Async mode): " & pThis & " , CDP-ErrorCode: " & ErrorCode, FromProcedureName

    If m_CDPAsyncPThisHexToInfo.Exists(keyHex) Then
        Dim info As Variant: info = m_CDPAsyncPThisHexToInfo(keyHex)
        Dim idx As Long: idx = CLng(info(0))
        Dim cmdId As Long: cmdId = CLng(info(1))

        m_CDPAsyncResultJsonByCmdId(cmdId) = MsgJson
        m_CDPAsyncErrorCodeByCmdId(cmdId) = ErrorCode

        ' slot free（結果は辞書側に残す）
        m_CDPAsyncHandlerUsed(idx) = False
        m_CDPAsyncPThisHexToInfo.Remove keyHex
        printMsg Debug_, "CB_OnCDPMessage: async result stored cmdId=" & cmdId & " slotIdx=" & idx & " keyHex=" & keyHex, FromProcedureName
        Exit Sub
    End If

    printMsg WARN_, "CB_OnCDPMessage: unmatched pThis (not sync, unknown async slot). pThis=" & pThis & " keyHex=" & keyHex & " CDP-ErrorCode=" & ErrorCode, FromProcedureName
End Sub

'----------------------------------------------------------------------
' ResultCDPForAsync
'   Async コマンド投げた側の cmdId で結果を後回収します。
'   WebView2 の場合、結果 JSON は MsgJson をそのまま保持しています。
'----------------------------------------------------------------------
Public Function ResultCDPForAsync(ByVal CommandID As Long, Optional ByRef ErrorExist As Boolean) As String
    Const FromProcedureName As String = ThisClassName & ".ResultCDPForAsync"

    ErrorExist = False
    If m_CDPAsyncResultJsonByCmdId Is Nothing Then
        printMsg Debug_, "ResultCDPForAsync: miss (result dict not initialized) CommandID=" & CommandID, FromProcedureName
        Exit Function
    End If

    If Not m_CDPAsyncResultJsonByCmdId.Exists(CommandID) Then
        printMsg Debug_, "ResultCDPForAsync: miss (not ready yet) CommandID=" & CommandID, FromProcedureName
        Exit Function
    End If

    ResultCDPForAsync = CStr(m_CDPAsyncResultJsonByCmdId(CommandID))

    Dim ec As Long: ec = 0
    If Not m_CDPAsyncErrorCodeByCmdId Is Nothing Then
        ec = CLng(m_CDPAsyncErrorCodeByCmdId(CommandID))
        ErrorExist = (ec <> 0)
    End If

    printMsg Debug_, "ResultCDPForAsync: hit CommandID=" & CommandID & " CDP-ErrorCode=" & ec & " jsonLen=" & Len(ResultCDPForAsync), FromProcedureName

    '回収したら破棄（pipe版と同様の取り出し口にする）
    m_CDPAsyncResultJsonByCmdId.Remove CommandID
    If Not m_CDPAsyncErrorCodeByCmdId Is Nothing Then m_CDPAsyncErrorCodeByCmdId.Remove CommandID
End Function

'***************************************************************************************************
'* 機能　　：WebView2 で発生したイベント情報が返ります
'---------------------------------------------------------------------------------------------------
'* 引数　　：ParameterObjectAsJson  非同期イベントJson文字列そのもの※`params`オブジェクトの中身のみ
'            pHandlerThis           どのポインターからの非同期イベントか
'---------------------------------------------------------------------------------------------------
'* 機能説明：WebView2 で発生した非同期イベント情報を処理します。
'            pHandlerThis（登録時に渡した handler の this）で`method`を特定し、
'
'* 注意事項：・Edge×Pipe×CDP の場合、methodやSessionID等の結果フルセットが届きますが、
'              WebView2の場合は、`params`オブジェクトの中身のみ返ります。
'              事前登録した`pHandlerThis`を基に、メソッド名を特定する必要があります。
'
'            ・CDPコマンド結果は、`CB_OnCDPMessage`で別途処理を設けています。ここのプロシージャはあくまでも、非同期イベントのみ担います。
'***************************************************************************************************
Friend Sub CB_OnDevToolsProtocolEvent(ByVal ParameterObjectAsJson As String, ByVal pHandlerThis As LongPtr)
    Const FromProcedureName As String = ThisClassName & ".CB_OnDevToolsProtocolEvent"


    '空文字の場合はそもそも、通さないようにする
    If Len(ParameterObjectAsJson) = 0 Then Exit Sub
    printMsg Trace_, "CDP-Event: " & ParameterObjectAsJson & " , pHandlerThis: " & pHandlerThis, FromProcedureName

    '1. ポインター値から、16進数表記に変換
    Dim Key As String: Key = PtrToHex(pHandlerThis)

    '2. その16進数から、`method`名の特定を試みる
    Dim methodName As String
    If m_DevToolsProtocolHandlerThisToEventName.Exists(Key) Then
        '`method`名を取り出す
        methodName = m_DevToolsProtocolHandlerThisToEventName(Key)
    Else
        '特定出来ない場合は、意味をなさないので抜ける
        Exit Sub
    End If

    '3. イベント発火させる
    printMsg Debug_, "CDP-EventMethod: " & methodName & " , Params: " & ParameterObjectAsJson, FromProcedureName
    RaiseEvent CDPEventReceived(methodName, ParameterObjectAsJson)
End Sub

'----------------------------------------------------------------------
' ProcessMessages  ― COM STA コールバック層に必要なメッセージの処理
'   IsReady になるまでの待機ループ内で繰り返し呼ぶ。
'   State 1（Environment取得済み）を検出したら Controller 作成を開始する。
'----------------------------------------------------------------------
Public Sub ProcessMessages(Optional ByVal iterations As Long = 100)
    Const FromProcedureName As String = ThisClassName & ".ProcessMessages"

    Dim msg(0 To 5) As LongPtr
    Dim i As Long
    Dim dispCount As Long
    dispCount = 0
    For i = 1 To iterations
        If PeekMessage(msg(0), 0, 0, 0, PM_REMOVE) Then
            TranslateMessage msg(0)
            DispatchMessage msg(0)
            dispCount = dispCount + 1
        Else
            Exit For
        End If
    Next i
    DoEvents

    ' 呼び出し頻度が高いため、実際に WM を配送したときだけ TRACE（空回しは同期待ち側の spin ログを参照）
    If dispCount > 0 Then
        printMsg Trace_, "ProcessMessages: requestedIter=" & CStr(iterations) & " dispatchedMsgs=" & CStr(dispCount) & " m_State=" & CStr(m_State), FromProcedureName
    End If

    ' State 1: Environment → Controller 作成
    If m_State = 1 Then BeginControllerCreation
    ' State 2: Controller → 表示・WebView2取得
    If m_State = 2 Then FinishWebViewSetup
End Sub



Public Sub SetVisible(ByVal isVisible As Boolean)
    If m_pController = 0 Then Exit Sub
    Dim visVal As Long: visVal = IIf(isVisible, 1, 0)
    COM_DispCallLong1 m_pController, OFF_CTRL_put_IsVisible, visVal
End Sub

'----------------------------------------------------------------------
' Quit : リソース解放
'----------------------------------------------------------------------
Public Sub quit()
    On Error Resume Next
    If m_pController <> 0 Then
        Dim retV As Variant
        DispCallFunc m_pController, OFF_CTRL_Close, CC_STDCALL, vbLong, 0, 0, 0, retV
    End If
    m_pWebView2 = 0
    m_pController = 0
    m_pEnvironment = 0
    Set m_CDPResultQueue = Nothing
    Set m_CDPAsyncPThisHexToInfo = Nothing
    Set m_CDPAsyncResultJsonByCmdId = Nothing
    Set m_CDPAsyncErrorCodeByCmdId = Nothing
    Set m_DevToolsProtocolEventSubscriptions = Nothing
    Set m_DevToolsProtocolEventReceiverHexToName = Nothing
    Set m_DevToolsProtocolHandlerThisToEventName = Nothing
    m_DTPEventHandlerThisCount = 0
    Set g_WebView2Core = Nothing
End Sub

'***************************************************************************************************
'* 機能　　：ログレベル付きのスマートなログをイミディエイトへ出力させるヘルパーメソッドです
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Dim logName As String, strFormattedMsg As String

    If isHeader Then strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-") _
    Else strFormattedMsg = " | " & brID & " | " & strMsg

   'Perform msg logging if needed
    If doLog Then
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".WebView2.log"
        If Dir(logPath, vbDirectory) = "" Then MkDir logPath
    End If

   'Perform msg output to Immediate Window if needed
    With LogControl
        Select Case LogLevel_
            Case LogLevelName.Trace_: .LogTrace strFormattedMsg, From, logPath, logName
            Case LogLevelName.Debug_: .LogDebug strFormattedMsg, From, logPath, logName
            Case LogLevelName.info_: .LogInfo strFormattedMsg, From, logPath, logName
            Case LogLevelName.WARN_: .LogWarn strFormattedMsg, From, logPath, logName
            Case LogLevelName.ERROR_: .LogError strFormattedMsg, From, logPath, logName, Err.Number
            Case Else: 'ログ出力無効化
        End Select
    End With


    'WebView2自身や、非同期処理自体に問題や、タイムアウトになったと判断された場合は、このツールでの処理(CDP-Json)ができないので、エラーで停止させておく
    If doRaiseError Then Err.Raise ERR_CoreWebView2, From, Description:=strMsg

End Sub

'----------------------------------------------------------------------
' LastError : 直前のエラー情報を返す
'----------------------------------------------------------------------
Public Property Get LastErrorCode() As Long:        LastErrorCode = m_LastErrorCode:   End Property
Public Property Get LastErrorDescription() As String: LastErrorDescription = m_LastErrorDesc: End Property
Public Property Get IsReady() As Boolean:           IsReady = (m_pWebView2 <> 0):     End Property

' =====================================================================
' Friend（Callbacks から呼ばれる）
' =====================================================================

'----------------------------------------------------------------------
' CB_EnvironmentCreated
'   WV2_EnvCB_Invoke → ここへ委譲。
'   ※ コールバック内で CreateCoreWebView2Controller を呼ぶと再入問題が起きるため、
'     pEnv を保存して m_State=1 をセットするだけにする。
'     実際の Controller 作成は ProcessMessages → BeginControllerCreation で行う。
'----------------------------------------------------------------------
Friend Sub CB_EnvironmentCreated(ByVal ErrorCode As Long, ByVal pEnv As LongPtr)
    If WV2logView Then Debug.Print "[WV2] CB_EnvironmentCreated: errorCode=" & ErrorCode & ", pEnv=" & Hex(pEnv)
    If ErrorCode <> 0 Or pEnv = 0 Then
        m_LastErrorCode = ErrorCode
        m_LastErrorDesc = "Environment creation failed: 0x" & Hex(ErrorCode)
        RaiseEvent InitializeFailed(ErrorCode, m_LastErrorDesc)
        Exit Sub
    End If
    COM_AddRef pEnv
    m_pEnvironment = pEnv
    m_State = 1     ' → BeginControllerCreation へのシグナル
End Sub

'----------------------------------------------------------------------
' BeginControllerCreation（ProcessMessages から呼ばれる）
' 注: DispCallFunc の prgpvarg は VARIANTARG*(16byte構造体) へのポインタ配列
'----------------------------------------------------------------------
Private Sub BeginControllerCreation()
    m_State = 9     ' 「作成中」マーク（二重呼び出し防止）

    ' --- 診断ログ ---
    Dim pCtrlHandler As LongPtr: pCtrlHandler = VarPtr(m_CtrlHandlerThis)
    If WV2logView Then Debug.Print "[WV2] BeginControllerCreation"
    If WV2logView Then Debug.Print "  hWnd          = 0x" & Hex(m_hWndParent)
    If WV2logView Then Debug.Print "  pEnvironment  = 0x" & Hex(m_pEnvironment)
    If WV2logView Then Debug.Print "  pCtrlHandler  = 0x" & Hex(pCtrlHandler)
    If WV2logView Then Debug.Print "  vtable ptr    = 0x" & Hex(m_CtrlHandlerThis.pVTable)

    ' --- VARIANT 構造体ベースで引数を渡す ---
    ' VARIANTARG: vt(2)+予約(6)+value(8) = 16 bytes
    ' DispCallFunc が *prgpvarg[i] の offset+8 から値を読む場合に対応
    Dim vHwnd(0 To 1) As LongPtr    ' [0]=vt+pad(8bytes), [1]=value(8bytes)
    Dim vHandler(0 To 1) As LongPtr

    ' vt フィールド (下位2バイト) に vbLongLong=20 をセット、残りは0
    Dim vtLongLong As Integer: vtLongLong = 20  ' VT_I8
    CopyMemory vHwnd(0), vtLongLong, 2
    vHwnd(1) = m_hWndParent

    CopyMemory vHandler(0), vtLongLong, 2
    vHandler(1) = pCtrlHandler

    Dim argT(1) As Integer
    Dim argP(1) As LongPtr
    argT(0) = vbLongLong: argP(0) = VarPtr(vHwnd(0))
    argT(1) = vbLongLong: argP(1) = VarPtr(vHandler(0))

    Dim retV As Variant
    Dim hr As Long
    hr = DispCallFunc(m_pEnvironment, OFF_ENV_CreateController, CC_STDCALL, vbLong, 2, argT(0), argP(0), retV)
    If WV2logView Then Debug.Print "  hr(DispCallFunc)=" & Hex(hr) & "  result(HRESULT)=" & Hex(CLng(retV))
End Sub

'----------------------------------------------------------------------
' CB_ControllerCreated
'   WV2_CtrlCB_Invoke → ここへ委譲。
'   ※ コールバック内での COM 呼び出しは VBE7.DLL クラッシュの原因になるため、
'     pCtrl を保存して m_State=2 をセットするだけにする。
'     実際の表示処理は ProcessMessages → FinishWebViewSetup で行う。
'----------------------------------------------------------------------
Friend Sub CB_ControllerCreated(ByVal ErrorCode As Long, ByVal pCtrl As LongPtr)
    If WV2logView Then Debug.Print "[WV2] CB_ControllerCreated: errorCode=" & ErrorCode & ", pCtrl=" & Hex(pCtrl)
    If ErrorCode <> 0 Or pCtrl = 0 Then
        m_LastErrorCode = ErrorCode
        m_LastErrorDesc = "Controller creation failed: 0x" & Hex(ErrorCode)
        RaiseEvent InitializeFailed(ErrorCode, m_LastErrorDesc)
        Exit Sub
    End If
    ' AddRef（0引数 DispCallFunc）はコールバック内でも安全（山引数ありの呼び出しがクラッシュするのでこちらだけ残す）
    COM_AddRef pCtrl
    m_pController = pCtrl
    m_State = 2
End Sub

'----------------------------------------------------------------------
' FinishWebViewSetup（ProcessMessages から呼ばれる）
'   BeginControllerCreation と同じ VARIANT構造体ベースの引数渡しで COM を呼ぶ。
'----------------------------------------------------------------------
Private Sub FinishWebViewSetup()
    m_State = 99
    If WV2logView Then Debug.Print "[WV2] FinishWebViewSetup: pController=0x" & Hex(m_pController)

    ' VARIANTARGバッファ: [0]=vt+予約(8bytes) [1]=value(8bytes)
    Dim vtII As Integer: vtII = 20  ' VT_I8
    Dim retV As Variant

    ' --- put_Bounds (RECT ポインタ渡し) ---
    Dim vBounds(0 To 1) As LongPtr
    CopyMemory vBounds(0), vtII, 2
    Dim pBoundsArg As LongPtr: pBoundsArg = VarPtr(m_Bounds)
    vBounds(1) = pBoundsArg
    Dim aT1(0) As Integer: aT1(0) = 20
    Dim aP1(0) As LongPtr: aP1(0) = VarPtr(vBounds(0))
    DispCallFunc m_pController, OFF_CTRL_put_Bounds, CC_STDCALL, vbLong, 1, aT1(0), aP1(0), retV
    If WV2logView Then Debug.Print "  put_Bounds HRESULT=0x" & Hex(CLng(retV))

    ' --- put_IsVisible (BOOL=1) ---
    Dim vVis(0 To 1) As LongPtr
    CopyMemory vVis(0), vtII, 2
    vVis(1) = 1
    Dim aT2(0) As Integer: aT2(0) = 20
    Dim aP2(0) As LongPtr: aP2(0) = VarPtr(vVis(0))
    DispCallFunc m_pController, OFF_CTRL_put_IsVisible, CC_STDCALL, vbLong, 1, aT2(0), aP2(0), retV
    If WV2logView Then Debug.Print "  put_IsVisible HRESULT=0x" & Hex(CLng(retV))

    ' --- get_CoreWebView2 (出力: ICoreWebView2*) ---
    Dim pWV2 As LongPtr
    Dim vWV2(0 To 1) As LongPtr
    CopyMemory vWV2(0), vtII, 2
    vWV2(1) = VarPtr(pWV2)
    Dim aT3(0) As Integer: aT3(0) = 20
    Dim aP3(0) As LongPtr: aP3(0) = VarPtr(vWV2(0))
    DispCallFunc m_pController, OFF_CTRL_get_CoreWebView2, CC_STDCALL, vbLong, 1, aT3(0), aP3(0), retV
    m_pWebView2 = pWV2
    If WV2logView Then Debug.Print "  get_CoreWebView2 HRESULT=0x" & Hex(CLng(retV)) & ", pWV2=0x" & Hex(m_pWebView2)

    ' Ready 通知
    RaiseEvent Ready
    If Len(m_InitialUrl) > 0 Then navigate m_InitialUrl
End Sub

' =====================================================================
' Private ─ vtable 構築
' =====================================================================
Private Sub BuildVtables()
    ' コールバック関数アドレスを Callbacks モジュールから取得
    Dim envFn()  As LongPtr
    Dim ctrlFn() As LongPtr
    WV2_FillFunctionPointers envFn, ctrlFn

    ' 環境ハンドラ vtable
    m_EnvHandlerVT.fn(0) = envFn(0)
    m_EnvHandlerVT.fn(1) = envFn(1)
    m_EnvHandlerVT.fn(2) = envFn(2)
    m_EnvHandlerVT.fn(3) = envFn(3)
    m_EnvHandlerThis.pVTable = VarPtr(m_EnvHandlerVT.fn(0))

    ' コントローラーハンドラ vtable
    m_CtrlHandlerVT.fn(0) = ctrlFn(0)
    m_CtrlHandlerVT.fn(1) = ctrlFn(1)
    m_CtrlHandlerVT.fn(2) = ctrlFn(2)
    m_CtrlHandlerVT.fn(3) = ctrlFn(3)
    m_CtrlHandlerThis.pVTable = VarPtr(m_CtrlHandlerVT.fn(0))
    Dim cdpFn() As LongPtr
    WV2_FillCDPFunctionPointers cdpFn
    m_CDPHandlerVT.fn(0) = cdpFn(0): m_CDPHandlerVT.fn(1) = cdpFn(1)
    m_CDPHandlerVT.fn(2) = cdpFn(2): m_CDPHandlerVT.fn(3) = cdpFn(3)
    m_CDPHandlerThis.pVTable = VarPtr(m_CDPHandlerVT.fn(0))

    ' Async 用も同じ関数テーブルを使うが、this を分離して判別できるようにする
    m_CDPAsyncHandlerVT.fn(0) = cdpFn(0): m_CDPAsyncHandlerVT.fn(1) = cdpFn(1)
    m_CDPAsyncHandlerVT.fn(2) = cdpFn(2): m_CDPAsyncHandlerVT.fn(3) = cdpFn(3)
    m_CDPAsyncHandlerThis.pVTable = VarPtr(m_CDPAsyncHandlerVT.fn(0))

    ' async 用の this をコマンドごとにユニーク化するための固定プール
    Dim i As Long
    For i = 0 To MAX_ASYNC_CDP_COMMANDS - 1
        m_CDPAsyncHandlerThisPool(i).pVTable = VarPtr(m_CDPAsyncHandlerVT.fn(0))
    Next i

    ' DevToolsProtocol event handler vtable
    Dim dtpFn() As LongPtr
    WV2_FillDevToolsProtocolEventFunctionPointers dtpFn
    m_DTPEventHandlerVT.fn(0) = dtpFn(0): m_DTPEventHandlerVT.fn(1) = dtpFn(1)
    m_DTPEventHandlerVT.fn(2) = dtpFn(2): m_DTPEventHandlerVT.fn(3) = dtpFn(3)
    m_DTPEventHandlerThis.pVTable = VarPtr(m_DTPEventHandlerVT.fn(0))
End Sub

' ICoreWebView2EnvironmentOptions 偽装 vtable を組み立て（追加ブラウザ引数用、Options2 含む）
Private Sub BuildEnvOptionsVtable()
    Dim fn() As LongPtr
    WV2_FillEnvOptionsFunctionPointers fn
    Dim i As Long
    For i = 0 To 12
        m_EnvOptionsVT.fn(i) = fn(i)
    Next
    m_EnvOptionsThis.pVTable = VarPtr(m_EnvOptionsVT.fn(0))
End Sub

' =====================================================================
' Private ─ COM ヘルパー
' =====================================================================

' AddRef（参照カウンタ増加）
Private Sub COM_AddRef(ByVal pObj As LongPtr)
    Dim retV As Variant
    DispCallFunc pObj, 1 * PTR_SIZE, CC_STDCALL, vbLong, 0, 0, 0, retV
End Sub



'***************************************************************************************************
'                           ■■■ 汎用的なDispCallFunc設計 ■■■
'***************************************************************************************************
' put_Bounds(RECT)  ─ x64 ABI では 16バイト構造体はポインタ渡し
'   重要: x64 の DispCallFunc にはポインタを VARIANT(16bytes) 形式で渡す必要がある
Private Sub COM_SetBounds(ByVal pCtrl As LongPtr, ByRef bounds As RECT_T)
    Dim pBounds As LongPtr: pBounds = VarPtr(bounds)
    Dim vBounds(0 To 1) As LongPtr
    Dim vtII As Integer: vtII = 20  ' VT_I8
    CopyMemory vBounds(0), vtII, 2
    vBounds(1) = pBounds            ' RECT構造体へのポインタを詰める

    Dim argT(0) As Integer: argT(0) = 20 ' vbLongLong
    Dim argP(0) As LongPtr: argP(0) = VarPtr(vBounds(0))
    Dim retV As Variant
    DispCallFunc pCtrl, OFF_CTRL_put_Bounds, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
End Sub

' 引数が LongPtr 1個の汎用DispCallFunc
Private Sub COM_DispCallPtr1(ByVal pObj As LongPtr, ByVal Offset As Long, ByVal arg0 As LongPtr)
    Dim vArg(0 To 1) As LongPtr
    Dim vtII As Integer: vtII = 20  ' VT_I8
    CopyMemory vArg(0), vtII, 2
    vArg(1) = arg0
    
    Dim argT(0) As Integer: argT(0) = 20
    Dim argP(0) As LongPtr: argP(0) = VarPtr(vArg(0))
    Dim retV As Variant
    DispCallFunc pObj, Offset, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
End Sub

' 引数が Long 1個の汎用DispCallFunc（put_IsVisible など）
Private Sub COM_DispCallLong1(ByVal pObj As LongPtr, ByVal Offset As Long, ByVal arg0 As Long)
    Dim vArg(0 To 1) As LongPtr
    Dim vtII As Integer: vtII = 3   ' VT_I4 / vbLong
    CopyMemory vArg(0), vtII, 2
    CopyMemory vArg(1), arg0, 4     ' value (lower 4 bytes of LongPtr slot)
    
    Dim argT(0) As Integer: argT(0) = 3 ' vbLong
    Dim argP(0) As LongPtr: argP(0) = VarPtr(vArg(0))
    Dim retV As Variant
    DispCallFunc pObj, Offset, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
End Sub



'***************************************************************************************************
'                     ■■■ WebView2に用意されてるメソッドのCOM設計 ■■■
'***************************************************************************************************
'* 処理名　：CallDevToolsProtocolMethod
'---------------------------------------------------------------------------------------------------
'* 詳細説明：CallDevToolsProtocolMethod(methodName, parametersAsJson, handler) ? 3 引数ともポインタ
'* URL     ：https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2#calldevtoolsprotocolmethod
'***************************************************************************************************
Private Function COM_CallDevToolsProtocolMethod(ByVal pWV2 As LongPtr, ByVal methodName As String, _
    ByVal paramsJson As String, ByVal pHandler As LongPtr) As Long

    Dim vtII As Integer: vtII = 20
    Dim vMethod(0 To 1) As LongPtr, vParams(0 To 1) As LongPtr, vHandler(0 To 1) As LongPtr

    CopyMemory vMethod(0), vtII, 2: vMethod(1) = StrPtr(methodName)
    CopyMemory vParams(0), vtII, 2: vParams(1) = StrPtr(paramsJson)
    CopyMemory vHandler(0), vtII, 2: vHandler(1) = pHandler

    Dim argT(0 To 2) As Integer: argT(0) = 20: argT(1) = 20: argT(2) = 20
    Dim argP(0 To 2) As LongPtr: argP(0) = VarPtr(vMethod(0)): argP(1) = VarPtr(vParams(0)): argP(2) = VarPtr(vHandler(0))
    Dim retV As Variant

    DispCallFunc pWV2, m_CDPMethodVTableIndex * PTR_SIZE, CC_STDCALL, vbLong, 3, argT(0), argP(0), retV

    COM_CallDevToolsProtocolMethod = CLng(retV)
End Function

'***************************************************************************************************
'* 処理名　：CallDevToolsProtocolMethodForSession
'---------------------------------------------------------------------------------------------------
'* 詳細説明：CallDevToolsProtocolMethodForSession(sessionId, methodName, parametersAsJson, handler)
'* URL     ：https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2_11#calldevtoolsprotocolmethodforsession
'***************************************************************************************************
Private Function COM_CallDevToolsProtocolMethodForSession(ByVal pWV2 As LongPtr, ByVal sessionID As String, _
    ByVal methodName As String, ByVal paramsJson As String, ByVal pHandler As LongPtr) As Long

    Dim vtII As Integer: vtII = 20
    Dim vSession(0 To 1) As LongPtr, vMethod(0 To 1) As LongPtr, vParams(0 To 1) As LongPtr, vHandler(0 To 1) As LongPtr

    CopyMemory vSession(0), vtII, 2: vSession(1) = StrPtr(sessionID)
    CopyMemory vMethod(0), vtII, 2: vMethod(1) = StrPtr(methodName)
    CopyMemory vParams(0), vtII, 2: vParams(1) = StrPtr(paramsJson)
    CopyMemory vHandler(0), vtII, 2: vHandler(1) = pHandler

    Dim argT(0 To 3) As Integer: argT(0) = 20: argT(1) = 20: argT(2) = 20: argT(3) = 20
    Dim argP(0 To 3) As LongPtr: argP(0) = VarPtr(vSession(0)): argP(1) = VarPtr(vMethod(0)): argP(2) = VarPtr(vParams(0)): argP(3) = VarPtr(vHandler(0))
    Dim retV As Variant

    DispCallFunc pWV2, m_CDPMethodForSessionVTableIndex * PTR_SIZE, CC_STDCALL, vbLong, 4, argT(0), argP(0), retV

    COM_CallDevToolsProtocolMethodForSession = CLng(retV)
End Function

'***************************************************************************************************
'* 処理名　：Navigate
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2#navigate
'***************************************************************************************************
Private Function COM_Navigate(ByVal pWV2 As LongPtr, ByVal Url As String) As Long
    Dim vtII As Integer: vtII = 20  ' VT_I8
    Dim vUrl(0 To 1) As LongPtr

    CopyMemory vUrl(0), vtII, 2: vUrl(1) = StrPtr(Url)

    Dim argT(0) As Integer: argT(0) = 20
    Dim argP(0) As LongPtr: argP(0) = VarPtr(vUrl(0))
    Dim retV As Variant

    DispCallFunc pWV2, OFF_WV2_Navigate, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV

    COM_Navigate = CLng(retV)
End Function

'***************************************************************************************************
'* 処理名　：ICoreWebView2DevToolsProtocolEventReceiver
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2devtoolsprotocoleventreceiver
'***************************************************************************************************
Private Function COM_GetDevToolsProtocolEventReceiver(ByVal pWV2 As LongPtr, ByVal eventName As String) As LongPtr
    Dim pReceiver As LongPtr

    Dim vEvent(0 To 1) As LongPtr
    Dim vtII As Integer: vtII = 20 'VT_I8
    CopyMemory vEvent(0), vtII, 2
    vEvent(1) = StrPtr(eventName)

    Dim vOut(0 To 1) As LongPtr
    CopyMemory vOut(0), vtII, 2
    vOut(1) = VarPtr(pReceiver)

    Dim argT(0 To 1) As Integer
    Dim argP(0 To 1) As LongPtr
    argT(0) = vbLongLong: argP(0) = VarPtr(vEvent(0))
    argT(1) = vbLongLong: argP(1) = VarPtr(vOut(0))

    Dim retV As Variant
    Dim hr As Long
    hr = DispCallFunc(pWV2, m_DevToolsProtocolEventReceiverVTableIndex * PTR_SIZE, CC_STDCALL, vbLong, 2, argT(0), argP(0), retV)
    If hr <> 0 Then
        COM_GetDevToolsProtocolEventReceiver = 0
        Exit Function
    End If

    COM_GetDevToolsProtocolEventReceiver = pReceiver
End Function



' =====================================================================
' Class lifecycle
' =====================================================================
Private Sub Class_Initialize()
    m_pEnvironment = 0
    m_pController = 0
    m_pWebView2 = 0
    m_hWndFrame = 0
    m_LastAppliedRight = 0
    m_LastAppliedBottom = 0
    m_LastErrorCode = 0
    m_State = 0
    Set m_CDPResultQueue = New Collection
    m_LastCDPResultErrorCode = 0

    Set m_CDPAsyncPThisHexToInfo = New Dictionary
    Set m_CDPAsyncResultJsonByCmdId = New Dictionary
    Set m_CDPAsyncErrorCodeByCmdId = New Dictionary

    Set m_DevToolsProtocolEventSubscriptions = Nothing
    Set m_DevToolsProtocolEventReceiverHexToName = Nothing
    Set m_DevToolsProtocolHandlerThisToEventName = Nothing
    m_DTPEventHandlerThisCount = 0

    'Assign an objID for easy debug identification
    brID = "BRID" & Format(Rnd * 1000, "000")


    '------------------------- 設定シートの各セルから設定値を取得し、クラス内変数に適用します(予定 ToDo) -------------------------
    With ShSetting01_StartBrowser
        m_CDPMethodVTableIndex = 36     ' 既定。環境でずれる場合は CDPMethodVTableIndex で 22,37,39 等を指定
        m_CDPMethodForSessionVTableIndex = 99
        ResultCDPTimeout = 30           'CDPからの結果を得る際のタイムアウト秒数
    
        ' GetDevToolsProtocolEventReceiver の vtable index（推定値。環境差が出る場合は変更用プロパティを使って調整）
        m_DevToolsProtocolEventReceiverVTableIndex = 42

        '2. パスが設定されてたら、ログファイル保存を有効化
        Dim LogFolderPath As String: LogFolderPath = .Range(.UseRangeName(9, "WebView2Browser.Class_Initialize")).value
        If LogFolderPath <> "" Then
            doLog = True                'ログファイルとして保存するか？
            logPath = LogFolderPath     '保存先のフォルダパスを設定
        End If

    End With

End Sub

Private Sub Class_Terminate()
    quit
End Sub
