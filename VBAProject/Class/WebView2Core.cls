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

' =====================================================================
' Win32 API Declares (64bit PtrSafe)
' =====================================================================
#If VBA7 Then

' WebView2Loader.dll ─ Edge に同梱されている DLL を使用
' Edge インストールフォルダ（通常は自動的に PATH に含まれないため、
' Initialize 内で LoadLibrary を用いてフルパスからロードする）
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
Private Const OFF_WV2_add_NavCompleted     As Long = 15 * 8 ' = 120
Private Const OFF_WV2_ExecuteScript        As Long = 38 * 8 ' ※要検証
Private Const OFF_WV2_get_BrowserProcessId As Long = 37 * 8 ' = 296  [propget] UINT32* value
' CallDevToolsProtocolMethod の vtable index。ランタイム/SDK でずれるためプロパティで上書き可（既定 36）
Private m_CDPMethodVTableIndex As Long

'ベースとなる作業フォルダパス
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

' =====================================================================
' CDP メッセージ キュー / イベントキャプチャ
' =====================================================================
' WebView2 の CDP callback は command result / event が混ざって届く可能性があるため、
' JSON を受信した時点で「method を持つもの=イベント」「それ以外=コマンド結果」に振り分けてキュー化します。
Private m_CDPEventQueue As Collection
Private m_CDPResultQueue As Collection
Private m_LastCDPResultErrorCode As Long

' DevToolsProtocol event receiver（現状は「1種類の eventName のみ」想定）
Private m_DevToolsProtocolEventName As String
Private m_DevToolsProtocolEventReceiverPtr As LongPtr
Private m_DevToolsProtocolEventToken As LongPtr

' DevToolsProtocol event receiver（複数購読対応）
Private m_DevToolsProtocolEventSubscriptions As Dictionary ' eventName -> array(receiverPtr, token, handlerThisPtr)
Private m_DevToolsProtocolEventReceiverHexToName As Dictionary ' receiverPtr(hex) -> eventName（pSender 用・フォールバック）
Private m_DevToolsProtocolHandlerThisToEventName As Dictionary ' handlerThis(hex) -> eventName（コールバックの pThis で振り分け）

' ICoreWebView2DevToolsProtocolEventReceivedEventHandler の vtable 偽装（同一 vtable を全ハンドラで共有）
Private m_DTPEventHandlerThis As VtblObj
Private m_DTPEventHandlerVT As VtblData4
Private Const MAX_DTP_EVENT_HANDLERS As Long = 256
Private m_DTPEventHandlerThisArray(0 To MAX_DTP_EVENT_HANDLERS - 1) As VtblObj ' 固定領域（再配置しない）
Private m_DTPEventHandlerThisCount As Long

' BrowserEvents: CDPBrowser と同様の Dictionary 構造でイベントを蓄積します
Private m_BrowserEvents As Dictionary
Private Const EventsDictionaryKeyName01 As String = "TotalEvents"
Private Const EventsDictionaryKeyName02 As String = "EventMethods"
Private Const EventsDictionaryKeyIndex As String = "__index__"

' JSON -> Dictionary 変換（イベント解析用）
Private jsConverter As New WebJsonConverter

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


' =====================================================================
' Public Properties（CDP vtable の動的対応用）
' =====================================================================
' CDPMethodVTableIndex: CallDevToolsProtocolMethod の vtable index（既定 36）。
' ランタイム版でずれた場合は 22,37,39 等を設定。バージョン検出＋テーブルで自動設定する拡張も可。
Public Property Get CDPMethodVTableIndex() As Long: CDPMethodVTableIndex = m_CDPMethodVTableIndex: End Property
Public Property Let CDPMethodVTableIndex(ByVal idx As Long): m_CDPMethodVTableIndex = idx: End Property

' GetDevToolsProtocolEventReceiver の vtable index（ランタイム/SDK 差異調整用）
Public Property Get DevToolsProtocolEventReceiverVTableIndex() As Long: DevToolsProtocolEventReceiverVTableIndex = m_DevToolsProtocolEventReceiverVTableIndex: End Property
Public Property Let DevToolsProtocolEventReceiverVTableIndex(ByVal idx As Long): m_DevToolsProtocolEventReceiverVTableIndex = idx: End Property

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

Private Function PtrToHex(ByVal p As LongPtr) As String
    'LongPtr を Dictionary のキーにするための簡易変換
    PtrToHex = "0x" & Hex$(p)
End Function

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


' =====================================================================
' Public Interface
' =====================================================================

'----------------------------------------------------------------------
' Initialize
'   hWndParent : WebView2 を埋め込む親ウィンドウハンドル
'   boundsLeft/Top/Width/Height : 表示領域（ピクセル）
'   initialUrl : 最初に遷移するURL（省略可）
'   additionalBrowserArguments : 追加ブラウザ引数（省略可）。例: "--disable-web-security"
'   戻り値 : True = 初期化コマンド送信成功（Ready イベント待ち）
'----------------------------------------------------------------------
Public Function Initialize( _
    ByVal hwndParent As LongPtr, _
    ByVal boundsLeft As Long, _
    ByVal boundsTop As Long, _
    ByVal boundsWidth As Long, _
    ByVal boundsHeight As Long, _
    Optional ByVal initialUrl As String = "about:blank", Optional UserDataName As String, Optional additionalBrowserArguments As String) As Boolean

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

    ' WebView2Loader.dll が見つからない環境用に Edge フォルダから明示ロード
    'EnsureWebView2LoaderDll

    ' 環境作成開始（非同期; WV2_EnvCB_Invoke が呼ばれると CB_EnvironmentCreated が動く）
    Dim hr As Long
    Dim userDataPath As String
    
    'Excelが使うWebView2の作業ディレクトリに合わせる
    userDataPath = Environ("LOCALAPPDATA") & "\Microsoft\Office\" & Application.Version & "\" & WorkSpaceFolderName & "\" & UserDataName

    Dim pOptions As LongPtr
    Dim useOptions As Boolean
    If Len(additionalBrowserArguments) > 0 Then
        g_EnvOptions_AdditionalArgs = additionalBrowserArguments
        BuildEnvOptionsVtable
        pOptions = VarPtr(m_EnvOptionsThis)
        useOptions = True
    Else
        pOptions = 0
        useOptions = False
        ' この起動では追加引数を使わないので、前回のフォールバック設定をクリア
        SetEnvironmentVariableW StrPtr("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"), 0
    End If

    Dim pUserData   As LongPtr: pUserData = StrPtr(userDataPath)
    Dim pHandler    As LongPtr: pHandler = VarPtr(m_EnvHandlerThis)

    hr = CreateCoreWebView2EnvironmentWithOptions(0, pUserData, pOptions, pHandler)

    ' E_INVALIDARG: 偽装 COM がマーシャリング等で拒否される。環境変数でフォールバック。
    ' pOptions=0 のとき、ローダーは WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS を参照する。
    If hr <> 0 And useOptions And hr = &H80070057 Then
        SetEnvironmentVariableW StrPtr("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"), StrPtr(additionalBrowserArguments)
        Debug.Print "追加引数: 環境変数でフォールバック"
        hr = CreateCoreWebView2EnvironmentWithOptions(0, pUserData, 0, pHandler)
        ' この時点で環境作成は開始されたので、他の Excel 内 WebView2 への影響を避けるため即座にクリア
        SetEnvironmentVariableW StrPtr("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"), 0
    End If

    If hr <> 0 Then
        m_LastErrorCode = hr
        If Len(m_LastErrorDesc) = 0 Then m_LastErrorDesc = "CreateCoreWebView2EnvironmentWithOptions failed: 0x" & Hex(hr)
        RaiseEvent InitializeFailed(hr, m_LastErrorDesc)
        Initialize = False
    Else
        Initialize = True
    End If
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

'----------------------------------------------------------------------
' CallDevToolsProtocolMethod : CDP メソッド実行。結果 JSON を同期待ちで返す
'   methodName: "Domain.method"（例: "Runtime.evaluate", "Page.navigate"）
'   paramsJson: パラメータ JSON（例: "{}" または "{""expression"":""1+1""}"）
'   診断: Immediate に [CDP] Invoke が出なければコールバック未呼び出し（vtable 要確認）
'----------------------------------------------------------------------
Public Function CallDevToolsProtocolMethod(ByVal methodName As String, ByVal paramsJson As String) As String
    If m_pWebView2 = 0 Then CallDevToolsProtocolMethod = "": Exit Function
    '同期呼び出しは「次に届くコマンド結果」だけを返す
    '（イベントが混ざって届いても誤判定しないよう、キューで分離します）
    If m_CDPResultQueue Is Nothing Then
        Set m_CDPResultQueue = New Collection
    Else
        Set m_CDPResultQueue = New Collection
    End If
    m_LastCDPResultErrorCode = 0
    Dim hr As Long
    hr = COM_CallDevToolsProtocolMethod(m_pWebView2, methodName, paramsJson, VarPtr(m_CDPHandlerThis))
    If hr <> 0 Then
        m_LastErrorCode = hr
        CallDevToolsProtocolMethod = "": Exit Function
    End If
    Dim t As Single: t = Timer
    Do
        ProcessMessages
        DoEvents

        If Not m_CDPResultQueue Is Nothing Then
            If m_CDPResultQueue.Count > 0 Then
                CallDevToolsProtocolMethod = CStr(m_CDPResultQueue(1))
                m_CDPResultQueue.Remove 1
                If m_LastCDPResultErrorCode <> 0 Then m_LastErrorCode = m_LastCDPResultErrorCode
                Exit Function
            End If
        End If
    Loop While (Timer - t) < 30

    Debug.Print "[CDP] timeout: command result not received from queue (try other vtable index 22,36,37,39)"
    CallDevToolsProtocolMethod = ""
End Function

'----------------------------------------------------------------------
' CallDevToolsProtocolMethodAsync : CDP メソッド実行（結果待ちしない）
'   methodName: "Domain.method"（例: "Runtime.callFunctionOn"）
'   paramsJson: パラメータ JSON（例: "{}"）
'   返り値: COM 呼び出し HRESULT（0=成功）
'----------------------------------------------------------------------
Public Function CallDevToolsProtocolMethodAsync(ByVal methodName As String, ByVal paramsJson As String) As Long
    If m_pWebView2 = 0 Then
        CallDevToolsProtocolMethodAsync = &H80004005 'E_FAIL
        Exit Function
    End If

    Dim hr As Long
    hr = COM_CallDevToolsProtocolMethod(m_pWebView2, methodName, paramsJson, VarPtr(m_CDPAsyncHandlerThis))
    If hr <> 0 Then m_LastErrorCode = hr
    CallDevToolsProtocolMethodAsync = hr
End Function

'----------------------------------------------------------------------
' SubscribeDevToolsProtocolEvent
'   WebView2 の DevToolsProtocol event を購読します。
'   現状は「1種類の eventName のみ」フォールバックする実装です。
'----------------------------------------------------------------------
Public Function SubscribeDevToolsProtocolEvent(ByVal eventName As String) As Boolean
    If m_pWebView2 = 0 Then SubscribeDevToolsProtocolEvent = False: Exit Function
    If Len(eventName) = 0 Then SubscribeDevToolsProtocolEvent = False: Exit Function

    '複数購読対応：既に登録済みなら何もしない
    If m_DevToolsProtocolEventSubscriptions Is Nothing Then
        Set m_DevToolsProtocolEventSubscriptions = New Dictionary
    End If
    If m_DevToolsProtocolEventSubscriptions.Exists(eventName) Then
        SubscribeDevToolsProtocolEvent = True
        Exit Function
    End If

    'フォールバック用（TakeEvents 側が params のみの場合に備える）
    m_DevToolsProtocolEventName = eventName

    ' イベントごとに別の handler this を用意（コールバックの pThis で振り分けするため）
    Dim idx As Long: idx = m_DTPEventHandlerThisCount
    If idx >= MAX_DTP_EVENT_HANDLERS Then
        SubscribeDevToolsProtocolEvent = False
        Exit Function
    End If
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
    m_DevToolsProtocolEventReceiverPtr = pReceiver '互換用（最後に登録された）

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

    m_DevToolsProtocolEventToken = token

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

'----------------------------------------------------------------------
' UnsubscribeDevToolsProtocolEvent
'   SubscribeDevToolsProtocolEvent で登録したハンドラを解除します。
'   eventName を省略/空の場合：全購読を解除します。
'----------------------------------------------------------------------
Public Function UnsubscribeDevToolsProtocolEvent(Optional ByVal eventName As String = vbNullString) As Boolean
    On Error GoTo iferr
    If m_DevToolsProtocolEventSubscriptions Is Nothing Then UnsubscribeDevToolsProtocolEvent = True: Exit Function
    If m_DevToolsProtocolEventSubscriptions.Count = 0 Then UnsubscribeDevToolsProtocolEvent = True: Exit Function

    ' ICoreWebView2DevToolsProtocolEventReceiver::remove_DevToolsProtocolEventReceived
    ' vtable: IUnknown(0..2), add(3), remove(4)
    Const OFF_RCV_remove_DevToolsProtocolEventReceived As Long = 4 * 8 'index 4

    '引数なし => 全解除
    If Len(eventName) = 0 Then
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
        Next k

        Set m_DevToolsProtocolEventSubscriptions = New Dictionary
        If Not m_DevToolsProtocolEventReceiverHexToName Is Nothing Then Set m_DevToolsProtocolEventReceiverHexToName = New Dictionary
        If Not m_DevToolsProtocolHandlerThisToEventName Is Nothing Then Set m_DevToolsProtocolHandlerThisToEventName = New Dictionary
        m_DTPEventHandlerThisCount = 0
        m_DevToolsProtocolEventToken = 0
        m_DevToolsProtocolEventReceiverPtr = 0
        m_DevToolsProtocolEventName = vbNullString
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
' ProcessMessages  ― COM STA コールバック層に必要なメッセージの処理
'   IsReady になるまでの待機ループ内で繰り返し呼ぶ。
'   State 1（Environment取得済み）を検出したら Controller 作成を開始する。
'----------------------------------------------------------------------
Public Sub ProcessMessages(Optional ByVal iterations As Long = 100)
    Dim msg(0 To 5) As LongPtr
    Dim i As Long
    For i = 1 To iterations
        If PeekMessage(msg(0), 0, 0, 0, PM_REMOVE) Then
            TranslateMessage msg(0)
            DispatchMessage msg(0)
        Else
            Exit For
        End If
    Next i
    DoEvents

    ' State 1: Environment → Controller 作成
    If m_State = 1 Then BeginControllerCreation
    ' State 2: Controller → 表示・WebView2取得
    If m_State = 2 Then FinishWebViewSetup
End Sub

'----------------------------------------------------------------------
' BrowserEvents（CDPイベント蓄積）
'   Set BrowserEvents = Nothing                -> イベント破棄
'   Set BrowserEvents = New Dictionary (Count=0) -> 初期構造を作ってキャプチャ開始
'   Set BrowserEvents = 既存辞書（TotalEvents/ EventMethods を持つ） -> 追加キャプチャ
'----------------------------------------------------------------------
Public Property Get BrowserEvents() As Dictionary
    Set BrowserEvents = m_BrowserEvents
End Property

Public Property Set BrowserEvents(ByVal ObjDic As Dictionary)
    If ObjDic Is Nothing Then
        Set m_BrowserEvents = Nothing
        Exit Property
    End If

    'まっさら dictionary -> 初期構造を作って開始
    If ObjDic.Count = 0 Then
        Set m_BrowserEvents = FormatDictionary_BrowserEvents()
        Exit Property
    End If

    '既存辞書の形式チェック（CDPBrowser準拠）
    If ObjDic.Exists(EventsDictionaryKeyName01) And ObjDic.Exists(EventsDictionaryKeyName02) And ObjDic.Count = 2 Then
        Set m_BrowserEvents = ObjDic
    Else
        Err.Raise vbObjectError + 1010, "WebView2Core.BrowserEvents", "Invalid BrowserEvents Dictionary structure."
    End If
End Property

'----------------------------------------------------------------------
' TakeEvents
'   CDP callback から積まれた「イベントキュー」を Dictionary に反映します。
'----------------------------------------------------------------------
Public Sub TakeEvents(Optional InternalUseMessages As Dictionary)
    If m_CDPEventQueue Is Nothing Then Exit Sub
    If m_CDPEventQueue.Count = 0 Then Exit Sub

    'キャプチャ無効なら破棄（キューだけは膨らませない）
    If m_BrowserEvents Is Nothing And InternalUseMessages Is Nothing Then
        Set m_CDPEventQueue = New Collection
        Exit Sub
    End If

    If Not m_BrowserEvents Is Nothing Then
        Dim eventsDict As Dictionary
        Set eventsDict = m_BrowserEvents(EventsDictionaryKeyName02)
    End If

    If Not InternalUseMessages Is Nothing Then
        Set InternalUseMessages = FormatDictionary_BrowserEvents()

        Dim eventsDictForInternal As Dictionary
        Set eventsDictForInternal = InternalUseMessages(EventsDictionaryKeyName02)
    End If

    Dim i As Long
    For i = 1 To m_CDPEventQueue.Count
        Dim rawJson As String: rawJson = CStr(m_CDPEventQueue(i))
        If Len(rawJson) = 0 Then GoTo nextI

        'イベント本体を Dictionary 化（WebView2 は params のみで届く想定のため、method は文中の "method" に依存しない）
        Dim eventData As Dictionary
        On Error GoTo parseErr
        Set eventData = jsConverter.ParseJson(rawJson)
        If eventData Is Nothing Then GoTo nextI

        '振り分け用 methodName：トップレベルの "method" キーのみ採用（params 内の "method" は HTTP メソッド等で誤検出されるため使わない）
        Dim methodName As String: methodName = vbNullString
        If eventData.Exists("method") Then
            Dim mVal As Variant: mVal = eventData("method")
            If VarType(mVal) = vbString Then methodName = Trim$(CStr(mVal))
        End If
        If Len(methodName) = 0 Then
            'params のみで来た場合（レガシー／pSender 未マッチ時）。単一購読時のみフォールバックで振り分け可能
            methodName = m_DevToolsProtocolEventName
            If Len(methodName) = 0 Then GoTo nextI
            'params のみの JSON なので eventData を { method, params: eventData } に再構成
            Dim paramsOnly As Dictionary
            Set paramsOnly = eventData
            Set eventData = New Dictionary
            eventData.Add "method", methodName
            eventData.Add "params", paramsOnly
        End If

        'WebView2 によっては params が文字列で渡る場合がある（二重エンコード）。オブジェクトに正規化する。※エラーになるので一旦コメントアウト
'        If eventData.Exists("params") Then
'            Dim p As Variant: p = eventData("params")
'            If VarType(p) = vbString Then
'                If Len(CStr(p)) > 0 Then
'                    On Error Resume Next
'                    Dim paramsObj As Dictionary
'                    Set paramsObj = jsConverter.ParseJson(CStr(p))
'                    If Not paramsObj Is Nothing Then eventData("params") = paramsObj
'                    On Error GoTo parseErr
'                End If
'            End If
'        End If

        Dim total As Long
        Dim eventList As Collection
        
        If Not m_BrowserEvents Is Nothing Then
            total = CLng(m_BrowserEvents(EventsDictionaryKeyName01)) + 1
            m_BrowserEvents(EventsDictionaryKeyName01) = total

            If eventData.Exists(EventsDictionaryKeyIndex) Then
                eventData(EventsDictionaryKeyIndex) = total
            Else
                eventData.Add EventsDictionaryKeyIndex, total
            End If

            If eventsDict.Exists(methodName) Then
                Set eventList = eventsDict(methodName)
            Else
                Set eventList = New Collection
                eventsDict.Add methodName, eventList
            End If
            eventList.Add eventData
        End If

        If Not InternalUseMessages Is Nothing Then
            total = CLng(InternalUseMessages(EventsDictionaryKeyName01)) + 1
            InternalUseMessages(EventsDictionaryKeyName01) = total

            If eventsDictForInternal.Exists(EventsDictionaryKeyIndex) Then
                eventData(EventsDictionaryKeyIndex) = total
            Else
                eventsDictForInternal.Add EventsDictionaryKeyIndex, total
            End If

            If eventsDictForInternal.Exists(methodName) Then
                Set eventList = eventsDictForInternal(methodName)
            Else
                Set eventList = New Collection
                eventsDictForInternal.Add methodName, eventList
            End If
            eventList.Add eventData
        End If

parseContinue:
        On Error GoTo 0
nextI:
    Next i

    '処理済みは破棄（destruction は将来拡張用。現状は常にクリア）
    Set m_CDPEventQueue = New Collection
    Exit Sub

parseErr:
    'JSONがパースできないものはスキップ（イベント形式が不安定な場合の保険）
    On Error GoTo 0
    GoTo nextI
End Sub

Public Property Get PendingEventCount() As Long
    If m_CDPEventQueue Is Nothing Then
        PendingEventCount = 0
    Else
        PendingEventCount = m_CDPEventQueue.Count
    End If
End Property

'----------------------------------------------------------------------
' Friend（WebView2Callbacks から呼ばれる）
'----------------------------------------------------------------------
Friend Sub CB_OnCDPMessage(ByVal ErrorCode As Long, ByVal MsgJson As String, ByVal pThis As LongPtr)
    If Len(MsgJson) = 0 Then Exit Sub
    If m_CDPEventQueue Is Nothing Then Set m_CDPEventQueue = New Collection
    If m_CDPResultQueue Is Nothing Then Set m_CDPResultQueue = New Collection

    ' 非同期呼び出しの完了結果は同期キューに混ぜない（結果取り違え防止）
    If pThis = VarPtr(m_CDPAsyncHandlerThis) Then Exit Sub

    ' "method" を含む -> CDP event として扱う
    If InStr(1, MsgJson, """method""", vbTextCompare) > 0 Then
        m_CDPEventQueue.Add MsgJson
    Else
        '同期結果（もしくは command result）
        m_LastCDPResultErrorCode = ErrorCode
        m_CDPResultQueue.Add MsgJson
    End If
End Sub

'----------------------------------------------------------------------
' CB_OnDevToolsProtocolEvent
'   DevToolsProtocol event receiver から「params(JSONのみ)」が渡ってくる想定。
'   pHandlerThis（登録時に渡した handler の this）で eventName を特定し、method 付き JSON に整形してキューへ積みます。
'----------------------------------------------------------------------
Friend Sub CB_OnDevToolsProtocolEvent(ByVal ParameterObjectAsJson As String, ByVal pHandlerThis As LongPtr)
    If Len(ParameterObjectAsJson) = 0 Then Exit Sub
    If m_CDPEventQueue Is Nothing Then Set m_CDPEventQueue = New Collection

    Dim eventName As String: eventName = vbNullString
    ' イベントごとの handler this で振り分け（複数購読で確実に識別）
    If Not m_DevToolsProtocolHandlerThisToEventName Is Nothing Then
        Dim Key As String: Key = PtrToHex(pHandlerThis)
        If m_DevToolsProtocolHandlerThisToEventName.Exists(Key) Then eventName = m_DevToolsProtocolHandlerThisToEventName(Key)
    End If
    If Len(eventName) = 0 And Not m_DevToolsProtocolEventReceiverHexToName Is Nothing Then
        ' フォールバック：pHandlerThis を receiver として扱う（旧挙動）
        Key = PtrToHex(pHandlerThis)
        If m_DevToolsProtocolEventReceiverHexToName.Exists(Key) Then eventName = m_DevToolsProtocolEventReceiverHexToName(Key)
        If Len(eventName) = 0 Then
            Dim pUnk As LongPtr: pUnk = COM_ToIUnknownPtr(pHandlerThis)
            If pUnk <> 0 Then
                Key = PtrToHex(pUnk)
                If m_DevToolsProtocolEventReceiverHexToName.Exists(Key) Then eventName = m_DevToolsProtocolEventReceiverHexToName(Key)
                COM_Release pUnk
            End If
        End If
    End If

    'method を付与して CDP(event 完全文) と同じ形に揃える
    'ParameterObjectAsJson の前後空白はトリム（COM が余分に付与する場合の対策）
    Dim paramsJson As String: paramsJson = Trim$(ParameterObjectAsJson)
    If Len(eventName) > 0 Then
        m_CDPEventQueue.Add "{""method"":""" & eventName & """,""params"":" & paramsJson & "}"
    Else
        'フォールバック（分類不能なら params のみを積む）
        m_CDPEventQueue.Add paramsJson
    End If
End Sub

Private Function FormatDictionary_BrowserEvents() As Dictionary
    Set FormatDictionary_BrowserEvents = New Dictionary
    FormatDictionary_BrowserEvents.Add EventsDictionaryKeyName01, 0
    FormatDictionary_BrowserEvents.Add EventsDictionaryKeyName02, New Dictionary
End Function

Private Function ExtractCdpMethodName(ByVal jsonString As String) As String
    '例: {"method":"Network.responseReceived", "params":{...}}
    Const KeyPattern As String = """method"":"""
    Dim p As Long: p = InStr(1, jsonString, KeyPattern, vbTextCompare)
    If p = 0 Then Exit Function
    Dim StartPos As Long: StartPos = p + Len(KeyPattern)
    Dim EndPos As Long: EndPos = InStr(StartPos, jsonString, """", vbTextCompare)
    If EndPos > StartPos Then ExtractCdpMethodName = Mid(jsonString, StartPos, EndPos - StartPos)
End Function

Public Sub SetVisible(ByVal isVisible As Boolean)
    If m_pController = 0 Then Exit Sub
    Dim visVal As Long: visVal = IIf(isVisible, 1, 0)
    COM_DispCallLong1 m_pController, OFF_CTRL_put_IsVisible, visVal
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
    Debug.Print "[WV2] ResizeToParent: hFrame=0x" & Hex(hWndFrame) & _
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
    Debug.Print "  put_Bounds({0,0," & rc.Right & "," & rc.Bottom & "}) HRESULT=0x" & Hex(CLng(retV))
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
        Debug.Print "[WV2] DoTimerResize: スキップ (bounds 変化なし: {0,0," & rc.Right & "," & rc.Bottom & "})"
        Exit Sub
    End If

    m_Bounds.Left = 0
    m_Bounds.Top = 0
    m_Bounds.Right = rc.Right
    m_Bounds.Bottom = rc.Bottom

    Debug.Print "[WV2] DoTimerResize: bounds={0,0," & m_Bounds.Right & "," & m_Bounds.Bottom & "}"
    COM_SetBounds m_pController, m_Bounds

    ' 適用済み値を記録（次回のループ検出用）
    m_LastAppliedRight = rc.Right
    m_LastAppliedBottom = rc.Bottom
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
    Set m_CDPEventQueue = Nothing
    Set m_CDPResultQueue = Nothing
    Set m_BrowserEvents = Nothing
    m_DevToolsProtocolEventName = vbNullString
    m_DevToolsProtocolEventReceiverPtr = 0
    m_DevToolsProtocolEventToken = 0
    Set m_DevToolsProtocolEventSubscriptions = Nothing
    Set m_DevToolsProtocolEventReceiverHexToName = Nothing
    Set m_DevToolsProtocolHandlerThisToEventName = Nothing
    m_DTPEventHandlerThisCount = 0
    Set g_WebView2Core = Nothing
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
    Debug.Print "[WV2] CB_EnvironmentCreated: errorCode=" & ErrorCode & ", pEnv=" & Hex(pEnv)
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
    Debug.Print "[WV2] BeginControllerCreation"
    Debug.Print "  hWnd          = 0x" & Hex(m_hWndParent)
    Debug.Print "  pEnvironment  = 0x" & Hex(m_pEnvironment)
    Debug.Print "  pCtrlHandler  = 0x" & Hex(pCtrlHandler)
    Debug.Print "  vtable ptr    = 0x" & Hex(m_CtrlHandlerThis.pVTable)

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
    Debug.Print "  hr(DispCallFunc)=" & Hex(hr) & "  result(HRESULT)=" & Hex(CLng(retV))
End Sub

'----------------------------------------------------------------------
' CB_ControllerCreated
'   WV2_CtrlCB_Invoke → ここへ委譲。
'   ※ コールバック内での COM 呼び出しは VBE7.DLL クラッシュの原因になるため、
'     pCtrl を保存して m_State=2 をセットするだけにする。
'     実際の表示処理は ProcessMessages → FinishWebViewSetup で行う。
'----------------------------------------------------------------------
Friend Sub CB_ControllerCreated(ByVal ErrorCode As Long, ByVal pCtrl As LongPtr)
    Debug.Print "[WV2] CB_ControllerCreated: errorCode=" & ErrorCode & ", pCtrl=" & Hex(pCtrl)
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
    Debug.Print "[WV2] FinishWebViewSetup: pController=0x" & Hex(m_pController)

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
    Debug.Print "  put_Bounds HRESULT=0x" & Hex(CLng(retV))

    ' --- put_IsVisible (BOOL=1) ---
    Dim vVis(0 To 1) As LongPtr
    CopyMemory vVis(0), vtII, 2
    vVis(1) = 1
    Dim aT2(0) As Integer: aT2(0) = 20
    Dim aP2(0) As LongPtr: aP2(0) = VarPtr(vVis(0))
    DispCallFunc m_pController, OFF_CTRL_put_IsVisible, CC_STDCALL, vbLong, 1, aT2(0), aP2(0), retV
    Debug.Print "  put_IsVisible HRESULT=0x" & Hex(CLng(retV))

    ' --- get_CoreWebView2 (出力: ICoreWebView2*) ---
    Dim pWV2 As LongPtr
    Dim vWV2(0 To 1) As LongPtr
    CopyMemory vWV2(0), vtII, 2
    vWV2(1) = VarPtr(pWV2)
    Dim aT3(0) As Integer: aT3(0) = 20
    Dim aP3(0) As LongPtr: aP3(0) = VarPtr(vWV2(0))
    DispCallFunc m_pController, OFF_CTRL_get_CoreWebView2, CC_STDCALL, vbLong, 1, aT3(0), aP3(0), retV
    m_pWebView2 = pWV2
    Debug.Print "  get_CoreWebView2 HRESULT=0x" & Hex(CLng(retV)) & ", pWV2=0x" & Hex(m_pWebView2)

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

' Navigate(url: LPCWSTR)
Private Function COM_Navigate(ByVal pWV2 As LongPtr, ByVal Url As String) As Long
    Dim pUrl As LongPtr: pUrl = StrPtr(Url)
    Dim vUrl(0 To 1) As LongPtr
    Dim vtII As Integer: vtII = 20  ' VT_I8
    CopyMemory vUrl(0), vtII, 2     ' vt(2) + pad(6)
    vUrl(1) = pUrl                  ' value(8)
    
    Dim argT(0) As Integer: argT(0) = 20 ' vbLongLong
    Dim argP(0) As LongPtr: argP(0) = VarPtr(vUrl(0))
    Dim retV As Variant
    DispCallFunc pWV2, OFF_WV2_Navigate, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
    COM_Navigate = CLng(retV)
End Function

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

' CallDevToolsProtocolMethod(methodName, parametersAsJson, handler) ? 3 引数ともポインタ
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

'----------------------------------------------------------------------
' COM helper: ICoreWebView2::GetDevToolsProtocolEventReceiver
'----------------------------------------------------------------------
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
' Private ─ WebView2Loader.dll の検索とロード
' =====================================================================
Private Sub EnsureWebView2LoaderDll()
    ' Edge インストール先から WebView2Loader.dll をロードする
    ' 既に PATH に入っていれば CreateCoreWebView2EnvironmentWithOptions がそのまま動く。
    ' 入っていない場合、以下の典型的パスを順に試みる。
    Dim candidates(2) As String
    candidates(0) = "C:\Program Files (x86)\Microsoft\EdgeWebView\Application\"
    candidates(1) = "C:\Program Files\Microsoft\EdgeWebView\Application\"
    candidates(2) = Environ("LOCALAPPDATA") & "\Microsoft\Edge\Application\"

    Dim i As Integer
    For i = 0 To UBound(candidates)
        Dim dllPath As String
        dllPath = FindLatestEdgeDll(candidates(i))
        If Len(dllPath) > 0 Then
            LoadLibrary StrPtr(dllPath)
            Exit For
        End If
    Next i
End Sub

Private Function FindLatestEdgeDll(ByVal basePath As String) As String
    ' basePath\<バージョン>\WebView2Loader.dll を Dir で探す（最初に見つかったもの）
    If Len(basePath) = 0 Then Exit Function
    Dim ver As String: ver = Dir(basePath, vbDirectory)
    Do While Len(ver) > 0
        If ver <> "." And ver <> ".." Then
            Dim candidate As String
            candidate = basePath & ver & "\WebView2Loader.dll"
            If Len(Dir(candidate)) > 0 Then
                FindLatestEdgeDll = candidate
                Exit Function
            End If
        End If
        ver = Dir()
    Loop
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
    Set m_CDPEventQueue = New Collection
    Set m_CDPResultQueue = New Collection
    m_LastCDPResultErrorCode = 0
    Set m_BrowserEvents = Nothing
    m_DevToolsProtocolEventName = vbNullString
    m_DevToolsProtocolEventReceiverPtr = 0
    m_DevToolsProtocolEventToken = 0

    Set m_DevToolsProtocolEventSubscriptions = Nothing
    Set m_DevToolsProtocolEventReceiverHexToName = Nothing
    Set m_DevToolsProtocolHandlerThisToEventName = Nothing
    m_DTPEventHandlerThisCount = 0

    '------------------------- 設定シートの各セルから設定値を取得し、クラス内変数に適用します(予定 ToDo) -------------------------
    With ShSetting01_StartBrowser
        m_CDPMethodVTableIndex = 36   ' 既定。環境でずれる場合は CDPMethodVTableIndex で 22,37,39 等を指定
    
        ' GetDevToolsProtocolEventReceiver の vtable index（推定値。環境差が出る場合は変更用プロパティを使って調整）
        m_DevToolsProtocolEventReceiverVTableIndex = 42
    
    End With

End Sub


Private Sub Class_Terminate()
    quit
End Sub
