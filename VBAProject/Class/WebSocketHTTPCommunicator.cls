VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebSocketHTTPCommunicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   Websocketの確立に必要なWindowsAPI宣言や処理を行います。WebSocketへのアップグレードまでは、ここのクラスモジュールで行います
'   winhttp5.1 では、http系統特化のためここでは、WinHttp.dll を直接操作する低レベル操作を行います
'
'           ここをカスタマイズすることで、cookieや認証ヘッダー対応もできたりします
'***************************************************************************************************
Option Explicit



'WebSocketに関するステータスグルーピング
'https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ns-winhttp-winhttp_web_socket_status
Private Type WINHTTP_WEB_SOCKET_STATUS
    dwBytesTransferred As Long
    eBufferType        As Long
End Type



'***************************************************************************************************
'   !! SafeTimer の SafeTimer.cls に相当するセクション !!
'   !! 以下の Public Event / WinHttpCallbackProc は vtable の先頭ユーザーメソッド位置に固定 !!
'   !! 絶対にこのブロックより上に新しいメソッド・プロパティを追加しないこと              !!
'   !! （vtable[7] の位置がずれるとクラッシュします）                                  !!
'***************************************************************************************************
'コールバックイベント（WithEvents で受け取る側がハンドルする）
Public Event OnCallback(ByVal dwInternetStatus As Long, _
                        ByVal dwBytesTransferred As Long, _
                        ByVal eBufferType As Long, _
                        ByVal dwStatusInformationLength As Long)

'フォームインスタンス（SafeTimer の m_form 相当。ハンドル取得＆ライフタイム管理に使用）
Private m_form As New WinHttpSetStatusCallbackForm

Private Type CallbackQueueItem
    HINTERNET As LongPtr
    dwInternetStatus As Long
    hasStatusInformation As Long
    statusInformation As WINHTTP_WEB_SOCKET_STATUS
End Type
Private Const CALLBACK_QUEUE_SIZE As Long = 4096
Private m_callbackQueue(0 To CALLBACK_QUEUE_SIZE - 1) As CallbackQueueItem
Private m_queueRead As Long
Private m_queueWrite As Long
Private m_droppedCallbacks As Long
Private m_callbackHitCount As Long
Private m_notifyHwnd As LongPtr



'***************************************************************************************************
'                           ■■■ WindowsAPI宣言と定数定義 ■■■
'***************************************************************************************************
'* 機能    ：`WinHttp`関連
'***************************************************************************************************
Private Declare PtrSafe Function WinHttpOpen Lib "WinHttp" ( _
   ByVal pszAgentW As LongPtr, _
   ByVal dwAccessType As Long, _
   ByVal pszProxyW As LongPtr, _
   ByVal pszProxyBypassW As LongPtr, _
   ByVal dwFlags As Long _
   ) As LongPtr
Private Const WINHTTP_ACCESS_TYPE_DEFAULT_PROXY = 0
Private Const WINHTTP_ACCESS_TYPE_NO_PROXY = 1 '＝プロキシを介さず、すべてのホスト名を直接解決。
Private Const WINHTTP_ACCESS_TYPE_NAMED_PROXY = 3
Private Const WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY = 4
Private Const WINHTTP_NO_PROXY_BYPASS = 0
Private Const WINHTTP_FLAG_ASYNC = &H10000000  ' セッションは常に非同期（SafeTimer 方式のコールバックに必須）

Private Declare PtrSafe Function WinHttpConnect Lib "WinHttp" ( _
   ByVal hSession As LongPtr, _
   ByVal pswzServerName As LongPtr, _
   ByVal nServerPort As Long, _
   ByVal dwReserved As Long _
   ) As LongPtr
Private Const INTERNET_DEFAULT_PORT = 0           ' use the protocol-specific default
Private Const INTERNET_DEFAULT_HTTP_PORT = 80     ' use the HTTP default
Private Const INTERNET_DEFAULT_HTTPS_PORT = 443   ' use the HTTPS default

Private Declare PtrSafe Function WinHttpOpenRequest Lib "WinHttp" ( _
   ByVal hConnect As LongPtr, _
   ByVal pwszVerb As LongPtr, _
   ByVal pwszObjectName As LongPtr, _
   ByVal pwszVersion As LongPtr, _
   ByVal pwszReferrer As LongPtr, _
   ByVal ppwszAcceptTypes As LongPtr, _
   ByVal dwFlags As Long _
   ) As LongPtr
Private Const WINHTTP_NO_REFERER As LongPtr = 0
Private Const WINHTTP_DEFAULT_ACCEPT_TYPES As LongPtr = 0
Private Const WINHTTP_FLAG_SECURE As Long = &H800000

Private Declare PtrSafe Function WinHttpSetOption Lib "WinHttp" ( _
   ByVal HINTERNET As LongPtr, _
   ByVal dwOption As Long, _
   ByVal lpBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long
Private Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET = 114
Private Const WINHTTP_OPTION_SECURE_PROTOCOLS = 84
Private Const WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_1 = &H200
Private Const WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2 = &H800

Private Declare PtrSafe Function WinHttpAddRequestHeaders Lib "WinHttp" ( _
  ByVal hRequest As LongPtr, _
  ByVal lpszHeaders As LongPtr, _
  ByVal dwHeadersLength As Long, _
  ByVal dwModifiers As Long _
) As Boolean
Private Const WINHTTP_ADDREQ_FLAG_ADD = &H20000000
Private Const WINHTTP_ADDREQ_FLAG_REPLACE = &H80000000
Private Const WINHTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

Private Declare PtrSafe Function WinHttpSendRequest Lib "WinHttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal lpszHeaders As LongPtr, _
   ByVal dwHeadersLength As Long, _
   ByVal lpOptional As LongPtr, _
   ByVal dwOptionalLength As Long, _
   ByVal dwTotalLength As Long, _
   ByVal dwContext As LongPtr _
   ) As Long
Private Const WINHTTP_NO_ADDITIONAL_HEADERS = 0
Private Const WINHTTP_NO_REQUEST_DATA = 0

Private Declare PtrSafe Function WinHttpReceiveResponse Lib "WinHttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal lpReserved As LongPtr _
   ) As Long

Private Declare PtrSafe Function WinHttpQueryHeaders Lib "WinHttp" ( _
  ByVal hRequest As LongPtr, _
  ByVal dwInfoLevel As Long, _
  ByVal pwszName As LongPtr, _
  ByRef lpBuffer As Long, _
  ByRef lpdwBufferLength As Long, _
  ByRef lpdwIndex As Long _
   ) As Long
Private Const WINHTTP_QUERY_STATUS_CODE = 19
Private Const WINHTTP_QUERY_FLAG_NUMBER = &H20000000
Private Const WINHTTP_HEADER_NAME_BY_INDEX = 0
Private Const WINHTTP_NO_OUTPUT_BUFFER = 0
Private Const WINHTTP_NO_HEADER_INDEX = 0
Private Const HTTP_STATUS_CONTINUE = 100           ' リクエストの続行OK
Private Const HTTP_STATUS_SWITCH_PROTOCOLS = 101   ' サーバーがアップグレードヘッダでプロトコルを切り替えた
Private Const HTTP_STATUS_OK = 200                 ' リクエスト完了

Private Declare PtrSafe Function WinHttpWebSocketCompleteUpgrade Lib "WinHttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal pContext As LongPtr _
   ) As LongPtr

Private Declare PtrSafe Function WinHttpCloseHandle Lib "WinHttp" ( _
   ByVal hRequest As LongPtr _
   ) As Long

Private Declare PtrSafe Function WinHttpSetStatusCallback Lib "WinHttp" ( _
  ByVal hWebSocket As LongPtr, _
  ByVal lpfnInternetCallback As LongPtr, _
  ByVal dwNotificationFlags As Long, _
  ByVal dwReserved As LongPtr _
   ) As LongPtr
Private Const WINHTTP_INVALID_STATUS_CALLBACK As LongPtr = -1
Private Const WINHTTP_CALLBACK_STATUS_READ_COMPLETE = &H80000
Private Const WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE = &H100000
Private Const WINHTTP_CALLBACK_STATUS_CLOSE_COMPLETE = &H2000000
Private Const WINHTTP_CALLBACK_STATUS_SHUTDOWN_COMPLETE = &H4000000
Private Const WINHTTP_CALLBACK_FLAG_WEBSOCKET_NOTIFICATIONS As Long = _
      WINHTTP_CALLBACK_STATUS_READ_COMPLETE _
    Or WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE _
    Or WINHTTP_CALLBACK_STATUS_CLOSE_COMPLETE _
    Or WINHTTP_CALLBACK_STATUS_SHUTDOWN_COMPLETE

Private Const AGENT As String = "Client"
Private Const GET_METHOD As String = "GET"
Private Const HTTP_VERSION  As String = "HTTP/1.1"

'***************************************************************************************************
'* 機能    ：`user32`関連
'---------------------------------------------------------------------------------------------------
'* 機能説明：自インスタンスの WndProc 復元 / 既定処理委譲に使用
'***************************************************************************************************
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal msg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Const GWLP_WNDPROC As Long = -4

Private hSessionHandle_     As LongPtr
Private hConnectionHandle_  As LongPtr
Private hRequestHandle_     As LongPtr
Private hwebsockethandle_   As LongPtr
Private m_prevWndProc       As LongPtr    'このインスタンスがフックする前にあった元の WndProc。Class_Terminate でSetWindowLongPtr により自分で復元する。



'***************************************************************************************************
'                           ■■■ コールバック関連処理 ■■■
'***************************************************************************************************
'!! Do not move. This must be the 1st user-defined method (vtable[7], offset +0x38 on x64) !!
'   SafeTimer::TimerProc に相当。
'   WinHttp からのコールバックを受け取り、OnCallback イベントを発火させます。
'   引数は WINHTTP_STATUS_CALLBACK のシグネチャと一致させること。
'***************************************************************************************************
Private Sub WinHttpCallbackProc(ByVal HINTERNET As LongPtr, _
                                 ByVal dwInternetStatus As Long, _
                                 ByVal lpvStatusInformation As LongPtr, _
                                 ByVal dwStatusInformationLength As Long)
    On Error Resume Next
    m_callbackHitCount = m_callbackHitCount + 1
    EnqueueCallback HINTERNET, dwInternetStatus, lpvStatusInformation, dwStatusInformationLength
    On Error GoTo 0
End Sub

Private Sub EnqueueCallback(ByVal HINTERNET As LongPtr, _
                            ByVal dwInternetStatus As Long, _
                            ByVal lpvStatusInformation As LongPtr, _
                            ByVal dwStatusInformationLength As Long)
    Dim nextWrite As Long
    nextWrite = (m_queueWrite + 1) Mod CALLBACK_QUEUE_SIZE
    If nextWrite = m_queueRead Then
        m_droppedCallbacks = m_droppedCallbacks + 1
        Exit Sub
    End If

    With m_callbackQueue(m_queueWrite)
        .HINTERNET = HINTERNET
        .dwInternetStatus = dwInternetStatus
        .hasStatusInformation = 0
        If lpvStatusInformation <> 0 And dwStatusInformationLength >= LenB(.statusInformation) Then
            'ReadMemLongPtr（PointerAccessor）で読み取り + ビット演算で分解
#If Win64 Then
            Dim rawWsStatus As LongLong
            rawWsStatus = WinHttpCallbackRouter.ReadMemLongPtr(lpvStatusInformation)
            .statusInformation.dwBytesTransferred = CLng(rawWsStatus And &HFFFFFFFF^)
            .statusInformation.eBufferType = CLng((rawWsStatus \ &H100000000^) And &HFFFFFFFF^)
#Else
            .statusInformation.dwBytesTransferred = CLng(WinHttpCallbackRouter.ReadMemLongPtr(lpvStatusInformation))
            .statusInformation.eBufferType = CLng(WinHttpCallbackRouter.ReadMemLongPtr(lpvStatusInformation + 4))
#End If
            .hasStatusInformation = 1
        End If
    End With
    m_queueWrite = nextWrite
End Sub

Public Sub HandlePostedWinHttpCallback(ByVal packedStatusAndLength As LongPtr, ByVal packedStatusInformation As LongPtr)
    On Error Resume Next
    Dim dwInternetStatus As Long
    Dim dwStatusInformationLength As Long
    dwInternetStatus = CLng(packedStatusAndLength)
#If Win64 Then
    Dim statusInfo As WINHTTP_WEB_SOCKET_STATUS
    ' ビット演算で LongLong → Long×2 に分解
    Dim packed64 As LongLong
    packed64 = packedStatusAndLength
    dwInternetStatus = CLng(packed64 And &HFFFFFFFF^)
    dwStatusInformationLength = CLng((packed64 \ &H100000000^) And &HFFFFFFFF^)
    packed64 = packedStatusInformation
    statusInfo.dwBytesTransferred = CLng(packed64 And &HFFFFFFFF^)
    statusInfo.eBufferType = CLng((packed64 \ &H100000000^) And &HFFFFFFFF^)
#Else
    dwStatusInformationLength = 0
#End If
    m_callbackHitCount = m_callbackHitCount + 1
#If Win64 Then
    If dwStatusInformationLength >= LenB(statusInfo) Then
        EnqueueCallback 0, dwInternetStatus, VarPtr(statusInfo), LenB(statusInfo)
    Else
        EnqueueCallback 0, dwInternetStatus, 0, 0
    End If
#Else
    EnqueueCallback 0, dwInternetStatus, packedStatusInformation, dwStatusInformationLength
#End If
    ' PostMessage が短時間に大量に積まれると Enqueue だけだとキューが溢れ、READ_COMPLETE が落ちる。
    ' WM 処理のたびに即 Dispatch してキューを常に空に近づける（巨大 UTF8 フラグメント連打対策）。
    DispatchQueuedCallbacks
    On Error GoTo 0
End Sub

Public Sub DispatchQueuedCallbacks()
    On Error Resume Next
    Do While m_queueRead <> m_queueWrite
        With m_callbackQueue(m_queueRead)
            If .hasStatusInformation <> 0 Then
                RaiseEvent OnCallback(.dwInternetStatus, .statusInformation.dwBytesTransferred, .statusInformation.eBufferType, LenB(.statusInformation))
            Else
                RaiseEvent OnCallback(.dwInternetStatus, 0, -1, 0)
            End If
        End With
        m_queueRead = (m_queueRead + 1) Mod CALLBACK_QUEUE_SIZE
    Loop
    On Error GoTo 0
End Sub

Public Property Get DroppedCallbackCount() As Long
    DroppedCallbackCount = m_droppedCallbacks
End Property

Public Sub ResetDroppedCallbackCount()
    m_droppedCallbacks = 0
End Sub

Public Property Get CallbackHitCount() As Long
    CallbackHitCount = WinHttpCallbackRouter.GetWinHttpMessageCount
End Property

Public Property Get NotifyHwnd() As LongPtr
    NotifyHwnd = m_notifyHwnd
End Property

Public Sub PumpWinHttpCallbacks()
    If m_notifyHwnd <> 0 Then WinHttpCallbackRouter.PumpPostedWinHttpMessages m_notifyHwnd
End Sub

Private Sub EnsureNotifyHwnd()
    If m_notifyHwnd <> 0 Then Exit Sub
    On Error Resume Next
    m_form.Show vbModeless
    m_form.Hide
    m_notifyHwnd = m_form.Controls.Add("Forms.Frame.1").[_GethWnd]
    On Error GoTo 0
End Sub



'***************************************************************************************************
'                               ■■■ WndProc フック■■■
'***************************************************************************************************
'* 機能    ：WinHttpCallbackRouter から「あなたの hWnd を私がフックしました。元の WndProc は
'*           これです」と通知を受けるためのエントリ。各インスタンスが自分の前任 WndProc を
'*           自分の状態として保持します。
'*---------------------------------------------------------------------------------------------------
'* 引数    ：hWnd          このインスタンスがフック対象とする hWnd (= m_notifyHwnd)
'*           prevWndProc   SetWindowLongPtr で置き換えた直前の WndProc アドレス
'*---------------------------------------------------------------------------------------------------
'* 注意    ：呼び元は WinHttpCallbackRouter.cls のみ想定 (Friend)。
'*           m_notifyHwnd は UpgradeToWebSocket 内ですでに自前で設定済みなので、そちらを尊重し
'*           ここでは hWnd 引数は照合用にのみ受け取り、保存は m_prevWndProc だけにする。
'***************************************************************************************************
Friend Sub RegisterAsHookTarget(ByVal hWnd As LongPtr, ByVal prevWndProc As LongPtr)
    If hWnd <> 0 And m_notifyHwnd = 0 Then m_notifyHwnd = hWnd
    m_prevWndProc = prevWndProc
End Sub

'***************************************************************************************************
'* 機能    ：このインスタンスが保持している前任 WndProc アドレスを取得します
'*---------------------------------------------------------------------------------------------------
'* Get返値 ：SetWindowLongPtr 直前の WndProc。フック未設置 / 既に復元済みなら 0。
'***************************************************************************************************
Friend Property Get prevWndProc() As LongPtr
    prevWndProc = m_prevWndProc
End Property

'***************************************************************************************************
'* 機能    ：WinHttpCallbackRouter.WinHttpBridgeWndProc から委譲されるメッセージディスパッチ。
'*           WM_APP_WINHTTP_CALLBACK は自分で消化し、それ以外は前任 WndProc に転送します。
'*---------------------------------------------------------------------------------------------------
'* 引数    ：hWnd, msg, wParam, lParam   通常の WndProc シグネチャ
'* 返り値  ：WndProc の戻り値 (LRESULT)
'*---------------------------------------------------------------------------------------------------
'* 設計意図：WinHttpCallbackRouter 側で hWnd→target を引き、メッセージ判定もそちらで
'*           行っていた。案 B ではメッセージ判定とCallWindowProc までインスタンス内で完結させ、
'*           module 側は純粋なルーティングだけに責務を絞る。
'***************************************************************************************************
Public Function DispatchHookedMessage(ByVal hWnd As LongPtr, _
                                       ByVal msg As Long, _
                                       ByVal wParam As LongPtr, _
                                       ByVal lParam As LongPtr) As LongPtr
    On Error Resume Next
    If msg = WinHttpCallbackRouter.AppWinHttpCallbackMsg Then
        HandlePostedWinHttpCallback wParam, lParam
        DispatchHookedMessage = 0
        Exit Function
    End If

    If m_prevWndProc <> 0 Then
        DispatchHookedMessage = CallWindowProc(m_prevWndProc, hWnd, msg, wParam, lParam)
    Else
        DispatchHookedMessage = 0
    End If
    On Error GoTo 0
End Function

'***************************************************************************************************
'* 機能    ：このインスタンスが SetWindowLongPtr で置き換えた WndProc を、自身で復元します。
'*           Class_Terminate および EmergencyUnregisterCallbacks 経由で
'*           WinHttpCallbackRouter.RemoveWinHttpMessageHook から呼ばれます。
'*---------------------------------------------------------------------------------------------------
'* 注意    ：Friend です。WinHttpCallbackRouter.cls からしか呼び出せません。
'*           idempotent (m_prevWndProc を 0 にクリアするので 2 回目以降は noop)。
'***************************************************************************************************
Friend Sub RestoreWndProcIfNeeded()
    On Error Resume Next
    If m_prevWndProc <> 0 And m_notifyHwnd <> 0 Then
        SetWindowLongPtr m_notifyHwnd, GWLP_WNDPROC, m_prevWndProc
    End If
    m_prevWndProc = 0
    On Error GoTo 0
End Sub



'***************************************************************************************************
'                                   ■■■ 主要プロパティ ■■■
'***************************************************************************************************
'* 機能　　：確立したwebsocketのハンドル値を取得をします
'---------------------------------------------------------------------------------------------------
'* Get返値 ：アクティブなwebsocketのハンドル値
'***************************************************************************************************
Public Property Get websockethandle() As LongPtr
    websockethandle = hwebsockethandle_
End Property



'***************************************************************************************************
'                                   ■■■ 公開メソッド ■■■
'***************************************************************************************************
'* 機能　　：通常の http-GET 接続を行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功/失敗
'* 引数    ：Host           接続先のホスト名
'            Path           接続先のパス
'            port           httpは80、httpsは443、カスタムも可
'            SecureFlag     httpsの場合はTrue
'---------------------------------------------------------------------------------------------------
'* 機能説明：常に非同期モード（WINHTTP_FLAG_ASYNC）でセッションを開く。
'*           WebSocketへのアップグレードまでのハンドシェイク自体は同期的に実行される。
'* 注意事項：・セキュアモードが正しくないと、処理に失敗します
'            ・クッキーや認証ヘッダー等は未対応です。匿名アクセス用です
'***************************************************************************************************
Public Function Connect(Host As String, Path As String, port As Long, SecureFlag As Boolean) As Boolean
    '常に非同期モードでセッションを開く（SafeTimer 方式のコールバック登録に必須）
    hSessionHandle_ = WinHttpOpen(StrPtr(AGENT), WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, 0, 0, WINHTTP_FLAG_ASYNC)
    If hSessionHandle_ = 0 Then GoTo quit
    hConnectionHandle_ = WinHttpConnect(hSessionHandle_, StrPtr(Host), port, 0)
    If hConnectionHandle_ = 0 Then GoTo quit
    
    'セキュアモードの切り替え判定
    Dim SecureMode As Long
    If SecureFlag Then SecureMode = WINHTTP_FLAG_SECURE Else SecureMode = 0

    '適用してリクエスト
    hRequestHandle_ = WinHttpOpenRequest(hConnectionHandle_, _
                                    StrPtr(GET_METHOD), StrPtr(Path), StrPtr(HTTP_VERSION), 0, 0, SecureMode)
    If hRequestHandle_ = 0 Then GoTo quit

    'WebSocket 用は UpgradeToWebSocket() で初回 Send/Receive を実行する。
    'ここではハンドル確立までを成功条件にする。
    Connect = True
Exit Function
quit:
    Debug.Print "Http.Connect失敗 LastDllError=" & Err.LastDllError
    Connect = False
End Function

'***************************************************************************************************
'* 機能　　：HTTP通信からWebSocket通信にUpGradeします
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功/失敗
'* 引数    ：なし（コールバック thunk は WinHttpCallbackRouter が VirtualAlloc 生成）
'***************************************************************************************************
Public Function UpgradeToWebSocket() As Boolean
    Dim result As Long
    Dim callbackResult As LongPtr
    If m_notifyHwnd = 0 Then
        EnsureNotifyHwnd
        Debug.Print "[Upgrade] F0: NotifyHwnd=" & m_notifyHwnd
    End If
    WinHttpCallbackRouter.InstallWinHttpMessageHook m_notifyHwnd, Me

    Dim requestCallbackResult As LongPtr
    Dim cbPtr As LongPtr: cbPtr = WinHttpCallbackRouter.GetWinHttpStatusCallbackPtr(Me)
    Debug.Print "[Upgrade] F: CallbackPtr=" & cbPtr
    If cbPtr = 0 Then GoTo quit

    result = WinHttpSetOption(hRequestHandle_, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0)
    Debug.Print "[Upgrade] A: SetOption=" & result
    If result = 0 Then GoTo quit

    result = WinHttpSendRequest(hRequestHandle_, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
    Debug.Print "[Upgrade] B: SendRequest=" & result
    If result = 0 Then GoTo quit

    result = WinHttpReceiveResponse(hRequestHandle_, 0)
    Debug.Print "[Upgrade] C: ReceiveResponse=" & result
    If result = 0 Then GoTo quit

    If IsResponseStatus(101) = False Then
        Debug.Print "[Upgrade] D: HTTP101 失敗"
        GoTo quit
    End If
    Debug.Print "[Upgrade] D: HTTP101 OK"

    hwebsockethandle_ = WinHttpWebSocketCompleteUpgrade(hRequestHandle_, m_notifyHwnd)
    Debug.Print "[Upgrade] E: WebSocketHandle=" & hwebsockethandle_ & "  NotifyHwnd=" & m_notifyHwnd
    If hwebsockethandle_ = 0 Then GoTo quit

    callbackResult = WinHttpSetStatusCallback(hwebsockethandle_, cbPtr, WINHTTP_CALLBACK_FLAG_WEBSOCKET_NOTIFICATIONS, 0)
    Debug.Print "[Upgrade] G: SetStatusCallback=" & callbackResult
    If callbackResult = WINHTTP_INVALID_STATUS_CALLBACK Then GoTo quit

    Debug.Print "[Upgrade] 成功"
    UpgradeToWebSocket = True

    'WebSocketハンドルの入手はできたので、HTTP通信関連のハンドルは破棄する
    CloseHInternetHandles
Exit Function

quit:
    Debug.Print "[Upgrade] FAIL ゴート: result=" & result & " requestCallbackResult=" & requestCallbackResult & " callbackResult=" & callbackResult & " LastDllError=" & Err.LastDllError
    On Error Resume Next
    WinHttpCallbackRouter.RemoveWinHttpMessageHook
    
    On Error GoTo 0
    UpgradeToWebSocket = False
End Function



'***************************************************************************************************
'                           ■■■ ヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：Httpレスポンスのステータスコード確認を行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功/失敗
'* 引数    ：CheckStatusNumber  期待するステータスコード
'***************************************************************************************************
Private Function IsResponseStatus(CheckStatusNumber As Long) As Boolean
    Dim statusCodeBuffer As Long
    Dim BufferSize As Long: BufferSize = 4
    Dim result As Long
    result = WinHttpQueryHeaders(hRequestHandle_, _
                    (WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER), _
                    WINHTTP_HEADER_NAME_BY_INDEX, _
                    statusCodeBuffer, BufferSize, WINHTTP_NO_HEADER_INDEX)
                    
    If result = False Then GoTo quit
    If statusCodeBuffer <> CheckStatusNumber Then GoTo quit

    IsResponseStatus = True
    
Exit Function
quit:
    If result = False Then Debug.Print result
    If statusCodeBuffer <> CheckStatusNumber Then Debug.Print _
        "ステータスコード:" & statusCodeBuffer & "/" & CheckStatusNumber & "以外"
    IsResponseStatus = False
End Function



'***************************************************************************************************
'                                   ■■■ 後始末系 ■■■
'***************************************************************************************************
'* 機能　　：HTTP通信ハンドル系統の後始末を行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebSocket も後始末する場合は、次項のプロシージャも呼び出して下さい
'***************************************************************************************************
Private Sub CloseHInternetHandles()
    WinHttpCloseHandle hRequestHandle_
    WinHttpCloseHandle hConnectionHandle_
    WinHttpCloseHandle hSessionHandle_
End Sub

'***************************************************************************************************
'* 機能　　：WebSocketハンドルの後始末を行います
'***************************************************************************************************
Public Sub CloseHWebsocketHandle()
    WinHttpSetStatusCallback hwebsockethandle_, 0, 0, 0     'コールバックを解除しておく
    WinHttpCloseHandle hwebsockethandle_
End Sub

'***************************************************************************************************
'* 機能　　：VBE でブレーク中～「リセット」直前に、イミディエイト等から呼ぶための緊急解除です。
'*           WinHttp の status callback と PostMessage 橋渡し用ウィンドウフックを外します。
'*           WebSocket ハンドルは閉じません（リセットで VBA 状態が消える前提の軽量停止）。
'---------------------------------------------------------------------------------------------------
'* 注意事項：・呼び出し後は通信状態が不整合になるため、通常フローでは使わず、デバッグ用に限定してください。
'            ・VBEリセット前提につき、インスタンス全て破棄します
'***************************************************************************************************
Public Sub EmergencyUnregisterCallbacks()
    On Error Resume Next
    WinHttpCallbackRouter.RemoveWinHttpMessageHook
    If hwebsockethandle_ <> 0 Then
        WinHttpSetStatusCallback hwebsockethandle_, 0, 0, 0
    End If
    On Error GoTo 0
End Sub



Private Sub Class_Terminate()
    WinHttpCallbackRouter.RemoveWinHttpMessageHook m_notifyHwnd
    m_notifyHwnd = 0
    CloseHInternetHandles
    CloseHWebsocketHandle
End Sub
