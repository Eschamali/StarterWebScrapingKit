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
'   Websocketの確立に必要なWindowsAPI宣言や処理を行います。WebSocketへのアップグレードまでは、ここのモジュールで行います
'   winhttp5.1 では、http系統特化のためここでは、WinHttp.dll を直接操作する低レベル操作を行います
'
'           ここをカスタマイズすることで、cookieや認証ヘッダー対応もできたりします
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                           ■■■ WindowsAPI宣言と定数定義 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Const ERROR_SUCCESS = 0
Private Const ERROR_INVALID_FUNCTION = 1
Private Const ERROR_INVALID_HANDLE = 6
Private Const ERROR_NOT_ENOUGH_MEMORY = 8

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
Private Const WINHTTP_FLAG_SYNC = &H0          ' セッションは同期　これを使用
Private Const WINHTTP_FLAG_ASYNC = &H10000000  ' セッション非同期　未実装

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
   ByVal dwContext As Long _
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
   ) As Long

Private Const WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS = &HFFFFFFFF


Private Const AGENT As String = "Client"
Private Const GET_METHOD As String = "GET"
Private Const HTTP_VERSION  As String = "HTTP/1.1"


Private hSessionHandle_ As LongPtr
Private hConnectionHandle_ As LongPtr
Private hRequestHandle_ As LongPtr
Private hwebsockethandle_ As LongPtr



'***************************************************************************************************
'                                   ■■■ プロパティ ■■■
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
'* 引数    ：Host           https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON なら、"github.com"
'            Path           https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON なら、"24000/ChromeControler-No-Selenium-WebDriver-VBAJSON"
'            port           http始まりなら80、https始まりなら、443ですが、カスタムも可能です。
'            SecureFlag     http始まりならFalse、https始まりなら、Trueですが、無視もできます。
'            Async          Trueで､非同期として､WebSocketを利用できるように処理します｡なお､WebSocketへのアップグレードまでは､同期的な処理とします｡
'---------------------------------------------------------------------------------------------------
'* 機能説明：最初はhttp通信を行って、200が返るか確認します
'* 注意事項：・secureモードが正しくないと、処理に失敗します
'            ・cookieや認証ヘッダー等は未対応です。匿名アクセス用です
'***************************************************************************************************
Public Function Connect(Host As String, Path As String, port As Long, SecureFlag As Boolean, Async As Boolean) As Boolean
    '非同期/同期の設定
    Dim UseMode As Long
    If Async Then UseMode = WINHTTP_FLAG_ASYNC Else UseMode = WINHTTP_FLAG_SYNC
    
    hSessionHandle_ = WinHttpOpen(StrPtr(AGENT), WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, 0, 0, UseMode)
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

    Dim result As Long
    result = WinHttpSendRequest(hRequestHandle_, WINHTTP_NO_ADDITIONAL_HEADERS, _
                    0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
    If (result = 0) Then GoTo quit

    result = WinHttpReceiveResponse(hRequestHandle_, 0)
    If (result = 0) Then GoTo quit

    If IsResponseStatus(200) = False Then GoTo quit
    
    '一連の処理やチェックを超えたら成功とみなす
    Connect = True
Exit Function
quit:
    Debug.Print "Http.Connect失敗"
    Connect = False
End Function

'***************************************************************************************************
'* 機能　　：HTTP通信からWebSocket通信にUpGradeします
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功/失敗
'* 引数    ：CallBackProcedure  外部から呼び出す標準モジュール内のプロシージャ名(AddressOf ○○)を指定します
'***************************************************************************************************
Public Function UpgradeToWebSocket(Optional CallBackProcedure As LongPtr) As Boolean
    Dim result As Long
    result = WinHttpSetOption(hRequestHandle_, _
                    WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0)
    If result = 0 Then GoTo quit
    
    result = WinHttpSendRequest(hRequestHandle_, _
                    WINHTTP_NO_ADDITIONAL_HEADERS, _
                    0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
    If result = 0 Then GoTo quit

    result = WinHttpReceiveResponse(hRequestHandle_, 0)
    If result = 0 Then GoTo quit
    
    '101でアップグレード確認
    If IsResponseStatus(101) = False Then GoTo quit

    'アップグレードしたハンドル値を取得します
    hwebsockethandle_ = WinHttpWebSocketCompleteUpgrade(hRequestHandle_, 0)
    If hwebsockethandle_ = 0 Then GoTo quit
    Debug.Print "Upgrade成功"

    'コールバック関数が指定されてる場合は、コールバック登録を行います
    If CallBackProcedure Then
        result = WinHttpSetStatusCallback(hwebsockethandle_, CallBackProcedure, _
                WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS, 0)
        If result < 0 Then GoTo quit
        
        Debug.Print "コールバック　登録成功"
    End If

    '全ての関門を超えたら、成功とする
    UpgradeToWebSocket = True

Exit Function
quit:
    If result = 0 Then Debug.Print "Upgrade：" & result
    If hwebsockethandle_ = 0 Then Debug.Print "Upgrade：" & Err.LastDllError
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
'* 機能　　：HTTP通信ハンドルの後始末を行います
'* 引数    ：RetentionWebsockethandle   True：Websocketのハンドル値だけは残す
'---------------------------------------------------------------------------------------------------
'* 注意事項：適切な掃除をしないと、使わなくなったWebsocketハンドルがどんどんメモリに溜まっていくため、ご注意ください。
'***************************************************************************************************
Public Sub CloseHInternetHandles(RetentionWebsockethandle As Boolean)
    If (hwebsockethandle_ <> 0 And Not (RetentionWebsockethandle)) Then
        WinHttpCloseHandle (hwebsockethandle_)
        hwebsockethandle_ = 0
    End If
    If (hRequestHandle_ <> 0) Then
        WinHttpCloseHandle (hRequestHandle_)
        hRequestHandle_ = 0
    End If
    If (hConnectionHandle_ <> 0) Then
        WinHttpCloseHandle (hConnectionHandle_)
        hConnectionHandle_ = 0
    End If
    If (hSessionHandle_ <> 0) Then
        WinHttpCloseHandle (hSessionHandle_)
        hSessionHandle_ = 0
    End If

    Dim result As Long
    result = Err.LastDllError
    If (result <> ERROR_SUCCESS) Then
          MsgBox "quitting with result: " & result
    End If
End Sub



Private Sub Class_Terminate()
    CloseHInternetHandles True
End Sub
