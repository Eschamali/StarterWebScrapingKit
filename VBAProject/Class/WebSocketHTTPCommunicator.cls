VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a1x1_HTTPCommunicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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


Private Const AGENT As String = "Client"
Private Const GET_METHOD As String = "GET"
Private Const HTTP_VERSION  As String = "HTTP/1.1"

Private hSessionHandle_ As LongPtr
Private hConnectionHandle_ As LongPtr
Private hRequestHandle_ As LongPtr
Private hwebsockethandle_ As LongPtr

Public Property Get websockethandle() As LongPtr
    websockethandle = hwebsockethandle_
End Property

'HTTP接続
Public Function Connect(Host As String, Path As String, port As Long) As Boolean

    hSessionHandle_ = WinHttpOpen(StrPtr(AGENT), WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, 0, 0, 0)
    If hSessionHandle_ = 0 Then GoTo quit
    hConnectionHandle_ = WinHttpConnect(hSessionHandle_, StrPtr(Host), port, 0)
    If hConnectionHandle_ = 0 Then GoTo quit
    hRequestHandle_ = WinHttpOpenRequest(hConnectionHandle_, _
                                    StrPtr(GET_METHOD), StrPtr(Path), StrPtr(HTTP_VERSION), 0, 0, 0)
    If hRequestHandle_ = 0 Then GoTo quit

    Dim result As Long
    result = WinHttpSendRequest(hRequestHandle_, WINHTTP_NO_ADDITIONAL_HEADERS, _
                    0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
    If (result = 0) Then GoTo quit

    result = WinHttpReceiveResponse(hRequestHandle_, 0)
    If (result = 0) Then GoTo quit

    If IsResponseStatus(200) = False Then GoTo quit
    
    Connect = True
Exit Function
quit:
    Debug.Print "Http.Connect失敗"
    Connect = False
End Function

''Httpレスポンスのステータスコード確認
Public Function IsResponseStatus(CheckStatusNumber As Long) As Boolean
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

Public Function UpgradeToWebSocket() As Boolean
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
    
    If IsResponseStatus(101) = False Then GoTo quit

    hwebsockethandle_ = WinHttpWebSocketCompleteUpgrade(hRequestHandle_, 0)
    If hwebsockethandle_ = 0 Then GoTo quit
    
'    Debug.Print "Upgrade成功"
    UpgradeToWebSocket = True

Exit Function
quit:
    If result = 0 Then Debug.Print "Upgrade：" & result
    If hwebsockethandle_ = 0 Then Debug.Print "Upgrade：" & Err.LastDllError
    UpgradeToWebSocket = False
End Function

Private Sub Class_Terminate()
    CloseHInternetHandles
End Sub

Public Sub CloseHInternetHandles()
    If (hwebsockethandle_ <> 0) Then
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
