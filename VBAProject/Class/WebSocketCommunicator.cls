VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebSocketCommunicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'                      Websocketとのやり取りの核となる処理です
'   winhttp5.1 では、http系統特化のためここでは、WinHttp.dll を直接操作する低レベル操作を行います
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                           ■■■ WindowsAPI宣言と定数定義 ■■■
'***************************************************************************************************
Private Const NO_ERROR As Long = 0
Private Declare PtrSafe Function WinHttpWebSocketSend Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal eBufferType As Long, _
   ByVal pvBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long

'https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ne-winhttp-winhttp_web_socket_buffer_type
Private Const WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE = 0 'バイナリメッセージ全体か最後の部分。
Private Const WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE = 1 'バイナリメッセージの一部分のみ｡
Private Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE = 2 'UTF-8メッセージ全体か最後の部分。
Private Const WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE = 3 'UTF-8 メッセージの一部分のみ。
Private Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE = 4 'CloseFrame。

Private Declare PtrSafe Function WinHttpWebSocketReceive Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef pvBuffer As Any, _
   ByVal dwBufferLength As Long, _
   ByRef pdwBytesRead As Long, _
   ByRef peBufferType As Long _
   ) As Long

Private Declare PtrSafe Function WinHttpWebSocketClose Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal usStatus As Integer, _
   ByVal pvReason As LongPtr, _
   ByVal dwReasonLength As Long _
   ) As Long
Private Const WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS = 1000
Private Const WINHTTP_WEB_SOCKET_ENDPOINT_TERMINATED_CLOSE_STATUS = 1001
Private Const WINHTTP_WEB_SOCKET_PROTOCOL_ERROR_CLOSE_STATUS = 1002
Private Const WINHTTP_WEB_SOCKET_INVALID_DATA_TYPE_CLOSE_STATUS = 1003
Private Const WINHTTP_WEB_SOCKET_EMPTY_CLOSE_STATUS = 1005
Private Const WINHTTP_WEB_SOCKET_ABORTED_CLOSE_STATUS = 1006
Private Const WINHTTP_WEB_SOCKET_INVALID_PAYLOAD_CLOSE_STATUS = 1007
Private Const WINHTTP_WEB_SOCKET_POLICY_VIOLATION_CLOSE_STATUS = 1008
Private Const WINHTTP_WEB_SOCKET_MESSAGE_TOO_BIG_CLOSE_STATUS = 1009
Private Const WINHTTP_WEB_SOCKET_UNSUPPORTED_EXTENSIONS_CLOSE_STATUS = 1010
Private Const WINHTTP_WEB_SOCKET_SERVER_ERROR_CLOSE_STATUS = 1011
Private Const WINHTTP_WEB_SOCKET_SECURE_HANDSHAKE_ERROR_CLOSE_STATUS = 1015

Private Declare PtrSafe Function WinHttpWebSocketQueryCloseStatus Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef usStatus As Integer, _
   ByRef pvReason As Any, _
   ByVal dwReasonLength As Long, _
   ByRef pdwReasonLengthConsumed As Long _
   ) As Long

''Websocket受信に使用(自作構造体)
Private Type response
    Buffer(4095) As Byte     '第1引数
    BufferLength As Long    '第2引数
    ReceiveBytes As Long    '第3引数
    Status As Long              '第4引数
    CurrentPointer As Long  '第5引数
    result As Long               '戻り値
    collect As Collection
End Type

Public Http As WebSocketHTTPCommunicator
Private GetErrorMes As New WinApiError
Private ExistingWebsockethandle_ As LongPtr


 
'***************************************************************************************************
'                                   ■■■ 公開メソッド ■■■
'***************************************************************************************************
'* 機能　　：websocketのハンドル値を新規に取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：指定したURIのwebsocketのハンドル値　※失敗時は、0です
'* 引数    ：Host           https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON なら、"github.com"
'            Path           https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON なら、"24000/ChromeControler-No-Selenium-WebDriver-VBAJSON"
'            port           http始まりなら80、https始まりなら、443ですが、カスタムも可能です。
'            SecureMode     http始まりならFalse、https始まりなら、Trueですが、無視もできます。
'---------------------------------------------------------------------------------------------------
'* 機能説明：ヘルパークラスを伝って、websocketのハンドル値を取得する仕組みとなっています
'* 注意事項：・secureモードが正しくないと、処理に失敗します
'            ・cookieや認証ヘッダー等は未対応です。匿名アクセス用です
'            ・既存のwebsocketのハンドル値を使う際は、次項のメソッドをご利用下さい
'            ・引数のデフォルトOptional構成は、httpsプロトコルです
'***************************************************************************************************
Public Function Init(Host As String, Path As String, Optional port As Long = 443, Optional SecureMode As Boolean = True) As LongPtr
    Set Http = New WebSocketHTTPCommunicator
    If Http.Connect(Host, Path, port, SecureMode) = False Then GoTo quit
    If Http.UpgradeToWebSocket() = False Then GoTo quit

    '成功したときのWebSocketハンドル値を返す
    Init = Http.websockethandle
Exit Function
quit:
    Init = 0
End Function

'クロームへメッセージ送信
Public Function SendMessage(message As String) As Boolean

    message = StrConv(message, vbFromUnicode)
    Dim ary() As Byte: ary = message
    Dim messageByteSize As Long: messageByteSize = LenB(message)
    Dim result As Long
    result = WinHttpWebSocketSend(Http.websockethandle, _
                        WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                        StrPtr(message), messageByteSize)
    If result = NO_ERROR Then
        SendMessage = True
    Else
        Debug.Print "送信エラー発生：(戻り値)" & result & ":(LastDllError)" & Err.LastDllError
        SendMessage = False
    End If
    
End Function

'クロームからのレスポンス(JSON)受信
Public Function GetMessage() As String
    Dim res As response
    InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength
    Set res.collect = New Collection

    Do 'レスポンス受信ループ
        res.result = WinHttpWebSocketReceive( _
                        Http.websockethandle, res.Buffer(res.CurrentPointer), res.BufferLength, _
                        res.ReceiveBytes, res.Status)
        If res.result <> NO_ERROR Then GoTo quit

        If res.Status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then
            res.collect.add res.Buffer
            GetMessage = MakeResponseMessage(res.collect)
            Exit Function
        End If
        
        UpdateBufferInfo res.CurrentPointer, res.BufferLength, res.ReceiveBytes
        If (res.BufferLength <= 0) Then
            res.collect.add res.Buffer
            InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength
        End If
    Loop
    
quit:
Debug.Print "受信エラー発生：(戻り値)" & res.result
GetMessage = "受信エラー発生"
End Function
'バッファ管理情報更新
Private Sub UpdateBufferInfo(CurrentPointer As Long, BufferLength As Long, ReceiveBytes)
    CurrentPointer = CurrentPointer + ReceiveBytes
    BufferLength = BufferLength - ReceiveBytes
End Sub
'バッファ初期化
Private Sub InitializeBuffer(responseBuffer As Variant, CurrentPointer As Long, BufferLength As Long)
    Erase responseBuffer
    CurrentPointer = 0
    BufferLength = UBound(responseBuffer) + 1
End Sub

'フラグメント組み立て：文字コードをANSI→UTF16へ変換：Null除去
Private Function MakeResponseMessage(collect As Collection) As String
    Dim ary As Variant
    Dim message As String
    For Each ary In collect
        message = message & ary
    Next
    message = StrConv(message, vbUnicode)
    MakeResponseMessage = Left(message, InStr(message, vbNullChar) - 1)
End Function



Public Function CloseWebSocket() As Boolean
    Dim result As Long
    result = WinHttpWebSocketClose(Http.websockethandle, _
                        WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS, 0, 0)
    If result <> NO_ERROR Then GoTo quit

    ' サーバーから返されたクローズステータスを確認する。
    Dim CloseReasonBuffer(123) As Byte
    Dim BufferSize As Long: BufferSize = UBound(CloseReasonBuffer) + 1
    Dim Status As Integer
    Dim CloseReasonLength As Long
    
    result = WinHttpWebSocketQueryCloseStatus(Http.websockethandle, Status, _
                        CloseReasonBuffer(0), BufferSize, CloseReasonLength)
    If result <> NO_ERROR Then GoTo quit
    If Status <> WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS Then GoTo quit

    CloseWebSocket = True
Exit Function

quit:
If result <> NO_ERROR Then Debug.Print result
Debug.Print "The server closed the connection with status code: " & Status
Dim strCloseReason As String
Dim i As Long
For i = LBound(CloseReasonBuffer) To UBound(CloseReasonBuffer)
    If CloseReasonBuffer(i) <> 0 Then
      strCloseReason = strCloseReason & Chr(CloseReasonBuffer(i))
    End If
Next
If Len(strCloseReason) > 0 Then
    Debug.Print " and reason: " & strCloseReason
End If
CloseWebSocket = False
End Function
