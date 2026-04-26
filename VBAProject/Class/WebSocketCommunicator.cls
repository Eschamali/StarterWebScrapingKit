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
'                     ■■■ コールバック受信用 WindowsAPI・構造体定義 ■■■
'   (旧: WebSocketAsyncHelpers.bas に置いていたものをインスタンス管理のためここへ移動)
'***************************************************************************************************
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)

'https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ns-winhttp-winhttp_web_socket_status
Private Type WINHTTP_WEB_SOCKET_STATUS
    dwBytesTransferred As Long
    eBufferType        As Long
End Type



'***************************************************************************************************
'                           ■■■ WindowsAPI宣言と定数定義 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function WinHttpWebSocketSend Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal eBufferType As Long, _
   ByVal pvBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long

'https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ne-winhttp-winhttp_web_socket_buffer_type
Private Const WINHTTP_WEB_SOCKET_Incoming_Requests_BUFFER_TYPE = -1 '受信リクエスト。※独自ステータスコード
Private Const WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE = 0     'バイナリメッセージ全体か最後の部分。
Private Const WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE = 1    'バイナリメッセージの一部分のみ｡
Private Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE = 2       'UTF-8メッセージ全体か最後の部分。
Private Const WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE = 3      'UTF-8 メッセージの一部分のみ。
Private Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE = 4              'CloseFrame。

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

'Websocket蓄積受信状況把握に使用
Private Const BufferToAllocate As Long = 16384  '大きいと 1 メッセージあたりのフラグメント数が減り CDP 長文に有利（旧 4096）
Private Type WebSocketReceiveManage
    Buffer(BufferToAllocate - 1) As Byte    '第1引数        `WinHttpWebSocketReceive`からの受信バイトデータ
    BufferLength As Long                    '第2引数        バッファー残量　※`UpdateBufferInfo`内で計算します
    ReceiveBytes As Long                    '第3引数        `WinHttpWebSocketReceive`から受信できたサイズ
    Status As Long                          '第4引数        WINHTTP_WEB_SOCKET_STATUS.eBufferType
    CurrentPointer As Long                  '第5引数        カーソル位置　※`UpdateBufferInfo`内で計算します
    result As Long                          '戻り値         `WinHttpWebSocketReceive`自体の実行結果
    collect As Collection                   'チャンク収集   バラバラのデータを蓄積させる用
End Type


Private ExistingWebsockethandle_ As LongPtr
Private WithEvents Http As WebSocketHTTPCommunicator   ' SafeTimer の WithEvents st As SafeTimer 相当
Attribute Http.VB_VarHelpID = -1
Private CharConv As New CharacterCodeConversion
Private ViewLog As New Logger

'--- グローバル変数から移行したインスタンス変数（旧: WebSocketAsyncHelpers.bas の Global 宣言） ---
Private m_res         As WebSocketReceiveManage   ' 旧: G_res
Private m_isReceiving As Boolean                  ' 旧: isReceiving
Private m_isDataReady As Boolean                  ' 旧: isDataReady

'--- ログファイル（CDPBrowser.Class_Initialize と同じく ShSetting01_StartBrowser の「ログファイル保存」） ---
Private m_logFolderPath As String
Private m_logFileName As String

Private Sub WsLogDebug(ByVal message As String, Optional ByVal From As String = "")
    If Len(m_logFolderPath) > 0 And Len(m_logFileName) > 0 Then
        ViewLog.LogDebug message, From, m_logFolderPath, m_logFileName
    Else
        ViewLog.LogDebug message, From
    End If
End Sub

Public Sub WsLogInfo(ByVal message As String, Optional ByVal From As String = "")
    If Len(m_logFolderPath) > 0 And Len(m_logFileName) > 0 Then
        ViewLog.LogInfo message, From, m_logFolderPath, m_logFileName
    Else
        ViewLog.LogInfo message, From
    End If
End Sub

Private Sub WsLogWarn(ByVal message As String, Optional ByVal From As String = "")
    If Len(m_logFolderPath) > 0 And Len(m_logFileName) > 0 Then
        ViewLog.LogWarn message, From, m_logFolderPath, m_logFileName
    Else
        ViewLog.LogWarn message, From
    End If
End Sub


'***************************************************************************************************
'                                 ■■■ プロパティメソッド ■■■
'***************************************************************************************************
'* 機能　　：既存のwebsocketのハンドル値を設定します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：このクラスオブジェクトに設定中のWebsockethandle値
'* LET引数 ：ExistingWebsockethandle    以前取得したWebsockethandle値
'---------------------------------------------------------------------------------------------------
'* 機能説明：以前に確立したwebsocketのハンドル値を渡すことで続きのやり取りができます。
'* 注意事項：ErrorCode:12030 が来た場合は、もう切断状態です。「接続処理（Connect）からやり直す（再接続）」 しかありません。
'            別プロセスを跨ぐ使用("/x excel.exe" による別プロセス起動など)はできません...
'***************************************************************************************************
Public Property Get ReConnect() As LongPtr
    ReConnect = ExistingWebsockethandle_
End Property

Public Property Let ReConnect(ExistingWebsockethandle As LongPtr)
    ExistingWebsockethandle_ = ExistingWebsockethandle
End Property



'***************************************************************************************************
'                                 ■■■ 数値から翻訳 ■■■
'***************************************************************************************************
'* 機能    ：WINHTTP_WEB_SOCKET_BUFFER_TYPE - eBufferType 一式
'---------------------------------------------------------------------------------------------------
'* URL　　 ：https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ne-winhttp-winhttp_web_socket_buffer_type
'***************************************************************************************************
Public Property Get Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE(TypeID As Long, ErrorSource As String) As String
    Select Case TypeID
        Case 0: Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "BINARY_MESSAGE_BUFFER_TYPE"
        Case 1: Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "BINARY_FRAGMENT_BUFFER_TYPE"
        Case 2: Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "UTF8_MESSAGE_BUFFER_TYPE"
        Case 3: Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "UTF8_FRAGMENT_BUFFER_TYPE"
        Case 4: Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "CLOSE_BUFFER_TYPE"
        Case Else: Err.Raise 4 + vbObjectError, ErrorSource, "未知の WINHTTP_WEB_SOCKET_BUFFER_TYPE が来ました。" & vbCrLf & "Code：" & TypeID
    End Select
End Property

'***************************************************************************************************
'* 機能    ：WINHTTP_STATUS_CALLBACK - dwInternetStatus 一式
'---------------------------------------------------------------------------------------------------
'* 注意事項：Websocketとは関係ないコールバック定数は一旦、コメントアウトしています
'---------------------------------------------------------------------------------------------------
'* URL　　 ：https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/nc-winhttp-winhttp_status_callback
'***************************************************************************************************
Public Property Get Name__WINHTTP_STATUS_CALLBACK(dwInternetStatus As Long, ErrorSource As String) As String
    Select Case dwInternetStatus
'        Case 1: Name__WINHTTP_STATUS_CALLBACK = "RESOLVING_NAME"
'        Case 2: Name__WINHTTP_STATUS_CALLBACK = "NAME_RESOLVED"
'        Case 4: Name__WINHTTP_STATUS_CALLBACK = "CONNECTING_TO_SERVER"
'        Case 8: Name__WINHTTP_STATUS_CALLBACK = "CONNECTED_TO_SERVER"
'        Case 16: Name__WINHTTP_STATUS_CALLBACK = "SENDING_REQUEST"
'        Case 32: Name__WINHTTP_STATUS_CALLBACK = "REQUEST_SENT"
'        Case 64: Name__WINHTTP_STATUS_CALLBACK = "RECEIVING_RESPONSE"
'        Case 128: Name__WINHTTP_STATUS_CALLBACK = "RESPONSE_RECEIVED"
'        Case 256: Name__WINHTTP_STATUS_CALLBACK = "CLOSING_CONNECTION"
'        Case 512: Name__WINHTTP_STATUS_CALLBACK = "CONNECTION_CLOSED"
'        Case 1024: Name__WINHTTP_STATUS_CALLBACK = "HANDLE_CREATED"
'        Case 2048: Name__WINHTTP_STATUS_CALLBACK = "HANDLE_CLOSING"
'        Case 4096: Name__WINHTTP_STATUS_CALLBACK = "DETECTING_PROXY"
'        Case 16384: Name__WINHTTP_STATUS_CALLBACK = "REDIRECT"
'        Case 32768: Name__WINHTTP_STATUS_CALLBACK = "INTERMEDIATE_RESPONSE"
'        Case 65536: Name__WINHTTP_STATUS_CALLBACK = "SECURE_FAILURE"
'        Case 131072: Name__WINHTTP_STATUS_CALLBACK = "HEADERS_AVAILABLE"
'        Case 262144: Name__WINHTTP_STATUS_CALLBACK = "DATA_AVAILABLE"
        Case 524288: Name__WINHTTP_STATUS_CALLBACK = "READ_COMPLETE"
        Case 1048576: Name__WINHTTP_STATUS_CALLBACK = "WRITE_COMPLETE"
'        Case 2097152: Name__WINHTTP_STATUS_CALLBACK = "REQUEST_ERROR"
'        Case 4194304: Name__WINHTTP_STATUS_CALLBACK = "SENDREQUEST_COMPLETE"
'        Case 16777216: Name__WINHTTP_STATUS_CALLBACK = "GETPRO_YFORURL_COMPLETE"
        Case 33554432: Name__WINHTTP_STATUS_CALLBACK = "CLOSE_COMPLETE"
        Case 67108864: Name__WINHTTP_STATUS_CALLBACK = "SHUTDOWN_COMPLETE"
'        Case 268435456: Name__WINHTTP_STATUS_CALLBACK = "SETTINGS_WRITE_COMPLETE"
'        Case 536870912: Name__WINHTTP_STATUS_CALLBACK = "SETTINGS_READ_COMPLETE"
'        Case Else: Err.Raise 3 + vbObjectError, ErrorSource, "未知の WINHTTP_CALLBACK_STATUS が来ました。" & vbCrLf & "Code：" & dwInternetStatus
        Case Else: Name__WINHTTP_STATUS_CALLBACK = "Unknow_Code:" & dwInternetStatus
    End Select
End Property



'***************************************************************************************************
'                                   ■■■ 公開メソッド ■■■
'***************************************************************************************************
'* 機能　　：WebSocket 接続を確立し、コールバックを登録します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：確立した WebSocket のハンドル値（失敗時は 0）
'* 引数    ：Host       接続先ホスト名
'            Path       接続先パス
'            port       http=80、https=443、カスタムも可（省略時=443）
'            SecureMode https なら True（省略時=True）
'---------------------------------------------------------------------------------------------------
'* 機能説明：常に非同期モードで動作します。コールバックは Http_OnCallback で自動受信されます。
'* 注意事項：・SecureMode が正しくないと失敗します
'            ・Cookie や認証ヘッダー等は未対応です（匿名アクセス用）
'***************************************************************************************************
Public Function Init(Host As String, Path As String, Optional port As Long = 443, Optional SecureMode As Boolean = True) As LongPtr
    Debug.Print "[Init] Step1: 接続開始"
    Set Http = New WebSocketHTTPCommunicator
    If Http.Connect(Host, Path, port, SecureMode) = False Then GoTo quit
    Debug.Print "[Init] Step2: Connect 成功"

    If Http.UpgradeToWebSocket() = False Then GoTo quit
    Debug.Print "[Init] Step3: UpgradeToWebSocket 成功"

    Init = Http.websockethandle
    Debug.Print "[Init] Step4: 完了。ハンドル=" & Init
Exit Function

quit:
    Debug.Print "[Init] FAIL: ゴートした"
    Init = 0
End Function

'***************************************************************************************************
'* 機能　　：受け取ったデータで、WebSocketへ送信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：SendData   String OR Byte 型を想定
'* 返り値  ：実行結果コード   ※成功時は、0です
'---------------------------------------------------------------------------------------------------
'* 機能説明：型判定を行い、テキストもWideCharToMultiByte を介したバイト配列変換を使用します
'* 注意事項：・テキストも所詮、バイナリデータなのでこのような方式にしています。
'            ・非同期モードでも使えるようにしていますがその場合、実行結果コードは常に0です。コールバック側で実際に確認する必要があります。
'***************************************************************************************************
Public Function SendMessage(SendData As Variant) As Long
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle
    
    '必要に応じて、バイト配列として変換する
    Dim UseData() As Byte
    Select Case VarType(SendData)
        Case vbString: UseData = CharConv.BytesFromString(CStr(SendData))   'テキストをバイト配列に変換する
        Case vbByte: UseData = SendData                                     'そのまんま使用
        Case Else: Err.Raise vbObjectError + 1, "WebSocketCommunicator.SendMessage", "この操作は、String型かByte型のみ可能です。" & vbCrLf & "数値のみで渡してる場合は事前に、`CStr`を介す必要があります。"
    End Select

    'バイト配列のサイズを取得
    Dim messageByteSize As Long: messageByteSize = UBound(UseData) + 1

    '求めたサイズで一括送信し、実行結果コードを返却
    '※呼び出し側で、エラーかの確認をしてください
    SendMessage = WinHttpWebSocketSend(UseWebsockethandle, _
                        WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                        VarPtr(UseData(0)), messageByteSize)
End Function



'***************************************************************************************************
'* 機能　　：WebSocketからの非同期受信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：GetBinaryMode
'               `True`  バイト配列として受け取る場合
'               `False` テキストとして受け取る場合
'
'            ResultCode     WinHttpWebSocketReceive の実行結果。実質返り値です。
'* 返り値  ：WebSocketからの受信メッセージ。※エラーだと、vbnullstring です
'---------------------------------------------------------------------------------------------------
'* 機能説明：・非同期モードではありますが実際は、コールバック側の処理が終わったら、もとの呼び出し側に戻って処理を続けるというシングル仕様です
'            ・同期受信処理と違って、受信データがなくてもブロッキングしません
'
'* 注意事項：・受信データが空かつ既に予約済みの場合に更にこのメソッドを呼ぶと、4317エラーが返ります
'            ・これを呼び出すと、全メッセージを受信するのではなく、呼び出す度に1回1回、受信される仕組みです
'            ・WinHttpWebSocketReceive に渡す際は、グローバル引数として渡す必要があります
'***************************************************************************************************
Public Function GetAsyncMessage(Optional GetBinaryMode As Boolean, Optional ByRef ResultCode As Long) As Variant
    If Not (Http Is Nothing) Then
        Http.DrainPostedCallbacks
        WsLogDebug "CallbackHitCount=" & Http.CallbackHitCount, "WebSocketCommunicator.GetAsyncMessage"
        If Http.DroppedCallbackCount > 0 Then
            WsLogWarn "WinHttp コールバックキュー溢れ: " & Http.DroppedCallbackCount & "件", "WebSocketCommunicator.GetAsyncMessage"
            Http.ResetDroppedCallbackCount
        End If
    End If

    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle
    
    'エラーコードの初期化
    ResultCode = 0
    
    '既に受信予約済みなら、新規予約はせずコールバック待ち
    If m_isDataReady Then
        WsLogDebug "受信予約済みのため、コールバック到着待ちです。", "WebSocketCommunicator.GetAsyncMessage"
        If Not (Http Is Nothing) Then Http.DrainPostedCallbacks
        If Not (m_isReceiving) Then
            GetAsyncMessage = vbNullString
            Exit Function
        End If
        m_isDataReady = False
    End If

    '非同期処理側で既に、受信済みデータがあるか？
    If Not (m_isReceiving) Then
        'ない場合は、初期化
        InitializeBuffer m_res.Buffer, m_res.CurrentPointer, m_res.BufferLength
        Set m_res.collect = New Collection

        '非同期モードの場合は、こっちも追加で初期化しておく
        m_res.Status = WINHTTP_WEB_SOCKET_Incoming_Requests_BUFFER_TYPE
        m_res.ReceiveBytes = 0
    End If

    'データが尽きるまで、レスポンス受信ループ
    Dim tmp() As Byte, InsiderRes As WebSocketReceiveManage
    Do
        '非同期処理側で既に、受信済みデータがある場合は、処理を飛ばします
        If Not (m_isReceiving) Then
            m_res.result = WinHttpWebSocketReceive( _
                            UseWebsockethandle, m_res.Buffer(m_res.CurrentPointer), m_res.BufferLength, _
                            m_res.ReceiveBytes, m_res.Status)

            '既に予約済み(4317)は非同期待機中として扱う
            If m_res.result = 4317 Then
                m_isDataReady = True
                ResultCode = 0
                GetAsyncMessage = vbNullString
                Exit Function
            End If

            'エラーなら即抜け
            If m_res.result Then ResultCode = m_res.result: Exit Function

            '予約のみ完了（まだ受信データなし）の場合は、次回コールバック待ち
            If m_res.Status = WINHTTP_WEB_SOCKET_Incoming_Requests_BUFFER_TYPE And m_res.ReceiveBytes = 0 Then
                m_isDataReady = True
                GetAsyncMessage = vbNullString
                Exit Function
            End If
        End If

        If Not (Http Is Nothing) Then Http.DrainPostedCallbacks
    
        'インスタンス変数と仲介しながら、バッファー管理処理へ移る
        tmp = IntermediaryProcessing_BufferManagement(InsiderRes)
    Loop Until InsiderRes.Status = WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE Or InsiderRes.Status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Or InsiderRes.Status = WINHTTP_WEB_SOCKET_Incoming_Requests_BUFFER_TYPE

    'テキスト変換するか判断して、返却
    If InsiderRes.Status = WINHTTP_WEB_SOCKET_Incoming_Requests_BUFFER_TYPE Then
        GetAsyncMessage = vbNullString
    ElseIf GetBinaryMode Then
        GetAsyncMessage = tmp
    Else
        GetAsyncMessage = Split(CharConv.BytesToString(tmp), vbNullChar)(0)
    End If
End Function



'***************************************************************************************************
'* 機能    ：WinHttp コールバックイベントハンドラ（SafeTimer の st_TimerCall 相当）
'---------------------------------------------------------------------------------------------------
'*           旧 WebSocketAsyncHelpers::WebSocketCallback の処理をここに移植
'*           G_res → m_res、isReceiving → m_isReceiving、isDataReady → m_isDataReady
'***************************************************************************************************
Private Sub Http_OnCallback(ByVal HINTERNET As LongPtr, _
                             ByVal dwInternetStatus As Long, _
                             ByVal lpvStatusInformation As LongPtr, _
                             ByVal dwStatusInformationLength As Long)
    Const ErrorSource As String = "WebSocketCommunicator.Http_OnCallback"
    
    'WebSocket 関連コールバックのみ処理する
    Select Case dwInternetStatus
        Case 524288, 1048576, 67108864, 33554432

            'WINHTTP_WEB_SOCKET_STATUS のポインタを基にコピー
            Dim WebSocketStatus As WINHTTP_WEB_SOCKET_STATUS
            If lpvStatusInformation <> 0 And dwStatusInformationLength >= LenB(WebSocketStatus) Then
                CopyMemory WebSocketStatus, ByVal lpvStatusInformation, LenB(WebSocketStatus)
            Else
                'CLOSE/SHUTDOWN 等では statusInformation が来ないケースがある
                WebSocketStatus.dwBytesTransferred = 0
                WebSocketStatus.eBufferType = WINHTTP_WEB_SOCKET_Incoming_Requests_BUFFER_TYPE
                WsLogWarn "statusInformation が Null/不足のため既定値で処理します。status=" & dwInternetStatus & ", len=" & dwStatusInformationLength, ErrorSource
            End If
            
            WsLogDebug "------------ WINHTTP_WEB_SOCKET_STATUS ------------", ErrorSource
            WsLogDebug "Bytes：" & WebSocketStatus.dwBytesTransferred, ErrorSource
            WsLogDebug "Type ：" & Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE(WebSocketStatus.eBufferType, ErrorSource) & "(" & WebSocketStatus.eBufferType & ")", ErrorSource
            WsLogDebug "---------------------------------------------------", ErrorSource
            WsLogDebug "WINHTTP_STATUS_CALLBACK：" & Name__WINHTTP_STATUS_CALLBACK(dwInternetStatus, ErrorSource) & "(" & dwInternetStatus & ")", ErrorSource

            Select Case dwInternetStatus
                Case 524288    'READ_COMPLETE
                    '受信バッファ管理情報の更新は READ_COMPLETE のときだけ行う。
                    'WRITE_COMPLETE 等でここを上書きすると、巨大レスポンス分割受信中に m_res が壊れる。
                    m_res.Status = WebSocketStatus.eBufferType
                    m_res.ReceiveBytes = WebSocketStatus.dwBytesTransferred
                    m_isReceiving = True
                    m_isDataReady = False
                    WsLogInfo "非同期処理により、受信メッセージを格納しました。", ErrorSource
                Case 1048576   'WRITE_COMPLETE
                    WsLogInfo "非同期処理により、送信の確認が取れました。", ErrorSource
                Case 67108864  'SHUTDOWN_COMPLETE
                    WsLogInfo "WebSocket の接続が正常にシャットダウンされました。", ErrorSource
                Case 33554432  'CLOSE_COMPLETE
                    WsLogInfo "WebSocket を閉じました。", ErrorSource
            End Select

        Case Else
            WsLogWarn "WebSocket 関連以外のコールバックが来たようです。WINHTTP_STATUS_CALLBACK：" & dwInternetStatus, ErrorSource
    End Select
End Sub

'***************************************************************************************************
'* 機能　　：Websocket通信を終了させます
'---------------------------------------------------------------------------------------------------
'* 引数    ：ForceClose
'               `True`  必ず、ハンドル破棄まで突っ走ります
'               `False` 途中でエラーがある場合は、そこで終わります
'***************************************************************************************************
Public Function CloseWebSocket(Optional ForceClose As Boolean) As Boolean
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle

    Dim result As Long
    result = WinHttpWebSocketClose(UseWebsockethandle, _
                        WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS, 0, 0)
    If result And Not (ForceClose) Then GoTo quit

    ' サーバーから返されたクローズステータスを確認する。
    Dim CloseReasonBuffer(123) As Byte
    Dim BufferSize As Long: BufferSize = UBound(CloseReasonBuffer) + 1
    Dim Status As Integer
    Dim CloseReasonLength As Long
    
    result = WinHttpWebSocketQueryCloseStatus(UseWebsockethandle, Status, _
                        CloseReasonBuffer(0), BufferSize, CloseReasonLength)
    If result And Not (ForceClose) Then GoTo quit
    If Status <> WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS And Not (ForceClose) Then GoTo quit
    
    'ハンドルを破棄
    Set Http = New WebSocketHTTPCommunicator
    Http.CloseHWebsocketHandle UseWebsockethandle

    CloseWebSocket = True
Exit Function

quit:
If result Then Debug.Print result
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



'***************************************************************************************************
'                                ■■■ ヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：バッファ初期化
'***************************************************************************************************
Private Sub InitializeBuffer(responseBuffer() As Byte, CurrentPointer As Long, BufferLength As Long)
    Erase responseBuffer
    CurrentPointer = 0
    BufferLength = UBound(responseBuffer) + 1
End Sub

'***************************************************************************************************
'* 機能　　：バッファ管理情報更新
'***************************************************************************************************
Private Sub UpdateBufferInfo(CurrentPointer As Long, BufferLength As Long, ReceiveBytes)
    CurrentPointer = CurrentPointer + ReceiveBytes
    BufferLength = BufferLength - ReceiveBytes
End Sub

'***************************************************************************************************
'* 機能　　：チャンクごとに区切られたコレクションバイト配列を一気にUTF-8変換します
'---------------------------------------------------------------------------------------------------
'* 引数    ：collect  コレクションバイト配列
'* 返り値  ：合体したバイト配列
'***************************************************************************************************
Private Function MakeResponseFullBytes(collect As Collection) As Byte()
    Dim totalSize As Long
    Dim chunk As Variant
    Dim fullBytes() As Byte
    Dim currentPos As Long
    Dim i As Long
    
    ' 1. まず合計サイズを計算する
    totalSize = 0
    For Each chunk In collect
        ' 配列のサイズを加算（UBound+1）
        totalSize = totalSize + (UBound(chunk) + 1)
    Next chunk
    
    ' 2. 巨大な配列を確保
    ReDim fullBytes(0 To totalSize - 1)

    ' 3. バイトデータをコピーして結合
    currentPos = 0
    For Each chunk In collect
        For i = 0 To UBound(chunk)
            fullBytes(currentPos) = chunk(i)
            currentPos = currentPos + 1
        Next
    Next

    '4. 返却
    MakeResponseFullBytes = fullBytes
End Function

'***************************************************************************************************
'* 機能　　：クラス内ユーザー定義型：WebSocketReceiveManage とグローバル変数側ユーザー定義型：G_WebSocketReceiveManage の橋渡し役を担います
'---------------------------------------------------------------------------------------------------
' *返り値  ：Param.Status(WINHTTP_WEB_SOCKET_BUFFER_TYPE)が、
'            ・0,2 ：完成品のバイナリデータ
'            ・他：Empty
'
'* 引数    ：ユーザー定義型 - WebSocketReceiveManage
'---------------------------------------------------------------------------------------------------
'* 機能説明：バッファー管理を行うための仲介処理をします
'***************************************************************************************************
Private Function IntermediaryProcessing_BufferManagement(ByRef tmp As WebSocketReceiveManage) As Byte()
    Dim i As Long
    
    'm_res（インスタンス変数）からクラス内ローカル変数へコピー
    With tmp
        For i = LBound(.Buffer) To UBound(.Buffer)
            .Buffer(i) = m_res.Buffer(i)
        Next
        .BufferLength = m_res.BufferLength
        Set .collect = m_res.collect
        .CurrentPointer = m_res.CurrentPointer
        .ReceiveBytes = m_res.ReceiveBytes
        .result = m_res.result
        .Status = m_res.Status
    End With
    
    'バッファー管理へ移し、結果を返却
    IntermediaryProcessing_BufferManagement = BufferManagement(tmp)
    
    '受信済みフラグをリセット
    m_isReceiving = False

    'バッファー処理した結果を m_res へ反映
    With m_res
        For i = LBound(.Buffer) To UBound(.Buffer)
            .Buffer(i) = tmp.Buffer(i)
        Next
        .BufferLength = tmp.BufferLength
        Set .collect = tmp.collect
        .CurrentPointer = tmp.CurrentPointer
        .ReceiveBytes = tmp.ReceiveBytes
        .result = tmp.result
        .Status = tmp.Status
    End With
End Function

'***************************************************************************************************
'* 機能　　：WinHttpWebSocketReceive のバッファー管理を担います
'---------------------------------------------------------------------------------------------------
' *返り値  ：Param.Status(WINHTTP_WEB_SOCKET_BUFFER_TYPE)が、
'            ・0,2 ：完成品のバイナリデータ
'            ・他：Empty
'
'* 引数    ：ユーザー定義型 - WebSocketReceiveManage
'---------------------------------------------------------------------------------------------------
'* 機能説明：デバッグログによるストリーミング受信状況の可視化を行っています。
'***************************************************************************************************
Private Function BufferManagement(ByRef Param As WebSocketReceiveManage) As Byte()
    'ログ把握用
    Const TraceName As String = "WebSocketCommunicator.BufferManagement"

    'バッファー情報の更新
    UpdateBufferInfo Param.CurrentPointer, Param.BufferLength, Param.ReceiveBytes


    '========================= 受信ステータス値　把握用 =========================
    WsLogDebug "------------ WebSocket 受信ステータス ------------", TraceName
    WsLogDebug "バッファー残量 ：" & Param.BufferLength & " byte", TraceName
    WsLogDebug "受信データ量　 ：" & Param.ReceiveBytes & " byte", TraceName
    WsLogDebug "バッファー位置 ：" & Param.CurrentPointer & " byte", TraceName
    WsLogDebug "WebSocketStatus：" & Param.Status, TraceName
    WsLogDebug "---------------------------------------------------", TraceName
    '============================================================================


    'WINHTTP_WEB_SOCKET_STATUS.eBufferType による分岐処理
    '※4 + 1種類のタイプが出ますが、4種類の部分はどっちもバイナリ処理のため、統一します。
    Select Case Param.Status
        '※非同期用に特別に用意したコードです
        Case WINHTTP_WEB_SOCKET_Incoming_Requests_BUFFER_TYPE
            m_isDataReady = True
            WsLogInfo "まだ、受信データがありません。受信予約を行いました。", TraceName
        
        Case WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE
            WsLogInfo "全ての" & IIf(Param.Status = WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, "バイナリ", "UTF-8") & "データが届きました。", TraceName
            
            '1. 最後のパケットデータを追加
            Param.collect.Add Param.Buffer
            WsLogInfo "最後のバッファーを一時Collectionキューに追加しました。", TraceName

            '2. 全てのチャンクを結合して、返却
            BufferManagement = MakeResponseFullBytes(Param.collect)
            WsLogInfo "チャンクを全て結合し、バイト配列として返却しました。", TraceName

            '3. 全てを受信し終えたステータスなので、初期化処理
            InitializeBuffer Param.Buffer, Param.CurrentPointer, Param.BufferLength
            Set Param.collect = New Collection
            WsLogInfo "バッファーを確保しました。", TraceName
        
        Case WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE, WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE
            '1. バッファー残量確認
            WsLogWarn "一部の" & IIf(Param.Status = WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE, "バイナリ", "UTF-8") & "データを受け取りました。完全なデータの受信には再度、受信リクエストを送る必要があります。", TraceName
            If Param.BufferLength <= 0 Then
                WsLogInfo "バッファーを使い果たしました。", TraceName
                
                '2. 満たしたパケットをコレクションへ追加
                Param.collect.Add Param.Buffer
                WsLogInfo "満たしたバッファーを一時Collectionキューに追加しました。", TraceName

                '3. バッファー初期化して、空きを確保
                InitializeBuffer Param.Buffer, Param.CurrentPointer, Param.BufferLength
                WsLogInfo "バッファーを確保しました。", TraceName
            Else
                WsLogInfo "バッファーに空きがあるため、次の受信処理に持ち越します。", TraceName
            End If

        'サーバーが閉じるフレームを送信しました。
        Case WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE
            WsLogInfo "切断リクエストが来たため、WebSocket 処理を終了します...", TraceName
            
            '具体的な処理はまだ未実装
        
        '`WINHTTP_WEB_SOCKET_STATUS.eBufferType`未定義のコード
        Case Else
            WsLogWarn "`WINHTTP_WEB_SOCKET_STATUS.eBufferType`未定義のコードが来てます：" & Param.Status, TraceName
    End Select
End Function



Private Sub Class_Initialize()
    Const FromProcedureName As String = "WebSocketCommunicator.Class_Initialize"
    On Error Resume Next
    With ShSetting01_StartBrowser
        Dim logFolder As String
        logFolder = Trim$(CStr(.Range(.UseRangeName(9, FromProcedureName)).value))
        If Len(logFolder) > 0 Then
            If Dir(logFolder, vbDirectory) = "" Then MkDir logFolder
            If Dir(logFolder, vbDirectory) <> "" Then
                m_logFolderPath = logFolder
                m_logFileName = "WebSocket_" & CStr(Year(Now)) & Format$(Month(Now), "00") & Format$(Day(Now), "00") & ".log"
            End If
        End If
    End With
    On Error GoTo 0
End Sub
