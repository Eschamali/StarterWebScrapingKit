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
Private Http As WebSocketHTTPCommunicator
Private CharConv As New CharacterCodeConversion
Private ViewLog As New Logger



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
        Case 2097152: Name__WINHTTP_STATUS_CALLBACK = "REQUEST_ERROR"
'        Case 4194304: Name__WINHTTP_STATUS_CALLBACK = "SENDREQUEST_COMPLETE"
'        Case 16777216: Name__WINHTTP_STATUS_CALLBACK = "GETPRO_YFORURL_COMPLETE"
        Case 33554432: Name__WINHTTP_STATUS_CALLBACK = "CLOSE_COMPLETE"
'        Case 67108864: Name__WINHTTP_STATUS_CALLBACK = "SHUTDOWN_COMPLETE"
'        Case 268435456: Name__WINHTTP_STATUS_CALLBACK = "SETTINGS_WRITE_COMPLETE"
'        Case 536870912: Name__WINHTTP_STATUS_CALLBACK = "SETTINGS_READ_COMPLETE"
'        Case Else: Err.Raise 3 + vbObjectError, ErrorSource, "未知の WINHTTP_CALLBACK_STATUS が来ました。" & vbCrLf & "Code：" & dwInternetStatus
        Case Else: Name__WINHTTP_STATUS_CALLBACK = "Unknow_Code:" & dwInternetStatus
    End Select
End Property



'***************************************************************************************************
'                                   ■■■ 公開メソッド ■■■
'***************************************************************************************************
'* 機能　　：websocketのハンドル値を新規に取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：指定したURIのwebsocketのハンドル値　※失敗時は、0です
'* 引数    ：Host               https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON なら、"github.com"
'            Path               https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON なら、"24000/ChromeControler-No-Selenium-WebDriver-VBAJSON"
'            port               http始まりなら80、https始まりなら、443ですが、カスタムも可能です。
'            SecureMode         http始まりならFalse、https始まりなら、Trueですが、無視もできます。
'            CallBackProcedure  外部から呼び出す標準モジュール内のプロシージャ名(AddressOf ○○)を指定します
'---------------------------------------------------------------------------------------------------
'* 機能説明：ヘルパークラスを伝って、websocketのハンドル値を取得する仕組みとなっています
'* 注意事項：・secureモードが正しくないと、処理に失敗します
'            ・cookieや認証ヘッダー等は未対応です。匿名アクセス用です
'            ・既存のwebsocketのハンドル値を使う際は、次項のメソッドをご利用下さい
'            ・引数のデフォルトOptional構成は、httpsプロトコルです
'***************************************************************************************************
Public Function Init(Host As String, Path As String, Optional port As Long = 443, Optional SecureMode As Boolean = True, Optional CallBackProcedure As LongPtr) As LongPtr
    '非同期/同期の設定
    Dim UseMode As Boolean
    If CallBackProcedure Then UseMode = True Else UseMode = False
    
    '通常のHTTP接続を行う
    Set Http = New WebSocketHTTPCommunicator
    If Http.Connect(Host, Path, port, SecureMode, UseMode) = False Then GoTo quit
    
    'WebSocketへの昇格処理を行う
    If Http.UpgradeToWebSocket(CallBackProcedure) = False Then GoTo quit

    '成功したときのWebSocketハンドル値を返す
    Init = Http.websockethandle
Exit Function


quit:
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
'* 機能　　：WebSocketからの同期受信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：GetBinaryMode
'               `True`  バイト配列として受け取る場合
'               `False` テキストとして受け取る場合
'
'            ResultCode     WinHttpWebSocketReceive の実行結果。実質返り値です。
'* 返り値  ：WebSocketからの受信メッセージ。※エラーだと、vbnullstring です
'---------------------------------------------------------------------------------------------------
'* 機能説明：ErrorCode：12152　対策により、WideCharToMultiByte を介したバイト配列変換を使用してます
'* 注意事項：・受信データがない状態でこれを呼び出すと、受信データが来るまで、フ　リ　ー　ズ　します。ネット切断することで、40s前後でフリーズ回避できます。
'            ・現時点は、プレーンテキスト受信のみ対応です
'            ・送信と違って受信は内部で、固定のバッファーサイズ受信によるストリーミング受信方式を取っています。そのため、受信データ量が多いと返るまで時間がかかります。
'***************************************************************************************************
Public Function GetSyncMessage(Optional GetBinaryMode As Boolean, Optional ByRef ResultCode As Long) As Variant
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle
    
    'ストリーミング受信準備による初期化
    Dim res As WebSocketReceiveManage
    InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength
    Set res.collect = New Collection
    ResultCode = 0

    'データが尽きるまで、レスポンス受信ループ
    Dim tmp() As Byte
    Do
        '受信処理
        res.result = WinHttpWebSocketReceive( _
                        UseWebsockethandle, res.Buffer(res.CurrentPointer), res.BufferLength, _
                        res.ReceiveBytes, res.Status)
        
        'エラーなら即抜け
        '※呼び出し側で、エラーかの確認をしてください
        If res.result Then ResultCode = res.result: Exit Function
        
        'バッファー管理処理へ移る
        tmp = BufferManagement(res)
    Loop Until res.Status = WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE Or res.Status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE

    'テキスト変換するか判断して、返却
    If GetBinaryMode Then GetSyncMessage = tmp Else GetSyncMessage = Split(CharConv.BytesToString(tmp), vbNullChar)(0)
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
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle
    
    'エラーコードの初期化
    ResultCode = 0
    
    '既に受信予約済みの場合は、即抜け
    If isDataReady Then ViewLog.LogWarn "既に受信予約済みです。", "WebSocketCommunicator.GetAsyncMessage": Exit Function

    '非同期処理側で既に、受信済みデータがあるか？
    If Not (isReceiving) Then
        'ない場合は、初期化
        InitializeBuffer G_res.Buffer, G_res.CurrentPointer, G_res.BufferLength
        Set G_res.collect = New Collection

        '非同期モードの場合は、こっちも追加で初期化しておく
        G_res.Status = WINHTTP_WEB_SOCKET_Incoming_Requests_BUFFER_TYPE
        G_res.ReceiveBytes = 0
    End If

    'データが尽きるまで、レスポンス受信ループ
    Dim tmp() As Byte, InsiderRes As WebSocketReceiveManage
    Do
        '非同期処理側で既に、受信済みデータがある場合は、処理を飛ばします
        '※次の受信データがある状態で、この処理が走ると、コールバック側のプロシージャへ遷移し、グローバル変数側に、メッセージが格納されます。
        If Not (isReceiving) Then
            G_res.result = WinHttpWebSocketReceive( _
                            UseWebsockethandle, G_res.Buffer(G_res.CurrentPointer), G_res.BufferLength, _
                            G_res.ReceiveBytes, G_res.Status)

            'エラーなら即抜け
            '※呼び出し側で、エラーかの確認をしてください
            If G_res.result Then ResultCode = G_res.result: Exit Function
        End If
    
        'グローバル変数側と仲介しながら、バッファー管理処理へ移る
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
    
    'コールバック側で受け取ったグローバル変数で、クラス内変数へコピーする
    With tmp
        '固定配列なので、愚直にループでコピー
        For i = LBound(.Buffer) To UBound(.Buffer)
            .Buffer(i) = G_res.Buffer(i)
        Next
        
        .BufferLength = G_res.BufferLength
        Set .collect = G_res.collect
        .CurrentPointer = G_res.CurrentPointer
        .ReceiveBytes = G_res.ReceiveBytes
        .result = G_res.result
        .Status = G_res.Status
    End With
    
    'コピーしたクラス内変数で、バッファー管理へ移し、結果を返却
    IntermediaryProcessing_BufferManagement = BufferManagement(tmp)
    
    'メッセージ処理したので、受信済みフラグを消す
    isReceiving = False

    'バッファー処理した結果を、グローバル変数側へ反映する
    With G_res
        '固定配列なので、愚直にループで転記
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
    ViewLog.LogDebug "------------ WebSocket 受信ステータス ------------", TraceName
    ViewLog.LogDebug "バッファー残量 ：" & Param.BufferLength & " byte", TraceName
    ViewLog.LogDebug "受信データ量　 ：" & Param.ReceiveBytes & " byte", TraceName
    ViewLog.LogDebug "バッファー位置 ：" & Param.CurrentPointer & " byte", TraceName
    ViewLog.LogDebug "WebSocketStatus：" & Param.Status, TraceName
    ViewLog.LogDebug "---------------------------------------------------", TraceName
    '============================================================================


    'WINHTTP_WEB_SOCKET_STATUS.eBufferType による分岐処理
    '※4 + 1種類のタイプが出ますが、4種類の部分はどっちもバイナリ処理のため、統一します。
    Select Case Param.Status
        '※非同期用に特別に用意したコードです
        Case WINHTTP_WEB_SOCKET_Incoming_Requests_BUFFER_TYPE
            isDataReady = True
            ViewLog.LogInfo "まだ、受信データがありません。受信予約を行いました。", TraceName
        
        Case WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE
            ViewLog.LogInfo "全ての" & IIf(Param.Status = WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, "バイナリ", "UTF-8") & "データが届きました。", TraceName
            
            '1. 最後のパケットデータを追加
            Param.collect.Add Param.Buffer
            ViewLog.LogInfo "最後のバッファーを一時Collectionキューに追加しました。", TraceName

            '2. 全てのチャンクを結合して、返却
            BufferManagement = MakeResponseFullBytes(Param.collect)
            ViewLog.LogInfo "チャンクを全て結合し、バイト配列として返却しました。", TraceName

            '3. 全てを受信し終えたステータスなので、初期化処理
            InitializeBuffer Param.Buffer, Param.CurrentPointer, Param.BufferLength
            Set Param.collect = New Collection
            ViewLog.LogInfo "バッファーを確保しました。", TraceName
        
        Case WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE, WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE
            '1. バッファー残量確認
            ViewLog.LogWarn "一部の" & IIf(Param.Status = WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE, "バイナリ", "UTF-8") & "データを受け取りました。完全なデータの受信には再度、受信リクエストを送る必要があります。", TraceName
            If Param.BufferLength <= 0 Then
                ViewLog.LogInfo "バッファーを使い果たしました。", TraceName
                
                '2. 満たしたパケットをコレクションへ追加
                Param.collect.Add Param.Buffer
                ViewLog.LogInfo "満たしたバッファーを一時Collectionキューに追加しました。", TraceName

                '3. バッファー初期化して、空きを確保
                InitializeBuffer Param.Buffer, Param.CurrentPointer, Param.BufferLength
                ViewLog.LogInfo "バッファーを確保しました。", TraceName
            Else
                ViewLog.LogInfo "バッファーに空きがあるため、次の受信処理に持ち越します。", TraceName
            End If

        'サーバーが閉じるフレームを送信しました。
        Case WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE
            ViewLog.LogInfo "切断リクエストが来たため、WebSocket 処理を終了します...", TraceName
            
            '具体的な処理はまだ未実装
        
        '`WINHTTP_WEB_SOCKET_STATUS.eBufferType`未定義のコード
        Case Else
            ViewLog.LogWarn "`WINHTTP_WEB_SOCKET_STATUS.eBufferType`未定義のコードが来てます：" & Param.Status, TraceName
    End Select
End Function
