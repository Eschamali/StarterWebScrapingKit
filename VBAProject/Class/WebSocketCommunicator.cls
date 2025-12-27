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

'WebSocket に関する定義
Private Http As WebSocketHTTPCommunicator
Private ExistingWebsockethandle_ As LongPtr



'***************************************************************************************************
'                                 ■■■ 設定/取得メソッド ■■■
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
'* 機能　　：WebSocketへ送信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：message            送信プレーンテキスト
'* 返り値  ：実行結果コード     ※成功時は、0です
'---------------------------------------------------------------------------------------------------
'* 機能説明：ErrorCode：12152　対策により、ADODB を介したバイト配列変換を使用してます
'* 注意事項：・現時点は、プレーンテキスト送信のみ対応です
'            ・非同期の場合は、コールバック側での確認も必要です
'***************************************************************************************************
Public Function SendMessage(message As String) As Long
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle

    'テキストメッセージをUTF-8バイト配列に変換
    'message = StrConv(message, vbFromUnicode)
    Dim ary() As Byte: ary = ShSetting02_StartWebSocket.StringToUtf8Bytes(message)     'Task：ここに直接バイナリデータを入れることで、画像とかもいけるかも？
    
    '変換したバイト配列のサイズを取得
    Dim messageByteSize As Long: messageByteSize = UBound(ary) + 1
    
    '求めたサイズを基に一括送信
    Dim result As Long
    result = WinHttpWebSocketSend(UseWebsockethandle, _
                        WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                        VarPtr(ary(0)), messageByteSize)


    '実行結果コードを返却
    SendMessage = result
End Function

'***************************************************************************************************
'* 機能　　：WebSocketからの同期受信を行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：WebSocketからの受信メッセージ。※エラーだと、vbnullstring です
'---------------------------------------------------------------------------------------------------
'* 機能説明：ErrorCode：12152　対策により、ADODB を介したバイト配列変換を使用してます
'* 注意事項：・受信データがない状態でこれを呼び出すと、受信データが来るまで、フ　リ　ー　ズ　します。ネット切断することで、40s前後でフリーズ回避できます。
'            ・現時点は、プレーンテキスト受信のみ対応です
'            ・送信と違って受信は内部で、固定のバッファーサイズ受信によるストリーミング受信方式を取っています。そのため、受信データ量が多いと返るまで時間がかかります。
'***************************************************************************************************
Public Function GetMessageForSync() As String
    'ログ関連定義
    Dim Error翻訳 As New WinApiError
    Dim LogView As New Logger
    Const ErrorSource As String = "WebSocketCommunicator.GetMessageForSync"

    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle
    
    'ストリーミング受信準備
    Dim res As WebSocketReceiveManage
    ShSetting02_StartWebSocket.InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength    'バッファー初期化
    Set res.collect = New Collection                                                                '一時蓄積データも初期化

    'データが尽きるまで、レスポンス受信ループ
    Dim tmp
    Do
        '受信処理
        res.result = WinHttpWebSocketReceive( _
                        UseWebsockethandle, res.Buffer(res.CurrentPointer), res.BufferLength, _
                        res.ReceiveBytes, res.Status)
        
        'エラーならGoToへ
        If res.result <> NO_ERROR Then GoTo quit
        LogView.LogInfo Error翻訳.GetMessage(res.result, "winhttp"), ErrorSource

        'バッファー情報の更新
        ShSetting02_StartWebSocket.UpdateBufferInfo res.CurrentPointer, res.BufferLength, res.ReceiveBytes
 
        '受信処理を託す
        tmp = ShSetting02_StartWebSocket.CommonWinHttpWebSocketReceive(res)
        
        '全ての受信を終えたときの処理
        If res.Status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then
            '完成品のプレーンテキストを返却
            GetMessageForSync = CStr(tmp)

            Exit Function
        ElseIf res.Status = WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE Then
            '完成品のバイナリデータを返却
            
            
            Exit Function
        End If
    Loop


'エラー時は、エラーメッセージを表示
quit:
LogView.LogError Error翻訳.GetMessage(res.result, "winhttp"), ErrorSource, res.result
End Function

'***************************************************************************************************
'* 機能　　：WebSocketからの受信予約を行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：エラーコード   ※成功時は、0です
'---------------------------------------------------------------------------------------------------
'* 機能説明：・非同期モードではありますが実際は、受信予約扱いです。それまでにメッセージがあれば、コールバック側でメッセージの確認が取れます
'            ・受信データが空でもフリーズせず、受信予約的な挙動となります
'
'* 注意事項：・現時点は、プレーンテキスト受信のみ対応です
'            ・受信データが空かつ既に予約済みの場合に更にこのメソッドを呼ぶと、4317エラーが返るようです
'            ・これを呼び出すと、全メッセージを受信するのではなく、呼び出す度に1件1件受信される仕組みです
'***************************************************************************************************
Public Function GetMessageForAsync() As Long
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle

    '受信予約処理
    G_res.result = WinHttpWebSocketReceive( _
                    UseWebsockethandle, G_res.Buffer(G_res.CurrentPointer), G_res.BufferLength, _
                    G_res.ReceiveBytes, G_res.Status)
                    

    '実行結果コードを返却
    GetMessageForAsync = G_res.result
End Function

'***************************************************************************************************
'* 機能　　：Websocket通信を終了させます
'***************************************************************************************************
Public Function CloseWebSocket() As Boolean
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle

    Dim result As Long
    result = WinHttpWebSocketClose(UseWebsockethandle, _
                        WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS, 0, 0)
    If result <> NO_ERROR Then GoTo quit

    ' サーバーから返されたクローズステータスを確認する。
    Dim CloseReasonBuffer(123) As Byte
    Dim BufferSize As Long: BufferSize = UBound(CloseReasonBuffer) + 1
    Dim Status As Integer
    Dim CloseReasonLength As Long
    
    result = WinHttpWebSocketQueryCloseStatus(UseWebsockethandle, Status, _
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
