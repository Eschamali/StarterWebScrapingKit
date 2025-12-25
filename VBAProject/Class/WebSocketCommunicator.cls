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

'Websocket受信に使用(自作構造体)
Private Type response
    Buffer(4095) As Byte    '第1引数
    BufferLength As Long    '第2引数
    ReceiveBytes As Long    '第3引数
    Status As Long          '第4引数
    CurrentPointer As Long  '第5引数
    result As Long          '戻り値
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
'* 機能　　：既存のwebsocketのハンドル値を設定します
'---------------------------------------------------------------------------------------------------
'* LET引数 ：ExistingWebsockethandle    以前取得したWebsockethandle値
'---------------------------------------------------------------------------------------------------
'* 機能説明：以前に確立したwebsocketのハンドル値を渡すことで続きのやり取りができます。
'* 注意事項：ErrorCode:12030 が来た場合は、もう切断状態です。「接続処理（Connect）からやり直す（再接続）」 しかありません。
'            別プロセスを跨ぐ使用("/x excel.exe" による別プロセス起動など)はできません...
'***************************************************************************************************
Public Property Let ReConnect(ExistingWebsockethandle As LongPtr)
    ExistingWebsockethandle_ = ExistingWebsockethandle
End Property

'***************************************************************************************************
'* 機能　　：WebSocketへ送信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：message            送信プレーンテキスト
'            ReConnectionHandle 再接続時のWebSocketハンドル値
'
'* 返り値  ：エラーコード   ※成功時は、0です
'---------------------------------------------------------------------------------------------------
'* 機能説明：ErrorCode：12152　対策により、ADODB を介したバイト配列変換を使用してます
'* 注意事項：現時点は、プレーンテキスト送信のみ対応です
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
    
    'エラーチェック
    If result = NO_ERROR Then
        '成功なら、0とする
        SendMessage = NO_ERROR
    Else
        'エラーなら､エラーコードで返す
        Debug.Print "送信エラー発生。ErrorCode：" & result & ",Description：" & GetErrorMes.GetMessage(result, "winhttp")
        SendMessage = result
    End If
End Function

'***************************************************************************************************
'* 機能　　：WebSocketからの同期受信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：ReConnectionHandle 再接続時のWebSocketハンドル値
'* 返り値  ：WebSocketからの受信メッセージ。※エラーだと、vbnullstring です
'---------------------------------------------------------------------------------------------------
'* 機能説明：ErrorCode：12152　対策により、ADODB を介したバイト配列変換を使用してます
'* 注意事項：・受信データがない状態でこれを呼び出すと、受信データが来るまで、フ　リ　ー　ズ　します。ネット切断することで、40s前後でフリーズ回避できます。
'            ・現時点は、プレーンテキスト受信のみ対応です
'            ・送信と違って受信は内部で、固定のバッファーサイズ受信によるストリーミング受信方式を取っています。そのため、受信データ量が多いと返るまで時間がかかります。
'***************************************************************************************************
Public Function GetMessageForSync() As String
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle
    
    'ストリーミング受信準備
    Dim res As response
    ShSetting02_StartWebSocket.InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength   'バッファー初期化
    Set res.collect = New Collection                                    '結合データも初期化

    'データが尽きるまで、レスポンス受信ループ
    Do
        '受信処理
        res.result = WinHttpWebSocketReceive( _
                        UseWebsockethandle, res.Buffer(res.CurrentPointer), res.BufferLength, _
                        res.ReceiveBytes, res.Status)
        
        'エラーならGoToへ
        If res.result <> NO_ERROR Then GoTo quit

        'メッセージを全て受信し終えたら、各パケットを結合
        If res.Status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then
            '最後のパケットを追加
            res.collect.Add res.Buffer
            
            '全てのパケットを結合して、プレーンテキストに変換して、返却
            GetMessageForSync = ShSetting02_StartWebSocket.MakeResponseMessage(res.collect)
            Exit Function
        End If
        
        'バッファー情報の更新
        ShSetting02_StartWebSocket.UpdateBufferInfo res.CurrentPointer, res.BufferLength, res.ReceiveBytes
        
        'バッファー残量判定
        If (res.BufferLength <= 0) Then
            '満たしたパケットをコレクションへ追加
            res.collect.Add res.Buffer
            
            'バッファー初期化
            ShSetting02_StartWebSocket.InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength
        End If
    Loop


'エラー時は、エラーメッセージを表示
quit:
Debug.Print "受信エラー発生。ErrorCode：" & res.result & ",Description：" & GetErrorMes.GetMessage(res.result, "winhttp")
End Function

'***************************************************************************************************
'* 機能　　：WebSocketからの受信予約を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：ReConnectionHandle 再接続時のWebSocketハンドル値
'* 返り値  ：WebSocketからの受信メッセージ。※エラーだと、vbnullstring です
'---------------------------------------------------------------------------------------------------
'* 機能説明：非同期モードではありますが実際は、受信予約扱いです。それまでにメッセージがあれば、コールバック側でメッセージの確認が取れます
'* 注意事項：現時点は、プレーンテキスト受信のみ対応です
'***************************************************************************************************
Public Function GetMessageForAsync() As String
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle
    
    'ストリーミング受信準備による初期化
    InitializeBuffer G_res.Buffer, G_res.CurrentPointer, G_res.BufferLength
    Set G_res.collect = New Collection

    'データが尽きるまで、レスポンス受信ループ
    Do
        '受信処理
    G_res.result = WinHttpWebSocketReceive( _
                    UseWebsockethandle, G_res.Buffer(G_res.CurrentPointer), G_res.BufferLength, _
                    G_res.ReceiveBytes, G_res.Status)
        
        'エラーならGoToへ
        If G_res.result <> NO_ERROR Then GoTo quit

        'メッセージを全て受信し終えたら、各パケットを結合
        If G_res.Status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then
            '最後のパケットを追加
            G_res.collect.Add G_res.Buffer
            
            '全てのパケットを結合して、プレーンテキストに変換して、返却
            GetMessage = MakeResponseMessage(G_res.collect)
            Exit Function
        End If
        
        'バッファー情報の更新
        UpdateBufferInfo G_res.CurrentPointer, G_res.BufferLength, G_res.ReceiveBytes
        
        'バッファー残量判定
        If (G_res.BufferLength <= 0) Then
            '満たしたパケットをコレクションへ追加
            G_res.collect.Add G_res.Buffer
            
            'バッファー初期化
            InitializeBuffer G_res.Buffer, G_res.CurrentPointer, G_res.BufferLength
        End If
    Loop


'エラー時は、エラーメッセージを表示
quit:
Debug.Print "受信エラー発生。ErrorCode：" & G_res.result & ",Description：" & GetErrorMes.GetMessage(G_res.result, "winhttp")
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
