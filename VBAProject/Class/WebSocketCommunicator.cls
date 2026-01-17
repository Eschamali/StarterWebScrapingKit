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
    Buffer(4095) As Byte    '第1引数
    BufferLength As Long    '第2引数
    ReceiveBytes As Long    '第3引数
    Status As Long          '第4引数
    CurrentPointer As Long  '第5引数
    result As Long          '戻り値
    collect As Collection
End Type

Public Http As WebSocketHTTPCommunicator
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
'* 機能説明：ErrorCode：12152　対策により、WideCharToMultiByte を介したバイト配列変換を使用してます
'* 注意事項：現時点は、プレーンテキスト送信のみ対応です
'***************************************************************************************************
Public Function SendMessage(message As String) As Long
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle

    'テキストメッセージをUTF-8バイト配列に変換
    Dim CharConv As New CharacterCodeConversion
    Dim ary() As Byte: ary = CharConv.BytesFromString(message)  'Task：ここに直接バイナリデータを入れることで、画像とかもいけるかも？
    'message = StrConv(message, vbFromUnicode)
    
    '変換したバイト配列のサイズを取得
    Dim messageByteSize As Long: messageByteSize = UBound(ary) + 1

    '求めたサイズで一括送信し、実行結果コードを返却
    '※呼び出し側で、エラーかの確認をしてください
    SendMessage = WinHttpWebSocketSend(UseWebsockethandle, _
                        WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                        VarPtr(ary(0)), messageByteSize)
End Function

'***************************************************************************************************
'* 機能　　：WebSocketからの受信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：ResultCode     WinHttpWebSocketReceive の実行結果。実質返り値です。
'* 返り値  ：WebSocketからの受信メッセージ。※エラーだと、vbnullstring です
'---------------------------------------------------------------------------------------------------
'* 機能説明：ErrorCode：12152　対策により、WideCharToMultiByte を介したバイト配列変換を使用してます
'* 注意事項：・受信データがない状態でこれを呼び出すと、受信データが来るまで、フ　リ　ー　ズ　します。ネット切断することで、40s前後でフリーズ回避できます。
'            ・現時点は、プレーンテキスト受信のみ対応です
'            ・送信と違って受信は内部で、固定のバッファーサイズ受信によるストリーミング受信方式を取っています。そのため、受信データ量が多いと返るまで時間がかかります。
'***************************************************************************************************
Public Function GetMessage(Optional ByRef ResultCode As Long) As String
    '再接続のハンドル情報があったらそっちを使用する
    Dim UseWebsockethandle As LongPtr
    If ExistingWebsockethandle_ Then UseWebsockethandle = ExistingWebsockethandle_ Else UseWebsockethandle = Http.websockethandle
    
    'ストリーミング受信準備
    Dim res As response
    InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength
    Set res.collect = New Collection

    'データが尽きるまで、レスポンス受信ループ
    Do
        '受信処理
        res.result = WinHttpWebSocketReceive( _
                        UseWebsockethandle, res.Buffer(res.CurrentPointer), res.BufferLength, _
                        res.ReceiveBytes, res.Status)
        
        'エラーなら即抜け
        '※呼び出し側で、エラーかの確認をしてください
        If res.result Then ResultCode = res.result: Exit Function

        'メッセージを全て受信し終えたら、各パケットを結合
        If res.Status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then
            '最後のパケットを追加
            res.collect.Add res.Buffer
            
            '全てのパケットを結合して、プレーンテキストに変換して、返却
            GetMessage = MakeResponseMessage(res.collect)
            Exit Function
        End If
        
        'バッファー情報の更新
        UpdateBufferInfo res.CurrentPointer, res.BufferLength, res.ReceiveBytes
        
        'バッファー残量判定
        If (res.BufferLength <= 0) Then
            '満たしたパケットをコレクションへ追加
            res.collect.Add res.Buffer
            
            'バッファー初期化
            InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength
        End If
    Loop
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
    If result Then GoTo quit

    ' サーバーから返されたクローズステータスを確認する。
    Dim CloseReasonBuffer(123) As Byte
    Dim BufferSize As Long: BufferSize = UBound(CloseReasonBuffer) + 1
    Dim Status As Integer
    Dim CloseReasonLength As Long
    
    result = WinHttpWebSocketQueryCloseStatus(UseWebsockethandle, Status, _
                        CloseReasonBuffer(0), BufferSize, CloseReasonLength)
    If result Then GoTo quit
    If Status <> WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS Then GoTo quit

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
Private Sub InitializeBuffer(responseBuffer As Variant, CurrentPointer As Long, BufferLength As Long)
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
'* 返り値  ：プレーンテキスト「UTF-8文字列(String)」
'---------------------------------------------------------------------------------------------------
'* 注意事項：当然ですが、各コレクションバイナリデータが、UTF-8バイト配列になっていないと、文字化けします
'***************************************************************************************************
Private Function MakeResponseMessage(collect As Collection) As String
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
    
    ' 3. バイトデータをコピーして結合       Task：バイナリデータそのものが欲しいならここで取り出す
    currentPos = 0
    For Each chunk In collect
        For i = 0 To UBound(chunk)
            fullBytes(currentPos) = chunk(i)
            currentPos = currentPos + 1
        Next
    Next

    ' 4. ここでUTF-8変換！（さっきの関数を使う）
    ' ※末尾のヌル文字(vbNullChar)除去も行う
    Dim CharConv As New CharacterCodeConversion
    MakeResponseMessage = Split(CharConv.BytesToString(fullBytes), vbNullChar)(0)
End Function
