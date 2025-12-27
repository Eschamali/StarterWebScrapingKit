VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShSetting02_StartWebSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'***************************************************************************************************
'                   WebSocket にまつわる共通処理をまとめたものです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                      ■■■ コールバック関数からの受け取り変数の定義 ■■■
'***************************************************************************************************
Private FromCallBackTmpReceiveData As Collection



'***************************************************************************************************
'                      ■■■ このシートのみ定義されてるセル名称定義 ■■■
'***************************************************************************************************
Property Get UseRangeName(id As Long, ErrorSource As String) As String
    Select Case id
        Case 1: UseRangeName = "WebSocketハンドル"
        Case 2: UseRangeName = "IsWebSocket"
        Case 3: UseRangeName = "ホスト名"
        Case 4: UseRangeName = "Port"
        Case 5: UseRangeName = "セキュアモード"
        Case 6: UseRangeName = "UriPath"
        Case 7: UseRangeName = "送信プレーンテキスト"
        Case Else: Err.Raise 1 + vbObjectError, ErrorSource, "渡されたセルIDは、定義されてません"
    End Select
End Property



'***************************************************************************************************
'                   ■■■ WINHTTP_STATUS_CALLBACK - dwInternetStatus 一式 ■■■
'---------------------------------------------------------------------------------------------------
'* 注意事項：Websocketとは関係ないコールバック定数は一旦、コメントアウトしています
'---------------------------------------------------------------------------------------------------
'* URL　　 ：https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/nc-winhttp-winhttp_status_callback
'***************************************************************************************************
Property Get Name__WINHTTP_STATUS_CALLBACK(dwInternetStatus As Long, ErrorSource As String) As String
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
        Case Else: Name__WINHTTP_STATUS_CALLBACK = "Unknow_Code"
    End Select
End Property



'***************************************************************************************************
'                               ■■■ 受け取り変数の制御設定 ■■■
'***************************************************************************************************
'* 機能　　：コールバック関数からの受信データを一旦受け取ったり、取得、追加処理を行います
'---------------------------------------------------------------------------------------------------
'* GET返値 ：現在の受信コレクションデータ
'* LET引数 ：コールバック関数からの受信データ。一旦、型指定はしません。
'---------------------------------------------------------------------------------------------------
'* 注意事項：これを使用する前に次項の New による初期化処理が必要です。
'***************************************************************************************************
Property Get ReceiveBox() As Collection
    Set ReceiveBox = FromCallBackTmpReceiveData
End Property
Property Let ReceiveBox(Data)
    FromCallBackTmpReceiveData.Add Data
End Property

'***************************************************************************************************
'* 機能　　：受信メッセージを全て削除し、初期化します
'***************************************************************************************************
Public Sub CleanReceiveBox()
    Set FromCallBackTmpReceiveData = New Collection
End Sub



'***************************************************************************************************
'                   ■■■ WINHTTP_WEB_SOCKET_BUFFER_TYPE - eBufferType 一式 ■■■
'---------------------------------------------------------------------------------------------------
'* URL　　 ：https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ne-winhttp-winhttp_web_socket_buffer_type
'***************************************************************************************************
Property Get Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE(TypeID As Long, ErrorSource As String) As String
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
'                 ■■■ 受信処理を円滑に行うためのヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：WinHttpWebSocketReceive の受信処理を賄います
'---------------------------------------------------------------------------------------------------
' *返り値  ：Param.Status(WINHTTP_WEB_SOCKET_BUFFER_TYPE)が、
'            ・0 ：完成品のバイナリデータ
'            ・2 ：完成品のプレーンテキストデータ
'            ・他：Empty
'
'* 引数    ：ユーザー定義型-WebSocketReceiveManage
'---------------------------------------------------------------------------------------------------
'* 機能説明：WebSocket バッファーの種類が「WINHTTP_STATUS_CALLBACK_READ_COMPLETE」での、チャンク処理をいい感じに共通化させます
'* 注意事項：・このヘルパープロシージャーを呼び出す前に、バッファー情報の更新を済ませること
'            ・Param.Status が、0 or 2  の場合は、呼び出す側で、`WebSocketReceiveManage.collect`のクリーンをすること
'***************************************************************************************************
Public Function CommonWinHttpWebSocketReceive(ByRef Param As WebSocketReceiveManage) As Variant
    'ログ把握用クラス
    Dim ViewLog As New Logger
    Const TraceName As String = "ShSetting02_StartWebSocket.CommonWinHttpWebSocketReceive"

    
    '========================= 受信ステータス値　把握用 =========================
    ViewLog.LogDebug "------------ WebSocket 受信ステータス ------------", TraceName
    ViewLog.LogDebug "バッファー残量 ：" & Param.BufferLength & " byte", TraceName
    ViewLog.LogDebug "受信データ量　 ：" & Param.ReceiveBytes & " byte", TraceName
    ViewLog.LogDebug "バッファー位置 ：" & Param.CurrentPointer & " byte", TraceName
    ViewLog.LogDebug "WebSocketStatus：" & Param.Status, TraceName
    ViewLog.LogDebug "---------------------------------------------------", TraceName
    '============================================================================

    
    Select Case Param.Status
        'BINARY_MESSAGE_BUFFER_TYPE：Buffer には、バイナリ メッセージ全体またはその最後の部分が含まれます。
        Case 0
            ViewLog.LogWarn "全てのバイナリデータが届きましたが、組み立てロジックはまだ未完成です...", TraceName

            '2. 受信し終えたステータスなので、初期化処理
            Me.InitializeBuffer Param.Buffer, Param.CurrentPointer, Param.BufferLength  'バッファー初期化
            ViewLog.LogInfo "バッファーを確保しました。", TraceName

        'UTF8_MESSAGE_BUFFER_TYPE  ：Buffer には、UTF-8    メッセージ全体またはその最後の部分が含まれます。
        Case 2
            ViewLog.LogInfo "全てのプレーンテキストデータが届きました。", TraceName
            
            '1. 最後のパケットデータを追加
            Param.collect.Add Param.Buffer
            ViewLog.LogInfo "最後のバッファーを一時Collectionキューに追加しました。", TraceName

            '2. 全てのパケットを結合して、プレーンテキストに変換し、返却
            CommonWinHttpWebSocketReceive = ShSetting02_StartWebSocket.MakeResponseMessage(Param.collect)
            ViewLog.LogInfo "パケットを全て結合し、プレーンテキストとして返却しました。", TraceName

            '3. 全てを受信し終えたステータスなので、初期化処理
            ShSetting02_StartWebSocket.InitializeBuffer Param.Buffer, Param.CurrentPointer, Param.BufferLength  'バッファー初期化
            ViewLog.LogInfo "バッファーを確保しました。", TraceName
        
        '1=BINARY_FRAGMENT_BUFFER_TYPE：Buffer にはバイナリ メッセージの一部のみが含まれます。
        '3=UTF8_FRAGMENT_BUFFER_TYPE  ：Buffer には UTF-8   メッセージの一部のみが含まれます。
        Case 1, 3
            '1. バッファー残量確認
            ViewLog.LogWarn "一部のデータを受け取りました。完全なデータの受信には再度、受信リクエストを送る必要があります。", TraceName
            If Param.BufferLength <= 0 Then
                ViewLog.LogInfo "バッファーを使い果たしました。", TraceName
                
                '2. 満たしたパケットをコレクションへ追加
                Param.collect.Add Param.Buffer
                ViewLog.LogInfo "満たしたバッファーを一時Collectionキューに追加しました。", TraceName

                '3. バッファー初期化して、空きを確保
                ShSetting02_StartWebSocket.InitializeBuffer Param.Buffer, Param.CurrentPointer, Param.BufferLength
                ViewLog.LogInfo "バッファーを確保しました。", TraceName
            Else
                ViewLog.LogInfo "バッファーに空きがあるため、次の受信処理に持ち越します。", TraceName
            End If
        
        'CLOSE_BUFFER_TYPE：サーバーが閉じるフレームを送信しました。
        Case 4
            ViewLog.LogInfo "切断リクエストが来たため、WebSocket 処理を終了します...", TraceName
            
            'WebSocketを閉じる
            
            
    End Select
End Function

'***************************************************************************************************
'* 機能　　：バッファ初期化
'***************************************************************************************************
Public Sub InitializeBuffer(responseBuffer() As Byte, CurrentPointer As Long, BufferLength As Long)
    '中身を消す("0"で埋め尽くす)
    Erase responseBuffer
    
    '書き込み位置を先頭にする
    CurrentPointer = 0
    
    '`responseBuffer`の確保サイズを取得します
    BufferLength = UBound(responseBuffer) + 1
End Sub

'***************************************************************************************************
'* 機能　　：バッファ管理情報更新
'***************************************************************************************************
Public Sub UpdateBufferInfo(CurrentPointer As Long, BufferLength As Long, ReceiveBytes As Long)
    '受信したバイト数分、書き込み位置をずらす
    CurrentPointer = CurrentPointer + ReceiveBytes

    '`responseBuffer`の空きサイズを更新
    BufferLength = BufferLength - ReceiveBytes
End Sub

'***************************************************************************************************
'* 機能　　：プレーンテキスト「文字列(String)」で送る際、UTF-8バイト配列(Byte())に変換します
'---------------------------------------------------------------------------------------------------
'* 引数    ：str        プレーンテキスト「文字列(String)」
'* 返り値  ：UTF-8バイト配列
'---------------------------------------------------------------------------------------------------
'* 機能説明：ErrorCode：12152　対策
'***************************************************************************************************
Public Function StringToUtf8Bytes(ByVal str As String) As Byte()
    '空文字なら即抜け
    If str = "" Then Exit Function

    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText str
        
        ' ストリームの先頭に戻ってバイト型に変更
        .Position = 0
        .Type = 1 ' adTypeBinary
        
        ' BOM（先頭3バイト）をスキップして読む
        .Position = 3
        StringToUtf8Bytes = .Read
        .Close
    End With
End Function

'***************************************************************************************************
'* 機能　　：バイナリデータをプレーンテキスト「UTF-8文字列(String)」に変換します
'---------------------------------------------------------------------------------------------------
'* 引数    ：bytes  バイナリデータ
'* 返り値  ：プレーンテキスト「UTF-8文字列(String)」
'---------------------------------------------------------------------------------------------------
'* 注意事項：当然ですが、バイナリデータが、UTF-8バイト配列になっていないと、文字化けします
'***************************************************************************************************
Public Function Utf8BytesToString(ByRef Bytes() As Byte) As String
    With CreateObject("ADODB.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .Write Bytes
        
        ' 先頭に戻ってテキストモードに変更
        .Position = 0
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        
        ' 文字列として読み出す
        Utf8BytesToString = .ReadText
        
        .Close
    End With
End Function

'***************************************************************************************************
'* 機能　　：チャンクごとに区切られたコレクションバイト配列を一気にUTF-8変換します
'---------------------------------------------------------------------------------------------------
'* 引数    ：collect  コレクションバイト配列
'* 返り値  ：プレーンテキスト「UTF-8文字列(String)」
'***************************************************************************************************
Public Function MakeResponseMessage(collect As Collection) As String
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
        Next i
    Next chunk
    
    ' 4. ここでUTF-8変換！（さっきの関数を使う）
    ' ※末尾のヌル文字(vbNullChar)除去は、UTF-8変換後にやるか、
    '   変換前にサイズ調整するかですが、ADODBはヌル文字もそのまま変換してくれるので
    '   変換後にReplaceかLeftで切るのが安全です。
    
    Dim resultStr As String
    resultStr = Utf8BytesToString(fullBytes)
    
    ' 5. ヌル文字除去（もしあれば）
    Dim nullPos As Long
    nullPos = InStr(resultStr, vbNullChar)
    If nullPos > 0 Then
        MakeResponseMessage = Left(resultStr, nullPos - 1)
    Else
        MakeResponseMessage = resultStr
    End If
End Function
