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
'                                  ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
'WebSokect関連API
Private Declare PtrSafe Function WinHttpWebSocketSend Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal eBufferType As Long, _
   ByVal pvBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long

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

Private Declare PtrSafe Function WinHttpWebSocketQueryCloseStatus Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef usStatus As Integer, _
   ByRef pvReason As Any, _
   ByVal dwReasonLength As Long, _
   ByRef pdwReasonLengthConsumed As Long _
   ) As Long


'`lpvStatusInformation`の中身を受け取るためのAPI
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)



'***************************************************************************************************
'                     ■■■ コールバック受信用 WindowsAPI・構造体定義 ■■■
'***************************************************************************************************
'https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ns-winhttp-winhttp_web_socket_status
Private Type WINHTTP_WEB_SOCKET_STATUS
    dwBytesTransferred As Long
    eBufferType        As Long
End Type

'Websocket蓄積受信状況把握に使用
Private Const BufferToAllocate      As Long = 2 ^ 14 '初期バッファー上限
Private Type WebSocketReceiveManage
    Buffer(BufferToAllocate - 1)    As Byte         '第1引数        `WinHttpWebSocketReceive`からの受信バイトデータ
    BufferLength                    As Long         '第2引数        バッファー残量　※`UpdateBufferInfo`内で計算します
    ReceiveBytes                    As Long         '第3引数        `WinHttpWebSocketReceive`から受信できたサイズ
    Status                          As Long         '第4引数        WINHTTP_WEB_SOCKET_STATUS.eBufferType
    CurrentPointer                  As Long         '第5引数        カーソル位置　※`UpdateBufferInfo`内で計算します
    collect                         As Collection   'チャンク収集   バラバラのデータを蓄積させる用
End Type
Private m_res   As WebSocketReceiveManage



'***************************************************************************************************
'                                  ■■■ 定数宣言 ■■■
'***************************************************************************************************
'ログトレース用
Private Const ThisClassName As String = "WebSocketCommunicator"

'受信ステータス
'https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ne-winhttp-winhttp_web_socket_buffer_type
Private Enum WINHTTP_WEB_SOCKET_BUFFER_TYPE
    BINARY_MESSAGE      'バイナリメッセージ全体か最後の部分。
    BINARY_FRAGMENT     'バイナリメッセージの一部分のみ｡
    UTF8_MESSAGE        'UTF-8メッセージ全体か最後の部分。
    UTF8_FRAGMENT       'UTF-8 メッセージの一部分のみ。
    CLOSE_              'サーバーが閉じるフレームを送信しました｡
End Enum

'主要なエラーコード
'※一旦は、送受信関連のみ。不整合な処理を防ぐ際の辻褄のあうエラー番号を用意してます
Private Enum WINHTTP_WEB_SOCKET_ERROR
    INVALID_OPERATION = 4317                '今の状態でその命令は無理
    WINHTTP_OPERATION_CANCELLED = 12017     '操作が取り消された
End Enum

'主要なWebSocket関連コールバックステータス
'https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/nc-winhttp-winhttp_status_callback
Private Enum WINHTTP_STATUS_CALLBACK
    READ_COMPLETE = 524288          'サーバーからデータが正常に読み取られました。   `WinHttpWebSocketReceive` 使用後、呼び出される合図
    WRITE_COMPLETE = 1048576        'データがサーバーに正常に送信できた。           `WinHttpWebSocketSend`    使用後、呼び出される合図
    CLOSE_COMPLETE = 33554432       '接続が正常に閉じられました。                   `WinHttpWebSocketClose`   使用後、呼び出される合図
    SHUTDOWN_COMPLETE = 67108864    '接続が正常にシャットダウンされました。         `WinHttpWebSocketShutdown`使用後、呼び出される合図
End Enum

'WebSokect関連のステータス
Private Const WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS = 1000

'その他
Private Const TIMEOUT_DEFAULT As Single = 30                '初期タイムアウト秒数



'***************************************************************************************************
'                                  ■■■ グローバル変数 ■■■
'***************************************************************************************************
'内部使用用オブジェクト
Private WithEvents Http             As WebSocketHTTPCommunicator    'SafeTimer の WithEvents st As SafeTimer 相当。WebSocketハンドルもここに格納されてる
Attribute Http.VB_VarHelpID = -1
Private CharConv                    As New CharacterCodeConversion  'バイト配列←→UTF-8変換用
Private ViewLog                     As New Logger                   'ログ用

'--- WebSokect関連の補助ステータス ---
Private Type StateWebSocket
    isReceiving         As Boolean  '前回の受信データが残っている場合は、`True`となり、プロシージャ`RequestWebSocketReceive`は、実行できなくなります。
    isReserved          As Boolean  'プロシージャ`GetAsyncMessage`を呼ぶと同時に`True`になり`、WinHttpWebSocketReceive`が実行された後、`READ_COMPLETE`が来れば`False`、来なければ`True`のままになります
    isSended            As Boolean  '`WinHttpWebSocketSend`が実行された後、正常に送信が行われると、`True`になります。プロシージャ`SendAsyncMessage`が呼ばれる度に、`False`にリセットされます
End Type
Private currentebWebSocket As StateWebSocket

'--- ログ設定状況 ---
Private Type StateLog
    id                  As String   'ユニーク名
    doLog               As Boolean  'ログファイルにも、デバッグメッセージを残すか
    logPath             As String   'ログファイルの保存フォルダパス
End Type
Private currentLog As StateLog

'最後のWebSocketから得た受信データ
Private Content_WinHttpWebSocketReceive() As Byte

'プロシージャ`GetAsyncMessage`にて、次のデータを受信する際のタイムアウト値
Private timerOut   As Long



'***************************************************************************************************
'                                ■■■ check用プロパティ ■■■
'***************************************************************************************************
'* 機能    ：前回の`WinHttpWebSocketSend`がうまくいったかを確認します
'---------------------------------------------------------------------------------------------------
'* 注意事項：初回は、`Flase`とします
'***************************************************************************************************
Public Property Get LastSendSuccess() As Boolean
    LastSendSuccess = currentebWebSocket.isSended
End Property

'***************************************************************************************************
'* 機能    ：`WinHttpReceiveResponse`による受信待機中かを確認します
'***************************************************************************************************
Public Property Get isWaitingReceiveResponse() As Boolean
    isWaitingReceiveResponse = currentebWebSocket.isReserved
End Property

'***************************************************************************************************
'* 機能    ：`WinHttpReceiveResponse`による受信データが届いたかを確認します
'***************************************************************************************************
Public Property Get LastReceiveExisting() As Boolean
    LastReceiveExisting = currentebWebSocket.isReceiving
End Property



'***************************************************************************************************
'                          ■■■ 受信データ取得用プロパティ ■■■
'***************************************************************************************************
'* 機能    ：`WinHttpReceiveResponse`による最後の受信データをバイト配列として、取得します
'---------------------------------------------------------------------------------------------------
'* 注意事項：最後の1件分しか保持しません
'***************************************************************************************************
Public Property Get LastReceiveContentByte() As Byte
    LastReceiveContentByte = Content_WinHttpWebSocketReceive
End Property

'***************************************************************************************************
'* 機能    ：`WinHttpReceiveResponse`による最後の受信データを文字列として、取得します
'---------------------------------------------------------------------------------------------------
'* 注意事項：最後の1件分しか保持しません
'***************************************************************************************************
Public Property Get LastReceiveContentString() As String
    LastReceiveContentString = CharConv.BytesToString(Content_WinHttpWebSocketReceive)
End Property



'***************************************************************************************************
'                                  ■■■ ログヘルパー ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
'----------------------------------------------------------------------------------
' A custom printMsg that besides printing to the Immediate Window, it will
' also print to a dated log file. This is useful at user level where
' the Immediate Window is not easily accessible. doRaiseError will also raise an
' error aside from printing the error message to the log file.
' Updated: 07/01/26 Long Vh: the message output is disabled by default to avoid
'                            spamming info. This can be turned on by setting globally
'                            noDbgMsg = False or noLogging = False
'----------------------------------------------------------------------------------

    Dim logName As String, strFormattedMsg As String

    If isHeader Then strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-") _
    Else strFormattedMsg = " | " & currentLog.id & " | " & strMsg

   'Perform msg logging if needed
    If currentLog.doLog Then
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".WebSokect.log"
        If Dir(currentLog.logPath, vbDirectory) = "" Then MkDir currentLog.logPath
    End If

   'Perform msg output to Immediate Window if needed
    With ViewLog
        Select Case LogLevel_
            Case LogLevelName.Trace_: .LogTrace strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.Debug_: .LogDebug strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.info_: .LogInfo strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.WARN_: .LogWarn strFormattedMsg, From, currentLog.logPath, logName
            Case LogLevelName.ERROR_: .LogError strFormattedMsg, From, currentLog.logPath, logName, Err.Number
            Case Else: 'ログ出力無効化
        End Select
    End With


    'エラー時に止まります
    If doRaiseError Then Err.Raise vbObjectError + 1, From, Description:=strMsg

End Sub



'***************************************************************************************************
'                                 ■■■ 数値から文字列化 ■■■
'***************************************************************************************************
'* 機能    ：WINHTTP_WEB_SOCKET_BUFFER_TYPE - eBufferType 一式
'---------------------------------------------------------------------------------------------------
'* 詳細説明：基本的には、ログ用です
'---------------------------------------------------------------------------------------------------
'* URL　　 ：https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ne-winhttp-winhttp_web_socket_buffer_type
'***************************************************************************************************
Public Property Get Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE(TypeID As Long, ErrorSource As String) As String
    Select Case TypeID
        Case WINHTTP_WEB_SOCKET_BUFFER_TYPE.BINARY_MESSAGE:     Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "BINARY_MESSAGE_BUFFER_TYPE"
        Case WINHTTP_WEB_SOCKET_BUFFER_TYPE.BINARY_FRAGMENT:    Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "BINARY_FRAGMENT_BUFFER_TYPE"
        Case WINHTTP_WEB_SOCKET_BUFFER_TYPE.UTF8_MESSAGE:       Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "UTF8_MESSAGE_BUFFER_TYPE"
        Case WINHTTP_WEB_SOCKET_BUFFER_TYPE.UTF8_FRAGMENT:      Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "UTF8_FRAGMENT_BUFFER_TYPE"
        Case WINHTTP_WEB_SOCKET_BUFFER_TYPE.CLOSE_:             Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE = "CLOSE_BUFFER_TYPE"
        Case Else: Err.Raise 4 + vbObjectError, ErrorSource, "未知の WINHTTP_WEB_SOCKET_BUFFER_TYPE が来ました。" & vbCrLf & "Code：" & TypeID
    End Select
End Property

'***************************************************************************************************
'* 機能    ：WINHTTP_STATUS_CALLBACK - dwInternetStatus 一式
'---------------------------------------------------------------------------------------------------
'* 詳細説明：基本的には、ログ用です
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
        Case WINHTTP_STATUS_CALLBACK.READ_COMPLETE: Name__WINHTTP_STATUS_CALLBACK = "READ_COMPLETE"
        Case WINHTTP_STATUS_CALLBACK.WRITE_COMPLETE: Name__WINHTTP_STATUS_CALLBACK = "WRITE_COMPLETE"
'        Case 2097152: Name__WINHTTP_STATUS_CALLBACK = "REQUEST_ERROR"
'        Case 4194304: Name__WINHTTP_STATUS_CALLBACK = "SENDREQUEST_COMPLETE"
'        Case 16777216: Name__WINHTTP_STATUS_CALLBACK = "GETPRO_YFORURL_COMPLETE"
        Case WINHTTP_STATUS_CALLBACK.CLOSE_COMPLETE: Name__WINHTTP_STATUS_CALLBACK = "CLOSE_COMPLETE"
        Case WINHTTP_STATUS_CALLBACK.SHUTDOWN_COMPLETE: Name__WINHTTP_STATUS_CALLBACK = "SHUTDOWN_COMPLETE"
'        Case 268435456: Name__WINHTTP_STATUS_CALLBACK = "SETTINGS_WRITE_COMPLETE"
'        Case 536870912: Name__WINHTTP_STATUS_CALLBACK = "SETTINGS_READ_COMPLETE"
'        Case Else: Err.Raise 3 + vbObjectError, ErrorSource, "未知の WINHTTP_CALLBACK_STATUS が来ました。" & vbCrLf & "Code：" & dwInternetStatus
        Case Else: Name__WINHTTP_STATUS_CALLBACK = "Unknow_Code:" & dwInternetStatus
    End Select
End Property



'***************************************************************************************************
'                               ■■■ 主要WebSocketメソッド ■■■
'***************************************************************************************************
'* 機能　　：WebSocket 接続を確立し、コールバックを登録します
'---------------------------------------------------------------------------------------------------
'* 引数    ：Host       接続先ホスト名
'            Path       接続先パス
'            port       ws=80、wss=443、カスタムも可（省略時=443）
'            SecureMode wss なら True（省略時=True）
'---------------------------------------------------------------------------------------------------
'* 機能説明：常に非同期モードで動作します。コールバックは Http_OnCallback で自動受信されます。
'* 注意事項：・SecureMode が正しくないと失敗します
'            ・Cookie や認証ヘッダー等は未対応です（匿名アクセス用）
'            ・接続を維持するには、グローバル変数 or Static変数 として、このクラスオブジェクトを定義してください。
'***************************************************************************************************
Public Sub connectionWebSocket(Host As String, Path As String, Optional port As Long = 443, Optional SecureMode As Boolean = True)
    Const FromProcedureName As String = ThisClassName & ".connectionWebSocket"


    '1. 最初はhttp(s)で接続確立
    Dim FullURI As String: FullURI = Host & Path & ":" & port
    printMsg Debug_, FullURI & "　に接続します...", FromProcedureName
    Set Http = New WebSocketHTTPCommunicator
    If Http.Connect(Host, Path, port, SecureMode) Then
        printMsg info_, IIf(SecureMode, "https://", "http://") & FullURI & "　への接続に成功しました。WebSocketへの Upgrade 処理を行います...", FromProcedureName
    Else
        printMsg ERROR_, IIf(SecureMode, "https://", "http://") & FullURI & "　への接続に失敗しました。", FromProcedureName, doRaiseError:=True
        Exit Sub
    End If

    '2. WebSocketへUpgrade
    printMsg Debug_, IIf(SecureMode, "wss://", "ws://") & FullURI & "　に Upgrade します...", FromProcedureName
    If Http.UpgradeToWebSocket() Then
        printMsg info_, IIf(SecureMode, "wss://", "ws://") & FullURI & "　に Upgrade、成功しました。", FromProcedureName
    Else
        printMsg ERROR_, IIf(SecureMode, "wss://", "ws://") & FullURI & "　に Upgrade、失敗しました。", FromProcedureName, doRaiseError:=True
        Exit Sub
    End If
End Sub

'***************************************************************************************************
'* 機能　　：受け取った引数のデータで、WebSocketへ送信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：SendData           String OR Byte 型を想定
'* 返り値  ：実行結果コード     ※成功時は、0です
'---------------------------------------------------------------------------------------------------
'* 機能説明：型判定を行い、テキストもWideCharToMultiByte を介したバイト配列変換を使用します
'* 注意事項：・テキストも所詮、バイナリデータなのでこのような方式にしています
'            ・必要に応じて、`LastSendStatus`にて、最後の送信がうまくいったかを確認してください
'            ・前回の送信がうまくいったかの確認はこのプロシージャでは行いません
'            ・`WinHttpWebSocketSend`の実行結果コードは基本的に0です。
'              送信がうまくいったは、`LastSendStatus`にて、確認できます。
'***************************************************************************************************
Public Function SendAsyncMessage(SendData As Variant) As Long
    Const FromProcedureName As String = ThisClassName & ".SendMessage"


    '0. ステータスをリセット
    currentebWebSocket.isSended = False

    '1. 必要に応じて、バイト配列として変換する
    Dim UseData() As Byte
    Dim UseType As Long
    Select Case VarType(SendData)
        Case vbString
            UseType = WINHTTP_WEB_SOCKET_BUFFER_TYPE.UTF8_MESSAGE
            UseData = CharConv.BytesFromString(CStr(SendData))   'テキストをバイト配列に変換する
            printMsg Debug_, "テキストデータをUTF-8バイト配列に変換して送信します...", FromProcedureName

        Case vbByte
            UseType = WINHTTP_WEB_SOCKET_BUFFER_TYPE.BINARY_MESSAGE
            UseData = SendData                                     'そのまんま使用
            printMsg Debug_, "バイナリデータとして送信します...", FromProcedureName

        Case Else
            Err.Raise vbObjectError + 1, "WebSocketCommunicator.SendMessage", "この操作は、String型かByte型のみ可能です。" & vbCrLf & "数値のみで渡してる場合は事前に、`CStr`を介す必要があります。"

    End Select

    '2. バイト配列のサイズを取得
    Dim messageByteSize As Long: messageByteSize = UBound(UseData) + 1

    '3. 求めたサイズで一括送信し、実行結果コードを返却
    SendAsyncMessage = WinHttpWebSocketSend(Http.websockethandle, _
                        UseType, _
                        VarPtr(UseData(0)), messageByteSize)
End Function

'***************************************************************************************************
'* 機能　　：WebSocketに対して、受信リクエストを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：WinHttpWebSocketReceive の実行結果。
'---------------------------------------------------------------------------------------------------
'* 機能説明：非同期モードでの、`WinHttpWebSocketReceive`は、予約扱いとして機能します
'* 注意事項：・あくまでも「予約」まで担います。実際のデータは、`GetAsyncMessage`にて、取り出し処理してください
'            ・以下の状態の場合は、`WinHttpWebSocketReceive`を呼び出しません
'            　・既に予約済み
'            　・前回の`WinHttpWebSocketReceive`で得たデータが残っている
'***************************************************************************************************
Public Function RequestWebSocketReceive() As Long
    Const FromProcedureName As String = ThisClassName & ".RequestWebSocketReceive"


    '1. いくつかフラグを確認し、受信予約可能かを確認します
    '1-1. 既に予約済み(受信待機)か？
    If currentebWebSocket.isReserved Then
        printMsg WARN_, "既に、`WinHttpWebSocketReceive`による受信予約がされています", FromProcedureName
        RequestWebSocketReceive = WINHTTP_WEB_SOCKET_ERROR.WINHTTP_OPERATION_CANCELLED  '既に予約してるので、取り消しサインとしてのエラー番号を返却
        Exit Function

    '1-2. 前回の受信データが残っているか？
    ElseIf currentebWebSocket.isReceiving Then
        printMsg WARN_, "前回の`WinHttpWebSocketReceive`で取得したデータがまだ残っています。`GetAsyncMessage`を使用してデータを取得し、再度実行してください。", FromProcedureName
        RequestWebSocketReceive = WINHTTP_WEB_SOCKET_ERROR.INVALID_OPERATION            '前回のデータがあるのに、また予約しようとして、不整合の基になるため、「そんな状態で予約は無理だよ」サインを返す
        Exit Function

    '1-3. 全てのチェックをパス出来たら、受信予約としてフラグ付け
    Else
        printMsg Debug_, "`WinHttpWebSocketReceive`で、受信予約を行います...", FromProcedureName
        currentebWebSocket.isReserved = True

    End If

    '2. このクラスオブジェクトの変数に格納するように受信リクエストし、実行結果を返却
    With m_res
        RequestWebSocketReceive = WinHttpWebSocketReceive( _
                                Http.websockethandle, .Buffer(m_res.CurrentPointer), .BufferLength, _
                                .ReceiveBytes, .Status)
    End With
End Function

'***************************************************************************************************
'* 機能　　：WebSocketからの非同期受信を行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：`WinHttpWebSocketReceive`の最終実行結果
'---------------------------------------------------------------------------------------------------
'* 機能説明：全てのデータを受け切るまでリクエストを垂れ流していきます。
'* 注意事項：・途中でエラーが発生した場合は、即抜けとなり、途中のデータは破棄されます
'            ・完成したデータは、プロパティ`LastReceiveContentString`で別途取得してください
'***************************************************************************************************
Public Function GetAsyncMessage() As Long
    Const FromProcedureName As String = ThisClassName & ".GetAsyncMessage"


    '1. いくつかフラグを確認し、データ受信可能かを確認します
    '1-1. まだ、予約(受信待機)中か？
    If currentebWebSocket.isReserved Then
        printMsg WARN_, "現在、受信データを待機しています", FromProcedureName
        GetAsyncMessage = WINHTTP_WEB_SOCKET_ERROR.WINHTTP_OPERATION_CANCELLED   '既に予約してるので、取り消しサインとしてのエラー番号を返却
        Exit Function

    '1-2. データが届いてないか？
    ElseIf Not currentebWebSocket.isReceiving Then
        printMsg WARN_, "受信データは、ありません。受信予約が必要です。", FromProcedureName
        GetAsyncMessage = WINHTTP_WEB_SOCKET_ERROR.INVALID_OPERATION             '前回のデータがないのに、取り出そうとして、不整合の基になるため、「そんな状態で取り出しは無理だよ」サインを返す
        Exit Function

    '1-3. 全てのチェックをパス出来たら、次の処理へ
    Else
        printMsg Debug_, "データが受信されたことを確認できました。処理に入ります。", FromProcedureName
    End If

    '2. データが尽きるまで、レスポンス受信ループ
    Do
        '2-1. バッファー管理処理へ移り、現時点での`m_res`にあるデータの取得を試みる
        BufferManagement

        '2-2. 断片データ状態の場合は、再予約を行い、断片データも収集できるようにする
        If m_res.Status = WINHTTP_WEB_SOCKET_BUFFER_TYPE.BINARY_FRAGMENT Or m_res.Status = WINHTTP_WEB_SOCKET_BUFFER_TYPE.UTF8_FRAGMENT Then
            '2-2-1. コールバック(ウィンドウメッセージ)によるイベント発火を試みる
            printMsg info_, "`FRAGMENT`を検知しました。次のデータが届くまで待機します...", FromProcedureName
            GetAsyncMessage = RequestWebSocketReceive

            '2-2-2. エラーなら、即抜け
            If GetAsyncMessage Then Exit Function

            '2-2-3. 次のデータが届くまでLoop　※タイムアウトしたら、中断されます
            Dim timerStart As Single
            timerStart = Timer
            Do Until currentebWebSocket.isReceiving
                printMsg Trace_, "Waiting for the next data...", FromProcedureName
                If Timer - timerStart > timerOut Then printMsg ERROR_, "Timeout waiting for the WebSocket to load the next fragment data.", FromProcedureName, doRaiseError:=True
                DoEvents
            Loop
        End If

        DoEvents

    Loop While currentebWebSocket.isReceiving
End Function

'***************************************************************************************************
'* 機能    ：WinHttp コールバックイベントハンドラ（SafeTimer の st_TimerCall 相当）
'***************************************************************************************************
Private Sub Http_OnCallback(ByVal HINTERNET As LongPtr, _
                             ByVal dwInternetStatus As Long, _
                             ByVal lpvStatusInformation As LongPtr, _
                             ByVal dwStatusInformationLength As Long)
    Const FromProcedureName As String = ThisClassName & ".Http_OnCallback"


    'WebSocket 関連コールバックのみ処理する
    Select Case dwInternetStatus
        Case CLOSE_COMPLETE, READ_COMPLETE, SHUTDOWN_COMPLETE, WRITE_COMPLETE

            'WINHTTP_WEB_SOCKET_STATUS のポインタを基にコピー
            Dim WebSocketStatus As WINHTTP_WEB_SOCKET_STATUS
            If lpvStatusInformation <> 0 And dwStatusInformationLength >= LenB(WebSocketStatus) Then
                CopyMemory WebSocketStatus, ByVal lpvStatusInformation, LenB(WebSocketStatus)
            Else
                'CLOSE/SHUTDOWN 等では statusInformation が来ないケースがある
                WebSocketStatus.dwBytesTransferred = 0
                WebSocketStatus.eBufferType = -1
                printMsg WARN_, "statusInformation が Null/不足のため既定値で処理します。status=" & dwInternetStatus & ", len=" & dwStatusInformationLength, FromProcedureName
            End If

            printMsg Debug_, "------------ WINHTTP_WEB_SOCKET_STATUS ------------", FromProcedureName
            printMsg Debug_, "Bytes：" & WebSocketStatus.dwBytesTransferred, FromProcedureName
            printMsg Debug_, "Type ：" & Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE(WebSocketStatus.eBufferType, FromProcedureName) & "(" & WebSocketStatus.eBufferType & ")", FromProcedureName
            printMsg Debug_, "---------------------------------------------------", FromProcedureName
            printMsg Debug_, "WINHTTP_STATUS_CALLBACK：" & Name__WINHTTP_STATUS_CALLBACK(dwInternetStatus, FromProcedureName) & "(" & dwInternetStatus & ")", FromProcedureName

            Select Case dwInternetStatus
                Case READ_COMPLETE
                    '受信バッファ管理情報の更新は READ_COMPLETE のときだけ行う。
                    'WRITE_COMPLETE 等でここを上書きすると、巨大レスポンス分割受信中に m_res が壊れる。
                    m_res.Status = WebSocketStatus.eBufferType
                    m_res.ReceiveBytes = WebSocketStatus.dwBytesTransferred

                    'ステータス更新
                    currentebWebSocket.isReceiving = True    'データが届いたので、データが残存してるフラグを`True`にする
                    currentebWebSocket.isReserved = False    'データは届いたので、受信予約フラグを`False`にする

                    'ログ化
                    printMsg Debug_, "非同期処理により、受信メッセージを格納しました。", FromProcedureName

                Case WRITE_COMPLETE
                    'ステータス更新
                    currentebWebSocket.isSended = True       '送信の確認が取れたので、そのフラグを`True`にする

                    'ログ化
                    printMsg Debug_, "非同期処理により、送信の確認が取れました。", FromProcedureName

                Case SHUTDOWN_COMPLETE
                    printMsg info_, "WebSocket の接続が正常にシャットダウンされました。", FromProcedureName

                Case CLOSE_COMPLETE
                    printMsg info_, "WebSocket を閉じました。", FromProcedureName

            End Select

        Case Else
            printMsg WARN_, "WebSocket 関連以外のコールバックが来たようです。WINHTTP_STATUS_CALLBACK：" & dwInternetStatus, FromProcedureName

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
    Dim result As Long
    result = WinHttpWebSocketClose(Http.websockethandle, _
                        WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS, 0, 0)
    If result And Not (ForceClose) Then GoTo quit

    ' サーバーから返されたクローズステータスを確認する。
    Dim CloseReasonBuffer(123) As Byte
    Dim BufferSize As Long: BufferSize = UBound(CloseReasonBuffer) + 1
    Dim Status As Integer
    Dim CloseReasonLength As Long
    
    result = WinHttpWebSocketQueryCloseStatus(Http.websockethandle, Status, _
                        CloseReasonBuffer(0), BufferSize, CloseReasonLength)
    If result And Not (ForceClose) Then GoTo quit
    If Status <> WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS And Not (ForceClose) Then GoTo quit
    
    'ハンドルを破棄
    Set Http = New WebSocketHTTPCommunicator
    Http.CloseHWebsocketHandle

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
'* 機能　　：チャンクごとに区切られたコレクションバイト配列を一気にバイト配列変換します
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
    
    If totalSize <= 0 Then
        Dim emptyBytes() As Byte
        MakeResponseFullBytes = emptyBytes
        Exit Function
    End If

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
'* 機能　　：先頭から指定サイズ分だけ切り出したバイト配列を返します
'***************************************************************************************************
Private Function SliceBytes(ByRef Source() As Byte, ByVal dataSize As Long) As Byte()
    Dim clipped() As Byte
    Dim i As Long

    If dataSize <= 0 Then
        SliceBytes = clipped
        Exit Function
    End If

    ReDim clipped(0 To dataSize - 1)
    For i = 0 To dataSize - 1
        clipped(i) = Source(i)
    Next
    SliceBytes = clipped
End Function

'***************************************************************************************************
'* 機能　　：WinHttpWebSocketReceive のバッファー管理を担います
'---------------------------------------------------------------------------------------------------
'* 機能説明：`WebSocketReceiveManage`型のクラス内変数の情報を駆使して、良い感じに、WebSocketからの受信データを蓄積させます
'***************************************************************************************************
Private Sub BufferManagement()
    'ログ把握用
    Const FromProcedureName As String = ThisClassName & ".BufferManagement"


    '残存データを処理したので、フラグをリセット
    currentebWebSocket.isReceiving = False

    'バッファー情報の更新
    UpdateBufferInfo m_res.CurrentPointer, m_res.BufferLength, m_res.ReceiveBytes


    '========================= 受信ステータス値　把握用 =========================
    printMsg Debug_, "------------ WebSocket 受信ステータス ------------", FromProcedureName
    printMsg Debug_, "バッファー残量 ：" & m_res.BufferLength & " byte", FromProcedureName
    printMsg Debug_, "受信データ量　 ：" & m_res.ReceiveBytes & " byte", FromProcedureName
    printMsg Debug_, "バッファー位置 ：" & m_res.CurrentPointer & " byte", FromProcedureName
    printMsg Debug_, "WebSocketStatus：" & m_res.Status, FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    '============================================================================


    'WINHTTP_WEB_SOCKET_STATUS.eBufferType による分岐処理
    '※4種類 + 1 のタイプが出ますが、処理の仕方は2種類のため、統一します。
    Select Case m_res.Status
        Case WINHTTP_WEB_SOCKET_BUFFER_TYPE.BINARY_MESSAGE, WINHTTP_WEB_SOCKET_BUFFER_TYPE.UTF8_MESSAGE
            '1. 今のCollectionに、既に貯まっている状態か？
            If m_res.collect.Count Then
                printMsg Debug_, "全ての" & IIf(m_res.Status = WINHTTP_WEB_SOCKET_BUFFER_TYPE.BINARY_MESSAGE, "バイナリ", "UTF-8") & "データが届きました。", FromProcedureName

                '1-1. 最後のパケットデータを追加（未使用領域を含めない）
                m_res.collect.Add SliceBytes(m_res.Buffer, m_res.CurrentPointer)
                printMsg Debug_, "最後のバッファーを一時Collectionキューに追加しました。", FromProcedureName

                '1-2. 全てのチャンクを結合して、完成データ用変数に適用
                Content_WinHttpWebSocketReceive = MakeResponseFullBytes(m_res.collect)
                printMsg Debug_, "チャンクを全て結合し、バイト配列として返却しました。", FromProcedureName

                '1-3. 次のためにCollectionを空にする
                Set m_res.collect = New Collection
                printMsg Debug_, "新規Collectionを確保しました。", FromProcedureName

            Else
                printMsg Debug_, "1回のバッファーで収まる" & IIf(m_res.Status = WINHTTP_WEB_SOCKET_BUFFER_TYPE.BINARY_MESSAGE, "バイナリ", "UTF-8") & "データが届きました。", FromProcedureName

                '1-1. 最後のパケットデータを直接返却（未使用領域を含めない）
                Content_WinHttpWebSocketReceive = SliceBytes(m_res.Buffer, m_res.CurrentPointer)
            End If

                '2. 全てを受信し終えたステータスなので、初期化処理
                InitializeBuffer m_res.Buffer, m_res.CurrentPointer, m_res.BufferLength
                printMsg Debug_, "バッファーを確保しました。", FromProcedureName

        Case WINHTTP_WEB_SOCKET_BUFFER_TYPE.BINARY_FRAGMENT, WINHTTP_WEB_SOCKET_BUFFER_TYPE.UTF8_FRAGMENT
            '1. バッファー残量確認
            printMsg info_, "一部の" & IIf(m_res.Status = WINHTTP_WEB_SOCKET_BUFFER_TYPE.BINARY_FRAGMENT, "バイナリ", "UTF-8") & "データを受け取りました。完全なデータの受信には再度、受信リクエストを送る必要があります。", FromProcedureName
            If m_res.BufferLength <= 0 Then
                printMsg Debug_, "バッファーを使い果たしました。", FromProcedureName

                '2. 満たしたパケットをコレクションへ直追加
                m_res.collect.Add m_res.Buffer
                printMsg Debug_, "満たしたバッファーを、Collectionキューに追加しました。", FromProcedureName

                '3. バッファー初期化して、空きを確保
                InitializeBuffer m_res.Buffer, m_res.CurrentPointer, m_res.BufferLength
                printMsg Debug_, "バッファーを確保しました。", FromProcedureName
            Else
                printMsg Debug_, "まだバッファーに空きがあるため、次の受信処理に持ち越します。", FromProcedureName
            End If

        'サーバーが閉じるフレームを送信しました。
        Case WINHTTP_WEB_SOCKET_BUFFER_TYPE.CLOSE_
            printMsg Debug_, "切断リクエストが来たため、WebSocket 処理を終了します...", FromProcedureName
            
            '具体的な処理はまだ未実装
        
        '`WINHTTP_WEB_SOCKET_STATUS.eBufferType`未定義のコード
        Case Else
            printMsg WARN_, "`WINHTTP_WEB_SOCKET_STATUS.eBufferType`未定義のコードが来てます：" & m_res.Status, FromProcedureName
    End Select
End Sub



Private Sub Class_Initialize()
    Const FromProcedureName As String = "WebSocketCommunicator.Class_Initialize"


    'Assign an objID for easy debug identification
    currentLog.id = "BRID" & Format(Rnd * 1000, "000")

    '------------------------- クラス内変数の一部を初期化 -------------------------
    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        '1. パスが設定されてたら、ログファイル保存を有効化
        Dim LogFolderPath As String: LogFolderPath = .Range(.UseRangeName(9, FromProcedureName)).value
        If LogFolderPath <> "" Then
            currentLog.doLog = True                'ログファイルとして保存するか？
            currentLog.logPath = LogFolderPath     '保存先のフォルダパスを設定
        End If
    End With

    'タイムアウト設定(デフォルトで、30s)
    timerOut = TIMEOUT_DEFAULT

    'バッファー情報を初期化
    InitializeBuffer m_res.Buffer, m_res.CurrentPointer, m_res.BufferLength
    Set m_res.collect = New Collection
End Sub
