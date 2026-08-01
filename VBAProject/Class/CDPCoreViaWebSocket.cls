VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPCoreViaWebSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   `WebSocket`経由での、`Chrome DevTools Protocol`のやり取りに特化した専用Classモジュールです。
'                                  基本的には、ローカル専用です。
'   もし、LANを超えたネットを介した制御もしたい場合は、「edge://inspect/#devices」や、SSH、VPN(Tailscale等)側で設定して下さい。
'
'                    `CDPCore.cls`とやり取りしやすいように、工夫して組んでます。
'　基本的には、このClassオブジェクトの使いまわしはしない想定で組んでます。`deserialize`で疑似的な使い回しが可能なので、そちらをご利用ください
'
'
'====================== なんで、「WinHttpWebSocket○○」系のAPIを使わないの？ ======================
'   「WinHttpWebSocket○○」の場合、`ioctlsocket`や`PeekNamedPipe`と言った「覗き見機能」はなく、下記の2択しか選択肢がありません
'   ・OSの推奨する「非同期コールバック」
'     VBE（開発画面）で一時停止したり、デバッグでマクロを途中で止めたりした瞬間に、Excelごと警告なしでクラッシュしてしまいます。
'     現時点では､手動リセット用プロシージャをわざわざ実行するしかなく､忘れた途端､クラッシュして作業を失う可能性があります｡
'     AIの力を持ってしてもまだ完全には解消出来てません｡
'
'   ・同期モード
'     Peekがないから、データがまだ届いていないときに「頂戴！」と1行呼んだ瞬間、
'     Excelが「応答なし（白い画面）」になって完全にフリーズ（ブロッキング）して、泣く泣くタスクマネージャーで落とします。
'
'
'   悩みに悩んだ結果、「WinSock」でほぼ1から作り上げました。
'   これにより､Pipe版と同じロジックとして運用することが可能になりました｡
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
' --- WinSock 2 API 宣言 ---
#If VBA7 Then
    Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Integer, ByRef lpWSAData As WSAData) As Long
    Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare PtrSafe Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal socktype As Long, ByVal Protocol As Long) As LongPtr
    Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As LongPtr) As Long
    Private Declare PtrSafe Function Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As LongPtr, ByRef Name As sockaddr_in, ByVal namelen As Long) As Long
    Private Declare PtrSafe Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As LongPtr, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
    Private Declare PtrSafe Function recv Lib "ws2_32.dll" (ByVal s As LongPtr, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
    Private Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare PtrSafe Function ioctlsocket Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare PtrSafe Function setsockopt Lib "ws2_32.dll" (ByVal s As LongPtr, ByVal Level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Long) As Long
#Else
    Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Integer, ByRef lpWSAData As WSAData) As Long
    Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
    Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal socktype As Long, ByVal protocol As Long) As Long
    Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
    Private Declare Function connect Lib "ws2_32.dll" (ByVal sID As Long, ByRef Name As sockaddr_in, ByVal namelen As Long) As Long
    Private Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
    Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
    Private Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Long) As Long
#End If



'***************************************************************************************************
'                                   ■■■ 各種構造体 ■■■
'***************************************************************************************************
' --- WinSock 2 構造体定義 ---
#If VBA7 And Win64 Then
    Private Type WSAData
        wVersion As Integer
        wHighVersion As Integer
        szDescription(0 To 256) As Byte
        szSystemStatus(0 To 128) As Byte
        iMaxSockets As Integer
        iMaxUdpDg As Integer
        lpVendorInfo As LongPtr
    End Type
#Else
    Private Type WSAData
        wVersion As Integer
        wHighVersion As Integer
        szDescription(0 To 256) As Byte
        szSystemStatus(0 To 128) As Byte
        iMaxSockets As Integer
        iMaxUdpDg As Integer
        lpVendorInfo As Long
    End Type
#End If

Private Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

'WebSocketのFrame単位を管理する独自構造体
Private Type WebSocketReceiveFrameManage
    PayloadLength   As Long     'Chromium側から実際に送られるサイズ
    CurrentCursor   As Long     '`PayloadLength`のうち、どこまで書き込んだ？
    ValidTags       As Boolean  '有効なタグ中か？ペイロードデータ自体の欠け対策です
    isFIN           As Boolean  'FINフラグ
    opcode          As Long     '「これはテキストだよ（1）」「バイナリだよ（2）」「切断要求だよ（8）」などのデータ種類
    FrameHeaderSize As Long     '今回のフレームヘッダーサイズ
End Type
Private m_WebSocketFrame As WebSocketReceiveFrameManage



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'WINHTTP_WEB_SOCKET_BUFFER_TYPE
'https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ne-winhttp-winhttp_web_socket_buffer_type
Public Enum WS_BUFFER_TYPE
    WS_NOT_MESSAGE = -3       ' 一度も受信してません            ※一度も`GetAsyncMessage`を呼び出してない初期値
    WS_PAYLOAD_FRAGMENT = -2  ' ペイロード自体が欠けてる場合    ※フレームヘッダー情報に載ってるペイロードデータサイズが、WinSockPipeに入ってない理不尽パターン
    WS_PAYLOAD_EMPTY = -1     ' ペイロード自体がない　          ※フレームヘッダー情報は届いたのに、ペイロードデータが届いてない理不尽パターン
    WS_BINARY_MESSAGE = 0     ' WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE
    WS_BINARY_FRAGMENT = 1    ' WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE
    WS_UTF8_MESSAGE = 2       ' WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE
    WS_UTF8_FRAGMENT = 3      ' WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE
    WS_CLOSE_MESSAGE = 4      ' WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE
End Enum

'主要な、winsock定数
Private Const AF_INET       As Long = 2             ' 【住所の形式】IPv4アドレス形式（「127.0.0.1」などのお馴染みのIPアドレス形式）を使用する宣言
Private Const SOCK_STREAM   As Long = 1             ' 【データの送り方】データの順序と到着が確実に保証される「コネクション型（ストリーム）」を使用する宣言
Private Const IPPROTO_TCP   As Long = 6             ' 【通信プロトコル】世界標準の信頼性の高い「TCP/IPプロトコル」を使用する宣言
Private Const FIONREAD      As Long = &H4004667F    ' 受信可能バイト数
Private Const SOL_SOCKET    As Long = 65535         ' ソケットレベルのオプション
Private Const SO_RCVBUF     As Long = &H1002&       ' 受信バッファサイズ設定フラグ
Private Const MSG_WAITALL   As Long = &H8           ' 要求したバイト数のデータがすべて揃うか、エラーや接続終了が発生するまで関数の処理をブロック（待機）する上級者向け。いわゆる、Pipe版と同等の挙動

'デフォルト値
Private Const D_port        As Long = 9222
Private Const D_host        As String = "127.0.0.1"

'その他
Private Const ThisClassName As String = "CDPCoreViaWebSocket"
Private Const InitialBuffer As Long = 2 ^ 20        ' WinSock/ADODB.Stream 初期バッファー上限



'***************************************************************************************************
'                                   ■■■ イベント定義 ■■■
'***************************************************************************************************
'WebSocket受信発信機
Public Event WebSocketReceive(res() As Byte, resSize As Long, EOS As Boolean)



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'--- 状態管理 ---
#If VBA7 Then
    Private m_hSocket As LongPtr    'WinSockハンドル
#Else
    Private m_hSocket As Long       'WinSockハンドル
#End If
Private m_WSAInitialized    As Boolean          'WinSockの初期化成功可否
Private m_IsConnected       As Boolean          '既にWebSocketに接続済みか？
Private m_LastReceiveContentBINARY() As Byte    '最後のWebSocketから得た完成受信データ(バイト配列)
Private m_LastBufferType    As WS_BUFFER_TYPE   '最後のWebSocketから得た受信データステータス

'--- ログ設定状況 ---
Private Type StateLog
    id                  As String   'ユニーク名
    doLog               As Boolean  'ログファイルにも、デバッグメッセージを残すか
    logPath             As String   'ログファイルの保存フォルダパス
End Type
Private currentLog As StateLog

'--- Pipe版に合わせたエラー有無判定 ---
Private StyleErrPeekNamedPipe   As Long  'データ存在有無(ioctlsocket)の確認でエラーになったか？
Private StyleErrReadFile        As Long  'データを取り出してる最中(recv)にエラーになったか？

'--- 各種クラスオブジェクト ---
Private WebSocketReceiveStream  As Object                       'WebSocketバイトストリームを蓄積するオブジェクト
Private Utf8Converter           As New CharacterCodeConversion  'UTF-8バイト配列に変換するオブジェクト
Private ViewLog                 As New Logger                   'ログ用



'***************************************************************************************************
'                                ■■■ check用プロパティ ■■■
'***************************************************************************************************
'* 機能    ：最後に受信したパケットのバッファタイプを確認します（WinHTTP仕様に準拠）
'---------------------------------------------------------------------------------------------------
'* 注意事項：初回は、`-3`とします。`GetAsyncMessage`を呼ばないと更新されません
'***************************************************************************************************
Public Property Get LastBufferType() As WS_BUFFER_TYPE
    LastBufferType = m_LastBufferType
End Property

'***************************************************************************************************
'* 機能    ：このClassオブジェクトが既に、WebSocketによる接続中かを確認します
'***************************************************************************************************
Public Property Get Connected() As Boolean
    Connected = m_IsConnected
End Property

'***************************************************************************************************
'* 機能    ：`GetAsyncMessage`プロシージャ使用時、`ioctlsocket`でエラーが起こった際のエラーコード
'***************************************************************************************************
Public Property Get IsErrorStylePeekNamedPipe() As Long
    IsErrorStylePeekNamedPipe = StyleErrPeekNamedPipe
End Property

'***************************************************************************************************
'* 機能    ：`GetAsyncMessage`プロシージャ使用時、`recv`でエラーが起こった際のエラーコード
'***************************************************************************************************
Public Property Get IsErrorStyleReadFile() As Long
    IsErrorStyleReadFile = StyleErrReadFile
End Property



'***************************************************************************************************
'                          ■■■ 受信データ取得用プロパティ ■■■
'***************************************************************************************************
'* 機能    ：最後の受信データをバイト配列として、取得します
'---------------------------------------------------------------------------------------------------
'* 注意事項：最後の1件分しか保持しません
'***************************************************************************************************
Public Property Get LastReceiveContentBINARY() As Byte()
    LastReceiveContentBINARY = m_LastReceiveContentBINARY
End Property

'***************************************************************************************************
'* 機能    ：最後の受信データをUTF-8文字列として、取得します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・最後の1件分しか保持しません
'            ・命名規則の都合上、`UTF8`とありますが実際は、VBAで扱える`String`型となります
'            ・ロジックの仕組み上、デバッグ用になると思います
'***************************************************************************************************
Public Property Get LastReceiveContentUTF8() As String
    LastReceiveContentUTF8 = Utf8Converter.BytesToString(m_LastReceiveContentBINARY)
End Property



'***************************************************************************************************
'                               ■■■ 新規接続関連 ■■■
'***************************************************************************************************
'* 機能　　：`WebSocket`方式としてCDP制御できる準備を行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：接続成功可否
'* 引数    ：UserName       利用者識別名称
'            Path           接続先パス      例：/devtools/page/8B4125AE4E31D852C475E291A9DE6D8C　※必ず「/」始まり！
'            targetId       `Target.TargetInfo`の`targetId` ※既存のタブ情報を流用する場合は指定すること
'            BiDiContext    WebDriverBiDiでのメインタブ情報 ※既存のタブ情報を流用する場合は指定すること
'            port           接続先ポート    例：9222
'            Host           接続先IP        例：127.0.0.1
'---------------------------------------------------------------------------------------------------
'* 機能説明：CDP準拠の引数構造です。基本は、`Path`のみで十分です
'* 注意事項：・既存の接続がある場合は、切断されます
'            ・`targetId`を指定した場合は、`CDPContext.cls`の`reattach`メソッドに適用して使用して下さい。省略時は、`CDPBrowser.cls`から
'***************************************************************************************************
Public Function ConnectCDP(UserName As String, Path As String, Optional targetID As String, Optional BiDiContext As String, Optional port As Long = D_port, Optional Host As String = D_host) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ConnectCDP"


    '1. winsock自体が生きてない場合は即抜け
    If Not m_WSAInitialized Then Exit Function

    '2. 既に接続済みの場合は切断後、再開
    If m_IsConnected Then printMsg WARN_, "This Class object already has a connection assigned to it. Disconnect the existing connection.", FromProcedureName: DisconnectCDP

    '3. ソケット生成
    m_hSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If m_hSocket = -1 Then printMsg ERROR_, "Failed to create WinSock handle. ErrorCode:" & Err.LastDllError, FromProcedureName, doRaiseError:=True

    '4. 丁度良いパフォーマンスとなるバッファーを設定
    Call setsockopt(m_hSocket, SOL_SOCKET, SO_RCVBUF, InitialBuffer, 4)

    '5. アドレス構造体の設定
    Dim addr As sockaddr_in
    addr.sin_family = AF_INET
    addr.sin_port = SwapBytes16(port) ' ポートのエンディアン変換 (純VBA)
    addr.sin_addr = inet_addr(Host)
    If addr.sin_addr = -1 Then printMsg ERROR_, "IP address translation failed. ErrorCode:" & Err.LastDllError, FromProcedureName, doRaiseError:=True

    '6. 接続実行 (TCPコネクションの確立)
    Dim ret As Long
    ret = Connect(m_hSocket, addr, LenB(addr))
    If ret <> 0 Then printMsg ERROR_, "Failed to connect to `Chrome DevTools Protocol`. ErrorCode:" & Err.LastDllError, FromProcedureName, doRaiseError:=True

    '7. WebSocketハンドシェイクの実行
    If Not SendHandshake(Host, port, Path) Then printMsg ERROR_, "The WebSocket-`Chrome DevTools Protocol` handshake failed. ErrorCode:" & Err.LastDllError, FromProcedureName, doRaiseError:=True

    '8. WebSocketハンドシェイクの実行結果確認
    If Not ReadHandshakeResponse Then printMsg ERROR_, "The results of the WebSocket-`Chrome DevTools Protocol` handshake could not be confirmed.", FromProcedureName, doRaiseError:=True

    '9. 設定情報を記録
    serialize UserName, targetID, BiDiContext

    '10. 各種ステータスを更新
    m_IsConnected = True
    ConnectCDP = True
    printMsg info_, "The `Chrome DevTools Protocol` control is ready. - " & Path, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：ポート番号のバイトスワップ (16bit)を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：Value  ポート番号
'---------------------------------------------------------------------------------------------------
'* 機能説明：htonsを使用した場合の32bit/64bitや符号付き整数のオーバーフロー問題を避けるために行います
'***************************************************************************************************
Private Function SwapBytes16(value As Long) As Integer
    Dim hi As Byte, lo As Byte
    lo = value And &HFF
    hi = (value And &HFF00) \ &H100

    Dim temp As Long
    temp = (lo * &H100&) Or hi

    ' 32768 (16進数で &H8000&) 以上の場合は、Integerの負数範囲にマッピング
    If temp >= 32768 Then
        SwapBytes16 = CInt(temp - 65536)
    Else
        SwapBytes16 = CInt(temp)
    End If
End Function



'***************************************************************************************************
'                              ■■■ WebSocket Upgrade関連 ■■■
'***************************************************************************************************
'* 機能　　：`WebSocket`方式として扱うフォーマットとして組み立てて、送信します
'---------------------------------------------------------------------------------------------------
'* 機能説明：引数は`ConnectWebSocket`と一緒なので省略します
'* 注意事項：ローカルIP運用を想定してるため、Keyは固定とします
'***************************************************************************************************
Private Function SendHandshake(Host As String, port As Long, Path As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".SendHandshake"


    Dim req As String
    req = "GET " & Path & " HTTP/1.1" & vbCrLf & _
          "Host: " & Host & ":" & port & vbCrLf & _
          "Upgrade: websocket" & vbCrLf & _
          "Connection: Upgrade" & vbCrLf & _
          "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==" & vbCrLf & _
          "Sec-WebSocket-Version: 13" & vbCrLf & vbCrLf

    'ASCIIバイト列に変換して送信
    Dim reqBytes() As Byte
    reqBytes = StrConv(req, vbFromUnicode)

    Dim sentBytes As Long
    sentBytes = Send(m_hSocket, reqBytes(0), UBound(reqBytes) + 1, 0)
    printMsg Debug_, "SendHandshake: " & Replace(req, vbCrLf, Chr(30)), FromProcedureName

    SendHandshake = (sentBytes > 0)
End Function

'***************************************************************************************************
'* 機能　　：前述の`SendHandshake`の応答結果を待ちます
'---------------------------------------------------------------------------------------------------
'* 注意事項：「edge://inspect/#remote-debugging」への接続の場合は、ダイアログに応答しないと一生フリーズします
'***************************************************************************************************
Private Function ReadHandshakeResponse() As Boolean
    Const FromProcedureName As String = ThisClassName & ".ReadHandshakeResponse"


    Dim response As String
    Dim buf(0 To 511) As Byte
    Dim bytesRead As Long
    Dim i As Long

    ' ハンドシェイク応答が終わる(CRLF * 2)まで、またはタイムアウトするまでブロッキングで読み取る
    Do
        bytesRead = recv(m_hSocket, buf(0), 512, 0)
        If bytesRead <= 0 Then Exit Do

        ' バイト配列を文字列に安全に復元
        Dim chunk As String
        chunk = Space$(bytesRead)
        For i = 1 To bytesRead
            Mid$(chunk, i, 1) = Chr$(buf(i - 1))
        Next i
        response = response & chunk     '文字列結合のアンチパターンではあるが、精々最初の接続1回しか使わないし、そこまで大きいサイズではないのでここは、妥協

        ' HTTPヘッダーの終端を検知
        If InStr(response, vbCrLf & vbCrLf) > 0 Then
            Exit Do
        End If
    Loop

    ' 「101 Switching Protocols」が含まれていればハンドシェイク成功
    printMsg Debug_, "Result-Handshake: " & Replace(response, vbCrLf, Chr(30)), FromProcedureName
    If InStr(response, "HTTP/1.1 101") > 0 Then
        ReadHandshakeResponse = True
    End If
End Function



'***************************************************************************************************
'                               ■■■ 送信関連 ■■■
'***************************************************************************************************
'* 機能    ：テキストデータで、WebSocketへ送信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：SendData           String 型
'* 返り値  ：実行結果コード     ※成功時は、0です
'---------------------------------------------------------------------------------------------------
'* 機能説明：テキストもWideCharToMultiByte を介したバイト配列変換を使用し、送信を行います
'* 注意事項：テキストも所詮、バイナリデータなので、UTF-8バイト配列に変換した上で、`SendAsyncMessageAsBINARY`を経由して、送信します
'***************************************************************************************************
Public Function SendCommandCDP(SendData As String) As Long
    SendCommandCDP = SendMessageAsBINARY(Utf8Converter.BytesFromString(SendData), True)
End Function

'***************************************************************************************************
'* 機能    ：バイト配列データで、WebSocketへ送信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：SendData           Byte 型の配列
'            TextType           True  : UTF-8テキスト送信 (Opcode: 0x01)
'                               False : バイナリ送信 (Opcode: 0x02)
'
'* 返り値  ：実行結果コード     ※成功時は、0です
'---------------------------------------------------------------------------------------------------
'* 注意事項：CDPでは、直バイナリとして送ることはないので、`Private`化してます
'***************************************************************************************************
Private Function SendMessageAsBINARY(SendData() As Byte, Optional TextType As Boolean) As Long
    ' 1. 送信Opcodeの決定 (テキスト=0x01, バイナリ=0x02)
    Dim opcode As Byte
    opcode = IIf(TextType, &H1, &H2)

    ' 2. WebSocketの仕様(RFC 6455)に準拠したマスク付きフレームを構築
    Dim wsFrame() As Byte
    wsFrame = BuildFrame(SendData, opcode)

    ' 3. WinSock経由でデータを全て送信
    SendMessageAsBINARY = RawSend(wsFrame)
End Function

'***************************************************************************************************
'* 機能　　：WebSocketフレームの構築（マスク処理含む）を行います
'***************************************************************************************************
Private Function BuildFrame(Payload() As Byte, opcode As Byte) As Byte()
    Dim payloadLen As Long
    payloadLen = UBound(Payload) - LBound(Payload) + 1

    ' ヘッダーサイズの計算
    ' 基本ヘッダー(2バイト) + 拡張サイズ領域(0/2/8バイト) + マスクキー(4バイト)
    Dim headerSize As Long
    Dim lenIndicator As Byte

    If payloadLen < 126 Then
        headerSize = 2 + 4
        lenIndicator = CByte(payloadLen)
    ElseIf payloadLen <= 65535 Then
        headerSize = 2 + 2 + 4
        lenIndicator = 126
    Else
        headerSize = 2 + 8 + 4
        lenIndicator = 127
    End If

    Dim frame() As Byte
    ReDim frame(0 To headerSize + payloadLen - 1)

    ' Byte 0: FIN (1) + RSV (000) + Opcode
    frame(0) = &H80 Or opcode

    ' Byte 1: MASK (1) + Length Indicator
    ' ※クライアントからサーバーへのフレームはMASKビット（&H80）が必須です
    frame(1) = &H80 Or lenIndicator

    Dim Offset As Long
    Offset = 2

    ' 拡張サイズ領域の書き込み（ビッグエンディアン）
    If lenIndicator = 126 Then
        frame(2) = CByte((payloadLen And &HFF00&) \ &H100&)
        frame(3) = CByte(payloadLen And &HFF&)
        Offset = 4
    ElseIf lenIndicator = 127 Then
        ' 32bitを超える長大なペイロード対応用（VBAのDouble/Long計算を安全に行う）
        Dim tempLen As Double
        tempLen = payloadLen
        frame(2) = 0: frame(3) = 0: frame(4) = 0: frame(5) = 0 ' 上位4バイトは0
        frame(6) = CByte(Int(tempLen / 2 ^ 24) And &HFF)
        frame(7) = CByte(Int(tempLen / 2 ^ 16) And &HFF)
        frame(8) = CByte(Int(tempLen / 2 ^ 8) And &HFF)
        frame(9) = CByte(tempLen And &HFF)
        Offset = 10
    End If

    ' 4バイトのランダムなマスクキーの生成
    Dim maskKey(0 To 3) As Byte
    Dim i As Long
    Static seeded As Boolean
    If Not seeded Then
        Randomize
        seeded = True
    End If

    For i = 0 To 3
        maskKey(i) = CByte(Int(Rnd * 256))
        frame(Offset + i) = maskKey(i)
    Next i
    Offset = Offset + 4

    ' ペイロードデータのマスク処理（XOR演算）をしながらコピー
    If payloadLen > 0 Then
        For i = 0 To payloadLen - 1
            frame(Offset + i) = Payload(i) Xor maskKey(i Mod 4)
        Next i
    End If

    BuildFrame = frame
End Function

'***************************************************************************************************
'* 機能　　：実際に、WinSockによる確実な分割/一括送信を行います
'***************************************************************************************************
Private Function RawSend(ByRef Data() As Byte) As Long
    Const FromProcedureName As String = ThisClassName & ".RawSend"


    Dim totalSent As Long
    Dim toSend As Long
    Dim ret As Long

    '送信サイズを求める
    toSend = UBound(Data) - LBound(Data) + 1

    ' 大容量データを送信する際、TCPの送信ウィンドウ制限による「分割送信」に対処するため
    ' 全バイトが完全にソケットに乗り切るまでループで送信を継続します
    Do While totalSent < toSend
        ret = Send(m_hSocket, Data(totalSent), toSend - totalSent, 0)
        If ret = -1 Then
            ' SOCKET_ERRORが発生した場合、エラーコードを返却
            RawSend = Err.LastDllError
            printMsg WARN_, "An error occurred while processing `send`. ErrorCode: " & RawSend & " - Detail: " & WinApiError.GetMessage(RawSend, "ws2_32"), FromProcedureName
            Exit Function
        End If
        totalSent = totalSent + ret
    Loop
End Function



'***************************************************************************************************
'                               ■■■ 受信関連 ■■■
'***************************************************************************************************
'* 機能    ：WebSocketからの非同期受信を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：読み取ったペイロードバイト数を返します。
'* 引数　　：EndOfData      `True` ：メッセージ全体またはその最後の部分
'                           `False`：メッセージの一部のみが含まれます
'
'            destruction    `True`で、フレームヘッダー情報を解析せずに、データを即破棄します。通常は、省略して下さい
'---------------------------------------------------------------------------------------------------
'* 詳細説明：必要に応じて、`WithEvents`による受け取りもご利用できます
'* 注意事項：・エラーコードは別途プロパティメソッドで取得してください
'            ・このプロシージャは、内部でループ処理は行いません。呼び出し側にて、引数やGETプロパティから状況を判断して、ループ処理を継続するかのロジックで組んでます
'            ・Pipe版と違って、「[タグ情報][ペイロード][タグ情報][ペイロード]...」というフレーム連鎖になってる都合上、
'            　処理の途中でVBE停止させると、ハンドルの使いまわし時に上手く機能しない場合があります。
'              その場合は、新規WebSocket接続からやり直すことをおすすめします。
'***************************************************************************************************
Public Function GetAsyncMessage(ByRef EndOfData As Boolean, Optional destruction As Boolean) As Long
    Const FromProcedureName As String = ThisClassName & ".GetAsyncMessage"


    '1. 今届いてるデータサイズを確認
    Dim bytesAvailable As Long
    Dim retIoctl As Long
    retIoctl = ioctlsocket(m_hSocket, FIONREAD, bytesAvailable)
    If retIoctl <> 0 Then
        ' ソケットがすでに死んでいる場合、エラー情報を記録して抜ける
        StyleErrPeekNamedPipe = Err.LastDllError
        printMsg WARN_, "An error occurred while processing `ioctlsocket`. ErrorCode: " & StyleErrPeekNamedPipe & " - Detail: " & WinApiError.GetMessage(StyleErrPeekNamedPipe, "ws2_32"), FromProcedureName
        Exit Function
    End If
    printMsg Trace_, "WinSockResSize: " & bytesAvailable, FromProcedureName

    '2. 状況に応じた分岐処理
    If bytesAvailable <= 0 Then
        '受信サイズ0以下なら即抜け
        printMsg Trace_, "The received data is empty.", FromProcedureName

        'ペイロード処理中の場合は、`False`に上書きしておきます
        '※破棄モードの場合は、スルー
        If Not destruction Then EndOfData = Not m_WebSocketFrame.ValidTags
        Exit Function

    ElseIf destruction Then
        '破棄モードの場合は、受信サイズ分を即取り出し
        Dim dSuccess As Long
        ReDim destructionData(bytesAvailable - 1) As Byte
        dSuccess = recv(m_hSocket, destructionData(0), bytesAvailable, MSG_WAITALL)

        '0以下の場合は失敗なので、エラー情報を記録
        If dSuccess <= 0 Then StyleErrReadFile = Err.LastDllError: printMsg WARN_, "An error occurred while using `recv` for disposal. ErrorCode: " & StyleErrReadFile & " - Detail: " & WinApiError.GetMessage(StyleErrReadFile, "ws2_32"), FromProcedureName

        '破棄したサイズを返却して、抜ける
        GetAsyncMessage = dSuccess
        Exit Function

    ElseIf Not m_WebSocketFrame.ValidTags Then
        'フレームヘッダーが無効(古い)状態なら更新させる
        printMsg Debug_, "This is not a valid frame header, so we will update it.", FromProcedureName

        '必要なフレームヘッダー情報が不足してる場合は、ここで抜ける
        'この場合、サイズ0、エラーなしとなります。
        If Not UpdateFrameHeader(bytesAvailable) Then Exit Function

        '更新後の、切断要求チェック (Opcode: 8)
        If m_WebSocketFrame.opcode = 8 Then
            Call DisconnectCDP
            m_LastBufferType = WS_BUFFER_TYPE.WS_CLOSE_MESSAGE ' プロパティを切断状態に
            printMsg WARN_, "A disconnection request has been received from `Chrome DevTools Protocol`.", FromProcedureName
            Exit Function
        End If
    End If

    '3. 必要な取り出しペイロードデータサイズを計算
    Dim TakePayloadSize As Long
    With m_WebSocketFrame
        If bytesAvailable <= 0 Then
            '`UpdateFrameHeader`プロシージャの減算処理過程で0以下になった場合。つまり、フレームヘッダー情報しか届かなかった場合
            TakePayloadSize = 0
            printMsg info_, "Only the frame header information was received. A request for further reception is required. TakeSize: " & TakePayloadSize, FromProcedureName

        ElseIf bytesAvailable >= .PayloadLength - .CurrentCursor Then
            'ペイロードサイズを超える受信データサイズを確認した場合は、不整合を防ぐため、ペイロードサイズのうち、未受信分のみに制限
            TakePayloadSize = .PayloadLength - .CurrentCursor
            printMsg Debug_, "I confirmed that the received data fit within the payload in a single reception. TakeSize: " & TakePayloadSize, FromProcedureName

        Else
            'ペイロードサイズに満たない場合は、そのまま適用
            TakePayloadSize = bytesAvailable
            printMsg info_, "The received data could not be confirmed to fit into the Payload in a single reception. A request for further reception is required. TakeSize: " & TakePayloadSize, FromProcedureName
        End If
    End With

    '4. 取り出せる状態であれば、実際に取り出します　※ここのフェーズで、`EndOfData`の更新を行います
    If TakePayloadSize > 0 Then
        Dim bSuccess As Long
        ReDim payloadBuffer(TakePayloadSize - 1) As Byte
        bSuccess = recv(m_hSocket, payloadBuffer(0), TakePayloadSize, MSG_WAITALL)

        '4-1. 0より大きければ、成功です
        If bSuccess > 0 Then
            '受信サイズを記録
            m_WebSocketFrame.CurrentCursor = m_WebSocketFrame.CurrentCursor + TakePayloadSize

            '全てのペイロードを受信出来たかを確認
            Dim AllPayloadReceive As Boolean
            AllPayloadReceive = (m_WebSocketFrame.CurrentCursor = m_WebSocketFrame.PayloadLength)

            'デバッグ調査用
'            Debug.Assert bSuccess = TakePayloadSize
'            Debug.Assert m_WebSocketFrame.CurrentCursor <= m_WebSocketFrame.PayloadLength

            'ストリームに書き込み、イベント発信
            WebSocketReceiveStream.Write payloadBuffer
            RaiseEvent WebSocketReceive(payloadBuffer, TakePayloadSize, AllPayloadReceive And m_WebSocketFrame.isFIN)
        Else
            '取り出しに失敗したら、エラー情報を記録して抜ける
            StyleErrReadFile = Err.LastDllError
            printMsg WARN_, "An error occurred while retrieving payload using `recv`. ErrorCode: " & StyleErrReadFile & " - Detail: " & WinApiError.GetMessage(StyleErrReadFile, "ws2_32"), FromProcedureName
            Exit Function
        End If

        '4-2. 全てのペイロードを受信出来た場合の仕上げ
        If AllPayloadReceive Then
            'ペイロード自体は全て取り出せたので、タグを無効化させて次回の取り出し時は、タグ情報を更新するようにする
            m_WebSocketFrame.ValidTags = False
            printMsg Debug_, "All payloads were received.", FromProcedureName

            'テキストとして送られた物か確認
            Dim TextType As Boolean
            TextType = (m_WebSocketFrame.opcode = &H1 Or m_WebSocketFrame.opcode = &H0) ' テキストまたは継続テキスト

            '`Base Header`としての、`FINフラグ`チェック
            If m_WebSocketFrame.isFIN Then
                '完成メッセージを取り出す
                With WebSocketReceiveStream
                    'メッセージが完全に揃ったので、取り出す末尾位置を設定する
                    .SetEOS

                    '先頭に戻して、`SetEOS`分まで読み込む
                    .Position = 0
                    m_LastReceiveContentBINARY = .Read

                    'ストリーム位置のリセット
                    .Position = 0
                End With

                ' 完了メッセージのタイプを設定
                m_LastBufferType = IIf(TextType, WS_BUFFER_TYPE.WS_UTF8_MESSAGE, WS_BUFFER_TYPE.WS_BINARY_MESSAGE)
                EndOfData = True
            Else
                ' まだ分割データの受信中
                m_LastBufferType = IIf(TextType, WS_BUFFER_TYPE.WS_UTF8_FRAGMENT, WS_BUFFER_TYPE.WS_BINARY_FRAGMENT)
                EndOfData = False
            End If
        Else
            ' ペイロード自体が欠けてます
            m_LastBufferType = WS_BUFFER_TYPE.WS_PAYLOAD_FRAGMENT
            printMsg Debug_, "Not all payloads have been received yet.", FromProcedureName
            EndOfData = False
        End If
    Else
        'ペイロードデータが届いてない
        m_LastBufferType = WS_BUFFER_TYPE.WS_PAYLOAD_EMPTY
        printMsg Debug_, "Payload data has not been received.", FromProcedureName
        EndOfData = False
    End If

    '5. 取り出したペイロードサイズとして、返却
    GetAsyncMessage = TakePayloadSize
End Function

'***************************************************************************************************
'* 機能　　：フレームヘッダー情報を更新します
'---------------------------------------------------------------------------------------------------
'* 引数    ：ResSize    `ioctlsocket`でチェックした受信サイズ。タグチェックに使い、追加受信分の減算処理もして、ペイロードサイズ分だけにするように調整します。参照渡しの恩恵です
'* 返り値  ：フレームヘッダーの読み取りに成功すれば`True`です。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://triple-underscore.github.io/RFC6455-ja.html#section-5.2
'* 注意事項：・ペイロードを全て読み終わった後にこれを呼び出すこと。ペイロードの欠け状態でこれを呼び出すと、バグります
'            ・コードをシンプルに保つため、「Extended payload」(中,大箱)部分の読み取りのみは、コンマ秒のブロッキングが起こることがありますが、ローカル環境なら無視できるレベルです
'***************************************************************************************************
Private Function UpdateFrameHeader(ByRef resSize As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".UpdateFrameHeader"


    '1. 最低限必要なフレームヘッダー情報が不足してる場合は、ここで抜ける。
    If resSize < 2 Then printMsg Trace_, "Frame header information has not been received.", FromProcedureName: Exit Function

    '2. フレームヘッダーを解析
    With m_WebSocketFrame
        '有効なフラグとして適用
        .ValidTags = True

        'カーソル位置をリセット
        .CurrentCursor = 0

        '2-2. 先頭の基本ヘッダー2バイトを受信
        '[!NOTE] 最初の2バイトに書かれている「タグ情報」の中身
        'WebSocketから送られてくるデータの先頭2バイトには、以下のような「このパケットが何者なのか」を示す超重要データがぎっしり詰まっています。
        '
        '---【1バイト目（8ビット）】── データの「状態」と「種類」---
        '・FINフラグ（先頭1ビット目）：「これが最後のパケットだよ！（1）」か「まだ続きがあるよ！（0）」か。
        '・RSV1フラグ(先頭2ビット目) ：「圧縮」してるかのフラグ。※CDPでは使用しないため、無視
        '・Opcode（後半4ビット分）   ：「これはテキストだよ（1）」「バイナリだよ（2）」「切断要求だよ（8）」などのデータ種類。
        '
        '---【2バイト目（8ビット）】── データの「長さのヒント」---
        '・MASKフラグ（先頭1ビット目）：マスク処理されているか。※CDPでは使用しないため、無視
        '・長さ指示子（後半7ビット分）：「本体のサイズ」または「中箱・大箱の指示フラグ」
        Dim header(0 To 1) As Byte
        Dim ret As Long
        ret = recv(m_hSocket, header(0), 2, MSG_WAITALL): .FrameHeaderSize = 2: resSize = resSize - 2
        If ret <= 0 Then
            '取り出しに失敗したら、エラー情報を記録して抜ける
            StyleErrReadFile = Err.LastDllError
            printMsg WARN_, "An error occurred while retrieving header using `recv`. ErrorCode: " & StyleErrReadFile & " - Detail: " & WinApiError.GetMessage(StyleErrReadFile, "ws2_32"), FromProcedureName
            Exit Function
        End If

        '2-3. 各種フラグ情報を取り出す
        .isFIN = ((header(0) And &H80) <> 0)        'メッセージ終端フラグ
        .opcode = (header(0) And &HF)               'ペイロードデータの種類

        '2-4. 今回の「完全なWebSocketフレームサイズ」を計算
        '[!NOTE] WebSocketの仕様を決めた偉い人たちは、「送るデータのサイズに合わせて、ヘッダーの大きさを自動的に変化させて、通信量を1バイトでも節約しよう」と考えました。
        'これが < 126、= 126、= 127 という3分岐の儀式の正体です。
        '
        '小箱（125バイト以下）：
        '短いテキストチャットなどを送る際、ヘッダーが無駄に大きいと邪魔です。そのため、最初の1バイト（指示子）だけでサイズを表現します。ヘッダーはわずか 2バイト で済みます。
        '
        '中箱（126バイト ~ 65535バイト）：
        '指示子に 126 と書かれていたら、それは「中箱」の合図です。
        '「サイズは、この直後に続く 2バイト（16bit） を見てね」という意味になり、ヘッダーは 4バイト に伸びます。
        '
        '大箱（65536バイト以上）：
        '指示子に 127 と書かれていたら、それは「大箱」の合図です。
        '「サイズは、この直後に続く 8バイト（64bit） を見てね」という意味になり、ヘッダーは最大で 10バイト まで巨大化します。
        'VBA側としては、相手がどの箱で送ってきたかに合わせて、ソケットから「追加で2バイト読み取るか、8バイト読み取るか」を変えなければならないため、この儀式が必要になります。
        Dim lenIndicator As Byte
        lenIndicator = (header(1) And &H7F)
        If lenIndicator < 126 Then
            printMsg Debug_, "Indicator is " & lenIndicator & ". Sets the header size to 2 bytes.", FromProcedureName


            '最初の1バイト（指示子）だけで、これからの受信サイズを表現してるので、そのまんま適用
            .PayloadLength = lenIndicator

        ElseIf lenIndicator = 126 Then
            printMsg Debug_, "Indicator is " & lenIndicator & ". Sets the header size to 4 bytes.", FromProcedureName


            '次の2バイト（16bit）に、これからの受信サイズを表現してるので、その分を追加で、取り出す
            Dim extLen16(0 To 1) As Byte
            ret = recv(m_hSocket, extLen16(0), 2, MSG_WAITALL): .FrameHeaderSize = .FrameHeaderSize + 2: resSize = resSize - 2
            If ret <= 0 Then
                '取り出しに失敗したら、エラー情報を記録して抜ける
                StyleErrReadFile = Err.LastDllError
                printMsg WARN_, "An error occurred while retrieving the payload size from the following 2 bytes using `recv`. ErrorCode: " & StyleErrReadFile & " - Detail: " & WinApiError.GetMessage(StyleErrReadFile, "ws2_32"), FromProcedureName
                Exit Function
            End If

            'うまいぐあいに、これからの受信サイズを取得
            .PayloadLength = (CLng(extLen16(0)) * &H100&) Or extLen16(1)

        ElseIf lenIndicator = 127 Then
            printMsg Debug_, "Indicator is " & lenIndicator & ". Sets the header size to 10 bytes.", FromProcedureName


            '次の8バイト（64bit） に、これからの受信サイズを表現してるので、その分を追加で、取り出す
            Dim extLen64(0 To 7) As Byte
            ret = recv(m_hSocket, extLen64(0), 8, MSG_WAITALL): .FrameHeaderSize = .FrameHeaderSize + 8: resSize = resSize - 8
            If ret <= 0 Then
                '取り出しに失敗したら、エラー情報を記録して抜ける
                StyleErrReadFile = Err.LastDllError
                printMsg WARN_, "An error occurred while retrieving the payload size from the following 8 bytes using `recv`. ErrorCode: " & StyleErrReadFile & " - Detail: " & WinApiError.GetMessage(StyleErrReadFile, "ws2_32"), FromProcedureName
                Exit Function
            End If

            'エンディアン変換儀式
            Dim tempLen As Double
            tempLen = (CDbl(extLen64(4)) * 2 ^ 24) + _
                      (CDbl(extLen64(5)) * 2 ^ 16) + _
                      (CDbl(extLen64(6)) * 2 ^ 8) + _
                       CDbl(extLen64(7))

            ' 上位4バイト（extLen64の0~3）が0以外、または計算結果がVBAのLongの上限(約21億)を超えている場合は、VBAの制限仕様なので停止
            If extLen64(0) <> 0 Or extLen64(1) <> 0 Or extLen64(2) <> 0 Or extLen64(3) <> 0 Or tempLen > 2147483647# Then
                Call DisconnectCDP
                printMsg ERROR_, "Payload size is too large to handle in VBA (exceeds 2GB). Disconnected", FromProcedureName, doRaiseError:=True
            End If

            'これからの受信サイズを取得
            .PayloadLength = CLng(tempLen)

        Else
            '念のため、用意
            Call DisconnectCDP
            printMsg ERROR_, "An invalid `Payload len` has been detected: " & lenIndicator, FromProcedureName, doRaiseError:=True

        End If
    End With

    '3. 全て更新し終えたので、成功フラグを返す
    UpdateFrameHeader = True
End Function



'***************************************************************************************************
'                                    ■■■ 切断関連 ■■■
'***************************************************************************************************
Public Sub DisconnectCDP()
    Const FromProcedureName As String = ThisClassName & ".DisconnectCDP"


    '1. 有効なハンドル値か確認
    If m_hSocket <> -1 And m_hSocket <> 0 Then
        '1-1. 切断処理
        Call closesocket(m_hSocket)

        '1-2. ハンドル値を無力化
        m_hSocket = -1
    End If

    '2. 接続フラグもリセット
    m_IsConnected = False
    printMsg info_, "It was cut off.", FromProcedureName
End Sub



'***************************************************************************************************
'                                  ■■■ ログヘルパー ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
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
        If LenB(Dir(currentLog.logPath, vbDirectory)) = 0 Then MkDir currentLog.logPath
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
'                                 ■■■ 新規接続アシスト ■■■
'***************************************************************************************************
'* 機能　　：「/json/list」へアクセスして、利用可能なすべてのウェブソケットターゲットのリストのうち、引数に基づいた`Page`接続のWebSocket接続まで行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：接続成功可否
'* 引数    ：UserName   利用者識別名称
'            url        URL
'            title      ページ名
'            port       接続先ポート    例：9222
'            Host       接続先IP        例：127.0.0.1
'---------------------------------------------------------------------------------------------------
'* 注意事項：・必須引数以外、全て省略すると、type=page の先頭タブに繋ぎに行きます
'            ・内部で「見つかるまでループ」と言ったことはしません。呼び出し側にて別途実装してください
'            ・このメソッドで接続する場合は、`CDPContext.cls`の`reattach`メソッドに適用して使用して下さい
'            ・このメソッドでの接続は`type:="page"`のみに限ります。
'***************************************************************************************************
Public Function AutoConnectPageCDP(UserName As String, Optional Url As String, Optional Title As String, Optional port As Long = D_port, Optional Host As String = D_host) As Boolean
    Const FromProcedureName As String = ThisClassName & ".AutoConnectPageCDP"


    '1. タブ単位でリクエストして、結果を取得
    Dim ResultRequest As WebResponse
    Set ResultRequest = checkDebugWs(True, Host, port)

    '2. ステータスチェック
    Dim websocketTargets As Collection
    With ResultRequest
        If .StatusCode = vbWebStatusOk Then
            '成功したら、タブ情報を取得
            Set websocketTargets = .Data
            printMsg info_, "Successfully retrieved tab information.", FromProcedureName
        Else
            '正常以外はエラーで停止
            printMsg ERROR_, .StatusDescription & " - HttpStatus: " & .StatusCode, FromProcedureName, doRaiseError:=True
        End If
    End With

    '3. スコア方式で目的のタブを見つける
    '「https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA/blob/main/Class%20Modules/CDPBrowser.cls#L667」のロジックを流用します
    Dim Obj
    Dim bestScore As Long, score As Long
    Dim candidateTabId  As String, tabId  As String, tabfound As Boolean
    For Each Obj In websocketTargets
        '3-1. タブの種類があってるか？
        If Obj("type") = "page" Then
            '3-2. 現スコアをリセット
            score = -1

            '3-3. 入力した引数に応じた確認
            Select Case True
                '第一優先：タブ名
                Case StrPtr(Title)
                    If Obj("title") = Title Then
                        score = 100
                    ElseIf InStr(Obj("title"), Title) Then
                        score = 50
                    Else
                        printMsg WARN_, "with name containing """ & Title & """", FromProcedureName
                    End If

                '第二優先：URL
                Case StrPtr(Url)
                    '一致位置検査
                    Dim StartPos As Long: StartPos = InStr(1, Obj("url"), Url)

                    '先頭一致
                    If StartPos = 1 Then
                        score = 90
                    '部分一致
                    ElseIf StartPos Then
                        score = 40
                    '一致せず
                    Else
                        printMsg WARN_, "with url containing """ & Url & """", FromProcedureName
                    End If

                '第三優先：直近の未接続タブ
                Case Else
                    score = 80
            End Select

            '3-4. 記録中のベストスコアを上回ったら更新
            If score > bestScore Then
                bestScore = score
                candidateTabId = Obj("id")
            End If
        End If
    Next

    '4. 目的の物っぽいのがあったら見つけたとします
    If bestScore > 0 Then
        tabfound = True
        tabId = candidateTabId
    End If

    '5. Returns nothing if not found
    'Nothing is preferred over throwing error
    Dim info As String
    If Not tabfound Then
        info = IIf(LenB(Title), Title, IIf(LenB(Url), Url, "(no new tab found)"))
        printMsg WARN_, "Unable to find the tab: " & info, FromProcedureName
        Exit Function
    Else
        printMsg info_, "Found the tab with ID: " & tabId, FromProcedureName
    End If

    '6. 見つけたパラメーターで、接続
    AutoConnectPageCDP = ConnectCDP(UserName, "/devtools/page/" & tabId, tabId, , port, Host)
End Function

'***************************************************************************************************
'* 機能　　：「/json/version」へアクセスして、ブラウザ接続のWebSocket接続まで行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：接続成功可否
'* 引数    ：UserName       利用者識別名称
'            ReuseContext   `True`で、Excelテーブルにあるメインタブ情報を流用します
'            port           接続先ポート    例：9222
'            Host           接続先IP        例：127.0.0.1
'---------------------------------------------------------------------------------------------------
'* 注意事項：このメソッドで接続する場合は、`CDPBrowser.cls`の`reattach`メソッドに適用して使用して下さい
'***************************************************************************************************
Public Function AutoConnectBrowserCDP(UserName As String, Optional ReuseContext As Boolean, Optional port As Long = D_port, Optional Host As String = D_host) As Boolean
    Const FromProcedureName As String = ThisClassName & ".AutoConnectBrowserCDP"


    '1. ブラウザ単位でリクエストして、結果を取得
    Dim ResultRequest As WebResponse
    Set ResultRequest = checkDebugWs(False, Host, port)

    '2. ステータスチェック
    With ResultRequest
        If .StatusCode = vbWebStatusOk Then
            '成功したら、`webSocketDebuggerUrl`を取得
            Dim webSocketDebuggerUrl As String
            webSocketDebuggerUrl = .Data("webSocketDebuggerUrl")
            printMsg info_, "Successfully retrieved Browser information.", FromProcedureName
        Else
            '正常以外はエラーで停止
            printMsg ERROR_, .StatusDescription & " - HttpStatus: " & .StatusCode, FromProcedureName, doRaiseError:=True
        End If
    End With

    '3. 末尾の"/"の位置を探索
    Dim StartPos As Long
    StartPos = InStrRev(webSocketDebuggerUrl, "/")

    '4. うまいぐあいに、接続パラメーターを取り出す
    Dim browserID As String
    browserID = Mid(webSocketDebuggerUrl, Len(webSocketDebuggerUrl) - StartPos + 2)

    '5. 既存のメインタブ情報を流用する場合はそれの取得を試み、接続を行う
    If ReuseContext Then
        '5-1. Excelテーブルからメインタブ情報を取得
        Dim tmp As Scripting.Dictionary
        Set tmp = ShSetting01_StartBrowser.TableBrowserContext(UserName, FromProcedureName)

        '5-2. データ有無チェック
        If Not tmp Is Nothing Then
            '6. あったらその情報を適用し、抜ける
            AutoConnectBrowserCDP = ConnectCDP(UserName, "/devtools/browser/" & browserID, tmp("targetID"), tmp("BiDi-context"), D_port, D_host)
            Exit Function
        End If
    End If

    '6. そのパラメーターで接続
    AutoConnectBrowserCDP = ConnectCDP(UserName, "/devtools/browser/" & browserID, , , D_port, D_host)
End Function

'***************************************************************************************************
'* 機能　　：事前に設定したパラメーターを元に、共通設定を施して、ws情報を取得するAPIアクセスを行います
'***************************************************************************************************
Private Function RequestHost(RequestParams As WebRequest, EndPointURL As String) As WebResponse
    '通信の設定
    Dim EndPointAPI As New WebClient
    With EndPointAPI
        'ホスト名URL
        .BaseUrl = EndPointURL

        'タイムアウト設定(5s程度)
        .TimeoutMS = 5000
    End With

    '事前設定と共通設定を使ってリクエストし、結果を返します
    Set RequestHost = EndPointAPI.Execute(RequestParams)
End Function

'***************************************************************************************************
'* 機能　　：タブ or ブラウザ単位でのws情報を取得する際のリクエストパラメータを生成し、リクエストします
'***************************************************************************************************
Private Function checkDebugWs(TabMode As Boolean, Host As String, port As Long) As WebResponse
    '1. リクエストパラメーターを設定する
    Dim RequestParams As New WebRequest
    With RequestParams
        'GET
        .Method = vbWebHttpGet

        'レスポンスのデータ形式を指定
        .ResponseFormat = vbWebFormatJson

        'パスパラメーターを指定
        .Resource = "/json/" & IIf(TabMode, "list", "version")
    End With

    '2. リクエストして、結果を返却
    Set checkDebugWs = RequestHost(RequestParams, "http://" & Host & ":" & port)
End Function

'***************************************************************************************************
'* 機能　　：「DevToolsActivePort」というファイルにアクセスして、今目の前のブラウザに接続処理します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：接続成功可否
'* 引数    ：UserName       利用者識別名称
'            ReuseContext   `True`で、Excelテーブルにあるメインタブ情報を流用します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://developer.chrome.com/blog/chrome-devtools-mcp-debug-your-browser-session?hl=ja
'* 注意事項：・このメソッドで接続する場合は、`CDPBrowser.cls`の`reattach`メソッドに適用して使用して下さい
'            ・現時点では、Edge OR Chrome の安定版への接続用に限ります
'***************************************************************************************************
Public Function AutoConnectDevToolsActivePort(UserName As String, Optional ReuseContext As Boolean) As Boolean
    Const FromProcedureName As String = ThisClassName & ".AutoConnectDevToolsActivePort"


    '1. `DevToolsActivePort`のファイルパスを特定
    Dim FilePath_DevToolsActivePort As String
    If ShSetting01_StartBrowser.UseRangeID(4, FromProcedureName) Then
        FilePath_DevToolsActivePort = Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\DevToolsActivePort"
    Else
        FilePath_DevToolsActivePort = Environ("LOCALAPPDATA") & "\Microsoft\Edge\User Data\DevToolsActivePort"
    End If

    '2. そのファイルがあるか？
    If LenB(Dir(FilePath_DevToolsActivePort)) Then
        'そのパスを読み込んで、接続情報を取得する
        '3. ファイルOpen
        Dim intFree As Integer
        intFree = FreeFile
        Open FilePath_DevToolsActivePort For Input As #intFree

        '4. 中身を全部読む
        Dim tmp
        tmp = Input(LOF(intFree), #intFree)

        '5. 閉じる
        Close #intFree

        '6. 分割
        '※最初にポート番号、その次にURLパスが記述されてます
        tmp = Split(tmp, vbLf)

        '7. 既存のメインタブ情報を流用する場合はそれの取得を試み、接続を行う
        If ReuseContext Then
            '7-1. Excelテーブルからメインタブ情報を取得
            Dim tmpRe As Scripting.Dictionary
            Set tmpRe = ShSetting01_StartBrowser.TableBrowserContext(UserName, FromProcedureName)
    
            '7-2. データ有無チェック
            If Not tmpRe Is Nothing Then
                '8. あったらその情報を適用し、抜ける
                AutoConnectDevToolsActivePort = ConnectCDP(UserName, CStr(tmp(1)), tmpRe("targetID"), tmpRe("BiDi-context"), CLng(tmp(0)))
                Exit Function
            End If
        End If

        '8. 接続
        AutoConnectDevToolsActivePort = ConnectCDP(UserName, CStr(tmp(1)), , , CLng(tmp(0)))
    Else
        'エラーで終了
        printMsg ERROR_, "Connection information is missing." & vbCrLf & "> " & FilePath_DevToolsActivePort, FromProcedureName, doRaiseError:=True
    End If
End Function



'***************************************************************************************************
'                            ■■■ 設定情報読み書き処理 ■■■
'***************************************************************************************************
'* 機能　　：接続情報をExcelテーブルに残します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       利用者識別名称
'            targetId       `Target.TargetInfo`の`targetId`
'            BiDiContext    WebDriverBiDiでのメインタブ情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`remote-debugging-pipe`と共存する都合上、一部は仮値にしてます
'            ・ここでの、Dictionary.Add の Item 引数は `(hoge)` と括弧必須。64bit Dictionary が LongPtr を
'            　参照渡しと誤解して 稀にクラッシュする不具合を回避するための強制値渡しです。
'***************************************************************************************************
Private Sub serialize(UserName As String, targetID As String, BiDiContext As String)
    '------------------ 1. パイプ情報の記録準備 ------------------
    '※主要となる情報以外は一旦、一律0とし、必要なデータを`Dictionary`に詰める
    Dim tmp As New Dictionary
    tmp.Add "hStdOutRd", 0
    tmp.Add "hStderrOutRd", 0
    tmp.Add "hStdInWr", 0
    tmp.Add "hCDPOutRd", (m_hSocket)
    tmp.Add "hCDPInWr", (m_hSocket)
    tmp.Add "hProcess", 0
    tmp.Add "dwProcessId", 0

    'Excelテーブルに、ハンドル情報を記録する
    Set ShSetting01_StartBrowser.TableBrowserHandle(UserName, "CDPCoreViaWebSocket.serialize") = tmp


    '------------------ 2. タブ情報の記録準備 ------------------
    tmp.RemoveAll
    tmp.Add "BiDi-context", BiDiContext
    tmp.Add "sessionID", vbNullString   '新規でWebSocket接続する場合は、`SessionID`の使い回しができないので、未設定とする
    tmp.Add "targetID", targetID

    'Excelテーブルに、タブ情報欄を確保する
    Set ShSetting01_StartBrowser.TableBrowserContext(UserName, "CDPCoreViaWebSocket.serialize") = tmp
End Sub

'***************************************************************************************************
'* 機能　　：接続情報をExcelテーブルから取得し、クラス内変数に適用します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   利用者識別名称
'* 返り値  ：ハンドル生存可否
'               -1:Excelテーブルに記録されてないため利用できません
'                0:成功
'               >0:WinSockエラー
'---------------------------------------------------------------------------------------------------
'* 詳細説明：WebSocketへの再接続の手間を減らす時にご利用ください。この場合、`SessionID`はKeepされます
'* 注意事項：うまくいかない場合は、新規接続からやり直して下さい。まぁ、おまけ程度として
'***************************************************************************************************
Public Function deserialize(UserName As String) As Long
    Const FromProcedureName As String = ThisClassName & ".deserialize"


    'Excelテーブルから、`UserName`に紐づくパイプハンドル情報を読み込む
    Dim PipeInfo As Dictionary
    Set PipeInfo = ShSetting01_StartBrowser.TableBrowserHandle(UserName, FromProcedureName)

    '取得に失敗した場合は、Excelのテーブルに記録されてないので抜ける
    If PipeInfo Is Nothing Then deserialize = -1: Exit Function

    'ハンドルに適用
    m_hSocket = PipeInfo("hCDPOutRd")

    '生存確認
    Dim tmp As Long
    If ioctlsocket(m_hSocket, FIONREAD, tmp) <> 0 Then deserialize = Err.LastDllError
    printMsg Debug_, WinApiError.GetMessage(deserialize, "ws2_32"), FromProcedureName
End Function



'***************************************************************************************************
'                         ■■■ クラスオブジェクトの初期化と破棄時の処理 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    ' WinSock 2.2 の初期化
    Dim wsaDat As WSAData
    If WSAStartup(&H202, wsaDat) = 0 Then
        m_WSAInitialized = True
    Else
        Err.Raise vbObjectError + 1001, "CDPCoreViaWebSocket.Class_Initialize", "WSAStartup の初期化に失敗しました。"
    End If

    '各種ステータスを初期化
    m_hSocket = -1
    m_LastBufferType = WS_NOT_MESSAGE
    currentLog.id = "WSID" & Format(Rnd * 1000, "000")

    '蓄積用バイト配列バッファーの初期化
    ReDim binBuffer(InitialBuffer - 1) As Byte
    Set WebSocketReceiveStream = CreateObject("ADODB.Stream")
    With WebSocketReceiveStream
        .Open
        .Type = 1           ' adTypeBinary
        .Write binBuffer    ' 疑似的に大きめのバッファーを確保
        .Position = 0       ' カーソルだけ先頭に戻す
        .SetEOS             ' 終端（EOS）としてマークする
    End With
End Sub

Private Sub Class_Terminate()
    'クリーンして...
    If m_WSAInitialized Then
        Call WSACleanup
        m_WSAInitialized = False
    End If

    'ストリームをClose
    WebSocketReceiveStream.Close
    Set WebSocketReceiveStream = Nothing
End Sub
