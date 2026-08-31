VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPCoreViaNativeMessaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'***************************************************************************************************
'   Chrome拡張機能の`chrome.debugger`×`Native Messaging`経由での、`Chrome DevTools Protocol`のやり取りに
'   特化した専用Classモジュールです。
'
'                    `CDPCore.cls`とやり取りしやすいように、工夫して組んでます。
'　基本的には、このClassオブジェクトの使いまわしはしない想定で組んでます。
'
'
'====================== どういう仕組み？ ======================
'   Chromeの`Native Messaging`は、拡張機能が`chrome.runtime.connectNative`した瞬間に、
'   OS側に登録されたホスト実行ファイルをChrome自身が`CreateProcess`相当で起動し、
'   そのプロセスの標準入力/標準出力をパイプとして直接掴む仕様です。
'
'   このプロジェクトでは、そのホスト実行ファイルとして`EXCEL.EXE`自身（もしくはそれを起動するラッパー）を
'   登録する運用を想定しています。つまり、このClassが動いてるExcelプロセス自身が、
'   Chromeに起動された「ネイティブホスト」そのものなので、
'   `CDPCoreViaWebSocket.cls`のように「これから接続しにいく」のではなく、
'   「既に渡されている自分の標準入出力ハンドルを回収する」のが`ConnectCDP`の役割になります。
'
'   ワイヤーフォーマットは、4バイト・リトルエンディアンの長さプレフィックス + UTF-8 JSON本体という、
'   `WebSocket`のRFC6455フレームよりずっと単純な形式です。
'
'   これにより､WebSocket版と同じロジックとして運用することが可能になりました｡
'   （`CDPCore.cls`側は、生バイト列を`CDPPipeStream`に蓄積→`ReadyCDPMessage`でテキスト化、という
'   　Pipe/WebSocketと全く同じ経路をそのまま使います。WebView2のような経路バイパスはしません）
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As LongPtr
Private Declare PtrSafe Function GetFileType Lib "kernel32" (ByVal hFile As LongPtr) As Long
Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As LongPtr, _
    lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    lpOverlapped As Long) As Long



'***************************************************************************************************
'                                   ■■■ 各種構造体 ■■■
'***************************************************************************************************
'Native Messagingの「4バイト長プレフィックス + ペイロード」のフレーム単位を管理する独自構造体
Private Type NativeMsgFrameManage
    PayloadLength       As Long     '今回来る1メッセージのペイロードサイズ
    CurrentCursor       As Long     '`PayloadLength`のうち、どこまで書き込んだ？
    ValidTags           As Boolean  '有効なヘッダー取得済みか？(=PayloadLengthが確定してるか)
End Type
Private m_Frame As NativeMsgFrameManage



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const STD_INPUT_HANDLE  As Long = -10
Private Const STD_OUTPUT_HANDLE As Long = -11
Private Const FILE_TYPE_PIPE    As Long = 3
Private Const FrameHeaderSize As Long = 4   'フレームヘッダーのバイト数(4バイト長プレフィックス)

'その他
Private Const ThisClassName As String = "CDPCoreViaNativeMessaging"



'***************************************************************************************************
'                                   ■■■ イベント定義 ■■■
'***************************************************************************************************
'Native Messaging受信発信機 ※`CDPCoreViaWebSocket.WebSocketReceive`と同じ形
Public Event NativeMessageReceived(res() As Byte, resSize As Long, EOS As Boolean)



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'--- 状態管理 ---
Private m_hStdIn            As LongPtr    '自プロセス自身の標準入力ハンドル(Chromeから見た書き込み先)
Private m_hStdOut           As LongPtr    '自プロセス自身の標準出力ハンドル(Chromeから見た読み取り元)
Private m_IsConnected       As Boolean    '既にNative Messagingとして利用可能な状態か？
Private m_LastReceiveContentBINARY() As Byte    '最後のNative Messagingから得た完成受信データ(バイト配列)

'--- ログ設定状況 ---
Private currentLog As StateLog

'--- Pipe/WebSocket版に合わせたエラー有無判定 ---
Private Type ErrorCode
    StyleErrWriteFile       As Long  'データ送信時(WriteFile)にてエラーになったか？
    StyleErrPeekNamedPipe   As Long  'データ存在有無(PeekNamedPipe)の確認でエラーになったか？
    StyleErrReadFile        As Long  'データを取り出してる最中(ReadFile)にエラーになったか？
End Type
Private LastNMError    As ErrorCode

'--- 各種クラスオブジェクト ---
Private NativeMessagingReceiveStream  As Object                       'Native Messagingバイトストリームを蓄積するオブジェクト
Private Utf8Converter As New CharacterCodeConversion  'UTF-8バイト配列に変換するオブジェクト



'***************************************************************************************************
'                                ■■■ check用プロパティ ■■■
'***************************************************************************************************
'* 機能    ：このClassオブジェクトが既に、Native Messagingとして利用可能かを確認します
'***************************************************************************************************
Public Property Get Connected() As Boolean
    Connected = m_IsConnected
End Property

'***************************************************************************************************
'* 機能    ：`GetAsyncMessage`プロシージャ使用時、`PeekNamedPipe`でエラーが起こった際のエラーコード
'***************************************************************************************************
Public Property Get IsErrorStylePeekNamedPipe() As Long
    IsErrorStylePeekNamedPipe = LastNMError.StyleErrPeekNamedPipe
End Property

'***************************************************************************************************
'* 機能    ：`GetAsyncMessage`プロシージャ使用時、`ReadFile`でエラーが起こった際のエラーコード
'***************************************************************************************************
Public Property Get IsErrorStyleReadFile() As Long
    IsErrorStyleReadFile = LastNMError.StyleErrReadFile
End Property

'***************************************************************************************************
'* 機能    ：標準入力パイプが生きてるかを確認します
'***************************************************************************************************
Public Function isAvailability() As Boolean
    Const FromProcedureName As String = ThisClassName & ".isAvailability"


    'パイプ生存判定(0バイトpeek)
    Dim tmp As Long
    If PeekNamedPipe(m_hStdIn, ByVal 0&, 0&, ByVal 0&, tmp, ByVal 0&) <> 0 Then isAvailability = True

    'ステータスを最新にしておく
    LastNMError.StyleErrPeekNamedPipe = Err.LastDllError
    printMsg Debug_, WinApiError.GetMessage(LastNMError.StyleErrPeekNamedPipe, "kernel32"), FromProcedureName
End Function

'***************************************************************************************************
'* 機能    ：Chromeに渡された標準入力ハンドル(診断用)
'***************************************************************************************************
Public Property Get Connection_hStdIn() As LongPtr
    Connection_hStdIn = m_hStdIn
End Property

'***************************************************************************************************
'* 機能    ：Chromeに渡された標準出力ハンドル(診断用)
'***************************************************************************************************
Public Property Get Connection_hStdOut() As LongPtr
    Connection_hStdOut = m_hStdOut
End Property



'***************************************************************************************************
'                          ■■■ 受信データ取得用プロパティ ■■■
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
'* 機能　　：`Native Messaging`方式としてCDP制御できる準備を行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：接続成功可否
'* 引数    ：ExpectedExtensionOrigin   （任意）呼び出し元拡張機能のorigin。参考照合のみに使い、判定自体には使いません
'---------------------------------------------------------------------------------------------------
'* 機能説明：他のtransportと違い「接続しにいく」のではなく、Chromeに起動された自プロセス自身が
'            既に受け取っている標準入出力ハンドルを回収・検証します
'* 注意事項：・通常のダブルクリック起動など、Chromeに起動されたのでない場合は`False`を返して即終了します
'            　（コマンドライン引数はChrome側で自由に変えられない/信頼できないため、判定には使いません。
'            　　`GetFileType`でパイプかどうかを見るのが唯一信頼できる判定方法です）
'***************************************************************************************************
Public Function ConnectCDP(Optional ExpectedExtensionOrigin As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ConnectCDP"


    '1. 既に接続済みの場合は切断後、再開
    If m_IsConnected Then printMsg WARN_, "This Class object already has a connection assigned to it. Disconnect the existing connection.", FromProcedureName: DisconnectCDP

    '2. 自プロセス自身の標準入出力ハンドルを回収
    m_hStdIn = GetStdHandle(STD_INPUT_HANDLE)
    m_hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)

    '3. ハンドルが有効か確認
    If m_hStdIn = 0 Or m_hStdIn = -1 Or m_hStdOut = 0 Or m_hStdOut = -1 Then
        printMsg info_, "Valid standard input/output handles could not be obtained.", FromProcedureName
        Exit Function
    End If

    '4. Chromeの`Native Messaging`は標準入力をパイプとして渡してくる。それ以外(コンソール等)なら通常起動と判断
    If GetFileType(m_hStdIn) <> FILE_TYPE_PIPE Then
        printMsg info_, "Standard input is not a pipe. This process was probably launched normally, not by Chrome's Native Messaging.", FromProcedureName
        Exit Function
    End If

    '5. 参考情報として、拡張機能のoriginをコマンドライン引数と照合 ※あくまで参考ログのみ。判定には使わない
    If Len(ExpectedExtensionOrigin) > 0 Then
        If InStr(1, Command$, ExpectedExtensionOrigin, vbTextCompare) = 0 Then
            printMsg WARN_, "The expected extension origin was not found in the command line arguments: " & Command$, FromProcedureName
        End If
    End If

    '6. 各種ステータスを更新
    m_IsConnected = True
    ConnectCDP = True
    printMsg info_, "The `Chrome DevTools Protocol` control is ready via Native Messaging.", FromProcedureName
End Function



'***************************************************************************************************
'                               ■■■ 送信関連 ■■■
'***************************************************************************************************
'* 機能    ：テキストデータで、Native Messaging経由の標準出力へ送信を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：SendData           String 型
'* 返り値  ：実行結果コード     ※成功時は、0です
'---------------------------------------------------------------------------------------------------
'* 機能説明：4バイト・リトルエンディアンの長さプレフィックス + UTF-8バイト配列、を1本のバイト配列として
'            組み立てた上で送信します
'* 注意事項：パイプへの同期`WriteFile`は、Windows側が内部で完了まで責任を持つため、`CDPCoreViaWebSocket`の
'            `RawSend`のような分割送信対応のループは不要です（`CDPCore.writeProcCDP`と同じ考え方）
'***************************************************************************************************
Public Function SendCommandCDP(SendData As String) As Long
    Const FromProcedureName As String = ThisClassName & ".SendCommandCDP"


    '1. UTF-8バイト配列へ変換
    Dim Payload() As Byte
    Payload = Utf8Converter.BytesFromString(SendData)

    Dim payloadLen As Long
    payloadLen = UBound(Payload) - LBound(Payload) + 1

    '2. 4バイト長プレフィックス(リトルエンディアン) + ペイロード、を1本のバイト配列に組み立て
    Dim frame() As Byte
    ReDim frame(0 To 3 + payloadLen)
    frame(0) = CByte(payloadLen And &HFF&)
    frame(1) = CByte((payloadLen \ &H100&) And &HFF&)
    frame(2) = CByte((payloadLen \ &H10000) And &HFF&)
    frame(3) = CByte((payloadLen \ &H1000000) And &HFF&)

    If payloadLen > 0 Then
        Dim i As Long
        For i = 0 To payloadLen - 1
            frame(4 + i) = Payload(i)
        Next i
    End If

    '3. Pipeへの`WriteFile`は1回のみ呼び出す(ループ送信しない)
    Dim lngWritten As Long
    Dim ResultCode As Long
    ResultCode = WriteFile(m_hStdOut, ByVal VarPtr(frame(0)), payloadLen + 4, lngWritten, ByVal 0&)

    '4. 失敗時は、実際のエラーコードを記録
    If ResultCode = 0 Then
        LastNMError.StyleErrWriteFile = Err.LastDllError
        printMsg WARN_, "An error occurred while processing `WriteFile`. ErrorCode: " & LastNMError.StyleErrWriteFile & " - Detail: " & WinApiError.GetMessage(LastNMError.StyleErrWriteFile, "kernel32"), FromProcedureName
    Else
        LastNMError.StyleErrWriteFile = 0
    End If

    '5. 結果を返却
    SendCommandCDP = LastNMError.StyleErrWriteFile
End Function



'***************************************************************************************************
'                               ■■■ 受信関連 ■■■
'***************************************************************************************************
'* 機能    ：Native Messaging(標準入力)からの非同期受信を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：読み取ったペイロードバイト数を返します。
'* 引数　　：EndOfData      `True` ：メッセージ全体またはその最後の部分
'                           `False`：メッセージの一部のみが含まれます
'
'            destruction    `True`で、フレームヘッダー情報を解析せずに、データを即破棄します。通常は、省略して下さい
'---------------------------------------------------------------------------------------------------
'* 詳細説明：必要に応じて、`WithEvents`による受け取りもご利用できます
'* 注意事項：・エラーコードは別途プロパティメソッドで取得してください
'            ・このプロシージャは、内部でループ処理は行いません。呼び出し側にて、引数やGETプロパティから状況を判断して、
'            　ループ処理を継続するかのロジックで組んでます(`CDPCoreViaWebSocket.GetAsyncMessage`と同じ契約)
'***************************************************************************************************
Public Function GetAsyncMessage(ByRef EndOfData As Boolean, Optional destruction As Boolean) As Long
    Const FromProcedureName As String = ThisClassName & ".GetAsyncMessage"


    '1. 今届いてるデータサイズを確認
    Dim bytesAvailable As Long
    Dim retPeek As Long
    retPeek = PeekNamedPipe(m_hStdIn, ByVal 0&, 0&, ByVal 0&, bytesAvailable, ByVal 0&)
    If retPeek = 0 Then
        'パイプが切れてる(相手先プロセス終了等)場合も含め、エラー情報を記録して抜ける
        LastNMError.StyleErrPeekNamedPipe = Err.LastDllError
        printMsg WARN_, "An error occurred while processing `PeekNamedPipe`. ErrorCode: " & LastNMError.StyleErrPeekNamedPipe & " - Detail: " & WinApiError.GetMessage(LastNMError.StyleErrPeekNamedPipe, "kernel32"), FromProcedureName
        Exit Function
    End If
    printMsg Trace_, "StdInResSize: " & bytesAvailable, FromProcedureName

    '2. 状況に応じた分岐処理
    If bytesAvailable <= 0 Then
        '受信サイズ0以下なら即抜け
        printMsg Trace_, "The received data is empty.", FromProcedureName

        'メッセージ処理中の場合は、`False`に上書きしておきます　※破棄モードの場合は、スルー
        If Not destruction Then EndOfData = Not m_Frame.ValidTags
        Exit Function

    ElseIf destruction Then
        '破棄モードの場合は、受信サイズ分を即取り出し
        Dim dSuccess As Long
        ReDim destructionData(bytesAvailable - 1) As Byte
        dSuccess = ReadFile(m_hStdIn, ByVal VarPtr(destructionData(0)), bytesAvailable, 0&, ByVal 0&)

        '0の場合は失敗なので、エラー情報を記録
        If dSuccess = 0 Then LastNMError.StyleErrReadFile = Err.LastDllError: printMsg WARN_, "An error occurred while using `ReadFile` for disposal. ErrorCode: " & LastNMError.StyleErrReadFile & " - Detail: " & WinApiError.GetMessage(LastNMError.StyleErrReadFile, "kernel32"), FromProcedureName

        '破棄したサイズを返却して、抜ける
        GetAsyncMessage = bytesAvailable
        Exit Function

    ElseIf Not m_Frame.ValidTags Then
        'フレームヘッダーが無効(古い)状態なら更新させる
        printMsg Debug_, "This is not a valid frame header, so we will update it.", FromProcedureName

        '最低限必要なフレームヘッダー情報が不足してる場合は、ここで抜ける
        'この場合、サイズ0、エラーなしとなります。
        If Not UpdateFrameHeader(bytesAvailable) Then Exit Function
    End If

    '3. 必要な取り出しペイロードデータサイズを計算
    Dim TakePayloadSize As Long
    With m_Frame
        If bytesAvailable <= 0 Then
            '`UpdateFrameHeader`の減算処理過程で0以下になった場合。つまり、ヘッダー情報しか届かなかった場合
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
        bSuccess = ReadFile(m_hStdIn, ByVal VarPtr(payloadBuffer(0)), TakePayloadSize, 0&, ByVal 0&)

        '4-1. 0以外なら、成功です
        If Not bSuccess = 0 Then
            '受信サイズを記録
            m_Frame.CurrentCursor = m_Frame.CurrentCursor + TakePayloadSize

            '全てのペイロードを受信出来たかを確認
            Dim AllPayloadReceive As Boolean
            AllPayloadReceive = (m_Frame.CurrentCursor = m_Frame.PayloadLength)

            'トリームに書き込み、イベント発信 ※ヘッダー部分(4バイト)は含めず、ペイロード部分のみを流す
            NativeMessagingReceiveStream.Write payloadBuffer
            RaiseEvent NativeMessageReceived(payloadBuffer, TakePayloadSize, AllPayloadReceive)
        Else
            '取り出しに失敗したら、エラー情報を記録して抜ける
            LastNMError.StyleErrReadFile = Err.LastDllError
            printMsg WARN_, "An error occurred while retrieving payload using `ReadFile`. ErrorCode: " & LastNMError.StyleErrReadFile & " - Detail: " & WinApiError.GetMessage(LastNMError.StyleErrReadFile, "kernel32"), FromProcedureName
            Exit Function
        End If

        '4-2. 全てのペイロードを受信出来た場合の仕上げ
        If AllPayloadReceive Then
            '次のメッセージのヘッダー取得に備えて、状態をリセット
            m_Frame.ValidTags = False
            EndOfData = True

            '完成メッセージを取り出す
            With NativeMessagingReceiveStream
                'メッセージが完全に揃ったので、取り出す末尾位置を設定する
                .SetEOS

                '先頭に戻して、`SetEOS`分まで読み込む
                .Position = 0
                m_LastReceiveContentBINARY = .Read

                'ストリーム位置のリセット
                .Position = 0
            End With

            printMsg Debug_, "All payloads were received.", FromProcedureName
        Else
            ' まだ分割データの受信中
            EndOfData = False

            printMsg Debug_, "Not all payloads have been received yet.", FromProcedureName
        End If

    Else
        'ペイロードデータが届いてない(ヘッダーのみ受信できた場合等)
        EndOfData = False
        printMsg Debug_, "Payload data has not been received.", FromProcedureName
    End If

    '5. 取り出したペイロードサイズとして、返却
    GetAsyncMessage = TakePayloadSize
End Function

'***************************************************************************************************
'* 機能　　：フレームヘッダー(4バイト長プレフィックス)情報を更新します
'---------------------------------------------------------------------------------------------------
'* 引数    ：resSize    `PeekNamedPipe`でチェックした受信サイズ。読み取った分だけ減算処理もして、
'                        ペイロードサイズ分だけになるように調整します。参照渡しの恩恵です
'* 返り値  ：4バイト長プレフィックスの読み取りが完了すれば`True`です。まだ揃っていない場合は`False`
'---------------------------------------------------------------------------------------------------
'* 注意事項：`ReadFile`を1回呼ぶごとに`resSize`分をまたぐ場合があるため、複数回の呼び出しにまたがって
'            `HeaderBuf`へ蓄積していく前提で組んでます
'***************************************************************************************************
Private Function UpdateFrameHeader(ByRef resSize As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".UpdateFrameHeader"


    '1. 最低限必要なフレームヘッダー情報が不足してる場合は、ここで抜ける。
    If resSize < FrameHeaderSize Then printMsg Trace_, "Frame header information has not been received.", FromProcedureName: Exit Function

    '2. フレームヘッダーを解析
    With m_Frame
        '2-1. 各種リセット
        .ValidTags = True   '有効なフラグとして適用
        .CurrentCursor = 0  'カーソル位置をリセット

        '2-2. 先頭の基本ヘッダー4バイトを受信
        Dim header(0 To 3) As Byte
        Dim ret As Long
        ret = ReadFile(m_hStdIn, ByVal VarPtr(header(0)), FrameHeaderSize, 0&, ByVal 0&): resSize = resSize - FrameHeaderSize
        If ret = 0 Then
            '取り出しに失敗したら、エラー情報を記録して抜ける
            LastNMError.StyleErrReadFile = Err.LastDllError
            printMsg WARN_, "An error occurred while retrieving header using `ReadFile`. ErrorCode: " & LastNMError.StyleErrReadFile & " - Detail: " & WinApiError.GetMessage(LastNMError.StyleErrReadFile, "kernel32"), FromProcedureName
            Exit Function
        End If

        '2-3. リトルエンディアンでペイロード長を確定
        '※上限は、64MiBまで。Chrome側のNative Messagingの仕様上、これ以上のサイズは送信できません
        .PayloadLength = CLng(header(0)) + CLng(header(1)) * &H100& + CLng(header(2)) * &H10000 + CLng(header(3)) * &H1000000
    End With

    '3. フレーム情報を表示する
    printMsg Debug_, "Frame header decoded. PayloadLength: " & m_Frame.PayloadLength, FromProcedureName

    '4. 全て更新し終えたので、成功フラグを返す
    UpdateFrameHeader = True
End Function



'***************************************************************************************************
'                                    ■■■ 切断関連 ■■■
'***************************************************************************************************
'* 機能    ：内部状態をリセットします
'---------------------------------------------------------------------------------------------------
'* 注意事項：継承済みの標準入出力ハンドルは、Chrome側が閉じるのに委ねます。ここでは強制`CloseHandle`しません
'***************************************************************************************************
Public Sub DisconnectCDP()
    Const FromProcedureName As String = ThisClassName & ".DisconnectCDP"


    m_hStdIn = 0
    m_hStdOut = 0
    m_IsConnected = False

    m_Frame.ValidTags = False

    'ストリームもClose
    NativeMessagingReceiveStream.Close
    Set NativeMessagingReceiveStream = Nothing

    printMsg info_, "It was cut off.", FromProcedureName
End Sub



'***************************************************************************************************
'                                  ■■■ ログヘルパー ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Const LogFileExtension As String = ".NativeMessaging.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath, vbObjectError + 1
End Sub



'***************************************************************************************************
'                         ■■■ クラスオブジェクトの初期化と破棄時の処理 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    '各種ステータスを初期化
    currentLog.LogID = "NMID" & Format(Rnd * 1000, "000")

    'ログ設定
    currentLog.seeRawSendMsgDbg = ShSetting01_StartBrowser.UseRangeID(8, FromProcedureName)
    currentLog.logPath = ShSetting01_StartBrowser.UseRangeID(9, FromProcedureName)

    '蓄積用バイト配列バッファーの初期化
    Set NativeMessagingReceiveStream = CreateObject("ADODB.Stream")
    With NativeMessagingReceiveStream
        .Type = 1           ' adTypeBinary
        .Open
    End With
End Sub

Private Sub Class_Terminate()
    DisconnectCDP
End Sub
