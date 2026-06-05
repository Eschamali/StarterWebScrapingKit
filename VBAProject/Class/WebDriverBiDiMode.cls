VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverBiDiMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'                    WebDriverBiDi制御モードとしてのメソッドやパラメータ運用を行います
'                       このClassオブジェクトは、低レベル操作として運用します
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'BiDiからのイベント情報を格納する際のDictionaryKey名設定
Private Const EventsDictionaryKeyName01 As String = "TotalEvents"
Private Const EventsDictionaryKeyName02 As String = "EventMethods"

'その他
Private Const TIMEOUT_DEFAULT       As Double = 30                  '初期タイムアウト秒数
Private Const ThisClassName         As String = "WebDriverBiDiMode" 'ログのトレース用
Private Const MaxStoreResultBiDi    As Long = 2 ^ 8                 'コマンド結果蓄積数



'***************************************************************************************************
'                                 ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'WebDriverBiDi制御用コアクラスオブジェクト
Private WithEvents m_BiDiCore   As WebDriverBiDiCore      'イベント受け取りオブジェクト

'各種便利クラスオブジェクト
Public jsConverter              As New WebJsonConverter

'BiDiコマンドの結果を蓄積する際の管理(`WithEvents`で直接は取らないようにする仕組み作り)
Private Type cacheResultBiDi
    response    As Dictionary   '「ID - それに対するCDPコマンドの結果」というDictionary
    Limit       As Long         'その`Dictionary`に詰め込む上限数
End Type
Private resultBiDi  As cacheResultBiDi

'ブラウザからの非同期イベント状況
Private m_Events    As Dictionary

'各種エラー情報
Private m_ErrorDetails  As Dictionary

'その他
Private timerOut        As Double  'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
Private LogID           As String
Private BiDiUserName    As String  'User-data-dir と同等



'***************************************************************************************************
'               ■■■ CDP非同期イベントからBiDi用の結果を受け取る処理 ■■■
'***************************************************************************************************
Private Sub m_CDPCore_CDPTabEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".m_CDPCore_CDPTabEvent"

    'BiDiでのやりとりの核となるイベント名
    Const BiDiForCDPEventName As String = "Runtime.bindingCalled"


    '1. このタブクラスで設定中の`SessionId`以外は即抜け
    If sessionID <> BiDiTab.sessionID Then Exit Sub

    '2. 対象のイベント以外も即抜け
    If methodName <> BiDiForCDPEventName Then Exit Sub
    If currentDebug.seeRawSendMsgDbg Then printMsg Trace_, "`Runtime.bindingCalled` was detected. Raw-Json : " & RawJson, FromProcedureName

    '3. PayloadのName部分を取得
    Dim NamePayload As String: NamePayload = ExtractName(RawJson)

    '4. `Name`値に応じた処理
    Dim tmp As String
    Select Case NamePayload
        'メインとなるWebDriverBiDiの結果
        Case BiDiResMain
            'エスケープ解除済みのpayload 部分を抽出
            tmp = ExtractPayload(RawJson)
            If currentDebug.seeRawSendMsgDbg Then printMsg Debug_, "Result - WebDriverBiDi: " & tmp, FromProcedureName

            '専用のCollectionに蓄積
            QueueResultBiDi.Add tmp

        'BiDi-Mapper 内部のデバッグログ用
        Case BiDiResDebug
            'エスケープ解除済みのpayload 部分を抽出し、ログ表示のみに留める
            tmp = ExtractPayload(RawJson)
            printMsg Debug_, "Debug - WebDriverBiDi: " & tmp, FromProcedureName

        'WebDriverBiDi 以外の`Runtime.bindingCalled`が返った模様
        '※この場合は、負荷の高いJsonパース処理はしない
        Case Else
            printMsg Debug_, "We detected `Runtime.bindingCalled` instances other than WebDriverBiDi. Name :" & NamePayload, FromProcedureName
    End Select

End Sub

'***************************************************************************************************
'* 機能　　：`Runtime.bindingCalled` のCDPメッセージから name 部分を抜き出します
'---------------------------------------------------------------------------------------------------
'* 返値　　：`name`部分
'* 引数　　：resCDP     CDPからの非同期イベント:`Runtime.bindingCalled`の内容そのまんま
'***************************************************************************************************
Private Function ExtractName(resCDP As String) As String
    Const FromProcedureName As String = ThisClassName & ".ExtractName"

    '--------- params-name項目から、`payload`の種類を特定 ---------
    Const SearchKey_ParamsName As String = "{""method"":""Runtime.bindingCalled"",""params"":{""name"":"""
    Dim StartPos As Long, EndPos As Long
    Dim CalledName As String

    'name項目を抽出
    StartPos = Len(SearchKey_ParamsName) + 1    '「{"method":"Runtime.bindingCalled","params":{"name":"」分、検索位置をずらす想定
    EndPos = InStr(StartPos, resCDP, """,""")
    ExtractName = Mid(resCDP, StartPos, EndPos - StartPos)
End Function

'***************************************************************************************************
'* 機能　　：`Runtime.bindingCalled` のCDPメッセージから payload 部分を高速に抜き出します
'---------------------------------------------------------------------------------------------------
'* 返値　　：エスケープ解除済み`payload`部分
'* 引数　　：resCDP     CDPからの非同期イベント:`Runtime.bindingCalled`の内容そのまんま
'***************************************************************************************************
Private Function ExtractPayload(resCDP As String) As String
    Const SearchKey As String = """payload"":"""
    Dim StartPos As Long: StartPos = InStr(1, resCDP, SearchKey)
    If StartPos = 0 Then Exit Function
    
    StartPos = StartPos + Len(SearchKey)
    
    ' payloadの終了位置を探す。通常はexecutionContextIdの前
    Dim EndPos As Long: EndPos = InStr(StartPos, resCDP, """,""executionContextId""")
    If EndPos = 0 Then EndPos = InStrRev(resCDP, """}") ' 万が一の末尾判定
    
    If EndPos > StartPos Then
        ' JSONエンコードされた文字列をデコード（本来はもっと複雑だが、ここではWebJsonConverter.ParseJsonに任せるために、
        ' いったん「JsonOfJson」の外殻だけ剥がす。ただしVBAのParserを完全に避けるならここが正念場）
        ' 実はParseJsonにこの部分文字列を渡すのが一番安全。
        ExtractPayload = Mid(resCDP, StartPos, EndPos - StartPos)
        
        ' 文字列内のエスケープ（\" -> "）を復元
        ' ※WebJsonConverter.ParseJson は「JSON全体」を期待するので、
        ' ここで完全に復元してしまうと、今度はその中の文字列が壊れる可能性がある。
        ' なので、ここでは「JSONとしての文字列値」を取得するにとどめ、
        ' クラス内のパース処理にそのまま渡す。
        ' ただし、CDPのJSONエンコードにより \" になっているものを " に戻す必要がある。
        ExtractPayload = Replace(ExtractPayload, "\""", """")
        ExtractPayload = Replace(ExtractPayload, "\\", "\")
    End If
End Function



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する所定の起動引数をつけて、WebDriverBiDiとして、起動させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：Name                           使用ブラウザ名※現時点では、Edgeとchromeのみです
'            appUrl                         最初の起動時に遷移したいURL
'            cleanActive                    `True` :デバッグブラウザを(あれば)Kill後、ブラウザを起動する
'                                           `False`:起動チェックをせずに、ブラウザを起動します
'
'            userProfile                    UserDataDir名を指定します
'            addArgs                        追加の起動引数を設けます
'            KioskMode                      0(省略)：通常モード(キオスクモードは使いません)
'                                           1      ：キオスクモード デジタル/対話型サイネージ
'                                           2      ：キオスクモード パブリック ブラウジング
'
'            sessionCapabilitiesRequest     `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'***************************************************************************************************
Public Sub StartBiDiMode(Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional cleanActive As Boolean = True, _
    Optional userProfile As String, Optional addArgs As String, Optional KioskMode As edgeKioskType, _
    Optional sessionCapabilitiesRequest As Dictionary)


    '1. 初期化処理
    BiDiUserName = userProfile
    Set m_BiDiCore = New WebDriverBiDiCore

    '2. WebDriverBiDi としてブラウザを起動させる
    m_BiDiCore.start Name, appUrl, cleanActive, BiDiUserName, addArgs, KioskMode, sessionCapabilitiesRequest

    '3. BiDi との新しいセッションを確立する
    Dim resultBiDi As Dictionary
    Set resultBiDi = sessionNew(sessionCapabilitiesRequest)

    '4. 接続先ブラウザ情報をログで表示
    ViewNewSession resultBiDi
End Sub

'***************************************************************************************************
'* 機能　　：既存の`BiDi-CDP Mapper`に再接続します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功可否論理値
'* 引数　　：userProfile                再アタッチしたいユーザー名(user-data-dirに基づく)
'            sessionCapabilitiesRequest `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'---------------------------------------------------------------------------------------------------
'* 機能説明：あとから、`BiDi-CDP Mapper`セッションに再接続する際に使い、セッション確立まで担います
'* 注意事項：ブラウザ自体が落ちてると、パイプが閉じてたり、無効といったエラーで止まります
'***************************************************************************************************
Public Function reattach(userProfile As String, Optional sessionCapabilitiesRequest As Dictionary) As Boolean
    Const FromProcedureName As String = ThisClassName & ".reattach"


    '1. CDPタブセッションに再接続
    Set m_CDPBrowser = New CDPBrowser
    Dim ResultConnectionCDP As Boolean
    ResultConnectionCDP = m_CDPBrowser.reattach(userProfile)

    'CDPタブセッション有無に応じた分岐
    If ResultConnectionCDP Then
        printMsg info_, "`Target.exposeDevToolsProtocol` instance is already running. Reattaching to existing BiDi session...", FromProcedureName

        '2. そのタブURLは、このクラスで用意した専用URLか？
        If LCase(MapperURL) = LCase(m_CDPBrowser.Url) Then
            '3. すでに当該タブで `BiDi-CDP Mapper`が起動済みかを確認する
            '→ typeof window.onBidiMessage === 'function' かどうかを調べる
            Dim jsCheck As String
            jsCheck = "typeof window.onBidiMessage === 'function';"

            Dim isMapperRunning As Boolean
            isMapperRunning = m_CDPBrowser.jsEval(jsCheck)

            '`BiDi-CDP Mapper`の起動状況に応じた分岐
            If isMapperRunning Then
                '4. `BiDi-CDP Mapper`は生きてるので、再バイディングを行う
                printMsg Debug_, "The `BiDi-CDP Mapper` instance is already running. Update the SessionID and reconnect to it.", FromProcedureName
                Call BiDiResponseBiding

            Else
                '4. お亡くなりなので、`BiDi-CDP Mapper`を再始動させます
                printMsg Debug_, "`BiDi-CDP Mapper` instance not found. ReBooting `BiDi-CDP Mapper` into reattached tab...", FromProcedureName
                StartBiDiCDPMapper True

                '5. BiDi との新しいセッションを確立し、ブラウザ情報をログで表示
                printMsg info_, "`BiDi-CDP Mapper` has been launched on the reattached tab", FromProcedureName
                ViewNewSession sessionNew(sessionCapabilitiesRequest)
            End If

        Else
            '3. 別URL状態なので、空ページに遷移させる
            '※`Target.exposeDevToolsProtocol`が有効の状態でのネットサーフィンは、危険なのでこの措置をとる
            '　超端的に言うと、ネット上のvbsコードがIEで実行されちゃって大変なことになるあれ
            printMsg WARN_, "The URL is not suitable for `Target.exposeDevToolsProtocol`. You will be redirected to an empty page.", FromProcedureName
            m_CDPBrowser.Url = MapperURL

            '4. `BiDi-CDP Mapper`を始動させる
            StartBiDiCDPMapper True

            '5. BiDi との新しいセッションを確立し、ブラウザ情報をログで表示
            printMsg info_, "`BiDi-CDP Mapper` has been launched on the reattached tab", FromProcedureName
            ViewNewSession sessionNew(sessionCapabilitiesRequest)
        End If
    Else
        '2. 既知タブ自体が失ってるので、新規タブで始動しなおす
        printMsg WARN_, "Mapper instance not found. Injecting `BiDi-CDP Mapper` into reattached new tab...", FromProcedureName
        StartBiDiCDPMapper

        '3. BiDi との新しいセッションを確立し、ブラウザ情報をログで表示
        printMsg info_, "`BiDi-CDP Mapper` has been launched on the reattached tab", FromProcedureName
        ViewNewSession sessionNew(sessionCapabilitiesRequest)
    End If

    '6. セッション一覧を取得して、既存のものがあるか確認する
    ' (BiDiは接続ごとにセッションIDがかわるわけではないが、安全のためセッション状態を確認)
    Dim resultBiDi As Dictionary
    Set resultBiDi = sessionStatus

    '7. mapperTab.js の仕様に沿った確認に従う
    If resultBiDi.Exists("ready") Then
        printMsg info_, "Successfully reattached to BiDi! Session status ready: " & resultBiDi("message"), FromProcedureName
        reattach = True
    Else
        printMsg ERROR_, "Failed to reattach to BiDi session.", FromProcedureName, doRaiseError:=True
        reattach = False
    End If
End Function

'***************************************************************************************************
'* 機能　　：新規タブにて、`BiDi-CDP Mapper`を始動させます
'* 引数　　：ReConnect  Excelテーブルにある`targetID`に対して再度、`BiDi-CDP Mapper`を再始動する際は、`True`にします
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`CDPBrowser`を`New`してからご利用ください
'            ・各パラメータは、`chromium-bidi\src\bidiServer\MapperCdpConnection.ts`に準拠してます
'***************************************************************************************************
Private Sub StartBiDiCDPMapper(Optional ReConnect As Boolean)
    Const FromProcedureName As String = ThisClassName & ".StartBiDiCDPMapper"


    Dim paramsCDP As New Dictionary
    With m_CDPBrowser
        '1. 新規タブに対して、BiDiコンソールタブを準備し、セッションターゲットをメイン化
        '※タブ表示モードの際は、そのタブがフォアグラウンドにならないように仕向けます
        If Not (ReConnect) Then Set WebDriverBiDi_mapperTab = .newTab(MapperURL, setMain:=True, isHidden:=Not (currentDebug.ViewDebugTab), isBackground:=True)
        printMsg Debug_, "Booting `BiDi-CDP Mapper`... (ViewTab: " & CStr(currentDebug.ViewDebugTab) & ")", FromProcedureName

        '2. 「BiDi-CDP Mapper」のタブ情報を取り出す
        BiDiTab.targetID = WebDriverBiDi_mapperTab.CurrentTargetID
        BiDiTab.sessionID = WebDriverBiDi_mapperTab.id

        '3. BiDi用バインディングの追加
        paramsCDP.Add "bindingName", "cdp"
        paramsCDP.Add "targetId", BiDiTab.targetID
        paramsCDP.Add "inheritPermissions", True
        .ExecuteCDP "Target.exposeDevToolsProtocol", paramsCDP, Not ReConnect   '再接続モードの場合は、エラー無視
    End With

    '4. JSからVBAへ返す経路を設定
    Call BiDiResponseBiding

    With WebDriverBiDi_mapperTab
        '5. パースしたJavaScriptをブラウザ上で実行
        .jsEval Take_mapperTabJSstring

        '6. ブートストラップ（初期化）JSの実行
        .jsEval "window.runMapperInstance('" & BiDiTab.targetID & "')"
    End With

    printMsg Debug_, "`BiDi-CDP Mapper` has been launched", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：BiDiからの結果やデバッグログを受け取るためのマッピング処理を行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`CDPBrowser`を`New`してからご利用ください
'            ・各パラメータは、`chromium-bidi\src\bidiServer\MapperCdpConnection.ts`に準拠してます
'            ・SessionIDが変わるたびにこれを行わないと、`BiDi-CDP Mapper`は生きてるのに、コマンドのやり取りが出来ない問題が起こります
'***************************************************************************************************
Private Sub BiDiResponseBiding()
    Const FromProcedureName As String = ThisClassName & ".BiDiResponseBiding"


    Dim paramsCDP As New Dictionary
    With WebDriverBiDi_mapperTab
        '1. JSからVBAへ、BiDiコマンドの結果/イベント payload を返す経路を設定
        paramsCDP.Add "name", BiDiResMain
        .ExecuteCDP "Runtime.addBinding", paramsCDP

        '2. JSからVBAへ、Mapper側で発生したログ（状況説明、警告、エラー寄り情報）を返す経路を設定
        If currentDebug.DiagnosticLog Then
            paramsCDP.RemoveAll
            paramsCDP.Add "name", BiDiResDebug
            .ExecuteCDP "Runtime.addBinding", paramsCDP
        End If
    End With

    printMsg Debug_, "I have configured the route to receive from `BiDi-CDP Mapper`.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：BiDiセッション接続時のブラウザ情報を表示します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ResultSessionNew   `session.new`時の結果Dictionaryオブジェクト
'---------------------------------------------------------------------------------------------------
'* 注意事項：・CDPのパイプ通信で1対1で直接繋がってしまっている都合上、一部項目は空となります
'            ・ログレベル - DEBUG のみ表示します
'***************************************************************************************************
Private Sub ViewNewSession(ResultSessionNew As Dictionary)
    Const FromProcedureName As String = ThisClassName & ".ViewNewSession"


    '現時点では、`platformName`は空欄なので、せっかくなので、このアプリケーションの名称にしてみます
    Dim pfName As String: pfName = ResultSessionNew("capabilities")("platformName")
    If pfName = "" Then pfName = Application.Name

    'ログビュー
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, " BiDi Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "sessionId             | " & ResultSessionNew("sessionId"), FromProcedureName
    printMsg Debug_, "acceptInsecureCerts   | " & ResultSessionNew("capabilities")("acceptInsecureCerts"), FromProcedureName
    printMsg Debug_, "browserName           | " & ResultSessionNew("capabilities")("browserName"), FromProcedureName
    printMsg Debug_, "browserVersion        | " & ResultSessionNew("capabilities")("browserVersion"), FromProcedureName
    printMsg Debug_, "platformName          | " & pfName, FromProcedureName
    printMsg Debug_, "setWindowRect         | " & ResultSessionNew("capabilities")("setWindowRect"), FromProcedureName
    printMsg Debug_, "webSocketUrl          | " & ResultSessionNew("capabilities")("webSocketUrl"), FromProcedureName
    printMsg Debug_, "userAgent             | " & ResultSessionNew("capabilities")("userAgent"), FromProcedureName
    printMsg Debug_, "mapperTab.js Version  | " & ShLibrary01_JS.Range(ShLibrary01_JS.UseRangeName(1, FromProcedureName)).value, FromProcedureName
End Sub



'***************************************************************************************************
'                           ■■■ BiDi-sessionコマンド関連 ■■■
'***************************************************************************************************
'* 機能　　：BiDiコマンド、`session.new`を実行します
'---------------------------------------------------------------------------------------------------
'* 返値　　：`session.new`時の結果Dictionaryオブジェクト
'* 引数    ：params     `session.new`の`Parameters`
'---------------------------------------------------------------------------------------------------
'* 注意事項：`BiDi-CDP Mapper`を始動後は、これを必ず実行しないと、BiDiコマンドの結果が返りません
'---------------------------------------------------------------------------------------------------
'* URL     ：https://w3c.github.io/webdriver-bidi/#command-session-new
'***************************************************************************************************
Private Function sessionNew(Optional params As Dictionary) As Dictionary
    '`session.new`の`Parameters`が省略されてる場合は、規定値を作成する
    If params Is Nothing Then
        Set params = New Dictionary
        params.Add "capabilities", New Dictionary
    End If

    '確立処理！
    Set sessionNew = invokeMethod("session.new", params)
End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンド、`session.status`を実行します
'---------------------------------------------------------------------------------------------------
'* 返値　　：`session.status`時の結果Dictionaryオブジェクト
'---------------------------------------------------------------------------------------------------
'* 注意事項：ソースコード上、結果は常に、「{ready: false, message: 'already connected'}」となります
'            →src\bidiMapper\modules\session\SessionProcessor.ts:47
'---------------------------------------------------------------------------------------------------
'* URL     ：https://w3c.github.io/webdriver-bidi/#command-session-status
'***************************************************************************************************
Private Function sessionStatus() As Dictionary
    Set sessionStatus = invokeMethod("session.status")
End Function



'***************************************************************************************************
'                              ■■■ BiDi送信準備関連 ■■■
'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (非同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "session.new")
'*           params        引数を持ったDictionary (省略可能)
'* 返値    ：送信したメッセージの ID
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを構成し、JSON化してブラウザへ投げます。結果の待機は行いません。
'***************************************************************************************************
Public Function invokeMethodAsync(ByVal methodName As String, Optional params As Dictionary) As Long
    Const FromProcedureName As String = ThisClassName & ".invokeMethodAsync"


    Dim paramStr As String
    Dim results As Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = m_JsonConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}", Async:=True)

    '実行したIDを返却
    invokeMethodAsync = results.keys(0)
    printMsg Debug_, "BiDi-AsyncInvoked """ & methodName & """ with ID: " & invokeMethodAsync, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "session.new")
'*           params        引数を持ったDictionary (省略可能)
'*           StopError     実行失敗時に`Err.Raise`で停止し、以降の処理を止めるか
'* 返値    ：ブラウザから返却されたBiDi結果 (パースされたDictionary)
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを送信後、ID返却されるまで待機します。
'***************************************************************************************************
Public Function invokeMethod(ByVal methodName As String, Optional params As Dictionary, Optional StopError As Boolean = True) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".invokeMethod"


    Dim paramStr As String
    Dim results As Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = m_JsonConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}", Async:=False)

    'Print error if there is
    If invokeError(results) Then
        If StopError Then
            printMsg ERROR_, "Failed to BiDi-invoke method """ & methodName & """" & vbCrLf & vbCrLf & "=== RawErrorResponse ===" & vbCrLf & m_JsonConverter.ConvertToJson(results), FromProcedureName, doRaiseError:=True
        Else
            printMsg WARN_, "Failed to BiDi-invoke method """ & methodName & """", FromProcedureName
        End If
    End If

    'Return result object(`CDPBrowser`の思想に従い、`Result`オブジェクトで返却します)
    If results.Exists("result") Then Set invokeMethod = results.Item("result")
    If currentDebug.seeRawSendMsgDbg Then printMsg Debug_, "BiDi-Invoked """ & methodName & """ with result: " & m_JsonConverter.ConvertToJson(results), FromProcedureName
End Function



'***************************************************************************************************
'                           ■■■ 核となるプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：`BiDiコマンド送信のコア処理
'---------------------------------------------------------------------------------------------------
'* 引数　　：strMessage             BiDi形式で送信するJsonプレーンテキスト
'            InternalUseMessages    非同期イベントをクラス内部で後程、使用したいときに、`New Dictionry`として、渡して有効化します。
'            Async                  `True` ：コマンドの結果を待ちません。Do Loopフェーズには入りません。
'                                   `False`：コマンドの結果を待ちます。Do Loopフェーズに入り、そのIDが来るまで待ちます。この場合、非同期に実行した結果のID回収もついでに行います
'
'* 返値　　：Dictionary形式の結果。Async=Trueの場合は ID:StatusCode等の簡易結果
'***************************************************************************************************
Private Function SendMessage(strMessage As String, Optional InternalUseMessages As Dictionary, Optional Async As Boolean) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".SendMessage"


    Dim lngCurrentId As Long
    lngCurrentId = m_MessageID

    'リクエスト数の上限に達したら、1にリセットしておく
    If m_MessageID + 1 > LimitID Then
        m_MessageID = 1
        printMsg info_, "Reset to `1` to prevent internal counter overflow.", FromProcedureName

        '不整合を防ぐため、残ってる結果は全消去
        currentResultBiDi.response.RemoveAll
        printMsg info_, "To prevent double storage, the history of asynchronous BiDi execution results has been completely deleted.", FromProcedureName
    Else
        m_MessageID = m_MessageID + 1   'We increase the global ID counter
    End If

    '末尾の「}」を削って id を付与する
    strMessage = Left(strMessage, Len(strMessage) - 1)
    strMessage = strMessage & ", ""id"":" & lngCurrentId & "}"

    printMsg Debug_, "sendBiDi-JsonCommand = " & strMessage, FromProcedureName
    printMsg Debug_, "runBiDi -Async       = " & Async, FromProcedureName

    'JSONのダブルクォートやエスケープ処理による構文乱れを防ぐためエスケープ
    Dim safeJson As String
    safeJson = Replace(strMessage, "\", "\\")
    safeJson = Replace(safeJson, "'", "\'")

    Dim paramsCDP As New Dictionary
    paramsCDP.Add "expression", "window.onBidiMessage('" & safeJson & "');"

    'CDP経由でBiDiコマンド送信
    '※JavaScript実行自体は同期処理。エラー時は、内部でStopがかかります。(chromium-bidi-main\src\bidiServer\MapperCdpConnection.ts　に基づく)
    '※また内部にて、`RaiseEvent`よるBiDi結果蓄積も行われます
    '※`.jsEval`は、内部のお節介によりオーバーヘッドが積み重なることから、直通メソッドで使用します
    WebDriverBiDi_mapperTab.ExecuteCDP "Runtime.evaluate", paramsCDP

    'BiDi自体が、非同期(結果はいらない)モードなら、整理券(ID)を渡してここで抜けます
    If Async Then
        Set SendMessage = New Dictionary

        '実行したコマンドIDと、送ったコマンド内容を渡して、抜けます
        SendMessage.Add lngCurrentId, paramsCDP("expression")
        Exit Function
    End If


   '------------------------------------------------
   'Read message returned by browser
   '------------------------------------------------
    Dim timerStart As Double: timerStart = TimerCounter

    'CDP受信処理→Collection内`Runtime.bindingCalled`処理→受信処理→Collection内`Runtime.bindingCalled`処理→…　というロジックでループ処理を行います
    'BiDi-Jsonコマンドを送った際のIDが返ってくるまで、ループ受信処理を行いますが、タイムアウトになったら、抜けます
    Dim ResultSendCommand As Boolean
    printMsg Trace_, "Starting Processes incoming messages from the browser...", FromProcedureName
    Do
        '先述の`.invokeMethod("Runtime.evaluate", paramsCDP)`にて既に、結果が蓄積されてる場合は、受信処理をSKIP
        If QueueResultBiDi.Count = 0 Then m_CDPCore.TakeEvents

        '受信数check
        Dim tmp As Dictionary
        If QueueResultBiDi.Count Then
            'データは届いたので、タイムアウト値をリセット
            timerStart = TimerCounter
            printMsg Trace_, "Data received, so reset the timeout value.", FromProcedureName

            '受信メッセージを活用するため、担当プロシージャへ遷移
            Set tmp = ReceiveFromCollectionPayload(InternalUseMessages, lngCurrentId)   'コマンド結果を受け取れたのに、再度実行されちゃうと結果が消えちゃうため、一時変数にまず格納しておく。`lngCurrentId`は、常に一定なので、オブジェクトありの状態は常に1度のみ

            'オブジェクトありなら、直近に送ったコマンド結果であることを確認できたので、受信ループ終了フラグを立て、改めて結果を返す
            If Not (tmp Is Nothing) Then
                ResultSendCommand = True
                Set SendMessage = tmp
            End If
        Else
            printMsg Trace_, "Data reception could not be confirmed. The timeout value will be incremented.", FromProcedureName

            'データがまだ、届いてないので非同期処理結果用の受信Boxを念の為、確認する
            '※`WithEvents`と併用すると、ここに格納される可能性があるためです
            Set tmp = ResultBiDiForAsync(lngCurrentId)
            If Not tmp Is Nothing Then
                '送ったコマンド結果であることを確認できたので、受信ループ終了フラグを立て、改めて結果を返す。1度結果を受け取ると同じIDでも、もう受け取れないので、True状態は常に1度のみ
                ResultSendCommand = True
                Set SendMessage = tmp
                printMsg info_, "Command ID:" & lngCurrentId & " was observed on the 'Asynchronous Result Dictionary side' rather than on the Collection. This is a phenomenon that is likely to occur when used in conjunction with `WithEvents`.", FromProcedureName
            Else
                printMsg Trace_, "I used `m_CDPBrowser.BiDiCapture` but it is empty.", FromProcedureName
            End If
        End If

        'いくつかのログ表示
        If Not ResultSendCommand Then
            'コマンド結果が、まだ来ない
            printMsg Trace_, "The BiDi-Json command result ID:" & lngCurrentId & " from `Runtime.bindingCalled` has not yet been received, so reception continues.", FromProcedureName
        End If

        '一定時間経ってもコマンド結果が来ないなら、タイムアウトエラーとして終了します
        If TimerCounter - timerStart > timerOut Then printMsg ERROR_, "Timeout waiting for BiDi to return message.", FromProcedureName, , True
        DoEvents

        'BiDi-Jsonコマンドの結果が得れたら、受信ループを終了します
    Loop Until ResultSendCommand


    printMsg Trace_, "BiDiCommand result ID:" & lngCurrentId & " was returned from BiDi-Json, so the receiving loop will end.", FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：`CDPBrowser.BiDiCapture`にある`Runtime.bindingCalled`データをを取り出し、引数に応じて、非同期で実行したIDの結果やイベントを別途格納します
'---------------------------------------------------------------------------------------------------
'* 引数　　：destruction   `True` :IDの解析やイベントの中身の解析を行わず、破棄します。解析しない分、多少パフォーマンス向上します
'                          `False`:IDの解析やイベントの中身の解析を行います
'***************************************************************************************************
Public Sub TakeEvents(Optional destruction As Boolean)
    Const FromProcedureName As String = ThisClassName & ".TakeEvents"


    Dim BiDiNum As Long

    printMsg Debug_, "Starts retrieving data from the BiDiCapture...", FromProcedureName
    Do
        'CDPイベント受信処理
        '※内部で、`RaiseEvent`が発火され、Collection変数に格納されます
        m_CDPCore.TakeEvents

        BiDiNum = QueueResultBiDi.Count
        printMsg Trace_, "BiDiNum = " & BiDiNum, FromProcedureName

        '受信数check
        If BiDiNum Then
            '破棄モードかの判定
            If destruction Then
                '解析などせず、すぐ初期化して破棄します
                '※Collectionのため、中身は直接見れません
                printMsg Trace_, "Discarding data in BiDiCapture...", FromProcedureName
                Set QueueResultBiDi = New Collection
            Else
                '中身の解析をし、IDやイベントを取り出します
                printMsg Trace_, "Analyzing the data in the BiDiCapture...", FromProcedureName
                Call ReceiveFromCollectionPayload      '単なるイベント収集のため、コマンドIDによる確認はしないように設定
            End If
        Else
            printMsg Trace_, "The message queue is currently empty.", FromProcedureName
        End If
    Loop Until BiDiNum = 0

    printMsg Debug_, "Finished getting data from the BiDiCapture.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：`CDPBrowser.BiDiCapture`に格納してあるコレクションを、1件1件処理し、状況に応じて内部で、イベントキャプチャしたり、コマンド結果の確認を行います
'---------------------------------------------------------------------------------------------------
'* 返値　　：`checkCommandID`に対するCDP-Jsonコマンド結果Dictionary形式
'* 引数　　：InternalUseMessages                非同期イベントをクラス内部で後程、使用したいときに、`New Dictionry`として、渡して有効化します。
'            checkCommandID                     確認するコマンドID
'---------------------------------------------------------------------------------------------------
'* 機能説明：`checkCommandID`に合わないIDが来ても`AccumulatedCommandResults`に蓄積されます
'***************************************************************************************************
Private Function ReceiveFromCollectionPayload(Optional ByRef InternalUseMessages As Dictionary, Optional ByRef checkCommandID As Long = -1) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".ReceiveFromCollectionPayload"


    '処理用の変数を用意
    Dim resBiDi
    Dim i As Long
    Dim workCol  As Collection

    '現時点での、BiDi結果を取り出し、クリーンアップする
    Set workCol = QueueResultBiDi
    Set QueueResultBiDi = New Collection

    printMsg Trace_, "Number of BiDi receptions:" & workCol.Count, FromProcedureName
    For Each resBiDi In workCol
        'ログ用に、カウント
        i = i + 1
        If currentDebug.seeRawSendMsgDbg Then printMsg Debug_, "BiDiRes(" & i & ") = " & resBiDi, FromProcedureName

        '空文字はそもそも受け付けないようにする
        If resBiDi <> "" Then
            '受信データのチェックを行うプロシージャへ遷移します。
            If BrowserReceivedDataCheck(CStr(resBiDi), InternalUseMessages, checkCommandID) Then
                'ここでcheckCommandID一致時のみ、パース（BiDi本体）を行う。
                Set ReceiveFromCollectionPayload = m_JsonConverter.ParseJson(CStr(resBiDi))
                printMsg Debug_, "The BiDicommand result was obtained from `Runtime.bindingCalled`! - ID:" & checkCommandID, FromProcedureName
            End If
        Else
            printMsg Debug_, "The payload is empty", FromProcedureName
        End If

        DoEvents
    Next

End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンドの受信データを高速にチェックし、仕分けます
'---------------------------------------------------------------------------------------------------
'* 返値    ：TRUE :受信ループを止める合図。
'               ・`BiDiJsonString`に、ExcelからのBiDi-Jsonコマンド結果がある
'               ・コマンド構文エラーによる、ID項目なし
'
'* 引数　　：BiDiJsonString         BiDiの生JSON文字列
'            InternalUseMessages    クラス内部プロシージャで、イベント内容を扱う用として蓄積させる参照変数
'*           checkCommandID         待ち受け中のコマンド結果ID
'***************************************************************************************************
Private Function BrowserReceivedDataCheck(BiDiJsonString As String, InternalUseMessages As Dictionary, checkCommandID As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".BrowserReceivedDataCheck"



    '======================================= チェック開始 =======================================
    Dim ResultCommandID As Long, AsynchronousEventName As String
    Dim JsonDictionaryObj As Dictionary
    Dim FlagValue As Long

    Const resultCommand As Long = 2 ^ 0 'IDのコマンド結果を検知出来た
    Const mismatchID    As Long = 2 ^ 1 'クラス内変数に記録してるIDとは違うコマンド結果が検知された。恐らく以前、非同期で実行したときの結果と思われる
    Const browserEvent  As Long = 2 ^ 2 '非同期イベントである
    Const EventCapture  As Long = 2 ^ 3 '非同期イベントキャプチャモードである
    Dim mask As Long: mask = resultCommand + mismatchID + browserEvent + EventCapture


    '--------- 1. BiDiからのイベントか？ ---------
    If GetContentsMethod(BiDiJsonString, AsynchronousEventName) Then
        FlagValue = FlagValue Or browserEvent   '非同期イベントとして、フラグ付け

        'VBAでのイベント発火を行う
        '※拡張機能の柔軟性を高めるため、ブラウザから得た非同期イベントは全て、イベント発火させます
        RaiseEvent detectionBiDiEvent(AsynchronousEventName, BiDiJsonString)


        '--------- 2. いずれかのDictionary変数が、非同期イベントキャプチャモードか？ ---------
        If (m_Events Is Nothing) And (InternalUseMessages Is Nothing) Then
            '両方とも`Nothing`なら、わざわざ中身を解析する意味がないので、パーサーは使わない
            printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName
        Else
            'イベントキャプチャモードなので、JsonPurserを使用する
            FlagValue = FlagValue Or EventCapture   'イベントキャプチャモードとして、フラグ付け
            printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
            Set JsonDictionaryObj = m_JsonConverter.ParseJson(BiDiJsonString)
        End If

    Else
        FlagValue = FlagValue Or resultCommand  'IDのコマンド結果を検知出来たとして、フラグ付け

        '--------- 2. `id`項目があるか？ ---------
        If GetContentsID(BiDiJsonString, ResultCommandID) Then
            '--------- 3. そのIDは、クラス内変数に記録してるIDと一致してるか？ ---------
            If ResultCommandID = checkCommandID Then
                '正しいコマンド結果が返ったので次項に、受信ループ終了フラグを返すようにする
                printMsg Trace_, "It was confirmed that the execution BiDicommand ID matches the one recorded in Excel. ID:" & checkCommandID, FromProcedureName
            Else
                '`WithEvents`によるズレなので、結果をキャッシュしておく
                FlagValue = FlagValue Or mismatchID '現在のコマンド結果IDと返ったコマンド結果IDが異なる場合のフラグ付け
                printMsg info_, "The BiDicommand result (ID:" & ResultCommandID & ") was returned, but it does not match the BiDicommand ID (" & checkCommandID & ") recorded in Excel. This is probably an asynchronous result. The results will be accumulated.", FromProcedureName

                'コマンド結果ID用受信Boxへ格納しておく
                '※パーサー処理負荷を抑えるため、Json文字列まんまとします
                currentResultBiDi.response.Add ResultCommandID, BiDiJsonString

                '格納件数チェックし、上限を超えたら、最古データを削除します
                If currentResultBiDi.response.Count > currentResultBiDi.Limit Then
                    ' 一番古いキーを取得して削除
                    printMsg Trace_, "Delete asynchronous processing result ID:" & currentResultBiDi.response.keys(0), FromProcedureName
                    currentResultBiDi.response.Remove currentResultBiDi.response.keys(0)
                    printMsg info_, "The accumulation limit has exceeded " & currentResultBiDi.Limit & ", so the oldest has been deleted.", FromProcedureName
                End If
            End If
        Else
            '`method`も`id`項目もない場合、BiDi-Jsonコマンドの構文エラーとして返った可能性があるとし、受信ループ終了とする
            printMsg WARN_, "Both the `id` and `method` fields are missing. This may be a syntax error.", FromProcedureName
        End If
    End If


    '======================================= このプロシージャから抜けるか判定 =======================================
    Select Case (FlagValue And mask)
        '送ったコマンドの結果は来たので、受信ループを止めて、抜ける
        Case resultCommand
            BrowserReceivedDataCheck = True
            Exit Function

        'コマンド結果だがID不一致。以前、非同期で実行したときの結果と思われるので、まだ受信ループは止めずに、抜ける
        Case resultCommand + mismatchID
            BrowserReceivedDataCheck = False
            Exit Function

        '受信ループは止めないが、イベントキャプチャしないモード(格納用のDictionaryがどちらもNothing or 古いSessionID)なので、次項の格納処理はせず、抜ける
        Case browserEvent
            BrowserReceivedDataCheck = False
            Exit Function

        'イベントキャプチャモードにより、Dictionaryオブジェクト格納フェーズへ行けるようにしておく。なお、受信ループは止めません
        Case browserEvent + EventCapture
            BrowserReceivedDataCheck = False

        '未知の組み合わせ
        Case Else
            printMsg ERROR_, "An unknown flag was passed. (" & (FlagValue And mask) & ")", FromProcedureName, doRaiseError:=True

    End Select


    '======================================= イベント情報を格納するフェーズへ =======================================
    'パブリック(呼び出し側使用)用
    If m_Events Is Nothing Then
        printMsg Trace_, "No event information is required on the module side.", FromProcedureName
    Else
        printMsg Trace_, "The module is requesting the accumulation of event information. Saving...", FromProcedureName
        OnNewEventReceived m_Events, JsonDictionaryObj, AsynchronousEventName
    End If

    'プライベート(クラス内変数)用
    If InternalUseMessages Is Nothing Then
        printMsg Trace_, "No event information needed inside the class.", FromProcedureName
    Else
        printMsg Trace_, "The class is requesting the accumulation of event information. Saving...", FromProcedureName
        OnNewEventReceived InternalUseMessages, JsonDictionaryObj, AsynchronousEventName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックし、`id`を高速抽出します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：jsonString         BiDiのレスポンス(JSON文字列そのまま)
'*            idValueContents   コマンド結果のID値を格納する変数(参照渡し)
'---------------------------------------------------------------------------------------------------
'* 注意事項：コマンド結果のエラー、成功の区別はしません。
'***************************************************************************************************
Private Function GetContentsID(jsonString As String, ByRef idValueContents As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".GetContentsID"


    '検索設定
    Const SearchKeyNameOK   As String = "success"   '検索文字列(コマンド成功時)
    Const SearchKeyNameNG   As String = "error"     '検索文字列(コマンド失敗時)
    Const StartSearchPos    As Long = 10            'あるべき位置


    '1. 文字列を探索
    Dim StartPos_OK As Long, StartPos_NG As Long
    StartPos_OK = InStr(StartSearchPos, jsonString, SearchKeyNameOK)
    StartPos_NG = InStr(StartSearchPos, jsonString, SearchKeyNameNG)

    '2. 2つの結果を見て1つにまとめる
    Dim StartPos As Long
    Dim SearchKeyName As String
    StartPos = IIf(StartPos_OK, StartPos_OK, StartPos_NG)
    SearchKeyName = IIf(StartPos_OK, SearchKeyNameOK, SearchKeyNameNG)

    '3. 所定の位置にない場合は、存在しないとし、抜ける
    '下記の例に基づき、10文字目にない場合は、コマンド結果ではないとして、扱います。
    '・{"type":"success", "id":... （コマンドの成功）
    '・{"type":"error", "id":... （コマンドの失敗）
    If StartPos = StartSearchPos Then
        GetContentsID = True

        '4. 開始位置をうまい具合にずらす
        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""",""id"":") '「","id":」分、ずらすイメージ

        '5. ID値の終点を求める
        Dim EndPos As Long
        EndPos = InStr(StartPos, jsonString, ",")

        '6. ID値を取り出す
        idValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, "This BiDi response is the BiDicommand result. ID:" & idValueContents, FromProcedureName
    Else
        GetContentsID = False
        printMsg Trace_, "This BiDi response does not have an `id`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックし、`method`を高速抽出します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：jsonString             BiDiのレスポンス(JSON文字列そのまま)
'*            methodValueContents  非同期イベントのメソッド名(参照渡し)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：前方一致、後方一致を確認します
'***************************************************************************************************
Private Function GetContentsMethod(jsonString As String, ByRef methodValueContents As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".GetContentsMethod"


    '検索設定
    Const SearchKeyName     As String = "event" '検索文字列
    Const StartSearchPos    As Long = 10        'あるべき位置


    '1. 前方文字列を探索
    Dim StartPos As Long, EndPos As Long
    StartPos = InStr(StartSearchPos, jsonString, SearchKeyName)

    '2. 所定の位置あるか確認する
    '下記の例に基づき、10文字目にない場合は、次に後方一致を確認
    '・{"type":"event","method":...
    If StartPos = StartSearchPos Then
        GetContentsMethod = True

        '3. 開始位置をうまい具合にずらす
        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""",""method"":""") '「","method":"」分、ずらすイメージ

        '4. method名の終点を求める
        EndPos = InStr(StartPos, jsonString, """,")

        '5. method名を取り出す
        methodValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, "This BiDi response is an asynchronous event: " & methodValueContents, FromProcedureName
    
    '下記の例に基づき、後方にない場合は、非同期イベントではないとして、扱います。
    '・{"method":"network.beforeRequestSent", ... (中略) ..., "type":"event"}
    ElseIf Right(jsonString, Len(SearchKeyName) + Len("""}")) = SearchKeyName & """}" Then
        GetContentsMethod = True

        '3. 開始位置をうまい具合にずらす
        StartPos = 1 + Len("{""method"":""") '「{"method":"」分、ずらすイメージ

        '4. method名の終点を求める
        EndPos = InStr(StartPos, jsonString, """,")

        '5. method名を取り出す
        methodValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, "This BiDi response is an asynchronous event: " & methodValueContents, FromProcedureName

    Else
        GetContentsMethod = False
        printMsg Trace_, "This BiDi response does not have an `method`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：コマンドの結果にエラーが含まれているかチェックします
'---------------------------------------------------------------------------------------------------
'* 返り値　：エラー有無
'* 引数　　：ResultsObj    `invokeMethod`の生結果のDictionary
'***************************************************************************************************
Private Function invokeError(ResultsObj As Dictionary) As Boolean
    Const FromProcedureName As String = ThisClassName & ".invokeError"


    Dim ErrorType As String, ErrorMessage As String, ErrorStackTrace As String
    '結果が空っぽになってる場合は、タイムアウトエラーか、コードのミスの可能性
    If ResultsObj Is Nothing Then
        'ログを出しておく
        ErrorType = "error" 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        printMsg WARN_, "InvokeMethod Error: [" & ErrorType & "] " & ErrorMessage, FromProcedureName

        'エラーフラグを立てる
        invokeError = True
        Set m_ErrorDetails = m_JsonConverter.ParseJson("{""type"":" & ErrorType & ",""message"":""" & ErrorMessage & """}")     'カスタムエラーメッセージJsonで返却

    'BiDiコマンドエラーは "type":"error" で返る
    ElseIf ResultsObj("type") = "error" Then
        'エラーを残す
        Set m_ErrorDetails = ResultsObj

        'ログを出しておく
        ErrorType = ResultsObj("error")
        ErrorMessage = ResultsObj("message")
        printMsg WARN_, "BiDi-InvokeMethod Error: [" & ErrorType & "] " & ErrorMessage, FromProcedureName

        '`StackTrace`があったら別枠で表示
        If ResultsObj.Exists("stacktrace") Then
            ErrorStackTrace = ResultsObj("stacktrace")
            printMsg WARN_, "       " & ErrorStackTrace, FromProcedureName
        End If

        'エラーフラグを立てる
        invokeError = True
    End If
End Function

'***************************************************************************************************
'* 機能　　：同期モードでの BiDi-Json コマンド結果のうち、最後のエラー情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後のエラー情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」のエラー情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます。
'            ・適切な`invokeMethod`を実行したつもりなのに返り値が、`Nothing`となるのが、エラーの目印となるので、それでこのエラー情報を見に行ってください。
'            ・`invokeMethod`の`StopError`が`False`だと実質機能しません。
'            ・非同期によるあとで取り出したパターンとは、連動しません。
'***************************************************************************************************
Public Property Get LastBiDiJsonError() As Dictionary
    Set LastBiDiJsonError = m_ErrorDetails
End Property

'***************************************************************************************************
'* 機能　　：非同期で実行されたIDを取り出します
'---------------------------------------------------------------------------------------------------
'* 引数　　：CommandID  非同期で実行した際の返り値(整理券番号)
'            ErrorExist エラー有無(参照渡しにより、実質返り値となります)
'
'* 返値    ：そのIDの実行結果のJson-Dictionary
'---------------------------------------------------------------------------------------------------
'* 注意事項：・事前に、`SendMessage` or `TakeEvents` が実行されてる必要があります
'            ・見つかったIDは、取り出した後、Dictionaryから自動で削除されます
'            ・ここで取り出した結果は、`LastBiDiJsonError`には、反映しません
'            ・結果がない場合は、`Nothing`として返り、`ErrorExist`は、False扱いとします
'***************************************************************************************************
Public Property Get ResultBiDiForAsync(commandID As Long, Optional ErrorExist As Boolean) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".ResultBiDiForAsync"


    '1. 存在するか？
    If currentResultBiDi.response.Exists(commandID) Then
        printMsg Trace_, "BiDi-CommandID: " & commandID & " , Raw-Json:" & currentResultBiDi.response(commandID), FromProcedureName

        '2. Jsonパーサーを施す
        Dim tmp As Dictionary
        Set tmp = m_JsonConverter.ParseJson(currentResultBiDi.response(commandID))

        '3. 成功か？
        If tmp("type") = "success" Then
            '正常に終了していると伝える
            ErrorExist = False
            printMsg Debug_, "This BiDi-command ID(" & commandID & ") has been completed successfully.", FromProcedureName
        Else
            'エラーありとして伝える
            ErrorExist = True
            printMsg WARN_, "This BiDi-command ID(" & commandID & ") is in error. ErrorDetails:" & currentResultBiDi.response(commandID)("error"), FromProcedureName
        End If

        '4. 返却自体は、そのままで
        Set ResultBiDiForAsync = tmp

        '5. 取り出したので消しておく
        currentResultBiDi.response.Remove commandID
        printMsg Trace_, "BiDi-command ID: " & commandID & " has been removed from `AccumulatedAsyncResults`.", FromProcedureName
    Else
        '2. 存在しない旨を伝える
        ErrorExist = False
        printMsg Trace_, "The BiDi command ID(" & commandID & ") passed is not running asynchronously or has been deleted.", FromProcedureName
    End If
End Property

'***************************************************************************************************
'* 機能　　：非同期で実行した、BiDi-Json コマンド結果を蓄積させる上限を決めます
'---------------------------------------------------------------------------------------------------
'* GET返値 ：設定中の上限数
'* LET引数 ：置き換える上限数
'---------------------------------------------------------------------------------------------------
'* 機能説明：`invokeMethodAsync`による結果をあとで、取り出す際の格納上限件数を来ます。デフォルトは、10件です。
'***************************************************************************************************
Property Get ResultAccumulatedLimitBiDiForAsync() As Long
    ResultAccumulatedLimitBiDiForAsync = currentResultBiDi.Limit
End Property

Property Let ResultAccumulatedLimitBiDiForAsync(Number As Long)
    currentResultBiDi.Limit = Number
End Property

'***************************************************************************************************
'* 機能　　：BiDi-Json コマンド結果を得るためのタイムアウト時間を設定します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：設定中のタイムアウト秒数
'* LET引数 ：置き換えるタイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDi-Json コマンド結果を得るのに待つ時間を設定します。デフォルトは、30sです。
'***************************************************************************************************
Property Get TimeOutSecond() As Long
    TimeOutSecond = timerOut
End Property

Property Let TimeOutSecond(TimeSec As Long)
    timerOut = TimeSec
End Property



'***************************************************************************************************
'                           ■■■ 非同期イベントに関する処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザからのイベント内容に関するプロパティを行います
'---------------------------------------------------------------------------------------------------
'* GET返値 ：呼び出した時点での、ブラウザイベント内容
'* SET引数 ：置き換えるDictionary変数。基本は、New Dictionary を渡してのリセットです
'---------------------------------------------------------------------------------------------------
'* 機能説明：・`Set BiDiEvents` に、`New Dictionary`を渡すと、イベント記録モードになります。
'            ・`Set BiDiEvents` に、`Nothing`を渡すと、イベント破棄モードになります。
'* 注意事項：Set プロパティ使用時、簡易的なDictionary構造のチェックを行います。構造が違うDictionaryを渡された場合、警告メッセージが出て、停止します。
'***************************************************************************************************
Property Get BiDiEvents() As Dictionary
    Set BiDiEvents = m_Events
End Property

Property Set BiDiEvents(ObjDic As Dictionary)
    Const FromProcedureName As String = ThisClassName & ".BiDiEvents"


    'SetされたDictionaryの構造をチェックします
    '1. そもそもDictionary オブジェクトが備わっているか？
    If ObjDic Is Nothing Then
        'Nothing 処理。以降は、イベントキャプチャはされません。
        Set m_Events = Nothing
        printMsg info_, "Event capture save data has been discarded. Event capture will not be performed.", FromProcedureName

    '2. まっさらの場合
    ElseIf ObjDic.Count = 0 Then
        '初期化処理。これにより、イベントキャプチャを開始します。
        Set m_Events = FormatDictionary_BiDiEvents
        printMsg info_, "A new event capture save data has been created. Starts event capture.", FromProcedureName

    '3. 指定のKeyがあるか？
    ElseIf ObjDic.Exists(EventsDictionaryKeyName01) And ObjDic.Exists(EventsDictionaryKeyName02) And ObjDic.Count = 2 Then
        '既存データをセット。これにより、このDictionaryにイベントキャプチャの追記が行われます。
        Set m_Events = ObjDic
        printMsg info_, "Existing event capture save data was successfully loaded. Resume event capture with this save data.", FromProcedureName

    '4. 不正なDictionary 形式なのでエラーで止める
    Else
        printMsg ERROR_, "Event capture save data structure is invalid.", FromProcedureName, doRaiseError:=True
    End If
End Property

'***************************************************************************************************
'* 機能　　：BiDiイベント情報を格納するためのフォーマットDictionaryを作成します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：フォーマット済みDictionary
'***************************************************************************************************
Private Function FormatDictionary_BiDiEvents() As Dictionary
    '新規オブジェクトセット
    Set FormatDictionary_BiDiEvents = New Dictionary

    '初期構造を作成します
    FormatDictionary_BiDiEvents.Add EventsDictionaryKeyName01, 0                 '合計イベントキャプチャ数
    FormatDictionary_BiDiEvents.Add EventsDictionaryKeyName02, New Dictionary    'イベントキャプチャDictionaryデータ
End Function

'***************************************************************************************************
'* 機能　　：イベントの「種類（メソッド名）」ごとの仕切りを入れ、そこにCollectionを入れていきます
'---------------------------------------------------------------------------------------------------
'* 実質返値：Accumulator
'* 引数　　：Accumulator        これまでの格納イベント。
'            eventData          Json-Dictionry済みのブラウザからのイベントデータ
'            EventMethodName    イベント名
'---------------------------------------------------------------------------------------------------
'* 機能説明：メソッド名をキーにし、それぞれの中では、時系列が保つように追加していきます。
'            これにより、「特定の"種類"のイベント」を、個別にチェックしやすくなります。
'            時系列データ(Index値)と合計キャプチャ数もおまけとして残しておきます。
'* 注意事項；このプロシージャを呼び出す前に、イベントデータであるかのチェックを済ませてください
'***************************************************************************************************
Private Sub OnNewEventReceived(Accumulator As Dictionary, eventData As Dictionary, EventMethodName As String)
    Const FromProcedureName As String = ThisClassName & ".OnNewEventReceived"


    ' 合計イベントキャプチャ数をカウントする
    Accumulator(EventsDictionaryKeyName01) = Accumulator(EventsDictionaryKeyName01) + 1
    printMsg Trace_, "Now Total of accumulated events :" & Accumulator(EventsDictionaryKeyName01), FromProcedureName

    ' インデックスも付与(1始まり)
    ' これにより、全体の時系列データも保持できます
    eventData.Add "__index__", Accumulator(EventsDictionaryKeyName01)


    '格納開始
    ' 1.イベントメソッドデータがあるDictionaryに遷移
    Dim EventsDictionary As Dictionary: Set EventsDictionary = Accumulator(EventsDictionaryKeyName02)

    ' 2. このイベント名の`Collection`が、すでにあるか？
    Dim eventList As Collection
    If EventsDictionary.Exists(EventMethodName) Then
        ' あれば、そのコレクションデータをセット
        Set eventList = EventsDictionary(EventMethodName)
        printMsg Debug_, "Stored event data from the browser. - " & EventMethodName, FromProcedureName
    Else
        ' なければ、新しい`Collection`をセット
        Set eventList = New Collection
        EventsDictionary.Add EventMethodName, eventList

        printMsg Debug_, "A new type of event. - " & EventMethodName, FromProcedureName
    End If

    ' 3. セットされたコレクションに、今回のイベントの詳細データを、追加する
    eventList.Add eventData
End Sub



'***************************************************************************************************
'                           ■■■ `CDPCore.cls`直通メソッド ■■■
'***************************************************************************************************
Public Function TimerCounter() As Double
    TimerCounter = m_CDPCore.LogControl.GetTicksMs
End Function

Public Sub sleep(Optional seconds As Double = 0.5)
    m_CDPCore.sleep seconds
End Sub



'***************************************************************************************************
'                           ■■■ `CDPTab.cls`直通メソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    WebDriverBiDi_mapperTab.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Public Sub quit()
    WebDriverBiDi_mapperTab.InheritanceCDPBrowser.quit
End Sub



'***************************************************************************************************
'                           ■■■ ワークシートからJS取得 ■■■
'***************************************************************************************************
'* 機能　　：指定のテーブルから、mapperTab.jsコードを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：mapperTab.jsコード
'---------------------------------------------------------------------------------------------------
'* 機能説明：テーブルに格納した分割文字列を結合して返却します
'***************************************************************************************************
Private Function Take_mapperTabJSstring() As String
    Dim tmp2
    Dim chunkNum As Long

    'JS置き場のワークシート制御へ
    With ShLibrary01_JS
        With .ListObjects(.UseTableName(1, "BiDiCore.InsertMapperTab"))
            'データがそもそもない場合は、エラーで止める
            If .ListRows.Count = 0 Then
                Err.Raise vbObjectError, , "mapperTab.js がセットされていません"
                Exit Function
            End If

            '一旦、2次元配列で受け取る
            tmp2 = .DataBodyRange
            chunkNum = UBound(tmp2)
        End With
    End With

    '1次元配列に変換
    ReDim tmp1(1 To chunkNum)
    Dim i As Long
    For i = 1 To chunkNum
        tmp1(i) = tmp2(i, 1)
    Next

    '返却
    Take_mapperTabJSstring = Join(tmp1, "")
End Function



'***************************************************************************************************
'                               ■■■ 更新関連 ■■■
'***************************************************************************************************
'* 機能　　：受け取った`mapperTab.js`のコードを、32,767文字ごとに分割し、テーブルに格納します
'***************************************************************************************************
Private Sub InsertMapperTab(Code As String)
    '1セルごとの格納文字数上限
    Const LimitCellCharacter As Long = 32767

    'テーブル追加時のアニメーション動作を止める
    Application.ScreenUpdating = False


    '1. 分割数を確認
    Dim NumberOfSplit As Long: NumberOfSplit = Len(Code) / LimitCellCharacter

    '2. 配列確保
    ReDim JsCodeChunk(NumberOfSplit, 0)

    '3. 分割して格納する
    Dim i As Long
    For i = 0 To NumberOfSplit
        JsCodeChunk(i, 0) = Mid(Code, LimitCellCharacter * i + 1, LimitCellCharacter)
    Next

    'JS置き場のワークシート制御へ
    With ShLibrary01_JS
        With .ListObjects(.UseTableName(1, "WebDriverBiDiCore.InsertMapperTab"))

            '4. テーブルの中身を全消去
            If .ListRows.Count Then .DataBodyRange.Delete
            
            '5. 挿入範囲の確定
            Dim TargetRange As Range
            Set TargetRange = .ListColumns(1).Range(2).Resize(NumberOfSplit + 1, 1)

            '6. 挿入前に書式を「文字列(@)」に設定
            TargetRange.NumberFormatLocal = "@"

            '7. 挿入
            TargetRange.value = JsCodeChunk
        End With
    End With


    '元に戻す
    Application.ScreenUpdating = True
End Sub

'***************************************************************************************************
'* 機能　　：`mapperTab.js`置き場のExcelテーブルの中身を削除します
'***************************************************************************************************
Private Sub DeleteMapperTab()
    'テーブル削除時のアニメーション動作を止める
    Application.ScreenUpdating = False


    'JS置き場のワークシート制御へ
    With ShLibrary01_JS
        With .ListObjects(.UseTableName(1, "WebDriverBiDiCore.DeleteMapperTab"))
            'テーブルの中身を全消去
            If .ListRows.Count Then .DataBodyRange.Delete
        End With
    End With


    '元に戻す
    Application.ScreenUpdating = True
End Sub

'***************************************************************************************************
'* 機能　　：ローカルファイルから、JSコードを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功論理値
'* 引数　　：FolderPath     フォルダパス
'            FileName       JSファイル名
'---------------------------------------------------------------------------------------------------
'* 機能説明：読み取ったファイルと、Excelテーブルに挿入した文字列との一致確認も行っていきます
'***************************************************************************************************
Public Function UpdateFromLocalFile(FolderPath As String, Optional FileName As String = MapperJSName) As Boolean
    'ファイルから、Jsコードを取得
    '※Windows方式の改行コード(CRLF)だと、`CR`と`LF`での分割が起こって不整合になる可能性があるため、Linux方式(LF)に統一させます
    Dim JsCode As String
    With m_CharConv
        JsCode = Replace(.BytesToString(.BytesFromSavedFile(FolderPath, FileName)), vbCrLf, vbLf)
    End With

    'テーブルに置く作業
    InsertMapperTab JsCode

    '正しく置けてるか確認
    UpdateFromLocalFile = (Take_mapperTabJSstring = JsCode)

    'ミスの場合はテーブルの中身を削除
    If Not (UpdateFromLocalFile) Then DeleteMapperTab

    'ローカルファイルで更新した旨を記録
    '※ミスの場合は、空欄にする
    With ShLibrary01_JS
        .Range(.UseRangeName(1, "WebDriverBiDiCore.UpdateFromLocalFile")).value = IIf(UpdateFromLocalFile, " Local", "")
    End With
End Function

'***************************************************************************************************
'* 機能　　：インターネット経由で、JSコードを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功論理値
'* 引数　　：Version        バージョン指定する場合は、指定してください。デフォルトは最新です
'---------------------------------------------------------------------------------------------------
'* 機能説明：`https://www.jsdelivr.com/`というオープンソースファイル向けの無料のCDN(コンテンツ配信ネットワーク)を利用して、
'            `chromium-bidi`の`mapperTab.js`を更新させます
'***************************************************************************************************
Public Function UpdateFromNPMFile(Optional Version As String = "latest") As Boolean
    Dim ResultResponse As WebResponse

    'リクエストパラメーターを設定する
    Dim リクエスト情報一式 As WebRequest: Set リクエスト情報一式 = New WebRequest
    With リクエスト情報一式
        'GET
        .Method = vbWebHttpGet

        'レスポンスのデータ形式を指定
        .ResponseFormat = vbWebFormatPlainText

        '`chromium-bidi`のnpmパスパラメーターを指定(ホスト名に続くURL)
        .Resource = "npm/" & npmName & "@{Version}/lib/iife/" & MapperJSName
        .AddUrlSegment "Version", Version
    End With

    'リクエストして、結果を返す
    With jsDelivrへリクエスト(リクエスト情報一式, True)
        '200確認
        If .StatusCode = vbWebStatusOk Then
            'テーブルに置く作業
            InsertMapperTab .Content

            '正しく置けてるか確認
            UpdateFromNPMFile = (Take_mapperTabJSstring = .Content)

            'ミスの場合はテーブルの中身を削除
            If Not (UpdateFromNPMFile) Then DeleteMapperTab

        Else
            MsgBox "npm経由による`" & MapperJSName & "`の取得に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & .Content, vbCritical, "ErrorCode:" & .StatusCode
            UpdateFromNPMFile = False
        End If
    End With
End Function

'***************************************************************************************************
'* 機能　　：インターネット経由で、現在のJSコードバージョンを確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：現在の最新バージョン
'---------------------------------------------------------------------------------------------------
'* 機能説明：`https://www.jsdelivr.com/`というオープンソースファイル向けの無料のCDN(コンテンツ配信ネットワーク)を利用して、
'            `chromium-bidi`の`mapperTab.js`を更新させます
'***************************************************************************************************
Public Function UpdateCheckNPMVersion() As String
    Dim ResultResponse As WebResponse

    'リクエストパラメーターを設定する
    Dim リクエスト情報一式 As WebRequest: Set リクエスト情報一式 = New WebRequest
    With リクエスト情報一式
        'GET
        .Method = vbWebHttpGet

        'レスポンスのデータ形式を指定
        .ResponseFormat = vbWebFormatJson

        '`chromium-bidi`のnpmパスパラメーターを指定(ホスト名に続くURL)
        .Resource = "v1/package/npm/" & npmName
    End With

    'リクエストして、結果を返す
    With jsDelivrへリクエスト(リクエスト情報一式, False)
        '200確認
        If .StatusCode = vbWebStatusOk Then
            UpdateCheckNPMVersion = .Data("tags")("latest")
        Else
            MsgBox "npm経由による`" & MapperJSName & "`のバージョン確認に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & .Content, vbCritical, "ErrorCode:" & .StatusCode
        End If
    End With
End Function

'***************************************************************************************************
'* 機能　　：jsDelivr関連のホスト設定です
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスオブジェクト - WebResponse   ここに結果が格納されます
'* 引数　　：クラスオブジェクト - WebRequest    事前に設定したリクエストパラメーター一式
'            GetContents                        `True` :ファイルそのもの用のホストととして
'                                               `False`:バージョン確認用のホストととして
'***************************************************************************************************
Private Function jsDelivrへリクエスト(リクエスト情報一式 As WebRequest, GetContents As Boolean) As WebResponse
    'エンドポイント通信の設定
    Dim EndPointAPI As New WebClient
    With EndPointAPI
        'ホスト名URL
        '実際のファイルか、バージョン確認かの切り替えをします
        .BaseUrl = IIf(GetContents, "https://cdn.jsdelivr.net/", "https://data.jsdelivr.com/")
    End With

    'リクエスト設定とホスト設定を使ってリクエストし、結果を返します
    Set jsDelivrへリクエスト = EndPointAPI.Execute(リクエスト情報一式)
End Function



Private Sub Class_Initialize()
    '------------------------- クラス内変数の一部を初期化 -------------------------
    'タイムアウト設定(デフォルトで、30s)
    timerOut = TIMEOUT_DEFAULT * 1000   'ミリ秒に補正

    '非同期でのコマンド結果を蓄積させる件数(デフォルトで、256件)
    Set currentResultBiDi.response = New Dictionary
    currentResultBiDi.Limit = 2 ^ 8

    '初期化
    Set QueueResultBiDi = New Collection

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        currentDebug.ViewDebugTab = .Range(.UseRangeName(11, "WebDriverBiDiCore.Class_Initialize")).value    '`BiDi-CDP Mapper`のコンソールタブを表示するか？
        currentDebug.seeRawSendMsgDbg = .Range(.UseRangeName(8, "WebDriverBiDiCore.Class_Initialize")).value 'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
        currentDebug.DiagnosticLog = .Range(.UseRangeName(13, "WebDriverBiDiCore.Class_Initialize")).value   'Mapper内部のデバッグログを記録するか？
    End With
End Sub
