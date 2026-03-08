VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverBiDiCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'               Automating Chromium-Based Browsers with WebDriverBiDi API and VBA
'--------------------------------------------------------------------------------------------------------------
' WebDriver BiDiの低レベル通信を担うクラス。
' CDPBrowser経由でBiDiマッパーを実行し、BiDiコマンドの送受信・イベント取得を行います。
' なるべく、`CDPBrowser`の大幅な変更を行わずに、`CDPBrowser`←→`BiDiCore`の行き来で、実現していきます
' 処理ロジックはほとんど、`CDPBrowser`を踏襲してます
'==============================================================================================================
Option Explicit



'内部参照用オブジェクト
Private m_CDPBrowser    As CDPBrowser
Private m_CharConv      As New CharacterCodeConversion
Private m_JsonConverter As New WebJsonConverter

'ステート管理
Private m_MessageID                 As Long         'BiDiコマンド結果を確認するためのID管理
Private m_Events                    As Dictionary   'BiDi向けの非同期イベント蓄積
Private m_AccumulatedCommandResults As Dictionary   '非同期で、BiDiコマンドを行った際の結果蓄積Dictionary
Private m_MaxCommandResults         As Long         '非同期で、BiDiコマンドを行った際の結果蓄積の上限
Private m_ErrorDetails              As Dictionary   '同期BiDiコマンドした際の最後のエラー情報
Private timerOut                    As Long         '同期BiDiコマンドした際のタイムアウト設定
Private ViewDebugTab                As Boolean      '「BiDi-CDP Mapper is controlling this tab」といったデバックタブ画面を表示するか？※非表示でも、「edge://inspect/#other」で存在確認はできます。

Private Const ThisClassName As String = "BiDiCore"  'ログのトレース用

'BiDiからのイベント情報を格納する際のDictionaryKey名設定
Private Const EventsDictionaryKeyName01 As String = "TotalEvents"
Private Const EventsDictionaryKeyName02 As String = "EventMethods"



'***************************************************************************************************
'                              ■■■ BiDiセッション確保関連 ■■■
'***************************************************************************************************
'* 機能　　：初期化処理
'---------------------------------------------------------------------------------------------------
'* 引数　　：StartURL                   最初の起動時に遷移したいURL
'*           SwitchUser                 設定シートとは別のユーザディレクトリを指定するときに使います
'            sessionCapabilitiesRequest `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'---------------------------------------------------------------------------------------------------
'* 機能説明：・BiDiマッパーのスクリプトをExcelテーブルから読み込み、対象のコンテキストで実行可能にします。
'            ・新規立ち上げ用に使い、セッション確立まで担います。
'***************************************************************************************************
Public Sub start(Optional StartURL As String, Optional SwitchUser As String, Optional sessionCapabilitiesRequest As Dictionary)
    Const FromProcedureName As String = ThisClassName & ".start"


    '1. セルから設定を適用
    Set m_CDPBrowser = 設定シートからの起動(StartURL, SwitchUser)

    '2. BiDi-CDP Mapper を始動させます
    Call StartBiDiCDPMapper

    '3. BiDi との新しいセッションを確立する
    '※ここは、BiDi用のコマンドとして使用
    Dim resultBiDi As Dictionary

    '`session.new`の`Parameters`が省略されてる場合は、規定値を作成する
    If sessionCapabilitiesRequest Is Nothing Then
        Set sessionCapabilitiesRequest = New Dictionary
        sessionCapabilitiesRequest.Add "capabilities", New Dictionary
    End If

    '確立処理！
    Set resultBiDi = invokeMethod("session.new", sessionCapabilitiesRequest)
    ViewNewSession resultBiDi
End Sub

'***************************************************************************************************
'* 機能　　：既存のBiDiセッションに再接続します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功可否論理値
'* 引数　　：userProfile    再アタッチしたいユーザー名(user-data-dirに基づく)
'---------------------------------------------------------------------------------------------------
'* 機能説明：あとから、再接続する際に使い、セッション確立まで担います。
'***************************************************************************************************
Public Function reattach(userProfile As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".reattach"


    '1. CDPセッションに再接続
    Dim ResultConnectionCDP As Boolean
    Set m_CDPBrowser = New CDPBrowser
    ResultConnectionCDP = m_CDPBrowser.reattach(userProfile)

    '失敗時は、ここで抜ける
    '※大抵は、デバックブラウザ自体がない
    If Not (ResultConnectionCDP) Then reattach = False: Exit Function

    '2. BiDiセッションに再接続
    With m_CDPBrowser
        '3. すでに当該タブで BiDi マッパーが起動済みかを確認する
        '→ typeof window.onBidiMessage === 'function' かどうかを調べる
        Dim jsCheck As String
        jsCheck = "typeof window.onBidiMessage === 'function';"

        Dim isMapperRunning As Boolean
        isMapperRunning = .jsEval(jsCheck)

        If isMapperRunning Then
            printMsg Debug_, "Mapper instance is already running. Reattaching to existing BiDi session...", FromProcedureName

            '4. `CDPBrowser`に、BiDiモードであると伝える
            Set .BiDiCapture = New Collection
        Else
            printMsg Debug_, "Mapper instance not found. Injecting `BiDi-CDP Mapper` into reattached tab...", FromProcedureName

            '4. BiDi-CDP Mapper を再始動させます
            Call StartBiDiCDPMapper

            printMsg Debug_, "`BiDi-CDP Mapper` has been launched on the reattached tab", FromProcedureName
        End If

        '5. セッション一覧を取得して、既存のものがあるか確認する
        ' (BiDiは接続ごとにセッションIDがかわるわけではないが、安全のためセッション状態を確認)
        Dim resultBiDi As Dictionary
        Set resultBiDi = invokeMethod("session.status")

        '6. mapperTab.js の仕様に沿った確認に従う
        If resultBiDi.Exists("ready") Then
            printMsg Debug_, "Successfully reattached to BiDi! Session status ready: " & resultBiDi("message"), FromProcedureName
            reattach = True
        Else
            printMsg ERROR_, "Failed to reattach to BiDi session.", FromProcedureName
            reattach = False
        End If
    End With
End Function

'***************************************************************************************************
'* 機能　　：新規タブにて、BiDi-CDP Mapperを始動させます
'---------------------------------------------------------------------------------------------------
'* 注意事項：`CDPBrowser`を`New`してからご利用ください
'***************************************************************************************************
Private Sub StartBiDiCDPMapper()
    Const FromProcedureName As String = ThisClassName & ".StartBiDiCDPMapper"


    With m_CDPBrowser
        '1. 新規タブに対して、BiDiコンソールタブを準備し、セッションターゲットをメイン化
        '※タブ表示モードの際は、そのタブがフォアグラウンドにならないように仕向けます
        .newTab , setMain:=True, isHidden:=Not (ViewDebugTab), isBackground:=True
        printMsg Trace_, "Booting `BiDi-CDP Mapper`... (ViewTab: " & CStr(ViewDebugTab) & ")", FromProcedureName

        '2. `CDPBrowser`に、BiDiモードであると伝える
        Set .BiDiCapture = New Collection


        '3. BiDi用バインディングの追加
        Dim paramsCDP As New Dictionary
        paramsCDP.Add "bindingName", "cdp"
        paramsCDP.Add "targetId", .CurrentTargetID
        paramsCDP.Add "inheritPermissions", True
        .invokeMethod "Target.exposeDevToolsProtocol", paramsCDP, True

        '4. JSからVBAへメッセージを送るための関数名を "sendBidiResponse" に設定
        paramsCDP.RemoveAll
        paramsCDP.Add "name", "sendBidiResponse"
        .invokeMethod "Runtime.addBinding", paramsCDP

        '5. パースしたJavaScriptをブラウザ上で実行
        .jsEval Take_mapperTabJSstring

        '6. ブートストラップ（初期化）JSの実行
        .jsEval "window.runMapperInstance('" & .CurrentTargetID & "')"
        printMsg Debug_, "`BiDi-CDP Mapper` has been launched", FromProcedureName
    End With

    printMsg Trace_, "Booted `BiDi-CDP Mapper`", FromProcedureName
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
    printMsg Debug_, "BiDi-Invoked """ & methodName & """ with result: " & m_JsonConverter.ConvertToJson(results), FromProcedureName
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


    m_MessageID = m_MessageID + 1
    Dim lngCurrentId As Long
    lngCurrentId = m_MessageID

    '末尾の「}」を削って id を付与する
    strMessage = Left(strMessage, Len(strMessage) - 1)
    strMessage = strMessage & ", ""id"":" & lngCurrentId & "}"

    printMsg Debug_, "sendBiDi-JsonCommand = " & strMessage, FromProcedureName

    'JSONのダブルクォートやエスケープ処理による構文乱れを防ぐためエスケープ
    Dim safeJson As String
    safeJson = Replace(strMessage, "\", "\\")
    safeJson = Replace(safeJson, "'", "\'")

    Dim paramsCDP As New Dictionary
    paramsCDP.Add "expression", "window.onBidiMessage('" & safeJson & "');"

    'CDP経由でBiDiコマンド送信
    '※JavaScript自体は同期実行。エラー時は、内部でStopがかかります
    Dim cdpResult As Dictionary
    Set cdpResult = m_CDPBrowser.invokeMethod("Runtime.evaluate", paramsCDP)

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
    Dim timerStart As Double: timerStart = Timer

    'CDP受信処理→Collection内`Runtime.bindingCalled`処理→受信処理→Collection内`Runtime.bindingCalled`処理→…　というロジックでループ処理を行います
    'BiDi-Jsonコマンドを送った際のIDが返ってくるまで、ループ受信処理を行いますが、タイムアウトになったら、抜けます
    Dim ResultSendCommand As Boolean
    printMsg Trace_, "Starting Processes incoming messages from the browser...", FromProcedureName
    Do
        'CDPイベント受信処理
        m_CDPBrowser.TakeEvents

        '受信の過程で入った、`BiDi`用の受信コレクションを取得
        Dim CurrentBiDiCol As Collection
        Set CurrentBiDiCol = m_CDPBrowser.BiDiCapture

        '受信数check
        If CurrentBiDiCol.Count Then
            'データあり。受信メッセージを活用するため、担当プロシージャへ遷移
            Set SendMessage = ReceiveFromCollectionPayload(CurrentBiDiCol, InternalUseMessages, lngCurrentId)

            'オブジェクトありなら、直近に送ったコマンド結果であることを確認できたので、受信ループ終了フラグを立てる
             ResultSendCommand = (Not (SendMessage Is Nothing))
        Else
            printMsg Trace_, "I used `m_CDPBrowser.BiDiCapture` but it is empty.", FromProcedureName
        End If


        '一定時間経ってもコマンド結果が来ないなら、タイムアウトエラーとして終了します
        If Timer - timerStart > timerOut Then printMsg ERROR_, "Timeout waiting for BiDi to return message.", FromProcedureName, , True
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
        m_CDPBrowser.TakeEvents

        '受信の過程で入った、`BiDi`用の受信コレクションを取得
        Dim CurrentBiDiCol As Collection
        Set CurrentBiDiCol = m_CDPBrowser.BiDiCapture

        BiDiNum = CurrentBiDiCol.Count
        printMsg Trace_, "BiDiNum = " & BiDiNum, FromProcedureName

        '受信数check
        If BiDiNum Then
            '破棄モードかの判定
            If destruction Then
                '解析などせず、そのまんま出して破棄します
                '※Collectionのため、中身は直接見れません
                printMsg Trace_, "Discarding data in BiDiCapture...", FromProcedureName
            Else
                '中身の解析をし、IDやイベントを取り出します
                printMsg Trace_, "Analyzing the data in the BiDiCapture...", FromProcedureName
                Call ReceiveFromCollectionPayload(CurrentBiDiCol)
            End If
        Else
            printMsg Trace_, "Message queue destruction completed.", FromProcedureName
        End If
    Loop Until BiDiNum = 0

    printMsg Debug_, "Finished getting data from the BiDiCapture.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：`CDPBrowser.BiDiCapture`に格納してあるコレクションを、1件1件処理し、状況に応じて内部で、イベントキャプチャしたり、コマンド結果の確認を行います
'---------------------------------------------------------------------------------------------------
'* 返値　　：`checkCommandID`に対するCDP-Jsonコマンド結果Dictionary形式
'* 引数　　：ReceiveRuntimebindingCalledCol     `CDPBrowser.BiDiCapture`
'            InternalUseMessages                非同期イベントをクラス内部で後程、使用したいときに、`New Dictionry`として、渡して有効化します。
'            checkCommandID                     確認するコマンドID
'---------------------------------------------------------------------------------------------------
'* 機能説明：`checkCommandID`に合わないIDが来ても`AccumulatedCommandResults`に蓄積されます
'***************************************************************************************************
Private Function ReceiveFromCollectionPayload(ByRef ReceiveRuntimebindingCalledCol As Collection, Optional ByRef InternalUseMessages As Dictionary, Optional ByRef checkCommandID As Long = -1) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".ReceiveFromCollectionPayload"


    Dim resBiDi
    Dim i As Long

    printMsg Trace_, "Number of BiDi receptions:" & ReceiveRuntimebindingCalledCol.Count, FromProcedureName
    For Each resBiDi In ReceiveRuntimebindingCalledCol
        'ログ用に、カウント
        i = i + 1
        printMsg Trace_, "Raw - BiDiRes(" & i & ") = " & resBiDi, FromProcedureName

        '【最適化】CDPのラッパーJSONをパースせず、文字列操作でpayloadを直接抜き出す（JsonOfJsonの1回目を回避）
        Dim payloadBiDi As String
        payloadBiDi = ExtractPayload(resBiDi)
        printMsg Debug_, "BiDiRes(" & i & ")-`payload` = " & payloadBiDi, FromProcedureName

        '空文字はそもそも受け付けないようにする
        If payloadBiDi <> "" Then
            '受信データのチェックを行うプロシージャへ遷移します。
            If BrowserReceivedDataCheck(payloadBiDi, InternalUseMessages, checkCommandID) Then
                'ここでcheckCommandID一致時のみ、パース（BiDi本体）を行う。
                Set ReceiveFromCollectionPayload = m_JsonConverter.ParseJson(CStr(payloadBiDi))
                printMsg Debug_, "The BiDicommand result was obtained from `Runtime.bindingCalled`! - ID:" & checkCommandID, FromProcedureName
            End If
        Else
            printMsg Debug_, "The payload is empty", FromProcedureName
        End If

        DoEvents
    Next
End Function

'***************************************************************************************************
'* 機能　　：`Runtime.bindingCalled` のCDPメッセージから payload 部分を高速に抜き出します
'***************************************************************************************************
Private Function ExtractPayload(ByVal resCDP As String) As String
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
                '多分、非同期で実行したコマンド結果と思われる
                'VBAによるシングルスレッドの制約上、上記以外のズレはありえないとおもう...
                FlagValue = FlagValue Or mismatchID '現在のコマンド結果IDと返ったコマンド結果IDが異なる場合のフラグ付け
                printMsg info_, "The BiDicommand result (ID:" & ResultCommandID & ") was returned, but it does not match the BiDicommand ID (" & checkCommandID & ") recorded in Excel. This is probably an asynchronous result. The results will be accumulated.", FromProcedureName

                'コマンド結果ID用受信Boxへ格納しておく
                '※パーサー処理負荷を抑えるため、Json文字列まんまとします
                m_AccumulatedCommandResults.Add ResultCommandID, BiDiJsonString

                '格納件数チェックし、上限を超えたら、最古データを削除します
                If m_AccumulatedCommandResults.Count > m_MaxCommandResults Then
                    ' 一番古いキーを取得して削除
                    m_AccumulatedCommandResults.Remove m_AccumulatedCommandResults.keys(0)
                    printMsg info_, "The accumulation limit has exceeded " & m_MaxCommandResults & ", so the oldest has been deleted.", FromProcedureName
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
        OnNewEventReceived m_Events, JsonDictionaryObj
    End If

    'プライベート(クラス内変数)用
    If InternalUseMessages Is Nothing Then
        printMsg Trace_, "No event information needed inside the class.", FromProcedureName
    Else
        printMsg Trace_, "The class is requesting the accumulation of event information. Saving...", FromProcedureName
        OnNewEventReceived InternalUseMessages, JsonDictionaryObj
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
        StartPos = StartSearchPos + Len("{""method"":""") '「{"method":"」分、ずらすイメージ

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


    'BiDiコマンドエラーは "type":"error" で返る
    Dim ErrorType As String, ErrorMessage As String, ErrorStackTrace As String
    If ResultsObj("type") = "error" Then
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
'***************************************************************************************************
Public Property Get ResultBiDiForAsync(CommandID As Long, ErrorExist As Boolean) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".ResultBiDiForAsync"


    '1. 存在するか？
    If m_AccumulatedCommandResults.Exists(CommandID) Then
        printMsg Debug_, "CommandID - " & CommandID & ",Raw-Json:" & m_AccumulatedCommandResults(CommandID), FromProcedureName

        '2. Jsonパーサーを施す
        Dim tmp As Dictionary
        Set tmp = m_JsonConverter.ParseJson(m_AccumulatedCommandResults(CommandID))

        '3. 成功か？
        If tmp("type") = "success" Then
            '正常に終了していると伝える
            ErrorExist = False
            printMsg Debug_, "This BiDi-command ID(" & CommandID & ") has been completed successfully.", FromProcedureName

            Set ResultBiDiForAsync = tmp("result")  '`result`オブジェクトとして返却
        Else
            'エラーありとして伝える
            ErrorExist = False
            printMsg WARN_, "This BiDi-command ID(" & CommandID & ") is in error.", FromProcedureName

            Set ResultBiDiForAsync = tmp            'そのまんま返却
        End If

        '4. 取り出したので消しておく
        m_AccumulatedCommandResults.Remove CommandID
        printMsg Trace_, "Deleted the resulting log corresponding to the execution BiDi-command ID. - ID: " & CommandID, FromProcedureName
    Else
        ErrorExist = True
        printMsg WARN_, "There is no target command ID result log. - ID: " & CommandID, FromProcedureName

        '2. カスタムエラーメッセージJsonで返却
        Set ResultBiDiForAsync = m_JsonConverter.ParseJson("{""type"":""error"",""id"":" & CommandID & ",""error"":-2,""message"":""There is no target command ID result log.""}")
    End If
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
'* 引数　　：Accumulator    これまでの格納イベント。
'            eventData      Json-Dictionry済みのブラウザからのイベントデータ
'---------------------------------------------------------------------------------------------------
'* 機能説明：メソッド名をキーにし、それぞれの中では、時系列が保つように追加していきます。
'            これにより、「特定の"種類"のイベント」を、個別にチェックしやすくなります。
'            時系列データ(Index値)と合計キャプチャ数もおまけとして残しておきます。
'* 注意事項；このプロシージャを呼び出す前に、イベントデータであるかのチェックを済ませてください
'***************************************************************************************************
Private Sub OnNewEventReceived(Accumulator As Dictionary, eventData As Dictionary)
    Const FromProcedureName As String = ThisClassName & ".OnNewEventReceived"


    'イベント名を取得
    Dim EventMethodName As String: EventMethodName = eventData("method")

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
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
Public Sub sleep(Optional seconds As Double = 0.5)
    m_CDPBrowser.sleep seconds
End Sub

Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    m_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Public Sub quit()
    m_CDPBrowser.quit
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
Function Take_mapperTabJSstring() As String
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
        With .ListObjects(.UseTableName(1, "BiDiCore.InsertMapperTab"))

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
'* 機能　　：ローカルファイルから、JSコードを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功論理値
'* 引数　　：FolderPath     フォルダパス
'            FileName       JSファイル名
'---------------------------------------------------------------------------------------------------
'* 機能説明：読み取ったファイルと、Excelテーブルに挿入した文字列との一致確認も行っていきます
'***************************************************************************************************
Public Function FromLocalFile(FolderPath As String, Optional FileName As String = "mapperTab.js") As Boolean
    'ファイルから、Jsコードを取得
    Dim JsCode As String
    With m_CharConv
        JsCode = .BytesToString(.BytesFromSavedFile(FolderPath, FileName))
    End With

    'テーブルに置く作業
    InsertMapperTab JsCode

    '正しく置けてるか確認
    FromLocalFile = (Take_mapperTabJSstring = JsCode)
End Function



Private Sub Class_Initialize()
    '------------ いくつかの初期設定をもうけます ------------
    'タイムアウト設定(デフォルトで、10s)
    timerOut = 10

    '非同期でのコマンド結果を蓄積させる件数
    m_MaxCommandResults = 10
    Set m_AccumulatedCommandResults = New Dictionary

    'コマンドIDリセット
    m_MessageID = 0

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        ViewDebugTab = .Range(.UseRangeName(11, "BiDiCore.Class_Initialize")).value
    End With
End Sub
