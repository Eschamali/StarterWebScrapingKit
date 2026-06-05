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
'                               ■■■ 再接続処理系 ■■■
'***************************************************************************************************
'* 機能　　：Excelテーブルにある既存のBiDiMapperタブ,パイプハンドル情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：userProfile                再アタッチしたいユーザー名(user-data-dirに基づく)
'            sessionCapabilitiesRequest `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'---------------------------------------------------------------------------------------------------
'* 注意事項：・パイプが生きてない場合は、このメソッドから再開できません
'            ・`sessionCapabilitiesRequest`は、新規に`BiDi-CDP Mapper`が起動された時のみ適用されます
'***************************************************************************************************
Public Function reattach(userProfile As String, Optional sessionCapabilitiesRequest As Dictionary) As Boolean
    Const FromProcedureName As String = ThisClassName & ".reattach"


    '1. 「BiDi-CDP Mapper」への再接続処理
    BiDiUserName = userProfile
    Set m_BiDiCore = New WebDriverBiDiCore
    Dim requestSessionNew As Boolean: requestSessionNew = m_BiDiCore.ReBootBiDiCDPMapper(BiDiUserName)

    '2. 結果に応じて`session.new`を行う
    If requestSessionNew Then
         printMsg Debug_, "A new instance of `BiDi-CDP Mapper` has been launched.", FromProcedureName
         sessionNew sessionCapabilitiesRequest
    End If

    '3. セッション一覧を取得して、既存のものがあるか確認する
    ' (BiDiは接続ごとにセッションIDがかわるわけではないが、安全のためセッション状態を確認)
    Dim resultBiDi As Dictionary
    Set resultBiDi = sessionStatus

    '4. mapperTab.js の仕様に沿った確認に従う
    If resultBiDi.Exists("ready") Then
        printMsg info_, "Successfully reattached to BiDi! Session status ready: " & resultBiDi("message"), FromProcedureName
        reattach = True
    Else
        printMsg ERROR_, "Failed to reattach to BiDi session.", FromProcedureName, doRaiseError:=True
        reattach = False
    End If
End Function


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
'                              ■■■ BiDicommand操作関連 ■■■
'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "session.new")
'*           params        引数を持ったDictionary (省略可能)
'*           StopBiDiError 実行失敗時に`Err.Raise`で停止し、以降の処理を止めるか
'* 返値    ：ブラウザから返却されたBiDi結果 (パースされたDictionary)
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを送信後、ID返却されるまで待機します。
'***************************************************************************************************
Public Function ExecuteBiDi(ByVal methodName As String, Optional params As Dictionary, Optional StopBiDiError As Boolean = True) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".ExecuteBiDi"


    Dim paramStr As String
    Dim results As Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = jsConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}")

    'Print error if there is
    If ExecuteError(results) Then
        If StopBiDiError Then
            printMsg ERROR_, "Failed to BiDi-invoke method """ & methodName & """" & vbCrLf & vbCrLf & "=== RawErrorResponse ===" & vbCrLf & jsConverter.ConvertToJson(results), FromProcedureName, doRaiseError:=True
        Else
            printMsg WARN_, "Failed to BiDi-invoke method """ & methodName & """", FromProcedureName
        End If
    Else
        'Return result object(`CDPBrowser`の思想に従い、`Result`オブジェクトで返却します)
        If results.Exists("result") Then Set ExecuteBiDi = results.Item("result")
    End If

    'デバッグ用に、As Is 出力
    If m_BiDiCore.ShowLogBiDiresponse Then printMsg Debug_, "ExecuteBiDi `" & methodName & "` with result: " & jsConverter.ConvertToJson(results), FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (非同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "session.new")
'*           params        引数を持ったDictionary (省略可能)
'* 返値    ：送信したメッセージの ID
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを構成し、JSON化してブラウザへ投げます。結果の待機は行いません。
'***************************************************************************************************
Public Function ExecuteBiDiAsync(ByVal methodName As String, Optional params As Dictionary, Optional StopError As Boolean = True) As Long
    Const FromProcedureName As String = ThisClassName & ".ExecuteBiDiAsync"


    Dim paramStr As String

    If params Is Nothing Then paramStr = "{}" Else paramStr = jsConverter.ConvertToJson(params)
    SendMessage "{""method"":""" & methodName & """,""params"":" & paramStr & "}", True

    '実行したIDを返却
    ExecuteBiDiAsync = m_BiDiCore.LastCommandID
    printMsg Debug_, "Async ExecuteCDP: " & methodName, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：`BiDiコマンド送信のコア処理
'---------------------------------------------------------------------------------------------------
'* 引数　　：strMessage             BiDi形式で送信するJsonプレーンテキスト
'            Async                  `True` ：コマンドの結果を待ちません。Do Loopフェーズには入りません。
'                                   `False`：コマンドの結果を待ちます。Do Loopフェーズに入り、そのIDが来るまで待ちます。この場合、非同期に実行した結果のID回収もついでに行います
'
'            dbgMsg                 `True` ：`CDPCore.cls - TakeEvents`にて、エラーがあったら止まります
'                                   `False`：`CDPCore.cls - TakeEvents`にて、エラーがあっても止まらず、`SendMessage`は`nothing`として返ります
'
'* 返値　　：Dictionary形式のBiDi結果
'---------------------------------------------------------------------------------------------------
'* 注意事項：Async:=True の場合は常に`nothing`で、返ります
'***************************************************************************************************
Private Function SendMessage(strMessage As String, Optional Async As Boolean, Optional dbgMsg As Boolean = True) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".SendMessage"


    '--------------------------------------------------
    'Send message to the browser
    '--------------------------------------------------
    'Pipe関連のエラー停止設定
    m_BiDiCore.StopPipeError = dbgMsg

    printMsg Debug_, "runBiDi  - Async       = " & Async, FromProcedureName

    m_BiDiCore.ReadyRunBiDi strMessage
    Dim currentCommandID As Long: currentCommandID = m_BiDiCore.LastCommandID '`ReadyRunBiDi`直後のcommandIDを取得

    'WriteFile にてエラーがあった場合
    If m_BiDiCore.LastErrorWriteFile Then printMsg ERROR_, "Error WriteFile in m_BiDiCore.writeProcCDP. ErrorCode:" & m_BiDiCore.LastErrorWriteFile & " - " & m_BiDiCore.ErrorDetail.GetMessage(m_BiDiCore.LastErrorWriteFile, "kernel32"), FromProcedureName, doRaiseError:=dbgMsg: Exit Function

    '非同期(結果はいらない)モードなら、ここで抜けます
    If Async Then Exit Function

    '------------------------------------------------
    'Read message returned by browser
    '------------------------------------------------
    'BiDi-Jsonコマンドを送った際のIDが返ってくるまで、ループ受信処理を行いますが、タイムアウトになったら、抜けます
    Dim timerStart As Double: timerStart = TimerCounter
    Dim tmp As String
    printMsg Trace_, "Starting Processes incoming messages from the browser...", FromProcedureName

    Do
        '1. 受信処理によるパイプを拝見する
        m_BiDiCore.TakeEvents

        '2. 受信処理にてエラーがあったら、即抜け
        If m_BiDiCore.LastErrorPeekNamedPipe > 0 Or m_BiDiCore.LastErrorReadFile > 0 Then printMsg ERROR_, "Possible reasons: (1) the browser is not actually running (ie. call string failure); (2) the browser window was accidentally closed; (3) the CDP Pipe connection to the browser was cut-off somehow.", FromProcedureName: Exit Function

        '3. 内部の`RaiseEvent detectionBiDiID`による発火で、Dictionaryへ蓄積されたか？
        tmp = ResultBiDiFromWithEvents(currentCommandID)
        If StrPtr(tmp) Then
            '送ったコマンド結果であることを確認できたのでJsonパース処理して結果を返す
            Set SendMessage = jsConverter.ParseJson(tmp)
            printMsg info_, "WebDriverBiDi have confirmed receipt of command result ID: " & currentCommandID, FromProcedureName

            'ループから抜ける
            Exit Do
        Else
            printMsg Trace_, "WebDriverBiDi have not yet been able to confirm command ID: " & currentCommandID, FromProcedureName
        End If

        '4. 一定時間経ってもコマンド結果が来ないなら、タイムアウトエラーとして終了します
        If TimerCounter - timerStart > timerOut Then printMsg ERROR_, "Timeout waiting for BiDi to return message.", FromProcedureName, , True
        DoEvents
    Loop

    printMsg Trace_, "BiDiCommand result ID:" & currentCommandID & " was returned from BiDi-Json, so the receiving loop will end.", FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：BiDi-Jsonコマンド実行結果のエラーチェックを行い、最後のエラー記録も担います
'---------------------------------------------------------------------------------------------------
'* 返り値　：エラー有無論理値
'* 引数　　：ResultsObj         パース処理済みのCDP-Jsonコマンド実行結果
'***************************************************************************************************
Private Function ExecuteError(ResultsObj As Dictionary) As Boolean
    Const FromProcedureName As String = ThisClassName & ".invokeError"


    Dim ErrorType As String, ErrorMessage As String, ErrorStackTrace As String

    'データの状態をチェック
    If ResultsObj Is Nothing Then
        '結果が空っぽになってる場合は、タイムアウトエラーか、コードのミスの可能性
        ExecuteError = True

        'カスタムなエラー情報として記載し、ログに出す
        ErrorType = "error" 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        printMsg WARN_, "InvokeMethod Error: [" & ErrorType & "] " & ErrorMessage, FromProcedureName

        '最後のエラー情報を記録
        Set m_ErrorDetails = New Dictionary
        m_ErrorDetails.Add "type", ErrorType
        m_ErrorDetails.Add "message", ErrorMessage

    ElseIf ResultsObj("type") = "error" Then
        'BiDiコマンドエラーは "type":"error" で返る
        ExecuteError = True

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
    End If
End Function

'***************************************************************************************************
'* 機能　　：`ExecuteError`メソッドを通した際の、最後のエラー情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後のエラー情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」のエラー情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます
'            ・適切な`ExecuteBiDi`を実行したつもりなのに返り値が、`Nothing`となるのが、エラーの目印となるので、それでこのエラー情報を見に行ってください
'            ・`ExecuteError`メソッドを通さないと機能しません
'            ・`Dictionary`オブジェクトとして得れるので必要に応じて`.RemoveAll`でエラークリアすることも可能です
'***************************************************************************************************
Public Property Get LastBiDiJsonError() As Dictionary
    Set LastBiDiJsonError = m_ErrorDetails
End Property

'***************************************************************************************************
'* 機能　　：`WithEvents`によって蓄積されたコマンドIDを取り出します
'---------------------------------------------------------------------------------------------------
'* 引数　　：CommandID  整理券番号
'* 返値    ：そのIDの実行結果のJson文字列
'---------------------------------------------------------------------------------------------------
'* 注意事項：・事前に、`TakeEvents` が実行されてる必要があります
'            ・見つかったIDは、取り出した後、Dictionaryから自動で削除されます
'            ・結果がない場合は、`vbnullstring`として返ります。`StrPtr`で判定可能です
'            ・エラー判定、JsonDictionay化は別途行って下さい
'***************************************************************************************************
Property Get ResultBiDiFromWithEvents(commandID As Long) As String
    Const FromProcedureName As String = ThisClassName & ".ResultBiDiFromWithEvents"


    '1. 存在するか？
    If resultBiDi.response.Exists(commandID) Then
        '2. Log化
        Dim tmp As String: tmp = resultBiDi.response(commandID)
        printMsg info_, "Browser were able to confirm the existence of BiDi command ID:" & commandID & "　→　" & tmp, FromProcedureName

        '3. 返却
        ResultBiDiFromWithEvents = tmp

        '4. 取り出したので消しておく
        resultBiDi.response.Remove commandID
        printMsg Trace_, "CDP-command ID: " & commandID & " has been removed from `ResultBiDi.response`.", FromProcedureName
    Else
        '2. 存在しない旨を伝える
        printMsg Trace_, "This BiDi-command ID(" & commandID & ") passed is not running asynchronously or has been deleted.", FromProcedureName
    End If
End Property

'***************************************************************************************************
'* 機能　　：BiDi-Json コマンド結果を蓄積させる上限を決めます
'---------------------------------------------------------------------------------------------------
'* LET引数 ：置き換える上限数
'---------------------------------------------------------------------------------------------------
'* 機能説明：格納上限件数を設定します。デフォルトは、256件です。
'***************************************************************************************************
Property Let SetLimitBiDi(Number As Long)
    resultBiDi.Limit = Number
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
