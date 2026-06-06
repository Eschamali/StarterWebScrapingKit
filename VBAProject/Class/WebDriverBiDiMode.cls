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
Attribute m_BiDiCore.VB_VarHelpID = -1

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
'                              ■■■ タブ操作関連 ■■■
'***************************************************************************************************
'* 機能　　：新規タブを返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - WebDriverBiDiContext
'* 引数　　：newWindow      `True` ：新しいウィンドウを開く新規タブとして扱います
'                           `False`：既存のウィンドウにて、タブを生成します
'
'            isBackground   `True` ：バックグラウンド（アクティブにしない）で開く
'                           `False`：アクティブで開く
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebDriverBiDi では、いきなりURLで直遷移はできないようです
'***************************************************************************************************
Public Function newTab(Optional newWindow As Boolean, Optional isBackground As Boolean) As WebDriverBiDiContext
    Const FromProcedureName As String = ThisClassName & ".newTab"


    ' 1. 指定したパラメーターで実行
    Dim bidiRes As Dictionary
    Set bidiRes = SendMessage("{""method"":""browsingContext.create"",""params"":{""type"":""" & IIf(newWindow, "window", "tab") & """,""background"":" & LCase(isBackground) & "}}")

    '2. 返却準備
    Set newTab = New WebDriverBiDiContext
    With newTab
        '1. このBiDi情報を渡す
        Set .InheritanceWebDriverBiDiMode = Me

        '2. 先ほど得た`context`を渡す
        .context = bidiRes("result")("context")
    End With

    printMsg info_, "newTab will now open a new BiDi instance", FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：既に開いているタブのセッションIDを、タブの名前またはURLに基づいて検索します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - WebDriverBiDiContext
'* 引数　　：tabName    タブ名
'            Url        URL
'            maxDepth   Tree の深さレベル。基本は、0にすると子フレーム(iframe)を含まず、トップレベルのタブだけ取得となります
'            doRetry    `True`で、出現まで何回かリトライします。
'---------------------------------------------------------------------------------------------------
'* 機能説明：tabName引数とurl引数は部分的な指定(部分一致)が可能です。maxDepth はデフォルトは、トップレベルのタブだけにしてます
'* 注意事項：・url引数を両方とも省略すると、getTabは最も近い未接続のタブに接続しようとします。
'            ・WebDriverBiDi では、タブ名での検索はできないようです
'***************************************************************************************************
Public Function getTab(Optional Url As String, Optional maxDepth As Long, Optional doRetry As Boolean) As WebDriverBiDiContext
    Const FromProcedureName As String = ThisClassName & ".getTab"


    Dim startTime As Double
    Dim bidiRes As Dictionary
    Dim contexts As Collection
    Dim matchData As Dictionary

    If doRetry Then startTime = TimerCounter

    'シンプルに、`Instr`方式の検索で行う
    Do
        ' 1. 指定した深さまでのツリーを取得
        Set bidiRes = SendMessage("{""method"":""browsingContext.getTree"",""params"":{""maxDepth"":" & maxDepth & "}}")

        If Not bidiRes Is Nothing Then
            Set contexts = bidiRes("result")("contexts")

            ' 2. 再帰関数を呼び出して全階層からマッチするものを探す
            Set matchData = FindContextRecursive(contexts, Url)

            ' 3. 見つかったらオブジェクト化して返す
            If Not matchData Is Nothing Then
                Set getTab = New WebDriverBiDiContext
                With getTab
                    '1. このBiDi情報を渡す
                    Set .InheritanceWebDriverBiDiMode = Me

                    '2. 先ほど得た`context`を渡す
                    .context = matchData("context")
                End With

                Exit Do
            End If
        End If

        If Not doRetry Then Exit Do
        If TimerCounter - startTime > timerOut Then Exit Do
        DoEvents
    Loop
End Function

'***************************************************************************************************
'* 機能　　：`browsingContext.getTree`の再帰処理による検索処理です
'---------------------------------------------------------------------------------------------------
'* 返り値　：contexts   browsingContext.InfoList
'* 引数　　：Url        URL
'---------------------------------------------------------------------------------------------------
'* 注意事項：再帰処理のため、深すぎるとエラーになりますが多分問題ないはずです
'***************************************************************************************************
Private Function FindContextRecursive(ByVal contexts As Collection, ByVal Url As String) As Dictionary
    Dim ctxData As Object
    Dim childMatch As Object
 
    ' contextsはCollection（配列）としてループ
    For Each ctxData In contexts
        
        ' A. このコンテキスト自体がマッチするかチェック
        Dim isMatch As Boolean
        If Url = "" Then
            isMatch = True ' 両方空なら最初に見つかったヤツ
        Else
            If Url <> "" And InStr(1, ctxData("url"), Url) > 0 Then isMatch = True
        End If

        If isMatch Then
            Set FindContextRecursive = ctxData
            Exit Function
        End If

        ' B. 子要素（children）があれば、さらにその中を探索
        If Not ctxData("children") Is Nothing Then
            If ctxData("children").Count > 0 Then
                Set childMatch = FindContextRecursive(ctxData("children"), Url)
                If Not childMatch Is Nothing Then
                    Set FindContextRecursive = childMatch
                    Exit Function
                End If
            End If
        End If
    Next

    Set FindContextRecursive = Nothing
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
    printMsg Debug_, "runBiDi  - Async       = " & Async, FromProcedureName

    m_BiDiCore.ReadyRunBiDi strMessage, dbgMsg                                  '`Runtime.evaluate`による、BiDi → CDP 変換実行時の内部エラー無視設定も指定して、両者を釣り合わせる
    Dim currentCommandID As Long: currentCommandID = m_BiDiCore.LastCommandID   '`ReadyRunBiDi`直後のcommandIDを取得

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
        m_BiDiCore.TakeEvents dbgMsg

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
    Loop

    printMsg Trace_, "BiDiCommand result ID:" & currentCommandID & " was returned from BiDi-Json, so the receiving loop will end.", FromProcedureName
End Function



'***************************************************************************************************
'                              ■■■ BiDiCommand結果関連 ■■■
'***************************************************************************************************
'* 機能　　：WebDriverBiDiの実行結果を一旦、Dictionaryへ蓄積させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：id             コマンドID
'            RawJson        コマンド結果の生Json文字列　※As Is. ID付きのままで渡されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：`ID`をKeyに、Dictionaryへ`RawJson`を格納します。
'* 注意事項：・Jsonは、文字列のままで渡されます。
'            ・Dictionaryへの蓄積には、上限があります。上限を超えると、結果履歴(特に、`ExecuteBiDiAsync`の結果)が全て削除されます
'***************************************************************************************************
Private Sub m_BiDiCore_detectionBiDiID(id As Long, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".m_BiDiCore_detectionBiDiID"


    '1. 格納件数チェックし、次の蓄積で上限を超えたら、全履歴を削除します
    If resultBiDi.response.Count + 1 > resultBiDi.Limit Then
        resultBiDi.response.RemoveAll
        printMsg WARN_, "The accumulation limit has exceeded " & resultBiDi.Limit & ", so All result history has been deleted to prevent performance degradation.", FromProcedureName
    End If

    '2. その後、結果を蓄積
    resultBiDi.response.Add id, RawJson
    printMsg Trace_, "The results of the WebDriverBiDi(ID: " & id & ") execution from the browser have been received. → " & RawJson, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：オーバーフロー対策によるcommandIDがリセットされると呼び出されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：IDが20億に達する場合も、結果履歴が全て削除されます
'***************************************************************************************************
Private Sub m_BiDiCore_ResetCommandID()
    resultBiDi.response.RemoveAll
End Sub

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
'                                       ■■■ 後始末系 ■■■
'***************************************************************************************************
Public Sub quit()
    Const FromProcedureName As String = ThisClassName & ".quit"


    '1. 終了コマンドを送る
    '※閉じると同時に、Pipeエラーが起こることがあるので、エラー無視しておく
    SendMessage "{""method"":""browser.close"",""params"":{}}", dbgMsg:=False

    '2. パイプハンドル、ブラウザセッション情報を消去しに行きます
    m_BiDiCore.ClosePipeCDP 'パイプハンドルを閉じる
    deleteCDPID             'テーブルからも削除する

    '3. it seems without waiting a bit the browser crashes and the next time it wants to recover from a crash
    sleep 1
    printMsg info_, "The WebDriverBiDi has been terminated.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：Excelのテーブルから、ブラウザへの接続パイプとタブ情報を削除します
'***************************************************************************************************
Private Sub deleteCDPID()
    Const FromProcedureName As String = ThisClassName & ".deleteCDPID"


    With ShSetting01_StartBrowser
        .DeleteCDPToTable(1, FromProcedureName) = BiDiUserName  'パイプ通信Handle情報
        .DeleteCDPToTable(2, FromProcedureName) = BiDiUserName  'ブラウザセッション情報
    End With
End Sub



'***************************************************************************************************
'                                ■■■ 非同期イベント関連 ■■■
'***************************************************************************************************
'* 機能　　：WebDriverBiDiに対する非同期イベントを処理します
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName 非同期イベントメソッド名
'            RawJson    非同期イベントの生Json文字列　※As Is. メソッド名付きのままで渡されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：非同期イベントキャプチャを行います
'* 注意事項：WebDriverBiDi の場合は、フィルター機能が内蔵済みです
'***************************************************************************************************
Private Sub m_BiDiCore_detectionBiDiEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".m_BiDiCore_detectionBiDiEvent"


    '------------------ 非同期イベントを収集します ------------------
    '1. イベントキャプチャしない場合は、即抜け
    If m_Events Is Nothing Then printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName: Exit Sub

    '2. パース処理
    printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
    Dim tmp As Dictionary: Set tmp = jsConverter.ParseJson(RawJson)

    '3. 専用のプロシージャにて、蓄積処理
    printMsg Trace_, "The module is requesting the accumulation of event information. Saving...", FromProcedureName
    OnNewEventReceived m_Events, tmp, methodName
End Sub



'***************************************************************************************************
'                                ■■■ 非同期イベント収集関連 ■■■
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
'                                   ■■■ ステータス系 ■■■
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
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
End Sub



'***************************************************************************************************
'                            ■■■ Session系統コマンド ■■■
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
    Set sessionNew = ExecuteBiDi("session.new", params)
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
    Set sessionStatus = ExecuteBiDi("session.status")
End Function



'***************************************************************************************************
'                              ■■■ PUBLIC PROPERTIES ■■■
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
'                              ■■■ Cheat PROPERTIES ■■■
'***************************************************************************************************
'* 機能　　：このClass内で定義してる`CDPBrowser.cls`を受け取り、WebDriverBiDiモードでありつつ、直接CDP制御できる仕組みを提供します
'---------------------------------------------------------------------------------------------------
'* 注意事項：利用する際は、「BiDi-CDP Mapper」タブ自体を操作しないように気を付ける必要がある上級者向けです
'***************************************************************************************************
Property Get DirectControlCDP() As CDPBrowser
    Set DirectControlCDP = m_BiDiCore.DirectControlCDP
End Property



'***************************************************************************************************
'                        ■■■ WebDriverBiDiCore.cls 直通メソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    m_BiDiCore.printMsg LogLevel_, LogID, strMsg, From, isHeader, doRaiseError
End Sub

Public Sub sleep(Optional seconds As Double = 0.5)
    m_BiDiCore.sleep seconds
End Sub

Public Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    m_BiDiCore.TakeEvents StopPipeError, destruction
End Sub

Public Function TimerCounter() As Double
    TimerCounter = m_BiDiCore.TimerCounter
End Function



'***************************************************************************************************
'                         ■■■ このClassオブジェクト生成時のアクション ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    'Assign an objID for easy debug identification
    LogID = "BDID" & Format(Rnd * 1000, "000")

    'タイムアウト設定(デフォルトで、30s)
    timerOut = TIMEOUT_DEFAULT * 1000   'ミリ秒に補正

    '`WithEvents`からのコマンド結果を蓄積させる件数設定(デフォルトで、256件)
    Set resultBiDi.response = New Dictionary
    resultBiDi.Limit = MaxStoreResultBiDi
End Sub
