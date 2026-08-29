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
'
'   `ExecuteBiDi`等で使用できるコマンド一覧は、右記にてアクセスして下さい → https://www.w3.org/TR/webdriver-bidi/
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const TIMEOUT_DEFAULT       As Double = 30                  '初期タイムアウト秒数
Private Const ThisClassName         As String = "WebDriverBiDiMode" 'ログのトレース用



'***************************************************************************************************
'                                 ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'WebDriverBiDi制御用コアクラスオブジェクト
Private WithEvents m_BiDiCore   As WebDriverBiDiCore      'イベント受け取りオブジェクト
Attribute m_BiDiCore.VB_VarHelpID = -1

'BiDiコマンドの結果を蓄積する変数(`WithEvents`で直接は取らないようにする仕組み作り)
Private ResultCommandBiDi As Dictionary  '「ID - それに対するBiDiコマンドの結果」というDictionary

'ブラウザからの非同期イベント状況
Private m_Events    As Dictionary

'各種エラー情報
Private m_ErrorDetails  As Dictionary

'ログ設定
Private currentLog      As StateLog

'その他
Private timerOut        As Double       'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
Private BiDiUserName    As String       'User-data-dir と同等
Private DefinitionBiDiContext As Long   '確保中の`WebDriverBiDiContext.cls` ※適切なメソッドを呼び出した場合に限る
Private InfoList        As Collection   '非同期イベント`browsingContext.contextCreated`で得た`browsingContext.InfoList`情報。基本的に、`WebDriverBiDiContext.cls - attachToContext`専用



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する所定の起動引数をつけて、WebDriverBiDiとして、起動させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：Name                           ブラウザ種別（`BrowserList.RunChrome` / `BrowserList.RunEdge`）
'            appUrl                         最初の起動時に遷移したいURL
'            userProfile                    UserDataDir名を指定します
'            addArgs                        追加の起動引数を設けます
'            sessionCapabilitiesRequest     `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'***************************************************************************************************
Public Sub StartBiDiMode(Optional Name As BrowserList = BrowserList.RunChrome, _
    Optional appUrl As String, _
    Optional userProfile As String, Optional addArgs As String, _
    Optional sessionCapabilitiesRequest As Dictionary)


    '1. 初期化処理
    BiDiUserName = userProfile
    Set m_BiDiCore = New WebDriverBiDiCore

    '2. WebDriverBiDi としてブラウザを起動させる
    m_BiDiCore.InitBiDi Name, appUrl, BiDiUserName, addArgs, sessionCapabilitiesRequest
    CDPHelpers.InheritLogSettings currentLog, m_BiDiCore.DirectControlCDP.InheritanceCDPCore.LogSettings

    '3. BiDi との新しいセッションを確立する
    Dim resultBiDi As BiDiCDPJson
    Set resultBiDi = sessionNew(sessionCapabilitiesRequest)

    '4. 接続先ブラウザ情報をログで表示
    ViewNewSession resultBiDi
End Sub

'***************************************************************************************************
'* 機能　　："browsingContext.contextCreated"による非同期イベント収集を有効/無効化します
'---------------------------------------------------------------------------------------------------
'* SET引数 ：格納させたいCollection変数。`Nothing`で渡せば、収集を無効化します
'---------------------------------------------------------------------------------------------------
'* 機能説明：・タブやiframeが開いた際の非同期イベントを流すかの切り替えです。
'            ・参照渡し方式により、わざわざ`Get`プロパティを介さず、取得可能です
'---------------------------------------------------------------------------------------------------
'* 注意事項：用が済んだら必ず、`Nothing`を渡して、参照カウンタを0にします
'***************************************************************************************************
Property Set RunCollect_InfoList(ArgCollect As Collection)
    '引数のポインタCollectionを指定させる
    Set InfoList = ArgCollect
End Property

'***************************************************************************************************
'* 機能　　：`browsingContext.contextCreated/browsingContext.contextDestroyed`の2種の非同期イベントを流しつつ、Context のカウントも行います
'---------------------------------------------------------------------------------------------------
'* SET引数 ：論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：・`WebDriverBiDiContext.cls`の生存状況把握に使います
'            ・うまい具合に、`WebDriverBiDiContext.cls`確保中のカウントも行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：`WebDriverBiDiContext.cls`を生成/破棄時に呼ぶ想定です。
'***************************************************************************************************
Property Let EnableDiscoverContexts(Flag As Boolean)
    Const FromProcedureName As String = ThisClassName & ".EnableDiscoverContexts"


    '1. カウント
    DefinitionBiDiContext = DefinitionBiDiContext + IIf(Flag, 1, -1)

    '2. 実行判定
    If DefinitionBiDiContext = 1 Or DefinitionBiDiContext = 0 Then
        '3. 購読準備
        Dim ListenEventList As New Collection
        ListenEventList.Add "browsingContext.contextCreated"
        ListenEventList.Add "browsingContext.contextDestroyed"

        '4. 実行
        Set sessionSubscribe(Flag) = ListenEventList
    End If

    '5. 下限は、0にしておく
    If DefinitionBiDiContext < 0 Then DefinitionBiDiContext = 0: printMsg WARN_, "`WebDriverBiDiContext.cls` no longer exists.", FromProcedureName

    '6. ログ表示
    printMsg info_, "Currently, " & DefinitionBiDiContext & " `WebDriverBiDiContext.cls` file is being defined.", FromProcedureName
End Property

'***************************************************************************************************
'* 機能　　：Excelテーブルに、BiDi-context情報を記録したり、読み取りを行います
'***************************************************************************************************
Public Property Let serializeMainTab(contextId As String)
    Const FromProcedureName As String = ThisClassName & ".MainTab - Let"


    ShSetting01_StartBrowser.UpdateBrowserContext(BiDiUserName, "BiDi-context", FromProcedureName) = contextId
End Property
Public Property Get serializeMainTab() As String
    Const FromProcedureName As String = ThisClassName & ".MainTab - Get"


    '1. Excelテーブルから、`UserName`に紐づくタブセッション情報を読み込む
    Dim tmp As Dictionary
    Set tmp = ShSetting01_StartBrowser.TableBrowserContext(BiDiUserName, FromProcedureName)

    '2. 返却
    If Not tmp Is Nothing Then serializeMainTab = tmp("BiDi-context")
End Property



'***************************************************************************************************
'                               ■■■ 再接続処理系 ■■■
'***************************************************************************************************
'* 機能　　：Excelテーブルにある既存のBiDiMapperタブ,パイプハンドル情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：userProfile                再アタッチしたいユーザー名(user-data-dirに基づく)
'            sessionCapabilitiesRequest `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'            WebSocketMode              WebSocketでBiDi制御する場合は、接続処理済みのClassオブジェクトを指定します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・パイプが生きてない場合は、このメソッドから再開できません
'            ・`sessionCapabilitiesRequest`は、新規に`BiDi-CDP Mapper`が起動された時のみ適用されます
'***************************************************************************************************
Public Function reattach(userProfile As String, Optional sessionCapabilitiesRequest As Dictionary, Optional WebSocketMode As CDPCoreViaWebSocket) As Boolean
    Const FromProcedureName As String = ThisClassName & ".reattach"


    '1. 「BiDi-CDP Mapper」への再接続処理
    BiDiUserName = userProfile
    Set m_BiDiCore = New WebDriverBiDiCore
    Dim requestSessionNew As Boolean: requestSessionNew = m_BiDiCore.ReBootBiDiCDPMapper(BiDiUserName, WebSocketMode)
    CDPHelpers.InheritLogSettings currentLog, m_BiDiCore.DirectControlCDP.InheritanceCDPCore.LogSettings

    '2. 結果に応じて`session.new`を行う
    If requestSessionNew Then
         printMsg Debug_, "A new instance of `BiDi-CDP Mapper` has been launched.", FromProcedureName
         sessionNew sessionCapabilitiesRequest
    End If

    '3. セッション一覧を取得して、既存のものがあるか確認する
    ' (BiDiは接続ごとにセッションIDがかわるわけではないが、安全のためセッション状態を確認)
    Dim resultBiDi As BiDiCDPJson
    Set resultBiDi = sessionStatus

    '4. mapperTab.js の仕様に沿った確認に従う
    If resultBiDi.ExistsKey("ready") Then
        printMsg info_, "Successfully reattached to BiDi! Session status ready: " & resultBiDi.StringKey("message"), FromProcedureName
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
'
'            setMain        `True`で、Excelに`context`を記録し、リアタッチに備えます
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebDriverBiDi では、いきなりURLで直遷移はできないようです
'***************************************************************************************************
Public Function newTab(Optional newWindow As Boolean, Optional isBackground As Boolean, Optional setMain As Boolean) As WebDriverBiDiContext
    Const FromProcedureName As String = ThisClassName & ".newTab"


    '1. パラメーター準備
    Dim BiDiparams As New Dictionary
    BiDiparams.Add "type", IIf(newWindow, "window", "tab")
    BiDiparams.Add "background", isBackground

    '2.実行
    Dim bidiRes As BiDiCDPJson
    Set bidiRes = ExecuteBiDi("browsingContext.create", BiDiparams)

    '3. 返却準備
    Set newTab = New WebDriverBiDiContext
    With newTab
        '3-1. このBiDiのイベントオブジェクトを渡す
        Set .InheritanceWebDriverBiDiCore = m_BiDiCore

        '4. このBiDi情報も渡す
        Set .InheritanceWebDriverBiDiMode = Me

        '5. 先ほど得た`context`を渡す
        .context = bidiRes.StringKey("context")

        '6. メイン化する場合は、Excelに記録しておく
        If setMain Then serializeMainTab = .context
    End With

    printMsg info_, "newTab will now open a new BiDi instance", FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：既に開いているタブのセッションIDを、タブの名前またはURLに基づいて検索します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - WebDriverBiDiContext
'* 引数　　：tabName        タブ名
'            Url            URL
'            maxDepth       Tree の深さレベル。基本は、0にすると子フレーム(iframe)を含まず、トップレベルのタブだけ取得となります
'            setMain        `True`で、Excelに`context`を記録し、リアタッチに備えます
'            doRetrySecond  指定秒以内に出るまでリトライを繰り返します。0で、1回のみです
'---------------------------------------------------------------------------------------------------
'* 機能説明：tabName引数とurl引数は部分的な指定(部分一致)が可能です。maxDepth はデフォルトは、トップレベルのタブだけにしてます
'* 注意事項：・url引数を両方とも省略すると、getTabは最も近い未接続のタブに接続しようとします。
'            ・WebDriverBiDi では、タブ名での検索はできないようです
'***************************************************************************************************
Public Function getTab(Optional Url As String, Optional maxDepth As Long, Optional setMain As Boolean, Optional doRetrySecond As Double) As WebDriverBiDiContext
    Const FromProcedureName As String = ThisClassName & ".getTab"


    Dim startTime As Double
    Dim bidiRes As BiDiCDPJson
    Dim contexts As BiDiCDPJson
    Dim matchData As BiDiCDPJson

    'パラメーターを用意
    Dim BiDiparams As New Dictionary
    BiDiparams.Add "maxDepth", maxDepth

    'タイマースタート
    If doRetrySecond Then startTime = CDPHelpers.TimerCounter
    
    'シンプルに、`Instr`方式の検索で行う
    Do
        ' 1. 指定した深さまでのツリーを取得
        Set bidiRes = ExecuteBiDi("browsingContext.getTree", BiDiparams)

        If Not bidiRes Is Nothing Then
            Set contexts = bidiRes.NodeKey("contexts")

            ' 2. 再帰関数を呼び出して全階層からマッチするものを探す
            Set matchData = FindContextRecursive(contexts, Url)

            ' 3. 見つかったらオブジェクト化して返す
            If Not matchData Is Nothing Then
                Set getTab = New WebDriverBiDiContext
                With getTab
                    '4. このこのBiDiのイベントオブジェクトを渡す
                    Set .InheritanceWebDriverBiDiCore = m_BiDiCore

                    '5. このBiDi情報を渡す
                    Set .InheritanceWebDriverBiDiMode = Me

                    '6. 先ほど得た`context`を渡す
                    .context = matchData("context")

                    '7. メイン化する場合は、Excelに記録しておく
                    If setMain Then serializeMainTab = matchData("context")
                End With

                Exit Do
            End If
        End If

        If Not doRetrySecond Then Exit Do
        If CDPHelpers.TimerCounter - startTime > doRetrySecond * 1000 Then Exit Do
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
Private Function FindContextRecursive(contexts As BiDiCDPJson, Url As String) As BiDiCDPJson
    Dim ctxData As BiDiCDPJson
    Dim childMatch As BiDiCDPJson
    Dim contextCount As Long

    ' contextsはCollection（配列）としてループ
    Dim i As Long
    contextCount = contexts.Count
    For i = 0 To contextCount
        Set ctxData = contexts.NodeIndex(i)


        ' A. このコンテキスト自体がマッチするかチェック
        Dim isMatch As Boolean
        If LenB(Url) = 0 Then
            isMatch = True ' 空なら最初に見つかったヤツ
        Else
            If InStr(1, ctxData.StringKey("url"), Url) > 0 Then isMatch = True
        End If

        If isMatch Then
            Set FindContextRecursive = ctxData
            Exit Function
        End If

        ' B. 子要素（children）があれば、さらにその中を探索
        If ctxData.NodeKey("children").Count > 0 Then
            Set childMatch = FindContextRecursive(ctxData.NodeKey("children"), Url)
            If Not childMatch Is Nothing Then
                Set FindContextRecursive = childMatch
                Exit Function
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
'
'* 返値    ：ブラウザから返却されたBiDi結果 (パースされたDictionary)
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを送信後、ID返却されるまで待機します。
'***************************************************************************************************
Public Function ExecuteBiDi(methodName As String, Optional params As Dictionary, Optional StopBiDiError As Boolean = True) As BiDiCDPJson
    Const FromProcedureName As String = ThisClassName & ".ExecuteBiDi"


    '1. まずは非同期実行で実行IDを回収し、IDが0になってないか確認
    Dim RunID As Long
    RunID = ExecuteBiDiAsync(methodName, params, StopBiDiError)
    If RunID = 0 Then Exit Function

    '2. 1回だけメッセージ取り出しを行う
    printMsg Debug_, "Waiting for the command result response for synchronous execution... : " & RunID, FromProcedureName
    TakeEvents StopBiDiError

    '3. BiDi-Jsonコマンドを送った際のIDが返ってくるまで、内部でループ受信処理を行います
    '前述の`TakeEvents`により
    '・結果が届いた：ループに入らず即回収
    '・結果がまだ　：ループに入って回収
    'といったパフォーマンス改善になります
    Dim ResultJsonStr As String
    ResultJsonStr = AutoWaitTakeResultBiDi(RunID, StopBiDiError)

    '4. 結果を、Node変換
    Dim resultJson As BiDiCDPJson
    Set resultJson = BiDiCDPJson.Parse(ResultJsonStr)

    '5. 実行時エラーチェック
    If ExecuteError(resultJson) Then
        'エラーによる`Err.Raise`で止めるか？
        If StopBiDiError Then
            '5-1. エラーになったmethod名と、結果生responseを表示して停止
            printMsg ERROR_, "Failed to ExecuteBiDi method """ & methodName & """" & vbCrLf & vbCrLf & "=== RawErrorResponse ===" & vbCrLf & ResultJsonStr, FromProcedureName, doRaiseError:=True
        Else
            '5-1. エラーになったmethod名だけログ出して続行
            printMsg WARN_, "Failed to ExecuteBiDi method """ & methodName & """", FromProcedureName
        End If
    Else
        '5-1. Return result object(`CDPBrowser`の思想に従い、`Result`オブジェクトで返却します)
        Set ExecuteBiDi = resultJson.NodeKey("result")
    End If

    '6. デバッグ用に、As Is 出力
    If currentLog.seeRawSendMsgDbg Then printMsg Debug_, "ExecuteBiDi `" & methodName & "` with result: " & ResultJsonStr, FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (非同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "session.new")
'*           params        引数を持ったDictionary (省略可能)
'            StopError     内部エラーがあったら止めるか？
'
'* 返値    ：送信したメッセージの ID    ※API失敗時は0で、成功時は1以上です。
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを構成し、JSON化してブラウザへ投げます。結果の待機は行いません。
'***************************************************************************************************
Public Function ExecuteBiDiAsync(methodName As String, Optional params As Dictionary, Optional StopError As Boolean = True) As Long
    Const FromProcedureName As String = ThisClassName & ".ExecuteBiDiAsync"


    '1. `nothing`の場合は、オブジェクトを生成。BiDiは、`params`が不要でも、「"params":{}」という記法が必要
    If params Is Nothing Then Set params = New Dictionary

    '2. コマンド組み立て
    Dim BiDiCommand As New Dictionary
    BiDiCommand.Add "method", methodName
    BiDiCommand.Add "params", params

    '3. 送信し、実行IDを取得
    ExecuteBiDiAsync = m_BiDiCore.ReadyRunBiDi(BiDiCommand, StopError)

    '4. 送信APIエラーチェック
    If m_BiDiCore.LastErrorWriteFile Then
        printMsg ERROR_, "Error WriteFile in m_BiDiCore.writeProcCDP. ErrorCode:" & m_BiDiCore.LastErrorWriteFile & " - " & WinApiError.GetMessage(m_BiDiCore.LastErrorWriteFile, "kernel32"), FromProcedureName, doRaiseError:=StopError

        ExecuteBiDiAsync = 0    '失敗なので、0にセット
        Exit Function
    End If

    '5. ログ化
    printMsg Debug_, "Async ExecuteBiDi: " & methodName & " - ID: " & ExecuteBiDiAsync, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：非同期実行(ExecuteBiDiAsync)した際のcommandIDに対して、結果が来るまで内部で待機します
'---------------------------------------------------------------------------------------------------
'* 返り値　：commandIDに対するBiDi-Jsonコマンド の実行結果(Json文字列)
'* 引数　　：CommandID      実行ID
'            StopApiError   `Pipe - ReadFile`/`WebSocket - recv`でエラーが起こったら停止します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・タイムアウトの場合は、エラー停止になります。結果に時間がかかりそうなコマンドの場合は事前に`TimeOutSecond`にて伸ばして下さい
'            ・WinAPI側のエラーの場合は、`vbnullstring`で返ります
'***************************************************************************************************
Public Function AutoWaitTakeResultBiDi(commandID As Long, Optional StopApiError As Boolean = True) As String
    Const FromProcedureName As String = ThisClassName & ".AutoWaitTakeResultBiDi"


    '1. すでに結果が格納済みの場合は、ここの段階で抜ける
    AutoWaitTakeResultBiDi = TakeResultBiDi(commandID)
    If StrPtr(AutoWaitTakeResultBiDi) Then printMsg Debug_, "The result has already been stored as a side effect of another process.", FromProcedureName: Exit Function

    '2. 今の経過時間を記録し、取り出し開始
    Dim timerStart As Double: timerStart = CDPHelpers.TimerCounter
    Do
        '3. 受信処理によるパイプを拝見する
        '※内部にて`DoEvents`が走ります
        m_BiDiCore.TakeEvents StopApiError

        '4. 内部の`RaiseEvent detectionBiDiID`による発火で、Dictionaryへ蓄積されたか？
        AutoWaitTakeResultBiDi = TakeResultBiDi(commandID)
        If StrPtr(AutoWaitTakeResultBiDi) Then
            '送ったコマンド結果であることを確認できたので、ループから抜ける
            printMsg info_, "WebDriverBiDi have confirmed receipt of command result ID: " & commandID, FromProcedureName
            Exit Do
        Else
            printMsg Trace_, "WebDriverBiDi have not yet been able to confirm command ID: " & commandID, FromProcedureName
        End If

        '5. 受信処理にてエラーがあったら、即抜け
        If m_BiDiCore.LastErrorPeekNamedPipe > 0 Or m_BiDiCore.LastErrorReadFile > 0 Then printMsg ERROR_, "Possible reasons: (1) the browser is not actually running (ie. call string failure); (2) the browser window was accidentally closed; (3) the CDP Pipe connection to the browser was cut-off somehow.", FromProcedureName, doRaiseError:=StopApiError: Exit Function

        '6. 一定時間経ってもコマンド結果が来ないなら、タイムアウトエラーとして終了します
        If CDPHelpers.TimerCounter - timerStart > timerOut Then printMsg ERROR_, "Timeout waiting for BiDi to return message.", FromProcedureName, , True
    Loop
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
'* 注意事項：Jsonは、文字列のままで渡されます。
'***************************************************************************************************
Private Sub m_BiDiCore_detectionBiDiID(id As Long, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".m_BiDiCore_detectionBiDiID"


    '結果を蓄積し、後で取り出せるようにする
    ResultCommandBiDi.Add id, RawJson
    printMsg Trace_, "The results of the WebDriverBiDi(ID: " & id & ") execution from the browser have been received. → " & RawJson, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：オーバーフロー対策によるcommandIDがリセットされると呼び出されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：IDが20億に達する場合も、結果履歴が全て削除されます
'***************************************************************************************************
Private Sub m_BiDiCore_ResetCommandID()
    ResultCommandBiDi.RemoveAll
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
Property Get TakeResultBiDi(commandID As Long) As String
    Const FromProcedureName As String = ThisClassName & ".TakeResultBiDi"


    '1. 存在するか？
    If ResultCommandBiDi.Exists(commandID) Then
        '2. Log化
        Dim tmp As String: tmp = ResultCommandBiDi(commandID)
        printMsg info_, "Browser were able to confirm the existence of BiDi command ID:" & commandID & "　→　" & tmp, FromProcedureName

        '3. 返却
        TakeResultBiDi = tmp

        '4. 取り出したので消しておく
        ResultCommandBiDi.Remove commandID
        printMsg Trace_, "BiDi-command ID: " & commandID & " has been removed from `ResultBiDi.response`.", FromProcedureName
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
Private Function ExecuteError(ResultsObj As BiDiCDPJson) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ExecuteError"


    Dim ErrorType As String, ErrorMessage As String, ErrorStackTrace As String

    'データの状態をチェック
    If ResultsObj.Count = 0 Then
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

    ElseIf ResultsObj.StringKey("type") = "error" Then
        'BiDiコマンドエラーは "type":"error" で返る
        ExecuteError = True


        'ログを出しておく
        ErrorType = ResultsObj.StringKey("error")
        ErrorMessage = ResultsObj.StringKey("message")
        printMsg WARN_, "BiDi-InvokeMethod Error: [" & ErrorType & "] " & ErrorMessage, FromProcedureName

        '`StackTrace`があったら別枠で表示
        If ResultsObj.ExistsKey("stacktrace") Then
            ErrorStackTrace = ResultsObj.StringKey("stacktrace")
            printMsg WARN_, "       " & ErrorStackTrace, FromProcedureName
        End If

        '最後のエラー情報を記録
        Set m_ErrorDetails = New Dictionary
        m_ErrorDetails.Add "error", ErrorType
        m_ErrorDetails.Add "message", ErrorMessage
        If StrPtr(ErrorStackTrace) Then m_ErrorDetails.Add "stacktrace", ErrorStackTrace
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
'* 機能　　：BiDi-Json コマンド結果を手動でクリアします
'---------------------------------------------------------------------------------------------------
'* 返値    ：クリア件数
'---------------------------------------------------------------------------------------------------
'* 機能説明：任意のタイミングでの、メモリクリーン等にどうぞ
'***************************************************************************************************
Public Function CleanResultCommandBiDi() As Long
    Const FromProcedureName As String = ThisClassName & ".CleanResultCommandBiDi"


    '1. クリア件数を返却
    CleanResultCommandBiDi = ResultCommandBiDi.Count

    '2. ALLクリア
    ResultCommandBiDi.RemoveAll
    printMsg info_, "The user cleared " & CleanResultCommandBiDi & " previous BiDi command results.", FromProcedureName
End Function



'***************************************************************************************************
'                                       ■■■ 後始末系 ■■■
'***************************************************************************************************
Public Sub quit()
    Const FromProcedureName As String = ThisClassName & ".quit"


    '1. 終了コマンドを送る
    '※閉じると同時に、Pipeエラーが起こることがあるので、エラー無視しておく
    ExecuteBiDi "browser.close", StopBiDiError:=False

    '2. パイプハンドル、ブラウザセッション情報を消去しに行きます
    m_BiDiCore.CloseHandleBrowser   'パイプハンドルを閉じる
    deleteCDPID                     'テーブルからも削除する

    '3. it seems without waiting a bit the browser crashes and the next time it wants to recover from a crash
    CDPHelpers.Sleep 1
    printMsg info_, "The WebDriverBiDi has been terminated.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：Excelのテーブルから、ブラウザへの接続パイプとタブ情報を削除します
'***************************************************************************************************
Private Sub deleteCDPID()
    Const FromProcedureName As String = ThisClassName & ".deleteCDPID"


    With ShSetting01_StartBrowser
        .DeleteBrowserHandle(FromProcedureName) = BiDiUserName   'パイプ通信Handle情報
        .DeleteBrowserContext(FromProcedureName) = BiDiUserName  'ブラウザセッション情報
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


    '------------------ 1. `Context`の生成情報を蓄積します ------------------
    '1-1. Select Case でスマートに、method名ごとの独自処理を行う
    Select Case methodName
        Case "browsingContext.contextCreated"
            '1-2. `InfoList`を収集するか判定
            If Not InfoList Is Nothing Then
                '1-3. そのまま、蓄積
                printMsg Debug_, "Collecting `browsingContext.contextCreated`... → " & RawJson, FromProcedureName
                InfoList.Add RawJson
            End If

        Case Else
            'None

    End Select


    '------------------ 2. 非同期イベントを収集します ------------------
    '2-1. イベントキャプチャしない場合は、即抜け
    If m_Events Is Nothing Then printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName: Exit Sub

    '2-2. パース処理
    printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
    Dim tmp As Dictionary: Set tmp = WebJsonConverter.Parse(RawJson).value

    '2-3. 専用のプロシージャにて、蓄積処理
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
Private Sub ViewNewSession(ResultSessionNew As BiDiCDPJson)
    Const FromProcedureName As String = ThisClassName & ".ViewNewSession"


    '現時点では、`platformName`は空欄なので、せっかくなので、このアプリケーションの名称にしてみます
    Dim Node_capabilities As BiDiCDPJson: Set Node_capabilities = ResultSessionNew.NodeKey("capabilities")
    Dim pfName As String: pfName = Node_capabilities.StringKey("platformName")
    If LenB(pfName) = 0 Then pfName = Application.Name

    'ログビュー
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, " BiDi Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "sessionId             | " & ResultSessionNew("sessionId"), FromProcedureName
    printMsg Debug_, "acceptInsecureCerts   | " & Node_capabilities.StringKey("acceptInsecureCerts"), FromProcedureName
    printMsg Debug_, "browserName           | " & Node_capabilities.StringKey("browserName"), FromProcedureName
    printMsg Debug_, "browserVersion        | " & Node_capabilities.StringKey("browserVersion"), FromProcedureName
    printMsg Debug_, "platformName          | " & pfName, FromProcedureName
    printMsg Debug_, "setWindowRect         | " & Node_capabilities.StringKey("setWindowRect"), FromProcedureName
    printMsg Debug_, "webSocketUrl          | " & Node_capabilities.StringKey("webSocketUrl"), FromProcedureName
    printMsg Debug_, "userAgent             | " & Node_capabilities.StringKey("userAgent"), FromProcedureName
    printMsg Debug_, "mapperTab.js Version  | " & ShLibrary01_JS.VersionMapperTabJS, FromProcedureName
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
Private Function sessionNew(Optional params As Dictionary) As BiDiCDPJson
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
Private Function sessionStatus() As BiDiCDPJson
    Set sessionStatus = ExecuteBiDi("session.status")
End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンド、`session.subscribe/unsubscribe`を実行します
'---------------------------------------------------------------------------------------------------
'* SET引数 ：`events`に渡すCollection変数
'     引数 ：subscribe  `True` :`session.subscribe`として実行   ※既定
'                       `False`:`session.unsubscribe`として実行
'---------------------------------------------------------------------------------------------------
'* 注意事項：何を購読中かは、ここでは管理しません
'***************************************************************************************************
Property Set sessionSubscribe(Optional subscribe As Boolean = True, events As Collection)
    'paramsに格納
    Dim params As New Dictionary
    params.Add "events", events

    '実行
    ExecuteBiDi IIf(subscribe, "session.subscribe", "session.unsubscribe"), params, False
End Property



'***************************************************************************************************
'                              ■■■ PUBLIC PROPERTIES ■■■
'***************************************************************************************************
'* 機能　　：BiDi-Json コマンド結果を得るためのタイムアウト時間を設定します
'---------------------------------------------------------------------------------------------------
'* LET引数 ：置き換えるタイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDi-Json コマンド結果を得るのに待つ時間を設定します。デフォルトは、30sです。
'***************************************************************************************************
Property Let TimeOutSecond(TimeSec As Double)
    timerOut = TimeSec * 1000       'ミリ秒に補正する
End Property



'***************************************************************************************************
'                              ■■■ 複製用プロパティ ■■■
'***************************************************************************************************
'* 機能　　：このClass内で定義してる`WebDriverBiDiCore.cls`を受け取ります
'***************************************************************************************************
Property Get InheritanceBiDiCore() As WebDriverBiDiCore
    Set InheritanceBiDiCore = m_BiDiCore
End Property



'***************************************************************************************************
'                           ■■■ CDPHelpers ラッパーメソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".bidi.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub



'***************************************************************************************************
'                        ■■■ WebDriverBiDiCore.cls 直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：メッセージ取り出し/制御系。`RaiseEvent`のコア
'***************************************************************************************************
'--- almighty ---
Public Sub TakeEvents(Optional StopApiError As Boolean = True, Optional destruction As Boolean)
    m_BiDiCore.TakeEvents StopApiError, destruction
End Sub

'--- manual ---
Public Sub TakeEvent()
    m_BiDiCore.TakeEvent
End Sub

Property Get NewResBiDi() As Boolean
    NewResBiDi = m_BiDiCore.NewResBiDi
End Property

Public Sub AnalyzeBiDi(Optional StopApiError As Boolean = True, Optional destruction As Boolean)
    m_BiDiCore.AnalyzeBiDi StopApiError, destruction
End Sub



'***************************************************************************************************
'                         ■■■ このClassオブジェクト生成時のアクション ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'Assign an objID for easy debug identification
    currentLog.LogID = "BDID" & Format(Rnd * 1000, "000")

    'タイムアウト設定(デフォルトで、30s)
    timerOut = TIMEOUT_DEFAULT * 1000   'ミリ秒に補正

    '`WithEvents`からのコマンド結果を蓄積させる初期化
    Set ResultCommandBiDi = New Dictionary
End Sub
