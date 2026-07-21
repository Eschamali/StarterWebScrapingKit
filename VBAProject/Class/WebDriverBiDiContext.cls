VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverBiDiContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'           WebDriverBiDi制御のうち、Context(タブ等)操作としてのメソッドやパラメータ運用を行います
'               必要最低限のメソッドを提供しますが実際の送信は、`WebDriverBiDiMode`側で行います
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const ThisClassName As String = "WebDriverBiDiContext" 'ログのトレース用



'***************************************************************************************************
'                                 ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'WebDriverBiDi全体制御用クラスオブジェクト
Private WithEvents m_BiDiCore   As WebDriverBiDiCore    'イベント受け取りオブジェクト
Attribute m_BiDiCore.VB_VarHelpID = -1
Private m_BiDiMode              As WebDriverBiDiMode    'BiDi全体

'各種ステータス
Private Type browsingContextInfo
    context          As String   '現在、制御中のContext(タブ)パラメーター
    Url              As String   '現時点でのURL
End Type
Private InfoList As browsingContextInfo
Private contextDestroyed As Boolean       'お亡くなりサイン

Private LogID  As String


Private lastExceptionDetailsJS  As BiDiCDPJson 'JavaScript実行時の最後のエラー情報(Node情報方式)



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する所定の起動引数をつけて、WebDriverBiDiとして、起動させ、非同期イベント：`browsingContext.contextCreated`を利用した初回タブ接続処理を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：Name                           使用ブラウザ名※現時点では、Edgeとchromeのみです
'            appUrl                         最初の起動時に遷移したいURL
'            userProfile                    UserDataDir名を指定します
'            addArgs                        追加の起動引数を設けます
'            KioskMode                      0(省略)：通常モード(キオスクモードは使いません)
'                                           1      ：キオスクモード デジタル/対話型サイネージ
'                                           2      ：キオスクモード パブリック ブラウジング
'
'            sessionCapabilitiesRequest     `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'***************************************************************************************************
Public Sub StartBiDiModeAndConnectTab(Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional userProfile As String, Optional addArgs As String, Optional KioskMode As edgeKioskType, _
    Optional sessionCapabilitiesRequest As Dictionary)


    '1. 初期化処理
    Set m_BiDiMode = New WebDriverBiDiMode

    '2. WebDriverBiDi としてブラウザを起動させ、`Session`を確立
    m_BiDiMode.StartBiDiMode Name, appUrl, userProfile, addArgs, KioskMode, sessionCapabilitiesRequest

    '3. イベントを継承
    Set m_BiDiCore = m_BiDiMode.InheritanceBiDiCore

    '4. タブの接続へ
    attachToContext
End Sub



'***************************************************************************************************
'                               ■■■ 再接続処理系 ■■■
'***************************************************************************************************
'* 機能　　：Excelテーブルにある既存のBiDi-contextタブ情報を利用して、再接続を試みます
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


    '1. 新オブジェクトを用意
    Set m_BiDiMode = New WebDriverBiDiMode

    '2. 識別名称に紐づく「BiDi-CDP Mapper」への再接続処理
    If Not m_BiDiMode.reattach(userProfile, sessionCapabilitiesRequest, WebSocketMode) Then printMsg ERROR_, "Failed to reattach to BiDi session.", FromProcedureName, doRaiseError:=True
    Set m_BiDiCore = m_BiDiMode.InheritanceBiDiCore

    '3. タブ情報を再セット
    InfoList.context = m_BiDiMode.serializeMainTab

    '4. 生存確認
    Dim BiDiparams As New Dictionary, BiDiResult As BiDiCDPJson
    BiDiparams.Add "root", InfoList.context
    BiDiparams.Add "maxDepth", 0
    Set BiDiResult = m_BiDiMode.ExecuteBiDi("browsingContext.getTree", BiDiparams, False)

    '5. 判定
    If BiDiResult Is Nothing Then
        '6. エラーログを出す
        printMsg ERROR_, "Failed to reattach to Previous BiDi-context.", FromProcedureName: Exit Function
        reattach = False
    Else
        '6. 購読処理
        m_BiDiMode.EnableDiscoverContexts = True
        reattach = True
    End If
End Function



'***************************************************************************************************
'                              ■■■ BiDi非同期イベント関連 ■■■
'***************************************************************************************************
'* 機能　　：BiDiの非同期イベント
'---------------------------------------------------------------------------------------------------
'* 機能説明：`Context`の生存確認に使います
'***************************************************************************************************
Private Sub m_BiDiCore_detectionBiDiEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".m_BiDiCore_detectionBiDiEvent"


    '------------------ 1. この`Context`のSessionステータスの更新を行います ------------------
    Dim BiDiEventNode As BiDiCDPJson
    Dim StartPos As Long, SearchWord As String
    Select Case methodName
        '`context`が消えた
        Case "browsingContext.contextDestroyed"
            '1. 自分の`context`か確認し、そうであれば、フラグ更新
            Set BiDiEventNode = BiDiCDPJson.Parse(RawJson)
            If BiDiEventNode.NodeKey("params").StringKey("context") = InfoList.context Then contextDestroyed = True: printMsg ERROR_, "The tab has been closed. This Class object is no longer usable. Please discard it.", FromProcedureName

    End Select
End Sub



'***************************************************************************************************
'                              ■■■ BiDicommand操作関連 ■■■
'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "browsingContext.navigate")
'*           params        引数を持ったDictionary (省略時、`context`の付与)
'*           StopBiDiError 実行失敗時に`Err.Raise`で停止し、以降の処理を止めるか
'
'* 返値    ：ブラウザから返却されたBiDi結果 (パースされたDictionary)
'---------------------------------------------------------------------------------------------------
'* 機能説明：`params`に、`context`を付け足した上で実行されます
'***************************************************************************************************
Public Function ExecuteBiDi(methodName As String, Optional params As Dictionary, Optional StopBiDiError As Boolean = True) As BiDiCDPJson
    '`context`情報の追加
    If params Is Nothing Then Set params = New Dictionary
    params.Add "context", InfoList.context

    '実行
    Set ExecuteBiDi = m_BiDiMode.ExecuteBiDi(methodName, params, StopBiDiError)
End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (非同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "browsingContext.navigate")
'*           params        引数を持ったDictionary (省略時、`context`の付与)
'*           StopError     内部エラーで停止し、以降の処理を止めるか
'
'* 返値    ：送信したメッセージの ID
'---------------------------------------------------------------------------------------------------
'* 機能説明：`params`に、`context`を付け足した上で実行し、結果の待機は行いません。
'***************************************************************************************************
Public Function ExecuteBiDiAsync(methodName As String, Optional params As Dictionary, Optional StopError As Boolean = True) As Long
    '`context`情報の追加
    If params Is Nothing Then Set params = New Dictionary
    params.Add "context", InfoList.context

    '実行
    ExecuteBiDiAsync = m_BiDiMode.ExecuteBiDiAsync(methodName, params, StopError)
End Function



'***************************************************************************************************
'                              ■■■ JavaScript コア機能 ■■■
'***************************************************************************************************
'* 機能　　：記述したJavaScriptをブラウザ内で実行します。JavaScript版低レベル操作としての位置付けです。
'            Webスクレイピングに関わる引数は全て設けています。
'---------------------------------------------------------------------------------------------------
'* 返り値　：value属性の中身を評価して、Variant型で返却。実行したJavaScriptコード、引数:Ownership や、IFEXCEPTION に依存します。
'* 引数　　：strString                      実行したいJavaScriptコードそのもの。`As Is`として実行されます。※ほとんどのエスケープ処理は、`WebJsonConverter`が担うため、ここでわざわざ手動エスケープの記述は必要ありません。
'
'            scriptHandle                   実行の基準`Handle ID`を指定します。これを指定すると`script.callFunction`として優先実行されます。
'
'            objectArguments                これを指定すると`script.callFunction`として優先実行されます。
'                                           予め、呼び出し側にてCollectionを組み立てて渡してください。
'                                           参考となる組み立て手法：https://w3c.github.io/webdriver-bidi/#cddl-type-scriptlocalvalue
'
'            IFEXCEPTION                    JavaScript実行中に「例外（Exception）」が発生した場合の代替値を指定します。名前の通り、IFERROR関数チックに使えます。
'            RealmTarget                    `iframe`内で、JavaScript実行をしたいときに指定します。
'
'            Ownership                      True ：script.ResultOwnership = "root"として扱い、type:"object" 等の場合は、`{"handle": "-4985739485.1.3"}`にある`-4985739485.1.3`というブラウザ内変数名的なものとして返えります。再度`jsEval`を使う際の引数`scriptHandle`に流用できます。「このJSコードは`SET`で...」という意識は不要です。
'                                           False：script.ResultOwnership = "none"として扱い、type:"object" 等の場合でも、`value`で受け取れるように試みます。実行コードによっては、`BiDiCDPJson.cls`形式で返されるため、`SET`で受け取る必要があります。なお、空のオブジェクトで返る場合は、対応してないデータ型の可能性があるためその際は、`True`に切り替えてください。
'
'            awaitPromise                   True ：JavaScriptそのものの実行結果が非同期処理であっても、処理が終わるまで待ってから結果を返すことができます
'                                           False：JavaScriptそのものの実行結果が非同期処理の場合、結果を待たずに続行されます
'
'            userActivation                 True ：「人間がクリックやキーボード操作をした結果」としてブラウザに偽装します。スクレピング対策が施されてる場合は、`True`にしておくといいかもしれません
'                                           False：特に施しません
'
'            serializationOptions           意地でも、type:"object" 等となっても、`value`で受け取れるように、細かい調整をします。予め、呼び出し側にてDictionaryを組み立てて渡してください。
'
'            RunAsyncBiDi                   True ：結果を待たないBiDiで実行します。実行時のIDが返却されます。主に、「この操作、`alert`関数が発動するな...」という場面で使用します。
'                                           False：結果を待つCDPで実行します。基本はこれでOK
'
'            StopException                  `True`で、JavaScriptコード内で例外が起こった場合、Err.Raiseで止まります。基本は、開発時のみ使用する想定です
'            StopBiDiError                  `True`で、BiDiエラー時、エラーで停止します
'---------------------------------------------------------------------------------------------------
'* 機能説明：This function commands the attached browser instance to execute a Javascript expression to perform required tasks.
'            Note: the original function has a bool retry variable which was removed for simplicity.
'            Added dbgMsg to read debug line on the go when needed.
'
'* 注意事項：・利用時は、「名前付き引数`xxx:=yyy`」として運用することをお勧めします
'            ・JavaScript実行中の例外確認は、「IsError関数」で判断してください。詳細の確認は、`LastJavaScriptException`プロパティで確認できます。「IsError関数」での確認が面倒なら、引数`IFEXCEPTION`を活用してください。
'            ・`Ownership:=true` の場合は、SET or LET で受け取る意識が必要になります。結果によって、`SET`or`LET`が変わって静的にできない場合は、Array関数で一旦受け止めて、VarType(hoge(0))で`SET`or`LET`判定することも可能です
'
'            ・「このJavaScriptコードの場合は、こういう返り値が来るな」といった知識がないと、返り値に困惑します。
'               開発時は、デバッグログを活用した試行錯誤をすることをお勧めします。
'
'            ・`scriptHandle`や`objectArguments`を指定する`script.callFunction`方式の場合は、`JavaScriptStr`を 「function() { this... }」 形式で記述する必要があります
'
'            ・実行したJavaScriptコードの結果が、プリミティブ型の場合、`scriptHandle`の設定問わずそのプリミティブ型の中身が返ります。
'              常に`scriptHandle`で返したい場合は、JavaScriptコードそのものを工夫する必要があります。
'
'            ・JavaScriptの場合、例外を利用した処理が多々あるため、運用時は`StopException`を省略しておくことを推奨します
'            ・BiDi実行自体のエラーキャッチは常に`True`となります
'            ・`IFEXCEPTION`はあくまでも、「例外」時のみ機能します。JavaScript自体は成功してるけど、結果がエラーの場合はISERROR関数で対処する必要があります。
'---------------------------------------------------------------------------------------------------
'* 参考URL ：https://w3c.github.io/webdriver-bidi/#command-script-evaluate
'***************************************************************************************************
Public Function jsEval(JavaScriptStr As String, Optional scriptHandle As String, Optional objectArguments As Collection, Optional IFEXCEPTION As Variant, _
                       Optional RealmTarget As String, _
                       Optional Ownership As Boolean, _
                       Optional awaitPromise As Boolean, _
                       Optional userActivation As Boolean, _
                       Optional serializationOptions As Dictionary, _
                       Optional RunAsyncBiDi As Boolean, _
                       Optional StopException As Boolean, _
                       Optional StopBiDiError As Boolean = True) As Variant
    Const FromProcedureName As String = ThisClassName & ".jsEval"

    Dim scriptEvaluateResult As BiDiCDPJson
    Dim BiDiCommand  As New Scripting.Dictionary
    Dim BiDiparams   As New Scripting.Dictionary, scriptTarget As New Dictionary


    'Construct message for sending ----------------------------------------------------------
    '`callFunction`か`evaluate`を決める判定に入ります
    Dim UseJsMethod As String

    ' 【パターン1】特定の要素を操作する場合
    Dim RemoteReference As New Dictionary:
    If StrPtr(scriptHandle) Then
        UseJsMethod = "script.callFunction"

        '必要なパラメーターをセット
        BiDiparams.Add "functionDeclaration", JavaScriptStr
        If Not objectArguments Is Nothing Then BiDiparams.Add "arguments", objectArguments   '引数があったら適用しておく
        RemoteReference.Add "handle", scriptHandle: BiDiparams.Add "this", RemoteReference

    ' 【パターン2】要素指定はないが、安全に引数(arguments)を渡したい場合
    ElseIf Not objectArguments Is Nothing Then
        UseJsMethod = "script.callFunction"

        '必要なパラメーターをセット
        BiDiparams.Add "functionDeclaration", JavaScriptStr
        BiDiparams.Add "arguments", objectArguments

    ' 【パターン3】引数も要素指定もない、ただのコード実行
    Else
        ' 最もシンプルな evaluate を使う
        UseJsMethod = "script.evaluate"

        '必要なパラメーターをセット
        BiDiparams.Add "expression", JavaScriptStr
    End If

    '以降は、共通設定を施す
    If StrPtr(RealmTarget) Then scriptTarget.Add "realm", RealmTarget Else scriptTarget.Add "context", InfoList.context
    BiDiparams.Add "target", scriptTarget
    BiDiparams.Add "awaitPromise", awaitPromise

    '文字数節約により、IF文をかます
    If Ownership Then BiDiparams.Add "resultOwnership", "root"
    If Not serializationOptions Is Nothing Then BiDiparams.Add "serializationOptions", serializationOptions
    If userActivation Then BiDiparams.Add "userActivation", userActivation
    '----------------------------------------------------------------------------------------


    'Send and check BiDiResult -------------------------------------------------------------------------------------------
    '引数に応じて、BiDi自体を同期/非同期で実行し、分岐します
    If RunAsyncBiDi Then
        '非同期の場合は、実行IDを返して、即抜け
        jsEval = m_BiDiMode.ExecuteBiDiAsync(UseJsMethod, BiDiparams, StopBiDiError)
        Exit Function
    Else
        'BiDiエラー判定
        Set scriptEvaluateResult = m_BiDiMode.ExecuteBiDi(UseJsMethod, BiDiparams, StopBiDiError)
        If scriptEvaluateResult Is Nothing Then
            '「機能しないコマンド」という意味合いでエラー返却
            jsEval = CVErr(xlErrNA)
            Exit Function
        End If
    End If
    '----------------------------------------------------------------------------------------


    'JavaScript実行結果部分の解析フェーズに移ります -------------------------------------------------------------------------------------------
    Dim strLogValue As String

    '例外（JS側のエラー）が発生していないかチェックし、あったら、ログを表示し、抜けます
    If scriptEvaluateResult.StringKey("type") = "exception" Then
        '最後の詳細例外情報を記録
        Set lastExceptionDetailsJS = scriptEvaluateResult.NodeKey("exceptionDetails")

        '実行したJavaScriptコードの改行を消して、長すぎる場合はカットする（整形）
        strLogValue = LogFormat(JavaScriptStr, 100)

        '例外による`Err.Raise`で止めるか？
        If StopException Then
            '`StopException`が有効の場合は、ログ表示用に、詳細の方を文字列化
            Dim exceptionDetails As String: exceptionDetails = lastExceptionDetailsJS.Stringify
            printMsg ERROR_, "An exception occurred while executing JavaScript: " & strLogValue & vbCrLf & vbCrLf & "<exceptionDetails>" & vbCrLf & exceptionDetails, FromProcedureName, doRaiseError:=True
        Else
            '`StopException`が無効の場合は、簡易部分の例外内容を出す
            printMsg WARN_, "ExceptionDescription: " & lastExceptionDetailsJS.StringKey("text"), FromProcedureName

            '返り値判定
            If IsMissing(IFEXCEPTION) Then
                '省略時は、エラー型として返します。https://learn.microsoft.com/ja-jp/office/vba/excel/Concepts/Cells-and-Ranges/cell-error-values
                '※例外の種類は一律「#VALUE!」扱いとします。あくまでも「JavaScript実行中に例外が起きた」という目印として判断してください
                jsEval = CVErr(xlErrValue)
                printMsg WARN_, "An exception occurred while executing JavaScript: " & strLogValue & "　It was returned as an error type.", FromProcedureName
            Else
                '引数:IFEXCEPTION が指定されてる場合はそっちを優先して返します
                If IsObject(IFEXCEPTION) Then Set jsEval = IFEXCEPTION Else jsEval = IFEXCEPTION
                printMsg WARN_, "An exception occurred while executing JavaScript: " & strLogValue & "　The argument `IFEXCEPTION` was specified, so it was returned.", FromProcedureName
            End If
        End If

        Exit Function
    End If

    'If the return type has a specific type we can return the result -------------------------------------------
    Dim tmp
    If JSresultAnalysis(scriptEvaluateResult.NodeKey("result"), tmp) Then
        Set jsEval = tmp
    Else
        jsEval = tmp
    End If
    '-----------------------------------------------------------------------------------------------------------

    'Output to Immediate Window if enabled -------------------------------------------------------------------
    ' VarTypeを使って型ごとにログ用の文字列を作る
    Select Case VarType(jsEval)
        Case vbDouble
            strLogValue = "[Double] " & jsEval

        Case vbString
            strLogValue = "[String] " & jsEval

        Case vbBoolean
            strLogValue = "[Boolean] " & jsEval

        Case vbDate
            strLogValue = "[Date] " & jsEval

        Case Else
            strLogValue = "[" & TypeName(jsEval) & "]"
    End Select

    'JavaScriptコードの実行結果の改行を消して、長すぎる場合はカットする（整形）
    strLogValue = LogFormat(strLogValue, 200)

    ' 最終的なデバッグ出力　※実行scriptも、100文字以内にしておく
    printMsg Debug_, "Script executed: " & LogFormat(JavaScriptStr, 100) & " | Returned: " & strLogValue, FromProcedureName
   '---------------------------------------------------------------------------------------------------------

End Function

'***************************************************************************************************
'* 機能　　：JavaScript実行で返ったJsonDictionaryから、結果値`value`だけを取り出すように試みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：scriptRemoteValue  `script.RemoteValue`でのJson-Nodeを指定します。　https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#type-RemoteObject
'            ReturnValue        結果値。参照渡しによる実質返り値です。`primitive values` OR `Object` 両方が想定されます。
'
'* 返り値　：ReturnValueがObjectであるかの論理値。`ReturnValue`がObject型であるかを呼び出し側で正しく代入するために必要です。
'---------------------------------------------------------------------------------------------------
'* 注意事項：返す優先順位は、`script.Handle` > `value`です
'***************************************************************************************************
Private Function JSresultAnalysis(scriptRemoteValue As BiDiCDPJson, ByRef ReturnValue As Variant) As Boolean
    Const FromProcedureName As String = ThisClassName & ".JSresultAnalysis"


    '------------------------ 1. 特定項目有無チェック ------------------------
    ' 格納型情報を取得
    Dim ValueType   As String: ValueType = scriptRemoteValue.StringKey("type")

    '`script.Handle`があったらそれを優先に返します。
    '※type=object系で、`resultOwnership`:="root"時、ここを通ります
    If scriptRemoteValue.ExistsKey("handle") Then
        ReturnValue = scriptRemoteValue.StringKey("handle")
        printMsg info_, "JS-Return type: " & ValueType, FromProcedureName
        printMsg info_, "It is returned as script.Handle.", FromProcedureName
        printMsg Debug_, "handle: " & ReturnValue, FromProcedureName

        Exit Function
    End If


    '------------------------ 2. JsonNodeの取得を試みる ------------------------
    '参考：https://w3c.github.io/webdriver-bidi/#type-script-RemoteValue
    '以降は、`resultOwnership:="none"`でないと通りません
    printMsg info_, "JS-Return type: " & ValueType, FromProcedureName

    Select Case ValueType
        Case "symbol", "function", "weakmap", "weakset", "generator", "error", "proxy", "promise", "typedarray", "arraybuffer"
            '`value`がないので、`Nothing`として返却
            Set ReturnValue = Nothing
            printMsg WARN_, "`" & ValueType & "` cannot be returned as a `value`.", FromProcedureName

        Case "array", "object", "map", "set", "nodelist", "htmlcollection", "window", "regexp", "node"
            'Json文字列として生成されるので、Node形式のObjectとして返却
            Set ReturnValue = scriptRemoteValue.NodeKey("value")

        Case "undefined"
            ' JSの undefined は VBAでは Empty が一番近い
            ReturnValue = Empty

        Case "null"
            ReturnValue = Null

        Case "string"
            ' 文字列として認識するように返す
            ReturnValue = scriptRemoteValue.StringKey("value")

        Case "number"
            '1. 一旦おまかせで取る
            ReturnValue = scriptRemoteValue("value")

            '2. VBAで扱えない特殊な"number"の場合は、警告を出し、「#NUM!」で返します
            If VarType(ReturnValue) = vbString Then
                printMsg WARN_, "This `number` can't be processed in VBA. It returns `#NUM!` instead. Actual JS result: " & ReturnValue, FromProcedureName
                ReturnValue = CVErr(xlErrNum)
            Else
                ReturnValue = CDbl(ReturnValue)
            End If

        Case "boolean"
            ' boolean型として認識するように返す
            ReturnValue = scriptRemoteValue.BoolKey("value")

        Case "bigint"
            ' BigInt は JSON に変換できない巨大整数のため、値を文字列としてそのまんま返却します
            ReturnValue = scriptRemoteValue.StringKey("value")
            printMsg WARN_, "Unserializable BigInt detected. Returning as String: " & ReturnValue, FromProcedureName

        Case "date"
            'BiDiの場合、`date`は文字列として返却されるため、VBAの`date`として返してあげる
            ReturnValue = UtcConverter.ParseIso(scriptRemoteValue.StringKey("value"))

        Case Else
            'ドキュメント化されてない型があったら、安全のため。止めておく
            printMsg ERROR_, "This data type cannot be handled in VBA. type: " & ValueType, FromProcedureName, doRaiseError:=True

    End Select

    'SETステートメントするかの判定を用意する
    JSresultAnalysis = IsObject(ReturnValue)

End Function

'***************************************************************************************************
'* 機能　　：ログメッセージ用に整形を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：logMes     整形したいログメッセージ
'            lengthMes  メインメッセージの上限文字数
'
'* 返り値　：lengthMes を超える場合　：「[上限文字数までのメッセージ内容]...  (Total: [文字数] chars)」。例：neko!umineko!ki... (Total: 1000 chars)
'                      を超えない場合：基本、logMes のまんまですが、改行コードは半角スペースに変換します。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：下記の整形を行います
'            ・改行コードを半角スペースに変換
'            ・`lengthMes`を超える文字数が来たら、合計文字数情報を付与して、返却
'
'* 注意事項：`lengthMes`が長すぎると、一部の箇所で使用する際、エラーになることがあります。
'            ・`Err.Raise`の`description`に長文を渡すと、「スタック領域不足」エラーになります
'***************************************************************************************************
Private Function LogFormat(ByVal logMes As String, lengthMes As Long) As String
    '1. 改行コードを半角スペースに置換
    logMes = Replace(logMes, vbNewLine, " ")

    '2. 指定文字数オーバーの場合はカットし、カット前の文字数情報を付与して返却
    Dim originalLength As Long: originalLength = Len(logMes)
    If originalLength > lengthMes Then LogFormat = Left(logMes, lengthMes) & "... (Total: " & originalLength & " chars)" Else LogFormat = logMes
End Function



'***************************************************************************************************
'                           ■■■ 各種BiDiヘルパーコマンド ■■■
'***************************************************************************************************
'* 機能　　：URL遷移を行います
'***************************************************************************************************
Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
    '1. 待機モードの設定
    Dim targetState As String
    Select Case till
        Case isLoading: targetState = "none"
        Case isInteractive: targetState = "interactive"
        Case isComplete: targetState = "complete"
    End Select

    '2. パラメーター組み立て
    Dim BiDiparams As New Dictionary
    BiDiparams.Add "url", strURL
    BiDiparams.Add "wait", targetState

    '3. 実行
    ExecuteBiDi "browsingContext.navigate", BiDiparams
End Sub

'***************************************************************************************************
'* 機能　　：初回の`browsingContext.contextCreated`にある`context`を取り出します
'***************************************************************************************************
Private Sub attachToContext()
    Const FromProcedureName As String = ThisClassName & ".attachToContext"


    '1. 箱を用意
    Dim FirstTabs As Collection: Set FirstTabs = New Collection

    '2. その箱を渡す
    Set m_BiDiMode.RunCollect_InfoList = FirstTabs

    '3. 購読処理
    m_BiDiMode.EnableDiscoverContexts = True

    '4. 初回購読は直ぐに、`browsingContext.contextCreated`が来るので、それを拝見
    Dim tmp
    Dim FirstTab As BiDiCDPJson
    For Each tmp In FirstTabs
        '4-1. Jsonパース処理
        Set FirstTab = BiDiCDPJson.Parse(CStr(tmp)).NodeKey("params")

        '4-2. 純粋な「新しいタブ」または「新しいウィンドウ」か判定し、そうであれば、`context`を適用
        If IsNull(FirstTab("parent")) Then
            With InfoList
                .context = FirstTab.StringKey("context")
                .Url = FirstTab("url")
            End With

            '4-3. Excelに記録
            m_BiDiMode.serializeMainTab = InfoList.context
            printMsg info_, "I was able to obtain a valid `context` from `browsingContext.contextCreated`. → " & InfoList.context, FromProcedureName
            Exit For
        End If
    Next

    '5. 後始末
    Set m_BiDiMode.RunCollect_InfoList = Nothing

    '6. `context`が取れてない場合は停止
    If StrPtr(InfoList.context) = 0 Then printMsg ERROR_, "A valid `context` could not be obtained from `browsingContext.contextCreated`.", FromProcedureName, doRaiseError:=True
End Sub



'***************************************************************************************************
'                                   ■■■ 黒魔術 ■■■
'***************************************************************************************************
'* 機能　　：現在のContext(タブ等)に対して、`CDPContext`オブジェクトに変換します
'---------------------------------------------------------------------------------------------------
'* 返値    ：このClassで設定中の`context`に紐づく、`CDPContext`オブジェクト　※変換失敗時は`Nothing`で返ります
'---------------------------------------------------------------------------------------------------
'* 詳細情報：`browsingContext.getTree`等で得れる`context`は、`targetId`と一緒である仕様を活かした変換術です
'* 注意事項：`goog:cdp.getSession`による`sessionId`共有は不可能みたいなので、`Target.attachToTarget`で新しいSessionIdとして制御します
'***************************************************************************************************
Public Function ConvertToCDPContext() As CDPContext
    Const FromProcedureName As String = ThisClassName & ".ConvertToCDPContext"


    '1. context に紐づく、新規`SessionId`の発行を試みる
    Dim newSessionID As String: newSessionID = m_BiDiCore.DirectControlCDP.attachToTab(InfoList.context)

    '2. 発行失敗時はエラーで即抜け
    If LenB(newSessionID) = 0 Then m_BiDiMode.printMsg WARN_, "Failed to issue a SessionID for CDP from the WebDriver BiDi context. > " & m_BiDiCore.DirectControlCDP.LastCDPJsonError("message"), FromProcedureName: Exit Function

    '3. 無事発行出来たら、変換手続き
    '------ CDPContext.clsとして返却する準備 ------
    m_BiDiCore.DirectControlCDP.EnableDiscoverTargets = True    '`CDPContext.cls`としての制御なので、カウントしておく
    Set ConvertToCDPContext = New CDPContext
    With ConvertToCDPContext
        '4. BiDiブラウザのパイプハンドル情報を渡す
        Set .InheritanceCDPCore = m_BiDiCore.DirectControlCDP.InheritanceCDPCore

        '5. BiDiブラウザ情報を渡す
        Set .InheritanceCDPBrowser = m_BiDiCore.DirectControlCDP

        '6. 先ほど発行したSessionIDを渡す
        .CurrentSessionID = newSessionID

        '7. ステータスの更新
        .getSessionInfo
    End With
End Function



'***************************************************************************************************
'                              ■■■ PUBLIC PROPERTIES ■■■
'***************************************************************************************************
'* 機能　　：`context`(タブ等)が閉鎖されたか確認します
'---------------------------------------------------------------------------------------------------
'* 返値    ：論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：このClassの`targetId`に対する`browsingContext.contextDestroyed`を受け取ると、`True`になります
'***************************************************************************************************
Property Get isContextClosed() As Boolean
    isContextClosed = contextDestroyed
End Property



'***************************************************************************************************
'                        ■■■ WebDriverBiDiCore.cls 直通メソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    m_BiDiCore.printMsg LogLevel_, LogID, strMsg, From, isHeader, doRaiseError
End Sub



'***************************************************************************************************
'                              ■■■ 複製/継承用プロパティ ■■■
'***************************************************************************************************
'* 機能　　：このClass内で定義してる`WebDriverBiDiMode.cls`を受け取ったり継承させます
'***************************************************************************************************
Property Get InheritanceWebDriverBiDiMode() As WebDriverBiDiMode
    Set InheritanceWebDriverBiDiMode = m_BiDiMode
End Property
Property Set InheritanceWebDriverBiDiMode(Target As WebDriverBiDiMode)
    Set m_BiDiMode = Target
End Property

'***************************************************************************************************
'* 機能　　：このClass内で定義してる`WebDriverBiDiCore.cls`を受け取ったり継承させます
'***************************************************************************************************
Property Get InheritanceWebDriverBiDiCore() As WebDriverBiDiCore
    Set InheritanceWebDriverBiDiCore = m_BiDiCore
End Property
Property Set InheritanceWebDriverBiDiCore(Target As WebDriverBiDiCore)
    Set m_BiDiCore = Target
End Property

'***************************************************************************************************
'* 機能　　：`browsingContext.getTree`等で得た`context`を適用させます
'***************************************************************************************************
Property Let context(param As String)
    InfoList.context = param
End Property
Property Get context() As String
    context = InfoList.context
End Property



'***************************************************************************************************
'                       ■■■ このClassオブジェクト生成/破棄時のアクション ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    'Assign an objID for easy debug identification
    LogID = "BTID" & Format(Rnd * 1000, "000")
End Sub

Private Sub Class_Terminate()
    m_BiDiMode.EnableDiscoverContexts = False
End Sub
