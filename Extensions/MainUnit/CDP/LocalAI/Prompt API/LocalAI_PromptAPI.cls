VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalAI_PromptAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           ブラウザのローカルAI(Prompt API)に対する機能拡張を提供するクラスファイルです
'   実装内容は全て、2026/4/15時点の「https://learn.microsoft.com/ja-jp/microsoft-edge/web-platform/prompt-api」を参考に記述してます
'           まだまだ、試験的扱いのためある日突然、動作しなくなることをご了承ください
'   現時点では、全てのサイトで動作することができません。「edge://」等と言った限られた場所のみ動作します
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部クラスオブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1
Private JsonConv                    As New WebJsonConverter

'内部使用用
Private Const ThisClassName                 As String = "LocalAI_PromptAPI" 'ログのトレース用
Private Const InitialAllocationStrBuffer    As Long = 2 ^ 6                 '初期バッファー上限

'`Runtime.bindingCalled`の`Name`定義用
Private Const BindingName01         As String = "sendDLprogressNotification"    'AIモデルデータのDL進捗用
Private Const BindingName02         As String = "pushAiStream"                  'promptStreaming の受け取り用
Private Const BindingName02_EOFcode As String = 30                              'promptStreaming 終端目印文字コード

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'responseをデバッグログに出すかの調整
Private EnableBrowserEvents         As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private ProgressValue               As Double       'AIモデルデータのDL進捗値
Private StreamingEOF                As Boolean      'ストリーミング受信が終わったかのサイン
Private StreamingCol                As Collection   'Streaming の蓄積用

'Prompt APIに関する設定群
Private Type PromptAPI
    objectId        As String       'AIへのセッションに対するObjectID
    initialPrompts  As Collection   'プロンプトが表示されたときにモデルで引き続き使用するユーザー/アシスタント操作の例を提供する
    topK            As Long         'モデルの出力分布から上位Kの次期トークンを選択するサンプリング手法です。 モデルはその後、このトークンのサブセットからサンプルを抽出して、シーケンス内の次のトークンを生成します。この方法は可能です 生成されたテキストのランダム性を低減し、より一貫性のあるものにするのに役立ちます
    temperature     As Double       'モデルの出力のランダム性を制御するパラメータです。低温(例:0.1) モデルをより決定論的で集中的にする一方で、温度が高くなる(例:1.0)と、よりランダムな状態にする。 創造的。温度を調整することで、一貫性と創造性の適切なバランスを見つけるのに役立ちます 生成されたテキスト
End Type
Private SettingAI As PromptAPI

'分割記述されたJavaScript配列管理用
Private Type StateFragmentedJS
    strBuffer()     As String   '一時的に蓄積するための配列です
    Index           As Long     '埋まってる位置
    expansionCount  As Long     '拡張確保回数
End Type
Private JSBox As StateFragmentedJS



'***************************************************************************************************
'                               ■■■ 公開プロパティ ■■■
'***************************************************************************************************
'* 機能　　：非同期イベント処理を行うかのスイッチングです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：スコープ範囲の調整にどうぞ。名称は、「Application.EnableEvents」から取りました
'***************************************************************************************************
Public Property Get EnableEvents() As Boolean
    EnableEvents = EnableBrowserEvents
End Property
Public Property Let EnableEvents(ByVal v As Boolean)
    EnableBrowserEvents = v
End Property

'***************************************************************************************************
'* 機能　　：Prompt APIへのObjectIDを取得/再設定をします
'***************************************************************************************************
Public Property Get objectidPromptAPI() As String
    objectidPromptAPI = SettingAI.objectId
End Property
Public Property Let objectidPromptAPI(AIobjectis As String)
    SettingAI.objectId = AIobjectis
End Property

'***************************************************************************************************
'* 機能　　：AIモデルデータのDL進捗値を取得します
'***************************************************************************************************
Public Property Get DLProgressValue() As Double
    DLProgressValue = ProgressValue
End Property

'***************************************************************************************************
'* 機能　　：promptStreamingにて蓄積したのを呼び出す度に先頭だけ取り出して、先頭を削除します
'***************************************************************************************************
Public Property Get StreamingTopTake() As String
    If StreamingCol.Count = 0 Then Exit Property

    StreamingTopTake = StreamingCol(1)
    StreamingCol.Remove (1)
End Property

'***************************************************************************************************
'* 機能　　：promptStreamingにて蓄積したのをコレクションとして丸ごと受け取ります
'***************************************************************************************************
Public Property Get StreamingColTake() As Collection
    Set StreamingColTake = StreamingCol
End Property

'***************************************************************************************************
'* 機能　　：promptStreamingのEOFサインを取得します
'***************************************************************************************************
Public Property Get StreamingEOFExist() As Boolean
    StreamingEOFExist = StreamingEOF
End Property

'***************************************************************************************************
'* 機能　　：SystemPrompt の設定をします
'---------------------------------------------------------------------------------------------------
'* GET返値 ：現時点での`SystemPrompt`の内容
'* LET引数 ：`SystemPrompt`内容
'---------------------------------------------------------------------------------------------------
'* 機能説明：`vbnullstring`で渡すと、削除扱いとします
'***************************************************************************************************
Public Property Let setSystemPrompt(SystemPrompt As String)
    Const FromProcedureName As String = ThisClassName & ".setSystemPrompt"


    Const UserRoleName As String = "system"

    '1. システムプロンプトを作成
    Dim prompt As New Dictionary
    prompt.Add "role", UserRoleName
    prompt.Add "content", SystemPrompt

    '`initialPrompts`の構造を解析
    '`initialPrompts`にプロンプト情報が既にあるか？
    If SettingAI.initialPrompts.Count > 0 Then
        '既に`system`ロールがあるか？
        If SettingAI.initialPrompts(1)("role") = UserRoleName Then
            '有効な値が渡されてるか？
            If StrPtr(SystemPrompt) Then
                '置き換え
                SettingAI.initialPrompts(1)("content") = SystemPrompt
                printMsg Debug_, "The existing system prompt has been replaced.", FromProcedureName
            Else
                '削除
                SettingAI.initialPrompts.Remove 1
                printMsg Debug_, "The system prompt was removed because it was passed as `null`.", FromProcedureName
            End If
        Else
            '有効な値が渡されてない場合は、何もしない
            If StrPtr(SystemPrompt) = 0 Then printMsg Debug_, "The system prompt was passed as `null`, so nothing happens.", FromProcedureName: Exit Property

            '先頭部分に置くようにする
            SettingAI.initialPrompts.Add prompt, Before:=1
            printMsg Debug_, "Added a system prompt to the beginning of the existing `initialPrompts`.", FromProcedureName
        End If
    Else
        '有効な値が渡されてない場合は、何もしない
        If StrPtr(SystemPrompt) = 0 Then printMsg Debug_, "The system prompt was passed as `null`, so nothing happens.", FromProcedureName: Exit Property

        'まだ何もないので新規追加
        SettingAI.initialPrompts.Add prompt
        printMsg Debug_, "Added a new system prompt.", FromProcedureName
    End If
End Property
Public Property Get setSystemPrompt() As String
    Const FromProcedureName As String = ThisClassName & ".setSystemPrompt"


    Const UserRoleName As String = "system"

    'initialPrompts の構造を解析
    'Collection の格納数チェック
    If SettingAI.initialPrompts.Count > 0 Then
        If SettingAI.initialPrompts(1)("role") = UserRoleName Then
            '既に`system`ロールがあるならそれを返す
            setSystemPrompt = SettingAI.initialPrompts(1)("content")
        End If
    End If
End Property

'***************************************************************************************************
'* 機能　　：initialPrompts を使用した N ショット プロンプトの設定をします
'---------------------------------------------------------------------------------------------------
'* GET返値 ：現時点での`initialPrompts`の内容　※Json-Stringで返します
'* LET引数 ：`content`内容
'---------------------------------------------------------------------------------------------------
'* 詳細説明：LETするたびに、"user"→"assistant"→"user"→... としてどんどん追加されます
'***************************************************************************************************
Public Property Let addInitialPrompt(contentInitialPrompt As String)
    '1. システムプロンプト有無チェック
    Const UserRoleName As String = "system"
    Dim SystemPromptExist As Boolean
    If SettingAI.initialPrompts.Count > 0 Then
        If SettingAI.initialPrompts(1)("role") = UserRoleName Then
            '既に`system`ロールがあるなら、フラグ適用
            SystemPromptExist = True
        End If
    End If

    '2. プロンプトを作成
    Dim prompt As New Dictionary
    prompt.Add "role", IIf((SettingAI.initialPrompts.Count + SystemPromptExist) Mod 2, "assistant", "user")
    prompt.Add "content", contentInitialPrompt

    '3. 追加
    SettingAI.initialPrompts.Add prompt
End Property
Public Property Get addInitialPrompt() As String
    'Json文字列で返却
    addInitialPrompt = JsonConv.ConvertToJson(SettingAI.initialPrompts)
End Property



'***************************************************************************************************
'                               ■■■ 便利プロシージャ ■■■
'***************************************************************************************************
Public Sub clearInitialPrompts()
    Const FromProcedureName As String = ThisClassName & ".clearInitialPrompts"


    Set SettingAI.initialPrompts = New Collection
    printMsg Debug_, "I have cleared all `initialPrompts`.", FromProcedureName
End Sub



'***************************************************************************************************
'                     ■■■ メインとなる機能拡張部分 非同期イベント編 ■■■
'***************************************************************************************************
'* 機能　　：非同期イベントを処理したいときにご利用ください。あなたのイメージをコードに落とし込む場所の1つです
'***************************************************************************************************
Private Sub ex_CDPBrowser_detectionCDPEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPBrowser_detectionCDPEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. 処理したい対象の非同期イベント名を、各Case に記述
    Select Case methodName
        Case "Runtime.bindingCalled"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName

            '1. PayloadのName部分を取得
            Dim NamePayload As String: NamePayload = ExtractName(EventRawJson)

            '2. `Name`値に応じた処理
            Dim tmp As String
            Select Case NamePayload
                Case BindingName01
                    'PayLoad部分を取り出します
                    tmp = ExtractPayload(EventRawJson)
                    printMsg Debug_, "payload :" & tmp, FromProcedureName

                    '進捗値を更新します
                    ProgressValue = tmp
                    printMsg Debug_, "Downloading... " & ProgressValue & "%", FromProcedureName

                Case BindingName02
                    'デコード済みのPayLoad部分を取り出します
                    tmp = JsonConv.ParseJson(EventRawJson)("params")("payload")
                    printMsg Debug_, "payload :" & tmp, FromProcedureName

                    'ストリームに追加
                    StreamingCol.Add tmp
                    'Debug.Print tmp;

                    'EOFサインが来たらフラグを更新
                    If tmp = Chr(BindingName02_EOFcode) Then StreamingEOF = True

                Case Else
                    printMsg Debug_, "We detected `Runtime.bindingCalled` instances other than AI. Name :" & NamePayload, FromProcedureName
            End Select

        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName
            Exit Sub

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
'                               ■■■ Prompt API - 準備編 ■■■
'***************************************************************************************************
'* 機能　　：API が有効かどうかを確認する
'---------------------------------------------------------------------------------------------------
'* 返り値  ：利用可否論理値
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Prompt API を使用する前に、LanguageModel オブジェクトの存在をテストして API が有効になっていることを確認します
'***************************************************************************************************
Public Function IsPromptApiAvailable() As Boolean
    '1. 所定のJavaScriptコードを記述
    Dim JsCode As String
    JsCode = "!LanguageModel"

    '2. CDPのJavaScript実行メソッドを使用し、ブラウザ内で、JavaScriptコードを実行
    Dim tmp
    tmp = ex_CDPBrowser.jsEval(JsCode)

    '3. 結果を整理
    'そもそも、用意されてないか？
    If IsError(tmp) Then
        '「Uncaught ReferenceError: LanguageModel is not defined」と出るので、`False`とする
        IsPromptApiAvailable = False
    Else
        '機能があれば、普通の論理値なので、それを返却
        '※元のJSが、否定演算子を使用してるため、結果も反転しておく
        IsPromptApiAvailable = Not tmp
    End If

End Function

'***************************************************************************************************
'* 機能　　：モデルを使用できるかどうかを確認する
'---------------------------------------------------------------------------------------------------
'* 返り値  ：現在は、下記の4種類が返ります
'            ・"unavailable"
'            ・"downloadable"
'            ・"downloading"
'            ・"available"
'---------------------------------------------------------------------------------------------------
'* 詳細説明：言語モデルとモデル ランタイムが Microsoft Edge によってダウンロードされたかを確認します
'* 注意事項：これを行う前に`IsPromptApiAvailable`の確認をしないとエラー落ちします
'***************************************************************************************************
Public Function CheckAvailability() As String
    '1. 所定のJavaScriptコードを記述
    Dim JsCode As String
    JsCode = "(async () => { return await LanguageModel.availability(); })()"

    '2. CDPのJavaScript実行メソッドを使用し、ブラウザ内で、JavaScriptコードを実行し、結果を返却
    CheckAvailability = ex_CDPBrowser.jsEval(JsCode, awaitPromise:=True, StopException:=True)
End Function

'***************************************************************************************************
'* 機能　　：モデルをDLし、進捗表示機能を有効化します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：実行に成功すると、`True`を返します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`monitor`オプションを使用して、モデルのダウンロードの進行状況を追跡できます
'***************************************************************************************************
Public Function ModelDownloadProgress() As Boolean
    With ex_CDPBrowser
        '1. ブラウザ→VBA への受け口を用意
        Dim params As New Dictionary
        params.Add "name", BindingName01
        .invokeMethod "Runtime.addBinding", params

        '2. JavaScript側でその窓口を叩くコードを準備
        Dim jsProgressCode As String
        jsProgressCode = "function(funcName) { " & _
           "  LanguageModel.create({ " & _
           "    monitor: m => { " & _
           "      m.addEventListener('downloadprogress', event => { " & _
           "        const progress = (event.loaded / event.total) * 100; " & _
           "        if (typeof window[funcName] === 'function') { " & _
           "          window[funcName](progress.toFixed(2)); " & _
           "        } " & _
           "      }); " & _
           "    } " & _
           "  }); " & _
           "  return true; " & _
           "}"

        '3. 引数を設定
        Dim args As New Collection
        Dim arg1 As New Dictionary: arg1.Add "value", BindingName01: args.Add arg1

        '4. Run!
        ModelDownloadProgress = .jsEval(jsProgressCode, contextId:=1, objectArguments:=args, userGesture:=True, StopException:=True)
    End With
End Function



'***************************************************************************************************
'                          ■■■ Prompt API - 一度切りトーク編 ■■■
'***************************************************************************************************
'* 機能　　：履歴を持たない一回限りのAIとのやり取りを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：チャットの返答内容
'* 引数　　：ChatString     チャット内容
'---------------------------------------------------------------------------------------------------
'* 注意事項：・必要に応じてタイムアウトに余裕を持たせる必要があります
'            ・システム プロンプト は、非対応です
'***************************************************************************************************
Public Function instantSession(ChatString As String) As String
    '1. JSコードの組み立て
    ' 1. セッション作成 2. 質問送信 3. セッション破棄 を一気に行います
    Dim JsCode As String
    JsCode = "async function(chatText) {" & _
           "  try {" & _
           "    const session = await LanguageModel.create();" & _
           "    const result = await session.prompt(chatText);" & _
           "    session.destroy();" & _
           "    return result;" & _
           "  } catch (e) {" & _
           "    return 'Error: ' + e.message;" & _
           "  }" & _
           "}"

    '2. 引数を設定
    Dim args As New Collection
    Dim arg1 As New Dictionary: arg1.Add "value", ChatString: args.Add arg1

    '3. 実行と結果取得
    instantSession = ex_CDPBrowser.jsEval(JsCode, contextId:=1, objectArguments:=args, awaitPromise:=True)
End Function

'***************************************************************************************************
'* 機能　　：履歴を持たない一回限りのAIとのやり取りを行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：ChatString     チャット内容
'---------------------------------------------------------------------------------------------------
'* 詳細説明：上記のプロシージャと違いこちらは、`Streaming`という滝みたいに流れる仕組みとなります
'* 注意事項：システム プロンプト は、非対応です
'***************************************************************************************************
Public Sub instantStreamingSession(ChatString As String)
    '0. ステータスをリセット
    StreamingEOF = False
    Set StreamingCol = New Collection

    '1. ブラウザ→VBA への受け口を用意
    Dim params As New Dictionary
    params.Add "name", BindingName02
    ex_CDPBrowser.invokeMethod "Runtime.addBinding", params

    '2. JSコードの組み立て
    ' window[funcName] で動的に呼び出すことで引数化が可能です'
    Dim JsCode As String
    JsCode = "async function(funcName, chatText, eofCode) {" & _
               "  const session = await LanguageModel.create();" & _
               "  const stream = session.promptStreaming(chatText);" & _
               "  for await (const chunk of stream) {" & _
               "    if (typeof window[funcName] === 'function') {" & _
               "      window[funcName](chunk);" & _
               "    }" & _
               "  }" & _
               "  if (typeof window[funcName] === 'function') {" & _
               "    window[funcName](String.fromCharCode(eofCode));" & _
               "  }" & _
               "  session.destroy();" & _
               "}"

    '3. 引数を設定
    Dim args As New Collection
    Dim arg1 As New Dictionary: arg1.Add "value", BindingName02:            args.Add arg1
    Dim arg2 As New Dictionary: arg2.Add "value", ChatString:               args.Add arg2
    Dim arg3 As New Dictionary: arg3.Add "value", BindingName02_EOFcode:    args.Add arg3

    '非同期で実行開始。結果を待たずに次の処理へ
    ex_CDPBrowser.jsEval JsCode, contextId:=1, objectArguments:=args
End Sub



'***************************************************************************************************
'                          ■■■ Prompt API - 継続トーク編 ■■■
'***************************************************************************************************
'* 機能　　：セッションを作成（初期化）する
'---------------------------------------------------------------------------------------------------
'* 返り値  ：このセッションの`ObjectID`
'---------------------------------------------------------------------------------------------------
'* 機能説明：このクラスに持ってるプロパティメソッドを使って設定したのを元に、セッションを作成します
'***************************************************************************************************
Public Sub createSession()
    'JSコードの組み立て
    Dim JsCode As String
    JsCode = "async function(promptsJson, tK, temp) {" & _
             "  const options = { " & _
             "    initialPrompts: JSON.parse(promptsJson), " & _
             "    topK: tK, " & _
             "    temperature: temp " & _
             "  };" & _
             "  return await LanguageModel.create(options);" & _
             "}"

    '引数を設定
    Dim args As New Collection
    With SettingAI
        Dim arg1 As New Dictionary: arg1.Add "value", JsonConv.ConvertToJson(.initialPrompts):  args.Add arg1
        Dim arg2 As New Dictionary: arg2.Add "value", .topK:                                    args.Add arg2
        Dim arg3 As New Dictionary: arg3.Add "value", .temperature:                             args.Add arg3
    End With

    'Run
    SettingAI.objectId = ex_CDPBrowser.jsEval(JsCode, contextId:=1, objectArguments:=args, awaitPromise:=True, StopException:=True)
End Sub

'***************************************************************************************************
'* 機能　　：上記の作ったセッションを元に、チャットを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：チャットの返答内容
'* 引数　　：ChatString     チャット内容
'---------------------------------------------------------------------------------------------------
'* 注意事項：・必要に応じてタイムアウトに余裕を持たせる必要があります
'            ・当然ですが、`createSession`をまず行うこと
'***************************************************************************************************
Public Function advancePrompt(ChatString As String) As String
    ' 1. JSコードの組み立て
    ' ※Runtime.callFunctionOn では、指定した objectId の実体が 'this' として扱われます
    Dim JsCode As String
    JsCode = "async function(text) { " & _
             "  return await this.prompt(text);" & _
             "}"

    ' 2. 引数を設定
    Dim args As New Collection
    Dim arg1 As New Dictionary: arg1.Add "value", ChatString:  args.Add arg1

    ' 3.Run
    advancePrompt = ex_CDPBrowser.jsEval(JsCode, SettingAI.objectId, objectArguments:=args, awaitPromise:=True, StopException:=True)
End Function

'***************************************************************************************************
'* 機能　　：上記の作ったセッションを元に、チャットを行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：ChatString     チャット内容
'---------------------------------------------------------------------------------------------------
'* 詳細説明：上記のプロシージャと違いこちらは、`Streaming`という滝みたいに流れる仕組みとなります
'***************************************************************************************************
Public Sub advancePromptStreaming(ChatString As String)
    '0. ステータスをリセット
    StreamingEOF = False
    Set StreamingCol = New Collection

    '1. ブラウザ→VBA への受け口を用意
    Dim params As New Dictionary
    params.Add "name", BindingName02
    ex_CDPBrowser.invokeMethod "Runtime.addBinding", params

    '2. JSコードの組み立て
    ' window[funcName] で動的に呼び出すことで引数化が可能です'
    Dim JsCode As String
    JsCode = "async function(funcName, chatText, eofCode) {" & _
               "  const stream = this.promptStreaming(chatText);" & _
               "  for await (const chunk of stream) {" & _
               "    if (typeof window[funcName] === 'function') {" & _
               "      window[funcName](chunk);" & _
               "    }" & _
               "  }" & _
               "  if (typeof window[funcName] === 'function') {" & _
               "    window[funcName](String.fromCharCode(eofCode));" & _
               "  }" & _
               "}"

    '3. 引数を設定
    Dim args As New Collection
    Dim arg1 As New Dictionary: arg1.Add "value", BindingName02:            args.Add arg1
    Dim arg2 As New Dictionary: arg2.Add "value", ChatString:               args.Add arg2
    Dim arg3 As New Dictionary: arg3.Add "value", BindingName02_EOFcode:    args.Add arg3

    '非同期で実行開始。結果を待たずに次の処理へ
    ex_CDPBrowser.jsEval JsCode, SettingAI.objectId, objectArguments:=args
End Sub


'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.invokeMethod`といった記述を防ぐ措置となります。
'---------------------------------------------------------------------------------------------------
'* 注意事項：・メインの`CDPBrowser`メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておくことをお勧めします
'            ・各詳細は、`CDPBrowser`の方をご覧ください
'***************************************************************************************************
'負荷を抑えるときにどうぞ
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_CDPBrowser.sleep seconds
End Sub

'ログメッセージを出すときにどうぞ
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

'CDP実行に必要です
Private Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set invokeMethod = ex_CDPBrowser.invokeMethod(methodName, params, alwaysBrowserContext, StopError)
End Function

'非同期イベント発火(RaiseEvent)に必要です
Private Sub TakeEvents(Optional destruction As Boolean)
    ex_CDPBrowser.TakeEvents destruction
End Sub



'***************************************************************************************************
'                             ■■■ ヘルパープロシージャ ■■■
'***************************************************************************************************
Private Property Let FragmentedJS(pieceData As String)
    Const FromProcedureName As String = ThisClassName & ".FragmentedJS"


    With JSBox
        '1. 上限を越したら、拡張する(高負荷ポイント)
        If .Index >= InitialAllocationStrBuffer * .expansionCount Then
            .expansionCount = .expansionCount + 1
            ReDim Preserve .strBuffer(InitialAllocationStrBuffer * .expansionCount - 1)
            printMsg Debug_, "The buffer has been expanded to " & InitialAllocationStrBuffer * .expansionCount, FromProcedureName
        End If

        '2. 配列に入れる
        .strBuffer(.Index) = pieceData
        printMsg Trace_, "Data has been placed in array " & .Index, FromProcedureName

        '3. カウント
        .Index = .Index + 1
    End With
End Property

Private Property Get FragmentedJS() As String
    Const FromProcedureName As String = ThisClassName & ".FragmentedJS"


    '1. 確保中の配列を全部結合する
    printMsg Trace_, "Now we'll begin the buffer merging process...", FromProcedureName
    Dim tmp As String: tmp = Join(JSBox.strBuffer, "")
    printMsg Trace_, "concatenated Raw-string: " & tmp, FromProcedureName

    '2. `vbnullchar`がある位置を先頭から探す
    Dim EndPos As Long: EndPos = searchNull(tmp)

    '3. その目印を元に、1件分返却
    FragmentedJS = Mid(tmp, 1, EndPos - 1)
    printMsg Debug_, "Trim complete. Result: " & FragmentedJS, FromProcedureName

    '4. インデックスリセット
    JSBox.Index = 0
    printMsg Trace_, "Buffer Index reset.", FromProcedureName

End Property

Private Function searchNull(checkString As String) As Long

    Dim lngPos As Long

    If Len(checkString) = 0 Then Exit Function
    lngPos = InStr(1, checkString, vbNullChar, vbBinaryCompare)
    If lngPos > 0 Then
        searchNull = lngPos
    Else
        searchNull = 0
    End If

End Function



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`CDPBrowser.cls`を継承させます
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPBrowser)
    '継承
    Set ex_CDPBrowser = Inheritance
End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    '配列の初期化
    With JSBox
        .expansionCount = 1
        ReDim .strBuffer(InitialAllocationStrBuffer * .expansionCount - 1)
        .Index = 0
    End With

    '初期設定化
    With SettingAI
        Set .initialPrompts = New Collection

        'ブラウザのデフォルト設定に従う
        .topK = 10
        .temperature = 0.7
    End With

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value    'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
    End With
End Sub
