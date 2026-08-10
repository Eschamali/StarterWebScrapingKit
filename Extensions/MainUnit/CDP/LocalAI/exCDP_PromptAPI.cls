VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_PromptAPI"
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
Private ex_CDPContext           As CDPContext
Attribute ex_CDPContext.VB_VarHelpID = -1
Private ex_CDPBrowser           As CDPBrowser
Private WithEvents ex_CDPCore   As CDPCore
Attribute ex_CDPCore.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "exCDP_PromptAPI" 'ログのトレース用

'`Runtime.bindingCalled`の`Name`定義用
Private Const BindingName01         As String = "sendDLprogressNotification"    'AIモデルデータのDL進捗用
Private Const BindingName02         As String = "pushAiStream"                  'promptStreaming の受け取り用
Private Const BindingName02_EOFcode As String = 30                              'promptStreaming 終端目印文字コード

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'responseをデバッグログに出すかの調整
Private EnableBrowserEvents         As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private ProgressValue               As Double       'AIモデルデータのDL進捗値
Private StreamingEOF                As Boolean      'ストリーミング受信が終わったかのサイン
Private CurrentSessionID            As String       '`CDPContext.cls`継承時のSesssionID

'`promptStreaming`のcursor管理
Private Type TextCursor
    strBuffer       As String   '`promptStreaming`から送られたテキストをどんどん蓄積する変数
    length          As Long     '現在のバッファ全体の長さ
    NextCursor      As Long     '次の取り出し開始位置
    EndCursor       As Long     '有効なデータ末尾位置
End Type
Private promptStreaming As TextCursor

'Prompt APIに関する設定群
Private Type PromptAPI
    objectId As String  'AIへのセッションに対するObjectID


End Type
Private SettingAI As PromptAPI



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
'* 機能　　：AIモデルデータのDL進捗値を取得します
'***************************************************************************************************
Public Property Get DLProgressValue() As Double
    DLProgressValue = ProgressValue
End Property

'***************************************************************************************************
'* 機能　　：promptStreamingにて蓄積したのを呼び出す度に、ある分だけ取り出して、次の取り出し位置を更新する機能です
'---------------------------------------------------------------------------------------------------
'* 注意事項：これは主に演出用兼進捗状況に使います。実際の完成データは、後術の`StreamingAllTake`をご利用下さい
'***************************************************************************************************
Public Property Get StreamingTopTake() As String
    With promptStreaming
        '1. データの末尾まで来たら即抜け
        If .NextCursor = .EndCursor Then Exit Property
        
        '2. 現在のcursor情報を基に取り出し
        StreamingTopTake = Mid$(.strBuffer, .NextCursor + 1, .EndCursor - .NextCursor)
        
        '3. cursor情報を更新
        .NextCursor = .EndCursor
    End With
End Property

'***************************************************************************************************
'* 機能　　：promptStreamingにて蓄積した1件のメッセージを受け取ります
'***************************************************************************************************
Public Property Get StreamingAllTake() As String
    With promptStreaming
        StreamingAllTake = Mid$(.strBuffer, 1, .EndCursor)
    End With
End Property

'***************************************************************************************************
'* 機能　　：promptStreamingのEOFサインを取得します
'***************************************************************************************************
Public Property Get StreamingEOFExist() As Boolean
    StreamingEOFExist = StreamingEOF
End Property



'***************************************************************************************************
'                     ■■■ メインとなる機能拡張部分 非同期イベント編 ■■■
'***************************************************************************************************
Private Sub ex_CDPCore_CDPContextEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPBrowser_detectionCDPEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. このコンテキスト(タブ)クラスで設定中の`SessionId`以外も即抜け
    If CurrentSessionID <> sessionID Then Exit Sub

    '2. 処理したい対象の非同期イベント名を、各Case に記述
    Select Case methodName
        Case "Runtime.bindingCalled"
            Dim ResParams As BiDiCDPJson
            Set ResParams = BiDiCDPJson.Parse(RawJson).NodeKey("params")
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & RawJson, FromProcedureName

            '1. PayloadのName部分を取得
            Dim NamePayload As String: NamePayload = ResParams.StringKey("name")

            '2. `Name`値に応じた処理
            Dim tmp As String
            Select Case NamePayload
                'モデルデータのDL進捗
                Case BindingName01
                    'PayLoad部分を取り出します
                    tmp = ResParams.StringKey("payload")
                    printMsg Debug_, "payload :" & tmp, FromProcedureName

                    '進捗値を更新します
                    ProgressValue = tmp
                    printMsg Debug_, "Downloading... " & ProgressValue & "%", FromProcedureName

                '`Streaming`による返却用
                Case BindingName02
                    'デコード済みのPayLoad部分を取り出します
                    tmp = ResParams.StringKey("payload")
                    printMsg Debug_, "payload :" & tmp, FromProcedureName

                    'ストリームに追加
                    StreamingAccumulate = tmp

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
'* 機能　　：promptStreamingからのレスポンスをどんどん蓄積させ、良い感じにcursor管理を行います
'***************************************************************************************************
Private Property Let StreamingAccumulate(str As String)
    Const FromProcedureName As String = ThisClassName & ".StreamingAccumulate"


    'サイズ0の場合は抜ける
    Dim resSize As Long
    resSize = Len(str)
    If resSize <= 0 Then printMsg WARN_, "No promptStreaming was found, so the buffer check will be skipped.", FromProcedureName: Exit Property

    With promptStreaming
        '1. 上限を越しそうになったら、拡張する(高負荷ポイント)
        '※オーバフローしたらこのツールでは対応できません
        If .EndCursor + resSize > .length Then
            .strBuffer = .strBuffer & String$(.length, vbNullChar)
            .length = .length * 2
            printMsg Debug_, "The buffer has been expanded: " & .length, FromProcedureName
        End If
        
        '2. MIDでうまいぐあいに、上書き
        Mid$(.strBuffer, .EndCursor + 1) = str

        '3. カーソル情報を更新
        .EndCursor = .EndCursor + resSize
    End With
End Property



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
    tmp = ex_CDPContext.jsEval(JsCode)

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
    CheckAvailability = ex_CDPContext.jsEval(JsCode, awaitPromise:=True, StopException:=True)
End Function

'***************************************************************************************************
'* 機能　　：モデルをDLし、進捗表示機能を有効化します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`monitor`オプションを使用して、モデルのダウンロードの進行状況を追跡できます
'***************************************************************************************************
Public Function ModelDownloadProgress() As String
    With ex_CDPContext
        '1. ブラウザ→VBA への受け口を用意
        Dim params As New Dictionary
        params.Add "name", BindingName01
        .ExecuteCDP "Runtime.addBinding", params

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
           "  return 'DOWNLOAD_STARTED'; " & _
           "}"

        '3. 引数を設定
        Dim args As New Collection
        Dim arg1 As New Dictionary: arg1.Add "value", BindingName01: args.Add arg1

        '4. Run!
        ModelDownloadProgress = .jsEval(jsProgressCode, objectArguments:=args, userGesture:=True, StopException:=True)
    End With
End Function



'***************************************************************************************************
'                          ■■■ Prompt API - 一度切りトーク編 ■■■
'***************************************************************************************************
'* 機能　　：履歴を持たない一回限りのAIとのやり取りを行います
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
    instantSession = ex_CDPContext.jsEval(JsCode, objectArguments:=args, awaitPromise:=True)
End Function

'***************************************************************************************************
'* 機能　　：履歴を持たない一回限りのAIとのやり取りを行います
'---------------------------------------------------------------------------------------------------
'* 詳細説明：上記のプロシージャと違いこちらは、`Streaming`という滝みたいに流れる仕組みとなります
'* 注意事項：システム プロンプト は、非対応です
'***************************************************************************************************
Public Function instantStreamingSession(ChatString As String) As String
    '0. ステータスをリセット
    StreamingEOF = False
    With promptStreaming
        .EndCursor = 0
        .NextCursor = 0
    End With

    '1. ブラウザ→VBA への受け口を用意
    Dim params As New Dictionary
    params.Add "name", BindingName02
    ex_CDPContext.ExecuteCDP "Runtime.addBinding", params

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
    instantStreamingSession = ex_CDPContext.jsEval(JsCode, objectArguments:=args)
End Function



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

'非同期イベント発火(RaiseEvent)に必要です
Private Sub TakeEvents(Optional destruction As Boolean)
    ex_CDPBrowser.TakeEvents destruction
End Sub



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`CDPCore.cls`等を継承させます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：CDP制御に関する全てのオブジェクトをここで継承させ、タブのSesssionIDも適用させてください
'* 注意事項：・タブの場合は、「CDPContext,CDPBrowser,CDPCore」の3つ、ブラウザの場合は、「CDPBrowser,CDPCore」の2つとなります
'           ・SessionId は、タブで継承する際に使用します
'***************************************************************************************************Friend Sub Init(Inheritance As CDPBrowser)
Friend Sub Init(Inheritance As CDPContext)
    '必要なCDPオブジェクトを継承
    Set ex_CDPContext = Inheritance
    Set ex_CDPBrowser = Inheritance.InheritanceCDPBrowser
    Set ex_CDPCore = Inheritance.InheritanceCDPCore

    'SessionIdを取得
    CurrentSessionID = ex_CDPContext.CurrentSessionID
End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True


    promptStreaming.strBuffer = String(2 ^ 16, vbNullChar)

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
seeRawSendMsgDbg = ShSetting01_StartBrowser.UseRangeID(8, FromProcedureName)    'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
End Sub
