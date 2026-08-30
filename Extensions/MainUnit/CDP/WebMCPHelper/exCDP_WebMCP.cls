VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_WebMCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   Chrome 146のリリースに合わせて、Webアプリの機能をAIエージェントに公開できる新標準 WebMCP
'                   を、CDPから簡単に制御できるラッパーClassとなります
'
'   `Experimental`につきある日突然、動作不良になることもあります。ご利用には注意が必要です
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                 ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'各種クラスオブジェクト
Private m_CDPContext As CDPContext      'コンテキスト(タブ)制御クラスオブジェクト
Private WithEvents m_CDPCore As CDPCore 'イベント受け取り用
Attribute m_CDPCore.VB_VarHelpID = -1

'ステート管理
Private currentLog          As StateLog      'ログ設定
Private EnableBrowserEvents As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private CurrentSessionID    As String       '`CDPContext.cls`継承時のSesssionID
Private toolsMap            As Dictionary   'コマンド名と`Object-WebMCP.Tool`との紐づけ。`toolName`が`Key`です
Private toolRespondedMap    As Dictionary   'WebMCP実行結果の格納。`invocationId`が`Key`です

'ログ用
Private Const ThisClassName As String = "exCDP_WebMCP"



'***************************************************************************************************
'                                   ■■■ イベント処理 ■■■
'***************************************************************************************************
'* 機能　　：WebMCP.Tool名とのマッピング処理を行います
'***************************************************************************************************
Private Sub m_CDPCore_CDPContextEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".m_CDPCore_CDPContextEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. このコンテキスト(タブ)クラスで設定中の`SessionId`以外も即抜け
    If CurrentSessionID <> sessionID Then Exit Sub

    '3. WebMCPに関する各イベント処理を行う
    Dim tmp As BiDiCDPJson
    Select Case methodName
        '`WebMCP`が実行されたログを記録します
        Case "WebMCP.toolInvoked"
            '3-1. Json文字化が入るため、非表示の場合は即抜けして、パフォーマンスを維持
            If Not currentLog.seeRawSendMsgDbg Then Exit Sub

            '3-2. 実行ログを表示させる準備
            Set tmp = BiDiCDPJson.Parse(RawJson).NodeKey("params")

            '3-3. 表示
            printMsg Debug_, "WebMCP:`" & tmp.StringKey("toolName") & "` has been executed. Input: " & tmp.StringKey("input"), FromProcedureName

        '`WebMCP`の実行結果を確保します
        Case "WebMCP.toolResponded"
            '3-1. WebMCPの実行結果を格納する準備
            Set tmp = BiDiCDPJson.Parse(RawJson).NodeKey("params")

            '3-2. `invocationId`を`Key`に格納
            toolRespondedMap.Add tmp.StringKey("invocationId"), RawJson
            printMsg Debug_, "Command results have arrived from WebMCP:`" & tmp.StringKey("invocationId") & "`", FromProcedureName

        '`WebMCP`が使えるツール名を記録します
        Case "WebMCP.toolsAdded"
            '3-1. Dictionaryにコマンド名と詳細をマッピングしていく準備
            Set tmp = BiDiCDPJson.Parse(RawJson).NodeKey("params").NodeKey("tools")

            '3-2. ループで必要な情報のみ取り出していく
            Dim NodeToken As Long, ToolParams As New Dictionary
            NodeToken = tmp.FirstChildToken
            Do While NodeToken > 0
                '3-2-1. ツール説明
                ToolParams.Add "description", tmp.TokenString(NodeToken, "description")

                '3-2-2. frameId       ※実行場所
                ToolParams.Add "frameId", tmp.TokenString(NodeToken, "frameId")

                '3-2-3. inputSchema   ※受付できる引数情報
                'ない場合は、`Nothing`になります
                ToolParams.Add "inputSchema", tmp.TokenNode(NodeToken, "inputSchema")

                '3-2-4. name をKeyに適用
                toolsMap.Add tmp.TokenString(NodeToken, "name"), ToolParams
                printMsg Debug_, "WebMCP: Added `" & tmp.TokenString(NodeToken, "name") & "`", FromProcedureName

                '3-2-5. リセット
                Set ToolParams = New Dictionary

                '3-2-6. 次の`Object-WebMCP.Tool`へ
                NodeToken = tmp.NextToken(NodeToken)
            Loop

        '`WebMCP`が消えたツール名を監視し、マッピングから削除します
        Case "WebMCP.toolsRemoved"
            '3-1. Dictionaryにコマンド名と詳細をマッピングしていく準備
            Set tmp = BiDiCDPJson.Parse(RawJson).NodeKey("params").NodeKey("tools")

            '3-2. ループでマッピングを削除していく
            NodeToken = tmp.FirstChildToken
            Do While NodeToken > 0
                '3-2-1.　削除
                toolsMap.Remove tmp.TokenString(NodeToken, "name")
                printMsg Debug_, "WebMCP: Removed `" & tmp.TokenString(NodeToken, "name") & "`", FromProcedureName

                '3-2-2. 次の`Target.TargetInfo`へ
                NodeToken = tmp.NextToken(NodeToken)
            Loop

        'ページ遷移時の過去のToolリストの初期化
        Case "Page.frameNavigated"
            '各種リストをクリーン
            toolsMap.RemoveAll
            toolRespondedMap.RemoveAll
            printMsg info_, "A page transition has occurred, so the mapping information will be reset.", FromProcedureName

    End Select
End Sub



'***************************************************************************************************
'                                   ■■■ 継承処理 ■■■
'***************************************************************************************************
'* 機能　　：`CDPContext.cls`の情報を継承し、`WebMCP`関連の機能を有効化します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ThisCDPContext  継承したい`CDPContext`そのもの
'***************************************************************************************************
Public Sub Init(ThisCDPContext As CDPContext)
    'コンテキスト(タブ)情報を継承
    Set m_CDPContext = ThisCDPContext

    'イベント受け取り用オブジェクトも継承
    Set m_CDPCore = ThisCDPContext.ThisCDPCore
    CDPHelpers.InheritLogSettings currentLog, ThisCDPContext.ThisCDPCore.LogSettings

    'セッションIDを取得
    CurrentSessionID = m_CDPContext.CurrentSessionID

    '「WebMCPの監視をONにして！」と命令する
    ExecuteCDP "WebMCP.enable"

    'ページ遷移時のToolListのリセットの監視に使います
    m_CDPContext.pageEnable
End Sub



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
'* 機能　　：現在使えるWebMCPコマンド名と詳細をログに表示させます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：出されたログをAIに渡せば、良い感じにコマンドを組み立ててくれます
'***************************************************************************************************
Public Sub printWebMCPToolsList()
    Const FromProcedureName As String = ThisClassName & ".printWebMCPToolsList"


    Dim KeyList, i As Long
    KeyList = toolsMap.keys

    For i = 0 To toolsMap.Count - 1
        '1. `nothing`対策用に別変数に、確保
        Dim tmp_inputSchema As BiDiCDPJson, inputSchemaStr As String
        Set tmp_inputSchema = toolsMap(KeyList(i))("inputSchema")
        If tmp_inputSchema Is Nothing Then inputSchemaStr = "※undefined" Else inputSchemaStr = tmp_inputSchema.Stringify

        '2. 見やすいように`name`を固定長にしておく
        Dim ToolName As String * 32: ToolName = KeyList(i)

        '3. ログ化
        printMsg info_, "#" & i + 1 & vbTab & _
                            "name: " & ToolName & " | " & _
                            "description: " & toolsMap(KeyList(i))("description") & " | " & _
                            "inputSchema: " & inputSchemaStr & " | " & _
                            "frameId: " & toolsMap(KeyList(i))("frameId"), FromProcedureName
    Next

    'ない場合は別途伝える
    If toolsMap.Count = 0 Then printMsg WARN_, "WebMCPTool is not available at this time.", FromProcedureName
End Sub

Public Function WebMCPToolsCount() As Long
    '返却
    WebMCPToolsCount = toolsMap.Count
End Function



'***************************************************************************************************
'                               ■■■ ポーリング関連 ■■■
'***************************************************************************************************
'* 機能　　：指定のツール名があるかを確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：存在有無
'* 引数　　：toolName   実行しようとするツール名
'***************************************************************************************************
Public Function isExistTool(ToolName As String) As Boolean
    isExistTool = toolsMap.Exists(ToolName)
End Function

'***************************************************************************************************
'* 機能　　：指定のツール名が来るまで、イベントを監視します
'---------------------------------------------------------------------------------------------------
'* 返り値　：自身のClassオブジェクト
'* 引数　　：toolName   実行しようとするツール名
'***************************************************************************************************
Public Function onExistTool(ToolName As String, Optional timeOutInSeconds As Double = 30, Optional raiseTimeoutError As Boolean = True) As exCDP_WebMCP
    Const FromProcedureName As String = ThisClassName & ".onExistTool"


    'カウントスタート
    Dim timeStart As Double: timeStart = TimerCounter
    Do Until isExistTool(ToolName)
        '1. タイムアウトチェック
        If TimerCounter - timeStart > timeOutInSeconds * 1000 Then
            printMsg ERROR_, "onExistTool timeout on WebMCPTool = " & ToolName, FromProcedureName, doRaiseError:=raiseTimeoutError
            Exit Function
        End If

        '2. イベント取り出し
        TakeEvents
    Loop

    '自身を返却
    Set onExistTool = Me
    printMsg info_, "onExists satisfied: [" & ToolName & "] is now available for interaction", FromProcedureName
End Function



'***************************************************************************************************
'                               ■■■ 実行のやり取り ■■■
'***************************************************************************************************
'* 機能　　：WebMCPを実行させます
'---------------------------------------------------------------------------------------------------
'* 返り値　：実行結果を受け取る際の整理券ID(invocationId)
'* 引数　　：toolName   実行したいツール名
'            input_     ツール名に対する引数(Dictionary形式)
'
'* 詳細説明：・引数は、`printWebMCPList`を参考に入力してください
'            ・`frameId`は、`toolName`を参考に自動で入力されます
'            ・CDP実行自体は、同期モードで行います
'            ・結果自体はこのメソッドでは受け取れません。次述の`ExecuteWebMCP`をご利用ください
'***************************************************************************************************
Public Function ExecuteWebMCPAsync(ToolName As String, Optional input_ As Dictionary) As String
    Const FromProcedureName As String = ThisClassName & ".ExecuteWebMCPAsync"


    '1. parameters を作成
    Dim params As New Dictionary
    params.Add "frameId", toolsMap(ToolName)("frameId")
    params.Add "toolName", ToolName
    If input_ Is Nothing Then params.Add "input", New Dictionary Else params.Add "input", input_

    '2. 実行して、整理券IDを返却
    ExecuteWebMCPAsync = ExecuteCDP("WebMCP.invokeTool", params).StringKey("invocationId")
    printMsg Debug_, "Async ExecuteWebMCPAsync: " & ToolName & " - invocationId: " & ExecuteWebMCPAsync, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：WebMCPを実行させ、結果待機まで行い、`WebMCP.toolResponded`を返却します
'---------------------------------------------------------------------------------------------------
'* 返り値　：WebMCP.toolResponded
'* 引数　　：toolName       実行したいツール名
'            input_         ツール名に対する引数(Dictionary形式)
'            TimeOutSecond  指定秒間、イベント監視します。デフォルトは30sです
'
'* 詳細説明：・引数は、`printWebMCPList`を参考に入力してください
'            ・`frameId`は、`toolName`を参考に自動で入力されます
'---------------------------------------------------------------------------------------------------
'* 注意事項：エラー判定は別途必要です
'***************************************************************************************************
Public Function ExecuteWebMCP(ToolName As String, Optional input_ As Dictionary, Optional TimeOutSecond As Double = 30) As BiDiCDPJson
    Const FromProcedureName As String = ThisClassName & ".ExecuteWebMCP"


    '1. 実行して、整理券IDを取得
    Dim ResultInvocationId As String
    ResultInvocationId = ExecuteWebMCPAsync(ToolName, input_)

    '2. 結果 or タイムアウトが来るまで待機ループ
    Dim timerStart As Double: timerStart = TimerCounter
    Do
        '2-1. 結果の取り出しを試みる
        Dim tmp As String: tmp = TakeResultWebMCP(ResultInvocationId)

        '2-2. 取り出せたら、Node-Json形式にして返却後、抜ける
        If StrPtr(tmp) Then
            Set ExecuteWebMCP = BiDiCDPJson.Parse(tmp).NodeKey("params")
            If currentLog.seeRawSendMsgDbg Then printMsg Debug_, "ExecuteWebMCP `" & ToolName & "` with result: " & ExecuteWebMCP.Stringify, FromProcedureName
            Exit Function
        End If

        '2-3. イベント取り出し
        TakeEvents
    Loop Until TimerCounter - timerStart > TimeOutSecond * 1000

    '3. 見つからない旨を伝え、実行をキャンセルする
    printMsg WARN_, "This WebMCP(" & ToolName & ")-invocationId: " & ResultInvocationId & " could not be verified within the given time.", FromProcedureName
    CancelWebMCP ResultInvocationId
    printMsg info_, "This WebMCP(" & ToolName & ")-invocationId: " & ResultInvocationId & " has been cancelled.", FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：`ExecuteWebMCP/ExecuteWebMCPAsync`で実行した結果を回収します
'---------------------------------------------------------------------------------------------------
'* 引数　　：invocationId     `ExecuteWebMCP`の返り値を指定
'---------------------------------------------------------------------------------------------------
'* 注意事項：・必要に応じて、`TakeEvents`で結果を待つ仕組みが必要です
'            ・見つかったIDは、取り出した後、Dictionaryから自動で削除されます
'            ・結果がない場合は、`vbnullchar`として返ります。`StrPtr`で判定可能です
'            ・エラー判定、JsonDictionay化は別途行って下さい
'***************************************************************************************************
Property Get TakeResultWebMCP(invocationId As String) As String
    Const FromProcedureName As String = ThisClassName & ".TakeResultWebMCP"


    '1. 存在するか？
    If toolRespondedMap.Exists(invocationId) Then
        '2. Log化
        Dim tmp As String: tmp = toolRespondedMap(invocationId)
        printMsg info_, "Browser were able to confirm the existence of WebMCP invocationId:" & invocationId & "　→　" & tmp, FromProcedureName

        '3. 返却
        TakeResultWebMCP = tmp

        '4. 取り出したので消しておく
        toolRespondedMap.Remove invocationId
        printMsg Trace_, "WebMCP-invocationId: " & invocationId & " has been removed from `toolRespondedMap`.", FromProcedureName
    Else
        '2. 存在しない旨を伝える
        printMsg Trace_, "This WebMCP-invocationId(" & invocationId & ") passed is not running asynchronously or has been deleted.", FromProcedureName
    End If
End Property

'***************************************************************************************************
'* 機能　　：`ExecuteWebMCPAsync`で実行したコマンドをキャンセルします
'---------------------------------------------------------------------------------------------------
'* 引数　　：invocationId     `ExecuteWebMCP`の返り値を指定
'***************************************************************************************************
Public Sub CancelWebMCP(invocationId As String)
    '1. parameters を作成
    Dim params As New Dictionary
    params.Add "invocationId", invocationId

    '2. 実行して、WebMCPをキャンセル
    ExecuteCDP "WebMCP.cancelInvocation", params
End Sub



'***************************************************************************************************
'                               ■■■ 各種直通メソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension As String = ".cdp.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub

Private Function ExecuteCDP(methodName As String, Optional params As Dictionary, Optional StopCDPError As Boolean = True) As BiDiCDPJson
    Set ExecuteCDP = m_CDPContext.ExecuteCDP(methodName, params, StopCDPError)
End Function

Private Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    m_CDPCore.TakeEvents StopPipeError, destruction
End Sub

Private Function TimerCounter() As Double
    TimerCounter = CDPHelpers.TimerCounter
End Function



'***************************************************************************************************
'                         ■■■ クラスオブジェクトの初期化と破棄時の処理 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    '各種リストを初期化
    Set toolsMap = New Dictionary
    Set toolRespondedMap = New Dictionary

    currentLog.LogID = "WMID" & Format(Rnd * 1000, "000")
End Sub

Private Sub Class_Terminate()
    '監視をOFF
    ExecuteCDP "WebMCP.disable", StopCDPError:=False
    m_CDPContext.pageEnable False
End Sub
