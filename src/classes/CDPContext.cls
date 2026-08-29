VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
' Automating Chromium-Based Browsers with Chrome Dev Protocol API and VBA
'------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
'       Daniel Polak (SYS, Supporting Your Systems B.V.)
' Last Update :
'       22/01/26 Long Vh:      - some tweaks to the .start method to improve the usage of user-data-dir switch
'                              - enhanced the framework with a new feature called multi-instances operation (see demo example demoMultiProfileOperation)
'                              - removed auto reattach feature from the framework and added .reattach method to suit
'                                the new multi-instances operation feature
'       31/12/25 Daniel Polak: - added BringWindowToTop and bringtoforeground
'       24/10/25 Daniel Polak: - sendMessage: added soft re-attach attempt when pipe error
'                              - snapPage: added CDP native screenshot capability which avoids use of external JavaScript library
'                              - getTab: improved tab selection
'                              - quit: avoid immediate hard kills, reduces race conditions, and makes shutdown behavior predictable and robust in non-responsive cases
'                              - searchNull: new faster version
'                              - use constants (ERR_TIMEOUT, ERR_PIPE, ERR_PROTOCOL, TIMEOUT_DEFAULT, TIMEOUT_NAVIGATE) instead of numbers for clarity
'       23/10/25 Daniel Polak: - replaced unnecessary use of FileSystemObject
'                              - added Option Explicit and added all required variable declarations
'                              - quit: delete the one and only cdpid.txt
'                              - Class_Initialize: prevent compilation errors
'                              - cleanUpSessions: solved bug
'                              - added DeleteCDPID so we use a function instead of repeating same code
'                              - 64-bit safety: window handles as LongPtr
'                              - start: added missing space before --homepage
'                              - jsAddLib: fixed bug in check whether JS library is loaded
'                              - getTab: fixed typo timeStart instead of timerStart and fixed bug 60000s instead of 60s
'                              - cleanUpSessions: only terminate automation-processes
'                              - snapPage: fixed bug where the all files in the Downloads folder were deleted instead of just the temporary screenshot
'                              - getBrowserPath: added fallback path when registry lookup fails
'       18/08/23 Long Vh:      - set core to Private i/o Public as it is meant to be Private.
'       06/06/23 Long Vh:      - applied recommendations by MitsOmitsaras for Access & Word compat.
'       26/05/23 Long Vh:      - made some tweaks for compatibility with Access & Word out-of-the-box
'       26/04/23 Long Vh:      - added .setMain for better tab management.
'       22/03/23 Long Vh:      - added many .getElementsBy methods.
'                              - added advanced methods to extend features of CDP apps.
'                              - changed deserialize/serialize methods to read from a temp file.
'                              - added .snapPage method for capturing an entire web screen.
'                              - added interactions with the new element class object.
'                              - many other changes to improve debug handlings.
' References  :
'       Microsoft Scripting Runtime
' Sources     :
'       Main Git: https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA
'       CDP Tutorials: https://github.com/cyrus-and/chrome-remote-interface/issues/282
'       CodeProject: https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA
'       Javascript Tutorials: https://www.w3schools.com/js/
'       Chrome CDP Documentation: https://chromedevtools.github.io/devtools-protocol/
'================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
''----- CDPだけでは実現できない、Windows自体の機能によるウィンドウ操作関連 -----
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExW" ( _
    ByVal hwndParent As LongPtr, _
    ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As LongPtr, _
    ByVal lpszWindow As LongPtr) As LongPtr

Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextW" ( _
    ByVal hWnd As LongPtr, _
    ByVal lpString As LongPtr, _
    ByVal nMaxCount As Long) As Long

Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nCmdShow As Long) As Boolean

Private Declare PtrSafe Function BringWindowToTop Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    lpdwProcessId As Long) As Long



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const TIMEOUT_DEFAULT   As Double = 30              '初期タイムアウト秒数
Private Const ThisClassName     As String = "CDPContext"    'ログのトレース用



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'ブラウザ制御クラスオブジェクト
Private WithEvents PipeCore As CDPCore      'イベント受け取りオブジェクト
Attribute PipeCore.VB_VarHelpID = -1
Private BrowserCore         As CDPBrowser   'Must be initialized for CDP engine

'各種便利クラスオブジェクト
Private CharConv    As New CharacterCodeConversion

'操作中のタブセッションステータス
Private Type sessionStatus
    detachedFromTarget  As Boolean  'Issued when detached from target for any reason (including detachFromTarget command). Can be issued multiple times per target if multiple sessions have been attached to it.
    targetCrashed       As Boolean  'Issued when a target has crashed.
    targetDestroyed     As Boolean  'Issued when a target is destroyed. ※`Target.setDiscoverTargets`の購読必須
End Type
Private ControlStatus As sessionStatus

'操作中のタブセッションエラー情報
Private Type SessionErrorDetails
    Status      As String
    ErrorCode   As Long
End Type
Private ControlError As SessionErrorDetails

'操作中のタブ情報一式
Private Type StateTab
    'コアとなる情報
    sessionID           As String   'The ID to the current session tab.
    targetID            As String   'The targetID of the current tab for ExecuteCDP
    browserContextId    As String   'どの`Profile`に属してるか

    '補助情報
    hWnd        As LongPtr  'Browser object window handle
    WID         As String   'Useful for windows repositioning
    Title       As String   'The focused tab name / window name of the active CDPBrowser
    Url         As String   'The focused tab url of the active CDPBrowser
End Type
Private brTab As StateTab

'操作中のTOP層タブとしてのJavaScript環境情報
Private Type StateTabForJavaScript
    ExecutionContextId  As Long     '実行コンテキスト情報。`Runtime.callFunctionOn - executionContextId`に使います
    origin              As String   '実行範囲
    Name                As String   '意味のある名称
    uniqueId            As String   'より厳格な実行コンテキスト情報。破棄判定に使います
    IsValid             As Boolean  'その`ContextId`は有効か？
End Type
Private brTabRuntime As StateTabForJavaScript

'CDPコマンドの結果を蓄積する変数
Private ResultCommandCDP_Context As Dictionary  '「ID - それに対するCDPコマンドの結果」というDictionary

'ブラウザからの非同期イベント状況
Private Type StateBrowserEvents
    Queue   As Dictionary  'ブラウザからのイベントデータを蓄積するための変数です。
    Filter  As Dictionary  'ブラウザからのイベントデータのうち、登録したイベント名のみ、`Queue`に蓄積させるフィルター情報です
End Type
Private publicBrowserEvents As StateBrowserEvents   'bas側で使用する用

'各種エラー情報
Private lastErrorDetailsCDP     As Dictionary  'CDP-Jsonコマンド実行時の最後のエラー情報
Private lastExceptionDetailsJS  As BiDiCDPJson 'JavaScript実行時の最後のエラー情報(Node情報方式)

'その他
Public KeepSession      As Boolean  'JavaScriptの実行内容(ObjectID)を別のプロシージャに跨ぐ等で、SessionIDを破棄しないようにするオプション
Private timerOut        As Double   'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
Private currentLog      As StateLog 'ログ設定



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザ起動と同時に、非同期イベント：`Target.targetCreated`を利用した初回タブ接続処理を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：userProfile    起動引数：`--user-data-dir`　につけるユーザーディレクトリ名を指定します
'            Name           ブラウザ種別（`BrowserList.RunChrome` / `BrowserList.RunEdge`）
'            appUrl         起動引数：`--app`　につけるURLを指定します
'            addArgs        追加の起動引数
'---------------------------------------------------------------------------------------------------
'* 機能説明：ブラウザ起動と同時に、タブ接続まで自動で行います
'***************************************************************************************************
Public Sub StartAndConnectTab(Optional Name As BrowserList = BrowserList.RunChrome, _
    Optional appUrl As String, _
    Optional userProfile As String, Optional addArgs As String)
    Const FromProcedureName As String = ThisClassName & ".StartAndConnectTab"

    '1. ブラウザ起動
    Set BrowserCore = New CDPBrowser
    BrowserCore.start Name, appUrl, userProfile, addArgs

    '2. WithEventsClassオブジェクトを受け取る
    Set Me.InheritanceCDPCore = BrowserCore.InheritanceCDPCore

    '3. Complete initialization ------------------------------------------------------
    attachToSession     'Find and attach the newly open Session ID
    getSessionInfo      'Store initial session info for printParams debugging
    BrowserCore.serializeForMainTab brTab.sessionID, brTab.targetID
    '------------------------------------------------------------------------------

   'Make sure the page is fully loaded for inAppMode ------------------------------------------------------------------------------------
   '07/01/25: Long Vh: simplified the check as there are web apps that does not end up loading the target url but points to another url.
   'As a result, the below check simply checks if there is "movement" in the name of the page instead
    Dim timerStart As Double, strLoc As String
    If LenB(appUrl) Then
        timerStart = CDPHelpers.TimerCounter
        Do
            If LenB(Me.Title) Then Exit Do
            If CDPHelpers.TimerCounter - timerStart > timerOut Then Err.Raise CDPCustomErrorCodes.TIMEOUT, , "Timeout waiting for the browser to load the homepage."
            DoEvents
        Loop
       wait
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------

   'Finally print the session details for debugging ------------------------------
    printParams
   '------------------------------------------------------------------------------

End Sub

Private Sub attachToSession()
'-----------------------------------------------------------------------
' This function must be called after start and before all other methods.
' This function attaches to a session of the browser.
' Updated: 03/05/23 - Added printMsg msgs & removed URL checking req.
'-----------------------------------------------------------------------
' Note: if needed to read the returned message result from sendMessage,
' type ?jsConverter.ConvertToJson(objAllMessages) in the ImmWnd
'-----------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".attachToSession"


    Dim FirstTabs As Collection
    Dim result

    Dim strKey As Variant

    On Error GoTo ifError


    '`Target.targetCreated`を収集できるように初期化
    Set FirstTabs = New Collection
1   Set BrowserCore.RunCollect_TargetCreated = FirstTabs
    BrowserCore.EnableDiscoverTargets = True
    printMsg info_, "Now looking for the browser window", FromProcedureName

   'Read message to confirm the requested url is found
   'Checking results__.Item("url") has been skipped for simplicity since v2.3
   'A new session always has the property attached = false and method = Target.targetCreated
   'This way of verification vastly improves the stability over checking by URL as sometimes
   'there are rogues targets with similar URL but they are not attachable (ie. no Session ID)
    Dim boolFound As Boolean, doc As BiDiCDPJson
    For Each result In FirstTabs
        ' 1. 文字列をパース
        Set doc = BiDiCDPJson.Parse(CStr(result))

        ' 2. 必要なノードを安全に辿る
        If doc.ExistsKey("params") Then
            ' NodeKey で階層を降り、targetInfo の存在を確認
            With doc.NodeKey("params")
                If .ExistsKey("targetInfo") Then
                    ' targetInfo 内部を型指定アクセサで判定
                    With .NodeKey("targetInfo")
                        ' BoolKey と StringKey を使うことで高速に値を取得
                        If .BoolKey("attached") = False And .StringKey("type") = "page" Then
                            boolFound = True

                            '制御パラメーターを付与
                            brTab.targetID = .StringKey("targetId")
                            brTab.browserContextId = .StringKey("browserContextId")
                            Exit For
                        End If
                    End With
                End If
            End With
        End If
    Next

2   If Not boolFound Then printMsg ERROR_, "Unable to find the browser session", FromProcedureName, doRaiseError:=True _
    Else printMsg info_, "Now attaching CDP pipes to the browser session", FromProcedureName

   'Establish connection with the CDP pipes
    brTab.sessionID = BrowserCore.attachToTab(brTab.targetID)

    '自動蓄積機能を無効化
3   Set BrowserCore.RunCollect_TargetCreated = Nothing

    printMsg info_, "Successfully attached to Session ID " & brTab.sessionID, FromProcedureName
    Exit Sub

ifError:

   'Record errmsg and print helpful infos
    Dim errorLine As Long, errorDesc As String
    errorLine = Erl
    errorDesc = Err.Description
    printMsg ERROR_, "attachToSession error at line " & errorLine & ": """ & errorDesc & """", FromProcedureName
    printMsg Debug_, "result = " & result, FromProcedureName
    printMsg Debug_, ".start call string: " & BrowserCore.CurrentBrowserPath, FromProcedureName

    printParams
    Err.Raise CDPCustomErrorCodes.Protocol, , "attachToSession error at line " & errorLine & ": """ & errorDesc & """"

End Sub



'***************************************************************************************************
'                               ■■■ 再接続処理系 ■■■
'***************************************************************************************************
'* 機能　　：Excelテーブルにある既存のパイプハンドル、ブラウザセッション情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：既存タブへの接続成功可否論理値
'* 引数　　：userProfile    再アタッチしたいユーザー名(user-data-dirに基づく)
'***************************************************************************************************
Public Function reattachPipe(userProfile As String, Optional reuseSession As Boolean) As Boolean
    '1. 新オブジェクトを用意
    Set BrowserCore = New CDPBrowser

    '2. 識別名称に紐づくパイプハンドル情報を再セット
    BrowserCore.reattachPipe userProfile

    '3. 既存のタブ(TargetID)への再接続を試みる
    reattachPipe = reattach(reuseSession)
End Function

'***************************************************************************************************
'* 機能　　：WebSocket、ブラウザセッション情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：既存タブへの接続成功可否論理値
'* 引数　　：userProfile    再アタッチしたいユーザー名(user-data-dirに基づく)
'***************************************************************************************************
Public Function reattachWebSocket(userProfile As String, WebSocketMode As CDPCoreViaWebSocket, Optional reuseSession As Boolean) As Boolean
    '1. 新オブジェクトを用意
    Set BrowserCore = New CDPBrowser

    '2. 識別名称に紐づくWebSocket情報を再セット
    BrowserCore.reattachWebSocket userProfile, WebSocketMode

    '3. 既存のタブ(TargetID)への再接続を試みる
    reattachWebSocket = reattach(reuseSession)
End Function

'***************************************************************************************************
'* 機能　　：WebView2、ブラウザセッション情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：既存タブへの接続成功可否論理値
'* 引数　　：userProfile    再アタッチしたいユーザー名(user-data-dirに基づく)
'***************************************************************************************************
Public Function reattachWebView2(userProfile As String, WebView2Mode As CDPCoreViaWebView2, Optional reuseSession As Boolean) As Boolean
    '1. 新オブジェクトを用意
    Set BrowserCore = New CDPBrowser

    '2. WebView2情報を再セット
    BrowserCore.reattachWebView2 userProfile, WebView2Mode

    '3. 既存のタブ(TargetID)への再接続を試みる
    reattachWebView2 = reattach(reuseSession)
End Function

'***************************************************************************************************
'* 機能　　：Excelテーブルにある既存のタブ(TargetID)情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：TargetIDの生存論理値
'* 引数　　：reuseSession   Excelに記録してる`SessionID`を使いまわすか？
'---------------------------------------------------------------------------------------------------
'* 機能説明：Reattach to the target CPID file. This method enables the management of multiple
'            CDP instances at once under different Chrome/Edge user profiles. The benefit of
'            this is that if one instance fails/crashes, it does not affect the other ones.
'            and thus enables the possibility of running multiple CDP automation operations
'            on entirely different tasks simultaneously.
'
'* 注意事項：・この処理はあくまで、Excelテーブルに記録されてる`targetID`への再接続まで担う処理です。
'            　この`targetID`自体が既に閉じてしまってる場合、結果は`False`で返りますが、VBAエラーがない場合、パイプハンドルは生きてるので、下記のメソッドで手動で復旧可能です。
'            　その際は、`setMain:=True`で、セッションを確立させてください。
'            　・CDPBrowser.newTab
'            　・CDPBrowser.getTab
'
'            ・非同期イベント情報を購読中の場合(例：Network.webSocketFrameReceived)は、`reuseSession:=True`推奨です
'***************************************************************************************************
Private Function reattach(Optional reuseSession As Boolean) As Boolean
    Const FromProcedureName As String = ThisClassName & ".reattach"


    '1. タブ情報を再セット
    Set Me.InheritanceCDPCore = BrowserCore.InheritanceCDPCore
    deserializeForTab reuseSession

    '2. `sessionId`が空欄の場合は、即抜け
    If LenB(brTab.sessionID) = 0 Then reattach = False: printMsg WARN_, "Failed to reattach to previous CPID.", FromProcedureName: Exit Function

    '3. sessionId生存check
    If refreshTargetInfo Is Nothing Then
        reattach = False


        printMsg WARN_, "Failed to reattach to previous CPID.", FromProcedureName
    Else
        reattach = True


        '`Context`(タブ等)の変化を捉えれるようにする
        BrowserCore.EnableDiscoverTargets = True

        'ステータス更新
        getSessionInfo
        printParams
    End If

End Function

'***************************************************************************************************
'* 機能　　：Excelのテーブルから、ブラウザのタブ情報を取得します
'---------------------------------------------------------------------------------------------------
'* 引数　　：reuseSession   `True` ：Excelに記録中のsessionIdを流用します
'                           `False`：`targetID`を基に、sessionIdを更新し、Excelに記録中の古いsessionIdの破棄も行い、上書きします
'***************************************************************************************************
Private Sub deserializeForTab(reuseSession As Boolean)
    Const FromProcedureName As String = ThisClassName & ".deserializeForTab"


    '1. Excelテーブルから、`UserName`に紐づくタブセッション情報を読み込む
    Dim objSerialize As Dictionary
    Set objSerialize = ShSetting01_StartBrowser.TableBrowserContext(BrowserCore.CurrentBrowserUser, FromProcedureName)

    '2. そもそもない場合は即抜け
    If objSerialize Is Nothing Then Exit Sub

    '3. タブのTargetID情報を適用しない場合は、即抜け
    brTab.targetID = objSerialize.Item("targetID")
    If LenB(brTab.targetID) = 0 Then printMsg ERROR_, "targetId is not specified.", FromProcedureName: Exit Sub

    '4. SessionIDの適用方法に応じた分岐を行う
    If reuseSession And LenB(objSerialize.Item("sessionID")) Then
        '既存のSessionIDを適用
        '4-1. 古いSessionIDを適用
        brTab.sessionID = objSerialize.Item("sessionID")
        printMsg info_, "I applied the SessionID(" & brTab.sessionID & ") from the Excel table.", FromProcedureName
    Else
        '古いSessionIDを破棄してから、新しいSessionIDを付与
        '4-1. 新しい`sessionId`を発行し、失敗時はエラーログを出して抜ける
        Dim newSessionID As String: newSessionID = BrowserCore.attachToTab(brTab.targetID)
        If LenB(newSessionID) = 0 Then printMsg ERROR_, "Failed to update sessionid with Targetid: " & brTab.targetID, FromProcedureName: Exit Sub

        '4-2. 一旦、古いSessionIDを適用
        brTab.sessionID = objSerialize.Item("sessionID")

        '4-3. 内部で、その古いSessionIDを破棄させ、新しいSessionIDを付与
        Me.CurrentSessionID = newSessionID
    End If
End Sub



'***************************************************************************************************
'                              ■■■ JavaScript コア機能 ■■■
'***************************************************************************************************
'* 機能　　：記述したJavaScriptをブラウザ内で実行します。JavaScript版低レベル操作としての位置付けです。
'            Webスクレイピングに関わる引数は全て設けています。
'---------------------------------------------------------------------------------------------------
'* 返り値　：value属性の中身を評価して、Variant型で返却。実行したJavaScriptコード、引数:returnByValue や、IFEXCEPTION に依存します。
'* 引数　　：JavaScriptStr                  実行したいJavaScriptコードそのもの。`As Is`として実行されます。※ほとんどのエスケープ処理は、`WebJsonConverter`が担うため、ここでわざわざ手動エスケープの記述は必要ありません。
'
'            objectId                       実行の基準要素オブジェクトIDを指定します。これを指定すると`Runtime.callFunctionOn`として優先実行されます。
'
'            objectArguments                これを指定すると`Runtime.callFunctionOn`として優先実行されます。
'                                           予め、呼び出し側にて`Collection` OR `Array(... , ...)` OR `hoge(0) as Dictionary(Dictionary型 1次元配列)`で組み立てて渡してください。なお、後者ほどパフォーマンスが向上します
'                                           参考となる組み立て手法：https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#type-CallArgument
'
'            IFEXCEPTION                    JavaScript実行中に「例外（Exception）」が発生した場合の代替値を指定します。名前の通り、IFERROR関数チックに使えます。
'            contextId                      `iframe`内で、JavaScript実行をしたいときに指定します。`objectId`が指定されてる場合は、無視されます
'                                           `getIFrameContextID`メソッド等で得れるのがそれです
'
'            returnByValue                  True ：type:"object" 等の場合でも、`value`で受け取れるようにCDPが試みます。実行コードによっては、`BiDiCDPJson.cls`形式で返されるため、`SET`で受け取る必要があります。なお、空のオブジェクトで返る場合は、対応してないデータ型の可能性があるためその際は、`false`に切り替えてください。
'                                           False：type:"object" 等の場合は、`{"objectId": "-4985739485.1.3"}`にある`-4985739485.1.3`というブラウザ内変数名的なものとして返えります。再度`jsEval`を使う際の引数`objectId`に流用できます。「このJSコードは`SET`で...」という意識は不要です。(既定)
'
'            awaitPromise                   True ：JavaScriptそのものの実行結果が非同期処理であっても、処理が終わるまで待ってから結果を返すことができます
'                                           False：JavaScriptそのものの実行結果が非同期処理の場合、結果を待たずに続行されます(既定)
'
'            userGesture                    True ：「人間がクリックやキーボード操作をした結果」としてブラウザに偽装します。スクレピング対策が施されてる場合は、`True`にしておくといいかもしれません
'                                           False：特に施しません(既定)
'                                           ※この引数はどうやらクールタイム制のようで、例えば最初に`userGesture:=True`すれば以降のJavaScript実行に`userGesture:=True`を付与しなくてもしばらくは、「人間がクリックやキーボード操作をした結果」とみなす仕様みたいです
'
'            allowUnsafeEvalBlockedByCSP    True ：サイトの外部JavaScript実行規制を無視して、全てのJavaScript実行を許可します(既定)
'                                           False：サイトの外部JavaScript実行規制に従います
'                                           ※この引数は`Runtime.evaluate`のみ有効です
'
'            serializationOptions           意地でも、type:"object" 等となっても、`value`で受け取れるように、細かい調整をします。予め、呼び出し側にてDictionaryを組み立てて渡してください。
'                                           このパラメーターがあると、`returnByValue`の引数は無視されます。
'
'            generatePreview                `objectId`で返った中身をプレーンテキスト化して、プレビューします。予め、ログレベルを`DEBUG`以下にしてください。基本は、開発時のみ`True`にしてください。
'
'            RunAsyncCDP                    True ：結果を待たないCDPで実行します。実行時のIDが返却されます。主に、「この操作、`alert`関数が発動するな...」という場面で使用します。
'                                           False：結果を待つCDPで実行します。基本はこれでOK(既定)
'
'            StopException                  `True`で、JavaScriptコード内で例外が起こった場合、Err.Raiseで止まります。基本は、開発時のみ使用する想定です
'            StopPipeError                  `True`で、Pipeハンドルエラー時、エラーで停止します
'---------------------------------------------------------------------------------------------------
'* 機能説明：This function commands the attached browser instance to execute a Javascript expression to perform required tasks.
'            Note: the original function has a bool retry variable which was removed for simplicity.
'            Added dbgMsg to read debug line on the go when needed.
'
'* 注意事項：・利用時は、「名前付き引数`xxx:=yyy`」として運用することをお勧めします
'            ・JavaScript実行中の例外確認は、「IsError関数」で判断してください。詳細の確認は、`LastJavaScriptException`プロパティで確認できます。「IsError関数」での確認が面倒なら、引数`IFEXCEPTION`を活用してください。
'            ・`returnByValue`=true の場合は、SET or LET で受け取る意識が必要になります。結果によって、`SET`or`LET`が変わって静的にできない場合は、Array関数で一旦受け止めて、VarType(hoge(0))で`SET`or`LET`判定することも可能です
'
'            ・「このJavaScriptコードの場合は、こういう返り値が来るな」といった知識がないと、返り値に困惑します。
'               開発時は、デバッグログを活用した試行錯誤をすることをお勧めします。
'
'            ・`objectId`や`objectArguments`を指定する`Runtime.callFunctionOn`方式の場合は、`JavaScriptStr`を 「function() { this... }」 のような形式で記述する必要があります
'
'            ・実行したJavaScriptコードの結果が、プリミティブ型の場合、`returnByValue`の設定問わずそのプリミティブ型の中身が返ります。
'              常に`objectId`で返したい場合は、JavaScriptコードそのものを工夫する必要があります。
'
'            ・JavaScriptの場合、例外を利用した処理が多々あるため、運用時は`StopException`を省略しておくことを推奨します
'            ・CDP実行自体のエラーキャッチは常に`True`となります
'            ・`IFEXCEPTION`はあくまでも、「例外」時のみ機能します。JavaScript自体は成功してるけど、結果がエラーの場合はISERROR関数で対処する必要があります。
'---------------------------------------------------------------------------------------------------
'* 参考URL ：https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#method-evaluate
'***************************************************************************************************
Public Function jsEval(JavaScriptStr As String, Optional objectId As String, Optional objectArguments As Variant, Optional IFEXCEPTION As Variant, _
                       Optional contextId As Long, _
                       Optional returnByValue As Boolean, _
                       Optional awaitPromise As Boolean, _
                       Optional userGesture As Boolean, _
                       Optional allowUnsafeEvalBlockedByCSP As Boolean = True, _
                       Optional serializationOptions As Dictionary, _
                       Optional generatePreview As Boolean, _
                       Optional RunAsyncCDP As Boolean, _
                       Optional StopException As Boolean, _
                       Optional StopApiError As Boolean = True) As Variant
    Const FromProcedureName As String = ThisClassName & ".jsEval"


    Dim CDPresult   As String
    Dim CDPparams   As New Dictionary

step1:

   'Construct message for sending ----------------------------------------------------------
    CDPparams.RemoveAll

    '`callFunctionOn`か`evaluate`を決める判定に入ります
    Dim UseJsMethod As String

    ' 【パターン1】特定の要素を操作する場合
    If StrPtr(objectId) Then
        UseJsMethod = "Runtime.callFunctionOn"

        '必要なパラメーターをセット
        CDPparams.Add "functionDeclaration", JavaScriptStr
        If Not IsMissing(objectArguments) Then CDPparams.Add "arguments", objectArguments   '引数があったら適用しておく

        'objectId があるので、場所(ContextID)を指定する必要はない
        CDPparams.Add "objectId", objectId

    ' 【パターン2】要素指定はないが、安全に引数(arguments)を渡したい場合
    ElseIf Not IsMissing(objectArguments) Then
        UseJsMethod = "Runtime.callFunctionOn"

        '必要なパラメーターをセット
        CDPparams.Add "functionDeclaration", JavaScriptStr
        CDPparams.Add "arguments", objectArguments

        ' objectId がないので、実行する「場所(executionContextId)」を教える必要がある
        '任意指定あるいは、`Runtime.executionContextCreated`で来た情報を基に指定します
        Dim UseContextId As Long
        If contextId Then
            '引数で指定された`contextId`
            UseContextId = contextId
        Else
            'このClassの`TargetID`に紐づく"Runtime.executionContextCreated"の`contextId`
            '無効(古い)状態なら、有効な`contextId`が来るまで待機します
            If Not brTabRuntime.IsValid Then printMsg info_, "The `executionContextId` required to execute `Runtime.callFunctionOn` is invalid. Attempting to update it via `TakeEvents`.", FromProcedureName: AutoWaitIsValidExecutionContextId

            '適用
            UseContextId = brTabRuntime.ExecutionContextId
        End If
        CDPparams.Add "executionContextId", UseContextId

    ' 【パターン3】引数も要素指定もない、ただのコード実行
    Else
        ' 最もシンプルな evaluate を使う
        UseJsMethod = "Runtime.evaluate"

        '必要なパラメーターをセット
        CDPparams.Add "expression", JavaScriptStr
        If contextId Then CDPparams.Add "contextId", contextId
        If Not allowUnsafeEvalBlockedByCSP Then CDPparams.Add "allowUnsafeEvalBlockedByCSP", allowUnsafeEvalBlockedByCSP
    End If

    '以降は、共通設定を施す
    '※文字数節約により、IF文をかます
    If returnByValue Then CDPparams.Add "returnByValue", returnByValue
    If awaitPromise Then CDPparams.Add "awaitPromise", awaitPromise
    If userGesture Then CDPparams.Add "userGesture", userGesture
    If generatePreview Then CDPparams.Add "generatePreview", generatePreview
    If Not serializationOptions Is Nothing Then CDPparams.Add "serializationOptions", serializationOptions

    '----------------------------------------------------------------------------------------

    'Send and check CDPresult -------------------------------------------------------------------------------------------
    'まずは、CDP自体を非同期で実行し、結果待機するか分岐します
    Dim RunJsID As Long
    RunJsID = ExecuteCDPAsync(UseJsMethod, CDPparams, StopApiError)
    If RunAsyncCDP Then
        'ID返却して抜ける
        jsEval = RunJsID
        Exit Function
    Else
        '結果待機
        CDPresult = AutoWaitTakeResultCDP(RunJsID, StopApiError)

        'タイムアウトやパイプハンドルの応答なしによる空文字判定
        If StrPtr(CDPresult) = 0 Then printMsg ERROR_, "Unable to contact the Browser instance!", FromProcedureName, doRaiseError:=True

        'CDP自体の結果ログ表示
        If currentLog.seeRawSendMsgDbg Then printMsg Trace_, "jsEval - ResultCDP - JsonString = " & CDPresult, FromProcedureName

        'エラーチェック
        Dim Node_CDPresult As BiDiCDPJson: Set Node_CDPresult = BiDiCDPJson.Parse(CDPresult)
        If Not ExecuteError(Node_CDPresult) Then GoTo step2   'CDP自体のエラーがなければ、`step2`へ
    End If
    '-----------------------------------------------------------------------------------------------------------------

   'Handle errors ---------------------------------------------------------------------------------------------------
    Dim errMsg As String, retried As Boolean
    errMsg = lastErrorDetailsCDP("message")
    printMsg WARN_, "jsEval Error: " & errMsg, FromProcedureName
    If errMsg = "'" & UseJsMethod & "' wasn't found" Then printMsg WARN_, "Possibly the main session has been abruptly closed.", FromProcedureName
    If errMsg = "Session with given id not found." And Not retried Then
        printMsg info_, "Attempting to update Session ID " & brTab.sessionID, FromProcedureName

        '`SessionId`の更新を試みる
        Dim newSessionID As String: newSessionID = BrowserCore.attachToTab(brTab.targetID)

        '更新に失敗時はエラーで停止
        If LenB(newSessionID) = 0 Then printMsg ERROR_, "Failed to update SessionID." & vbCrLf & "> " & BrowserCore.LastCDPJsonError("message"), FromProcedureName, doRaiseError:=True

        '更新させて、リトライ
        Me.CurrentSessionID = newSessionID
        printMsg info_, "New Session ID obtained: " & brTab.sessionID, FromProcedureName
        retried = True: GoTo step1
    End If

    '実行したJavaScriptコードの改行を消して、長すぎる場合はカットする（整形）
    Dim strLogValue As String
    strLogValue = LogFormat(JavaScriptStr, 100)
    printMsg ERROR_, "The Browser was unable to evaluate the JS expression: " & strLogValue & vbCrLf & vbCrLf & "<ErrorDetail>" & vbCrLf & errMsg, FromProcedureName, doRaiseError:=True
   '-----------------------------------------------------------------------------------------------------------------

step2:

    'JavaScript実行結果部分の解析フェーズに移ります ----------------------------------------------------------------
    '各種`Return Object`を取り出します
    Dim Node_ReturnObject       As BiDiCDPJson: Set Node_ReturnObject = Node_CDPresult.NodeKey("result")

    Dim Node_RemoteObject       As BiDiCDPJson: Set Node_RemoteObject = Node_ReturnObject.NodeKey("result")
    Dim Node_ExceptionDetails   As BiDiCDPJson: If Node_ReturnObject.ExistsKey("exceptionDetails") Then Set Node_ExceptionDetails = Node_ReturnObject.NodeKey("exceptionDetails")  '`optional`なのでまずは、判定をかます

    '`RemoteObject`部分を文字ログ化
    If currentLog.seeRawSendMsgDbg Then printMsg Debug_, "jsEval - RemoteObject - JsonString = " & Node_RemoteObject.Stringify, FromProcedureName

    '例外（JS側のエラー）が発生していないかチェックし、あったら、ログを表示し、抜けます
    If Not Node_ExceptionDetails Is Nothing Then
        '最後の詳細例外情報を記録
        Set lastExceptionDetailsJS = Node_ExceptionDetails

        '実行したJavaScriptコードの改行を消して、長すぎる場合はカットする（整形）
        strLogValue = LogFormat(JavaScriptStr, 100)

        '例外による`Err.Raise`で止めるか？
        If StopException Then
            '`StopException`が有効の場合は、ログ表示用に、詳細の方を文字列化
            Dim exceptionDetails As String: exceptionDetails = lastExceptionDetailsJS.Stringify
            printMsg ERROR_, "An exception occurred while executing JavaScript: " & strLogValue & vbCrLf & vbCrLf & "<exceptionDetails>" & vbCrLf & exceptionDetails, FromProcedureName, doRaiseError:=True
        Else
            '`StopException`が無効の場合は、簡易部分の例外内容を出す
            printMsg WARN_, "ExceptionType: " & Node_RemoteObject.StringKey("className") & " , " & "ExceptionDescription: " & Node_RemoteObject.StringKey("description"), FromProcedureName

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
    If JSresultAnalysis(Node_RemoteObject, generatePreview, tmp) Then
        Set jsEval = tmp
    Else
        Let jsEval = tmp
    End If
   '-----------------------------------------------------------------------------------------------------------

step3:

    'Output to Immediate Window if enabled -------------------------------------------------------------------
    ' VarTypeを使って型ごとにログ用の文字列を作る
    Select Case VarType(jsEval)
        Case vbDouble
            strLogValue = "[Double] " & jsEval

        Case vbString
            strLogValue = "[String] " & jsEval

        Case vbBoolean
            strLogValue = "[Boolean] " & jsEval

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
'* 引数　　：RemoteObject   `Runtime.RemoteObject`でのJson-Dictionaryを指定します。　https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#type-RemoteObject
'            ObjectPreview  オブジェクトツリー構造の一部をテキスト化してプレビューします。デバッグ用途のみ
'            ReturnValue    結果値。参照渡しによる実質返り値です。`primitive values` OR `Object` 両方が想定されます。
'
'* 返り値　：ReturnValueがObjectであるかの論理値。`ReturnValue`がObject型であるかを呼び出し側で正しく代入するために必要です。
'---------------------------------------------------------------------------------------------------
'* 注意事項：・返す優先順位は、`deepSerializedValue` > `objectId` > `value`です
'            ・`description`等といった補助情報やエラー情報は返しません。
'***************************************************************************************************
Private Function JSresultAnalysis(RemoteObject As BiDiCDPJson, ObjectPreview As Boolean, ByRef ReturnValue As Variant) As Boolean
    Const FromProcedureName As String = ThisClassName & ".JSresultAnalysis"
    Const ErrorLandmark     As String = "error"


    '------------------------ 1. 特定項目有無チェック ------------------------
    '`deepSerializedValue`があったらそれを最優先に返します
    If RemoteObject.ExistsKey("deepSerializedValue") Then
        '見つかったので、`deepSerializedValue`を返して、ここで抜ける
        Set ReturnValue = RemoteObject("deepSerializedValue")
        printMsg info_, "JS-Return type: object-deepSerializedValue", FromProcedureName

        'Objectで返したのでその目印もつけておく
        JSresultAnalysis = True
        Exit Function

    '`objectId`があったら、こっちを返します
    '※type=objectで、`returnByValue:=False`なら`null`以外、ここを通ります
    ElseIf RemoteObject.ExistsKey("objectId") Then
        '見つかったので、`objectId`を返す
        ReturnValue = RemoteObject.StringKey("objectId")
        printMsg info_, "JS-Return type: objectId", FromProcedureName

        '返ったObjectID情報を列挙する
        Dim objectSubtype As String: objectSubtype = RemoteObject.StringKey("subtype")
        printMsg Debug_, String(100, "="), FromProcedureName
        printMsg Debug_, "objectId information", FromProcedureName
        printMsg Debug_, String(100, "-"), FromProcedureName
        printMsg Debug_, "objectId          : " & ReturnValue, FromProcedureName
        printMsg Debug_, "objectType        : " & RemoteObject.StringKey("type"), FromProcedureName
        printMsg Debug_, "objectSubtype     : " & objectSubtype, FromProcedureName
        printMsg Debug_, "objectClassName   : " & RemoteObject.StringKey("className"), FromProcedureName
        printMsg Debug_, "objectDescription : " & RemoteObject.StringKey("description"), FromProcedureName
        If ObjectPreview Then printMsg Debug_, "objectPreview     : " & RemoteObject.NodeKey("preview").Stringify, FromProcedureName
        printMsg Debug_, String(100, "="), FromProcedureName

        'エラー情報があればそのログを出しておく
        '※「try { ... } catch(e) { return e }」といった書き方をしないとこのログは表示されない想定
        If objectSubtype = ErrorLandmark Then printMsg WARN_, "This objectId(" & ReturnValue & ") contains an error. ErrorType: " & RemoteObject.StringKey("className") & " , " & "ErrorDescription: " & RemoteObject.StringKey("description"), FromProcedureName

        Exit Function
    End If


    '------------------------ 2. primitive values OR JsonNodeの取得を試みる ------------------------
    'primitive values の可能性があるので、それの取得を試みます
    '以降は、`returnByValue:=True`でないと通りません
    '例外として、`object-null`の場合は`returnByValue:=False`でもここを通って、null返却します

    ' 格納型情報を取得
    Dim ValueType       As String: ValueType = RemoteObject.StringKey("type")
    Dim ExistsSubtype   As Boolean: ExistsSubtype = RemoteObject.ExistsKey("subtype")
    Dim ValueSubtype    As String: If ExistsSubtype Then ValueSubtype = RemoteObject.StringKey("subtype")
    printMsg Debug_, "Type info. type: " & ValueType & " , subtype: " & ValueSubtype, FromProcedureName

    '型に応じた分岐処理
    Select Case ValueType
        Case "object"
            ' object型の場合はsubtypeで更に細かく分岐
            Select Case ValueSubtype
                Case "null"
                    ReturnValue = Null
                    printMsg info_, "JS-Return type: object-null", FromProcedureName

                Case Else
                    '`returnByValue:=True`の場合は`null`以外全て、Json文字列として生成されます
                    'よって、Node形式として返却します
                    Set ReturnValue = RemoteObject.NodeKey("value")
                    printMsg info_, "JS-Return type: object(returnByValue:=True)", FromProcedureName

            End Select

        Case "function"
            ' 関数が返ってきた場合は、「{"type":"function","value":{}}」なので、`SET`
            Set ReturnValue = RemoteObject.NodeKey("value")
            printMsg info_, "JS-Return type: function", FromProcedureName

        Case "undefined"
            ' JSの undefined は VBAでは Empty が一番近い
            ReturnValue = Empty
            printMsg info_, "JS-Return type: undefined", FromProcedureName

        Case "string"
            ' 文字列として認識するように返す
            ReturnValue = RemoteObject.StringKey("value")
            printMsg info_, "JS-Return type: string", FromProcedureName

        Case "number"
            'VBAで扱える型か？
            If RemoteObject.ExistsKey("unserializableValue") Then
                'VBAで扱えない特殊な"number"の場合は、警告を出し、「#NUM!」で返します。
                printMsg WARN_, "This `number` can't be processed in VBA. It returns `#NUM!` instead. Actual JS result: " & RemoteObject("unserializableValue") & " , description: " & RemoteObject("description"), FromProcedureName
                ReturnValue = CVErr(xlErrNum)
            Else
                ' Double型として認識するように返す
                ' ※JavaScriptは、整数専用等の型がないため
                ReturnValue = CDbl(RemoteObject.NumberKey("value"))
            End If

            printMsg info_, "JS-Return type: number", FromProcedureName

        Case "boolean"
            ' boolean型として認識するように返す
            ReturnValue = RemoteObject.BoolKey("value")
            printMsg info_, "JS-Return type: boolean", FromProcedureName

        Case "symbol"
            '※`symbol`は、`returnByValue:=True`にすると、「Object couldn't be returned by value.」というCDPエラーで返るので通らないと思うが念のため。
            ReturnValue = RemoteObject.Item("value")
            printMsg info_, "JS-Return type: symbol", FromProcedureName

        Case "bigint"
            ' BigInt は JSON に変換できない巨大整数のため、unserializableValue の中にある値をそのまんま文字列として、返却します
            ReturnValue = RemoteObject.StringKey("unserializableValue")
            printMsg WARN_, "Unserializable BigInt detected. Returning as String: " & ReturnValue & " , description: " & RemoteObject("description"), FromProcedureName
            printMsg info_, "JS-Return type: bigint", FromProcedureName

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
'* 機能　　：JavaScript実行場所を特定する`executionContextId`に有効なサインが来るまで待機します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・タイムアウト時はエラーとなります
'            ・待機時間は、`timerOut`に依存します
'***************************************************************************************************
Private Sub AutoWaitIsValidExecutionContextId()
    Const FromProcedureName As String = ThisClassName & ".AutoWaitIsValidExecutionContextId"


    '1. 現在の経過時間を取得
    Dim StartTimer As Double: StartTimer = CDPHelpers.TimerCounter
    Do
        '2. タイムアウトチェック
        If CDPHelpers.TimerCounter - StartTimer > timerOut Then printMsg ERROR_, "A valid `ExecutionContextId` was not received within the time limit.", FromProcedureName, doRaiseError:=True

        '3. イベント取り出し
        printMsg Trace_, "Waiting for a valid `ExecutionContextId`...", FromProcedureName
        TakeEvents  '内部で`DoEvents`が走ります
    Loop Until brTabRuntime.IsValid
End Sub

Public Sub jsAddLib(strURL As String)
'----------------------------------------------------------------------------------
' Load the external JS Library to the document for use of specialized JS methods.
' This is very useful when built-in JS commands are not sufficient to execute a
' task easily, such as when taking a screenshot of an entire webpage while this can
' be done relatively easy using external JS libraries. This function shall enables
' massive possibility of extended features.
' https://stackoverflow.com/a/57572107
' Updated: 23/03/2023
'          23/10/2025 Daniel Polak - fixed bug in check whether JS library is loaded
'----------------------------------------------------------------------------------
' Note: the script library should only be added after the page is fully loaded. If
' it is loaded when the HTML document is refereshing, it risks adding the script to
' the old HTML document which will be replaced by the new one causing the script to
' be inexistent and resulting in "ReferenceError" due to its functions not defined.
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".jsAddLib"


   'Check if already loaded, if so then no need to add again
    If getElementByQuery("script[src='" & strURL & "']").varResult <> vbNull Then
        printMsg WARN_, "The HTML already contains the JS Library: """ & strURL & """", FromProcedureName
        Exit Sub
    End If

   'Add script loading function
   '(Javascript Original)
    '    function loadScript(name) {
    '        var script = document.createElement('script');
    '        script.type = 'text/javascript'; script.src = name;
    '        document.head.appendChild(script);
    '        script.onload = () => {};                              //added to make sure script has been loaded
    '        return 'Script Library fully loaded';
    '    }
    jsEval "function loadScript(name) { var script = document.createElement('script'); script.type = 'text/javascript'; script.src = name; document.head.appendChild(script); script.onload = () => {}; return 'Script library loaded'; }"

   'Load the script URL
    jsEval "loadScript('" & strURL & "')"
    printMsg info_, "Successfully added the JS Library: """ & strURL & """", FromProcedureName

End Sub

Public Sub jsAddScript(FolderPath As String, ScriptName As String)
'----------------------------------------------------------------------------------
' Load the contents of the local script file directly to the HTML page. This is
' useful when you have a custom local script to add in a certain feature to a web
' application (eg adding a custom popup message).
' Updated: 01/12/2025
'----------------------------------------------------------------------------------
' Note: the library should only be added after the page is fully loaded. If
' it is loaded when the HTML document is refereshing, it risks adding the script to
' the old HTML document which will be replaced by the new one causing the script to
' be inexistent and resulting in "ReferenceError" due to its functions not defined.
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".jsAddScript"


    Dim strLocalScriptPath As String
    strLocalScriptPath = FolderPath & "\" & ScriptName
    If Dir(strLocalScriptPath) = "" Then Exit Sub

    jsEval CharConv.BytesToString(CharConv.BytesFromSavedFile(FolderPath, ScriptName)) 'Load the script contents
    printMsg info_, "Successfully added the local script: """ & strLocalScriptPath & """", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：同期モードでの JavaScript結果のうち、最後の例外情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後の「例外（Exception）」情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」の例外情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます。
'            ・これはあくまでも、`例外`のみ扱います。「try { ... } catch(e) { return e }」系統の取得は対象外です。
'            ・非同期によるあとで取り出したパターンとは、連動しません。
'***************************************************************************************************
Property Get LastJavaScriptException() As BiDiCDPJson
    Set LastJavaScriptException = lastExceptionDetailsJS
End Property



'***************************************************************************************************
'                              ■■■ CDPcommand操作関連 ■■■
'***************************************************************************************************
'* 機能　　：CDPのMethod名と、parametersを渡して、ブラウザへ命令を行い、結果の回収も同時に行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド の実行結果`result`ノード情報。失敗時は、`Nothing`となります。
'* 引数　　：methodName             method名。"Network.getCookies"や"DOM.getDocument"というのを記入します
'            params                 Json-Dictionary形式で、KeyとItemを事前に定義したオブジェクト変数を指定します。
'            StopError              `True`で、実行失敗時に`Err.Raise`で停止し、以降の処理を止めます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：This function enables advanced method invocation under CDP.
'            For a list of all methods available, see:
'               https://chromedevtools.github.io/devtools-protocol/tot/Browser/
'            Error message will be prompted automatically by ExecuteError in the Immediate Window.
'            The return object is Json String that has been converted to a dictionary object.
'            Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'
'* 更新履歴：・28/04/23 - Set this to Public as it is useful
'            ・20/03/23 - Added printMsg msgs & set to Private
'
'* 注意事項：・CDPの運営思想の都合上、多数の`Deprecated`,`Experimental`が存在します。
'              これらのタグのmethodを使用する際は、`StopError`を`True`にすることを推奨します。失敗時の以降の処理の予期せぬデータ破損などを防ぐためです。
'
'            ・`StopError`の規定値は、`True`なので、特別な理由がない限りは、省略推奨です
'            ・このメソッドは、コマンド結果が来るまで待機するメソッドです。送るコマンド内容や、サイトの挙動によっては、タイムアウトエラーになる場合があります。
'              その場合は、後述の結果を待たないタイプのメソッド`ExecuteCDPAsync`をご利用ください。
'***************************************************************************************************
Public Function ExecuteCDP(methodName As String, Optional params As Dictionary, Optional StopCDPError As Boolean = True) As BiDiCDPJson
    Const FromProcedureName As String = ThisClassName & ".ExecuteCDP"


    '1. まずは非同期実行で実行IDを回収し、IDが0になってないか確認
    Dim RunID As Long
    RunID = ExecuteCDPAsync(methodName, params, StopCDPError)
    If RunID = 0 Then Exit Function

    '2. 1回だけメッセージ取り出しを行う
    printMsg Debug_, "Waiting for the command result response for synchronous execution... : " & RunID, FromProcedureName
    TakeEvents StopCDPError

    '3. CDP-Jsonコマンドを送った際のIDが返ってくるまで、内部でループ受信処理を行います
    '前述の`TakeEvents`により
    '・結果が届いた：ループに入らず即回収
    '・結果がまだ　：ループに入って回収
    'といったパフォーマンス改善になります
    Dim ResultJsonStr As String
    ResultJsonStr = AutoWaitTakeResultCDP(RunID, StopCDPError)

    '4. 結果を、Node変換
    Dim resultJson As BiDiCDPJson
    Set resultJson = BiDiCDPJson.Parse(ResultJsonStr)

    '5. 実行時エラーチェック
    If ExecuteError(resultJson) Then
        'エラーによる`Err.Raise`で止めるか？
        If StopCDPError Then
            '5-1. エラーになったmethod名と、結果生responseを表示して停止
            printMsg ERROR_, "Failed to ExecuteCDP method """ & methodName & """" & vbCrLf & vbCrLf & "=== RawErrorResponse ===" & vbCrLf & ResultJsonStr, FromProcedureName, doRaiseError:=StopCDPError
        Else
            '5-1. エラーになったmethod名だけログ出して続行
            printMsg WARN_, "Failed to ExecuteCDP method """ & methodName & """", FromProcedureName
        End If
    Else
        '5-1. Return result object
        Set ExecuteCDP = resultJson.NodeKey("result")
    End If

    '6. デバッグ用に、As Is 出力
    If currentLog.seeRawSendMsgDbg Then printMsg Debug_, "ExecuteCDP `" & methodName & "` with result: " & ResultJsonStr, FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`ExecuteCDP`の非同期実行(コマンド結果を待たない)バージョンです
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド 実行時のID値(整理券番号のイメージでOK) ※API失敗時は0で、成功時は1以上です。
'* 引数　　：※`ExecuteCDP`と一緒のため、省略
'***************************************************************************************************
Public Function ExecuteCDPAsync(methodName As String, Optional params As Dictionary, Optional StopApiError As Boolean = True) As Long
    Const FromProcedureName As String = ThisClassName & ".ExecuteCDPAsync"


    '1. コマンド組み立て
    Dim CDPcommand As New Dictionary
    CDPcommand.Add "method", methodName
    If Not params Is Nothing Then CDPcommand.Add "params", params       '`nothing`の場合は、`params`項目を省略

    '2. ブラウザへ送信し、実行IDを取得
    ExecuteCDPAsync = PipeCore.ReadyRunCDP(CDPcommand, brTab.sessionID) 'Send message to the Context(tab) instance

    '3. 送信APIエラーチェック
    If PipeCore.LastErrorWriteFile Then
        printMsg ERROR_, "Error WriteFile in CDPCore.writeProcCDP. ErrorCode:" & PipeCore.LastErrorWriteFile & " - " & WinApiError.GetMessage(PipeCore.LastErrorWriteFile, "kernel32"), FromProcedureName, doRaiseError:=StopApiError

        ExecuteCDPAsync = 0 '失敗なので、0にセット
        Exit Function
    End If

    '4. ログ化
    printMsg Debug_, "Async ExecuteCDP: " & methodName & " - ID: " & ExecuteCDPAsync, FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：非同期実行(ExecuteCDPAsync)した際のcommandIDに対して、結果が来るまで内部で待機します
'---------------------------------------------------------------------------------------------------
'* 返り値　：commandIDに対する結果Json文字列
'* 引数　　：CommandID      実行ID
'            StopApiError   `Pipe - ReadFile`/`WebSocket - recv`でエラーが起こったら停止します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・タイムアウトの場合は、エラー停止になります。結果に時間がかかりそうなコマンドの場合は事前に`TimeOutSecond`にて伸ばして下さい
'            ・WinAPI側のエラーの場合は、`vbnullstring`で返ります
'***************************************************************************************************
Public Function AutoWaitTakeResultCDP(commandID As Long, Optional StopApiError As Boolean = True) As String
    Const FromProcedureName As String = ThisClassName & ".AutoWaitTakeResultCDP"


    '1. すでに結果が格納済みの場合は、ここの段階で抜ける
    AutoWaitTakeResultCDP = TakeResultCDP(commandID)
    If StrPtr(AutoWaitTakeResultCDP) Then printMsg Debug_, "The result has already been stored as a side effect of another process.", FromProcedureName: Exit Function

    '2. 今の経過時間を記録し、取り出し開始
    Dim timerStart As Double: timerStart = CDPHelpers.TimerCounter
    Do
        '3. 受信処理によるパイプを拝見する
        '※内部にて`DoEvents`が走ります
        TakeEvents StopApiError

        '4. 内部の`RaiseEvent`による発火で、Dictionaryへ蓄積されたか？
        AutoWaitTakeResultCDP = TakeResultCDP(commandID)
        If StrPtr(AutoWaitTakeResultCDP) Then
            '送ったコマンド結果であることを確認できたので、ループから抜ける
            printMsg info_, "Browser have confirmed receipt of command result ID: " & commandID, FromProcedureName
            Exit Do
        Else
            printMsg Trace_, "Browser have not yet been able to confirm command ID: " & commandID, FromProcedureName
        End If

        '5. 受信処理にてエラーがあったら、即抜け
        If PipeCore.LastErrorPeekNamedPipe > 0 Or PipeCore.LastErrorReadFile > 0 Then printMsg ERROR_, "Possible reasons: (1) the browser is not actually running (ie. call string failure); (2) the browser window was accidentally closed; (3) the CDP Pipe connection to the browser was cut-off somehow.", FromProcedureName, doRaiseError:=StopApiError: Exit Function

        '6. 一定時間経ってもコマンド結果が来ないなら、エラーで停止します
        If CDPHelpers.TimerCounter - timerStart > timerOut Then Err.Raise CDPCustomErrorCodes.TIMEOUT, FromProcedureName, "Timeout waiting for browser to return message." Else printMsg Trace_, (CDPHelpers.TimerCounter - timerStart) / 1000 & " seconds have passed...", FromProcedureName
    Loop
End Function



'***************************************************************************************************
'                              ■■■ CDPcommand結果関連 ■■■
'***************************************************************************************************
'* 機能　　：このタブに対するCDPレスポンス結果を一旦、Dictionaryへ蓄積させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：id             コマンドID
'            RawJson        コマンド結果の生Json文字列　※As Is. ID付きのままで渡されます
'            sessionID      届いたSessionID
'---------------------------------------------------------------------------------------------------
'* 機能説明：`ID`をKeyに、Dictionaryへ`RawJson`を格納します。
'* 注意事項：Jsonは、文字列のままで渡されます
'***************************************************************************************************
Private Sub PipeCore_CDPContextID(id As Long, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPBrowserEvent"


    '1. このタブクラスで設定中の`SesionId`以外は即抜け
    If sessionID <> brTab.sessionID Then Exit Sub

    '2. その後、結果を蓄積し、後で取り出せるようにする
    ResultCommandCDP_Context.Add id, RawJson
    printMsg Trace_, "The results of the CDP(ID: " & id & ") execution from the tab(sessionId: " & sessionID & ") have been received. → " & RawJson, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：オーバーフロー対策によるcommandIDがリセットされると呼び出されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：IDが20億に達する場合も、結果履歴が全て削除されます
'***************************************************************************************************
Private Sub PipeCore_ResetCommandID()
    ResultCommandCDP_Context.RemoveAll
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
Property Get TakeResultCDP(commandID As Long) As String
    Const FromProcedureName As String = ThisClassName & ".TakeResultCDP"


    '1. 存在するか？
    If ResultCommandCDP_Context.Exists(commandID) Then
        '2. Log化
        Dim tmp As String: tmp = ResultCommandCDP_Context(commandID)
        printMsg info_, "Browser were able to confirm the existence of CDP command ID:" & commandID & "　→　" & tmp, FromProcedureName

        '3. 返却
        TakeResultCDP = tmp

        '4. 取り出したので消しておく
        ResultCommandCDP_Context.Remove commandID
        printMsg Trace_, "CDP-command ID: " & commandID & " has been removed from `ResultCDP_Browser.response`.", FromProcedureName
    Else
        '2. 存在しない旨を伝える
        printMsg Trace_, "This CDP-command ID(" & commandID & ") passed is not running asynchronously or has been deleted.", FromProcedureName
    End If
End Property

'***************************************************************************************************
'* 機能　　：CDP-Jsonコマンド実行結果のエラーチェックを行い、最後のエラー記録も担います
'---------------------------------------------------------------------------------------------------
'* 返り値　：エラー有無論理値
'* 引数　　：result         パース済CDP-Jsonコマンド実行結果
'            ErrorCode      エラーコード　      ※参照渡しにより返されます
'            ErrorMessage   エラーメッセージ    ※参照渡しにより返されます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・Check if ExecuteCDP fails and what the error message returned is.
'            　Error message will be shown via printMsg
'
'            ・または、第2引数以降の変数からでもエラー情報をクイックに受け取れます
'---------------------------------------------------------------------------------------------------
'* Credit  ：https://github.com/PerditionC/VBAChromeDevProtocol
'***************************************************************************************************
Private Function ExecuteError(result As BiDiCDPJson, Optional ErrorCode As Long, Optional ErrorMessage As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ExecuteError"


    Dim ErrorData As Variant

    'データの状態をチェック
    If result.Count = 0 Then
        'タイムアウトもしくは、パイプは生きてるが、応答なしによる結果無し
        ExecuteError = True


        'カスタムなエラー情報として記載
        ErrorCode = -1 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        ErrorData = vbEmpty
        printMsg WARN_, "ExecuteCDP Error: [" & ErrorCode & "] " & ErrorMessage, FromProcedureName

        '最後のエラー情報を記録
        Set lastErrorDetailsCDP = New Dictionary
        lastErrorDetailsCDP.Add "code", ErrorCode
        lastErrorDetailsCDP.Add "message", ErrorMessage

    Else
        If result.ExistsKey("error") Then
            'CDPコマンドエラー
            ExecuteError = True


            '`error`オブジェクトの中身で記録
            Dim errNode As BiDiCDPJson
            Set errNode = result.NodeKey("error")

            '主要なエラー情報を抜き出す
            ErrorCode = errNode.NumberKey("code")
            ErrorMessage = errNode.StringKey("message")
            printMsg WARN_, "ExecuteCDP Error: [" & ErrorCode & "] " & ErrorMessage, FromProcedureName

            '追加情報がある場合は、それもログ化
            If errNode.ExistsKey("data") Then
                ErrorData = errNode("data")
                printMsg WARN_, "       " & ErrorData, FromProcedureName
            Else
                ErrorData = vbEmpty ' vbNullString? is lastErrorDetailsCDP("data") a string or object?
            End If

            '最後のエラー情報を記録
            Set lastErrorDetailsCDP = New Dictionary
            lastErrorDetailsCDP.Add "code", ErrorCode
            lastErrorDetailsCDP.Add "message", ErrorMessage
        Else
            'CDPコマンド正常終了
            ExecuteError = False


            'クリアしておく
            ErrorCode = 0
            ErrorMessage = vbNullString
            ErrorData = vbEmpty
        End If
    End If

End Function

'***************************************************************************************************
'* 機能　　：`ExecuteError`メソッドを通した際の、最後のエラー情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後のエラー情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」のエラー情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます
'            ・適切な`ExecuteCDP`を実行したつもりなのに返り値が、`Nothing`となるのが、エラーの目印となるので、それでこのエラー情報を見に行ってください
'            ・`ExecuteError`メソッドを通さないと機能しません
'            ・`Dictionary`オブジェクトとして得れるので必要に応じて`.RemoveAll`でエラークリアすることも可能です
'***************************************************************************************************
Property Get LastCDPJsonError() As Dictionary
    Set LastCDPJsonError = lastErrorDetailsCDP
End Property

'***************************************************************************************************
'* 機能　　：CDP-Json コマンド結果を手動でクリアします
'---------------------------------------------------------------------------------------------------
'* 返値    ：クリア件数
'---------------------------------------------------------------------------------------------------
'* 機能説明：任意のタイミングでの、メモリクリーン等にどうぞ
'***************************************************************************************************
Public Function CleanResultCommandCDP() As Long
    Const FromProcedureName As String = ThisClassName & ".CleanResultCommandCDP"


    '1. クリア件数を返却
    CleanResultCommandCDP = ResultCommandCDP_Context.Count

    '2. ALLクリア
    ResultCommandCDP_Context.RemoveAll
    printMsg info_, "The user cleared " & CleanResultCommandCDP & " previous CDP command results.", FromProcedureName
End Function



'***************************************************************************************************
'                                       ■■■ 後始末系 ■■■
'***************************************************************************************************
Public Sub closeTab()
'------------------------------------------------------------------
' Close the current tab of the session. Requires targetId which can
' be retrieved by using refreshTargetID. This is the targetId of the
' active tab under focus of the automation pipe.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".closeTab"


    'Get the original serialized tab ID
    Dim MainTargetID As String, tmp As Dictionary
    Set tmp = ShSetting01_StartBrowser.TableBrowserContext(BrowserCore.CurrentBrowserUser, FromProcedureName)
    If Not tmp Is Nothing Then MainTargetID = tmp("targetID")

    'メインタブは閉じないように仕向ける
    If brTab.targetID = MainTargetID Then
        printMsg WARN_, "Unable to close this tab " & brTab.targetID & " as it is the main session tab", FromProcedureName
        Exit Sub
    End If

    Dim params As New Dictionary
    params("targetId") = brTab.targetID
    ExecuteCDP "Target.closeTarget", params
    printMsg info_, brTab.targetID & " was closed", FromProcedureName

End Sub



'***************************************************************************************************
'                                ■■■ 非同期イベント関連 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する非同期イベントを受け取ります
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName 非同期イベントメソッド名
'            RawJson    非同期イベントの生Json文字列　※As Is. メソッド名付きのままで渡されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：非同期イベントキャプチャと、セッション状態の更新を行います
'***************************************************************************************************
Private Sub PipeCore_CDPBrowserEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPBrowserEvent"


    '------------------ 1. この`Context`のSessionステータスの更新を行います ------------------
    Dim tmp As BiDiCDPJson
    Select Case methodName
        '`SessionID`が切れたが、`targetId`は生存
        Case "Target.detachedFromTarget"
            '1-1. 自分の`SessionID`か確認し、そうであれば、フラグ更新
            If BiDiCDPJson.Parse(RawJson).NodeKey("params").StringKey("sessionId") = brTab.sessionID Then ControlStatus.detachedFromTarget = True: printMsg WARN_, "Your session has expired. You need to renew it.", FromProcedureName

        '`SessionID`は切れてないが、操作不能。ただし、`Page.reload`を使えば、息を吹き返す（復活する）かもしれない！状態
        Case "Target.targetCrashed"
            '1-1. 自分の`targetId`か確認し、そうであれば、Jsonパース処理して項目ごとに分類
            Set tmp = BiDiCDPJson.Parse(RawJson).NodeKey("params")
            If tmp.StringKey("targetId") = brTab.targetID Then
                ControlStatus.targetCrashed = True

                '1-2. エラー情報を格納
                ControlError.Status = tmp.StringKey("status")
                ControlError.ErrorCode = tmp.StringKey("errorCode")
                printMsg WARN_, "The tab crashed. Using `Page.reload` might fix it. errorCode: " & ControlError.ErrorCode & " , status: " & ControlError.Status, FromProcedureName
            End If

        '`targetId`が消えて、`SessionID`も道ずれ破棄
        Case "Target.targetDestroyed"
            '1-1. 自分の`targetId`か確認し、そうであれば、フラグ更新
            If BiDiCDPJson.Parse(RawJson).NodeKey("params").StringKey("targetId") = brTab.targetID Then ControlStatus.targetDestroyed = True: printMsg ERROR_, "The tab has been closed. This Class object is no longer usable. Please discard it.", FromProcedureName

        '一部のパラメーター(URL等)を更新します
        Case "Target.targetInfoChanged"
            '1-1. 自分の`targetId`か確認し、そうであれば、Jsonパース処理して項目ごとに分類
            Set tmp = BiDiCDPJson.Parse(RawJson).NodeKey("params").NodeKey("targetInfo")
            If tmp.StringKey("targetId") = brTab.targetID Then
                '1-2. 各種パラメーター更新
                With brTab
                    .Title = tmp.StringKey("title")
                    .Url = tmp.StringKey("url")
                    .browserContextId = tmp.StringKey("browserContextId")
                End With

                printMsg info_, "The information on this tab has been updated.", FromProcedureName
            End If

    End Select


    '------------------ 2. 非同期イベントを収集します ------------------
    '2-1. イベントキャプチャしない場合は、即抜け
    If publicBrowserEvents.Queue Is Nothing Then printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName: Exit Sub

    '2-2. その非同期イベント名称は、フィルターに含まれない場合も即抜け
    printMsg Trace_, "Checking public event filter settings...", FromProcedureName
    If Not checkFilterEvent(publicBrowserEvents.Filter, methodName) Then Exit Sub

    '2-3. パース処理
    Dim tmpDic As New Dictionary
    printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
    Set tmpDic = WebJsonConverter.Parse(RawJson).value

    '2-4. 専用のプロシージャにて、蓄積処理
    printMsg Trace_, "The module is requesting the accumulation of event information. Saving...", FromProcedureName
    OnNewEventReceived publicBrowserEvents.Queue, tmpDic, methodName
End Sub

'***************************************************************************************************
'* 機能　　：このタブに対する非同期イベントを受け取ります
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName 非同期イベントメソッド名
'            RawJson    非同期イベントの生Json文字列　※As Is. メソッド名付きのままで渡されます
'            sessionID  届いたSessionID
'---------------------------------------------------------------------------------------------------
'* 機能説明：非同期イベントキャプチャを行います
'***************************************************************************************************
Private Sub PipeCore_CDPContextEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPContextEvent"


    '1. このタブクラスで設定中の`SessionId`以外は即抜け
    If sessionID <> brTab.sessionID Then Exit Sub

    '2. このタブクラス独自のイベント処理を行う
    Dim tmpNode As BiDiCDPJson
    Select Case methodName
        'JavaScript実行環境情報の更新を行います
        Case "Runtime.executionContextCreated"
            '2-1. このタブクラスのTargetIDと一致するか？
            Set tmpNode = BiDiCDPJson.Parse(RawJson).NodeKey("params").NodeKey("context")
            If tmpNode.NodeKey("auxData").StringKey("frameId") = brTab.targetID Then
                '2-2. 各種statusを上書き
                With brTabRuntime
                    .IsValid = True '使えるよサイン

                    .ExecutionContextId = tmpNode.NumberKey("id")
                    .origin = tmpNode.StringKey("origin")
                    .Name = tmpNode.StringKey("name")
                    .uniqueId = tmpNode.StringKey("uniqueId")
                End With

                '2-3. ログ化
                printMsg info_, "The JavaScript execution environment for targetID:" & brTab.targetID & " has been updated.", FromProcedureName
            End If

        'JavaScript実行環境情報の破棄更新を行います
        Case "Runtime.executionContextDestroyed"
            '2-1. このClassで保持中の`executionContextUniqueId`と一致するか？
            If brTabRuntime.uniqueId = BiDiCDPJson.Parse(RawJson).NodeKey("params").StringKey("executionContextUniqueId") Then
                brTabRuntime.IsValid = False
                printMsg info_, "The JavaScript environment owned by this class has been destroyed.", FromProcedureName
            End If

        'JavaScript実行環境情報の破棄更新を行います。こちらは、コンテキスト(タブ)全体です
        Case "Runtime.executionContextsCleared"
            brTabRuntime.IsValid = False
            printMsg info_, "The JavaScript environment associated with this class was destroyed due to a page transition or other reasons.", FromProcedureName

    End Select


    '------------------ 3. 非同期イベントを収集します ------------------
    '3-1. イベントキャプチャしない場合は、即抜け
    If publicBrowserEvents.Queue Is Nothing Then printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName: Exit Sub

    '3-2. その非同期イベント名称は、フィルターに含まれない場合も即抜け
    printMsg Trace_, "Checking public event filter settings...", FromProcedureName
    If Not checkFilterEvent(publicBrowserEvents.Filter, methodName) Then Exit Sub

    '3-3. パース処理
    printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
    Dim tmp As Dictionary: Set tmp = WebJsonConverter.Parse(RawJson).value

    '3-4. 専用のプロシージャにて、蓄積処理
    printMsg Trace_, "The module is requesting the accumulation of event information. Saving...", FromProcedureName
    OnNewEventReceived publicBrowserEvents.Queue, tmp, methodName
End Sub



'***************************************************************************************************
'                                ■■■ 非同期イベント収集関連 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザからのイベント内容に関するプロパティを行います
'---------------------------------------------------------------------------------------------------
'* GET返値 ：呼び出した時点での、ブラウザイベント内容
'* SET引数 ：置き換えるDictionary変数。基本は、New Dictionary を渡してのリセットです
'---------------------------------------------------------------------------------------------------
'* 機能説明：・`Set BrowserEvents` に、`New Dictionary`を渡すと、イベント記録モードになります。
'            ・`Set BrowserEvents` に、`Nothing`を渡すと、イベント破棄モードになります。
'* 注意事項：Set プロパティ使用時、簡易的なDictionary構造のチェックを行います。構造が違うDictionaryを渡された場合、警告メッセージが出て、停止します。
'***************************************************************************************************
Property Get BrowserEvents() As Dictionary
    Set BrowserEvents = publicBrowserEvents.Queue
End Property

Property Set BrowserEvents(ObjDic As Dictionary)
    Const FromProcedureName As String = ThisClassName & ".BrowserEvents"


    'SetされたDictionaryの構造をチェックします
    '1. そもそもDictionary オブジェクトが備わっているか？
    If ObjDic Is Nothing Then
        'Nothing 処理。以降は、イベントキャプチャはされません。
        Set publicBrowserEvents.Queue = Nothing
        printMsg info_, "Event capture save data has been discarded. Event capture will not be performed.", FromProcedureName

    '2. まっさらの場合
    ElseIf ObjDic.Count = 0 Then
        '初期化処理。これにより、イベントキャプチャを開始します。
        Set publicBrowserEvents.Queue = FormatDictionary_BrowserEvents
        printMsg info_, "A new event capture save data has been created. Starts event capture.", FromProcedureName

    '3. 指定のKeyがあるか？
    ElseIf ObjDic.Exists(EventsDictionaryKeyName01) And ObjDic.Exists(EventsDictionaryKeyName02) And ObjDic.Count = 2 Then
        '既存データをセット。これにより、このDictionaryにイベントキャプチャの追記が行われます。
        Set publicBrowserEvents.Queue = ObjDic
        printMsg info_, "Existing event capture save data was successfully loaded. Resume event capture with this save data.", FromProcedureName

    '4. 不正なDictionary 形式なのでエラーで止める
    Else
        printMsg ERROR_, "Event capture save data structure is invalid.", FromProcedureName, doRaiseError:=True
    End If
End Property

'***************************************************************************************************
'* 機能　　：ブラウザからのイベントのうち、指定のイベントメソッド名のみ、Jsonパース、蓄積をさせる設定です
'---------------------------------------------------------------------------------------------------
'* 引数    ：DelMode    `True`  指定イベントメソッド名のフィルターを削除し以後、そのイベントメソッド名の蓄積はしません
'                       `False` 指定イベントメソッド名のフィルターを追加し以後、フィルターに追加された分のイベントメソッド名のみ蓄積できます
'---------------------------------------------------------------------------------------------------
'* 機能説明：特に`Network`イベントは、大量のイベントが来てJsonパース処理の負荷を抑えるのに効果的です
'* 注意事項：フィルター数が0の場合は、フィルター機能は果たしません
'***************************************************************************************************
Property Let SetFilterEvents(Optional DelMode As Boolean, EventName As String)
    Const FromProcedureName As String = ThisClassName & ".SetFilterEvents"


    '1. イベント名が空欄か？
    If LenB(EventName) = 0 Then
        '1-2. 削除モードか？
        If DelMode Then
            'LETのイベントメソッド名が空欄の場合は、全削除します
            publicBrowserEvents.Filter.RemoveAll
            printMsg info_, "All event method filters have been removed.", FromProcedureName
        Else
            printMsg WARN_, "Event method names that are left blank cannot be added to the filter.", FromProcedureName
        End If

    Else
        '2. 既にイベントメソッド名を持ってるか
        If publicBrowserEvents.Filter.Exists(EventName) Then
            '2-2. 削除モードか？
            If DelMode Then
                'LETで指定したイベントメソッド名を削除します
                publicBrowserEvents.Filter.Remove (EventName)
                printMsg info_, "The specified event method name `" & EventName & "` has been removed from the filter.", FromProcedureName
            Else
                printMsg WARN_, "The specified event method name `" & EventName & "` has already been added to the filter.", FromProcedureName
            End If
        Else
            '2-2. 削除モードか？
            If DelMode Then
                printMsg WARN_, "The specified event method name `" & EventName & "` is not registered in the filter.", FromProcedureName
            Else
                'LETで指定したイベントメソッドを追加します
                publicBrowserEvents.Filter.Add EventName, True
                printMsg info_, "The specified event method name `" & EventName & "` has been added to the filter.", FromProcedureName
            End If
        End If
    End If

    '3. ログ用に表示
    Dim FilterCount As Long: FilterCount = publicBrowserEvents.Filter.Count
    printMsg Debug_, "Current number of filters: " & FilterCount, FromProcedureName
    If FilterCount = 0 Then printMsg info_, "Because no filter is applied, all event methods are captured.", FromProcedureName
End Property

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
'* 機能　　：ブラウザイベント情報を格納する前に、フィルター情報を取得して、イベント蓄積させるか判定させます
'---------------------------------------------------------------------------------------------------
'* 返り値  ：蓄積可否論理値
'* 引数　　：FilterDictionary   対象のフィルターDictionary
'            EventMethodName    非同期イベント名称
'***************************************************************************************************
Private Function checkFilterEvent(FilterDictionary As Dictionary, EventMethodName As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".checkFilterEvent"


    '0. オブジェクト自体が設定されてない場合は一律、キャプチャしない判定とします
    If FilterDictionary Is Nothing Then checkFilterEvent = False: printMsg Trace_, "Since no filter object is set, capturing all event methods is disabled.", FromProcedureName: Exit Function

    '1. フィルター情報があるか？
    If FilterDictionary.Count Then
        '2. そのイベントメソッド名があるか？
        If FilterDictionary.Exists(EventMethodName) Then
            'あれば、capture対象
            checkFilterEvent = True
            printMsg Trace_, "Filter registration detection: " & EventMethodName, FromProcedureName
        Else
            'なければ、captureしない
            checkFilterEvent = False
            printMsg Trace_, "Filter not registered: " & EventMethodName, FromProcedureName
        End If
    Else
        'フィルター情報がない場合は、全てキャプチャさせます
        checkFilterEvent = True
        printMsg Trace_, "Since no filter is set, all event methods will be captured.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザイベント情報を格納するためのフォーマットDictionaryを作成します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：フォーマット済みDictionary
'***************************************************************************************************
Private Function FormatDictionary_BrowserEvents() As Dictionary
    '新規オブジェクトセット
    Set FormatDictionary_BrowserEvents = New Dictionary

    '初期構造を作成します
    FormatDictionary_BrowserEvents.Add EventsDictionaryKeyName01, 0                 '合計イベントキャプチャ数
    FormatDictionary_BrowserEvents.Add EventsDictionaryKeyName02, New Dictionary    'イベントキャプチャDictionaryデータ
End Function



'***************************************************************************************************
'                                   ■■■ ステータス系 ■■■
'***************************************************************************************************
Public Sub printParams()
'------------------------------------------------------------------
' Print all parameters of the current session for debuggings.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".printParams"


    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, currentLog.LogID & " Tab - Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "Session ID        | sessionID    | " & brTab.sessionID, FromProcedureName
    printMsg Debug_, "Window Handle     | brHWnd       | " & brTab.hWnd, FromProcedureName
    printMsg Debug_, "Window ID         | brWId        | " & brTab.WID, FromProcedureName
    printMsg Debug_, "Tab URL           | brUrl        | " & brTab.Url, FromProcedureName
    printMsg Debug_, "Tab Title         | brName       | " & brTab.Title, FromProcedureName
    printMsg Debug_, "Tab ID            | targetID     | " & brTab.targetID, FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

End Sub

Public Sub getSessionInfo()
'----------------------------------------------------------------------------------
' Obtain session info on the fly for printParams.
' Updated: 22/03/2023
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".getSessionInfo"


    'Get infos
    Dim TargetInfo As BiDiCDPJson: Set TargetInfo = refreshTargetInfo
    brTab.Url = TargetInfo.StringKey("url")
    brTab.Title = TargetInfo.StringKey("title")
    brTab.hWnd = getHandle

    '対象タブが非表示の場合、`Browser.getWindowForTarget`が効かなくなるので、判定をかます
    Dim tmp As BiDiCDPJson: Set tmp = ExecuteCDP("Browser.getWindowForTarget", StopCDPError:=False)
    If tmp Is Nothing Then printMsg info_, "The tab you are connecting to may be hidden.", FromProcedureName Else brTab.WID = tmp.StringKey("windowId")

    'タブIDを記録
    brTab.targetID = TargetInfo.StringKey("targetId")

    'ProfileIDを記録
    brTab.browserContextId = TargetInfo.StringKey("browserContextId")

    'JavaScript関連の変化も捉えれるようにする
    '※`reattach`を想定して一旦OFFにして、再度ONにすることで、`Runtime.executionContextCreated`等が飛んでくれます
    runtimeEnable False
    runtimeEnable True
End Sub

'***************************************************************************************************
'* 機能　　：現時点で開いている全てのiframe情報をログに表示します
'---------------------------------------------------------------------------------------------------
'* 機能説明：名前またはURLに基づいて検索しヒットしたら、`executionContextId`を返却します
'* 注意事項：検索範囲は、親の子iframeまでです。親の子iframe内のiframe...といった多重ネスト表示は非対応です
'***************************************************************************************************
Public Sub printChildFrames()
    Const FromProcedureName As String = ThisClassName & ".printChildFrames"


    Dim NodeToken As Long
    Dim iframeCount As Long
    Dim childFrames As BiDiCDPJson

    '1. 実行して、まずは結果を取得
    Dim tmp As BiDiCDPJson
    Set tmp = ExecuteCDP("Page.getFrameTree").NodeKey("frameTree")

    '2. iframeがない場合は即抜け
    If Not tmp.ExistsKey("childFrames") Then printMsg info_, "'" & brTab.Title & "' does not have an iframe.", FromProcedureName: Exit Sub

    '3.Read info
    printMsg info_, "---------------------------------------------------", FromProcedureName
    printMsg info_, "childFrames Information", FromProcedureName
    printMsg info_, "---------------------------------------------------", FromProcedureName

    With tmp.NodeKey("childFrames")
        NodeToken = .FirstChildToken  ' 配列の先頭トークン
        Do While NodeToken > 0
            '`childFrames`の各`frame`要素を取り出す
            Dim childFrame As BiDiCDPJson
            Set childFrame = .NodeFromToken(NodeToken).NodeKey("frame")
            iframeCount = iframeCount + 1

            With childFrame
                printMsg info_, "#" & iframeCount & vbTab & _
                            "type: " & .StringKey("mimeType") & " | " & _
                            "title: " & .StringKey("name") & " | " & _
                            "url: " & .StringKey("url") & " | " & _
                            "frameID: " & .StringKey("id"), FromProcedureName
            End With

            '次の`Target.TargetInfo`へ
            NodeToken = .NextToken(NodeToken)
        Loop
    End With
End Sub

Private Function getHandle() As LongPtr
'-----------------------------------------------------------------------------------
' Added by Long: obtain the window handle of the current browser session
' Updated: 02/03/23 - Add mechanism to work with browser windows that are blank.
'-----------------------------------------------------------------------------------

    Dim bufferString As String
    Dim titleLen As Long
    Dim hWnd As LongPtr

    'Determine current tab title, if error: the tab is a blank tab
    Dim sessionTitle As String, windowTitle As String
    sessionTitle = jsEval("document.getElementsByTagName(""title"")[0].innerHTML", IFEXCEPTION:="about:blank")

    'Obtain handles to the correct browser window
    'Both Edge and Chrome use the same class chromeWindowClass
    hWnd = FindWindowEx(0&, 0&, StrPtr(BrowserCore.CurrentBrowserWindowClass), 0&)        'To return Chrome / Edge window handles
    Do
        '1. 取得したウィンドウハンドル(hwnd)の「プロセスID」を取得する
        Dim currentPID As Long
        GetWindowThreadProcessId hWnd, currentPID

        '2. プロセスIDが、探しているブラウザのものと一致するか？
        If currentPID = PipeCore.Connection_dwProcessId Then
            '一致すれば少なくとも、デバッグブラウザのウィンドウハンドルなので一旦確定させる
            getHandle = hWnd

            '3. Get Window Title based on specified handles
            'Requires string buffering declaration as string
            bufferString = String(1024, " ")
            titleLen = GetWindowText(hWnd, StrPtr(bufferString), Len(bufferString))
            windowTitle = Left(bufferString, titleLen)

            'Verify if matching with the sessionTitle and is really an edge/chrome window
            If LenB(windowTitle) Then If InStr(windowTitle, sessionTitle) Then Exit Do
        End If

        'Continue with the next handle value
        hWnd = FindWindowEx(0&, hWnd, StrPtr(BrowserCore.CurrentBrowserWindowClass), 0&)

    Loop Until hWnd = 0 ' no more matching windows

End Function



'***************************************************************************************************
'                           ■■■ 各種CDPヘルパーコマンド ■■■
'***************************************************************************************************
Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
'-----------------------------------------------------------------------
' This sub is a quick wrapper to navigate to an url.
' Updated 23/03/23
'-----------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".navigate"


   'Check if already navigated
    Dim curURL As String
    curURL = jsEval("window.location.href")
    If curURL = checkURI(strURL) Then
        printMsg WARN_, "Already navigated to """ & strURL & """", FromProcedureName
        Exit Sub
    End If

   'If not yet then do navigate
    Dim CDPparams As New Dictionary
    CDPparams.Add "url", strURL
    ExecuteCDP "Page.navigate", CDPparams

    'should also add a maxTimeLimit to Wait
    wait till:=till
    printMsg info_, "Navigated to: " & strURL, FromProcedureName

End Sub

Public Sub wait(Optional till As ReadyState = isComplete, Optional dbgState As Boolean = False)
'--------------------------------------------------------------------------------------------------
' Wait until browser reaches a specific state. Default is until it fully loads ("complete")
' Possible states:  loading / interactive / complete
' Complete state is preferable when certain information is only showed when fully loaded.
' Otherwise, an interactive state is good enough and can potentially save a lot of wait time.
' However, do note that some web apps only fully displays elements at complete readystate.
' To read ReadyState on the go, set the dbgState to True to print the state to the Immediate Window
' Original Name: waitCompletion
' Updated: 26/04/23 - Enhanced 'till' to Enum type for ease of use
'                   - Added exit for targetState = interactive in case CDP misses the state change
'--------------------------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".wait"


    Dim strState As String, targetState As String
    targetState = EnumToStringReadyState(till)

    '初期値を設定しておく
    '最初はほぼ必ずloadingだから、`targetState`が`loading`になってたら、わざわざチェックする必要性はない
    strState = EnumToStringReadyState(isLoading)


    Do Until strState = targetState
        strState = jsEval("document.readyState", StopException:=dbgState)
        printMsg Debug_, "Current ReadyState: " & strState, FromProcedureName   'option to print ReadyState info during iteration
        If targetState = EnumToStringReadyState(isInteractive) And strState = EnumToStringReadyState(isComplete) Then Exit Do   'also exit if better than interactive is satisfied
        CDPHelpers.Sleep 0.1                                                               'reduce sleep will speed up but will cost cpu power
    Loop

    printMsg Debug_, "Wait '" & targetState & "' satisfied", FromProcedureName

End Sub

Public Sub show(Optional state As WinState = asNormal, Optional xywh As String = "0 0 0 0")
'-----------------------------------------------------------------------------------
' Display the window with various optional state. The default state is "normal".
' Can also use after .attach to bring the window topmost as well. By default at
' start, the browser window's state will be that of the previous session run.
' xywh is used for inputing left top width height parameters:
' eg.1: xywh = "100 200" left = 100 and top = 200, width and height unchanged
' eg.2: xywh = "100 100 300 400" left = top = 100, width = 300 and height = 400
' eg.3: xywh = "0 0 0 0" the method will ignore and not make any resizing
'-----------------------------------------------------------------------------------
' Note: This method targets the window based on its name so it is best applied when
' the browser window is not refreshing a URL which may potentially change the name.
' Updated: 02/03/23 - Changed CInt to Int to prevent Overflow error.
'-----------------------------------------------------------------------------------

    If brTab.hWnd = 0 Then brTab.hWnd = getHandle
    ShowWindow brTab.hWnd, doShow 'Make visible
    ShowWindow brTab.hWnd, state

   'Resize to specific position and size
   'Will not work properly if window is currently maximized
    If xywh <> "0 0 0 0" Then

        CDPHelpers.Sleep 0.3      'Sleep a bit for previous Win Api executions to finish, else the resize effect will be accidentally cancelled out

        Dim params As New Dictionary
        If LenB(brTab.WID) = 0 Then brTab.WID = ExecuteCDP("Browser.getWindowForTarget").NumberKey("windowId")
        params("windowId") = Int(brTab.WID)     '02/03/23 CInt() will throw error, Int() is better

        Dim bounds As New Dictionary
        Dim x As Long, y As Long, w As Long, h As Long
        On Error Resume Next
        x = CLng(Split(xywh, " ")(0))
        y = CLng(Split(xywh, " ")(1))
        w = CLng(Split(xywh, " ")(2))
        h = CLng(Split(xywh, " ")(3))
        On Error GoTo 0

        If x <> 0 Then bounds("left") = x
        If y <> 0 Then bounds("top") = y
        If w <> 0 Then bounds("width") = w
        If h <> 0 Then bounds("height") = h
        Set params("bounds") = bounds
        ExecuteCDP "Browser.setWindowBounds", params

    End If

End Sub

Public Sub notify(msg As String, Optional ViewSecond As Long = 10)
'----------------------------------------------------------------------------------
' Send a custom notification message on the bottom right corner of the screen. This
' popup message is nicer than alert() and it is not modal: ie. does not the user to
' interact with it while pausing the automation.
' Credit: Claude AI
' Updated: 01/12/2025
'----------------------------------------------------------------------------------

    ' 1. 引数をパッキング（CDPの仕様に合わせた形）
    Dim args(1) As Dictionary
    Dim arg1 As New Dictionary: arg1.Add "value", msg:                  Set args(0) = arg1
    Dim arg2 As New Dictionary: arg2.Add "value", ViewSecond * 1000:    Set args(1) = arg2

    ' 2. JS関数本体（無名関数として定義し、その場で実行させる）
    Dim funcJS As String
    funcJS = "function(msg, duration) {" & _
                "let c = document.getElementById('statusContainer');" & _
                "if (!c) {" & _
                    "c = document.createElement('div');" & _
                    "c.id = 'statusContainer';" & _
                    "c.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:9999;';" & _
                    "document.body.appendChild(c);" & _
                "}" & _
                "const m = document.createElement('div');" & _
                "m.style.cssText = 'background:#4CAF50;color:white;padding:16px 24px;margin-top:10px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3);white-space:nowrap;transition:all 0.3s;opacity:0;transform:translateX(400px);font-size:16px;font-family:system-ui,-apple-system,BlinkMacSystemFont,Roboto,sans-serif;font-weight:300;line-height:1.5;';" & _
                "m.textContent = msg;" & _
                "c.appendChild(m);" & _
                "setTimeout(() => { m.style.opacity = '1'; m.style.transform = 'translateX(0)'; }, 10);" & _
                "setTimeout(() => {" & _
                    "m.style.opacity = '0';" & _
                    "m.style.transform = 'translateX(400px)';" & _
                    "setTimeout(() => c.removeChild(m), 300);" & _
                "}, duration);" & _
            "}"

    '3. Run!
    jsEval funcJS, objectArguments:=args

End Sub

Public Sub snapPage(FolderPath As String, FileName As String, Optional getFullPage As Boolean = False)
'----------------------------------------------------------------------------------
' Capture a screenshot of the entire HTML page from the top to bottom. The method
' employs html2canvas external library is which loaded via addJsLibrary. To snap
' only a certain element within the page, indicate the same with the targetElement
' argument (default is "document.body" for the entire page).
' Original source: https://stackoverflow.com/a/51478809
' Updated: 14/04/23
'          23/10/25 Daniel Polak - fixed bug where the all files in the Downloads folder were deleted instead of just the temporary screenshot
'          24/10/25 Daniel Polak - added CDP native screenshot capability which avoids use of external JavaScript library
'          07/01/26 Long Vh      - removed the use of html2canvas js lib in favor of using the native CDP method to capture the page. Overhauled the function design.
'----------------------------------------------------------------------------------
' Note : if getFullPage = false, snapPage will only get a screenshot of the current
'        view. If true, it will get all the way from the start to the bottom. See
'        demo.getSnapShot sub procedure for example.
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".snapPage"


    Dim params As New Dictionary
    Dim r As BiDiCDPJson
    Dim b64 As String, tmpPng As String

   'Verify if the page can be captured with screenshot
    Set r = ExecuteCDP("Page.getLayoutMetrics")
    If r Is Nothing Then
        printMsg WARN_, "Unable to perform snapPage on this screen as it has no layout metrics.", FromProcedureName
        Exit Sub
    End If

   'Perform the snip
    If getFullPage Then params.Add "captureBeyondViewport", True
    Set r = ExecuteCDP("Page.captureScreenshot", params)

   'Save to the same folder if not indicated
    If Not r Is Nothing Then
        b64 = r.StringKey("data")
        SaveBase64ToFile b64, FolderPath, FileName
        printMsg info_, "Screenshot saved to: " & FolderPath & "\" & FileName, FromProcedureName
    End If

End Sub

Public Sub activate()
'----------------------------------------------------------------------------------
' Switch focus to the current tab being attached to under the current BRID.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

    Dim params As New Dictionary
    params("targetId") = brTab.targetID
    ExecuteCDP "Target.activateTarget", params

End Sub

Public Sub handleDialog(Accept As Boolean, Optional promptText As String)
'----------------------------------------------------------------------------------
' Accepts or dismisses a JavaScript initiated dialog (alert, confirm, prompt, or
' onbeforeunload). accept = True to accept the dialog, else dimiss it. promptText
' is the text to enter into the dialog prompt before accepting, used only if this is a
' prompt dialog.
' Note: this method requires the execution of "pageEnable" before the dialog appears
' so that the dialog event is probably exposed and captured as an event.
' Updated: 23/07/2025 - added as requested by Jane Zhang for her project
'----------------------------------------------------------------------------------

    Dim params As New Dictionary

    params("accept") = Accept
    If StrPtr(promptText) Then params("promptText") = promptText
    ExecuteCDP "Page.handleJavaScriptDialog", params

End Sub

Public Sub pageEnable(Optional enable As Boolean = True)
'----------------------------------------------------------------------------------
' Enable the page for events listening. Required for handleDialog for instance.
'----------------------------------------------------------------------------------

    ExecuteCDP IIf(enable, "Page.enable", "Page.disable"), , False

End Sub

'***************************************************************************************************
'* 機能　　：`Runtime`系のイベント機能を有効/無効化を行います
'***************************************************************************************************
Public Sub runtimeEnable(Optional enable As Boolean = True)
    ExecuteCDP IIf(enable, "Runtime.enable", "Runtime.disable"), , False
End Sub

'***************************************************************************************************
'* 機能　　：`SessionID`に紐づく`Target.TargetInfo`オブジェクトを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`Target.TargetInfo`オブジェクト。　※取得失敗時は、`Nothing`で返ります。大抵は、context(タブ等)が消えてることが多いです
'***************************************************************************************************
Private Function refreshTargetInfo() As BiDiCDPJson
    '1. このClassで設定中の`SessionID`を基に情報取得を試みる
    Dim TargetInfo As BiDiCDPJson
    Set TargetInfo = ExecuteCDP("Target.getTargetInfo")

    '2. エラーがなければ、`targetInfo`の中身を返却
    If TargetInfo.ExistsKey("targetInfo") Then Set refreshTargetInfo = TargetInfo.NodeKey("targetInfo")

End Function

'***************************************************************************************************
'* 機能　　：`targetId`に対する`DevTools`ウィンドウを開きます
'---------------------------------------------------------------------------------------------------
'* 引数　　：panelId    最初に開きたい項目があれば指定
'---------------------------------------------------------------------------------------------------
'* 機能説明：F12が効かないページでも`DevTools`ウィンドウを開かせます
'* 注意事項：WebView2,Electronの場合は、効かない場合があります
'***************************************************************************************************
Public Sub openDevTools(Optional panelId As String)
    '`CDPBrowser`側で実行させる
    BrowserCore.openDevTools brTab.targetID, panelId
End Sub

'***************************************************************************************************
'* 機能　　：現在のページに埋め込まれている`iframe`に対する`executionContextId`を返却します
'---------------------------------------------------------------------------------------------------
'* 返り値　：検索条件に紐づく`jsEval - ContextId`で扱える`executionContextId`    ※1つもない場合は、`0`で返却
'* 引数　　：iframeName     iframe名    第一優先(<iframe name="[これを指定]")
'            Url            URL         第二優先
'            doRetrySecond  指定秒以内に出るまでリトライを繰り返します。0で、1回のみです
'---------------------------------------------------------------------------------------------------
'* 機能説明：名前またはURLに基づいて検索しヒットしたら、`executionContextId`を返却します
'* 注意事項：・これらの引数を両方とも省略すると、`getIFrameContextID`は最初のiframeの`executionContextId`を返却します
'            ・返った`executionContextId`を`CDPElement`で便利に扱いたい場合は別途`getIFrame`メソッドをご利用ください
'            ・検索範囲は、親の子iframeまでです。親の子iframe内のiframe...といった多重ネストは非対応です。その場合は自力で「ExecuteCDP("Page.getFrameTree").NodeKey("frameTree")」をして目的のを取り出してください。
'***************************************************************************************************
Public Function getIFrameContextID(Optional iframeName As String, Optional Url As String, Optional doRetrySecond As Double) As Long
    Const FromProcedureName As String = ThisClassName & ".getIFrameContextID"


    Dim NodeToken_childFrames   As Long
    Dim Node_childFrames        As BiDiCDPJson
    Dim frameID                 As String, candidateframeID As String
    Dim iframefound             As Boolean
    Dim bestScore As Long, score As Long
    Dim timerStart              As Double


    '1. タイマースタートして探索を開始する
    If doRetrySecond > 0 Then timerStart = CDPHelpers.TimerCounter
    Do
        '2. 実行して、まずは`frameTree`を取得
        Dim frameTreeNode As BiDiCDPJson
        Set frameTreeNode = ExecuteCDP("Page.getFrameTree").NodeKey("frameTree")

        '3. iframeがない場合は即抜け
        If Not frameTreeNode.ExistsKey("childFrames") Then printMsg info_, "'" & brTab.Title & "' does not have an iframe.", FromProcedureName: Exit Function

        '4. `childFrames`オブジェクトリストを取得します
        Set Node_childFrames = frameTreeNode.NodeKey("childFrames")
        NodeToken_childFrames = Node_childFrames.FirstChildToken  ' 配列の先頭トークン

        '5. スコア方式を利用して、理想の対象iframe情報を取得します
        Const MaxScore As Long = 100
        bestScore = -1
        candidateframeID = vbNullString
        Do While NodeToken_childFrames > 0
            '5-1. `childFrames`の各`frame`要素を取り出す
            Dim childFrame As BiDiCDPJson
            Set childFrame = Node_childFrames.NodeFromToken(NodeToken_childFrames).NodeKey("frame")

            '5-2. 現スコアをリセット
            score = -1

            '5-3. 入力した引数に応じた確認
            Select Case True
                '第一優先：iframe名
                Case StrPtr(iframeName)
                    Dim frameName As String: frameName = childFrame.StringKey("name")

                    '完全一致
                    If frameName = iframeName Then
                        score = 100
                    '部分一致
                    ElseIf InStr(1, frameName, iframeName) Then
                        score = 50
                    '一致せず
                    Else
                        printMsg WARN_, "with name containing """ & iframeName & """", FromProcedureName
                    End If

                '第二優先：URL
                Case StrPtr(Url)
                    Dim frameNameUrl As String: frameNameUrl = childFrame.StringKey("url")

                    '一致位置検査
                    Dim StartPos As Long: StartPos = InStr(1, frameNameUrl, Url)

                    '先頭一致
                    If StartPos = 1 Then
                        score = 90
                    '部分一致
                    ElseIf StartPos Then
                        score = 40
                    '一致せず
                    Else
                        printMsg WARN_, "with url containing """ & Url & """", FromProcedureName
                    End If

                '第三優先：直近のiframeタブ
                Case Else
                    score = 80

            End Select

            '5-4. 記録中のベストスコアを上回ったら更新
            If score > bestScore Then
                bestScore = score
                candidateframeID = childFrame.StringKey("id")
            End If

            '5-5. 最大スコアが出てるなら、捜査打ち切り
            If bestScore = MaxScore Then Exit Do

            '5-6. 次の`childFrames`へ
            NodeToken_childFrames = Node_childFrames.NextToken(NodeToken_childFrames)
        Loop

        '6. 目的の物っぽいのがあったら見つけたとします
        If bestScore > 0 Then
            iframefound = True
            frameID = candidateframeID
        End If

    '7. Do retry as long as not timeout
    Loop Until iframefound Or CDPHelpers.TimerCounter - timerStart > doRetrySecond * 1000

    '8. Returns nothing if not found
    'Nothing is preferred over throwing error
    If Not iframefound Then
        Dim info As String
        info = IIf(LenB(iframeName), iframeName, IIf(LenB(Url), Url, "(no new tab found)"))
        printMsg WARN_, "Unable to find the iframe: " & info, FromProcedureName
        Exit Function
    Else
        printMsg info_, "Found the iframe with ID: " & frameID & ". (Certainty：" & bestScore & "%)", FromProcedureName
    End If

    '9. 得たIDを基に、`contextId`を生成し返却
    Dim CDPparams As New Dictionary
    CDPparams.Add "frameId", frameID
    getIFrameContextID = ExecuteCDP("Page.createIsolatedWorld", CDPparams).NumberKey("executionContextId")
End Function



'***************************************************************************************************
'                              ■■■ WindowsAPI操作系 ■■■
'***************************************************************************************************
Public Sub hide()
'------------------------------------------------------------------
' This is the equivalent of show "hidden" but this function is
' added to make it easy to access and use the same feature.
'------------------------------------------------------------------

    If brTab.hWnd = 0 Then brTab.hWnd = getHandle
    ShowWindow brTab.hWnd, asHidden

End Sub

Public Function bringToForeground()
'------------------------------------------------------------------
' Brings the browser window to the foreground
' Updated: Long Vh - 06/01/26 should be a function like others
'------------------------------------------------------------------

    If brTab.hWnd = 0 Then brTab.hWnd = getHandle
    If brTab.hWnd <> 0 Then
        ShowWindow brTab.hWnd, doShow
        BringWindowToTop brTab.hWnd
        SetForegroundWindow brTab.hWnd
    End If

End Function



'***************************************************************************************************
'                           ■■■ `CDPElement.cls`継承準備 ■■■
'***************************************************************************************************
Public Function getElementByID(strID As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class from the element ID, similar to the getElementByID of JS.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

    '1. ID値で、検索実行
    Dim SearchJS As String, resultSearch As Variant
    SearchJS = "document.getElementById(""" & strID & """)"
    resultSearch = jsEval(SearchJS)

    '2. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '3. 結果を`CDPElement.cls`に格納
    Set getElementByID = New CDPElement
    getElementByID.Init Me, JsResultVarType, objectId, SearchJS

End Function

Public Function getElementByXPath(strXPath As String) As CDPElement
'---------------------------------------------------------------------------------------
' Get an element class from XPath. The element class is init in CDPElement module.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' XPath Syntax help: https://www.w3schools.com/xml/xpath_syntax.asp
' Updated: 02/03/2023
' Examples -----------------------------------------------------------------------------
' eg.1 "//b[text()='T-BILL CURVE']": tag "b" with innertext to be 'T-BILL CURVE'
' eg.2 "//input[@name='username']": tag "input" & attribute "name" to be 'username'
' eg.3 "//a[@a1='abc' and @a2='xyz']": tag "a" with attribute a1 = abc and a2 = xyz
' eg.4 "//div[@id='abc']/div": get first div child of the div with id = abc
' eg.5 "//div[@id='abc'].//div": get first div descendant of the div with id = abc
' eg.6 "div": get first div child from the prescribed contextNode (default = 'document')
'---------------------------------------------------------------------------------------

    '1. Verify double slash syntax
    'People normally get this wrong when using XPath
    If InStr(1, strXPath, "\") Then strXPath = Replace(strXPath, "\", "/")

    '2. Xpath方式で検索実行
    Dim SearchJS As String, resultSearch As Variant
    SearchJS = "document.evaluate(""" & strXPath & """, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
    resultSearch = jsEval(SearchJS)

    '3. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '4. 結果を`CDPElement.cls`に格納
    Set getElementByXPath = New CDPElement
    getElementByXPath.Init Me, JsResultVarType, objectId, SearchJS

End Function

Public Function getElementsByXPath(strXPath As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by XPath. For further info on the particular XPath
' syntax used for retrieving the elements, refer to Mozilla doc below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

    '1. Verify double slash syntax
    'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")

    '2. 複数Xpath方式で検索実行
    Dim SearchJS As String, resultsSearch As Variant
    SearchJS = "document.evaluate(""" & strXPath & """, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)"
    resultsSearch = jsEval(SearchJS)

    '3. 結果を判定。文字型なら、ObjectIDを確保
    Dim collOid As String
    If VarType(resultsSearch) = vbString Then
        collOid = resultsSearch
    Else
        'ない場合は、即抜け
        Exit Function
    End If

    '4. Get snapshotLength via callFunctionOn
    Set getElementsByXPath = New Collection
    Dim length_ As Long, N As Long
    length_ = jsEval("function(){ return this.snapshotLength }", objectId:=collOid)

    '5. Get each item's objectId via callFunctionOn
    Dim tElement As CDPElement
    Dim itemArgs(0) As Dictionary, itemArg As Dictionary
    Dim itemResult As Variant, itemObjectID As String, itemResultVarType As VbVarType
    For N = 0 To length_ - 1
        '5-1. `callFunctionOn`に渡す引数を初期化
        Set itemArg = New Dictionary

        '5-2. `callFunctionOn`に渡す引数をセット
        itemArg.Add "value", N
        Set itemArgs(0) = itemArg

        '5-3. 引数付きで実行
        itemResult = jsEval("function(n){ return this.snapshotItem(n) }", objectId:=collOid, objectArguments:=itemArgs)

        '5-4. 念のため、有効な`ObjectID`かチェック
        itemResultVarType = VarType(itemResult)
        If itemResultVarType = vbString Then
            itemObjectID = itemResult

            '5-5. 結果を`CDPElement.cls`に格納し、コレクションに追加
            Set tElement = New CDPElement
            tElement.Init Me, itemResultVarType, itemObjectID
            getElementsByXPath.Add tElement
        End If
    Next

End Function

Public Function getElementByQuery(strQuery As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class by QuerySelector - similar to that of the same in JS.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation. Refer to the below for some examples.
' https://www.w3schools.com/jsref/met_document_queryselector.asp
' https://stackoverflow.com/a/29937844
' https://www.w3.org/TR/selectors-3/#context
' Updated: 08/11/2023
' Examples ------------------------------------------------------------------------
' eg.1 "#abc": get the first element with id attribute = abc
' eg.2 "tb.xyz": get the first element with tag tb and with class xyz
' eg.3 "a[name='abc']": get the first element with tag a and attribute name = abc
' eg.4 "div > p": get the first element with tag 'p' and has parent with tag 'div'
' eg.5 "[a1='abc'][a2='xyz']": get element with attribute a1 = abc and a2 = 123
' eg.6 "[a1='abc'] [a2='xyz']": get element with attribute a2 = xyz and is a child
' of an element with attribute a1 = abc
' eg.7 "#processType [id*='select_option']": get child element with id text that
' contains "select_option" and is a child of element with id "processType"
' eg.8 "#upload:not([disabled])": get the first element with id=upload and does not
' have attribute 'disabled'.
' eg.9 "#upload:disabled": get the first element with id=upload and has attribute
' 'disabled'.
' eg.10 "#id1 tr:last-child [id*='invoiceno']": first get element with id text
' 'id1', then look for the last child with tag 'tr', then finally look within the
' child's children for the first element with id containing 'invoiceno'.
'----------------------------------------------------------------------------------

    '1. querySelector方式で検索実行
    Dim SearchJS As String, resultSearch As Variant
    SearchJS = "document.querySelector(""" & strQuery & """)"
    resultSearch = jsEval(SearchJS)

    '2. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '3. 結果を`CDPElement.cls`に格納
    Set getElementByQuery = New CDPElement
    getElementByQuery.Init Me, JsResultVarType, objectId, SearchJS

End Function

Public Function getElementsByQuery(strQuery As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by QuerySelectorAll. The syntax is as below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

    '1. querySelectorAll方式で検索実行
    Dim SearchJS As String, resultsSearch As Variant
    SearchJS = "document.querySelectorAll(""" & strQuery & """)"
    resultsSearch = jsEval(SearchJS)

    '2. 結果を判定。文字型なら、ObjectIDを確保
    Dim collOid As String
    If VarType(resultsSearch) = vbString Then
        collOid = resultsSearch
    Else
        'ない場合は、即抜け
        Exit Function
    End If

    '3. Get length via callFunctionOn
    Set getElementsByQuery = New Collection
    Dim length_ As Long, N As Long
    length_ = jsEval("function(){ return this.length }", objectId:=collOid)

    '4. Get each item's objectId via callFunctionOn
    Dim tElement As CDPElement
    Dim itemArgs(0) As Dictionary, itemArg As Dictionary
    Dim itemResult As Variant, itemObjectID As String, itemResultVarType As VbVarType
    For N = 0 To length_ - 1
        '4-1. `callFunctionOn`に渡す引数を初期化
        Set itemArg = New Dictionary

        '4-2. `callFunctionOn`に渡す引数をセット
        itemArg.Add "value", N
        Set itemArgs(0) = itemArg

        '4-3. 引数付きで実行
        itemResult = jsEval("function(n){ return this.item(n) }", objectId:=collOid, objectArguments:=itemArgs)

        '4-4. 念のため、有効な`ObjectID`かチェック
        itemResultVarType = VarType(itemResult)
        If itemResultVarType = vbString Then
            itemObjectID = itemResult

            '4-5. 結果を`CDPElement.cls`に格納し、コレクションに追加
            Set tElement = New CDPElement
            tElement.Init Me, itemResultVarType, itemObjectID
            getElementsByQuery.Add tElement
        End If
    Next

End Function

'***************************************************************************************************
'* 機能　　：`getIFrameContextID`メソッドで得た`executionContextId`を基に、`CDPElement`で扱えるようにします
'---------------------------------------------------------------------------------------------------
'* 返り値　：`executionContextId`に基づく`CDPElement`
'* 引数　　：executionContextId     `getIFrameContextID`メソッドで得た`executionContextId`
'***************************************************************************************************
Public Function getIFrame(ExecutionContextId As Long) As CDPElement
    '1. 要素として返却するように`document`を実行
    Dim SearchJS As String, resultSearch As Variant
    SearchJS = "document"
    resultSearch = jsEval(SearchJS, contextId:=ExecutionContextId)

    '2. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '3. 結果を`CDPElement.cls`に格納
    Set getIFrame = New CDPElement
    getIFrame.Init Me, JsResultVarType, objectId, SearchJS
End Function



'***************************************************************************************************
'                               ■■■ その他ヘルパーメソッド ■■■
'***************************************************************************************************
Private Sub SaveBase64ToFile(strB64 As String, FolderPath As String, FileName As String)
'----------------------------------------------------------------------------------
' Decode a base64 string and save to file without external dependencies (WindowsAPI:CryptStringToBinary).
'----------------------------------------------------------------------------------

    'Base64 から、バイト配列に変換
    Dim tmp() As Byte
    tmp = WebCrypto.Decode(strB64, edfBase64)

    'そのバイト配列のまま、保存
    CharConv.BytesToSaveFile tmp, FolderPath, FileName

End Sub



'***************************************************************************************************
'                              ■■■ 複製/継承用プロパティ ■■■
'***************************************************************************************************
'* 機能　　：このClass内で定義してる`CDPCore.cls`を受け取ったり継承させます
'***************************************************************************************************
Property Get InheritanceCDPCore() As CDPCore
    Set InheritanceCDPCore = PipeCore
End Property
Property Set InheritanceCDPCore(Target As CDPCore)
    '1. コア本体を継承
    Set PipeCore = Target

    '2. ログ設定もついでに継承
    CDPHelpers.InheritLogSettings currentLog, Target.LogSettings
End Property

'***************************************************************************************************
'* 機能　　：このClass内で定義してる`CDPBrowser.cls`を受け取ったり継承させます
'***************************************************************************************************
Property Get InheritanceCDPBrowser() As CDPBrowser
    Set InheritanceCDPBrowser = BrowserCore
End Property
Property Set InheritanceCDPBrowser(Target As CDPBrowser)
    Set BrowserCore = Target
End Property



'***************************************************************************************************
'                              ■■■ PUBLIC PROPERTIES ■■■
'***************************************************************************************************
Property Get Url() As String
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    Url = brTab.Url

End Property

Property Let Url(newURL As String)
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    navigate newURL
'    brTab.Url = newURL     ※実際のURL適用は、イベント側で更新させます

End Property

Property Get Title() As String
'------------------------------------------------------------
' Get the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    'JavaScriptで今のタイトルを取得し、更新
    '※`Target.targetInfoChanged`だけでは、最新の`Title`が取れないのでJavaScriptで取ってきます
    brTab.Title = jsEval("document.title")
    Title = brTab.Title

End Property

Property Let Title(newName As String)
'------------------------------------------------------------
' Set the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".Title"


    jsEval "document.title = '" & newName & "'"
    printMsg info_, "Tab name changed to " & newName, FromProcedureName
    brTab.Title = newName    '※JavaScriptによる変更では、`Target.targetInfoChanged`が来ないので、手動で更新

End Property

Property Get CurrentSessionID() As String
'------------------------------------------------------------
' Get the session ID of the current browser
' Added 02/03/2023
'------------------------------------------------------------

    CurrentSessionID = brTab.sessionID

End Property

Property Let CurrentSessionID(newSessionID As String)
'------------------------------------------------------------
' Set the session ID of the current browser to a new value
' Added 02/03/2023
'------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".CurrentSessionID"


    '再設定されるので、元のSessionIDがあれば、破棄させる
    If LenB(brTab.sessionID) Then
        printMsg info_, "A property was called that receives a new SessionID (" & newSessionID & "). The SessionID (" & brTab.sessionID & ") held by this class will now be discarded.", FromProcedureName
        BrowserCore.DiscardSessionID brTab.sessionID
    End If

    '渡されたIDで、再設定
    brTab.sessionID = newSessionID

    'メインタブの場合は、新SessionIDも記録しておく
    With ShSetting01_StartBrowser
        Dim MainTargetID As String, tmp As Dictionary
        Set tmp = ShSetting01_StartBrowser.TableBrowserContext(BrowserCore.CurrentBrowserUser, FromProcedureName)
        If Not tmp Is Nothing Then MainTargetID = tmp("targetID")

        If MainTargetID = brTab.targetID Then .UpdateBrowserContext(BrowserCore.CurrentBrowserUser, "sessionID", FromProcedureName) = brTab.sessionID
    End With

    getSessionInfo        'update all session infos
    printMsg info_, currentLog.LogID & " updated Session ID to " & brTab.sessionID, FromProcedureName

End Property

Property Get html() As String
'------------------------------------------------------------
' Get the html text of the entire document
' Added 02/03/2023
'------------------------------------------------------------

    html = jsEval("document.documentElement.innerHTML")

End Property

'***************************************************************************************************
'* 機能　　：このクラスで定義中のtargetIDを取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：「どのタブか」を特定するための専用ID
'***************************************************************************************************
Property Get CurrentTargetID() As String
    CurrentTargetID = brTab.targetID
End Property

'***************************************************************************************************
'* 機能　　：このクラスで定義中の`browserContextId`を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：「どのProfile」に属してるかを特定するための専用ID
'***************************************************************************************************
Property Get CurrentBrowserContextId() As String
    CurrentBrowserContextId = brTab.browserContextId
End Property

'***************************************************************************************************
'* 機能　　：CDP-Json コマンド結果を得るためのタイムアウト時間を設定します
'---------------------------------------------------------------------------------------------------
'* LET引数 ：置き換えるタイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 機能説明：最初のブラウザ起動時の遷移完了判定や、CDP-Json コマンド結果を得るのに待つ時間を設定します。デフォルトは、30sです。
'***************************************************************************************************
Property Let TimeOutSecond(TimeSec As Double)
    timerOut = TimeSec * 1000       'ミリ秒に補正する
End Property

'***************************************************************************************************
'* 機能　　：このClassで起動したブラウザのウィンドウハンドルを取得します
'---------------------------------------------------------------------------------------------------
'* 引数　　：requestUpdate      `True` :必ずウィンドウハンドルを調査したうえで、ウィンドウハンドルを取得します
'                               `False`:既存の値があればそれを流用し、なければ再調査します
'* 返値    ：ブラウザのウィンドウハンドル
'---------------------------------------------------------------------------------------------------
'* 機能説明：このハンドルを使った高度なウィンドウの調整などができます
'***************************************************************************************************
Property Get BrowserWindowHandle(Optional alwaysRequest As Boolean) As LongPtr
    If alwaysRequest Then
        BrowserWindowHandle = getHandle
    Else
        '既存の値があれば、それを流用、ない場合は、特定処理に入ります
        If brTab.hWnd Then BrowserWindowHandle = brTab.hWnd Else BrowserWindowHandle = getHandle
    End If
End Property

'***************************************************************************************************
'* 機能　　：このClassで起動したブラウザのウィンドウIDを取得します
'---------------------------------------------------------------------------------------------------
'* 引数　　：requestUpdate      `True` :必ずウィンドウIDを調査したうえで、ウィンドウIDを取得します
'                               `False`:既存の値があればそれを流用し、なければ再調査します
'* 返値    ：ブラウザのウィンドウID
'---------------------------------------------------------------------------------------------------
'* 機能説明：このIDを使ったCDP-Jsonとしてのウィンドウ調整("Browser.setWindowBounds"等)ができます
'***************************************************************************************************
Property Get BrowserWindowID(Optional alwaysRequest As Boolean) As Long
    If alwaysRequest Then
        BrowserWindowID = ExecuteCDP("Browser.getWindowForTarget").Item("windowId")
    Else
        '既存の値があれば、それを流用、ない場合は、特定処理に入ります
        If brTab.WID Then BrowserWindowID = brTab.WID Else BrowserWindowID = ExecuteCDP("Browser.getWindowForTarget").Item("windowId")
    End If
End Property

'***************************************************************************************************
'* 機能　　：`targetId`(タブ等)が切断されたか確認します
'---------------------------------------------------------------------------------------------------
'* 返値    ：論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：このClassの`sessionId`に対する`Target.detachedFromTarget`を受け取ると、`True`になります
'***************************************************************************************************
Property Get isTargetDisconnected() As Boolean
    isTargetDisconnected = ControlStatus.detachedFromTarget
End Property

'***************************************************************************************************
'* 機能　　：`targetId`(タブ等)がクラッシュされたか確認します
'---------------------------------------------------------------------------------------------------
'* 返値    ：論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：このClassの`targetId`に対する`Target.targetCrashed`を受け取ると、`True`になります
'***************************************************************************************************
Property Get isTargetCrashed() As Boolean
    isTargetCrashed = ControlStatus.targetCrashed
End Property

'***************************************************************************************************
'* 機能　　：`targetId`(タブ等)が閉鎖されたか確認します
'---------------------------------------------------------------------------------------------------
'* 返値    ：論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：このClassの`targetId`に対する`Target.targetDestroyed`を受け取ると、`True`になります
'***************************************************************************************************
Property Get isTargetClosed() As Boolean
    isTargetClosed = ControlStatus.targetDestroyed
End Property



'***************************************************************************************************
'                           ■■■ CDPHelpers ラッパーメソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub



'***************************************************************************************************
'                               ■■■ CDPCore.cls 直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：メッセージ取り出し/制御系。`RaiseEvent`のコア
'***************************************************************************************************
'--- almighty ---
Private Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    PipeCore.TakeEvents StopPipeError, destruction
End Sub

'--- manual ---
Private Sub TakeEvent()
    PipeCore.TakeEvent
End Sub

Private Property Get NewResCDP() As Boolean
    NewResCDP = PipeCore.NewResCDP
End Property

Private Sub AnalyzeCDP(Optional StopApiError As Boolean = True, Optional destruction As Boolean)
    PipeCore.AnalyzeCDP StopApiError, destruction
End Sub



'***************************************************************************************************
'                               ■■■ CDPBrowser.cls 直通メソッド ■■■
'***************************************************************************************************
Private Function checkURI(URI As String) As String
    checkURI = BrowserCore.checkURI(URI)
End Function



'***************************************************************************************************
'                       ■■■ このClassオブジェクト生成/破棄時のアクション ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------
' A quick way to quickly check if an automated session is already active and open.
' If so, automatically attempt to reattach to it.
' Updated: 07/06/23
'          22/10/25 Daniel Polak - prevent compilation errors because depending on whether we are running in Excel, Word or Access only one of ThisWorkbook, ActiveDocument and CurrentProject is available
'          22/01/26 Long Vh - removed the auto deserialization feature to running multiple instances of browser profiles.
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'Assign an objID for easy debug identification
    currentLog.LogID = "TBID" & Format(Rnd * 1000, "000")

    'public用イベントメソッドフィルターを有効化
    Set publicBrowserEvents.Filter = New Dictionary

    'タイムアウト設定(デフォルトで、30s)
    timerOut = TIMEOUT_DEFAULT * 1000   'ミリ秒に補正

    '`WithEvents`からのコマンド結果を蓄積させる初期化
    Set ResultCommandCDP_Context = New Dictionary

End Sub

'***************************************************************************************************
'* 機能　　：このClassを破棄する際、持ってる`SessionID`も併せて破棄します
'***************************************************************************************************
Private Sub Class_Terminate()
    Const FromProcedureName As String = ThisClassName & ".Class_Terminate"


    '`Target.detachFromTarget`を送るか判定
    If ControlStatus.detachedFromTarget Then
        'もうすでに閉じてるので、わざわざ破棄コマンドは送らない
        printMsg info_, "The session ID (" & brTab.sessionID & ") has been discarded.", FromProcedureName
    Else
        'SessionIDを保持するオプションになってるか確認
        If KeepSession Then
            printMsg Debug_, "The class object will be destroyed, but the `SessionID`(" & brTab.sessionID & ") will be preserved depending on the configuration.", FromProcedureName
        Else
            printMsg Debug_, "When the class object is destroyed, the `SessionID`(" & brTab.sessionID & ") is also destroyed.", FromProcedureName

            'SessionIDを破棄しておく
            BrowserCore.DiscardSessionID brTab.sessionID
        End If
    End If

    'カウントダウンはさせておく
    BrowserCore.EnableDiscoverTargets = False
End Sub
'===============================================================================================
' Experimental Functions
'===============================================================================================
