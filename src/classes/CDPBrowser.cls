VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'                         ブラウザそのものに関するメソッドやパラメータ運用を行います
'                  このClassオブジェクトは、`--user-data-dir`単位で管理する前提とします
'
'   `ExecuteCDP`等で使用できるコマンド一覧は、右記にてアクセスして下さい → https://chromedevtools.github.io/devtools-protocol/
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long



'***************************************************************************************************
'                                   ■■■ 各種構造体 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] タブの種類。命名は右記に準拠します → https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/devtools_agent_host_impl.cc?ss=chromium&q=f:devtools%20-f:out%20%22::kTypeTab%5B%5D%22
'---------------------------------------------------------------------------------------------------'
Public Enum DevToolsAgentHost_KType
    kTab = 1
    kPage                   '普段目にするタブがこれ
    kFrame                  'iFrameと同等
    kDedicatedWorker
    kSharedWorker
    kServiceWorker
    kWorklet
    kBrowser
    kGuest
    kOther                  '非表示タブ
    kAuctionWorklet
    kAssistiveTechnology
    kBrowserUI
End Enum



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'このツール固有情報
Private Const CDPVer                    As String = "3.1.0"
Private Const ArgRemoteDebuggingName    As String = "remote-debugging-pipe"


'その他
Private Const TIMEOUT_DEFAULT   As Double = 30              '初期タイムアウト秒数
Private Const ThisClassName     As String = "CDPBrowser"    'ログのトレース用



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'CDP制御用コアクラスオブジェクト
Private WithEvents PipeCore As CDPCore              'Must be initialized for CDP engine
Attribute PipeCore.VB_VarHelpID = -1

'操作中のブラウザ情報一式
'※`reattach`系から始めた場合は一部の情報は取り出せません
Private Type StateBrowser
    Type        As String   'Browser type: either 'chrome' or 'edge'
    User        As String   'Browser user profile
    Path        As String   'Browser default installation path
    LogID       As String   'The ID of the current CDPBrowser object. Useful when reading debugs of multiple tabs.
    Ver         As String   'The current browser version
    CallString  As String   'The initial call string used by .start
End Type
Private br As StateBrowser

'CDPコマンドの結果を蓄積する変数
Private ResultCommandCDP_Browser As Dictionary  '「ID - それに対するCDPコマンドの結果」というDictionary

'各種エラー情報
Private lastErrorDetailsCDP     As Dictionary  'CDP-Jsonコマンド実行時の最後のエラー情報

'その他
Private timerOut                As Double       'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
Private TargetInfos             As Collection   '非同期イベント`Target.targetCreated`で得た`Target.TargetInfo`情報。基本的に、`CDPContext.cls - attachToSession`専用
Private DefinitionCDPContext    As Long         '確保中の`CDPContext.cls` ※適切なメソッドを呼び出した場合に限る
Private ProtocolWhiteList                       'URL遷移時の許可されたプロトコルリスト(1次元配列)
Private logFolderPath           As String       'ログファイルの保存フォルダパス



'***************************************************************************************************
'                       ■■■ Enum → 文字列 変換プロシージャ ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] タブの種類を文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringDevToolsAgentHost_KType(param As DevToolsAgentHost_KType) As String
    Select Case param
        Case DevToolsAgentHost_KType.kTab:                  EnumToStringDevToolsAgentHost_KType = "tab"
        Case DevToolsAgentHost_KType.kPage:                 EnumToStringDevToolsAgentHost_KType = "page"
        Case DevToolsAgentHost_KType.kFrame:                EnumToStringDevToolsAgentHost_KType = "iframe"
        Case DevToolsAgentHost_KType.kDedicatedWorker:      EnumToStringDevToolsAgentHost_KType = "worker"
        Case DevToolsAgentHost_KType.kSharedWorker:         EnumToStringDevToolsAgentHost_KType = "shared_worker"
        Case DevToolsAgentHost_KType.kServiceWorker:        EnumToStringDevToolsAgentHost_KType = "service_worker"
        Case DevToolsAgentHost_KType.kWorklet:              EnumToStringDevToolsAgentHost_KType = "worklet"
        Case DevToolsAgentHost_KType.kBrowser:              EnumToStringDevToolsAgentHost_KType = "browser"
        Case DevToolsAgentHost_KType.kGuest:                EnumToStringDevToolsAgentHost_KType = "webview"
        Case DevToolsAgentHost_KType.kOther:                EnumToStringDevToolsAgentHost_KType = "other"
        Case DevToolsAgentHost_KType.kAuctionWorklet:       EnumToStringDevToolsAgentHost_KType = "auction_worklet"
        Case DevToolsAgentHost_KType.kAssistiveTechnology:  EnumToStringDevToolsAgentHost_KType = "assistive_technology"
        Case DevToolsAgentHost_KType.kBrowserUI:            EnumToStringDevToolsAgentHost_KType = "browser_ui"
    End Select
End Function



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する所定の起動引数をつけて起動させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：Name           ブラウザ種別（`BrowserList.RunChrome` / `BrowserList.RunEdge`）
'            appUrl         起動引数：`--app`　につけるURLを指定します
'            userProfile    起動引数：`--user-data-dir`　につけるユーザーディレクトリ名を指定します
'            addArgs        追加の起動引数
'---------------------------------------------------------------------------------------------------
'* 機能説明
' This function initiates a new browser instance.
' By default, this function initializes chrome as the default choice.
' To init edge, pass BrowserList.RunEdge to the first parameter (eg. browser.start BrowserList.RunEdge).
' To open in App Mode, input the starting URL for appURL parameter as app mode cannot be open on a blank tab.
' If cleanActive is False, the method will skip cleaning currently open instances. This will save
' initialization time but if an instance is already running and not closed, pipe error will be triggered.
' If reAttach is True, the method will always try to reattach to the previous automation session if it is
' still active. If not, it will then move on to start a brand new session. If False, it will always open a new
' session regardless whether there is an on going one or not.
' To add additional arguments to the call string, use the addArguments parameter for instance:
' .start addArgs:="--disable-info-bar" (https://peter.sh/experiments/chromium-command-line-switches/)
' Note: not all arguments work the same for Chrome and Edge. Some may work for one but not the others.
' Updated: 22/03/23 Long Vh:      - Enhanced to dynamically retrieve installation path of the browser
'          26/03/23 Long Vh:      - Made some quality-of-life changes
'          25/05/25 Long Vh:      - set default profile to "CDP" so that it can work with the latest chrome update
'          23/10/25 Daniel Polak: - Added missing space before --homepage
'          22/01/26 Long Vh:      - Removed automatic reattachment from the start method as this restricts
'                                   the possibility of running multiple CDP browser instances under different
'                                   profiles - which is a very useful feature.
'                                 - Tweaked the addProfile section and the cleanUpSessions to reflect the
'                                   recent Chromium updates and the Edge policy discovery.
'***************************************************************************************************
Public Sub start(Optional Name As BrowserList = BrowserList.RunChrome, _
    Optional appUrl As String, _
    Optional userProfile As String, _
    Optional addArgs As String)
    Const FromProcedureName As String = ThisClassName & ".start"


    '1. パイプ関連の初期化処理しついでに、起動完了までのタイム測定もしちゃう
    Set PipeCore = New CDPCore
    Dim timeStart As Double
    timeStart = CDPHelpers.TimerCounter

    '2. Excelテーブルにユーザー名に紐づくハンドル情報が取れちゃったら、新規起動なのでClose処理しておく
    br.User = userProfile
    If PipeCore.deserialize(br.User) Then PipeCore.CloseHandleBrowser

    '3. 起動依頼
    br.CallString = PipeCore.ReadyInitCDP(br.User, Name, appUrl, addArgs)

    '4. ブラウザセッションパイプ情報をExcelのテーブルに、記録します
    PipeCore.serialize br.User

    '5. 起動中のブラウザ情報を変数に記録
    getBrowserInfo

    '6. ブラウザステータス表示
    printParams
    printMsg info_, "Browser initialization completed after " & (CDPHelpers.TimerCounter - timeStart) / 1000 & "s", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：`targetCreated/targetInfoChanged/targetDestroyed`の3種の非同期イベントを流しつつ、Context のカウントも行います
'---------------------------------------------------------------------------------------------------
'* SET引数 ：論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：・`CDPContext.cls`の生存状況把握に使います
'            ・うまい具合に、`CDPContext.cls`確保中のカウントも行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：`CDPContext.cls`を生成/破棄時に呼ぶ想定です。
'***************************************************************************************************
Property Let EnableDiscoverTargets(Flag As Boolean)
    Const FromProcedureName As String = ThisClassName & ".EnableDiscoverTargets"


    '1. カウント
    DefinitionCDPContext = DefinitionCDPContext + IIf(Flag, 1, -1)

    '2. パラメータ準備
    If DefinitionCDPContext = 1 Or DefinitionCDPContext = 0 Then
        Dim CDPparams As New Dictionary
        CDPparams.Add "discover", Flag
        ExecuteCDP "Target.setDiscoverTargets", CDPparams, False
    End If

    '3. 下限は、0にしておく
    If DefinitionCDPContext < 0 Then DefinitionCDPContext = 0: printMsg WARN_, "`CDPContext.cls` no longer exists.", FromProcedureName

    '4. ログ表示
    printMsg info_, "Currently, " & DefinitionCDPContext & " `CDPContext.cls` file is being defined.", FromProcedureName
End Property

'***************************************************************************************************
'* 機能　　："Target.targetCreated"による非同期イベント収集を有効/無効化します
'---------------------------------------------------------------------------------------------------
'* SET引数 ：格納させたいCollection変数。`Nothing`で渡せば、収集を無効化します
'---------------------------------------------------------------------------------------------------
'* 機能説明：・タブやiframeが開いた際の非同期イベントを流すかの切り替えです。
'            ・参照渡し方式により、わざわざ`Get`プロパティを介さず、取得可能です
'---------------------------------------------------------------------------------------------------
'* 注意事項：用が済んだら必ず、`Nothing`を渡して、参照カウンタを0にします
'***************************************************************************************************
Property Set RunCollect_TargetCreated(ArgCollect As Collection)
    '引数のポインタCollectionを指定させる
    Set TargetInfos = ArgCollect
End Property



'***************************************************************************************************
'                               ■■■ 再接続処理系 ■■■
'***************************************************************************************************
'* 機能　　：Excelテーブルにある既存のパイプハンドル情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：userProfile        再アタッチしたいユーザー名(user-data-dirに基づく)
'---------------------------------------------------------------------------------------------------
'* 注意事項：パイプが生きてない場合は、VBAエラーで止まります
'***************************************************************************************************
Public Sub reattachPipe(userProfile As String)
    Const FromProcedureName As String = ThisClassName & ".reattachPipe"


    '1. 識別名称を設定
    br.User = userProfile

    '2. CDP制御するための初期化
    Set PipeCore = New CDPCore

    '3. 1を基に、Excelテーブルからハンドル情報を取得
    '※載ってない物はエラーとする
    If Not PipeCore.deserialize(br.User) Then printMsg ERROR_, Chr(34) & br.User & """ cannot reattach because there is no pipe handle information in the Excel table.", FromProcedureName, doRaiseError:=True

    '4. 生存確認フェーズへ
    reattach
End Sub

'***************************************************************************************************
'* 機能　　：事前に接続済みの`WebSocket`オブジェクトを渡して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：WebSocket生存論理値
'* 引数　　：userProfile        再アタッチしたいユーザー名(user-data-dirに基づく)
'            WebSocketMode      CDPへ接続済みの`CDPCoreViaWebSocket`オブジェクト
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebSocketオブジェクトが機能してない場合は、VBAエラーで止まります
'***************************************************************************************************
Public Sub reattachWebSocket(userProfile As String, WebSocketMode As CDPCoreViaWebSocket)
    Const FromProcedureName As String = ThisClassName & ".reattachWebSocket"


    '1. 識別名称を設定
    br.User = userProfile

    '2. CDP制御するための初期化
    Set PipeCore = New CDPCore

    '3. 引数を基に、WebSocketの接続情報を適用
    '※オブジェクトなしの場合は、エラーとする
    If WebSocketMode Is Nothing Then printMsg ERROR_, "WebSocket connection information is not configured.", FromProcedureName, doRaiseError:=True
    Set PipeCore.WebSocketMode = WebSocketMode

    '4. 生存確認フェーズへ
    reattach
End Sub

'***************************************************************************************************
'* 機能　　：事前に構成済みの`WebView2`オブジェクトを渡して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：WebSocket生存論理値
'* 引数　　：userProfile        再アタッチしたいユーザー名(user-data-dirに基づく)
'            WebView2Mode       構成済みの`WebView2Mode`オブジェクト
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebView2オブジェクトが機能してない場合は、VBAエラーで止まります
'***************************************************************************************************
Public Sub reattachWebView2(userProfile As String, WebView2Mode As CDPCoreViaWebView2)
    Const FromProcedureName As String = ThisClassName & ".reattachWebView2"


    '1. 識別名称を設定
    br.User = userProfile

    '2. CDP制御するための初期化
    Set PipeCore = New CDPCore

    '3. 引数を基に、WEBView2オブジェクト情報を適用
    '※オブジェクトなしの場合は、エラーとする
    If WebView2Mode Is Nothing Then printMsg ERROR_, "WebView2 connection information is not configured.", FromProcedureName, doRaiseError:=True
    Set PipeCore.WebView2Mode = WebView2Mode

    '4. 生存確認フェーズへ
    reattach
End Sub

'***************************************************************************************************
'* 機能　　：ハンドルが生きれるか確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：生存論理値
'---------------------------------------------------------------------------------------------------
'* 注意事項：ハンドルが機能してない場合は、VBAエラーで止まります
'***************************************************************************************************
Private Sub reattach()
    Const FromProcedureName As String = ThisClassName & ".reattach"


    If isLiveBrowser Then
        '不整合を防ぐため、Pipeに残存しているデータを破棄しておく
        TakeEvents destruction:=True

        'ステータス更新
        getBrowserInfo
        printParams
    Else
        printMsg ERROR_, "Failed to reattach to previous Pipe handle. user-data-dir=" & br.User, FromProcedureName, doRaiseError:=True
    End If

End Sub

'***************************************************************************************************
'* 機能　　：パイプハンドルの生存確認を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：パイプ生存論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：PeekNamedPipe/ioctlsocket のハンドルが生きてるかをチェックします
'***************************************************************************************************
Public Function isLiveBrowser() As Boolean
    isLiveBrowser = PipeCore.isAvailability
End Function



'***************************************************************************************************
'                           ■■■ 各種接続情報記録/削除設定 ■■■
'***************************************************************************************************
'* 機能　　：Excelのテーブルに、メインタブの接続情報を記録します
'---------------------------------------------------------------------------------------------------
'* 引数　　：sessionID_mainTab      メインタブでのSessionId
'            targetID_mainTab       メインタブでのtargetId
'---------------------------------------------------------------------------------------------------
'* 機能説明：
' This method transcribes Chrome's Json message into a string.
' The serialized string can then be used to automate an active
' session with a known serial.
' Updated: 03/05/23 - Replace sessionID with targetID since it is
'                     constant and stable on every CDP session
'          22/01/26 - Added specification to user profile name to
'                     the cpid file name to allow targetting of
'                     specific profile for reattachment.
'***************************************************************************************************
Public Sub serializeForMainTab(sessionID_mainTab As String, targetID_mainTab As String)
    Const FromProcedureName As String = ThisClassName & ".serialize"


    '1. 必要なデータを`Dictionary`に詰める
    Dim BrowserInfo As New Dictionary
    BrowserInfo.Add "BiDi-context", vbNullString  '`WebDriverBiDi`用に枠だけ確保
    BrowserInfo.Add "sessionID", sessionID_mainTab
    BrowserInfo.Add "targetID", targetID_mainTab

    '2. Excelのテーブルへ記録する
    Set ShSetting01_StartBrowser.TableBrowserContext(br.User, FromProcedureName) = BrowserInfo
End Sub

'***************************************************************************************************
'* 機能　　：Excelのテーブルから、ブラウザへの接続パイプとタブ情報を削除します
'***************************************************************************************************
Private Sub deleteCDPID()
    Const FromProcedureName As String = ThisClassName & ".deleteCDPID"


    With ShSetting01_StartBrowser
        .DeleteBrowserHandle(FromProcedureName) = br.User   'パイプ通信Handle情報
        .DeleteBrowserContext(FromProcedureName) = br.User  'ブラウザセッション情報
    End With
End Sub



'***************************************************************************************************
'                              ■■■ タブ操作関連 ■■■
'***************************************************************************************************
Public Function newTab(Optional Url As String, Optional newWindow As Boolean, Optional setMain As Boolean, Optional isHidden As Boolean, Optional browserContextId As String, Optional isBackground As Boolean) As CDPContext
'------------------------------------------------------------------
' Create a new blank tab within the same instance.
' url: The initial URL the page will be navigated to. An empty
' string indicates about:blank.
' newWindow: If true will open a separate window instead of a tab.
' setMain: If true will set the new tab as the main session tab.
' isBackground: If true, creates the new target in background.
' In the original CDP document, each new tab is called "Target" and
' the equivalence of this method is Target.createTarget
' This returns the Session ID of the newly created tab.
' Updated: 20/03/23 Return the result as a new Browser object
'          26/04/23 Format url argument to accept simplified input
'          22/01/26 Added .reattach where appropriate
'          07/03/26 Added hidden argument
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".newTab"


    '1. パラメーター準備
    Dim params As New Dictionary
    params("url") = checkURI(Url)
    If newWindow Then params("newWindow") = newWindow
    If isHidden Then params("hidden") = isHidden
    If LenB(browserContextId) Then params("browserContextId") = browserContextId
    If isBackground Then params("background") = isBackground

    '2. 実行
    Dim results As BiDiCDPJson
    Set results = ExecuteCDP("Target.createTarget", params)

    '3. Return the new tab's Session ID
    Dim tabId As String, sessionID_ As String
    tabId = results.StringKey("targetId")
    sessionID_ = attachToTab(tabId)

    '4. メイン化する場合は、Excelに記録しておく
    If setMain Then serializeForMainTab sessionID_, tabId


    '------ 5. 専用のクラスとして返却する準備 ------
    'Switch attachment to the new tab
    '22/01/26: From v3.1, reattachment is no longer automatic so .reattach is needed
    printMsg info_, "newTab will now open a new CDP instance", FromProcedureName
    EnableDiscoverTargets = True    '`CDPContext.cls`を生成するので、カウントアップ
    Set newTab = New CDPContext
    With newTab
        '1. このブラウザのパイプハンドル情報を渡す
        Set .InheritanceCDPCore = PipeCore

        '2. このブラウザ情報を渡す
        Set .InheritanceCDPBrowser = Me

        '3. 先ほど生成した`sessionId`を渡す
        .CurrentSessionID = sessionID_

        '4. Urlが設定されてる場合は、読み込み完了まで待機する
        If LenB(Url) Then .wait
    End With
End Function

'***************************************************************************************************
'* 機能　　：既に開いているタブのセッションIDを、タブの名前またはURLに基づいて検索します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - CDPBrowser
'* 引数　　：tabName        タブ名  第一優先
'            Url            URL     第二優先
'            setMain        `True`で、Excelのタブ情報テーブルに上書きします。タブのリアタッチに利用します
'            SearchType     kPage,kFrame といった、タブの種類を指定します
'            doRetrySecond  指定秒以内に出るまでリトライを繰り返します。0で、1回のみです
'---------------------------------------------------------------------------------------------------
'* 更新履歴：06/06/23               - Added doRetry argument
'            03/05/23               - Combined the method with the getTabNew
'            24/10/25 Daniel Polak  - Predictable, precise, and script-friendly tab selection that scales well when there are many similar tabs open
'                                     Exact title and URL-prefix matching avoids accidentally selecting tabs with similar titles/URLs
'            23/10/25 Daniel Polak  - Fixed typo timeStart instead of timerStart and fixed bug 60000s instead of 60s
'            01/08/26 Eschamali     - iframeへの切り替えができるように改良
'            22/01/26 Long Vh       - Added .reattach where appropriate
'---------------------------------------------------------------------------------------------------
'* 機能説明：tabName引数とurl引数は部分的な指定(部分一致)が可能です。SearchType はデフォルトは、"page"にしてます。
'* 注意事項：これらの引数を両方とも省略すると、getTabは最も近い未接続のタブに接続しようとします。
'***************************************************************************************************
Public Function getTab(Optional tabName As String, Optional Url As String, Optional setMain As Boolean, Optional SearchTypeID As DevToolsAgentHost_KType = kPage, Optional doRetrySecond As Double) As CDPContext
    Const FromProcedureName As String = ThisClassName & ".getTab"


    Dim NodeToken As Long
    Dim tabId As String, candidateTabId As String
    Dim targets As BiDiCDPJson
    Dim timerStart As Double
    Dim tabfound As Boolean
    Dim bestScore As Long, score As Long
    Dim info As String


    '1. タブTypeを決める
    Dim SearchType As String: SearchType = EnumToStringDevToolsAgentHost_KType(SearchTypeID)

    '2. タイマースタートして探索を開始する
    If doRetrySecond > 0 Then timerStart = CDPHelpers.TimerCounter
    Do
        '3. 実行して、`Target.TargetInfo`オブジェクトリストを取得します
        Set targets = ExecuteCDP("Target.getTargets").NodeKey("targetInfos")
        NodeToken = targets.FirstChildToken  ' 配列の先頭トークン

        '4. スコア方式を利用して、理想の対象タブ情報を取得します
        Const MaxScore As Long = 100
        bestScore = -1
        candidateTabId = vbNullString
        Do While NodeToken > 0
            '4-1. タブの種類があってるか？　※未指定の場合は、タブの種類を問わない検索で
            If targets.TokenString(NodeToken, "type") = SearchType Or StrPtr(SearchType) = 0 Then
                '4-2. 現スコアをリセット
                score = -1

                '4-3. 入力した引数に応じた確認
                Select Case True
                    '第一優先：タブ名
                    Case StrPtr(tabName)
                        Dim TargetInfo_title As String: TargetInfo_title = targets.TokenString(NodeToken, "title")

                        '完全一致
                        If TargetInfo_title = tabName Then
                            score = 100
                        '部分一致
                        ElseIf InStr(1, TargetInfo_title, tabName) Then
                            score = 50
                        '一致せず
                        Else
                            printMsg WARN_, "with name containing """ & tabName & """", FromProcedureName
                        End If

                    '第二優先：URL
                    Case StrPtr(Url)
                        Dim TargetInfo_url As String: TargetInfo_url = targets.TokenString(NodeToken, "url")

                        '一致位置検査
                        Dim StartPos As Long: StartPos = InStr(1, TargetInfo_url, Url)

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

                    '第三優先：直近の未接続タブ
                    Case Else
                        If targets.TokenBool(NodeToken, "attached") = False Then
                            score = 80
                        Else
                            score = 30: printMsg Debug_, "as no new unattached tab could be found", FromProcedureName
                        End If

                End Select

                '4-4. 記録中のベストスコアを上回ったら更新
                If score > bestScore Then
                    bestScore = score
                    candidateTabId = targets.TokenString(NodeToken, "targetId")
                End If

                '4-5. 最大スコアが出てるなら、捜査打ち切り
                If bestScore = MaxScore Then Exit Do
            End If

            '4-6. 次の`Target.TargetInfo`へ
            NodeToken = targets.NextToken(NodeToken)
        Loop

        '5. 目的の物っぽいのがあったら見つけたとします
        If bestScore > 0 Then
            tabfound = True
            tabId = candidateTabId
        End If

    '6. Do retry as long as not timeout
    Loop Until tabfound Or CDPHelpers.TimerCounter - timerStart > doRetrySecond * 1000

    '7. Returns nothing if not found
    'Nothing is preferred over throwing error
    If Not tabfound Then
        info = IIf(LenB(tabName), tabName, IIf(LenB(Url), Url, "(no new tab found)"))
        printMsg WARN_, "Unable to find the tab: " & info, FromProcedureName
        Exit Function
    Else
        printMsg info_, "Found the tab with ID: " & tabId & ". (Certainty：" & bestScore & "%)", FromProcedureName
    End If

    '8.Return the session object for reference
    Dim sessionID_ As String
    sessionID_ = attachToTab(tabId)

    '9. メイン化する場合は、Excelに記録しておく
    If setMain Then serializeForMainTab sessionID_, tabId


    '------ 10.専用のクラスとして返却する準備 ------
    '22/01/26 .reattach br.User is needed from v3.1 onwards to enable multi-instances automation
    EnableDiscoverTargets = True    '`CDPContext.cls`を生成するので、カウントアップ
    Set getTab = New CDPContext
    With getTab
        '1. このブラウザのパイプハンドル情報を渡す
        Set .InheritanceCDPCore = PipeCore

        '2. このブラウザ情報を渡す
        Set .InheritanceCDPBrowser = Me

        '3. 先ほど生成した`sessionId`を渡す
        .CurrentSessionID = sessionID_
    End With
End Function

'***************************************************************************************************
'* 機能　　：`targetId`に対する`DevTools`ウィンドウを開きます
'---------------------------------------------------------------------------------------------------
'* 引数　　：targetId   対象のタブID
'            panelId    最初に開きたい項目があれば指定
'---------------------------------------------------------------------------------------------------
'* 機能説明：F12が効かないページでも`DevTools`ウィンドウを開かせます
'* 注意事項：WebView2,Electronの場合は、効かない場合があります
'***************************************************************************************************
Public Sub openDevTools(TrgetID As String, Optional panelId As String)
    'パラメーター準備
    Dim params As New Dictionary
    params.Add "targetId", TrgetID
    If StrPtr(panelId) Then params.Add "panelId", panelId

    '実行
    ExecuteCDP "Target.openDevTools", params
End Sub

'***************************************************************************************************
'* 機能　　：現時点で開いている全てのターゲット情報をログに表示します
'***************************************************************************************************
Public Sub printTargetInfos()
    Const FromProcedureName As String = ThisClassName & ".printTargetInfos"


    Dim NodeToken As Long
    Dim TabCount As Long
    Dim targets As BiDiCDPJson
    Set targets = ExecuteCDP("Target.getTargets").NodeKey("targetInfos")

   'Read info
    printMsg info_, "---------------------------------------------------", FromProcedureName
    printMsg info_, "targetInfos Information", FromProcedureName
    printMsg info_, "---------------------------------------------------", FromProcedureName

    With targets
        NodeToken = .FirstChildToken  ' 配列の先頭トークン
        Do While NodeToken > 0
            TabCount = TabCount + 1

            printMsg info_, "#" & TabCount & vbTab & _
                        "type: " & .TokenString(NodeToken, "type") & " | " & _
                        "attached: " & .TokenString(NodeToken, "attached") & " | " & _
                        "title: " & .TokenString(NodeToken, "title") & " | " & _
                        "url: " & .TokenString(NodeToken, "url") & " | " & _
                        "target ID: " & .TokenString(NodeToken, "targetId"), FromProcedureName

            '次の`Target.TargetInfo`へ
            NodeToken = .NextToken(NodeToken)
        Loop
    End With
End Sub

Public Function PageCount() As Long
'------------------------------------------------------------------------------
' Print all current Pages info. Useful for debugging.
' Updated : 05/02/25 - changed from printTabs to PageCount to get the count too
'------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".PageCount"


    Dim NodeToken As Long
    Dim targets As BiDiCDPJson
    Set targets = ExecuteCDP("Target.getTargets").NodeKey("targetInfos")

   'Read info
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, "Pages Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    With targets
        NodeToken = .FirstChildToken  ' 配列の先頭トークン
        Do While NodeToken > 0
            If .TokenString(NodeToken, "type") = "page" Then
                PageCount = PageCount + 1

                printMsg Debug_, "#" & PageCount & vbTab & _
                            "type: " & .TokenString(NodeToken, "type") & " | " & _
                            "attached: " & .TokenString(NodeToken, "attached") & " | " & _
                            "title: " & .TokenString(NodeToken, "title") & " | " & _
                            "url: " & .TokenString(NodeToken, "url") & " | " & _
                            "target ID: " & .TokenString(NodeToken, "targetId"), FromProcedureName
    
            End If

            '次の`Target.TargetInfo`へ
            NodeToken = .NextToken(NodeToken)
        Loop
    End With
End Function



'***************************************************************************************************
'                              ■■■ CDPcommand操作関連 ■■■
'***************************************************************************************************
'* 機能　　：CDPのMethod名と、parametersを渡して、ブラウザへ命令を行い、結果の回収も同時に行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド の実行結果`result`ノード情報。失敗時は、`Nothing`となります。
'* 引数　　：methodName             method名。"Network.getCookies"や"DOM.getDocument"というのを記入します
'            params                 Json-Dictionary形式で、KeyとItemを事前に定義したオブジェクト変数を指定します。
'            StopCDPError           `True`で、実行失敗時に`Err.Raise`で停止し、以降の処理を止めます
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
    If PipeCore.ShowLogCDPresponse Then printMsg Debug_, "ExecuteCDP `" & methodName & "` with result: " & ResultJsonStr, FromProcedureName

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
    If Not params Is Nothing Then CDPcommand.Add "params", params   '`nothing`の場合は、`params`項目を省略

    '2. ブラウザへ送信し、実行IDを取得
    ExecuteCDPAsync = PipeCore.ReadyRunCDP(CDPcommand)  'Send message to the browser instance

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
'* 機能　　：ブラウザに対するCDPレスポンス結果を一旦、Dictionaryへ蓄積させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：id             コマンドID
'            RawJson        コマンド結果の生Json文字列　※As Is. ID付きのままで渡されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：`ID`をKeyに、Dictionaryへ`RawJson`を格納します。
'* 注意事項：Jsonは、文字列のままで渡されます。
'***************************************************************************************************
Private Sub PipeCore_CDPBrowserID(id As Long, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPBrowserID"


    '結果を蓄積し、後で取り出せるようにする
    ResultCommandCDP_Browser.Add id, RawJson
    printMsg Trace_, "The results of the CDP(ID: " & id & ") execution from the browser have been received. → " & RawJson, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：オーバーフロー対策によるcommandIDがリセットされると呼び出されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：IDが20億に達すると、結果履歴が全て削除されます
'***************************************************************************************************
Private Sub PipeCore_ResetCommandID()
    ResultCommandCDP_Browser.RemoveAll
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
    If ResultCommandCDP_Browser.Exists(commandID) Then
        '2. Log化
        Dim tmp As String: tmp = ResultCommandCDP_Browser(commandID)
        printMsg info_, "Browser were able to confirm the existence of CDP command ID:" & commandID & "　→　" & tmp, FromProcedureName

        '3. 返却
        TakeResultCDP = tmp

        '4. 取り出したので消しておく
        ResultCommandCDP_Browser.Remove commandID
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
    CleanResultCommandCDP = ResultCommandCDP_Browser.Count

    '2. ALLクリア
    ResultCommandCDP_Browser.RemoveAll
    printMsg info_, "The user cleared " & CleanResultCommandCDP & " previous CDP command results.", FromProcedureName
End Function



'***************************************************************************************************
'                                       ■■■ 後始末系 ■■■
'***************************************************************************************************
Public Sub quit()
'-----------------------------------------------------------------------------------
' Closes the currently attached browser instance
' Original Name: closeBrowser
' Update: 24/10/2025: Daniel Polak - avoid immediate hard kills, reduce race conditions, and make shutdown behavior predictable and robust in non-responsive cases
'-----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".quit"


    Dim sngT As Double

    '1. 終了コマンドを送る
    '※閉じると同時に、Pipeエラーが起こることがあるので、エラー無視しておく
    ExecuteCDP "Browser.close", StopCDPError:=False

    '2. パイプ生存確認
    Dim PipeExists As Boolean
    If PipeCore.LastErrorPeekNamedPipe > 0 Or PipeCore.LastErrorReadFile > 0 Or PipeCore.LastErrorWriteFile > 0 Then
        '2-1. もうパイプは閉じてる
        PipeExists = False
    Else
        '2-1. どうもブラウザは生きてるようなのでコマンドを送って確認する
        sngT = CDPHelpers.TimerCounter
        Do While PipeExists And CDPHelpers.TimerCounter - sngT < 5000
            PipeExists = isLiveBrowser

            '2-2. なくなってたら、即抜け
            If Not PipeExists Then Exit Do
            CDPHelpers.Sleep 0.1
        Loop
    End If

    '3. 一定時間たっても閉じない場合は、応答なしとし、プロセス終了をかけて、無理やり閉じる
    If PipeExists Then
        printMsg WARN_, "Browser didn't close gracefully within 5s; invoking cleanUpSessions.", FromProcedureName
        TerminateDebugBrowser
    End If

    '4. パイプハンドル、ブラウザセッション情報を消去しに行きます
    PipeCore.CloseHandleBrowser '各種ハンドルを閉じる
    deleteCDPID                 'テーブルからも削除する

    '5. it seems without waiting a bit the browser crashes and the next time it wants to recover from a crash
    CDPHelpers.Sleep 1
    printMsg info_, "The browser has been terminated.", FromProcedureName

End Sub



'***************************************************************************************************
'* 機能　　：`CDPCore.InitCDP`で起動したデバッグブラウザプロセスハンドルに対して、プロセス終了をかけます
'---------------------------------------------------------------------------------------------------
'* 注意事項：CDPコマンド`Browser.close`を実行したのに、閉じれないアクシデント時に使用して下さい
'***************************************************************************************************
Private Sub TerminateDebugBrowser()
    TerminateProcess PipeCore.Connection_hProcess, 0
End Sub



'***************************************************************************************************
'                                ■■■ 非同期イベント関連 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する非同期イベントを処理します
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName 非同期イベントメソッド名
'            RawJson    非同期イベントの生Json文字列　※As Is. メソッド名付きのままで渡されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：`Target.targetCreated`に対する情報の蓄積を行います
'***************************************************************************************************
Private Sub PipeCore_CDPBrowserEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPBrowserEvent"


    '1. Select Case でスマートに、method名ごとの独自処理を行う
    Select Case methodName
        '1-1. `attachToSession`が行えるように準備する
        Case "Target.targetCreated"
            '1-2. `TargetInfo`を収集しない場合は、即抜け
            If TargetInfos Is Nothing Then Exit Sub

            '1-3. そのまま、蓄積
            printMsg Debug_, "Collecting `Target.targetCreated`... → " & RawJson, FromProcedureName
            TargetInfos.Add RawJson

        Case Else
            'None

    End Select
End Sub



'***************************************************************************************************
'                                   ■■■ ステータス系 ■■■
'***************************************************************************************************
Public Sub printParams()
'------------------------------------------------------------------
' Print all parameters of the current session for debuggings.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".printParams"


    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, br.LogID & " Browser - Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "Browser Type      | brType       | " & br.Type, FromProcedureName
    printMsg Debug_, "Browser Ver       | brVer        | " & br.Ver, FromProcedureName
    printMsg Debug_, "Browser Path      | brPath       | " & br.Path, FromProcedureName
    printMsg Debug_, "Core StdInWr      | hStdInWr     | " & PipeCore.ConnectionPipe_hStdInWr, FromProcedureName
    printMsg Debug_, "Core StdOutRd     | hStdOutRd    | " & PipeCore.ConnectionPipe_hStdOutRd, FromProcedureName
    printMsg Debug_, "Core hStderrOutRd | hStderrOutRd | " & PipeCore.ConnectionPipe_hStderrOutRd, FromProcedureName
    printMsg Debug_, "Core CDPInWr      | hCDPInWr     | " & PipeCore.ConnectionPipe_hCDPInWr, FromProcedureName
    printMsg Debug_, "Core CDPOutRd     | hCDPOutRd    | " & PipeCore.ConnectionPipe_hCDPOutRd, FromProcedureName
    printMsg Debug_, "Core dwProcessId  | dwProcessId  | " & PipeCore.Connection_dwProcessId, FromProcedureName
    printMsg Debug_, "CDP Version       | CDPVer       | " & CDPVer, FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：`Browser.getVersion`によるブラウザ情報を取得します
'***************************************************************************************************
Private Sub getBrowserInfo()
    br.Type = ExecuteCDP("Browser.getVersion")("product")
    br.Ver = Split(br.Type, "/")(1)
    If PipeCore.Connection_dwProcessId > 0 Then br.Path = getBrowserPath(br.Type)  'プロセスID情報がない場合は、取らない
End Sub



'***************************************************************************************************
'                             ■■■ SessionID発行と破棄 ■■■
'***************************************************************************************************
'* 機能　　：SessionID の発行を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：`TargetID`に対する`sessionId`  ※失敗時は、`vbnullstring`となります
'* 引数　　：tabId  TargetID
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Attach to a specific tab with its tabId (a.k.a targetId)
'            flatten: Enables "flat" access to the session via specifying sessionId attribute
'            in the commands. Set to True by default to make attachToTab work.
'            Returns: a session ID string that is needed for CDP automation.
'***************************************************************************************************
Public Function attachToTab(tabId As String) As String
    '1. コマンド準備
    Dim CDPparams As New Dictionary
    CDPparams.Add "targetId", tabId
    CDPparams.Add "flatten", True

    '2. 実行
    Dim tmp As BiDiCDPJson
    Set tmp = ExecuteCDP("Target.attachToTarget", CDPparams, False)

    '3. 結果があれば、返却
    If Not tmp Is Nothing Then attachToTab = tmp.StringKey("sessionId")
End Function

'***************************************************************************************************
'* 機能　　：SessionID の破棄を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：破棄成功論理値
'* 引数　　：SessionID      破棄したいSessionID
'---------------------------------------------------------------------------------------------------
'* 詳細説明：不必要なSessionIDを破棄することで、ブラウザの安定性が向上し、奇妙な不具合も消えるかもしれません
'* 注意事項：・ブラウザ権限で実行するため、メイン化してる`SessionID`も破棄すると、処理が止まるため、使う箇所は慎重に...
'            ・念のため成功論理値を返してますが基本、確認するほどでもないと思います
'***************************************************************************************************
Public Function DiscardSessionID(sessionID As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".DiscardSessionID"


    '1. コマンド準備
    Dim CDPparams As New Dictionary
    CDPparams.Add "sessionId", sessionID

    '2. 破棄コマンド実行
    Dim tmp As BiDiCDPJson
    Set tmp = ExecuteCDP("Target.detachFromTarget", CDPparams, False)

    '3. ログとして出しておく
    If tmp Is Nothing Then
        printMsg Debug_, "Failed to destroy SessionID: " & sessionID, FromProcedureName
    Else
        printMsg info_, "Successfully destroyed the SessionID: " & sessionID, FromProcedureName
    End If
End Function



'***************************************************************************************************
'                               ■■■ その他ヘルパーメソッド ■■■
'***************************************************************************************************
'* 機能　　：URIのプロトコルチェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数の内容そのまんま
'* 引数　　：URI   ブラウザのURLに渡す予定の文字列
'---------------------------------------------------------------------------------------------------
'* 注意事項：・セキュリティ重視により、ホワイトリスト外のプロトコルが来たら、エラーで止まります
'            ・既に開いてるURIに関して(`getTab`メソッド)はチェック対象外とします
'***************************************************************************************************
Public Function checkURI(URI As String) As String
    Const FromProcedureName As String = ThisClassName & ".checkURI"


    '0. 空文字の場合は、チェックしません
    printMsg Debug_, "Checking the protocol of the URI...", FromProcedureName
    If LenB(URI) = 0 Then
        printMsg Debug_, "Checking is skipped because the URI is empty.", FromProcedureName

        checkURI = URI
        Exit Function
    End If

    '1. 先頭始まりかチェック
    Dim i As Long
    For i = LBound(ProtocolWhiteList) To UBound(ProtocolWhiteList)
        '先頭始まりなら、引数の内容そのまんまを返して抜ける
        If Left$(URI, Len(ProtocolWhiteList(i))) = ProtocolWhiteList(i) Then
            printMsg Debug_, "Protocol check passed：" & ProtocolWhiteList(i), FromProcedureName

            checkURI = URI
            Exit Function
        End If
    Next

    '2. プロトコル名をチェック
    Dim ProtocolName, ErrorMes As String
    ProtocolName = Split(URI, ":")
    If UBound(ProtocolName) = 0 Then
        ErrorMes = "URLの形式が正しくありません。" & vbCrLf & "安全のため省略せず、`https://`や`file://`のような、プロトコル名を、URLの先頭に付けてください。" & vbCrLf & vbCrLf & "[問題のURL]" & vbCrLf & URI
    Else
        ErrorMes = "プロトコル`" & ProtocolName(0) & "`は、ホワイトリストに登録されていません。" & vbCrLf & "意図しないアプリケーションの起動などを防ぐため、処理を中断しました。" & vbCrLf & vbCrLf & "【解決策】" & vbCrLf & "このプロトコルへのアクセスが、本当に必要である場合は、シート「ブラウザ起動設定」のテーブル「ProtocolWhiteList」に、`" & ProtocolName(0) & "`を追加してから、もう一度、お試しください。"
    End If

    '4. エラーで中断
    printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True

End Function



'***************************************************************************************************
'                              ■■■ 複製用プロパティ ■■■
'***************************************************************************************************
'* 機能　　：このClass内で定義してる`CDPCore.cls`を受け取ります
'***************************************************************************************************
Property Get InheritanceCDPCore() As CDPCore
    Set InheritanceCDPCore = PipeCore
End Property



'***************************************************************************************************
'                              ■■■ PUBLIC PROPERTIES ■■■
'***************************************************************************************************
'* 機能　　：このClass内で起動してるブラウザパス値を受け取ります
'***************************************************************************************************
Property Get CurrentBrowserPath() As String
    CurrentBrowserPath = br.Path
End Property

'***************************************************************************************************
'* 機能　　：このClass内で起動してるブラウザの`user-data-dir`値を受け取ります
'***************************************************************************************************
Property Get CurrentBrowserUser() As String
    CurrentBrowserUser = br.User
End Property

'***************************************************************************************************
'* 機能　　：このClass内で起動してるブラウザのWindowクラス値を受け取ります
'***************************************************************************************************
Property Get CurrentBrowserWindowClass() As String
    CurrentBrowserWindowClass = chromeWindowClass
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
'                           ■■■ CDPHelpers ラッパーメソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, br.LogID, isHeader, doRaiseError, logFolderPath
End Sub



'***************************************************************************************************
'                               ■■■ CDPCore.cls 直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：メッセージ取り出し/制御系。`RaiseEvent`のコア
'***************************************************************************************************
'--- almighty ---
Public Sub TakeEvents(Optional StopApiError As Boolean = True, Optional destruction As Boolean)
    PipeCore.TakeEvents StopApiError, destruction
End Sub

'--- manual ---
Public Sub TakeEvent()
    PipeCore.TakeEvent
End Sub

Property Get NewResCDP() As Boolean
    NewResCDP = PipeCore.NewResCDP
End Property

Public Sub AnalyzeCDP(Optional StopApiError As Boolean = True, Optional destruction As Boolean)
    PipeCore.AnalyzeCDP StopApiError, destruction
End Sub



'***************************************************************************************************
'                         ■■■ このClassオブジェクト生成時のアクション ■■■
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
    br.LogID = "BRID" & Format(Rnd * 1000, "000")


    '------------------------- クラス内変数の一部を初期化 -------------------------
    'ホワイトプロトコルリストを取得
    ProtocolWhiteList = ShSetting01_StartBrowser.TableProtocolWhiteList

    'ログファイル保存先
    logFolderPath = ShSetting01_StartBrowser.UseRangeID(9, FromProcedureName)

    'タイムアウト設定(デフォルトで、30s)
    timerOut = TIMEOUT_DEFAULT * 1000   'ミリ秒に補正

    '`WithEvents`からのコマンド結果を蓄積させる初期化
    Set ResultCommandCDP_Browser = New Dictionary
    '------------------------------------------------------------------------------
End Sub



#If DEBUG_MODE Then
'***************************************************************************************************
'                                 ■■■ Test Only Area ■■■
'***************************************************************************************************
' 単体テストをしやすくするための専用テストメソッド。
' 事前に「条件付きコンパイル引数」にて、`DEBUG_MODE = 1`をすることで有効化される
'***************************************************************************************************
Property Set InheritanceCDPCore(t As CDPCore)
    Set PipeCore = t
End Property
#End If
