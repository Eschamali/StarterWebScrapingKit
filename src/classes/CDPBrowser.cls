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
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'このツール固有情報
Private Const ToolVer           As String = "4.2.0"

'その他
Private Const ThisClassName     As String = "CDPBrowser"    'ログのトレース用



'***************************************************************************************************
'                                 ■■■ 拡張用イベント定義 ■■■
'***************************************************************************************************
Public Event ExEvent(methodName As String, RawJson As String)    'ブラウザに対する非同期イベント(拡張用)



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'CDP制御用コアクラスオブジェクト
Private WithEvents PipeCore As CDPCore              'Must be initialized for CDP engine
Attribute PipeCore.VB_VarHelpID = -1

'操作中のブラウザ情報一式
'※Browser.getVersion で得れる情報とほぼ同じ構造です
Private Type StateBrowser
    'オリジナル
    JsVersion           As String   'JavaScriptV8エンジンのバージョン番号
    Product             As String   'このChromiumの製品名とバージョン番号
    ProtocolVersion     As String   'CDP自体のバージョン
    Revision            As String   'このChromiumのコミットID
    UserAgent           As String   'このChromiumのユーザーエージェント

    '独自
    Ver         As String   'The current browser version
End Type
Private br As StateBrowser

'DeveloperToolsAvailabilityに関する情報
Private Type DeveloperToolsAvailability
    Allowlist   As Variant
    Blocklist   As Variant
    Fallback    As Variant
    Collected   As Boolean  'ポリシー情報を取得済みか？ ※起動済みブラウザへの接続時等、`OpenBrowser`を通らない場合はFalse
End Type
Private CurrentDeveloperToolsAvailability As DeveloperToolsAvailability

'`DeveloperToolsAvailabilityAllowlist/Blocklist`のURLパターン1件分の解析結果
'※Chromium`url_matcher::util::FilterToComponents`の簡易版です
Private Type UrlPatternInfo
    Pattern         As String   '元のパターン文字列
    Allow           As Boolean  'Allowlist側のパターンか？
    Scheme          As String   'スキーム                 ※空ならすべてのスキーム
    Host            As String   'ホスト                   ※サブドメイン一致の場合は先頭に"."。"*"の場合は空
    MatchSubdomains As Boolean  'サブドメインも一致させるか？ ※先頭"."指定やIPアドレスの場合はFalse
    port            As Long     'ポート番号               ※0ならすべてのポート
    Path            As String   'パス(前方一致)           ※空ならすべてのパス
    Query           As String   'クエリ条件               ※空なら条件なし
    QueryCount      As Long     'クエリ条件の数
    IsValid         As Boolean  '解析できたか？(一致検索の結果としては、一致したか？)
End Type

'CDPコマンドの結果を蓄積する変数
Private ResultCommandCDP_Browser As Dictionary  '「ID - それに対するCDPコマンドの結果」というDictionary

'各種エラー情報
Private lastErrorDetailsCDP     As Dictionary  'CDP-Jsonコマンド実行時の最後のエラー情報

'その他
Private timerOut                As Double       'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
Private TargetInfos             As Collection   '非同期イベント`Target.targetCreated`で得た`Target.TargetInfo`情報。基本的に、`CDPContext.cls - attachToSession`専用
Private DefinitionCDPContext    As Long         '確保中の`CDPContext.cls` ※適切なメソッドを呼び出した場合に限る
Private m_LogID                 As String       'ログ識別ID
Private CustomUAMode            As Boolean      'このツール独自のユーザーエージェントを刻印するか？



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する所定の起動引数をつけて起動させます
'---------------------------------------------------------------------------------------------------
'* 返り値　：起動時の初期URL
'* 引数　　：caps    ブラウザ起動時のオプション設定オブジェクト。ワンショットで、最終的に`Nothing`になります
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
Public Function OpenBrowser(caps As BrowserCapabilities) As String
    '1. プロトコルコアオブジェクトを受けとる
    Set PipeCore = caps.ClsGet_CDPCore

    '2. 起動後に使う設定値を確保
    Dim UserName   As String: UserName = caps.Set_UserDataDirName
    Dim PolicyList As CDPHost: Set PolicyList = caps.ClsGet_CDPHost
    CustomUAMode = caps.Set_EnableUserAgentOverride

    '3. 起動モードに応じた分岐
    If caps.Set_UseWebSocket Then
        '3-1. WebSocketとして起動
        '※内部で､ポリシー情報も記録します
        Dim WebSocketCore As CDPCoreViaWebSocket: Set WebSocketCore = New CDPCoreViaWebSocket
        OpenBrowser = WebSocketCore.OpenBrowserAndConnectCDP(caps)

        '3-2. WebSocketモードとして切り替え
        '※内部で、接続情報(WinSockハンドル等)パラメーターが適用/記録されます
        Set PipeCore.RemoteDebuggingPortMode(UserName) = WebSocketCore
    Else
        '3-1. Pipeとして起動
        '※内部で、接続情報(Pipeハンドル等)パラメーターが適用/記録され、ポリシー情報も記録します
        OpenBrowser = PipeCore.OpenBrowserAndConnectCDP(caps)
    End If

    '4. 起動中のブラウザ情報を変数に記録
    getBrowserInfo

    '5. 事前に`DeveloperToolsAvailability`情報を取得する
    CurrentDeveloperToolsAvailability.Allowlist = PolicyList.CurrentDeveloperToolsAvailabilityAllowlist
    CurrentDeveloperToolsAvailability.Blocklist = PolicyList.CurrentDeveloperToolsAvailabilityBlocklist
    CurrentDeveloperToolsAvailability.Fallback = PolicyList.CurrentDeveloperToolsAvailability
    CurrentDeveloperToolsAvailability.Collected = PolicyList.isCheckedPolicy

    '6. ブラウザステータス表示
    printParams
End Function

'***************************************************************************************************
'* 機能　　：`targetCreated/targetInfoChanged/targetDestroyed`の3種の非同期イベントを流しつつ、Context のカウントも行います
'---------------------------------------------------------------------------------------------------
'* SET引数 ：論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：・`CDPContext.cls`の生存状況把握に使います
'            ・うまい具合に、`CDPContext.cls`確保中のカウントも行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`CDPContext.cls`を生成/破棄時に呼ぶ想定です。
'            ・`discover:=False`の状態から`discover:=True`にすると、今存在するタブが、`targetCreated`として一斉に送信されます。大量タブがあるといっぱい飛んで面倒ごとになるので、適切な参照カウントをすること
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
'* 機能　　：事前に通信経路を設定した`ReattachOptions.cls`を渡して、制御を再開します
'---------------------------------------------------------------------------------------------------
'* 引数　　：reattachInfo   リアタッチ系のオプションオブジェクト
'---------------------------------------------------------------------------------------------------
'* 注意事項：・ハンドルが生きてない場合は、VBAエラーで止まります
'            ・WebSocketモードの場合は、`destruction:=False`推奨です。ヘッダー情報にサイズ情報があるためです
'***************************************************************************************************
Public Sub reattach(reattachInfo As ReattachOptions)
    Const FromProcedureName As String = ThisClassName & ".reattach"


    '1. プロトコルコアオブジェクトを受けとる
    Set PipeCore = reattachInfo.ClsGet_CDPCore

    '2. handleが生きてるかを確認します
    If isLiveBrowser Then
        '2-1. 溜まってるデータを取り出しておく
        TakeEvents destruction:=reattachInfo.Set_Destruction

        '2-2. ステータス更新
        getBrowserInfo
        printParams
    Else
        printMsg ERROR_, "CDP communication failed with the remotely touched username: '" & Me.CurrentBrowserUser & "' .", FromProcedureName, doRaiseError:=True
    End If

    '3. 再設定が必要なパラメーターを適用
    CustomUAMode = reattachInfo.Set_EnableUserAgentOverride

    '4. 設定が済んだので、破棄
    '※デフォルトインスタンスも破棄しないと、正常な参照カウントダウンができないので、合わせて行う
    Set reattachInfo = Nothing
    Set ReattachOptions = Nothing
End Sub

'***************************************************************************************************
'* 機能　　：ハンドルの生存確認を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：ハンドル生存論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：通信経路のハンドルが生きてるかをチェックします
'***************************************************************************************************
Public Function isLiveBrowser() As Boolean
    isLiveBrowser = PipeCore.isAvailability
End Function



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
    params("url") = Url
    If newWindow Then params("newWindow") = newWindow
    If isHidden Then params("hidden") = isHidden
    If LenB(browserContextId) Then params("browserContextId") = browserContextId
    If isBackground Then params("background") = isBackground

    '2. 実行
    Dim results As BiDiCDPJConv
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
        '1. アクティブなプロトコルオブジェクトを渡す
        .Init Me, PipeCore

        '2. 先ほど生成した`sessionId`を渡す
        .CurrentSessionID = sessionID_

        '3. URLが指定されてる場合は読み込み状況を更新
        If LenB(Url) Then .PageReadyState (True)
    End With
End Function

'***************************************************************************************************
'* 機能　　：既に開いているタブのセッションIDを、タブの名前またはURLに基づいて検索します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - CDPContext
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
    Dim targets As BiDiCDPJConv
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
        '1. アクティブなプロトコルオブジェクトを渡す
        .Init Me, PipeCore

        '2. 先ほど生成した`sessionId`を渡す
        .CurrentSessionID = sessionID_
    End With
End Function

'***************************************************************************************************
'* 機能　　：指定タブ(TargetID/FrameId)から開かれたタブ(主にポップアップウィンドウ等)を検索するのに特化した専用メソッド版です
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - CDPContext。条件に合うタブがない場合は`nothing`です。
'* 引数　　：PageID         TargetID/FrameId
'            Strict         True ：`PageID=FrameId`として扱います
'                           False：`PageID=TargetID`として扱います
'
'            doRetrySecond  何秒間、探索するかを決めます。0で、1回のみです
'---------------------------------------------------------------------------------------------------
'* 機能説明：1. bas側で事前に、`Set Me.RunCollect_TargetCreated = New Collection`し、タブ生成監視を開始します
'            2. 別タブ/ウィンドウが開くであろう`CDPContext`にて、その操作(クリック等)をします
'            3. `CDPContext.CurrentTargetID`を`PageID`に渡します。iframe側でクリックした場合は、`Strict:=True`にします
'            4. クリック後、直ぐに`Target.targetCreated`が来ないことがあるので必要に応じて、`doRetrySecond`でタイムアウトの調整をします
'            5. 条件に合う`Target.TargetInfo`が来たらそれを`CDPContext`として返却し、`Collection-TargetInfos`から、そのヒットしたのを除去します
'            6. 用が済んだら、`Set Me.RunCollect_TargetCreated = Nothing`と打って、タブ生成監視を終了します
'
'* 注意事項：ページ待機も合わせて行う場合は、`CDPContext-Wait`によるJavaScript簡易待機の方を使って下さい。
'            `CDPContext-WaitEvents`では、レースコンディションの都合上、機能しません。
'***************************************************************************************************
Public Function getTabForOpenerID(PageID As String, Optional Strict As Boolean, Optional doRetrySecond As Double) As CDPContext
    Const FromProcedureName As String = ThisClassName & ".getTabForOpenerID"


    Dim tabId As String
    Dim TargetInfoCount As Long

    '1. 検索キー名を決める
    Dim SearchKey As String
    SearchKey = IIf(Strict, "openerFrameId", "openerId")

    '2. タイマースタートして、For Each で蓄積した`Target.targetCreated`分、回す探索を開始する
    Dim timerStart As Double
    If doRetrySecond > 0 Then timerStart = CDPHelpers.TimerCounter
    Do
        Dim TargetInfoStr
        TargetInfoCount = 0 'Do~Loopでの、カウントリセット
        For Each TargetInfoStr In TargetInfos
            '2-1. Json文字列をパース
            Dim TargetInfo As BiDiCDPJConv
            Set TargetInfo = BiDiCDPJConv.Parse(CStr(TargetInfoStr)).NodeKey("params").NodeKey("targetInfo")
            TargetInfoCount = TargetInfoCount + 1

            '2-2. 対象のOpenerがあるか？
            If PageID = TargetInfo.StringKey(SearchKey) Then
                '見つかったので、見つかった`targetID`を取得
                tabId = TargetInfo.StringKey("targetId")

                '次回、重複して取らないように、リストから削除
                TargetInfos.Remove TargetInfoCount
                Exit Do
            End If
        Next

        '3. タイムアウト判定
        If CDPHelpers.TimerCounter - timerStart > doRetrySecond * 1000 Then
            'ログ出して抜ける
            printMsg WARN_, "The tab for " & SearchKey & ": " & PageID & " could not be found.", FromProcedureName
            Exit Function
        Else
            'イベント回収
            TakeEvents
        End If
    Loop

    '4 .Return the session object for reference
    Dim sessionID_ As String
    sessionID_ = attachToTab(tabId)

    '------ 5.専用のクラスとして返却する準備 ------
    EnableDiscoverTargets = True    '`CDPContext.cls`を生成するので、カウントアップ
    Set getTabForOpenerID = New CDPContext
    With getTabForOpenerID
        '1. アクティブなプロトコルオブジェクトを渡す
        .Init Me, PipeCore

        '2. 先ほど生成した`sessionId`を渡す
        .CurrentSessionID = sessionID_
    End With
End Function

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
    Set ShSetting01_StartBrowser.TableBrowserContext(Me.CurrentBrowserUser, FromProcedureName) = BrowserInfo
End Sub

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
    Dim targets As BiDiCDPJConv
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
    Dim targets As BiDiCDPJConv
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
Public Function ExecuteCDP(methodName As String, Optional params As Dictionary, Optional StopCDPError As Boolean = True) As BiDiCDPJConv
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
    Dim resultJson As BiDiCDPJConv
    Set resultJson = BiDiCDPJConv.Parse(ResultJsonStr)

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
    If CDPHelpers.currentLog.seeRawSendMsgDbg Then printMsg Debug_, "ExecuteCDP `" & methodName & "` with result: " & ResultJsonStr, FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`ExecuteCDP`の非同期実行(コマンド結果を待たない)バージョンです
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド 実行時のID値(整理券番号のイメージでOK) ※API失敗時は0で、成功時は1以上です。
'* 引数　　：※`ExecuteCDP`と一緒のため、省略
'***************************************************************************************************
Public Function ExecuteCDPAsync(methodName As String, Optional params As Dictionary, Optional StopApiError As Boolean = True) As Long
    Const FromProcedureName As String = ThisClassName & ".ExecuteCDPAsync"


    '1. コマンド情報をブラウザへ送信し、実行IDを取得
    ExecuteCDPAsync = PipeCore.ReadyRunCDP(methodName, params) 'Send message to the browser instance

    '2. 送信APIエラーチェック
    If PipeCore.LastErrorWriteFile Then
        printMsg ERROR_, "Error WriteFile in CDPCore.writeProcCDP. ErrorCode:" & PipeCore.LastErrorWriteFile & " - " & WinApiError.GetMessage(PipeCore.LastErrorWriteFile, "kernel32"), FromProcedureName, doRaiseError:=StopApiError

        ExecuteCDPAsync = 0 '失敗なので、0にセット
        Exit Function
    End If

    '3. ログ化
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
Private Function ExecuteError(result As BiDiCDPJConv, Optional ByRef ErrorCode As Long, Optional ByRef ErrorMessage As String) As Boolean
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
            Dim errNode As BiDiCDPJConv
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
Public Sub CloseBrowser()
'-----------------------------------------------------------------------------------
' Closes the currently attached browser instance
' Original Name: closeBrowser
' Update: 24/10/2025: Daniel Polak - avoid immediate hard kills, reduce race conditions, and make shutdown behavior predictable and robust in non-responsive cases
'-----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".CloseBrowser"


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
    ShSetting01_StartBrowser.DeleteBrowserContext(FromProcedureName) = Me.CurrentBrowserUser    'メインタブ情報を削除
    DisconnectCDP                                                                               '各種ハンドルも閉じる

    '5. it seems without waiting a bit the browser crashes and the next time it wants to recover from a crash
    CDPHelpers.Sleep 1
    printMsg info_, "The browser has been terminated.", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：ローカルで起動したデバッグブラウザプロセスハンドルに対して、プロセス終了をかけます
'---------------------------------------------------------------------------------------------------
'* 注意事項：・CDPコマンド`Browser.close`を実行したのに、閉じれないアクシデント時に使用して下さい
'            ・WebView2では使えません
'***************************************************************************************************
Private Sub TerminateDebugBrowser()
    TerminateProcess PipeCore.Connection_hProcess, 0
End Sub

'***************************************************************************************************
'* 機能　　：ブラウザは終了せずに、接続だけ切り離し、テーブル情報からも削除します
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebSocketモード以外でこれを呼び出すと、二度と再接続できなくなりますのでご注意下さい
'***************************************************************************************************
Public Sub DisconnectCDP()
    PipeCore.CloseHandleBrowser
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
            '1-2. `TargetInfo`を収集しない場合は、蓄積処理しない
            If Not TargetInfos Is Nothing Then
                '1-3. そのまま、蓄積
                printMsg Debug_, "Collecting `Target.targetCreated`... → " & RawJson, FromProcedureName
                TargetInfos.Add RawJson
            End If

        Case Else
            'None

    End Select

    '2. 拡張用Classへのイベント発信
    RaiseEvent ExEvent(methodName, RawJson)
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
    printMsg Debug_, m_LogID & " Browser - Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "Browser Product   | " & br.Product, FromProcedureName
    printMsg Debug_, "Browser Ver       | " & br.Ver, FromProcedureName
    printMsg Debug_, "Browser JsVersion | " & br.JsVersion, FromProcedureName
    printMsg Debug_, "Browser CDPVer    | " & br.ProtocolVersion, FromProcedureName
    printMsg Debug_, "Browser Revision  | " & br.Revision, FromProcedureName
    printMsg Debug_, "Browser UserAgent | " & br.UserAgent, FromProcedureName
    printMsg Debug_, "Browser UserData  | " & Me.CurrentBrowserUser, FromProcedureName
    printMsg Debug_, "Core dwProcessId  | " & PipeCore.Connection_dwProcessId, FromProcedureName
    printMsg Debug_, "Tool Version      | " & ToolVer, FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：`Browser.getVersion`によるブラウザ情報を取得します
'***************************************************************************************************
Private Sub getBrowserInfo()
    '1. バージョン番号を取得
    Dim BrowserInfos As BiDiCDPJConv: Set BrowserInfos = ExecuteCDP("Browser.getVersion")

    '2. VBA用にキャッシュする
    With br
        'オリジナルのまま
        .JsVersion = BrowserInfos.StringKey("jsVersion")
        .Product = BrowserInfos.StringKey("product")
        .ProtocolVersion = BrowserInfos.StringKey("protocolVersion")
        .Revision = BrowserInfos.StringKey("revision")
        .UserAgent = BrowserInfos.StringKey("userAgent")

        '少し便利に
        br.Ver = Split(.Product, "/")(1)    'そのChromium製品のバージョン番号
    End With
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
'* 注意事項：・ポリシー`DeveloperToolsAvailability`系があると、「Not allowed」で失敗します。
'              その場合は、WebView2ルートをご利用ください
'            ・「Not allowed」で失敗した場合は、起動時に記録したポリシー情報(レジストリのみ)から、原因を推定してエラー表示します
'***************************************************************************************************
Public Function attachToTab(tabId As String) As String
    Const FromProcedureName As String = ThisClassName & ".attachToTab"


    '1. コマンド準備
    Dim CDPparams As New Dictionary
    CDPparams.Add "targetId", tabId
    CDPparams.Add "flatten", True

    '2. 実行
    '※独自のエラー処理が必要なのでここでは一旦、無視させる
    Dim tmp As BiDiCDPJConv
    Set tmp = ExecuteCDP("Target.attachToTarget", CDPparams, False)

    '3. 結果があれば、返却
    If Not tmp Is Nothing Then
        attachToTab = tmp.StringKey("sessionId")
    Else
        '3-1. エラーメッセージを取り出す
        Dim ErrMes As String
        If Not LastCDPJsonError Is Nothing Then ErrMes = LastCDPJsonError("message")

        '3-2. 「Not allowed」の場合は、ポリシー由来の原因を推定して添える
        If ErrMes = "Not allowed" Then
            ErrMes = ErrMes & vbCrLf & vbCrLf & ExplainDevToolsBlock(tabId)

            '3-3. 停止
            printMsg ERROR_, "Failed to retrieve `SessionID`. Reason: " & ErrMes, FromProcedureName, doRaiseError:=True
        Else
            printMsg WARN_, "Failed to retrieve `SessionID`. Reason: " & ErrMes, FromProcedureName
        End If
    End If
End Function

'***************************************************************************************************
'* 機能　　：SessionID の破棄を、非同期で行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：SessionID      破棄したいSessionID
'---------------------------------------------------------------------------------------------------
'* 詳細説明：不必要なSessionIDを破棄することで、ブラウザの安定性が向上し、奇妙な不具合も消えるかもしれません
'* 注意事項：ブラウザ権限で実行するため、メイン化してる`SessionID`も破棄すると、処理が止まるため、使う箇所は慎重に...
'***************************************************************************************************
Public Sub DiscardSessionID(sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".DiscardSessionID"


    '1. コマンド準備
    Dim CDPparams As New Dictionary
    CDPparams.Add "sessionId", sessionID

    '2. 破棄コマンド実行
    ExecuteCDPAsync "Target.detachFromTarget", CDPparams, False
End Sub



'***************************************************************************************************
'                               ■■■ ヘルパーメソッド ■■■
'***************************************************************************************************
'* 機能　　：`Target.attachToTarget`が「Not allowed」になった原因を、ポリシー情報から推定します
'---------------------------------------------------------------------------------------------------
'* 返り値　：推定結果のメッセージ
'* 引数　　：tabId  TargetID
'---------------------------------------------------------------------------------------------------
'* 機能説明：Chromium`DeveloperToolsPolicyChecker::GetDevToolsAvailabilityForUrl`と同じ順で判定します
'            1. URL が Allowlist に一致 → 許可
'               ※両リストに一致した場合は、より具体的なパターンが勝ちます。同点なら Allowlist の勝ち
'            2. Allowlist があり、Blocklist が無い → Allowlist にない URL はブロック
'            3. URL が Blocklist に一致 → ブロック
'            4. それ以外 → `DeveloperToolsAvailability`の値に従う
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも「推定」です。最終的な判定は、ブラウザ側の応答が正です
'            ・レジストリ(HKLM/HKCU)のポリシーのみが対象です。クラウド配信のポリシーは見ていません
'            ・URLパターンの照合は簡易版です(Chromium`URLBlocklist`の主要ルールのみ)
'            ・起動済みブラウザへの接続時等、ポリシー情報を取得していない場合は、推定せずに確認先のみ案内します
'***************************************************************************************************
Private Function ExplainDevToolsBlock(tabId As String) As String
    Const OutsideRegistry As String = "The cause is likely outside the registry policies (e.g. cloud-delivered policies, or a non-policy restriction)."


    '1. 対象タブのURLを取得
    Dim TargetUrl As String: TargetUrl = GetTargetUrl(tabId)

    '※ポリシー情報を取得していない場合は、推定できないので確認先のみ案内して抜ける
    If Not CurrentDeveloperToolsAvailability.Collected Then
        ExplainDevToolsBlock = "=== Policy diagnosis (registry: HKLM/HKCU) ===" & vbCrLf & _
                               "Target URL                          : " & TargetUrl & vbCrLf & _
                               "Probable cause                      : Policy info was not collected (e.g. attached to an already-running browser), so the cause cannot be estimated. " & _
                               "It may be blocked by DeveloperToolsAvailability / DeveloperToolsAvailabilityAllowlist / DeveloperToolsAvailabilityBlocklist." & vbCrLf & _
                               "How to check                        :" & vbCrLf & _
                               "  1. Open chrome://policy (Edge also accepts it and redirects to edge://policy)." & vbCrLf & _
                               "  2. If it is blocked, check these keys with the Registry Editor (reading them does not require administrator rights):" & vbCrLf & _
                               "       Edge  : HKLM\SOFTWARE\Policies\Microsoft\Edge, HKCU\SOFTWARE\Policies\Microsoft\Edge" & vbCrLf & _
                               "       Chrome: HKLM\SOFTWARE\Policies\Google\Chrome, HKCU\SOFTWARE\Policies\Google\Chrome" & vbCrLf & _
                               "  3. If neither is available, ask your system administrator."
        Exit Function
    End If

    '2. 判定材料を用意
    With CurrentDeveloperToolsAvailability
        Dim HasAllowlist As Boolean: HasAllowlist = IsArray(.Allowlist)
        Dim HasBlocklist As Boolean: HasBlocklist = IsArray(.Blocklist)

        '各リストで、URLに一致する最も優先度の高いパターンを探す
        Dim AllowHit As UrlPatternInfo: AllowHit = FindHighestPriorityPattern(TargetUrl, .Allowlist, True)
        Dim BlockHit As UrlPatternInfo: BlockHit = FindHighestPriorityPattern(TargetUrl, .Blocklist, False)

        '両方に一致した場合は、より優先度の高い方が勝つ
        Dim AllowWins As Boolean
        If AllowHit.IsValid Then
            If BlockHit.IsValid Then
                AllowWins = Not PatternTakesPrecedence(BlockHit, AllowHit)
            Else
                AllowWins = True
            End If
        End If

        '3. Chromiumと同じ順で原因を推定する
        Dim Cause As String
        If AllowWins Then
            '3-1. Allowlist に一致 → 本来は許可されるはず
            Cause = "[1] The URL matches the DeveloperToolsAvailabilityAllowlist pattern '" & AllowHit.Pattern & "', so it should be allowed. " & OutsideRegistry

        ElseIf HasAllowlist And Not HasBlocklist Then
            '3-2. Allowlist のみ構成 → Allowlist にない URL はブロック
            Cause = "[2] DeveloperToolsAvailabilityAllowlist is configured but DeveloperToolsAvailabilityBlocklist is not, so any URL not in the Allowlist is blocked."

        ElseIf BlockHit.IsValid Then
            '3-3. Blocklist に一致 → ブロック
            Cause = "[3] The URL matches the DeveloperToolsAvailabilityBlocklist pattern '" & BlockHit.Pattern & "'."

        ElseIf IsEmpty(.Fallback) Then
            '3-4. フォールバック(DeveloperToolsAvailability)が未設定 → 既定値0と同じ扱い
            Cause = "[4] No URL rule matched, and DeveloperToolsAvailability is not set (default = 0: blocked only for force-installed extensions). " & IIf(LCase$(TargetUrl) Like "*extension://*", "The target looks like an extension page, so it may be a force-installed extension.", OutsideRegistry)

        Else
            '3-4. フォールバック(DeveloperToolsAvailability)の値に従う
            Select Case CLng(.Fallback)
                Case 2
                    Cause = "[4] No URL rule matched, and DeveloperToolsAvailability = 2: DevTools are disallowed for all URLs."
                Case 1
                    Cause = "[4] No URL rule matched, and DeveloperToolsAvailability = 1: DevTools are allowed. " & OutsideRegistry
                Case 0
                    Cause = "[4] No URL rule matched, and DeveloperToolsAvailability = 0: blocked only for force-installed extensions. " & IIf(LCase$(TargetUrl) Like "*extension://*", "The target looks like an extension page, so it may be a force-installed extension.", OutsideRegistry)
                Case Else
                    Cause = "[4] No URL rule matched, and DeveloperToolsAvailability has an unknown value: " & CStr(.Fallback)
            End Select
        End If

        '4. 判定材料と一緒に返却
        ExplainDevToolsBlock = "=== Policy diagnosis (registry: HKLM/HKCU) ===" & vbCrLf & _
                               "Target URL                          : " & TargetUrl & vbCrLf & _
                               "DeveloperToolsAvailabilityAllowlist : " & PolicyValueToString(.Allowlist) & vbCrLf & _
                               "DeveloperToolsAvailabilityBlocklist : " & PolicyValueToString(.Blocklist) & vbCrLf & _
                               "DeveloperToolsAvailability          : " & PolicyValueToString(.Fallback) & vbCrLf & _
                               "Probable cause                      : " & Cause
    End With
End Function

'***************************************************************************************************
'* 機能　　：TargetID から、現在のURLを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：URL  ※取得できない場合は、`vbNullString`
'* 引数　　：tabId  TargetID
'***************************************************************************************************
Private Function GetTargetUrl(tabId As String) As String
    Dim CDPparams As New Dictionary
    CDPparams.Add "targetId", tabId

    Dim tmp As BiDiCDPJConv
    Set tmp = ExecuteCDP("Target.getTargetInfo", CDPparams, False)
    If Not tmp Is Nothing Then GetTargetUrl = tmp.NodeKey("targetInfo").StringKey("url")
End Function

'***************************************************************************************************
'* 機能　　：ポリシー値を、ログ表示用の文字列にします
'---------------------------------------------------------------------------------------------------
'* 返り値　：表示用文字列  ※リストは`[a, b]`、未設定は`(not set)`
'* 引数　　：PolicyValue  ポリシー値
'***************************************************************************************************
Private Function PolicyValueToString(ByVal PolicyValue As Variant) As String
    If IsArray(PolicyValue) Then
        PolicyValueToString = "[" & Join(PolicyValue, ", ") & "]"
    ElseIf IsEmpty(PolicyValue) Then
        PolicyValueToString = "(not set)"
    Else
        PolicyValueToString = CStr(PolicyValue)
    End If
End Function

'***************************************************************************************************
'* 機能　　：URLパターンのリストから、URLに一致する最も優先度の高いパターンを探します
'---------------------------------------------------------------------------------------------------
'* 返り値　：一致したパターンの解析結果  ※一致なしの場合は、`IsValid = False`
'* 引数　　：Url          照合するURL
'            PatternList  URLパターンの配列  ※未設定(Empty)なら一致なし
'            isAllowlist  Allowlist側のリストか？
'***************************************************************************************************
Private Function FindHighestPriorityPattern(Url As String, PatternList As Variant, isAllowlist As Boolean) As UrlPatternInfo
    If Not IsArray(PatternList) Then Exit Function

    Dim Best As UrlPatternInfo
    Dim i As Long
    For i = LBound(PatternList) To UBound(PatternList)
        Dim Candidate As UrlPatternInfo
        Candidate = ParseUrlPattern(CStr(PatternList(i)), isAllowlist)
        If Candidate.IsValid Then
            If MatchUrlPattern(Url, Candidate) Then
                If Not Best.IsValid Then
                    Best = Candidate
                ElseIf PatternTakesPrecedence(Candidate, Best) Then
                    Best = Candidate
                End If
            End If
        End If
    Next
    FindHighestPriorityPattern = Best
End Function

'***************************************************************************************************
'* 機能　　：URLパターン(URLBlocklist形式)を解析します
'---------------------------------------------------------------------------------------------------
'* 返り値　：解析結果  ※解析できない場合は、`IsValid = False`
'* 引数　　：Pattern      URLパターン  例：`*`、`https://`、`example.com`、`.example.com`、`https://example.com:8080/path?key=value`
'            isAllowlist  Allowlist側のパターンか？
'---------------------------------------------------------------------------------------------------
'* 機能説明：Chromium`url_matcher::util::FilterToComponents`の主要ルールを移植しています
'            ・`*`は、すべてのURLに一致
'            ・`scheme:*`または`scheme://*`は、そのスキームのすべてのURLに一致
'            ・ホストは、既定でサブドメインも一致。先頭に`.`を付けると、そのホストのみ一致。IPアドレスもそのホストのみ一致
'            ・パスは前方一致
'***************************************************************************************************
Private Function ParseUrlPattern(ByVal Pattern As String, isAllowlist As Boolean) As UrlPatternInfo
    Dim info As UrlPatternInfo
    info.Pattern = Pattern
    info.Allow = isAllowlist

    Dim Rest As String: Rest = Trim$(Pattern)
    If LenB(Rest) = 0 Then Exit Function

    '1. `*`は、すべてのURLに一致
    If Rest = "*" Then
        info.MatchSubdomains = True
        info.IsValid = True
        ParseUrlPattern = info
        Exit Function
    End If

    '2. スキーム部分を切り出す
    Dim Pos As Long: Pos = InStr(Rest, "://")
    If Pos > 0 Then
        info.Scheme = LCase$(Left$(Rest, Pos - 1))
        Rest = Mid$(Rest, Pos + 3)
    ElseIf Right$(Rest, 2) = ":*" And InStr(Rest, "/") = 0 Then
        '`scheme:*`形式
        info.Scheme = LCase$(Left$(Rest, Len(Rest) - 2))
        Rest = "*"
    ElseIf LCase$(Left$(Rest, 5)) = "data:" Then
        '`data:`は、中身を前方一致
        info.Scheme = "data"
        info.Path = Mid$(Rest, 6)
        info.MatchSubdomains = True
        info.IsValid = True
        ParseUrlPattern = info
        Exit Function
    End If

    '3. `scheme://*`形式、`file://`は、ホストなしで扱う
    If Rest = "*" And LenB(info.Scheme) > 0 Then
        info.MatchSubdomains = True
        info.IsValid = True
        ParseUrlPattern = info
        Exit Function
    ElseIf info.Scheme = "file" Then
        info.Path = Replace(Rest, "\", "/")
        If Left$(info.Path, 1) <> "/" Then info.Path = "/" & info.Path
        info.MatchSubdomains = True
        info.IsValid = True
        ParseUrlPattern = info
        Exit Function
    End If

    '4. フラグメント、クエリ、パスの順に切り出す
    Pos = InStr(Rest, "#")
    If Pos > 0 Then Rest = Left$(Rest, Pos - 1)

    Pos = InStr(Rest, "?")
    If Pos > 0 Then
        info.Query = Mid$(Rest, Pos + 1)
        Rest = Left$(Rest, Pos - 1)
        info.QueryCount = CountQueryConditions(info.Query)
    End If

    Pos = InStr(Rest, "/")
    If Pos > 0 Then
        info.Path = Mid$(Rest, Pos)
        Rest = Left$(Rest, Pos - 1)
    End If

    '5. ホストとポートに分ける
    Dim HostPart As String, PortPart As String
    If Not SplitHostPort(Rest, HostPart, PortPart) Then Exit Function
    If LenB(HostPart) = 0 Then Exit Function    'ホストは必須

    If LenB(PortPart) > 0 Then
        If PortPart Like "*[!0-9]*" Or Len(PortPart) > 5 Then Exit Function
        If CLng(PortPart) <= 0 Or CLng(PortPart) > 65535 Then Exit Function
        info.port = CLng(PortPart)
    End If

    '6. ホストの一致方法を決める
    If HostPart = "*" Then
        'すべてのホスト
        info.Host = vbNullString
        info.MatchSubdomains = True
    ElseIf Left$(HostPart, 1) = "." Then
        '先頭`.`は、そのホストのみ
        info.Host = LCase$(Mid$(HostPart, 2))
        info.MatchSubdomains = False
    ElseIf Left$(HostPart, 1) = "[" Or Not HostPart Like "*[!0-9.]*" Then
        'IPアドレスは、そのホストのみ
        info.Host = LCase$(HostPart)
        info.MatchSubdomains = False
    Else
        'サブドメインも一致させる。ドメインの区切りで一致させるため、先頭に`.`を付ける
        info.Host = "." & LCase$(HostPart)
        info.MatchSubdomains = True
    End If

    info.IsValid = True
    ParseUrlPattern = info
End Function

'***************************************************************************************************
'* 機能　　：URLが、URLパターンに一致するか判定します
'---------------------------------------------------------------------------------------------------
'* 返り値　：True：一致
'* 引数　　：Url   照合するURL
'            Info  `ParseUrlPattern`で解析したURLパターン
'***************************************************************************************************
Private Function MatchUrlPattern(ByVal Url As String, info As UrlPatternInfo) As Boolean
    '1. URLを分解
    Dim uScheme As String, uHost As String, uPort As Long, uPath As String, uQuery As String
    If Not SplitUrl(Url, uScheme, uHost, uPort, uPath, uQuery) Then Exit Function

    '2. スキーム
    If LenB(info.Scheme) > 0 Then
        If uScheme <> info.Scheme Then Exit Function
    End If

    '3. ホスト
    If LenB(info.Host) > 0 Then
        If info.MatchSubdomains Then
            If Right$("." & uHost, Len(info.Host)) <> info.Host Then Exit Function
        Else
            If uHost <> info.Host Then Exit Function
        End If
    End If

    '4. ポート
    If info.port <> 0 And uPort <> info.port Then Exit Function

    '5. パス(前方一致)
    If LenB(info.Path) > 0 Then
        If Left$(uPath, Len(info.Path)) <> info.Path Then Exit Function
    End If

    '6. クエリ
    If LenB(info.Query) > 0 Then
        If Not MatchQuery(uQuery, info.Query) Then Exit Function
    End If

    MatchUrlPattern = True
End Function

'***************************************************************************************************
'* 機能　　：2つのURLパターンのうち、左側が優先されるか判定します
'---------------------------------------------------------------------------------------------------
'* 返り値　：True：lhs が優先
'* 引数　　：lhs, rhs  比較するURLパターン
'---------------------------------------------------------------------------------------------------
'* 機能説明：Chromium`FilterTakesPrecedence`の移植です。上から順に比べます
'            1. Blocklist の`*`は、最も優先度が低い
'            2. そのホストのみ一致(先頭`.`等)の方が優先
'            3. ホストが長い方が優先
'            4. パスが長い方が優先
'            5. クエリ条件が多い方が優先
'            6. 同点なら Allowlist 側が優先
'***************************************************************************************************
Private Function PatternTakesPrecedence(lhs As UrlPatternInfo, rhs As UrlPatternInfo) As Boolean
    If IsWildcardBlockPattern(rhs) Then PatternTakesPrecedence = True: Exit Function
    If IsWildcardBlockPattern(lhs) Then Exit Function

    If lhs.MatchSubdomains <> rhs.MatchSubdomains Then PatternTakesPrecedence = Not lhs.MatchSubdomains: Exit Function
    If Len(lhs.Host) <> Len(rhs.Host) Then PatternTakesPrecedence = Len(lhs.Host) > Len(rhs.Host): Exit Function
    If Len(lhs.Path) <> Len(rhs.Path) Then PatternTakesPrecedence = Len(lhs.Path) > Len(rhs.Path): Exit Function
    If lhs.QueryCount <> rhs.QueryCount Then PatternTakesPrecedence = lhs.QueryCount > rhs.QueryCount: Exit Function

    PatternTakesPrecedence = lhs.Allow And Not rhs.Allow
End Function

'***************************************************************************************************
'* 機能　　：Blocklist の`*`(すべてのURLに一致)か判定します
'***************************************************************************************************
Private Function IsWildcardBlockPattern(info As UrlPatternInfo) As Boolean
    With info
        IsWildcardBlockPattern = Not .Allow And LenB(.Scheme) = 0 And LenB(.Host) = 0 And LenB(.Path) = 0 And LenB(.Query) = 0 And .port = 0 And .MatchSubdomains
    End With
End Function

'***************************************************************************************************
'* 機能　　：URLを、スキーム/ホスト/ポート/パス/クエリに分解します
'---------------------------------------------------------------------------------------------------
'* 返り値　：True：分解できた
'* 引数　　：Url                                  分解するURL
'            Scheme, Host, Port, Path, Query      [ByRef]分解結果  ※ポートは省略時、スキームの既定値
'---------------------------------------------------------------------------------------------------
'* 注意事項：`about:blank`や`data:`等のホストがないURLは、`scheme:`以降をパスとして扱います
'***************************************************************************************************
Private Function SplitUrl(ByVal Url As String, ByRef Scheme As String, ByRef Host As String, ByRef port As Long, ByRef Path As String, ByRef Query As String) As Boolean
    '1. スキーム
    Dim Pos As Long: Pos = InStr(Url, ":")
    If Pos <= 1 Then Exit Function
    Scheme = LCase$(Left$(Url, Pos - 1))
    Dim Rest As String: Rest = Mid$(Url, Pos + 1)

    '2. フラグメント、クエリを切り出す
    Pos = InStr(Rest, "#")
    If Pos > 0 Then Rest = Left$(Rest, Pos - 1)

    Pos = InStr(Rest, "?")
    If Pos > 0 Then
        Query = Mid$(Rest, Pos + 1)
        Rest = Left$(Rest, Pos - 1)
    End If

    '3. ホストがないURL(`about:blank`等)
    If Left$(Rest, 2) <> "//" Then
        Path = Rest
        SplitUrl = True
        Exit Function
    End If
    Rest = Mid$(Rest, 3)

    '4. パスを切り出す
    Pos = InStr(Rest, "/")
    If Pos > 0 Then
        Path = Mid$(Rest, Pos)
        Rest = Left$(Rest, Pos - 1)
    Else
        Path = "/"
    End If

    '5. ユーザー情報(`user:pass@`)を除く
    Pos = InStrRev(Rest, "@")
    If Pos > 0 Then Rest = Mid$(Rest, Pos + 1)

    '6. ホストとポート
    Dim PortPart As String
    If Not SplitHostPort(Rest, Host, PortPart) Then Exit Function
    Host = LCase$(Host)

    If LenB(PortPart) > 0 And Not PortPart Like "*[!0-9]*" And Len(PortPart) <= 5 Then
        port = CLng(PortPart)
    Else
        Select Case Scheme
            Case "http", "ws":    port = 80
            Case "https", "wss":  port = 443
            Case "ftp":           port = 21
            Case Else:            port = 0
        End Select
    End If

    SplitUrl = True
End Function

'***************************************************************************************************
'* 機能　　：`host:port`を、ホストとポートに分けます
'---------------------------------------------------------------------------------------------------
'* 返り値　：True：分けられた
'* 引数　　：HostPort        `host:port`形式の文字列  ※IPv6の`[::1]:8080`にも対応
'            Host, PortPart  [ByRef]分けた結果  ※ポートなしの場合、PortPartは空
'***************************************************************************************************
Private Function SplitHostPort(ByVal HostPort As String, ByRef Host As String, ByRef PortPart As String) As Boolean
    Dim Pos As Long
    If Left$(HostPort, 1) = "[" Then
        'IPv6
        Pos = InStr(HostPort, "]")
        If Pos = 0 Then Exit Function
        Host = Left$(HostPort, Pos)
        If Mid$(HostPort, Pos + 1, 1) = ":" Then PortPart = Mid$(HostPort, Pos + 2)
    Else
        Pos = InStrRev(HostPort, ":")
        If Pos > 0 Then
            Host = Left$(HostPort, Pos - 1)
            PortPart = Mid$(HostPort, Pos + 1)
        Else
            Host = HostPort
        End If
    End If
    SplitHostPort = True
End Function

'***************************************************************************************************
'* 機能　　：URLのクエリが、URLパターンのクエリ条件をすべて満たすか判定します
'---------------------------------------------------------------------------------------------------
'* 返り値　：True：すべて満たす
'* 引数　　：UrlQuery      URLのクエリ部分(`?`より後ろ)
'            PatternQuery  URLパターンのクエリ条件(`?`より後ろ)
'---------------------------------------------------------------------------------------------------
'* 機能説明：条件は`&`区切りで、それぞれ次のように照合します
'            ・`key`        ：そのキーがある
'            ・`key=value*` ：そのキーの値が前方一致
'            ・`key=value`  ：そのキーの値が完全一致
'***************************************************************************************************
Private Function MatchQuery(ByVal UrlQuery As String, ByVal PatternQuery As String) As Boolean
    Dim Conditions() As String: Conditions = Split(PatternQuery, "&")
    Dim Elements() As String: Elements = Split(UrlQuery, "&")

    Dim i As Long, j As Long
    For i = LBound(Conditions) To UBound(Conditions)
        If LenB(Conditions(i)) > 0 Then
            Dim Found As Boolean: Found = False
            For j = LBound(Elements) To UBound(Elements)
                If MatchQueryElement(Elements(j), Conditions(i)) Then
                    Found = True
                    Exit For
                End If
            Next
            If Not Found Then Exit Function
        End If
    Next
    MatchQuery = True
End Function

'***************************************************************************************************
'* 機能　　：クエリ要素1つが、クエリ条件1つを満たすか判定します  ※詳細は`MatchQuery`参照
'***************************************************************************************************
Private Function MatchQueryElement(ByVal Element As String, ByVal Condition As String) As Boolean
    Dim ePos As Long: ePos = InStr(Element, "=")
    Dim cPos As Long: cPos = InStr(Condition, "=")

    '1. キーを比較
    Dim eKey As String: eKey = IIf(ePos > 0, Left$(Element, ePos - 1), Element)
    Dim cKey As String: cKey = IIf(cPos > 0, Left$(Condition, cPos - 1), Condition)
    If eKey <> cKey Then Exit Function

    '2. 値の条件がなければ、キーの一致だけでOK
    If cPos = 0 Then MatchQueryElement = True: Exit Function

    '3. 値を比較
    Dim eValue As String
    If ePos > 0 Then eValue = Mid$(Element, ePos + 1)
    Dim cValue As String: cValue = Mid$(Condition, cPos + 1)
    If Right$(cValue, 1) = "*" Then
        cValue = Left$(cValue, Len(cValue) - 1)
        MatchQueryElement = (Left$(eValue, Len(cValue)) = cValue)
    Else
        MatchQueryElement = (eValue = cValue)
    End If
End Function

'***************************************************************************************************
'* 機能　　：クエリ条件の数を数えます  ※優先順位の比較用
'***************************************************************************************************
Private Function CountQueryConditions(ByVal PatternQuery As String) As Long
    Dim Conditions() As String: Conditions = Split(PatternQuery, "&")
    Dim i As Long
    For i = LBound(Conditions) To UBound(Conditions)
        If LenB(Conditions(i)) > 0 Then CountQueryConditions = CountQueryConditions + 1
    Next
End Function



'***************************************************************************************************
'                              ■■■ PUBLIC PROPERTIES ■■■
'***************************************************************************************************
'* 機能　　：このClass内で起動してるブラウザの`user-data-dir`値を受け取ります
'***************************************************************************************************
Property Get CurrentBrowserUser() As String
    CurrentBrowserUser = PipeCore.CurrentUserName
End Property

'***************************************************************************************************
'* 機能　　：このClass内で起動してるデフォルト値ブラウザの`userAgent`値を受け取ります
'***************************************************************************************************
Property Get CurrentBrowseruserAgent() As String
    CurrentBrowseruserAgent = br.UserAgent
End Property

'***************************************************************************************************
'* 機能　　：カスタムユーザーエージェントを適用するか？
'***************************************************************************************************
Property Get UseCustomUAMode() As Boolean
    UseCustomUAMode = CustomUAMode
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
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, CDPHelpers.currentLog.logPath
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
    'Assign an objID for easy debug identification
    m_LogID = "BRID" & Format(Rnd * 1000, "000")

    'タイムアウト設定(デフォルトで、30s)
    timerOut = TIMEOUT_DEFAULT * 1000   'ミリ秒に補正

    '`WithEvents`からのコマンド結果を蓄積させる初期化
    Set ResultCommandCDP_Browser = New Dictionary
End Sub
