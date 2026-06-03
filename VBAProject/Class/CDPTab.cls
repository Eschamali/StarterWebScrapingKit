VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPTab"
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


'===================================
' Win APIs Declarations
'===================================
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


'===================================
' Constants
'===================================
'JSの`document.readyState`の状態一式
Public Enum ReadyState      'Used for .wait method
    isLoading = 0           'equivalence of the browser's "loading" state
    isInteractive = 1       'equivalence of the browser's "interactive" state
    isComplete = 2          'equivalence of the browser's "complete" state
End Enum

'ブラウザからの非同期イベント情報を格納する際のDictionaryKey名設定
Private Const EventsDictionaryKeyName01 As String = "TotalEvents"
Private Const EventsDictionaryKeyName02 As String = "EventMethods"

'その他
Private Const TIMEOUT_DEFAULT   As Double = 30          '初期タイムアウト秒数
Private Const ThisClassName     As String = "CDPTab"    'ログのトレース用
Private Const MaxStoreResultCDP As Long = 2 ^ 8         'ブラウザからのコマンド結果蓄積数



'===================================
' Global Variables
'===================================
'ブラウザ制御クラスオブジェクト
Private WithEvents PipeCore As CDPCore      'イベント受け取りオブジェクト
Private BrowserCore         As CDPBrowser   'Must be initialized for CDP engine

'各種クラスオブジェクト
Private CharConv    As New CharacterCodeConversion
Private DataConv    As New WebCrypto        'Base64のネイティブ制御用
Private ErrorDetail As New WinApiError      'エラーコードから、詳細を取得するやつ

'操作中のタブ情報一式
Private Type StatebBrowserTab
    hWnd        As LongPtr  'Browser object window handle
    WID         As String   'Useful for windows repositioning
    LogID       As String   'The ID of the current CDPBrowser object. Useful when reading debugs of multiple tabs.
    Name        As String   'The focused tab name / window name of the active CDPBrowser
    Url         As String   'The focused tab url of the active CDPBrowser
    reattach    As Boolean  'Remember reAttach preference from .reattach method - updated by Long 22/01/2026
End Type
Private br As StatebBrowserTab

'操作中のタブ情報一式
Private Type StateTab
    sessionID   As String   'The ID to the current session tab.
    targetID    As String   'The targetID of the current tab for ExecuteCDP
End Type
Private brTab As StateTab

'CDPコマンドの結果を一時蓄積する際の管理(`WithEvents`によるズレ防止の仕組み作り)
Private Type cacheResultCDP
    response    As Dictionary   '「ID - それに対するCDPコマンドの結果」というDictionary
    Limit       As Long         'その`Dictionary`に詰め込む上限数
End Type
Private ResultCDP_Tab As cacheResultCDP

'ブラウザからの非同期イベント状況
Private Type StateBrowserEvents
    Queue   As Dictionary  'ブラウザからのイベントデータを蓄積するための変数です。
    Filter  As Dictionary  'ブラウザからのイベントデータのうち、登録したイベント名のみ、`Queue`に蓄積させるフィルター情報です
End Type
Private publicBrowserEvents As StateBrowserEvents   'bas側で使用する用

'各種エラー情報
Private lastErrorDetailsCDP     As Dictionary  'CDP-Jsonコマンド実行時の最後のエラー情報
Private lastExceptionDetailsJS  As Dictionary  'JavaScript実行時の最後のエラー情報

'その他
Public KeepSession  As Boolean  'JavaScriptの実行内容(ObjectID)を別のプロシージャに跨ぐ等で、SessionIDを破棄しないようにするオプション
Private thisApp     As Object   'A dymanic variable used in CDPBrowser initialize procedure to determine if the VBA host is an Excel / Word / Access
Private timerOut    As Double   'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
Private TargetInfos As Collection               '非同期イベント`Target.targetCreated`で得た`Target.TargetInfo`情報。基本的に、`attachToSession`専用



'***************************************************************************************************
'                              ■■■ レスポンス蓄積/処理 ■■■
'***************************************************************************************************
'* 機能　　：このタブに対するCDPレスポンス結果を一旦、Dictionaryへ蓄積させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：id             コマンドID
'            RawJson        コマンド結果の生Json文字列　※As Is. ID付きのままで渡されます
'            sessionID      届いたSessionID
'---------------------------------------------------------------------------------------------------
'* 機能説明：`ID`をKeyに、Dictionaryへ`RawJson`を格納します。
'* 注意事項：・Jsonは、文字列のままで渡されます
'            ・Dictionaryへの蓄積には、上限があります。上限を超えると、古いものから削除されていきます
'***************************************************************************************************
Private Sub PipeCore_CDPTabID(id As Long, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPBrowserEvent"


    '1. このタブクラスで設定中の`SesionId`以外は即抜け
    If sessionID <> brTab.sessionID Then Exit Sub

    '2. 格納件数チェックし、次の蓄積で上限を超えたら、全履歴を削除します
    If ResultCDP_Tab.response.Count + 1 > ResultCDP_Tab.Limit Then
        ResultCDP_Tab.response.RemoveAll
        printMsg WARN_, "The accumulation limit has exceeded " & ResultCDP_Tab.Limit & ", so All result history has been deleted to prevent performance degradation.", FromProcedureName
    End If

    '3. 結果を格納
    ResultCDP_Tab.response.Add id, RawJson
    printMsg Trace_, "The results of the CDP(ID: " & id & ") execution from the tab(sessionId: " & sessionID & ") have been received. → " & RawJson, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：ブラウザに対する非同期イベントを受け取ります
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName 非同期イベントメソッド名
'            RawJson    非同期イベントの生Json文字列　※As Is. メソッド名付きのままで渡されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：このClass独自の非同期イベント処理を行います
'***************************************************************************************************
Private Sub PipeCore_CDPBrowserEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPBrowserEvent"


    '1. Select Case でスマートに、method名ごとの独自処理を行う
    Select Case methodName
        '1-1. `attachToSession`が行えるように準備する
        Case "Target.targetCreated"
            '1-2. `TargetInfo`を収集しない場合は、飛ばす
            If Not TargetInfos Is Nothing Then
                '1-3. そのまま、蓄積
                printMsg Debug_, "Collecting `Target.targetCreated`... → " & RawJson, FromProcedureName
                TargetInfos.Add RawJson
            End If

        Case Else
            'None

    End Select

    '------------------ 2. 非同期イベントを収集します ------------------
    '2-1. イベントキャプチャしない場合は、即抜け
    If publicBrowserEvents.Queue Is Nothing Then printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName: Exit Sub

    '2-2. その非同期イベント名称は、フィルターに含まれない場合も即抜け
    printMsg Trace_, "Checking public event filter settings...", FromProcedureName
    If Not checkFilterEvent(publicBrowserEvents.Filter, methodName) Then Exit Sub

    '2-3. パース処理
    printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
    Dim tmp As Dictionary: Set tmp = BrowserCore.jsConverter.ParseJson(RawJson)

    '2-4. 専用のプロシージャにて、蓄積処理
    printMsg Trace_, "The module is requesting the accumulation of event information. Saving...", FromProcedureName
    OnNewEventReceived publicBrowserEvents.Queue, tmp, methodName
End Sub

'***************************************************************************************************
'* 機能　　：このタブに対する非同期イベントを受け取ります
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName 非同期イベントメソッド名
'            RawJson    非同期イベントの生Json文字列　※As Is. メソッド名付きのままで渡されます
'            sessionID  届いたSessionID
'---------------------------------------------------------------------------------------------------
'* 機能説明：このClass独自の非同期イベント処理を行います
'***************************************************************************************************
Private Sub PipeCore_CDPTabEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".PipeCore_CDPBrowserEvent"


    '1. このタブクラスで設定中の`SessionId`以外は即抜け
    If sessionID <> brTab.sessionID Then Exit Sub

    '------------------ 2. 非同期イベントを収集します ------------------
    '2-1. イベントキャプチャしない場合は、即抜け
    If publicBrowserEvents.Queue Is Nothing Then printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName: Exit Sub

    '2-2. その非同期イベント名称は、フィルターに含まれない場合も即抜け
    printMsg Trace_, "Checking public event filter settings...", FromProcedureName
    If Not checkFilterEvent(publicBrowserEvents.Filter, methodName) Then Exit Sub

    '2-3. パース処理
    printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
    Dim tmp As Dictionary: Set tmp = BrowserCore.jsConverter.ParseJson(RawJson)

    '2-4. 専用のプロシージャにて、蓄積処理
    printMsg Trace_, "The module is requesting the accumulation of event information. Saving...", FromProcedureName
    OnNewEventReceived publicBrowserEvents.Queue, tmp, methodName
End Sub



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザ起動と同時に、非同期イベント：`Target.targetCreated`を利用した初回タブ接続処理を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：userProfile    起動引数：`--user-data-dir`　につけるユーザーディレクトリ名を指定します
'            Name           ブラウザ名。※現時点では、Chrome OR Edge のみ
'            appUrl         起動引数：`--app`　につけるURLを指定します
'            cleanActive    `userProfile`で起動中のブラウザプロセスを落として起動させます
'            addArgs        追加の起動引数
'            KioskMode      Edgeのキオスクモードを使う際にセットします
'---------------------------------------------------------------------------------------------------
Public Sub StartAndConnectTab(Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional cleanActive As Boolean = True, _
    Optional userProfile As String, Optional addArgs As String, Optional KioskMode As edgeKioskType)
    Const FromProcedureName As String = ThisClassName & ".StartAndConnectTab"

    '1. ブラウザ起動
    Set BrowserCore = New CDPBrowser
    BrowserCore.start userProfile, Name, appUrl, cleanActive, addArgs, KioskMode

    '2. WithEventsClassオブジェクトを受け取る
    Set PipeCore = BrowserCore.InheritanceCDPCore

    '3. Complete initialization ------------------------------------------------------
    attachToSession     'Find and attach the newly open Session ID
    getSessionInfo      'Store initial session info for printParams debugging
    BrowserCore.serializeForMainTab brTab.sessionID, brTab.targetID
    '------------------------------------------------------------------------------

   'Make sure the page is fully loaded for inAppMode ------------------------------------------------------------------------------------
   '07/01/25: Long Vh: simplified the check as there are web apps that does not end up loading the target url but points to another url.
   'As a result, the below check simply checks if there is "movement" in the name of the page instead
    Dim timerStart As Double, strLoc As String
    If appUrl <> "" Then
        timerStart = TimerCounter
        Do
            If Me.Name <> "" Then Exit Do
            If TimerCounter - timerStart > timerOut Then Err.Raise CDPCustomErrorCodes.TIMEOUT, , "Timeout waiting for the browser to load the homepage."
            DoEvents
        Loop
       wait
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------

   'Finally print the session details for debugging ------------------------------
    printParams
   '------------------------------------------------------------------------------

End Sub

Private Sub serialize()
'------------------------------------------------------------------
' This method transcribes Chrome's Json message into a string.
' The serialized string can then be used to automate an active
' session with a known serial.
' Updated: 03/05/23 - Replace sessionID with targetID since it is
'                     constant and stable on every CDP session
'          22/01/26 - Added specification to user profile name to
'                     the cpid file name to allow targetting of
'                     specific profile for reattachment.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".serialize"


    '1. 必要なデータを`Dictionary`に詰める
    Dim BrowserInfo As New Dictionary
    BrowserInfo.Add "lngLastID", brTab.lngLastID
    BrowserInfo.Add "sessionID", brTab.sessionID
    BrowserInfo.Add "targetID", brTab.targetID

    '2. Excelのテーブルへ記録する
    Set ShSetting01_StartBrowser.serializeToTable(br.User, 2, FromProcedureName) = BrowserInfo
End Sub

'***************************************************************************************************
'* 機能　　：Excelのテーブルから、ブラウザのタブ情報を取得します
'---------------------------------------------------------------------------------------------------
'* 引数　　：reuseSession   `True` ：Excelに記録中のsessionIdを流用します
'                           `False`：`targetID`を基に、sessionIdを更新し、Excelに記録中の古いsessionIdの破棄も行い、上書きします
'***************************************************************************************************
Private Sub deserializeForTab(reuseSession As Boolean)
    Const FromProcedureName As String = ThisClassName & ".deserializeForTab"


    '1. Excelテーブルから、`UserName`に紐づくタブセッション情報を読み込む
    Dim objSerialize As Scripting.Dictionary
    Set objSerialize = ShSetting01_StartBrowser.deserializeFromTable(BrowserCore.CurrentBrowserUser, 2, FromProcedureName)

    '2. そもそもない場合は即抜け
    If objSerialize Is Nothing Then Exit Sub

    '3. タブのTargetID情報を適用
    brTab.targetID = objSerialize.Item("targetID")

    '4. SessionIDの適用方法に応じた分岐を行う
    If reuseSession And objSerialize.Item("sessionID") <> "" Then
        '既存のSessionIDを適用
        '4-1. 古いSessionIDを適用
        brTab.sessionID = objSerialize.Item("sessionID")
        printMsg info_, "I applied the SessionID(" & brTab.sessionID & ") from the Excel table.", FromProcedureName
    Else
        '古いSessionIDを破棄してから、新しいSessionIDを付与
        '4-1. 新しい`sessionId`を発行し、失敗時はエラーログを出して抜ける
        Dim newSessionID As String: newSessionID = BrowserCore.attachToTab(brTab.targetID)
        If newSessionID = "" Then printMsg ERROR_, "Failed to update sessionid with Targetid: " & brTab.targetID, FromProcedureName: Exit Sub

        '4-2. 一旦、古いSessionIDを適用
        brTab.sessionID = objSerialize.Item("sessionID")

        '4-3. 内部で、その古いSessionIDを破棄させ、新しいSessionIDを付与
        Me.id = newSessionID
    End If
End Sub

'***************************************************************************************************
'* 機能　　：記述したJavaScriptをブラウザ内で実行します。JavaScript版低レベル操作としての位置付けです。
'            Webスクレイピングに関わる引数は全て設けています。
'---------------------------------------------------------------------------------------------------
'* 返り値　：Variant型。実行したJavaScriptコード、引数:returnByValue や、IFEXCEPTION に依存します。
'* 引数　　：strString                      実行したいJavaScriptコードそのもの。`As Is`として実行されます。※ほとんどのエスケープ処理は、`WebJsonConverter`が担うため、ここでわざわざ手動エスケープの記述は必要ありません。
'
'            objectId                       実行の基準要素オブジェクトIDを指定します。これを指定すると`Runtime.callFunctionOn`として優先実行されます。
'
'            objectArguments                これを指定すると`Runtime.callFunctionOn`として優先実行されます。
'                                           予め、呼び出し側にてCollectionを組み立てて渡してください。
'
'            contextId                      `iframe`内で、JavaScript実行をしたいときに指定します。`objectId`が指定されてる場合は、無視されます
'            IFEXCEPTION                    JavaScript実行中に「例外（Exception）」が発生した場合の代替値を指定します。名前の通り、IFERROR関数チックに使えます。
'
'            returnByValue                  True ：type:"object" 等の場合でも、`value`で受け取れるようにCDPが試みます。実行コードによっては、Dictionary or Collection形式で返されるため、`SET`で受け取る必要があります。なお、空のオブジェクトで返る場合は、対応してないデータ型の可能性があるためその際は、`false`に切り替えてください。
'                                           False：type:"object" 等の場合は、`{"objectId": "-4985739485.1.3"}`にある`-4985739485.1.3`というブラウザ内変数名的なものとして返えります。再度`jsEval`を使う際の引数`objectId`に流用できます。「このJSコードは`SET`で...」という意識は不要です。
'
'            awaitPromise                   True ：JavaScriptそのものの実行結果が非同期処理であっても、処理が終わるまで待ってから結果を返すことができます
'                                           False：JavaScriptそのものの実行結果が非同期処理の場合、結果を待たずに続行されます
'
'            userGesture                    True ：「人間がクリックやキーボード操作をした結果」としてブラウザに偽装します。スクレピング対策が施されてる場合は、`True`にしておくといいかもしれません
'                                           False：特に施しません
'
'            allowUnsafeEvalBlockedByCSP    True ：サイトの外部JavaScript実行規制を無視して、全てのJavaScript実行を許可します
'                                           False：サイトの外部JavaScript実行規制に従います
'
'            serializationOptions           意地でも、type:"object" 等となっても、`value`で受け取れるように、細かい調整をします。予め、呼び出し側にてDictionaryを組み立てて渡してください。
'                                           このパラメーターがあると、`returnByValue`の引数は無視されます。
'
'            generatePreview                `objectId`で返った中身をプレーンテキスト化して、プレビューします。予め、ログレベルを`DEBUG`以下にしてください。基本は、開発時のみ`True`にしてください。
'
'            RunAsyncCDP                    True ：結果を待たないCDPで実行します。実行時のIDが返却されます。主に、「この操作、`alert`関数が発動するな...」という場面で使用します。
'                                           False：結果を待つCDPで実行します。基本はこれでOK
'
'            StopException                  `True`で、JavaScriptコード内で例外が起こった場合、Err.Raiseで止まります。基本は、開発時のみ使用する想定です
'
'            dbgMsg                         Allow the option to bypass error if intended
'---------------------------------------------------------------------------------------------------
'* 機能説明：This function commands the attached browser instance to execute a Javascript expression to perform required tasks.
'            Note: the original function has a bool retry variable which was removed for simplicity.
'            Added dbgMsg to read debug line on the go when needed.
'
'* 注意事項：・利用時は、「名前付き引数`xxx:=yyy`」として運用することをお勧めします
'            ・JavaScript実行中の例外確認は、「IsError関数」で判断してください。詳細の確認は、`LastJavaScriptException`プロパティで確認できます。「IsError関数」での確認が面倒なら、引数`IFEXCEPTION`を活用してください。
'            ・`returnByValue`=true の場合は、SET or LET で受け取る意識が必要になります
'
'            ・「このJavaScriptコードの場合は、こういう返り値が来るな」といった知識がないと、返り値に困惑します。
'               開発時は、デバッグログを活用した試行錯誤をすることをお勧めします。
'
'            ・`objectId`や`objectArguments`を指定する`Runtime.callFunctionOn`方式の場合は、`strString`を 「function() { ... }」 形式で記述する必要があります
'
'            ・実行したJavaScriptコードの結果が、プリミティブ型の場合、`returnByValue`の設定問わずそのプリミティブ型の中身が返ります。
'              常に`objectId`で返したい場合は、JavaScriptコードそのものを工夫する必要があります。
'
'            ・JavaScriptの場合、例外を利用した処理が多々あるため、運用時は`StopException`を省略しておくことを推奨します
'            ・CDP実行自体のエラーキャッチは常に`True`となります
'            ・`IFEXCEPTION`はあくまでも、「例外」時のみ機能します。JavaScript自体は成功してるけど、結果がエラーの場合はISERROR関数で対処する必要があります。
'---------------------------------------------------------------------------------------------------
'* 元URL   ：https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#method-evaluate
'***************************************************************************************************
Public Function jsEval(strString As String, Optional objectId As String, Optional contextId As Long, Optional IFEXCEPTION As Variant, _
                       Optional objectArguments As Collection, _
                       Optional returnByValue As Boolean, _
                       Optional awaitPromise As Boolean, _
                       Optional userGesture As Boolean, _
                       Optional allowUnsafeEvalBlockedByCSP As Boolean, _
                       Optional serializationOptions As Dictionary, _
                       Optional generatePreview As Boolean, _
                       Optional RunAsyncCDP As Boolean, _
                       Optional StopException As Boolean, _
                       Optional dbgMsg As Boolean = True) As Variant
    Const FromProcedureName As String = ThisClassName & ".jsEval"


    Dim CDPresults  As Scripting.Dictionary
    Dim strMessage  As String
    Dim objMessage  As New Scripting.Dictionary
    Dim CDPparams   As New Scripting.Dictionary

step1:

   'Construct message for sending ----------------------------------------------------------
    CDPparams.RemoveAll

    '`callFunctionOn`か`evaluate`を決める判定に入ります
    Dim UseJsMethod As String

    ' 【パターン1】特定の要素を操作する場合
    If StrPtr(objectId) Then
        UseJsMethod = "Runtime.callFunctionOn"

        '必要なパラメーターをセット
        CDPparams.Add "functionDeclaration", strString
        If Not objectArguments Is Nothing Then CDPparams.Add "arguments", objectArguments   '引数があったら適用しておく

        'objectId があるので、場所(ContextID)を指定する必要はない
        CDPparams.Add "objectId", objectId

    ' 【パターン2】要素指定はないが、安全に引数(arguments)を渡したい場合
    ElseIf Not objectArguments Is Nothing Then
        UseJsMethod = "Runtime.callFunctionOn"

        '必要なパラメーターをセット
        CDPparams.Add "functionDeclaration", strString
        CDPparams.Add "arguments", objectArguments

        ' objectId がないので、実行する「場所(executionContextId)」を教える必要がある
        If contextId Then CDPparams.Add "executionContextId", contextId

    ' 【パターン3】引数も要素指定もない、ただのコード実行
    Else
        ' 最もシンプルな evaluate を使う
        UseJsMethod = "Runtime.evaluate"

        '必要なパラメーターをセット
        CDPparams.Add "expression", strString
        If contextId Then CDPparams.Add "contextId", contextId
    End If

    '以降は、共通設定を施す
    '※文字数節約により、IF文をかます
    If returnByValue Then CDPparams.Add "returnByValue", returnByValue
    If awaitPromise Then CDPparams.Add "awaitPromise", awaitPromise
    If userGesture Then CDPparams.Add "userGesture", userGesture
    If allowUnsafeEvalBlockedByCSP Then CDPparams.Add "allowUnsafeEvalBlockedByCSP", allowUnsafeEvalBlockedByCSP
    If generatePreview Then CDPparams.Add "generatePreview", generatePreview
    If Not serializationOptions Is Nothing Then CDPparams.Add "serializationOptions", serializationOptions

    'CDP-Jsonを組み立て
    objMessage.RemoveAll
    objMessage.Add "method", UseJsMethod
    objMessage.Add "params", CDPparams

    strMessage = BrowserCore.jsConverter.ConvertToJson(objMessage)
   '----------------------------------------------------------------------------------------

    'Send and check CDPresult -------------------------------------------------------------------------------------------
    '引数に応じて、CDP自体を同期/非同期で実行し、分岐します
    Set CDPresults = SendMessage(strMessage, RunAsyncCDP, dbgMsg)
    If RunAsyncCDP Then
        '`WriteFile`のエラーをチェック
        If CDPresults.Items(0) Then
            printMsg ERROR_, "Failed to AsyncCDP-WriteFile JavaScript: " & strString & vbCrLf & vbCrLf & "=== WriteFileErrorDetail ===" & vbCrLf & "ErrorCode:" & CDPresults.Items(0) & " - " & ErrorDetail.GetMessage(CDPresults.Items(0)), FromProcedureName, doRaiseError:=True
        Else
            'ID返却して抜ける
            jsEval = CLng(CDPresults.keys(0))
        End If

        printMsg Debug_, "Async CDP JavaScript with WriteFile Result: " & CDPresults.Items(0) & " - " & ErrorDetail.GetMessage(CDPresults.Items(0), "kernel32") & " ID: " & CDPresults.keys(0), FromProcedureName
        Exit Function
    Else
        If PipeCore.ShowLogCDPresponse Then printMsg Trace_, "jsEval - ResultCDP - JsonString = " & BrowserCore.jsConverter.ConvertToJson(CDPresults), FromProcedureName
        If CDPresults Is Nothing Then printMsg ERROR_, "Unable to contact the Browser instance!", FromProcedureName, doRaiseError:=True
        If Not CDPresults.Exists("error") Then GoTo step2   'CDP自体のエラーがなければ、`step2`へ
    End If
    '-----------------------------------------------------------------------------------------------------------------

   'Handle errors ---------------------------------------------------------------------------------------------------
    Dim errMsg As String, retried As Boolean
    errMsg = CDPresults("error")("message")
    printMsg WARN_, "jsEval Error: " & errMsg, FromProcedureName
    If errMsg = "'" & UseJsMethod & "' wasn't found" Then printMsg WARN_, "Possibly the main session has been abruptly closed.", FromProcedureName
    If errMsg = "Session with given id not found." And Not retried Then
        printMsg info_, "Attempting to update Session ID " & brTab.sessionID, FromProcedureName
        Me.id = BrowserCore.attachToTab(brTab.targetID)
        printMsg info_, "New Session ID obtained: " & brTab.sessionID, FromProcedureName
        retried = True: GoTo step1
    End If

    '実行したJavaScriptコードの改行を消して、長すぎる場合はカットする（整形）
    Dim strLogValue As String
    strLogValue = LogFormat(strString, 100)
    printMsg ERROR_, "The Browser was unable to evaluate the JS expression: " & strLogValue & vbCrLf & vbCrLf & "<ErrorDetail>" & vbCrLf & errMsg, FromProcedureName, doRaiseError:=True
   '-----------------------------------------------------------------------------------------------------------------

step2:

    'JavaScript実行結果部分の解析フェーズに移ります ----------------------------------------------------------------
    Dim RemoteObject As Dictionary
    Set RemoteObject = CDPresults.Item("result").Item("result")
    If PipeCore.ShowLogCDPresponse Then printMsg Debug_, "jsEval - RemoteObject - JsonString = " & BrowserCore.jsConverter.ConvertToJson(RemoteObject), FromProcedureName

    '例外（JS側のエラー）が発生していないかチェックし、あったら、ログを表示し、抜けます
    If CDPresults.Item("result").Exists("exceptionDetails") Then
        '最後の詳細例外情報を記録
        Set lastExceptionDetailsJS = CDPresults("result")("exceptionDetails")

        '実行したJavaScriptコードの改行を消して、長すぎる場合はカットする（整形）
        strLogValue = LogFormat(strString, 100)

        '例外による`Err.Raise`で止めるか？
        If StopException Then
            '`StopException`が有効の場合は、ログ表示用に、詳細の方を文字列化
            Dim exceptionDetails As String: exceptionDetails = BrowserCore.jsConverter.ConvertToJson(lastExceptionDetailsJS)
            printMsg ERROR_, "An exception occurred while executing JavaScript: " & strLogValue & vbCrLf & vbCrLf & "<exceptionDetails>" & vbCrLf & exceptionDetails, FromProcedureName, doRaiseError:=True
        Else
            '`StopException`が無効の場合は、簡易部分の例外内容を出す
            printMsg WARN_, "ExceptionType: " & RemoteObject("className") & " , " & "ExceptionDescription: " & RemoteObject("description"), FromProcedureName

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
    If JSresultAnalysis(RemoteObject, serializationOptions, Not returnByValue, generatePreview, tmp) Then
        Set jsEval = tmp
    Else
        jsEval = tmp
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

    ' 最終的なデバッグ出力
    printMsg Debug_, "Script executed: " & strString & " | Returned: " & strLogValue, FromProcedureName
   '---------------------------------------------------------------------------------------------------------

End Function

'***************************************************************************************************
'* 機能　　：JavaScript実行で返ったJsonDictionaryから、結果値`value`だけを取り出すように試みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：RemoteObject                       `Runtime.RemoteObject`でのJson-Dictionaryを指定します。　https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#type-RemoteObject
'            PriorityReturn_deepSerializedValue `deepSerializedValue`があったら、そっちを最優先 に返すようにします。`serializationOptions`に依存します。
'            PriorityReturn_objectId            `objectId`           があったら、そっちを第2優先に返すようにします。`returnByValue`に依存します。
'            ReturnValue                        結果値。参照渡しによる実質返り値です。`primitive values` OR `Object` 両方が想定されます。
'
'* 返り値　：ReturnValueがObjectであるかの論理値。`ReturnValue`がObject型であるかを呼び出し側で正しく代入するために必要です。
'---------------------------------------------------------------------------------------------------
'* 注意事項：返す値は、`value` OR `objectId` OR `deepSerializedValue` に限ります。`description`等といった補助情報やエラー情報は返しません。
'***************************************************************************************************
Private Function JSresultAnalysis(RemoteObject As Dictionary, PriorityReturn_deepSerializedValue As Dictionary, PriorityReturn_objectId As Boolean, ObjectPreview As Boolean, ByRef ReturnValue As Variant) As Boolean
    Const FromProcedureName As String = ThisClassName & ".JSresultAnalysis"
    Const ErrorLandmark     As String = "error"


    '------------------------ 1. 特定項目有無チェック ------------------------
    '`serializationOptions`が指定されてたら、`deepSerializedValue`を最優先に返します
    If Not PriorityReturn_deepSerializedValue Is Nothing Then
        printMsg info_, "The parameter `serializationOptions` is specified. Check if `deepSerializedValue` exists.", FromProcedureName

        '`deepSerializedValue`があるか？
        If RemoteObject.Exists("deepSerializedValue") Then
            '見つかったので、`deepSerializedValue`を返して、ここで抜ける
            Set ReturnValue = RemoteObject("deepSerializedValue")
            printMsg info_, "JS-Return type: object-deepSerializedValue", FromProcedureName

            'Objectで返したのでその目印もつけておく
            JSresultAnalysis = True
            Exit Function
        Else
            printMsg WARN_, "`The parameter `serializationOptions` was specified, but `deepSerializedValue` was missing.", FromProcedureName
        End If
    End If

    '`objectId`を返す優先になってたら、こっちを返します
    If PriorityReturn_objectId Then
        printMsg Trace_, "Since `returnByValue` is `False`, we check for the existence of `objectId`.", FromProcedureName

        '`objectId`があるか？
        If RemoteObject.Exists("objectId") Then
            '見つかったので、`objectId`を返す
            ReturnValue = CStr(RemoteObject("objectId"))
            printMsg info_, "JS-Return type: objectId", FromProcedureName

            '返ったObjectID情報を列挙する
            Dim objectSubtype As String: objectSubtype = RemoteObject("subtype")
            printMsg Debug_, String(100, "="), FromProcedureName
            printMsg Debug_, "objectId information", FromProcedureName
            printMsg Debug_, String(100, "="), FromProcedureName
            printMsg Debug_, "objectId          : " & ReturnValue, FromProcedureName
            printMsg Debug_, "objectType        : " & RemoteObject("type"), FromProcedureName
            printMsg Debug_, "objectSubtype     : " & objectSubtype, FromProcedureName
            printMsg Debug_, "objectClassName   : " & RemoteObject("className"), FromProcedureName
            printMsg Debug_, "objectDescription : " & RemoteObject("description"), FromProcedureName
            If ObjectPreview Then printMsg Debug_, "objectPreview     : " & BrowserCore.jsConverter.ConvertToJson(RemoteObject("preview")), FromProcedureName
            printMsg Debug_, String(100, "="), FromProcedureName

            'エラー情報があればそのログを出しておく
            '※「try { ... } catch(e) { return e }」といった書き方をしないとこのログは表示されない想定
            If objectSubtype = ErrorLandmark Then printMsg WARN_, "This objectId(" & ReturnValue & ") contains an error. ErrorType: " & RemoteObject("className") & " , " & "ErrorDescription: " & RemoteObject("description"), FromProcedureName

            Exit Function
        Else
            'primitive values OR `object-null` OR `undefined` の可能性あり
            printMsg Debug_, "`objectId` was not found. It may be a primitive value or `object-null` or `undefined`, so we will try to retrieve it.", FromProcedureName
        End If
    Else
        printMsg Trace_, "`returnByValue` is `True`. Attempting to retrieve the primitive value.", FromProcedureName
    End If


    '------------------------ 2. primitive values OR JsonDictionaryの取得を試みる ------------------------
    'primitive values の可能性があるので、それの取得を試みます
    '以降は基本的に、`returnByValue:=True`でないと通りません
    'ObjectIDとして返そうとしたが`object-null`,`undefined`の場合もここを通って、返却します

    ' 格納型情報を取得
    Dim ValueType       As String: ValueType = RemoteObject.Item("type")
    Dim ExistsSubtype   As Boolean: ExistsSubtype = RemoteObject.Exists("subtype")
    Dim ValueSubtype    As String: ValueSubtype = RemoteObject.Item("subtype")
    printMsg Debug_, "Type info. type: " & ValueType & " , subtype: " & ValueSubtype, FromProcedureName

    '型に応じた分岐処理
    Select Case ValueType
        Case "object"
            ' object型の場合はsubtypeで更に細かく分岐
            Select Case ValueSubtype
                Case "null"
                    ReturnValue = Null
                    printMsg info_, "JS-Return type: object-null", FromProcedureName

                Case "array"
                    ' 配列やDOM要素(returnByValue時)は、valueの中に中身が入っている
                    ' 配列の場合は Scripting.Dictionary (連番) が入ります
                    '※誇らしげに説明してますが、`returnByValue:=True`で、`array`だと、subtype そのものがなくなります...が一応書いておきます。
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-array", FromProcedureName

                Case "node"
                    '多分、来ない...が一応書いておきます。
                    '`returnByValue:=True`で、`node`も、subtype そのものがなくなります...
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-node", FromProcedureName

                Case "regexp"
                    '以下同文
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-regexp", FromProcedureName

                Case "map"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-map", FromProcedureName

                Case "date"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-date", FromProcedureName

                Case "set"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-set", FromProcedureName

                Case "weakmap"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-weakmap", FromProcedureName

                Case "weakset"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-weakset", FromProcedureName

                Case "iterator"     '※ObjectIDモードでも空欄です?!
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-iterator", FromProcedureName
                
                Case "generator"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-generator", FromProcedureName

                Case "promise"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-promise", FromProcedureName

                Case "proxy"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-proxy", FromProcedureName

                Case "typedarray"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-typedarray", FromProcedureName

                Case "arraybuffer"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-arraybuffer", FromProcedureName

                Case "dataview"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-dataview", FromProcedureName

                Case "webassemblymemory"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-webassemblymemory", FromProcedureName

                Case "trustedtype"  '※ObjectIDモードでも空欄です?!
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-trustedtype", FromProcedureName

                Case ErrorLandmark
                    '`returnByValue:=True`では意味をなさないようです...が念のため
                    printMsg WARN_, "This value contains an error. ErrorType: " & RemoteObject("className") & " , " & "ErrorDescription: " & RemoteObject("description"), FromProcedureName
                    ReturnValue = CVErr(xlErrValue)
                    printMsg info_, "JS-Return type: object-error", FromProcedureName

                Case Else
                    '"subtype"自体がない場合は、`array`として扱います
                    '`array`,`typedarray`が機能してます。他は空({})です
                    If Not ExistsSubtype Then
                        ' 配列やDOM要素(returnByValue時)は、valueの中に中身が入っている
                        ' 配列の場合は Scripting.Dictionary (連番) が入ります
                        ' returnByValue=True なら value に Dictionary が入っているはず
                        printMsg WARN_, "It's of type Object, but no `subtype` was found. It will be treated as an `array`.", FromProcedureName
                        Set ReturnValue = RemoteObject.Item("value")
                        printMsg info_, "JS-Return type: object-array", FromProcedureName

                    Else
                        'ドキュメント化されてない型にも備えておく
                        printMsg WARN_, "This data type cannot be handled in VBA. type: " & ValueType & " , subtype: " & ValueSubtype, FromProcedureName
    
                        '返せる物があれば、返しておく
                        If RemoteObject.Exists("value") Then
                            ' オブジェクト(Dictionary)を返す場合は Set が必要
                            If IsObject(RemoteObject.Item("value")) Then
                                Set ReturnValue = RemoteObject.Item("value")
                            Else
                                ReturnValue = RemoteObject.Item("value")
                            End If
                        ElseIf RemoteObject.Exists("unserializableValue") Then
                            ReturnValue = CStr(RemoteObject.Item("unserializableValue"))
                        Else
                            ReturnValue = RemoteObject.Item("description")
                        End If
                    End If
            End Select

        Case "function"
            ' 関数が返ってきた場合は、「{"type":"function","value":{}}」なので、`SET`
            Set ReturnValue = RemoteObject.Item("value")
            printMsg info_, "JS-Return type: function", FromProcedureName

        Case "undefined"
            ' JSの undefined は VBAでは Empty が一番近い
            ReturnValue = Empty
            printMsg info_, "JS-Return type: undefined", FromProcedureName

        Case "string"
            ' 文字列として認識するように返す
            ReturnValue = CStr(RemoteObject.Item("value"))
            printMsg info_, "JS-Return type: string", FromProcedureName

        Case "number"
            'VBAで扱える型か？
            If RemoteObject.Exists("unserializableValue") Then
                'VBAで扱えない特殊な"number"の場合は、警告を出し、「#NUM!」で返します。
                printMsg WARN_, "This `number` can't be processed in VBA. It returns `#NUM!` instead. Actual JS result: " & RemoteObject("unserializableValue") & " , description: " & RemoteObject("description"), FromProcedureName
                ReturnValue = CVErr(xlErrNum)
            Else
                ' Double型として認識するように返す
                ' ※JavaScriptは、整数専用等の型がないため
                ReturnValue = CDbl(RemoteObject.Item("value"))
            End If

            printMsg info_, "JS-Return type: number", FromProcedureName

        Case "boolean"
            ' boolean型として認識するように返す
            ReturnValue = CBool(RemoteObject.Item("value"))
            printMsg info_, "JS-Return type: boolean", FromProcedureName

        Case "symbol"
            '※`symbol`は、`returnByValue:=True`にすると、「Object couldn't be returned by value.」というCDPエラーで返るので通らないと思うが念のため。
            ReturnValue = RemoteObject.Item("value")
            printMsg info_, "JS-Return type: symbol", FromProcedureName

        Case "bigint"
            ' BigInt は JSON に変換できない巨大整数のため、unserializableValue の中にある値をそのまんま返却します
            ReturnValue = CStr(RemoteObject.Item("unserializableValue"))
            printMsg WARN_, "Unserializable BigInt detected. Returning as String: " & ReturnValue & " , description: " & RemoteObject("description"), FromProcedureName
            printMsg info_, "JS-Return type: bigint", FromProcedureName

        Case Else
            'ドキュメント化されてない型にも備えておく
            printMsg WARN_, "This data type cannot be handled in VBA. type: " & ValueType, FromProcedureName

            '返せる物があれば、返しておく
            If RemoteObject.Exists("value") Then
                ' オブジェクト(Dictionary)を返す場合は Set が必要
                If IsObject(RemoteObject.Item("value")) Then
                    Set ReturnValue = RemoteObject.Item("value")
                Else
                    ReturnValue = RemoteObject.Item("value")
                End If
            ElseIf RemoteObject.Exists("unserializableValue") Then
                ReturnValue = CStr(RemoteObject.Item("unserializableValue"))
            Else
                ReturnValue = RemoteObject.Item("description")
            End If
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

Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
'-----------------------------------------------------------------------
' This sub is a quick wrapper to navigate to an url.
' Updated 23/03/23
'-----------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".navigate"


   'Check if already navigated
    Dim curURL As String
    curURL = jsEval("window.location.href", dbgMsg:=False)
    If curURL = checkURI(strURL) Then
        printMsg WARN_, "Already navigated to """ & strURL & """", FromProcedureName
        Exit Sub
    End If

   'If not yet then do navigate
    Dim results As Scripting.Dictionary
    Set results = SendMessage("{""method"":""Page.navigate"",""params"":{""url"":""" & strURL & """}}")

    'should also add a maxTimeLimit to Wait
    wait till:=till
    getSessionInfo 'update new session info for printParams
    printMsg info_, "Navigated to: " & strURL, FromProcedureName

End Sub


Private Function isLiveTab() As Boolean
'-----------------------------------------------------------------------------------
' Check if there is ongoing connection with an attached instance
' Original Name: connectionAlive
'-----------------------------------------------------------------------------------

    On Error GoTo iferr
    Dim strLoc As String
    strLoc = jsEval("window.location.href", dbgMsg:=False)
    isLiveTab = True
    Exit Function

iferr:
    isLiveTab = False

End Function

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
    Select Case till
        Case isLoading: targetState = "loading"
        Case isInteractive: targetState = "interactive"
        Case isComplete: targetState = "complete"
    End Select

    '初期値を設定しておく
    '最初はほぼ必ずloadingだから、`targetState`が`loading`になってたら、わざわざチェックする必要性はない
    strState = "loading"

    Do Until strState = targetState
        strState = jsEval("document.readyState", dbgMsg:=False)                 'dbgMsg = False to not spam Immediate Window
        printMsg Debug_, "Current ReadyState: " & strState, FromProcedureName   'option to print ReadyState info during iteration
        If targetState = "interactive" And strState = "complete" Then Exit Do   'also exit if better than interactive is satisfied
        sleep 0.1                                                               'reduce sleep will speed up but will cost cpu power
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

    If br.hWnd = 0 Then br.hWnd = getHandle
    ShowWindow br.hWnd, doShow 'Make visible
    ShowWindow br.hWnd, state

   'Resize to specific position and size
   'Will not work properly if window is currently maximized
    If xywh <> "0 0 0 0" Then

        sleep 0.3      'Sleep a bit for previous Win Api executions to finish, else the resize effect will be accidentally cancelled out

        Dim params As New Scripting.Dictionary
        If br.WID = "" Then br.WID = ExecuteCDP("Browser.getWindowForTarget").Item("windowId")
        params("windowId") = Int(br.WID)     '02/03/23 CInt() will throw error, Int() is better

        Dim bounds As New Scripting.Dictionary
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

Public Sub hide()
'------------------------------------------------------------------
' This is the equivalent of show "hidden" but this function is
' added to make it easy to access and use the same feature.
'------------------------------------------------------------------

    If br.hWnd = 0 Then br.hWnd = getHandle
    ShowWindow br.hWnd, asHidden

End Sub

Public Function bringToForeground()
'------------------------------------------------------------------
' Brings the browser window to the foreground
' Updated: Long Vh - 06/01/26 should be a function like others
'------------------------------------------------------------------

    If br.hWnd = 0 Then br.hWnd = getHandle
    If br.hWnd <> 0 Then
        ShowWindow br.hWnd, doShow
        BringWindowToTop br.hWnd
        SetForegroundWindow br.hWnd
    End If

End Function

Public Sub closeTab()
'------------------------------------------------------------------
' Close the current tab of the session. Requires targetId which can
' be retrieved by using getTargetId. This is the targetId of the
' active tab under focus of the automation pipe.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".closeTab"


   'Get the original serialized tab ID
    Dim mainID As String, serial As Dictionary
    Set serial = ShSetting01_StartBrowser.deserializeFromTable(BrowserCore.CurrentBrowserUser, 2, FromProcedureName)
    mainID = serial.Item("targetID")

    'メインタブは閉じないように仕向ける
    If brTab.targetID = mainID Then
        printMsg WARN_, "Unable to close this tab " & brTab.targetID & " as it is the main session tab", FromProcedureName
        Exit Sub
    End If

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    ExecuteCDP "Target.closeTarget", params
    printMsg info_, brTab.targetID & " was closed", FromProcedureName

End Sub

Public Function getElementByID(strID As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class from the element ID, similar to the getElementByID of JS.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Get the object (no global variable assignment - objectId only)
    Dim varJS As String: Dim result As Variant
    varJS = "document.getElementById(""" & strID & """)"
    result = jsEval(varJS, returnByValue:=False)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByID = New CDPElement
    getElementByID.Init Me, varJS, VarType(result), oid

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

   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")

   'Get the object (no global variable assignment - objectId only)
    Dim varJS As String: Dim result As Variant
    varJS = "document.evaluate(""" & strXPath & """, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
    result = jsEval(varJS, returnByValue:=False)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByXPath = New CDPElement
    getElementByXPath.Init Me, varJS, VarType(result), oid

End Function

Public Function getElementsByXPath(strXPath As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by XPath. For further info on the particular XPath
' syntax used for retrieving the elements, refer to Mozilla doc below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")

   'Get the XPathResult objectId (ORDERED_NODE_SNAPSHOT_TYPE)
    Dim collResult As Variant
    collResult = jsEval("document.evaluate(""" & strXPath & """, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)", returnByValue:=False)
    If VarType(collResult) <> vbString Then Exit Function
    Dim collOid As String: collOid = CStr(collResult)

   'Get snapshotLength via callFunctionOn
    Dim elements As Collection: Set elements = New Collection
    Dim length_ As Long, N As Long
    length_ = jsEval("function(){ return this.snapshotLength }", objectId:=collOid)

   'Get each item's objectId via callFunctionOn
    Dim tElement As CDPElement
    Dim itemArgs As Collection, itemArg As Scripting.Dictionary
    Dim itemResult As Variant, itemOid As String
    For N = 0 To length_ - 1
        Set itemArgs = New Collection
        Set itemArg = New Scripting.Dictionary
        itemArg.Add "value", N
        itemArgs.Add itemArg
        itemResult = jsEval("function(n){ return this.snapshotItem(n) }", objectId:=collOid, objectArguments:=itemArgs, returnByValue:=False)
        itemOid = "": If VarType(itemResult) = vbString Then itemOid = CStr(itemResult)
        Set tElement = New CDPElement
        tElement.Init Me, , VarType(itemResult), itemOid
        elements.Add tElement
    Next

    Set getElementsByXPath = elements

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

   'Get the object (no global variable assignment - objectId only)
    Dim varJS As String: Dim result As Variant
    varJS = "document.querySelector(""" & strQuery & """)"
    result = jsEval(varJS, returnByValue:=False)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByQuery = New CDPElement
    getElementByQuery.Init Me, varJS, VarType(result), oid

End Function

Public Function getElementsByQuery(strQuery As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by QuerySelectorAll. The syntax is as below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Get the NodeList objectId
    Dim collResult As Variant
    collResult = jsEval("document.querySelectorAll(""" & strQuery & """)", returnByValue:=False)
    If VarType(collResult) <> vbString Then Exit Function
    Dim collOid As String: collOid = CStr(collResult)

   'Get length via callFunctionOn
    Dim result As Variant, elements As Collection: Set elements = New Collection
    Dim length_ As Long, N As Long
    length_ = jsEval("function(){ return this.length }", objectId:=collOid)

   'Get each item's objectId via callFunctionOn
    Dim tElement As CDPElement
    Dim itemArgs As Collection, itemArg As Scripting.Dictionary
    Dim itemResult As Variant, itemOid As String
    For N = 0 To length_ - 1
        Set itemArgs = New Collection
        Set itemArg = New Scripting.Dictionary
        itemArg.Add "value", N
        itemArgs.Add itemArg
        itemResult = jsEval("function(n){ return this.item(n) }", objectId:=collOid, objectArguments:=itemArgs, returnByValue:=False)
        itemOid = "": If VarType(itemResult) = vbString Then itemOid = CStr(itemResult)
        Set tElement = New CDPElement
        tElement.Init Me, , VarType(itemResult), itemOid
        elements.Add tElement
    Next

    Set getElementsByQuery = elements

End Function

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

Public Sub jsAddScript(strLocalScriptPath As String)
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


    If Dir(strLocalScriptPath) = "" Then Exit Sub

    Dim fso As New FileSystemObject
    jsEval fso.OpenTextFile(strLocalScriptPath).ReadAll 'Load the script contents
    printMsg info_, "Successfully added the local script: """ & strLocalScriptPath & """", FromProcedureName

    Set fso = Nothing

End Sub

'***************************************************************************************************
'* 機能　　：Excelテーブルにある既存のパイプハンドル、ブラウザセッション情報を利用して、再接続を試みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功可否論理値
'* 引数　　：reuseSession   `True` ：Excelに記録中のsessionIdを流用します
'                           `False`：`targetID`を基に、sessionIdを更新し、Excelに記録中の古いsessionIdの破棄も行い、上書きします
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
'            　・newTab
'            　・getTab
'
'            ・非同期イベント情報を購読中の場合(例：Network.webSocketFrameReceived)は、`reuseSession:=True`推奨です
'***************************************************************************************************
Public Function reattach(userProfile As String, Optional reuseSession As Boolean) As Boolean
    Const FromProcedureName As String = ThisClassName & ".reattach"


    '1. 新オブジェクトを用意
    Set BrowserCore = New CDPBrowser

    '2. 識別名称に紐づくパイプハンドル情報を再セットし、パイプ生存確認を行う
    reattach = BrowserCore.reattach(userProfile)
    Set PipeCore = BrowserCore.InheritanceCDPCore

    '3. パイプが生きてない場合は、VBAエラーで停止させる
    If Not reattach Then printMsg ERROR_, "Failed to reattach to previous Pipe handle. user-data-dir=" & userProfile, FromProcedureName, doRaiseError:=True

    '4. タブ情報を再セット
    deserializeForTab reuseSession

    '5. `sessionId`が空欄の場合は、即抜け
    If brTab.sessionID = "" Then reattach = False: printMsg WARN_, "Failed to reattach to previous CPID.", FromProcedureName: Exit Function

    '6. sessionId生存check
    If isLiveTab Then
        reattach = True

        'ステータス更新
        getSessionInfo
        printParams
    Else
        reattach = False
        printMsg WARN_, "Failed to reattach to previous CPID.", FromProcedureName
    End If

End Function

Public Sub notify(msg)
'----------------------------------------------------------------------------------
' Send a custom notification message on the bottom right corner of the screen. This
' popup message is nicer than alert() and it is not modal: ie. does not the user to
' interact with it while pausing the automation.
' Credit: Claude AI
' Updated: 01/12/2025
'----------------------------------------------------------------------------------

   'Make sure backslash if any is escaped
    msg = Replace(msg, "\", "\\")

    jsEval "function showStatus(msg, duration = 10000) {" & _
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

    jsEval "showStatus('" & msg & "')"

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


    Dim params As New Scripting.Dictionary
    Dim r As Scripting.Dictionary
    Dim b64 As String, tmpPng As String

   'Verify if the page can be captured with screenshot
    Set r = ExecuteCDP("Page.getLayoutMetrics")
    If ExecuteError(r) Then
        printMsg WARN_, "Unable to perform snapPage on this screen as it has no layout metrics.", FromProcedureName
        Exit Sub
    End If

   'Perform the snip
    If getFullPage Then params.Add "captureBeyondViewport", True
    Set r = ExecuteCDP("Page.captureScreenshot", params)

   'Save to the same folder if not indicated
    If Not ExecuteError(r) Then
        b64 = r("data")
        SaveBase64ToFile b64, FolderPath, FileName
        printMsg info_, "Screenshot saved to: " & FolderPath & "\" & FileName, FromProcedureName
    End If

End Sub

Public Sub printParams()
'------------------------------------------------------------------
' Print all parameters of the current session for debuggings.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".printParams"


    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, br.LogID & " Tab - Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "Session ID        | sessionID    | " & brTab.sessionID, FromProcedureName
    printMsg Debug_, "Window Handle     | brHWnd       | " & br.hWnd, FromProcedureName
    printMsg Debug_, "Window ID         | brWId        | " & br.WID, FromProcedureName
    printMsg Debug_, "Tab URL           | brUrl        | " & br.Url, FromProcedureName
    printMsg Debug_, "Tab Title         | brName       | " & br.Name, FromProcedureName
    printMsg Debug_, "Tab ID            | targetID     | " & brTab.targetID, FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

End Sub

Public Sub activate(Optional focus As Boolean = True)
'----------------------------------------------------------------------------------
' Switch focus to the current tab being attached to under the current BRID.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    If focus Then ExecuteCDP "Target.activateTarget", params, True

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

    Dim params As New Scripting.Dictionary

    params("accept") = CBool(Accept)
    If StrPtr(promptText) Then params("promptText") = promptText
    ExecuteCDP "Page.handleJavaScriptDialog", params

End Sub

Public Sub pageEnable(Optional enable As Boolean = True)
'----------------------------------------------------------------------------------
' Enable the page for events listening. Required for handleDialog for instance.
'----------------------------------------------------------------------------------

    ExecuteCDP IIf(enable, "Page.enable", "Page.disable")

End Sub

'***************************************************************************************************
'* 機能　　：CDPのMethod名と、parametersを渡して、低レベルにてブラウザ制御をします
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド の実行結果(Json-Dictionary 形式)。失敗時は、`Nothing`となります。
'* 引数　　：methodName             method名。"Network.getCookies"や"DOM.getDocument"というのを記入します
'            params                 Json-Dictionary形式で、KeyとItemを事前に定義したオブジェクト変数を指定します。
'            alwaysBrowserContext   `True` ：`sessionID`の有無問わず、常にブラウザ自身を対象したコマンド実行モードとなります
'                                   `False`：`sessionID`があれば、その`SessionID`に対するコマンド実行。なければ、ブラウザ自身を対象したコマンド実行となります。
'
'            StopError              `True`で、実行失敗時に`Err.Raise`で停止し、以降の処理を止めます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：This function enables advanced method invocation under CDP.
'            For a list of all methods available, see:
'               https://chromedevtools.github.io/devtools-protocol/tot/Browser/
'            Error message will be prompted automatically by ErrorCheckCDP in the Immediate Window.
'            The return object is Json String that has been converted to a dictionary object.
'            Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'
'* 更新履歴：・18/02/26 - SessionIDがあっても、BrowserContext優先にコマンド実行できるように改良。エラーで止めるかのスイッチングを導入
'            ・28/04/23 - Set this to Public as it is useful
'            ・20/03/23 - Added printMsg msgs & set to Private
'
'* 注意事項：・CDPの運営思想の都合上、多数の`Deprecated`,`Experimental`が存在します。
'              これらのタグのmethodを使用する際は、`StopError`を`True`にすることを推奨します。失敗時の以降の処理の予期せぬデータ破損などを防ぐためです。
'
'            ・`StopError`の規定値は、`True`なので、特別な理由がない限りは、省略推奨です
'            ・このメソッドは、コマンド結果が来るまで待機するメソッドです。送るコマンド内容や、サイトの挙動によっては、タイムアウトエラーになる場合があります。
'              その場合は、後述の結果を待たないタイプのメソッド`ExecuteCDPAsync`をご利用ください。
'***************************************************************************************************
Public Function ExecuteCDP(methodName As String, Optional params As Scripting.Dictionary, Optional StopCDPError As Boolean = True) As Scripting.Dictionary
    Const FromProcedureName As String = ThisClassName & ".ExecuteCDP"


    Dim paramStr As String
    Dim results As Scripting.Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = BrowserCore.jsConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}", dbgMsg:=StopCDPError)

   'Print error if there is
    If ExecuteError(results) Then
        'エラーによる`Err.Raise`で止めるか？
        If StopCDPError Then
            'エラーになったmethod名と、結果生responseを表示して停止
            printMsg ERROR_, "Failed to invoke method """ & methodName & """" & vbCrLf & vbCrLf & "=== RawErrorResponse ===" & vbCrLf & BrowserCore.jsConverter.ConvertToJson(results), FromProcedureName, doRaiseError:=True
        Else
            'エラーになったmethod名だけログ出して続行
            printMsg WARN_, "Failed to invoke method """ & methodName & """", FromProcedureName
        End If
    Else
        'Return result object
        If results.Exists("result") Then Set ExecuteCDP = results.Item("result")
    End If

    'デバッグ用に、As Is 出力
    If PipeCore.ShowLogCDPresponse Then printMsg Debug_, "Executed """ & methodName & """ with result: " & BrowserCore.jsConverter.ConvertToJson(results), FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`ExecuteCDP`の非同期実行(コマンド結果を待たない)バージョンです
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド 実行時のID値(整理券番号のイメージでOK)
'* 引数　　：※`ExecuteCDP`と一緒のため、省略
'***************************************************************************************************
Public Function ExecuteCDPAsync(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As Long
    Const FromProcedureName As String = ThisClassName & ".ExecuteCDPAsync"


    Dim paramStr As String
    Dim results As Scripting.Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = BrowserCore.jsConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}", True, StopError)

    '非同期の場合は、`WriteFile`のエラーをチェックするだけとなります
    '※1件しかAddしてないので、直Index指定する
    If results.Items(0) Then
        'エラーによる`Err.Raise`で止めるか？
        If StopError Then
            'エラーになったmethod名と、`WriteFile`の実行エラーコードを表示して停止
            printMsg ERROR_, "Failed to invokeAsync-WriteFile method """ & methodName & """" & vbCrLf & vbCrLf & "=== WriteFileErrorDetail ===" & vbCrLf & "ErrorCode:" & results.Items(0) & " - " & ErrorDetail.GetMessage(results.Items(0)), FromProcedureName, doRaiseError:=True
        Else
            'エラーになったmethod名だけログ出して続行
            printMsg WARN_, "Failed to invokeAsync-WriteFile method """ & methodName & """　ErrorCode:" & results.Items(0) & " - " & ErrorDetail.GetMessage(results.Items(0), "kernel32"), FromProcedureName
        End If
    End If

    '実行したIDを返却
    ExecuteCDPAsync = results.keys(0)
    printMsg Debug_, "Async executed """ & methodName & """ with WriteFile Result: " & results.Items(0) & " - " & ErrorDetail.GetMessage(results.Items(0), "kernel32") & " ID: " & results.keys(0), FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：CDP にメッセージを送信し、Chromium ブラウザでメッセージが受信されているかどうかを確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：Async が、`False` → CDP-Jsonコマンド の実行結果(Json-Dictionary 形式)
'                      `True`  → 実行したコマンドIDと、WriteFile結果コードのペアーDictionary
'
'* 引数　　：strMessage             Chromium ブラウザに送るCDP-Jsonコマンド(プレーンテキスト)
'
'            Async                  `True` ：コマンドの結果を待ちません。Do Loopフェーズには入りません。
'                                   `False`：コマンドの結果を待ちます。Do Loopフェーズに入り、そのIDが来るまで待ちます。この場合、非同期に実行した結果のID回収もついでに行います
'
'            dbgMsg                 `True` ：`CDPCore.cls - TakeEvents`にて、エラーがあったら止まります
'                                   `False`：`CDPCore.cls - TakeEvents`にて、エラーがあっても止まらず、`SendMessage`は`nothing`として返ります
'---------------------------------------------------------------------------------------------------
'* 更新履歴：・03/05/23 - Added soft-reset on peeknamepipe error
'            ・03/22/23 - Added error handling for better debugging
'            ・24/10/25 Daniel Polak - Added soft re-attach attempt when pipe error
'---------------------------------------------------------------------------------------------------
'* 注意事項：引数-Async の値によって、Dictionaryの構造が変わる特殊なメソッドです。別途公開メソッドで、役割分担しております
'***************************************************************************************************
Private Function SendMessage(strMessage As String, Optional Async As Boolean, Optional dbgMsg As Boolean = True) As Scripting.Dictionary
    Const FromProcedureName As String = ThisClassName & ".SendMessage"


    Dim intRes As Long
    Dim strRes As String
    Dim results As Scripting.Dictionary       'to store a collection of IDs obtained from CDP

   '--------------------------------------------------
   'Send message to the browser
   '--------------------------------------------------
    printMsg Debug_, "runCDP  - Async       = " & Async, FromProcedureName

    Dim WriteResult As Long: WriteResult = PipeCore.ReadyRunCDP(strMessage, brTab.sessionID)    'Send message to the browser instance
    Dim currentCommandID As Long: currentCommandID = PipeCore.LastCommandID                     '`ReadyRunCDP`直後のcommandIDを取得
    
    printMsg Debug_, "Result  - WriteFile: " & WriteResult, FromProcedureName

    '非同期(結果はいらない)モードなら、整理券(ID)を渡してここで抜けます
    If Async Then
        Set SendMessage = New Dictionary

        '実行したコマンドIDと、WriteFileの結果コードを渡して、抜けます
        SendMessage.Add currentCommandID, WriteResult
        Exit Function
    End If

   '------------------------------------------------
   'Read message returned by browser
   '------------------------------------------------
    'CDP-Jsonコマンドを送った際のIDが返ってくるまで、ループ受信処理を行いますが、タイムアウトになったら、抜けます
    Dim timerStart As Double: timerStart = TimerCounter
    Dim tmp As String
    printMsg Trace_, "Starting Processes incoming messages from the browser...", FromProcedureName
    Do
        '1. 受信処理によるパイプを拝見する
        TakeEvents StopPipeError:=dbgMsg

        '2. 受信処理にてエラーがあったら、即抜け
        If PipeCore.LastErrorPeekNamedPipe > 0 Or PipeCore.LastErrorReadFile > 0 Then printMsg ERROR_, "Possible reasons: (1) the browser is not actually running (ie. call string failure); (2) the browser window was accidentally closed; (3) the CDP Pipe connection to the browser was cut-off somehow.", FromProcedureName: Exit Function

        '3. 内部の`RaiseEvent CDPBrowserID`による発火で、Dictionaryへ蓄積されたか？
        tmp = ResultCDPFromWithEvents(currentCommandID)
        If StrPtr(tmp) Then
            '送ったコマンド結果であることを確認できたのでJsonパース処理して結果を返す
            Set SendMessage = BrowserCore.jsConverter.ParseJson(tmp)
            printMsg info_, "Tab have confirmed receipt of command result ID: " & currentCommandID, FromProcedureName

            'ループから抜ける
            Exit Do
        Else
            printMsg Trace_, "Tab have not yet been able to confirm command ID: " & currentCommandID, FromProcedureName
        End If

        '4. 一定時間経ってもコマンド結果が来ないなら、エラーで停止します
        If TimerCounter - timerStart > timerOut Then Err.Raise CDPCustomErrorCodes.TIMEOUT, FromProcedureName, "Timeout waiting for browser to return message." Else printMsg Trace_, (TimerCounter - timerStart) / 1000 & " seconds have passed...", FromProcedureName
        DoEvents
    Loop

    printMsg Trace_, "Command result ID:" & currentCommandID & " was returned from CDP-Json, so the receiving loop will end.", FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：CDP-Jsonコマンド実行結果のエラーチェックを行い、最後のエラー記録も担います
'---------------------------------------------------------------------------------------------------
'* 返り値　：エラー有無論理値
'* 引数　　：result         パース処理済みのCDP-Jsonコマンド実行結果
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
Private Function ExecuteError(result As Scripting.Dictionary, Optional ErrorCode As Long, Optional ErrorMessage As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ExecuteError"


    Dim ErrorData As Variant

    'データの状態をチェック
    If result Is Nothing Then
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

    ElseIf result.Exists("error") Then
        'CDPコマンドエラー
        ExecuteError = True

        '`error`オブジェクトの中身で記録
        Set lastErrorDetailsCDP = result("error")

        '主要なエラー情報を抜き出す
        ErrorCode = lastErrorDetailsCDP("code")
        ErrorMessage = lastErrorDetailsCDP("message")
        printMsg WARN_, "ExecuteCDP Error: [" & ErrorCode & "] " & ErrorMessage, FromProcedureName

        '追加情報がある場合は、それもログ化
        If lastErrorDetailsCDP.Exists("data") Then
            ErrorData = lastErrorDetailsCDP("data")
            printMsg WARN_, "       " & ErrorData, FromProcedureName
        Else
            ErrorData = vbEmpty ' vbNullString? is lastErrorDetailsCDP("data") a string or object?
        End If
    Else
        'CDPコマンド正常終了
        ExecuteError = False

        'クリアしておく
        ErrorCode = 0
        ErrorMessage = vbNullString
        ErrorData = vbEmpty
    End If

End Function



'=============================================================================
' PUBLIC PROPERTIES
'=============================================================================
Property Get Url() As String
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    br.Url = jsEval("document.URL", dbgMsg:=False)
    Url = br.Url

End Property

Property Let Url(newURL As String)
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    navigate newURL
    br.Url = newURL

End Property

Property Get Name() As String
'------------------------------------------------------------
' Get the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    br.Name = jsEval("document.title", dbgMsg:=False)
    Name = br.Name

End Property

Property Let Name(newName As String)
'------------------------------------------------------------
' Set the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".Name"


    jsEval "document.title = '" & newName & "'", dbgMsg:=False
    printMsg info_, "Tab name changed to " & newName, FromProcedureName
    br.Name = newName

End Property

Property Get id() As String
'------------------------------------------------------------
' Get the session ID of the current browser
' Added 02/03/2023
'------------------------------------------------------------

    id = brTab.sessionID

End Property

Property Let id(newID As String)
'------------------------------------------------------------
' Set the session ID of the current browser to a new value
' Added 02/03/2023
'------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".id"


    '再設定されるので、元のSessionIDがあれば、破棄させる
    If StrPtr(brTab.sessionID) Then
        printMsg info_, "A property was called that receives a new SessionID (" & newID & "). The SessionID (" & brTab.sessionID & ") held by this class will now be discarded.", FromProcedureName
        BrowserCore.DiscardSessionID brTab.sessionID
    End If

    '渡されたIDで、再設定
    brTab.sessionID = newID

    '新SessionIDを記録しておく
    ShSetting01_StartBrowser.serializeSingleToBrowserInfoTable(BrowserCore.CurrentBrowserUser, "sessionID", FromProcedureName) = brTab.sessionID

    getSessionInfo        'update all session infos
    printMsg info_, br.LogID & " updated Session ID to " & brTab.sessionID, FromProcedureName

End Property

Property Get html() As String
'------------------------------------------------------------
' Get the html text of the entire document
' Added 02/03/2023
'------------------------------------------------------------

    html = jsEval("document.documentElement.innerHTML")

End Property

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
    If EventName = "" Then
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
'* 機能　　：このクラスで定義中のtargetIDを取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：「どのタブか」を特定するための専用ID
'***************************************************************************************************
Property Get CurrentTargetID() As String
    CurrentTargetID = brTab.targetID
End Property

'***************************************************************************************************
'* 機能　　：CDP-Json コマンド結果を得るためのタイムアウト時間を設定します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：設定中のタイムアウト秒数
'* LET引数 ：置き換えるタイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 機能説明：最初のブラウザ起動時の遷移完了判定や、CDP-Json コマンド結果を得るのに待つ時間を設定します。デフォルトは、30sです。
'***************************************************************************************************
Property Get TimeOutSecond() As Double
    TimeOutSecond = timerOut / 1000 'ミリ秒に補正する
End Property

Property Let TimeOutSecond(TimeSec As Double)
    timerOut = TimeSec * 1000       'ミリ秒に補正する
End Property

'***************************************************************************************************
'* 機能　　：非同期で実行した、CDP-Json コマンド結果を蓄積させる上限を決めます
'---------------------------------------------------------------------------------------------------
'* GET返値 ：設定中の上限数
'* LET引数 ：置き換える上限数
'---------------------------------------------------------------------------------------------------
'* 機能説明：`ExecuteCDPAsync`による結果をあとで、取り出す際の格納上限件数を来ます。デフォルトは、256件です。
'***************************************************************************************************
Property Get ResultAccumulatedLimitCDPForAsync() As Long
    ResultAccumulatedLimitCDPForAsync = currentResultCDP.Limit
End Property

Property Let ResultAccumulatedLimitCDPForAsync(Number As Long)
    currentResultCDP.Limit = Number
End Property

'***************************************************************************************************
'* 機能　　：同期モードでの CDP-Json コマンド結果のうち、最後のエラー情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後のエラー情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」のエラー情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます。
'            ・適切な`ExecuteCDP`を実行したつもりなのに返り値が、`Nothing`となるのが、エラーの目印となるので、それでこのエラー情報を見に行ってください。
'            ・`ExecuteCDP`の`StopError`が`False`だと実質機能しません。
'            ・非同期によるあとで取り出したパターンとは、連動しません。
'***************************************************************************************************
Property Get LastCDPJsonError() As Dictionary
    Set LastCDPJsonError = lastErrorDetailsCDP
End Property

'***************************************************************************************************
'* 機能　　：同期モードでの JavaScript結果のうち、最後の例外情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後の「例外（Exception）」情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」の例外情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます。
'            ・これはあくまでも、`例外`のみ扱います。「try { ... } catch(e) { return e }」系統の取得は対象外です。
'            ・非同期によるあとで取り出したパターンとは、連動しません。
'***************************************************************************************************
Property Get LastJavaScriptException() As Dictionary
    Set LastJavaScriptException = lastExceptionDetailsJS
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
Property Get ResultCDPFromWithEvents(commandID As Long) As String
    Const FromProcedureName As String = ThisClassName & ".ResultCDPFromWithEvents"


    '1. 存在するか？
    If ResultCDP_Browser.response.Exists(commandID) Then
        '2. Log化
        Dim tmp As String: tmp = ResultCDP_Browser.response(commandID)
        printMsg info_, "Browser were able to confirm the existence of CDP command ID:" & commandID & "　→　" & tmp, FromProcedureName

        '3. 返却
        ResultCDPFromWithEvents = tmp

        '4. 取り出したので消しておく
        ResultCDP_Browser.response.Remove commandID
        printMsg Trace_, "CDP-command ID: " & commandID & " has been removed from `ResultCDP_Browser.response`.", FromProcedureName
    Else
        '2. 存在しない旨を伝える
        printMsg Trace_, "This CDP-command ID(" & commandID & ") passed is not running asynchronously or has been deleted.", FromProcedureName
    End If
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
        If br.hWnd Then BrowserWindowHandle = br.hWnd Else BrowserWindowHandle = getHandle
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
        If br.WID Then BrowserWindowID = br.WID Else BrowserWindowID = ExecuteCDP("Browser.getWindowForTarget").Item("windowId")
    End If
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


    Dim result
    Dim results_ As Scripting.Dictionary
    Dim results__ As Scripting.Dictionary
    Dim results___ As Scripting.Dictionary

    Dim strKey As Variant

    On Error GoTo ifError

    '"Target.targetCreated"を収集できるように初期化
    Set TargetInfos = New Collection

    '実行
1   BrowserCore.Collect_TargetCreated = True
    printMsg info_, "Now looking for the browser window", FromProcedureName

   'Read message to confirm the requested url is found
   'Checking results__.Item("url") has been skipped for simplicity since v2.3
   'A new session always has the property attached = false and method = Target.targetCreated
   'This way of verification vastly improves the stability over checking by URL as sometimes
   'there are rogues targets with similar URL but they are not attachable (ie. no Session ID)
    Dim boolFound As Boolean
    For Each result In TargetInfos
        Set results_ = BrowserCore.jsConverter.ParseJson(result).Item("params")
        If results_.Exists("targetInfo") Then
            Set results__ = results_.Item("targetInfo")
            If results__.Item("attached") = False And results__.Item("type") = "page" Then
                boolFound = True
                Exit For
            End If
        End If
    Next
2   If Not boolFound Then printMsg ERROR_, "Unable to find the browser session", FromProcedureName, doRaiseError:=True _
    Else printMsg info_, "Now attaching CDP pipes to the browser session", FromProcedureName

   'Establish connection with the CDP pipes
3   brTab.targetID = results__.Item("targetId")
4   brTab.sessionID = BrowserCore.attachToTab(brTab.targetID)
5   Set results___ = SendMessage("{""method"":""Runtime.enable"",""params"":{}}")
6   BrowserCore.Collect_TargetCreated = False

    printMsg info_, "Successfully attached to Session ID " & brTab.sessionID, FromProcedureName
    Exit Sub

ifError:

   'Record errmsg and print helpful infos
    Dim errorLine As Long, errorDesc As String
    errorLine = Erl
    errorDesc = Err.Description
    printMsg ERROR_, "attachToSession error at line " & errorLine & ": """ & errorDesc & """", FromProcedureName
    printMsg Debug_, "result = " & result, FromProcedureName
    printMsg Debug_, "results_ = " & BrowserCore.jsConverter.ConvertToJson(results_), FromProcedureName
    printMsg Debug_, "results__ = " & BrowserCore.jsConverter.ConvertToJson(results__), FromProcedureName
    printMsg Debug_, "results___ = " & BrowserCore.jsConverter.ConvertToJson(results___), FromProcedureName
    printMsg Debug_, ".start call string: " & BrowserCore.CurrentBrowserPath, FromProcedureName

    printParams
    Err.Raise CDPCustomErrorCodes.Protocol, , "attachToSession error at line " & errorLine & ": """ & errorDesc & """"

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
    sessionTitle = jsEval("document.getElementsByTagName(""title"")[0].innerHTML", dbgMsg:=False, IFEXCEPTION:="about:blank")

    'Obtain handles to the correct browser window
    'Both Edge and Chrome use the same class chromeWindowClass
    hWnd = FindWindowEx(0&, 0&, StrPtr(chromeWindowClass), 0&)        'To return Chrome / Edge window handles
    Do
        '1. 取得したウィンドウハンドル(hwnd)の「プロセスID」を取得する
        Dim currentPID As Long
        GetWindowThreadProcessId hWnd, currentPID

        '2. プロセスIDが、探しているブラウザのものと一致するか？
        If currentPID = core.CollectionProcessId Then
            '一致すれば少なくとも、デバッグブラウザのウィンドウハンドルなので一旦確定させる
            getHandle = hWnd

            '3. Get Window Title based on specified handles
            'Requires string buffering declaration as string
            bufferString = String(1024, " ")
            titleLen = GetWindowText(hWnd, StrPtr(bufferString), Len(bufferString))
            windowTitle = Left(bufferString, titleLen)

            'Verify if matching with the sessionTitle and is really an edge/chrome window
            If windowTitle <> "" Then If InStr(windowTitle, sessionTitle) Then Exit Do
        End If

        'Continue with the next handle value
        hWnd = FindWindowEx(0&, hWnd, StrPtr(chromeWindowClass), 0&)

    Loop Until hWnd = 0 ' no more matching windows

End Function

Private Function getTargetID() As String
'----------------------------------------------------------------------------------
' Retrieve the targetId of the current tab (a.k.k "target" in CDP Documentation).
' This targetId is useful for tab automation, such as the .closeTab method.
'----------------------------------------------------------------------------------

    Dim results As Scripting.Dictionary
    Set results = ExecuteCDP("Target.getTargetInfo")
    getTargetID = results("targetInfo").Item("targetId")

End Function

Private Sub getSessionInfo()
'----------------------------------------------------------------------------------
' Obtain session info on the fly for printParams.
' Updated: 22/03/2023
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".getSessionInfo"


   'Activate to getHandle correctly
    activate False

   'Get infos
    br.Url = Url
    br.Name = Name
    br.hWnd = getHandle

    '対象タブが非表示の場合、`Browser.getWindowForTarget`が効かなくなるので、判定をかます
    Dim tmp As Dictionary: Set tmp = ExecuteCDP("Browser.getWindowForTarget", StopError:=False)
    If tmp Is Nothing Then printMsg info_, "The tab you are connecting to may be hidden.", FromProcedureName Else br.WID = tmp.Item("windowId")

    br.Type = ExecuteCDP("Browser.getVersion").Item("product")
    br.Ver = Split(br.Type, "/")(1)
    br.Path = getBrowserPath(br.Type)
    brTab.targetID = getTargetID

End Sub

Private Sub SaveBase64ToFile(strB64 As String, FolderPath As String, FileName As String)
'----------------------------------------------------------------------------------
' Decode a base64 string and save to file without external dependencies (WindowsAPI:CryptStringToBinary).
'----------------------------------------------------------------------------------

    'Base64 から、バイト配列に変換
    Dim tmp() As Byte
    tmp = DataConv.Decode(strB64, edfBase64)

    'そのバイト配列のまま、保存
    CharConv.BytesToSaveFile tmp, FolderPath, FileName

End Sub



'***************************************************************************************************
'                               ■■■ CDPCore.cls 直通メソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, LogID As String, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    PipeCore.printMsg LogLevel_, LogID, strMsg, From, isHeader, doRaiseError
End Sub

Private Sub TakeEvents(Optional destruction As Boolean, Optional StopPipeError As Boolean)
    PipeCore.TakeEvents destruction, StopPipeError
End Sub



'***************************************************************************************************
'                               ■■■ CDPBrowser.cls 直通メソッド ■■■
'***************************************************************************************************


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
    br.id = "TBID" & Format(Rnd * 1000, "000")

    'public用イベントメソッドフィルターを有効化
    Set publicBrowserEvents.Filter = New Dictionary

End Sub

'***************************************************************************************************
'* 機能　　：このClassを破棄する際、持ってる`SessionID`も併せて破棄します
'***************************************************************************************************
Private Sub Class_Terminate()
    Const FromProcedureName As String = ThisClassName & ".Class_Terminate"

    'SessionIDを保持するオプションになってるか確認
    If KeepSession Then
        printMsg Debug_, "The class object will be destroyed, but the `SessionID`(" & brTab.sessionID & ") will be preserved depending on the configuration.", FromProcedureName
    Else
        printMsg Debug_, "When the class object is destroyed, the `SessionID`(" & brTab.sessionID & ") is also destroyed.", FromProcedureName

        'SessionIDを破棄しておく
        DiscardSessionID brTab.sessionID
    End If

End Sub
'===============================================================================================
' Experimental Functions
'===============================================================================================
