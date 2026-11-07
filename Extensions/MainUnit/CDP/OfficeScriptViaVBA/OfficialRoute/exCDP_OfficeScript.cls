VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_OfficeScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           `CDPBrowser.cls` に対する機能拡張 - Officeスクリプト実行（OfficeScriptViaVBA）
'***************************************************************************************************
'* 機能　　：Excelの「コード エディター」(Officeスクリプト)へCDPで接続し、
'            VBAから渡したスクリプトを書き込んで実行、結果を回収します
'---------------------------------------------------------------------------------------------------
'* 使い方　：
'    '1. bas側で`CDPBrowser`を用意(接続方法は Demo_OfficeScript の ConnectDesktopExcel / ConnectWebExcel を参照)
'    Dim Browser As CDPBrowser: Set Browser = ConnectDesktopExcel()
'
'    '2. 継承して、コードエディターを確保(Web版Excelは WebMode:=True)
'    Dim os As New exCDP_OfficeScript
'    If os.Init(Browser, WebMode:=False) Then
'        '同期実行
'        Debug.Print os.RunScript(RawOfficeScript)       ' → Succeeded / Failed / Timeout ...
'        Debug.Print os.LastLogs
'        Debug.Print os.LastExceptionMessage             ' → スクリプト内で例外が出た場合のメッセージ
'
'        '非同期実行（整理券方式）
'        Dim ticket As Long: ticket = os.RunScriptAsync(RawOfficeScript)
'        Do Until os.TryGetResult(ticket, status): DoEvents: Loop
'    End If
'
'    '3. 後片付け(先にこのClassを手放してから、bas側で切断)
'    os.Dispose
'    Browser.DisconnectCDP
'---------------------------------------------------------------------------------------------------
'* 仕組み　：
'    ・コードエディターは、WebView2上のReactアプリ(Monaco Editor + Redux)
'    ・Monacoは`window.monaco`として公開されていないため、DOM(.monaco-editor)からReactファイバーを
'      遡って、エディターインスタンス(`stateNode.editor`)を掴む
'    ・実行状態/出力ログは、react-reduxの`Provider`が持つ`store`から読む
'        state.run.runningStatus      : None → InProgress → Succeeded / Failed / Unsupported
'        state.bottomPanel.outputLogs : console.log や実行時エラーの出力
'        state.messageBar             : 「スクリプトが正常に実行されました。」等の通知
'    ・実行は、ツールバーの「実行」ボタンをクリックする(エディター上の未保存の内容がそのまま実行される)
'    ・ページ内に`window.__exCDP_OfficeScript`というヘルパーを注入し、実行結果は整理券番号ごとに
'      ページ側で保持する。VBA側は整理券番号で、好きなタイミングで結果を取りに行ける
'    ・スクリプト本体は別タブ「Office JS Sandbox」で実行され、例外の詳細(message/stack)は
'      そちらの`console.error("Command Error:\n...")`にだけ出る。エディター側には汎用メッセージしか来ないため、
'      Sandbox側の`console.error`をフック(window.__exCDP_OfficeScriptSandbox)して回収する
'      → `LastExceptionDetails` / `LastExceptionMessage`
'    ・Web版(`Init`でWebMode:=True)も、中身は同じコードエディター(同じReact/Redux構成)。違いは次の通り
'        接続     : リモートデバッグを有効にしたEdgeへ、`DevToolsActivePort`のブラウザ用WebSocketで繋ぐ(bas側)
'        エディター: タブではなく、Excelページ内のiframe(URL: .../index.html#/edit)
'        例外     : メッセージはエディターの出力ログにそのまま出るので、そこから回収する(stackは取れない)
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - 事前に、Excelで[自動化]タブからスクリプトを開き、「コード エディター」を表示しておくこと
'    - エディターで開いているスクリプトの中身は、渡したスクリプトで上書きされる(自動保存される)
'    - コードエディターは1つなので、同時に実行できるのは1本だけ。実行中の`RunScriptAsync`は失敗(0)する
'    - `main`に引数(workbook以外)を持つスクリプトは、パラメーター入力ダイアログが出るので対象外
'    - ReactファイバーやRedux Storeの構造は、Office側の更新で変わる可能性がある(POC由来の実装)
'    - `CDPBrowser`の接続/切断はbas側の責務。このClassは切断しない。
'      `CDPContext`は破棄時に`Target.detachFromTarget`等を送るため、`DisconnectCDP`の「前」に`Dispose`(または破棄)すること
'    - 非同期で`Application.OnTime`等により結果回収する場合、このClassのインスタンスと`CDPBrowser`は
'      モジュールレベル変数で保持しておくこと
'    - Web版は、事前にEdgeのリモートデバッグを有効にしておくこと(`edge://inspect/#remote-debugging`で許可、等)
'    - 【既知の制約・デスクトップ版のみ】「Office JS Sandbox」タブは、Excel起動後の最初の実行時に初めて作られる。
'      そのため、Sandboxが無い状態で`Init`したインスタンスでは例外の詳細が取れない
'      (`LastExceptionDetails`はNothing。Failedの判定やエディター側の出力には影響なし)。
'      1回実行した後に`Init`し直せば取れるようになる。
'      確実に拾うにはSandbox作成時の自動アタッチ(Target.setAutoAttach等)が必要だが、設計が固まるまで保留
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private ex_Browser      As CDPBrowser
Private ex_Editor       As CDPContext
Private ex_Sandbox      As CDPContext   '例外詳細の回収用(取れなくても動作は継続)

'内部使用用
Private Const ThisClassName         As String = "exCDP_OfficeScript"   'ログのトレース用

'コードエディター/Sandboxの探し方(デスクトップ版)
Private Const DEFAULT_EDITOR_TITLE      As String = "コード エディター"
Private Const SANDBOX_TITLE             As String = "Office JS Sandbox"

'コードエディターの探し方(Web版)
Private Const DEFAULT_WEB_EDITOR_URL        As String = "index.html#/edit"             'コードエディターiframeのURL(部分一致)

'ページ内ヘルパー
Private Const HELPER_NAME           As String = "window.__exCDP_OfficeScript"
Private Const HELPER_VERSION        As String = "1.0.3"     'JSを変更したら上げる(ページに残った旧ヘルパーを差し替えるため)
Private Const SANDBOX_HELPER_NAME       As String = "window.__exCDP_OfficeScriptSandbox"
Private Const SANDBOX_HELPER_VERSION    As String = "1.0.0"
Private Const LOG_SETTLE_MS         As Long = 3000      '完了検知後、結果のメッセージバーが出るのを待つ上限(ミリ秒)

'待機設定
Private Const SYNC_TIMEOUT_SEC      As Double = 10      'エディター→Redux Storeへの反映待ち上限(秒)
Private Const POLL_INTERVAL_SEC     As Double = 0.5     '結果ポーリング間隔(秒)

'ログ設定
Private m_LogID                     As String   'ログ識別ID

'接続中のルート
Private m_IsWeb                     As Boolean  'True:Web版(ブラウザのExcel) / False:デスクトップ版

'直近に回収した結果
Private m_LastTicket                As Long
Private m_LastStatus                As String
Private m_LastLogs                  As String
Private m_LastOutputs               As Collection
Private m_LastExceptionDetails      As BiDiCDPJConv     'Sandbox側で捕まえた例外({"name","message","stack","raw","time"})



'***************************************************************************************************
'                                  ■■■ 準備 ■■■
'***************************************************************************************************
'* 機能　　：Excel内のWebView2を落とし、確実にデバッグ経路を確保するようにします
'---------------------------------------------------------------------------------------------------
'* 注意事項：・WebView2のデバッグポートは、環境変数`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`で開けるため、
'              確実にデバッグ経路を確保できるように、一旦Excel内のWebView2を落として再起動を促します。
'              一定回数以内であれば、自動でWebView2が再起動されます。
'***************************************************************************************************
Sub OpenDebugPortWebView2(Optional UserDataDir As String)
    '1. 専用の環境変数を付与し、デバッグ経路を開ける状態にする
    Dim WebView2Debug As CDPCoreViaWebSocket: Set WebView2Debug = New CDPCoreViaWebSocket
    WebView2Debug.EnsureWebView2DebugPort = 0

    '2. WebView2プロセスを落とすためだけに、WebView2専用のClassを用意
    Dim DropWebView2 As WebView2Host: Set DropWebView2 = New WebView2Host

    '3. ExcelのWebView2プロセスだけを落とすように、`--user-data-dir`でスコープをセット
    If StrPtr(UserDataDir) = 0 Then UserDataDir = Environ("LOCALAPPDATA") & "\Microsoft\Office\16.0\Wef\webview2\2\1"
    BrowserCapabilities.Set_UserDataDirName = UserDataDir

    '4. 情報を継承し、落とす
    DropWebView2.Init BrowserCapabilities, 0
    DropWebView2.cleanUpSessions
    Set BrowserCapabilities = Nothing
End Sub




'***************************************************************************************************
'                                  ■■■ 接続状態 ■■■
'***************************************************************************************************
'* 機能　　：コードエディター/Sandboxへの参照を手放します(`CDPBrowser`の切断はしません)
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`CDPContext`は破棄時(Class_Terminate)に`Target.detachFromTarget`等を送るため、
'              bas側で`CDPBrowser.DisconnectCDP`する「前」に、このメソッドを呼ぶ(またはこのClassを破棄する)こと
'            ・破棄時(Class_Terminate)にも自動で呼ばれます
'***************************************************************************************************
Public Sub Dispose()
    DiscardSandboxConsole   '次回`Init`時の、コンソール履歴の一斉再送を減らす
    Set ex_Sandbox = Nothing
    Set ex_Editor = Nothing
    Set ex_Browser = Nothing
    m_IsWeb = False
End Sub

'***************************************************************************************************
'* 機能　　：`Init`済み(コードエディターを確保済み)かどうか
'***************************************************************************************************
Public Property Get IsConnected() As Boolean
    IsConnected = Not ex_Editor Is Nothing
End Property

'***************************************************************************************************
'* 機能　　：Web版(`Init`でWebMode:=True)として初期化したかどうか
'***************************************************************************************************
Public Property Get IsWeb() As Boolean
    IsWeb = m_IsWeb
End Property



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分編 ■■■
'***************************************************************************************************
'* 機能　　：Officeスクリプトを書き込んで実行し、完了まで待ちます(同期実行)
'---------------------------------------------------------------------------------------------------
'* 返り値　：Succeeded / Failed / Unsupported / Timeout / NotStarted / Error
'* 引数　　：RawOfficeScript    Officeスクリプト(TypeScript)の全文
'            TimeoutSec         完了待ちの上限(秒)
'---------------------------------------------------------------------------------------------------
'* 注意事項：結果の詳細は、`LastLogs` / `LastOutputs` で確認できます
'***************************************************************************************************
Public Function RunScript(RawOfficeScript As String, Optional TimeoutSec As Double = 120) As String
    Dim TicketID As Long
    TicketID = RunScriptAsync(RawOfficeScript)
    If TicketID = 0 Then
        RunScript = m_LastStatus    'NotStarted
        Exit Function
    End If

    RunScript = WaitResult(TicketID, TimeoutSec)
End Function

'***************************************************************************************************
'* 機能　　：Officeスクリプトを書き込んで実行を開始し、整理券番号を返します(非同期実行)
'---------------------------------------------------------------------------------------------------
'* 返り値　：整理券番号(1以上)。開始できなかった場合は 0 (理由は`LastLogs`)
'* 引数　　：RawOfficeScript    Officeスクリプト(TypeScript)の全文
'---------------------------------------------------------------------------------------------------
'* 注意事項：結果は`TryGetResult` / `WaitResult`に整理券番号を渡して回収します
'***************************************************************************************************
Public Function RunScriptAsync(RawOfficeScript As String) As Long
    Const FromProcedureName As String = ThisClassName & ".RunScriptAsync"


    '0. 前提チェック
    If Not IsConnected Then SetNotStarted "Not initialized. Call `Init` first.", FromProcedureName: Exit Function

    Dim ret As String
    ret = EnsureHelper()
    If ret <> "OK" Then SetNotStarted "The code editor is not ready: " & ret, FromProcedureName: Exit Function

    '実行中に書き換えると、実行中のスクリプトとエディターの内容がズレるので、先に弾く
    If EvalString("String(" & HELPER_NAME & ".isBusy())") = "true" Then SetNotStarted "Another script is running.", FromProcedureName: Exit Function

    'Sandbox側のフックが外れていれば(ページ再読み込み等)、掛け直す
    EnsureSandboxHook

    '1. エディターへ書き込み(JSの文字列リテラル化は\uXXXXエスケープ込み)
    Dim v As Variant: v = RawOfficeScript
    ret = EvalString(HELPER_NAME & ".setCode(" & WebJsonConverter.serialize(v) & ")")
    If ret <> "OK" Then SetNotStarted "setCode failed: " & ret, FromProcedureName: Exit Function

    '2. React(onChange)→Redux Storeへの反映待ち
    Dim StartSec As Double: StartSec = Timer
    Do Until EvalString("String(" & HELPER_NAME & ".isSynced())") = "true"
        If ElapsedSec(StartSec) > SYNC_TIMEOUT_SEC Then SetNotStarted "The editor content was not reflected in the store.", FromProcedureName: Exit Function
        sleep 0.2
    Loop

    '3. 実行ボタンをクリック(ページ側で整理券番号が発行される)
    ret = EvalString(HELPER_NAME & ".run()")
    If Not IsNumeric(ret) Then SetNotStarted "run failed: " & ret, FromProcedureName: Exit Function

    RunScriptAsync = CLng(ret)
    printMsg info_, "The Office Script has started. Ticket: " & RunScriptAsync, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：整理券番号の実行結果を1回だけ確認します(待機しない)
'---------------------------------------------------------------------------------------------------
'* 返り値　：完了していれば True。未完了なら False
'* 引数　　：TicketID   `RunScriptAsync`の返り値
'            Status     (ByRef) 完了時の状態。Succeeded / Failed / Unsupported / Error 等
'---------------------------------------------------------------------------------------------------
'* 注意事項：完了時は`LastTicket` / `LastStatus` / `LastLogs` / `LastOutputs` /
'            `LastExceptionDetails` / `LastExceptionMessage` が更新されます
'            ページの再読み込み等で結果が失われた場合は、True & Status="Error" を返します
'***************************************************************************************************
Public Function TryGetResult(TicketID As Long, Optional ByRef Status As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".TryGetResult"


    '0. 前提チェック
    If Not IsConnected Then
        SetResult TicketID, "Error", "Not connected.", New Collection
        Status = m_LastStatus: TryGetResult = True
        Exit Function
    End If
    If EvalString("typeof " & HELPER_NAME) <> "object" Then
        SetResult TicketID, "Error", "The helper is missing. The page may have been reloaded.", New Collection
        Status = m_LastStatus: TryGetResult = True
        Exit Function
    End If

    '1. 状態確認
    Dim ret As String
    ret = EvalString(HELPER_NAME & ".result(" & TicketID & ")")
    If Len(ret) = 0 Then Exit Function      '未完了

    If Left$(ret, 3) = "NG:" Then
        SetResult TicketID, "Error", ret, New Collection
    Else
        '2. 完了：ログ類を回収
        SetResult TicketID, ret, _
                  Replace(EvalString(HELPER_NAME & ".logs(" & TicketID & ")"), vbLf, vbCrLf), _
                  SplitToCollection(EvalString(HELPER_NAME & ".outputs(" & TicketID & ")"), Chr$(30))

        '3. 例外の詳細を回収(デスクトップ版はSandbox側のフック、Web版はエディターの出力ログから)
        If m_IsWeb Then
            CollectExceptionWeb TicketID
        Else
            CollectException TicketID
        End If
    End If

    printMsg info_, "Ticket " & TicketID & " finished: " & m_LastStatus, FromProcedureName
    Status = m_LastStatus
    TryGetResult = True
End Function

'***************************************************************************************************
'* 機能　　：整理券番号の実行完了まで待ちます
'---------------------------------------------------------------------------------------------------
'* 返り値　：Succeeded / Failed / Unsupported / Error / Timeout
'* 引数　　：TicketID   `RunScriptAsync`の返り値
'            TimeoutSec 待機の上限(秒)
'---------------------------------------------------------------------------------------------------
'* 注意事項：Timeoutしてもスクリプトの実行自体は続いているため、同じ整理券番号で再度待機できます
'***************************************************************************************************
Public Function WaitResult(TicketID As Long, Optional TimeoutSec As Double = 120) As String
    Const FromProcedureName As String = ThisClassName & ".WaitResult"


    Dim Status As String
    Dim StartSec As Double: StartSec = Timer
    Do
        If TryGetResult(TicketID, Status) Then
            WaitResult = Status
            Exit Function
        End If
        sleep POLL_INTERVAL_SEC     '内部でDoEventsも走る
    Loop While ElapsedSec(StartSec) < TimeoutSec

    '時間切れ
    SetResult TicketID, "Timeout", "Timed out after " & TimeoutSec & " seconds. runningStatus: " & EvalString(HELPER_NAME & ".runningStatus()"), New Collection
    printMsg WARN_, "Ticket " & TicketID & " timed out.", FromProcedureName
    WaitResult = m_LastStatus
End Function

'***************************************************************************************************
'* 機能　　：コードエディターでスクリプトが実行中かどうか(VBA以外からの実行も含む。クリック直後の開始待ちも実行中扱い)
'***************************************************************************************************
Public Property Get IsRunning() As Boolean
    If Not IsConnected Then Exit Property
    If EnsureHelper() <> "OK" Then Exit Property
    IsRunning = (EvalString("String(" & HELPER_NAME & ".isBusy())") = "true")
End Property



'***************************************************************************************************
'                                ■■■ 直近の結果 ■■■
'***************************************************************************************************
'* 機能　　：直近に回収した(または開始に失敗した)整理券番号
'***************************************************************************************************
Public Property Get LastTicket() As Long
    LastTicket = m_LastTicket
End Property

'***************************************************************************************************
'* 機能　　：直近の状態。Succeeded / Failed / Unsupported / Timeout / NotStarted / Error
'***************************************************************************************************
Public Property Get LastStatus() As String
    LastStatus = m_LastStatus
End Property

'***************************************************************************************************
'* 機能　　：直近のログ全文(状態遷移 / console出力 / 問題 / メッセージバー / Sandbox側の例外)
'***************************************************************************************************
Public Property Get LastLogs() As String
    LastLogs = m_LastLogs
End Property

'***************************************************************************************************
'* 機能　　：直近の出力ログ(console.log やエラーメッセージ)を1件ずつ格納したCollection
'***************************************************************************************************
Public Property Get LastOutputs() As Collection
    Set LastOutputs = m_LastOutputs
End Property

'***************************************************************************************************
'* 機能　　：直近の実行で、スクリプト内で発生した例外の詳細
'---------------------------------------------------------------------------------------------------
'* 返り値　：{"name":"Error","message":"...","stack":"...","raw":"Command Error:...","time":1234567890123}
'            例外が無い、またはSandboxタブに接続できていない場合は Nothing
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`CDPContext.LastJavaScriptException`と同じく`BiDiCDPJConv`で返します
'            ・1回の実行で複数捕まえた場合は、最初の"Command Error"(根本原因)を保持します。
'              全件は`LastLogs`の[exception]行で確認できます
'            ・stackの行番号は、Office側でラップされた後のコード基準です
'            ・【既知の制約・デスクトップ版】Excel起動後の最初の実行(Sandboxタブがまだ無い状態で`Init`した場合)は取れません
'            ・Web版(`Init`でWebMode:=True)は、エディターの出力ログ(severity=error)から回収します。
'              最初の実行から取れますが、"name"と"stack"は空文字です
'***************************************************************************************************
Public Property Get LastExceptionDetails() As BiDiCDPJConv
    Set LastExceptionDetails = m_LastExceptionDetails
End Property

'***************************************************************************************************
'* 機能　　：直近の実行で発生した例外のメッセージ(`LastExceptionDetails`の"message")。無ければ空文字
'***************************************************************************************************
Public Property Get LastExceptionMessage() As String
    If m_LastExceptionDetails Is Nothing Then Exit Property
    LastExceptionMessage = m_LastExceptionDetails.StringKey("message")
End Property



'***************************************************************************************************
'                                ■■■ 内部処理 ■■■
'***************************************************************************************************
'* 機能　　：ページ内ヘルパーが無い(またはバージョン違い)なら注入し、エディター/Storeが掴めるか確認します
'* 返り値　："OK" または "NG: ..."
'***************************************************************************************************
Private Function EnsureHelper() As String
    If EvalString("String(" & HELPER_NAME & " && " & HELPER_NAME & ".version)") <> HELPER_VERSION Then
        Dim ret As String: ret = EvalString(BuildHelperJS())
        If ret <> "installed" Then
            EnsureHelper = "NG: helper install failed: " & ret
            Exit Function
        End If
    End If
    EnsureHelper = EvalString(HELPER_NAME & ".ready()")
End Function

'***************************************************************************************************
'* 機能　　：「Office JS Sandbox」タブへ接続し、`console.error`のフックを掛けます
'---------------------------------------------------------------------------------------------------
'* 注意事項：・失敗しても実行自体には影響しないため、警告のみで続行します(`LastExceptionDetails`がNothingになるだけ)
'            ・Sandboxタブは、Excel起動後の最初の実行時に作られます。それより前の`Init`では見つかりません(既知の制約)
'***************************************************************************************************
Private Sub ConnectSandbox()
    Const FromProcedureName As String = ThisClassName & ".ConnectSandbox"


    On Error GoTo Failed

    '1. タブ確保(部分一致で別タブを掴んでいないか、タイトルで確認)
    Set ex_Sandbox = ex_Browser.getTab(SANDBOX_TITLE)
    If Not ex_Sandbox Is Nothing Then
        If EvalStringOn(ex_Sandbox, "document.title") <> SANDBOX_TITLE Then Set ex_Sandbox = Nothing
    End If
    If ex_Sandbox Is Nothing Then
        printMsg WARN_, "The tab `" & SANDBOX_TITLE & "` was not found. Exception details will not be available.", FromProcedureName
        Exit Sub
    End If

    '2. フック
    If Not EnsureSandboxHook() Then
        printMsg WARN_, "Failed to hook `console.error` in `" & SANDBOX_TITLE & "`. Exception details will not be available.", FromProcedureName
        Set ex_Sandbox = Nothing
        Exit Sub
    End If

    '3. 接続時に受け取ったコンソール履歴は不要なので、Sandbox側の蓄積を捨てておく
    DiscardSandboxConsole
    Exit Sub

Failed:
    printMsg WARN_, "Could not connect to `" & SANDBOX_TITLE & "`. Error " & Err.Number & ": " & Err.Description, FromProcedureName
    Set ex_Sandbox = Nothing
End Sub

'***************************************************************************************************
'* 機能　　：Sandboxタブに蓄積されたコンソール履歴を破棄します(`Runtime.discardConsoleEntries`)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Sandboxタブは Excel を閉じるまで残り、実行のたびにコンソール出力が蓄積される。
'            `CDPContext`は接続時に`Runtime.enable`するため、その蓄積が`Runtime.consoleAPICalled`として
'            一斉に再送され、`Init`が数十秒かかるようになる。接続直後と`Dispose`時に捨てて、これを防ぐ
'* 注意事項：例外の回収は`console.error`のフック側で保持しているため、ここで捨てても影響しない。
'            失敗しても動作には影響しないため、エラーは無視する
'***************************************************************************************************
Private Sub DiscardSandboxConsole()
    If ex_Sandbox Is Nothing Then Exit Sub

    On Error Resume Next
    ex_Sandbox.ExecuteCDP "Runtime.discardConsoleEntries", StopCDPError:=False
    On Error GoTo 0
End Sub

'***************************************************************************************************
'* 機能　　：Sandboxタブのフックが無い(またはバージョン違い)なら掛けます
'* 返り値　：フック済みならTrue。Sandboxタブ未接続ならFalse
'***************************************************************************************************
Private Function EnsureSandboxHook() As Boolean
    If ex_Sandbox Is Nothing Then Exit Function

    If EvalStringOn(ex_Sandbox, "String(" & SANDBOX_HELPER_NAME & " && " & SANDBOX_HELPER_NAME & ".version)") <> SANDBOX_HELPER_VERSION Then
        If EvalStringOn(ex_Sandbox, BuildSandboxHookJS()) <> "installed" Then Exit Function
    End If
    EnsureSandboxHook = True
End Function

'***************************************************************************************************
'* 機能　　：整理券番号の実行開始以降に、Sandbox側で捕まえた例外を回収します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`m_LastExceptionDetails`に最初の"Command Error"を保持し、`m_LastLogs`に全件を[exception]行で追記します
'***************************************************************************************************
Private Sub CollectException(TicketID As Long)
    Const FromProcedureName As String = ThisClassName & ".CollectException"


    If ex_Sandbox Is Nothing Then Exit Sub

    '1. 実行開始時刻(エディター側のDate.now)。同じPCの時計なので、Sandbox側の記録時刻と比較できる
    Dim SinceMs As String
    SinceMs = EvalString(HELPER_NAME & ".startedAt(" & TicketID & ")")
    If Len(SinceMs) = 0 Or Not IsNumeric(SinceMs) Then Exit Sub

    '2. フックが外れていたら(Sandboxの再読み込み等)、回収できない
    If EvalStringOn(ex_Sandbox, "String(" & SANDBOX_HELPER_NAME & " && " & SANDBOX_HELPER_NAME & ".version)") <> SANDBOX_HELPER_VERSION Then
        printMsg WARN_, "The `console.error` hook in `" & SANDBOX_TITLE & "` was lost. Exception details are not available.", FromProcedureName
        Exit Sub
    End If

    '3. 最初の例外を保持
    Dim Json As String
    Json = EvalStringOn(ex_Sandbox, SANDBOX_HELPER_NAME & ".first(" & SinceMs & ")")
    If Len(Json) = 0 Or Left$(Json, 3) = "NG:" Then Exit Sub
    Set m_LastExceptionDetails = BiDiCDPJConv.Parse(Json)

    '4. 全件をログへ追記
    Dim Lines As String
    Lines = EvalStringOn(ex_Sandbox, SANDBOX_HELPER_NAME & ".lines(" & SinceMs & ")")
    If Len(Lines) Then m_LastLogs = m_LastLogs & vbCrLf & "[exception] " & Replace(Lines, vbLf, vbCrLf & "[exception] ")

    printMsg info_, "Exception captured in `" & SANDBOX_TITLE & "`: " & m_LastExceptionDetails.StringKey("message"), FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：Web版：エディターの出力ログ(severity=error)から、例外のメッセージを回収します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Web版は、例外のメッセージがエディターの出力ログにそのまま出る(デスクトップ版は汎用メッセージのみ)。
'            一方、Sandbox(iframe + Worker)は実行ごとに作り直されるため、フックできず stack は取れない。
'            `m_LastExceptionDetails`は {"time","name":"","message","stack":"","raw"} の形で保持する
'***************************************************************************************************
Private Sub CollectExceptionWeb(TicketID As Long)
    Const FromProcedureName As String = ThisClassName & ".CollectExceptionWeb"


    '1. 最初のエラー出力を保持
    Dim Json As String
    Json = EvalString(HELPER_NAME & ".firstError(" & TicketID & ")")
    If Len(Json) = 0 Or Left$(Json, 3) = "NG:" Then Exit Sub
    Set m_LastExceptionDetails = BiDiCDPJConv.Parse(Json)

    '2. 全件をログへ追記(デスクトップ版と同じく[exception]行で)
    Dim Lines As String
    Lines = EvalString(HELPER_NAME & ".errorLines(" & TicketID & ")")
    If Len(Lines) Then m_LastLogs = m_LastLogs & vbCrLf & "[exception] " & Replace(Lines, vbLf, vbCrLf & "[exception] ")

    printMsg info_, "Exception captured in the code editor output: " & m_LastExceptionDetails.StringKey("message"), FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：JSを評価して文字列で返します。例外時は "NG: ..." を返します
'***************************************************************************************************
Private Function EvalString(JavaScriptStr As String) As String
    EvalString = EvalStringOn(ex_Editor, JavaScriptStr)
End Function

Private Function EvalStringOn(Target As CDPContext, JavaScriptStr As String) As String
    Dim v As Variant
    v = Target.jsEval(JavaScriptStr)
    If IsError(v) Then
        EvalStringOn = "NG: JavaScript exception"
        If Not Target.LastJavaScriptException Is Nothing Then EvalStringOn = EvalStringOn & ": " & Target.LastJavaScriptException.Stringify
    ElseIf IsNull(v) Or IsEmpty(v) Then
        EvalStringOn = ""
    Else
        EvalStringOn = CStr(v)
    End If
End Function

Private Sub SetResult(TicketID As Long, Status As String, Logs As String, Outputs As Collection)
    m_LastTicket = TicketID
    m_LastStatus = Status
    m_LastLogs = Logs
    Set m_LastOutputs = Outputs
    Set m_LastExceptionDetails = Nothing
End Sub

Private Sub SetNotStarted(Reason As String, From As String)
    SetResult 0, "NotStarted", Reason, New Collection
    printMsg WARN_, Reason, From
End Sub

Private Function SplitToCollection(Text As String, Separator As String) As Collection
    Set SplitToCollection = New Collection
    If Len(Text) = 0 Then Exit Function

    Dim Item As Variant
    For Each Item In Split(Text, Separator)
        SplitToCollection.Add CStr(Item)
    Next
End Function

Private Function ElapsedSec(StartSec As Double) As Double
    ElapsedSec = Timer - StartSec
    If ElapsedSec < 0 Then ElapsedSec = ElapsedSec + 86400#   '日付またぎ
End Function

'***************************************************************************************************
'* 機能　　：コードエディターのページへ注入するヘルパー(window.__exCDP_OfficeScript)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ready()        エディター/Storeが掴めれば 'OK'
'            setCode(t)     エディターの全文を置換(Ctrl+Zで戻せる編集として)
'            isSynced()     アクティブなスクリプトのStore上の本文が、setCodeの内容と一致したか
'            isBusy()       実行中か(直前にクリックしたが、まだ開始が反映されていない場合も含む)
'            run()          Storeを購読してから「実行」ボタンをクリックし、整理券番号を返す
'            result(id)     完了していれば最終状態、未完了なら空文字。完了後は結果のメッセージバー表示を待つ
'                           (「実行中」→「結果」の2件が揃うか、LOG_SETTLE_MS経過で確定)
'            logs(id)       状態遷移/出力ログ/問題/メッセージバー(改行区切り)
'            outputs(id)    出力ログのみ(\u001e区切り)
'            startedAt(id)  実行開始時刻(Date.now)。Sandbox側の例外を時刻で絞り込むのに使う
'            firstError(id) 出力ログのうち最初のエラー(severity=error)をJSON文字列で。Web版の例外回収用
'            errorLines(id) 出力ログのエラーを改行区切りで。Web版の例外回収用
'***************************************************************************************************
Private Function BuildHelperJS() As String
    Dim js As String
    js = js & "(() => {"
    js = js & "const V = '" & HELPER_VERSION & "', P = window.__exCDP_OfficeScript;"
    js = js & "if (P && P.version === V) return 'installed';"
    js = js & "if (P && P.unsub) P.unsub();"
    js = js & "const O = { version: V, seq: 0, cur: 0, runs: {}, unsub: null, settleMs: " & LOG_SETTLE_MS & " };"
    js = js & "const norm = t => String(t).replace(/\r\n?/g, '\n');"
    js = js & "const fmt = v => typeof v === 'string' ? v : JSON.stringify(v);"
    js = js & "const fiberOf = el => { const k = Object.keys(el).find(k => k.startsWith('__reactFiber$')); return k ? el[k] : null; };"
    js = js & "const isEd = v => !!v && typeof v === 'object' && typeof v.getModel === 'function' && typeof v.executeEdits === 'function';"

    'Monacoエディターインスタンス(.monaco-editor から React ファイバーを遡る)
    js = js & "O.editor = () => {"
    js = js & "  for (let el = document.querySelector('.monaco-editor'); el; el = el.parentElement) {"
    js = js & "    for (let f = fiberOf(el), d = 0; f && d < 40; f = f.return, d++) {"
    js = js & "      const sn = f.stateNode;"
    js = js & "      if (sn && !(sn instanceof Node)) { if (isEd(sn.editor)) return sn.editor; if (isEd(sn)) return sn; }"
    js = js & "    }"
    js = js & "  }"
    js = js & "  return null;"
    js = js & "};"

    'Redux Store(react-redux の Provider の props.store)
    js = js & "O.store = () => {"
    js = js & "  if (O._store) return O._store;"
    js = js & "  for (const n of document.body.querySelectorAll('*')) {"
    js = js & "    for (let f = fiberOf(n); f; f = f.return) {"
    js = js & "      const p = f.memoizedProps;"
    js = js & "      if (p && p.store && typeof p.store.getState === 'function') return (O._store = p.store);"
    js = js & "    }"
    js = js & "  }"
    js = js & "  return null;"
    js = js & "};"
    js = js & "O.ready = () => !O.store() ? 'NG: redux store not found' : (!O.editor() ? 'NG: monaco editor not found (open a script in the code editor)' : 'OK');"
    js = js & "O.runningStatus = () => { const st = O.store(); return st ? String(st.getState().run.runningStatus) : 'NG: redux store not found'; };"

    '書き込み
    js = js & "O.setCode = (text) => {"
    js = js & "  const ed = O.editor(); if (!ed) return 'NG: editor not found';"
    js = js & "  const m = ed.getModel(); if (!m) return 'NG: model not found';"
    js = js & "  const t = norm(text); O.expected = t;"
    js = js & "  ed.pushUndoStop();"
    js = js & "  ed.executeEdits('exCDP_OfficeScript', [{ range: m.getFullModelRange(), text: t, forceMoveMarkers: true }]);"
    js = js & "  ed.pushUndoStop();"
    js = js & "  return norm(m.getValue()) === t ? 'OK' : 'NG: editor value mismatch';"
    js = js & "};"

    'Storeへの反映確認(scripts.scripts[id] と mruScripts の両方を見る)
    js = js & "O.isSynced = () => {"
    js = js & "  const st = O.store(); if (!st || O.expected === undefined) return false;"
    js = js & "  const s = st.getState(); const id = s.editor.active; const c = [];"
    js = js & "  if (s.scripts.scripts && s.scripts.scripts[id]) c.push(s.scripts.scripts[id]);"
    js = js & "  for (const x of (s.scripts.mruScripts || [])) if (x && x.storageScript && x.storageScript.id === id) c.push(x);"
    js = js & "  return c.some(x => { const b = (x.localUpdates && typeof x.localUpdates.body === 'string') ? x.localUpdates.body : x.storageScript.body; return norm(b) === O.expected; });"
    js = js & "};"

    '実行中判定(クリック直後でStoreにまだ開始が反映されていない5秒間も実行中扱い)
    js = js & "O.isBusy = () => {"
    js = js & "  if (O.runningStatus() === 'InProgress') return true;"
    js = js & "  const r = O.runs[O.cur];"
    js = js & "  return !!(r && !r.status && Date.now() - r.startedAt < 5000);"
    js = js & "};"

    '実行(ボタンのaria-labelは日本語UI「実行」/英語UI「Run」)
    js = js & "O.runButton = () => [...document.querySelectorAll('button')].find(b => ['実行', 'Run'].includes(b.getAttribute('aria-label')));"
    js = js & "O.run = () => {"
    js = js & "  const st = O.store(); if (!st) return 'NG: store not found';"
    js = js & "  if (O.isBusy()) return 'NG: another script is running';"
    js = js & "  const b = O.runButton(); if (!b) return 'NG: run button not found';"
    js = js & "  if (b.disabled || b.getAttribute('aria-disabled') === 'true') return 'NG: run button disabled';"
    js = js & "  const prevRun = O.runs[O.cur];"
    js = js & "  if (prevRun && !prevRun.status) { prevRun.status = 'Abandoned'; prevRun.doneAt = Date.now(); }"
    js = js & "  const s0 = st.getState();"
    js = js & "  const id = ++O.seq;"
    js = js & "  const r = O.runs[id] = { id: id, status: '', startedAt: Date.now(), doneAt: 0, lastMsgAt: 0, history: [], messages: [], logStart: s0.bottomPanel.outputLogs.length, final: null };"
    js = js & "  let prev = s0.run.runningStatus, prevMsg = (s0.messageBar && s0.messageBar.isVisible) ? fmt(s0.messageBar.text || '') : '';"
    js = js & "  if (O.unsub) { O.unsub(); O.unsub = null; }"
    js = js & "  O.unsub = st.subscribe(() => {"
    js = js & "    const s = st.getState();"
    js = js & "    if (s.bottomPanel.outputLogs.length < r.logStart) r.logStart = 0;"
    js = js & "    const mb = (s.messageBar && s.messageBar.isVisible) ? fmt(s.messageBar.text || '') : '';"
    js = js & "    if (mb && mb !== prevMsg) { r.messages.push(mb); r.lastMsgAt = Date.now(); }"
    js = js & "    prevMsg = mb;"
    js = js & "    const cur = s.run.runningStatus;"
    js = js & "    if (cur === prev) return;"
    js = js & "    prev = cur; r.history.push(cur);"
    js = js & "    if (!r.status && cur !== 'None' && cur !== 'InProgress') { r.status = cur; r.doneAt = Date.now(); }"
    js = js & "  });"
    js = js & "  O.cur = id;"
    js = js & "  b.click();"
    js = js & "  return String(id);"
    js = js & "};"

    '結果の確定(完了後、settleMs経過したらログをスナップショットして購読解除)
    js = js & "O.finalize = (r) => {"
    js = js & "  if (r.final) return r.final;"
    js = js & "  const s = O.store().getState();"
    js = js & "  const outputs = s.bottomPanel.outputLogs.slice(r.logStart).filter(l => l && l.message !== undefined).map(l => fmt(l.message));"
    js = js & "  const problems = (s.bottomPanel.problemsLogs || []).map(p => fmt(p && p.message !== undefined ? p.message : p));"
    js = js & "  const errors = s.bottomPanel.outputLogs.slice(r.logStart).filter(l => l && l.message !== undefined && l.severity === 'error').map(l => fmt(l.message));"
    js = js & "  r.final = { outputs: outputs, problems: problems, errors: errors };"
    js = js & "  if (O.cur === r.id && O.unsub) { O.unsub(); O.unsub = null; }"
    js = js & "  return r.final;"
    js = js & "};"
    js = js & "O.result = (id) => {"
    js = js & "  const r = O.runs[id]; if (!r) return 'NG: unknown ticket ' + id;"
    js = js & "  if (!r.status) return '';"
    js = js & "  const now = Date.now();"
    js = js & "  const settled = (r.messages.length >= 2 && r.lastMsgAt >= r.doneAt && now - r.lastMsgAt >= 200) || now - r.doneAt >= O.settleMs;"
    js = js & "  if (!r.final && !settled) return '';"
    js = js & "  O.finalize(r);"
    js = js & "  return r.status;"
    js = js & "};"
    js = js & "O.outputs = (id) => { const r = O.runs[id]; return (r && r.final) ? r.final.outputs.join('\u001e') : ''; };"
    js = js & "O.logs = (id) => {"
    js = js & "  const r = O.runs[id]; if (!r) return 'NG: unknown ticket ' + id;"
    js = js & "  const f = r.final || { outputs: [], problems: [] };"
    js = js & "  const lines = ['ticket: ' + id + ' / status: ' + (r.status || 'pending'), 'history: ' + r.history.join(' -> ')];"
    js = js & "  for (const x of f.outputs) lines.push('[output] ' + x);"
    js = js & "  for (const x of f.problems) lines.push('[problem] ' + x);"
    js = js & "  for (const x of r.messages) lines.push('[messageBar] ' + x);"
    js = js & "  return lines.join('\n');"
    js = js & "};"
    js = js & "O.startedAt = (id) => { const r = O.runs[id]; return r ? String(r.startedAt) : ''; };"
    js = js & "O.firstError = (id) => { const r = O.runs[id]; if (!r || !r.final || !r.final.errors.length) return ''; const m = r.final.errors[0]; return JSON.stringify({ time: r.doneAt, name: '', message: m, stack: '', raw: m }); };"
    js = js & "O.errorLines = (id) => { const r = O.runs[id]; return (r && r.final) ? r.final.errors.join('\n') : ''; };"

    js = js & "window.__exCDP_OfficeScript = O;"
    js = js & "return 'installed';"
    js = js & "})()"
    BuildHelperJS = js
End Function

'***************************************************************************************************
'* 機能　　：「Office JS Sandbox」タブへ注入する`console.error`フック(window.__exCDP_OfficeScriptSandbox)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Sandbox側は、スクリプトの例外を次の形式で`console.error`に出す(globalConsole.error経由)
'              Errorインスタンス : "Command Error:\n{name}\n{cause} {message}\n{stack}"
'              それ以外          : "Command Error:\n{JSON}"   ※JSONは {"message","name","stack"} 等
'            これを {time, name, message, stack, raw} に分解して直近100件を保持する。元の`console.error`も呼ぶ
'
'            first(t)   時刻t以降で最初の"Command Error"(無ければ最初の1件)をJSON文字列で。無ければ空文字
'            lines(t)   時刻t以降の全件を「name: message」の改行区切りで
'***************************************************************************************************
Private Function BuildSandboxHookJS() As String
    Dim js As String
    js = js & "(() => {"
    js = js & "const V = '" & SANDBOX_HELPER_VERSION & "', P = window.__exCDP_OfficeScriptSandbox;"
    js = js & "if (P && P.version === V) return 'installed';"
    '旧版のフックが残っていたら、ネイティブの console.error から掛け直す(二重ラップ防止)
    js = js & "const S = { version: V, native: P ? P.native : console.error, errors: [] };"
    js = js & "const str = a => { if (typeof a === 'string') return a; if (a instanceof Error) return a.stack || String(a); try { return JSON.stringify(a); } catch (e) { return String(a); } };"
    js = js & "const parse = (text) => {"
    js = js & "  const d = { time: Date.now(), name: '', message: '', stack: '', raw: text };"
    js = js & "  const m = /^Command Error:\n([\s\S]*)$/.exec(text);"
    js = js & "  if (!m) { d.message = text; return d; }"
    js = js & "  try {"
    js = js & "    const o = JSON.parse(m[1]);"
    js = js & "    if (o && typeof o === 'object') { d.name = String(o.name || ''); d.message = (typeof o.message === 'string') ? o.message : str(o.message !== undefined ? o.message : o); d.stack = String(o.stack || ''); }"
    js = js & "    else d.message = str(o);"
    js = js & "    return d;"
    js = js & "  } catch (e) {}"
    js = js & "  const l = m[1].split('\n');"
    js = js & "  d.name = l[0] || ''; d.message = (l[1] || '').replace(/^undefined /, ''); d.stack = l.slice(2).join('\n');"
    js = js & "  return d;"
    js = js & "};"
    js = js & "console.error = function (...args) {"
    js = js & "  try { S.errors.push(parse(args.map(str).join(' '))); if (S.errors.length > 100) S.errors.shift(); } catch (e) {}"
    js = js & "  return S.native.apply(this, args);"
    js = js & "};"
    js = js & "S.list = (t) => S.errors.filter(e => e.time >= Number(t));"
    js = js & "S.first = (t) => { const l = S.list(t); const c = l.find(e => /^Command Error:/.test(e.raw)) || l[0]; return c ? JSON.stringify(c) : ''; };"
    js = js & "S.lines = (t) => S.list(t).map(e => (e.name ? e.name + ': ' : '') + e.message).join('\n');"
    js = js & "window.__exCDP_OfficeScriptSandbox = S;"
    js = js & "return 'installed';"
    js = js & "})()"
    BuildSandboxHookJS = js
End Function



'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
Private Sub sleep(Optional seconds As Double = 0.5)
    CDPHelpers.Sleep seconds
End Sub

Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Const LogFileExtension As String = ".cdp.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, CDPHelpers.currentLog.logPath
End Sub



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：bas側で接続済みの`CDPBrowser.cls`を継承し、「コード エディター」を確保します
'---------------------------------------------------------------------------------------------------
'* 返り値　：コードエディターを確保できたか
'* 引数　　：Inheritance    接続済み(reattach済み)の`CDPBrowser`
'            WebMode        False(既定)：デスクトップ版Excel。OfficeアドインのWebView2に接続した`CDPBrowser`を渡す
'                           True       ：Web版Excel。Excelを開いているEdgeに接続した`CDPBrowser`を渡す
'            EditorKey      コードエディターの探し方。省略時は次の既定値
'                             デスクトップ版：タブタイトル(部分一致)「コード エディター」。英語UIは "Code Editor" 等
'                             Web版        ：iframeのURL(部分一致)「index.html#/edit」
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`CDPBrowser`の接続/切断はbas側の責務です。このClassは切断しません
'              (切断前に`Dispose`を呼ぶか、このClassを破棄すること)
'            ・Web版のコードエディターはタブではなく、Excelページ内のiframe(別プロセス)として存在する。
'              Web版Excelを複数開いている場合は、最初に見つかったものを確保する
'            ・デスクトップ版のみ、例外詳細の回収用に「Office JS Sandbox」タブも確保する(見つからなくても続行)
'***************************************************************************************************
Friend Function Init(Inheritance As CDPBrowser, Optional WebMode As Boolean = False, Optional ByVal EditorKey As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".Init"


    '0. 確保済みなら一旦手放す
    Dispose
    If Inheritance Is Nothing Then
        printMsg WARN_, "`CDPBrowser` is Nothing. Connect it on the bas side first.", FromProcedureName
        Exit Function
    End If

    '1. 既定値の適用
    If Len(EditorKey) = 0 Then EditorKey = IIf(WebMode, DEFAULT_WEB_EDITOR_URL, DEFAULT_EDITOR_TITLE)
    printMsg info_, "Looking for the Office Scripts code editor (" & IIf(WebMode, "web", "desktop") & "): " & EditorKey, FromProcedureName

    On Error GoTo Failed

    '2. 継承
    Set ex_Browser = Inheritance

    '3. コードエディターを確保(デスクトップ版はタブをタイトルで、Web版はiframeをURLで)
    If WebMode Then
        Set ex_Editor = ex_Browser.getTab(Url:=EditorKey, SearchTypeID:=kFrame)
    Else
        Set ex_Editor = ex_Browser.getTab(EditorKey)
    End If
    If ex_Editor Is Nothing Then
        printMsg WARN_, "The code editor (`" & EditorKey & "`) was not found. Open a script in the code editor.", FromProcedureName
        ex_Browser.printTargetInfos
        GoTo Cleanup
    End If

    '4. ヘルパー注入 & エディター/Storeが掴めるか確認
    Dim ret As String: ret = EnsureHelper()
    If ret <> "OK" Then
        printMsg WARN_, "The code editor is not ready: " & ret, FromProcedureName
        GoTo Cleanup
    End If

    '5. 例外詳細の回収用に、Sandboxタブへも接続(デスクトップ版のみ。失敗しても続行)
    m_IsWeb = WebMode
    If Not WebMode Then ConnectSandbox

    printMsg info_, "The Office Scripts code editor is ready (" & IIf(WebMode, "web", "desktop") & ").", FromProcedureName
    Init = True
    Exit Function

Failed:
    printMsg WARN_, "Init failed. Error " & Err.Number & ": " & Err.Description, FromProcedureName
    Resume Cleanup

Cleanup:
    On Error GoTo 0
    Dispose
End Function

'***************************************************************************************************
'* 機能　　：初期設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    m_LogID = "OSID" & Format(Rnd * 1000, "000")
    Set m_LastOutputs = New Collection
End Sub

'***************************************************************************************************
'* 機能　　：破棄時は、コードエディター/Sandboxへの参照を手放します(`CDPBrowser`は切断しません)
'***************************************************************************************************
Private Sub Class_Terminate()
    Dispose
End Sub
