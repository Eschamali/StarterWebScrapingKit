VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_OfficeJs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           `CDPBrowser.cls` に対する機能拡張 - Office.js ダイレクト実行（OfficeScriptViaVBA）
'***************************************************************************************************
'* 機能　　：Excel の「コード エディター」(Officeスクリプト)のページには Office.js が読み込まれていて、Excel 本体と
'            つながっています。そのページで `Excel.run` を直接呼び、VBA から Office.js の API を使います
'---------------------------------------------------------------------------------------------------
'* 使い方　：
'    '1. bas側で`CDPBrowser`を用意(Demo_OfficeScript の ConnectDesktopExcel / ConnectWebExcel を参照)
'    Dim ojs As New exCDP_OfficeJs
'    If ojs.Init(Browser, WebMode:=False) Then
'        '2. `Excel.run(async (context) => { ... })` の中身を渡す。return した値が文字列で返る
'        If ojs.Run("const r = context.workbook.getSelectedRange(); r.load('address'); await context.sync(); return r.address;") Then
'            Debug.Print ojs.LastResult
'        Else
'            Debug.Print ojs.LastError
'        End If
'    End If
'
'    '(演出) ojs.ShowInEditor = True にすると、Run のたびに実行するコードをエディターに表示する
'
'    '3. 後片付け(先にこのClassを手放してから、bas側で切断)
'    ojs.Dispose
'    Browser.DisconnectCDP
'---------------------------------------------------------------------------------------------------
'* exCDP_OfficeScript との違い：
'    ・エディターに書き込んで「実行」ボタンを押すのではなく、ページの中の Office.js を直接呼びます。
'      Officeスクリプトとしての保存や実行の手順を通らないため、1回あたりミリ秒～数百ミリ秒で終わります
'    ・デスクトップ版で VBA が同期で待っていても止まりません(exCDP_OfficeScript では、読み取りの多いスクリプトが
'      VBA の待機中に進まないことがありました)
'    ・書くのは Officeスクリプト(ExcelScript)ではなく Office.js(Excel JavaScript API)です。
'      値を読むときは `load` してから `await context.sync()` が要ります
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - 事前に、Excel でスクリプトを開き「コード エディター」を表示しておくこと(エディターの中身は書き換えない)
'    - Microsoft が公開している使い方ではありません。コードエディターのページに Office.js が読み込まれていることを
'      借りているため、Office 側の更新で使えなくなる可能性があります
'    - デスクトップ版は「コード エディター」以外に、「Office JS Sandbox」「Excel Copilot Chat Add-in」「操作レコーダー」の
'      タブでも同じように Excel.run が通ることを確認しています(TargetKey でタブのタイトルを指定)
'    - `CDPBrowser`の接続/切断はbas側の責務。このClassは切断しない。`DisconnectCDP`の「前」に`Dispose`(または破棄)すること
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private ex_Browser      As CDPBrowser
Private ex_Target       As CDPContext   'Office.js が読み込まれたページ(既定はコードエディター)

'内部使用用
Private Const ThisClassName         As String = "exCDP_OfficeJs"   'ログのトレース用

'Office.js があるページの探し方
Private Const DEFAULT_DESKTOP_TITLE As String = "コード エディター"     'デスクトップ版：タブのタイトル(部分一致)
Private Const DEFAULT_WEB_URL       As String = "index.html#/edit"     'Web 版：iframe の URL(部分一致)
Private Const DEFAULT_TIMEOUT_SEC   As Double = 60

'ログ設定
Private m_LogID                     As String   'ログ識別ID

'接続中のルート
Private m_IsWeb                     As Boolean

'演出：Run のたびに、実行するコードをエディターに表示する
Private m_ShowInEditor              As Boolean
Private m_EditorCaption             As String

'直近の結果
Private m_LastResult                As String
Private m_LastError                 As String
Private m_LastElapsedSec            As Double



'***************************************************************************************************
'                                  ■■■ 接続状態 ■■■
'***************************************************************************************************
'* 機能　　：Office.js のページへの参照を手放します(`CDPBrowser`の切断はしません)
'***************************************************************************************************
Public Sub Dispose()
    Set ex_Target = Nothing
    Set ex_Browser = Nothing
    m_IsWeb = False
End Sub

'***************************************************************************************************
'* 機能　　：`Init`済み(Office.js のページを確保済み)かどうか
'***************************************************************************************************
Public Property Get IsConnected() As Boolean
    IsConnected = Not ex_Target Is Nothing
End Property

'***************************************************************************************************
'* 機能　　：Web版(`Init`でWebMode:=True)として初期化したかどうか
'***************************************************************************************************
Public Property Get IsWeb() As Boolean
    IsWeb = m_IsWeb
End Property



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
'                           ■■■ メインとなる機能拡張部分編 ■■■
'***************************************************************************************************
'* 機能　　：`Excel.run(async (context) => { ExcelRunBody })` を実行し、完了まで待ちます
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功したら True。結果は`LastResult`、失敗の理由は`LastError`
'* 引数　　：ExcelRunBody   `Excel.run`に渡す関数の中身。`context`(RequestContext)が使えます。
'                           return した値が、文字列ならそのまま、それ以外は JSON 文字列にして`LastResult`に入ります
'            TimeoutSec     待つ上限(秒)。JavaScript 側でこの時間を過ぎたら "NG: timeout" で打ち切ります
'---------------------------------------------------------------------------------------------------
'* 注意事項：・Office.js のエラーは、`LastError`に "NG: <code または message> <debugInfo>" の形で入ります
'            ・ExcelRunBody は、そのまま JavaScript として埋め込みます。外から受け取った文字列を混ぜないこと
'***************************************************************************************************
Public Function Run(ExcelRunBody As String, Optional TimeoutSec As Double = DEFAULT_TIMEOUT_SEC) As Boolean
    Const FromProcedureName As String = ThisClassName & ".Run"


    m_LastResult = ""
    m_LastError = ""
    m_LastElapsedSec = 0
    If Not IsConnected Then
        m_LastError = "NG: Not initialized. Call `Init` first."
        printMsg WARN_, m_LastError, FromProcedureName
        Exit Function
    End If

    '0. 演出：実行するコードをエディターに表示(結果は待たない)
    If m_ShowInEditor Then ShowCodeInEditor ExcelRunBody

    '1. Excel.run で包む(タイムアウトと例外は、JavaScript 側で "NG: ..." の文字列にして返す)
    Dim js As String
    js = Replace(WrapperJS(), "__TIMEOUT_MS__", CStr(CLng(TimeoutSec * 1000)))
    js = Replace(js, "__BODY__", ExcelRunBody)

    '2. 実行して、Promise の完了まで待つ
    '   ※CDPContextのコマンド待ちは既定30秒で打ち切られるため、TimeoutSec に合わせて一時的に延ばす
    Dim StartSec As Double: StartSec = Timer
    Dim v As Variant, ErrNum As Long, ErrDesc As String
    ex_Target.TimeOutSecond = TimeoutSec + 10
    On Error Resume Next
    v = ex_Target.jsEval(js, awaitPromise:=True)
    ErrNum = Err.Number: ErrDesc = Err.Description
    On Error GoTo 0
    ex_Target.TimeOutSecond = TIMEOUT_DEFAULT
    If ErrNum <> 0 Then Err.Raise ErrNum, FromProcedureName, ErrDesc
    m_LastElapsedSec = Timer - StartSec
    If m_LastElapsedSec < 0 Then m_LastElapsedSec = m_LastElapsedSec + 86400#   '日付またぎ

    '3. 結果
    If IsError(v) Then
        m_LastError = "NG: JavaScript exception"
        If Not ex_Target.LastJavaScriptException Is Nothing Then m_LastError = m_LastError & ": " & ex_Target.LastJavaScriptException.Stringify
    ElseIf IsObject(v) Then
        m_LastError = "NG: unexpected object result"
    ElseIf IsNull(v) Or IsEmpty(v) Then
        Run = True
    ElseIf Left$(CStr(v), 3) = "NG:" Then
        m_LastError = CStr(v)
    Else
        m_LastResult = CStr(v)
        Run = True
    End If

    If Not Run Then printMsg WARN_, m_LastError, FromProcedureName
End Function

'***************************************************************************************************
'                             ■■■ 演出(エディターへの表示) ■■■
'***************************************************************************************************
'* 機能　　：True にすると、`Run`のたびに、実行するコードをコードエディターに表示します(既定は False)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：見た目のための演出です。実際の実行はエディターではなく、Office.js で直接行います。
'            表示は結果を待たない CDP(`jsEval`の`RunAsyncCDP:=True`)で送るため、`Run`の速さはほぼ変わりません
'* 注意事項：・エディターの中身は表示用のコードで上書きされ、Officeスクリプトとして自動保存されます
'            ・表示するのは Office.js のコードなので、Officeスクリプトのルール(引数の型など)に合わず、
'              エディターの「問題」に警告が出ます。実行には影響しません
'            ・Init で「コード エディター」以外のページ(Sandbox など)を選んだ場合は、表示されません
'***************************************************************************************************
Public Property Let ShowInEditor(Flag As Boolean)
    m_ShowInEditor = Flag
End Property
Public Property Get ShowInEditor() As Boolean
    ShowInEditor = m_ShowInEditor
End Property

'***************************************************************************************************
'* 機能　　：エディターに表示するコードの見出し(1行)。`Run`の前に設定すると、表示の3行目に入ります
'***************************************************************************************************
Public Property Let EditorCaption(Caption As String)
    m_EditorCaption = Caption
End Property
Public Property Get EditorCaption() As String
    EditorCaption = m_EditorCaption
End Property



'***************************************************************************************************
'                                ■■■ 直近の結果 ■■■
'***************************************************************************************************
'* 機能　　：直近の`Run`の結果(return した値。文字列以外は JSON)
'***************************************************************************************************
Public Property Get LastResult() As String
    LastResult = m_LastResult
End Property

'***************************************************************************************************
'* 機能　　：直近の`Run`が失敗した理由("NG: ...")。成功したら空文字
'***************************************************************************************************
Public Property Get LastError() As String
    LastError = m_LastError
End Property

'***************************************************************************************************
'* 機能　　：直近の`Run`に VBA から見てかかった時間(秒)
'***************************************************************************************************
Public Property Get LastElapsedSec() As Double
    LastElapsedSec = m_LastElapsedSec
End Property



'***************************************************************************************************
'                                ■■■ 内部処理 ■■■
'***************************************************************************************************
'* 機能　　：`Excel.run`を包む JavaScript。__TIMEOUT_MS__ と __BODY__ を差し替えて使う
'***************************************************************************************************
Private Function WrapperJS() As String
    Dim s As String
    s = s & "(() => {" & vbLf
    s = s & "    const TIMEOUT_MS = __TIMEOUT_MS__;" & vbLf
    s = s & "    if (typeof Excel === ""undefined"" || typeof Excel.run !== ""function"") { return Promise.resolve(""NG: Office.js (Excel.run) is not available on this page""); }" & vbLf
    s = s & "    const run = Excel.run(async (context) => {" & vbLf
    s = s & "__BODY__" & vbLf
    s = s & "    });" & vbLf
    s = s & "    const guard = new Promise((resolve) => setTimeout(() => resolve(""NG: timeout "" + TIMEOUT_MS + "" ms""), TIMEOUT_MS));" & vbLf
    s = s & "    return Promise.race([run.then((r) => (r === undefined || r === null) ? """" : (typeof r === ""string"" ? r : JSON.stringify(r))), guard])" & vbLf
    s = s & "        .catch((e) => ""NG: "" + String((e && (e.code || e.message)) || e) + ((e && e.debugInfo) ? "" "" + JSON.stringify(e.debugInfo) : """"));" & vbLf
    s = s & "})()" & vbLf
    WrapperJS = s
End Function


'***************************************************************************************************
'* 機能　　：演出用。実行するコードを、見出しを付けてコードエディターに表示します(結果は待たない)
'***************************************************************************************************
Private Sub ShowCodeInEditor(ExcelRunBody As String)
    Dim Text As String
    Text = "// ===== VBA -> CDP -> Office.js (Excel.run) =====" & vbLf & _
           "// 表示用です。実行はこのエディターではなく、Office.js で直接行っています" & vbLf & _
           "// " & m_EditorCaption & IIf(Len(m_EditorCaption), "  ", "") & Format$(Now, "yyyy/mm/dd hh:nn:ss") & vbLf & _
           "Excel.run(async (context) => {" & vbLf & _
           ExcelRunBody & vbLf & _
           "});" & vbLf

    Dim v As Variant: v = Text
    On Error Resume Next
    ex_Target.jsEval Replace(ShowJS(), "__TEXT__", WebJsonConverter.serialize(v)), RunAsyncCDP:=True
    On Error GoTo 0
End Sub

'***************************************************************************************************
'* 機能　　：演出用。Monaco エディターの全文を置き換える JavaScript。__TEXT__ に JSON 文字列リテラルを入れて使う
'***************************************************************************************************
Private Function ShowJS() As String
    Dim s As String
    s = s & "(() => {" & vbLf
    s = s & "    const W = window, K = ""__exCDP_OfficeJsShow"";" & vbLf
    s = s & "    const isEd = (v) => !!v && typeof v === ""object"" && typeof v.getModel === ""function"" && typeof v.executeEdits === ""function"";" & vbLf
    s = s & "    let ed = W[K];" & vbLf
    s = s & "    if (!isEd(ed) || !ed.getModel()) {" & vbLf
    s = s & "        ed = null;" & vbLf
    s = s & "        const fiberOf = (el) => { const k = Object.keys(el).find((x) => x.startsWith(""__reactFiber$"")); return k ? el[k] : null; };" & vbLf
    s = s & "        for (let el = document.querySelector("".monaco-editor""); el && !ed; el = el.parentElement) {" & vbLf
    s = s & "            for (let f = fiberOf(el), d = 0; f && d < 40 && !ed; f = f.return, d++) {" & vbLf
    s = s & "                const sn = f.stateNode;" & vbLf
    s = s & "                if (sn && !(sn instanceof Node)) { if (isEd(sn.editor)) { ed = sn.editor; } else if (isEd(sn)) { ed = sn; } }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if (!ed || !ed.getModel()) { return ""NG: editor not found""; }" & vbLf
    s = s & "        W[K] = ed;" & vbLf
    s = s & "    }" & vbLf
    s = s & "    const m = ed.getModel(), t = __TEXT__;" & vbLf
    s = s & "    if (m.getValue() === t) { return ""same""; }" & vbLf
    s = s & "    ed.executeEdits(""exCDP_OfficeJs"", [{ range: m.getFullModelRange(), text: t, forceMoveMarkers: true }]);" & vbLf
    s = s & "    ed.setScrollTop(0);" & vbLf
    s = s & "    return ""OK"";" & vbLf
    s = s & "})()" & vbLf
    ShowJS = s
End Function


'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Const LogFileExtension As String = ".cdp.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, CDPHelpers.currentLog.logPath
End Sub



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：bas側で接続済みの`CDPBrowser.cls`を継承し、Office.js が読み込まれたページを確保します
'---------------------------------------------------------------------------------------------------
'* 返り値　：Office.js(`Excel.run`)が使えるページを確保できたか
'* 引数　　：Inheritance    接続済み(reattach済み)の`CDPBrowser`
'            WebMode        False(既定)：デスクトップ版Excel / True：Web版Excel(Edge)
'            TargetKey      ページの探し方。省略時は次の既定値
'                             デスクトップ版：タブのタイトル(部分一致)「コード エディター」
'                             Web版        ：iframeのURL(部分一致)「index.html#/edit」
'***************************************************************************************************
Friend Function Init(Inheritance As CDPBrowser, Optional WebMode As Boolean = False, Optional ByVal TargetKey As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".Init"


    '0. 確保済みなら一旦手放す
    Dispose
    If Inheritance Is Nothing Then
        printMsg WARN_, "`CDPBrowser` is Nothing. Connect it on the bas side first.", FromProcedureName
        Exit Function
    End If
    If Len(TargetKey) = 0 Then TargetKey = IIf(WebMode, DEFAULT_WEB_URL, DEFAULT_DESKTOP_TITLE)

    On Error GoTo Failed

    '1. ページを確保(デスクトップ版はタブをタイトルで、Web版はiframeをURLで)
    Set ex_Browser = Inheritance
    If WebMode Then
        Set ex_Target = ex_Browser.getTab(Url:=TargetKey, SearchTypeID:=kFrame)
    Else
        Set ex_Target = ex_Browser.getTab(TargetKey)
    End If
    If ex_Target Is Nothing Then
        printMsg WARN_, "The page (`" & TargetKey & "`) was not found. Open a script in the code editor.", FromProcedureName
        GoTo Cleanup
    End If

    '2. Office.js(Excel.run)が使えるか確認
    Dim v As Variant
    v = ex_Target.jsEval("String(typeof Excel !== 'undefined' && typeof Excel.run === 'function')")
    If IsError(v) Then GoTo NotAvailable
    If CStr(v) <> "true" Then GoTo NotAvailable

    m_IsWeb = WebMode
    printMsg info_, "Office.js is ready on `" & TargetKey & "` (" & IIf(WebMode, "web", "desktop") & ").", FromProcedureName
    Init = True
    Exit Function

NotAvailable:
    printMsg WARN_, "Office.js (Excel.run) is not available on `" & TargetKey & "`.", FromProcedureName
    GoTo Cleanup

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
    m_LogID = "OJID" & Format(Rnd * 1000, "000")
End Sub

'***************************************************************************************************
'* 機能　　：破棄時は、ページへの参照を手放します(`CDPBrowser`は切断しません)
'***************************************************************************************************
Private Sub Class_Terminate()
    Dispose
End Sub
