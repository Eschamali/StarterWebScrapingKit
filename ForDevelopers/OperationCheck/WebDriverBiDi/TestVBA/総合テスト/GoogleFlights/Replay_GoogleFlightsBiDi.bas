Attribute VB_Name = "Replay_GoogleFlightsBiDi"
'===================================================================================================
' Google Flights 手動操作ログ(BiDi Discovery Log) を Kit の BiDi モードで再現
'   ※Kit(CDP)版 Replay_GoogleFlights.bas の移植。構成・関数名・待機ロジックは CDP 版と同一にし、
'     CDP 固有の部分だけを WebDriver BiDi に置き換えている（CDP 版との比較用）
'---------------------------------------------------------------------------------------------------
' 操作の流れ（ログの [USER-ACTION] seq 番号）
'   0. https://www.google.com/travel/flights へ遷移                         [NAV-TARGET]
'   1. 種別コンボ(Round trip) → "One way"                                   seq3, seq5
'   2. "Where from?" → ダイアログに入力 → 候補 "Sapporo"                     seq7, (入力), seq9
'   3. "Where to?"   → ダイアログに入力 → 候補 "Paris, France"               seq14, (入力), seq16
'   4. Search → 日付ダイアログ → 2026-11-22 のセル → Done                    seq21, seq23, seq24
'   5. Search → 検索結果("Top flights")                                     seq26, seq27
'
' 待機の考え方（CDP 版と同一）
'   ・ネットワーク信号 : 操作「前」に Arm → 応答開始(2xx) の URL 部分一致
'                        → 同じ request ID の応答完了 で HIT
'   ・ネットワーク静穏 : 追跡対象の通信を beforeRequestSent ～ responseCompleted/fetchError で
'                        request ID ごとに数え、0件か（アイドル除外パターンは数えない）
'   ・DOM 静穏         : ページに MutationObserver を仕込み、最終変化からの経過 ms を読む
'   ・安定判定         : 全信号HIT + 通信0件 + DOM静穏>=minStableMs の状態が、
'                        さらに minStableMs 連続したら STABLE
'
' CDP 版からの置き換え
'   ・起動        : CDPContext.OpenBrowserAndConnectTab → WebDriverBiDiContext.OpenBrowserAndConnectTab
'   ・通信イベント: Network.requestWillBeSent / responseReceived / loadingFinished / loadingFailed
'                 → network.beforeRequestSent / responseStarted / responseCompleted / fetchError
'   ・追跡対象    : resourceType(XHR|Fetch|Document) → request.destination / initiatorType で判定
'                   （Kit 内蔵 mapper の変換規則：Document→"document"/"iframe"、Fetch→initiatorType="fetch"、
'                     XHR は destination="" かつ initiatorType=null。後者は Ping/Media 等も含むため近似）
'   ・処理順      : BiDi はリダイレクトの途中でも responseCompleted が来るため、メソッド別ではなく
'                   受信順（Kit が付与する __index__）でイベントを処理する
'   ・クリック    : CDPElement.sendClick(Input.dispatchMouseEvent) → WebDriverBiDiElement.sendClick(input.performActions)
'   ・入力        : CDPElement.sendString(Input.insertText) → WebDriverBiDiElement.sendString(1文字ずつ keyDown/keyUp)
'   ・追跡範囲    : 本タブの context の通信のみ（子 iframe・Worker は数えない）
'
' 前提・注意
'   ・m_Ctx.ThisWebDriverBiDiMode.BiDiEvents（キャプチャ用）を本モジュールが占有し、取り出すたびに空にします。
'===================================================================================================
Option Explicit

'---------------------------------------------------------------------------------------------------
' 入力値（CDP 版と同一）
'   ログは入力文字列そのものを記録していない（valueLen のみ）。
'   ・出発地: change 時 valueLen=7、選択候補 aria-label="Sapporo"、H028ib 応答の推移(New York →
'     San Francisco → San Pedro Sula → Sapporo)から "Sapporo" と判断
'   ・目的地: valueLen=5、選択候補 aria-label="Paris, France" から "Paris" と判断
'   違っていればここを直してください。
'---------------------------------------------------------------------------------------------------
'記録は英語UIで行い、XPath も英語の aria-label 前提。ブラウザの言語設定に左右されないよう hl=en で固定
Private Const START_URL      As String = "https://www.google.com/travel/flights?hl=en"
Private Const ORIGIN_TEXT    As String = "Sapporo"
Private Const ORIGIN_OPTION  As String = "Sapporo"
Private Const DEST_TEXT      As String = "Paris"
Private Const DEST_OPTION    As String = "Paris, France"
Private Const DEPART_ISO     As String = "2026-11-22"   'ログの gridcell data-iso。過去日になったら変更

'---------------------------------------------------------------------------------------------------
' XPath（getElementByXPath が JS のダブルクォートで包むため、XPath 内はシングルクォートのみ）
'   動的な id（i6, c222, h0T7hb-9 等）は使わず、aria-label / role / data-iso で特定
'---------------------------------------------------------------------------------------------------
Private Const XP_TICKET_TYPE   As String = "//div[@role='combobox'][.//span[@aria-label='Change ticket type.']]"
Private Const XP_ONE_WAY       As String = "//ul[@role='listbox'][@aria-label='Select your ticket type.']/li[@role='option'][contains(., 'One way')]"
Private Const XP_FROM_MAIN     As String = "//input[@aria-label='Where from?'][not(ancestor::div[@role='dialog'])]"
Private Const XP_FROM_DIALOG   As String = "//div[@role='dialog'][@aria-label='Enter your origin']//input[@role='combobox']"
Private Const XP_TO_MAIN       As String = "//input[starts-with(@aria-label,'Where to?')][not(ancestor::div[@role='dialog'])]"
Private Const XP_TO_DIALOG     As String = "//div[@role='dialog'][@aria-label='Enter your destination']//input[@role='combobox']"
Private Const XP_SEARCH        As String = "//button[@aria-label='Search']"
Private Const XP_DONE          As String = "//button[starts-with(@aria-label,'Done.')]"
Private Const XP_TOP_FLIGHTS   As String = "//h3[normalize-space(.)='Top flights']"

'---------------------------------------------------------------------------------------------------
' ネットワーク関連（CDP 版と同一）
'---------------------------------------------------------------------------------------------------
'アイドル判定から除外（ログで idle=skip のもの + SPA-CONSENSUS-SUSPECT が付いた browserinfo）
Private Const IDLE_IGNORE As String = "www.google.com/log?|play.google.com/log?|/generate_204|ogads-pa.clients6.google.com/|/FlightsFrontendUi/browserinfo"

'購読する BiDi イベント
Private Const EV_BEFORE    As String = "network.beforeRequestSent"
Private Const EV_STARTED   As String = "network.responseStarted"
Private Const EV_COMPLETED As String = "network.responseCompleted"
Private Const EV_FAILED    As String = "network.fetchError"

'---------------------------------------------------------------------------------------------------
' 状態
'---------------------------------------------------------------------------------------------------
Private m_Ctx       As WebDriverBiDiContext
Private m_Inflight  As Dictionary   'request ID -> url（アイドル判定対象で未完了のもの）
Private m_ArmHit    As Dictionary   'pattern    -> Boolean（HIT済みか）
Private m_ArmReq    As Dictionary   'request ID -> pattern（URL一致済み・完了待ち）
Private m_T0        As Double       'ログ表示用の基準時刻(ms)

'===================================================================================================
' メイン
'===================================================================================================
Public Sub ReplayGoogleFlightsBiDi()
    Set m_Ctx = StartContextWithArgs("--disable-blink-features=AutomationControlled")
    m_T0 = CDPHelpers.TimerCounter

    '0. 遷移前にネットワーク監視を開始（遷移中の通信も数えるため）
    NetInit
    '起動引数が効いているかの確認（AutomationControlled を無効化できていれば false / undefined）
    Debug.Print Stamp & "[CHECK] navigator.webdriver=" & SafeJs("String(navigator.webdriver)", "(exception)")

    '遷移。エラーページ(chrome-error://)や BiDi 側のエラーなら1回だけ再試行する
    Dim attempt As Long, navErr As String
    For attempt = 1 To 2
        On Error Resume Next
        m_Ctx.navigate START_URL
        navErr = Err.Description: Err.Clear
        On Error GoTo 0
        If LenB(navErr) = 0 And Left$(SafeJs("location.href", ""), 15) <> "chrome-error://" Then Exit For
        Debug.Print Stamp & "[NAV-ERROR] attempt=" & attempt & " | " & IIf(LenB(navErr), navErr, "chrome-error://")
        If attempt = 2 Then Err.Raise vbObjectError + 514, "ReplayGoogleFlightsBiDi", "ページを読み込めませんでした: " & START_URL
        CDPHelpers.Sleep 2
    Next
    WaitSettled 500, 10000

    '1. 種別 → One way
    ClickVisible XP_TICKET_TYPE
    WaitSettled 500, 10000
    ClickVisible XP_ONE_WAY
    WaitSettled 500, 10000

    '2. 出発地
    SelectPlace XP_FROM_MAIN, XP_FROM_DIALOG, "Enter your origin", ORIGIN_TEXT, ORIGIN_OPTION, "rpcids=tDoGIe"

    '3. 目的地（ログでは目的地確定時に tDoGIe と BVAT3 が並んで飛ぶ。BVAT3 を信号にする）
    SelectPlace XP_TO_MAIN, XP_TO_DIALOG, "Enter your destination", DEST_TEXT, DEST_OPTION, "rpcids=BVAT3"

    '4. Search（日付未指定なので日付ダイアログが開く）→ 日付 → Done
    Dim xpDate As String
    xpDate = "//div[@role='gridcell'][@data-iso='" & DEPART_ISO & "']//div[@role='button']"
    Arm "GetCalendarPicker"
    ClickVisible XP_SEARCH
    WaitSettled 500, 10000, xpDate

    ClickVisible xpDate
    WaitSettled 500, 10000
    ClickVisible XP_DONE
    WaitSettled 500, 10000

    '5. Search → 検索結果（今回のログでは要求から応答完了まで約14秒のため、上限は CDP 版の30秒より延ばす）
    Arm "GetShoppingResults"
    ClickVisible XP_SEARCH
    WaitSettled 500, 60000, XP_TOP_FLIGHTS

    Debug.Print Stamp & "[DONE] url=" & SafeJs("location.href", "") & " | title=" & SafeJs("document.title", "")
    'm_Ctx.ThisWebDriverBiDiMode.CloseBrowser
End Sub

'---------------------------------------------------------------------------------------------------
' 設定シートの値に起動引数を足して起動する（StartBiDiModeContext と同じ流れを、引数追加だけ自前で行う）
'   ・New BrowserCapabilities の時点で設定シートの「追加引数」等が読み込まれる
'   ・Set_AdditionalBrowserArguments は文字列で、コマンドライン末尾に連結される
'   ・シート側に同じ引数が既にあれば二重に付けない
'---------------------------------------------------------------------------------------------------
Private Function StartContextWithArgs(extraArgs As String) As WebDriverBiDiContext
    Dim caps As New BrowserCapabilities
    If InStr(1, caps.Set_AdditionalBrowserArguments, extraArgs, vbTextCompare) = 0 Then
        caps.Set_AdditionalBrowserArguments = Trim$(caps.Set_AdditionalBrowserArguments & " " & extraArgs)
    End If
    Debug.Print "[START] args=" & caps.Set_AdditionalBrowserArguments

    Set StartContextWithArgs = New WebDriverBiDiContext
    StartContextWithArgs.OpenBrowserAndConnectTab caps
End Function

'---------------------------------------------------------------------------------------------------
' 出発地/目的地の共通手順：入力欄クリック → ダイアログで入力 → 候補選択（CDP 版と同一）
'---------------------------------------------------------------------------------------------------
Private Sub SelectPlace(xpMain As String, xpDialogInput As String, dialogLabel As String, _
                        inputText As String, optionLabel As String, confirmPattern As String)
    Dim xpOption As String
    xpOption = "//div[@role='dialog'][@aria-label='" & dialogLabel & "']//li[@role='option'][@aria-label='" & optionLabel & "']"

    'ダイアログを開く（開いた直後に空文字の H028ib が1発飛ぶので、ここで静まるのを待つ）
    ClickVisible xpMain
    WaitSettled 500, 10000, xpDialogInput

    '入力 → 候補の表示を待つ（値で特定できる候補 li を表示ゲートにする）
    Arm "rpcids=H028ib"
    FindVisible(xpDialogInput).sendString inputText
    WaitSettled 500, 10000, xpOption

    '候補選択 → 確定時の通信を待つ（初回ログで minStableMs=1000 が効いた箇所）
    Arm confirmPattern
    ClickVisible xpOption
    WaitSettled 1000, 10000
End Sub

'===================================================================================================
' 要素操作
'===================================================================================================
'XPath に一致する要素のうち表示中の1件を返す（見つかるまで待機）
'  1周目: 表示中かつ aria-hidden 祖先なし / 2周目: 表示中のみ（BiDi版の PreferVisibleMatch 相当）
Private Function FindVisible(xp As String, Optional timeoutMs As Double = 10000) As WebDriverBiDiElement
    Dim js As String
    js = "(function(x){var r=document.evaluate(x,document,null,XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,null);" & _
         "function vis(e){if(!e.getClientRects().length)return false;var s=getComputedStyle(e);return s.visibility!=='hidden';}" & _
         "var i,e;for(i=0;i<r.snapshotLength;i++){e=r.snapshotItem(i);if(vis(e)&&!e.closest('[aria-hidden=""true""]'))return i+1;}" & _
         "for(i=0;i<r.snapshotLength;i++){if(vis(r.snapshotItem(i)))return i+1;}return 0;})(""" & xp & """)"

    Dim t0 As Double: t0 = CDPHelpers.TimerCounter
    Dim idx As Variant
    Do
        idx = m_Ctx.jsEval(js, IFEXCEPTION:=0, StopBiDiError:=False)
        If VarType(idx) = vbDouble Then If idx > 0 Then Exit Do
        If CDPHelpers.TimerCounter - t0 > timeoutMs Then
            Dim diag As String: diag = DiagXPath(xp)
            Debug.Print Stamp & "[LOCATE-MISS] lastResult=" & TypeName(idx) & " | " & diag
            Err.Raise vbObjectError + 513, "FindVisible", "表示中の要素が見つかりません: " & xp & vbCrLf & vbCrLf & _
                      "lastResult=" & TypeName(idx) & vbCrLf & Replace(diag, " | ", vbCrLf)
        End If
        PumpNetwork
        CDPHelpers.Sleep 0.1
    Loop

    Set FindVisible = m_Ctx.getElementByXPath("(" & xp & ")[" & CLng(idx) & "]")
End Function

'要素が見つからない時の診断情報：ページの言語・URL・XPath一致数・combobox の aria-label 一覧（CDP 版と同一）
Private Function DiagXPath(xp As String) As String
    Dim js As String
    js = "(function(x){var out=[];try{out.push('lang='+document.documentElement.lang);out.push('url='+location.href);" & _
         "var r=document.evaluate(x,document,null,XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,null);out.push('match='+r.snapshotLength);" & _
         "var c=document.querySelectorAll('[role=combobox]'),a=[];for(var i=0;i<c.length&&i<8;i++){" & _
         "var t=c[i].getAttribute('aria-label');if(!t){var l=c[i].getAttribute('aria-labelledby');" & _
         "if(l){t=l.split(' ').map(function(id){var n=document.getElementById(id);return n?(n.getAttribute('aria-label')||n.textContent):'';}).join('/');}}" & _
         "a.push(c[i].tagName+':'+t);}out.push('comboboxes('+c.length+')='+a.join(' ; '));" & _
         "}catch(e){out.push('error='+e.message);}return out.join(' | ');})(""" & xp & """)"
    DiagXPath = SafeJs(js, "(diag exception)")
End Function

'ログの操作は全て実クリック(pointerdown→click)なので、input.performActions 経由の sendClick を使う
Private Sub ClickVisible(xp As String)
    Debug.Print Stamp & "[CLICK] " & xp
    FindVisible(xp).sendClick
End Sub

'jsEval の結果を文字列で返す（JS 例外・BiDi エラー時は fallback）
Private Function SafeJs(js As String, fallback As String) As String
    Dim v As Variant
    v = m_Ctx.jsEval(js, IFEXCEPTION:=fallback, StopBiDiError:=False)
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then SafeJs = fallback Else SafeJs = CStr(v)
End Function

'===================================================================================================
' ネットワーク監視
'===================================================================================================
Private Sub NetInit()
    Set m_Inflight = New Dictionary
    Set m_ArmHit = New Dictionary
    Set m_ArmReq = New Dictionary

    'network イベントを購読し、Kit のイベント記録を開始
    m_Ctx.SubscribeBiDiEvent = Array(EV_BEFORE, EV_STARTED, EV_COMPLETED, EV_FAILED)
    Set m_Ctx.ThisWebDriverBiDiMode.BiDiEvents = New Dictionary
End Sub

'待ちたい通信を「操作の前に」登録する
Private Sub Arm(pattern As String)
    If m_ArmHit.Exists(pattern) Then m_ArmHit.Remove pattern
    m_ArmHit.Add pattern, False
    Debug.Print Stamp & "[ARM] net pattern=" & pattern
End Sub

Private Sub DisarmAll()
    m_ArmHit.RemoveAll
    m_ArmReq.RemoveAll
End Sub

'届いているイベントを取り出して状態に反映し、記録を空にする
'  ※BiDi はリダイレクト途中でも responseCompleted が来るため、メソッド別ではなく受信順で処理する
Private Sub PumpNetwork()
    With m_Ctx.ThisWebDriverBiDiMode
        .TakeEvents

        Dim total As Long: total = .BiDiEvents(EventsKey_TotalEvents)
        If total = 0 Then Exit Sub

        '1. 対象4種のイベントを受信順（__index__）に並べる
        Dim ordered() As Variant: ReDim ordered(1 To total)
        Dim methods As Dictionary: Set methods = .BiDiEvents(EventsKey_EventMethods)
        Dim m As Variant, ev As Variant
        For Each m In Array(EV_BEFORE, EV_STARTED, EV_COMPLETED, EV_FAILED)
            If methods.Exists(m) Then
                For Each ev In methods(m)
                    Set ordered(ev("__index__")) = ev
                Next
            End If
        Next

        '2. 記録を空にしてから処理（処理中に届いた分は次回に回る）
        Set .BiDiEvents = New Dictionary
    End With

    Dim i As Long
    For i = 1 To total
        If IsObject(ordered(i)) Then ProcessEvent ordered(i)
    Next
End Sub

Private Sub ProcessEvent(ev As Variant)
    Dim methodName As String: methodName = SafeStr(ev("method"))
    Dim p As Dictionary: Set p = ev("params")

    '本タブの context 以外（子 iframe・Worker 等）は数えない
    If SafeStr(p("context")) <> m_Ctx.context Then Exit Sub

    Dim rid As String, url As String
    rid = SafeStr(p("request")("request"))

    Select Case methodName
        Case EV_BEFORE
            url = SafeStr(p("request")("url"))
            'リダイレクトは同じ request ID で再度来るので、未登録の時だけ数える
            If IsTrackedRequest(p("request")) And Not IsIgnored(url) Then
                If Not m_Inflight.Exists(rid) Then m_Inflight.Add rid, url
            End If

        Case EV_STARTED
            url = SafeStr(p("response")("url"))
            '2xx の応答ヘッダーが届いたものだけ候補にする
            '（この時点で応答は返っているので、後の fetchError も HIT 扱いにできる）
            Dim st As Long: st = CLng(Val(SafeStr(p("response")("status"))))
            If st >= 200 And st < 300 Then
                Dim k As Variant
                For Each k In m_ArmHit.keys
                    If Not m_ArmHit(k) And Not m_ArmReq.Exists(rid) Then
                        If InStr(url, k) > 0 Then m_ArmReq.Add rid, k
                    End If
                Next
            End If

        Case EV_COMPLETED
            If m_ArmReq.Exists(rid) Then
                m_ArmHit(m_ArmReq(rid)) = True
                Debug.Print Stamp & "[HIT] net pattern=" & m_ArmReq(rid) & " | request=" & rid
                m_ArmReq.Remove rid
            End If
            If m_Inflight.Exists(rid) Then m_Inflight.Remove rid

        Case EV_FAILED
            '2xx 応答の後の ERR_ABORTED は HIT 扱い。
            '  batchexecute / GetShoppingResults はストリーム応答で、ページが必要分を読んだ後に
            '  打ち切ると ERR_ABORTED になる（BiDi 記録ログでも 200 と ERR_ABORTED が同じ id で並んでいる）。
            '  それ以外の失敗（応答ヘッダー前の失敗を含む）は従来どおり HIT にしない。
            Dim errText As String: errText = SafeStr(p("errorText"))
            If m_ArmReq.Exists(rid) Then
                If InStr(errText, "ERR_ABORTED") > 0 Then
                    m_ArmHit(m_ArmReq(rid)) = True
                    Debug.Print Stamp & "[HIT] net pattern=" & m_ArmReq(rid) & " | request=" & rid & " | aborted-after-2xx"
                Else
                    Debug.Print Stamp & "[FAIL] net pattern=" & m_ArmReq(rid) & " | " & errText
                End If
                m_ArmReq.Remove rid
            End If
            If m_Inflight.Exists(rid) Then m_Inflight.Remove rid
    End Select
End Sub

'CDP 版の resourceType(XHR|Fetch|Document) 相当を、BiDi の request.destination / initiatorType で判定
'  Document → destination="document"/"iframe" ／ Fetch → initiatorType="fetch"
'  XHR      → destination="" かつ initiatorType=null（Ping/Media/Other 等も含む近似）
Private Function IsTrackedRequest(req As Variant) As Boolean
    Dim dest As String, initType As String
    If req.Exists("destination") Then dest = SafeStr(req("destination"))
    If req.Exists("initiatorType") Then initType = SafeStr(req("initiatorType"))

    Select Case dest
        Case "document", "iframe": IsTrackedRequest = True
        Case "":                   IsTrackedRequest = (initType = "" Or initType = "fetch")
    End Select
End Function

Private Function IsIgnored(url As String) As Boolean
    Dim v As Variant
    For Each v In Split(IDLE_IGNORE, "|")
        If InStr(url, v) > 0 Then IsIgnored = True: Exit Function
    Next
End Function

Private Function AllArmedHit() As Boolean
    Dim k As Variant
    For Each k In m_ArmHit.keys
        If Not m_ArmHit(k) Then Exit Function
    Next
    AllArmedHit = True
End Function

'JSON の null や欠落でも落ちないよう文字列化する
Private Function SafeStr(v As Variant) As String
    If IsObject(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    SafeStr = CStr(v)
End Function

'===================================================================================================
' DOM 静穏プローブ（CDP 版と同一。jsEval のみ BiDi エラーで止めない指定を追加）
'===================================================================================================
'最終 DOM 変化からの経過 ms。プローブが無ければ（新しい document 等）仕込み直して 0 を返す
Private Function DomQuietMs() As Double
    Const JS_READ As String = "(function(){var p=window.__kitProbe;return p?performance.now()-p.last:-1;})()"
    Const JS_INSTALL As String = "(function(){if(window.__kitProbe)return 0;window.__kitProbe={last:performance.now()};" & _
        "new MutationObserver(function(){window.__kitProbe.last=performance.now();})" & _
        ".observe(document.documentElement,{subtree:true,childList:true,attributes:true,characterData:true});return 0;})()"

    Dim v As Variant
    v = m_Ctx.jsEval(JS_READ, IFEXCEPTION:=-1, StopBiDiError:=False)
    If VarType(v) <> vbDouble Then v = -1
    If v < 0 Then
        m_Ctx.jsEval JS_INSTALL, IFEXCEPTION:=0, StopBiDiError:=False
        Debug.Print Stamp & "[PROBE] installed"
        v = 0
    End If
    DomQuietMs = v
End Function

Private Function IsVisible(xp As String) As Boolean
    Dim js As String, v As Variant
    js = "(function(x){var r=document.evaluate(x,document,null,XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,null);" & _
         "for(var i=0;i<r.snapshotLength;i++){var e=r.snapshotItem(i);" & _
         "if(e.getClientRects().length&&getComputedStyle(e).visibility!=='hidden')return true;}return false;})(""" & xp & """)"
    v = m_Ctx.jsEval(js, IFEXCEPTION:=False, StopBiDiError:=False)
    If VarType(v) = vbBoolean Then IsVisible = v
End Function

'===================================================================================================
' 安定待機（CDP 版と同一）
'===================================================================================================
'  全信号HIT（+ visXPath 表示）かつ 通信0件 かつ DOM静穏>=minStableMs を満たす状態が、
'  minStableMs 連続したら STABLE。満たさないポーリングが1回でもあれば即リセット。
Private Function WaitSettled(minStableMs As Double, timeoutMs As Double, Optional visXPath As String) As Boolean
    Dim t0 As Double, stableStart As Double, quiet As Double
    Dim netOK As Boolean, visOK As Boolean, idleOK As Boolean
    Dim visLogged As Boolean

    Debug.Print Stamp & "[WAIT-BEGIN] minStableMs=" & minStableMs & " | maxTimeoutMs=" & timeoutMs & _
                IIf(LenB(visXPath), " | vis=" & visXPath, "")
    t0 = CDPHelpers.TimerCounter

    Do
        PumpNetwork
        quiet = DomQuietMs

        netOK = AllArmedHit
        If LenB(visXPath) = 0 Then
            visOK = True
        ElseIf Not visOK Then
            visOK = IsVisible(visXPath)   '一度表示されたら以後はラッチ
            If visOK And Not visLogged Then Debug.Print Stamp & "[HIT] vis": visLogged = True
        End If
        idleOK = (m_Inflight.Count = 0)

        If netOK And visOK And idleOK And quiet >= minStableMs Then
            If stableStart = 0 Then stableStart = CDPHelpers.TimerCounter
            If CDPHelpers.TimerCounter - stableStart >= minStableMs Then
                Debug.Print Stamp & "[WAIT-END] STABLE | elapsedMs=" & Format$(CDPHelpers.TimerCounter - t0, "0") & _
                            " | domQuietMs=" & Format$(quiet, "0")
                WaitSettled = True
                Exit Do
            End If
        Else
            stableStart = 0
        End If

        If CDPHelpers.TimerCounter - t0 > timeoutMs Then
            Debug.Print Stamp & "[WAIT-END] TIMEOUT | net=" & netOK & " vis=" & visOK & _
                        " inflight=" & m_Inflight.Count & " domQuietMs=" & Format$(quiet, "0") & PendingList
            Exit Do
        End If

        CDPHelpers.Sleep 0.1
    Loop

    DisarmAll
End Function

'タイムアウト時の調査用：未完了の通信と未HITの信号
Private Function PendingList() As String
    Dim k As Variant, s As String
    For Each k In m_Inflight.keys
        s = s & vbCrLf & "    inflight: " & Left$(m_Inflight(k), 160)
    Next
    For Each k In m_ArmHit.keys
        If Not m_ArmHit(k) Then s = s & vbCrLf & "    armed-miss: " & k
    Next
    PendingList = s
End Function

Private Function Stamp() As String
    Stamp = "[" & Format$((CDPHelpers.TimerCounter - m_T0) / 1000, "0.00") & "s] "
End Function
