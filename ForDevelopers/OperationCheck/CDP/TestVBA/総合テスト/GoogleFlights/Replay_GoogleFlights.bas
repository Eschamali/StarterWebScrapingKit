Attribute VB_Name = "Replay_GoogleFlights"
'===================================================================================================
' Google Flights 手動操作ログ(BiDi Discovery Log, recorderCtx=166931A0...)の Kit版(CDP)再現
'---------------------------------------------------------------------------------------------------
' 操作の流れ（ログの [USER-ACTION] seq 番号）
'   0. https://www.google.com/travel/flights へ遷移                         [NAV-TARGET]
'   1. 券種コンボ(Round trip) → "One way"                                   seq3, seq5
'   2. "Where from?" → ダイアログに入力 → 候補 "Sapporo"                     seq7, (入力), seq9
'   3. "Where to?"   → ダイアログに入力 → 候補 "Paris, France"               seq14, (入力), seq16
'   4. Search → 日付ダイアログ → 2026-11-22 のセル → Done                    seq21, seq23, seq24
'   5. Search → 検索結果("Top flights")                                     seq26, seq27
'
' 待機の考え方（ログの SPA 待機を CDP で再構成）
'   ・ネットワーク信号 : 操作「前」に Arm → Network.responseReceived の URL 部分一致
'                        → 同じ requestId の Network.loadingFinished で HIT
'   ・ネットワーク静穏 : XHR/Fetch/Document の requestWillBeSent ? loadingFinished/Failed を
'                        requestId で数え、0件か（アイドル除外パターンは数えない）
'   ・DOM 静穏         : ページに MutationObserver を仕込み、最終変化からの経過 ms を読む
'   ・安定判定         : 全信号HIT + 通信0件 + DOM静穏>=minStableMs の状態が、
'                        さらに minStableMs 連続したら STABLE
'
' 前提・制約
'   ・Pipe / WebSocket 接続を想定。WebView2 接続では Network.* を
'     SubscribeCdpEvent で個別購読しない限りイベントが届きません。
'   ・ctx.BrowserEvents（キャプチャ用Queue）を本モジュールが占有し、取り出すたびに空にします。
'===================================================================================================
Option Explicit

'---------------------------------------------------------------------------------------------------
' 入力値
'   ログは入力文字列そのものを記録していない（valueLen のみ）。
'   ・出発地: change 時 valueLen=7、選択候補 aria-label="Sapporo"、H028ib 応答の推移(New York →
'     San Francisco → San Pedro Sula → Sapporo)から "Sapporo" と判断
'   ・目的地: valueLen=5、選択候補 aria-label="Paris, France" から "Paris" と判断
'   違っていればここを直してください。
'---------------------------------------------------------------------------------------------------
Private Const START_URL      As String = "https://www.google.com/travel/flights"
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
' ネットワーク関連
'---------------------------------------------------------------------------------------------------
'アイドル判定から除外（ログで idle=skip のもの + SPA-CONSENSUS-SUSPECT が付いた browserinfo）
Private Const IDLE_IGNORE As String = "www.google.com/log?|play.google.com/log?|/generate_204|ogads-pa.clients6.google.com/|/FlightsFrontendUi/browserinfo"
'アイドル判定で数える resourceType（ログの excludeResources=True に相当）
Private Const TRACK_TYPES As String = "|XHR|Fetch|Document|"

'---------------------------------------------------------------------------------------------------
' 状態
'---------------------------------------------------------------------------------------------------
Private m_Ctx       As CDPContext
Private m_Inflight  As Dictionary   'requestId -> url（アイドル判定対象で未完了のもの）
Private m_ArmHit    As Dictionary   'pattern   -> Boolean（HIT済みか）
Private m_ArmReq    As Dictionary   'requestId -> pattern（URL一致済み・完了待ち）
Private m_T0        As Double       'ログ表示用の基準時刻(ms)

'===================================================================================================
' メイン
'===================================================================================================
Public Sub ReplayGoogleFlights()
    Set m_Ctx = ShSetting01_StartBrowser.StartCDPModeContext
    m_T0 = CDPHelpers.TimerCounter

    '0. 遷移前にネットワーク監視を開始（遷移中の通信も数えるため）
    NetInit
    m_Ctx.navigate START_URL
    WaitSettled 500, 10000

    '1. 券種 → One way
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

    '5. Search → 検索結果（ログでは応答まで約7.5秒。上限30秒）
    Arm "GetShoppingResults"
    ClickVisible XP_SEARCH
    WaitSettled 500, 30000, XP_TOP_FLIGHTS

    Debug.Print Stamp & "[DONE] url=" & m_Ctx.Url & " | title=" & m_Ctx.Title
    'm_Ctx.ThisCDPBrowser.CloseBrowser
End Sub

'---------------------------------------------------------------------------------------------------
' 出発地/目的地の共通手順：入力欄クリック → ダイアログで入力 → 候補を選択
'---------------------------------------------------------------------------------------------------
Private Sub SelectPlace(xpMain As String, xpDialogInput As String, dialogLabel As String, _
                        inputText As String, optionLabel As String, confirmPattern As String)
    Dim xpOption As String
    xpOption = "//div[@role='dialog'][@aria-label='" & dialogLabel & "']//li[@role='option'][@aria-label='" & optionLabel & "']"

    'ダイアログを開く（開いた直後に空文字の H028ib が1回飛ぶので、ここで静まるのを待つ）
    ClickVisible xpMain
    WaitSettled 500, 10000, xpDialogInput

    '入力 → 候補の表示を待つ（値で特定できる候補 li を表示ゲートにする）
    Arm "rpcids=H028ib"
    FindVisible(xpDialogInput).sendString inputText
    WaitSettled 500, 10000, xpOption

    '候補を選択 → 確定時の通信を待つ（初回ログで minStableMs=1000 だった箇所）
    Arm confirmPattern
    ClickVisible xpOption
    WaitSettled 1000, 10000
End Sub

'===================================================================================================
' 要素操作
'===================================================================================================
'XPath に一致する要素のうち表示中の1件を返す（見つかるまで待機）
'  1巡目: 表示中かつ aria-hidden 祖先なし / 2巡目: 表示中のみ（ログの LOCATE-PREFER-VISIBLE 相当）
Private Function FindVisible(xp As String, Optional timeoutMs As Double = 10000) As CDPElement
    Dim js As String
    js = "(function(x){var r=document.evaluate(x,document,null,XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,null);" & _
         "function vis(e){if(!e.getClientRects().length)return false;var s=getComputedStyle(e);return s.visibility!=='hidden';}" & _
         "var i,e;for(i=0;i<r.snapshotLength;i++){e=r.snapshotItem(i);if(vis(e)&&!e.closest('[aria-hidden=""true""]'))return i+1;}" & _
         "for(i=0;i<r.snapshotLength;i++){if(vis(r.snapshotItem(i)))return i+1;}return 0;})(""" & xp & """)"

    Dim t0 As Double: t0 = CDPHelpers.TimerCounter
    Dim idx As Variant
    Do
        idx = m_Ctx.jsEval(js, IFEXCEPTION:=0)
        If VarType(idx) = vbDouble Then If idx > 0 Then Exit Do
        If CDPHelpers.TimerCounter - t0 > timeoutMs Then
            Err.Raise vbObjectError + 513, "FindVisible", "表示中の要素が見つかりません: " & xp
        End If
        PumpNetwork
        CDPHelpers.Sleep 0.1
    Loop

    Set FindVisible = m_Ctx.getElementByXPath("(" & xp & ")[" & CLng(idx) & "]")
End Function

'ログの操作は全て実クリック(pointerdown→click)なので、Input.dispatchMouseEvent 系の sendClick を使う
Private Sub ClickVisible(xp As String)
    Debug.Print Stamp & "[CLICK] " & xp
    FindVisible(xp).sendClick
End Sub

'===================================================================================================
' ネットワーク監視
'===================================================================================================
Private Sub NetInit()
    Set m_Inflight = New Dictionary
    Set m_ArmHit = New Dictionary
    Set m_ArmReq = New Dictionary

    'キャプチャ対象を Network の4種類に限定（それ以外は Queue に入れず、JSON パースもしない）
    m_Ctx.SetFilterEvents = "Network.requestWillBeSent"
    m_Ctx.SetFilterEvents = "Network.responseReceived"
    m_Ctx.SetFilterEvents = "Network.loadingFinished"
    m_Ctx.SetFilterEvents = "Network.loadingFailed"
    Set m_Ctx.BrowserEvents = New Dictionary   'キャプチャ開始

    m_Ctx.ExecuteCDP "Network.enable"
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

'届いているイベントを取り出して状態に反映し、Queue を空にする
'  ※VBA はシングルスレッドで、Queue への追加は TakeEvents の中でしか起きないため、
'    処理後に作り直しても取りこぼしは生じない
Private Sub PumpNetwork()
    m_Ctx.ThisCDPBrowser.TakeEvents

    Dim methods As Dictionary
    Set methods = m_Ctx.BrowserEvents("EventMethods")
    If methods.Count = 0 Then Exit Sub

    'Queue はメソッド名ごとに分かれていて相互の順序はないため、開始→応答→完了の順で処理する
    ProcessEvents methods, "Network.requestWillBeSent"
    ProcessEvents methods, "Network.responseReceived"
    ProcessEvents methods, "Network.loadingFinished"
    ProcessEvents methods, "Network.loadingFailed"

    Set m_Ctx.BrowserEvents = New Dictionary
End Sub

Private Sub ProcessEvents(methods As Dictionary, methodName As String)
    If Not methods.Exists(methodName) Then Exit Sub

    Dim ev As Variant, p As Dictionary, rid As String, url As String
    For Each ev In methods(methodName)
        Set p = ev("params")
        rid = p("requestId")

        Select Case methodName
            Case "Network.requestWillBeSent"
                url = p("request")("url")
                Dim typ As String: typ = vbNullString
                If p.Exists("type") Then typ = p("type")
                'リダイレクトは同じ requestId で再度来るので、未登録の時だけ数える
                If InStr(TRACK_TYPES, "|" & typ & "|") > 0 And Not IsIgnored(url) Then
                    If Not m_Inflight.Exists(rid) Then m_Inflight.Add rid, url
                End If

            Case "Network.responseReceived"
                url = p("response")("url")
                Dim k As Variant
                For Each k In m_ArmHit.keys
                    If Not m_ArmHit(k) And Not m_ArmReq.Exists(rid) Then
                        If InStr(url, k) > 0 Then m_ArmReq.Add rid, k
                    End If
                Next

            Case "Network.loadingFinished"
                If m_ArmReq.Exists(rid) Then
                    m_ArmHit(m_ArmReq(rid)) = True
                    Debug.Print Stamp & "[HIT] net pattern=" & m_ArmReq(rid) & " | requestId=" & rid
                    m_ArmReq.Remove rid
                End If
                If m_Inflight.Exists(rid) Then m_Inflight.Remove rid

            Case "Network.loadingFailed"
                If m_ArmReq.Exists(rid) Then
                    Debug.Print Stamp & "[FAIL] net pattern=" & m_ArmReq(rid) & " | " & p("errorText")
                    m_ArmReq.Remove rid   '失敗は HIT 扱いにしない（後続の同パターン応答を待つ）
                End If
                If m_Inflight.Exists(rid) Then m_Inflight.Remove rid
        End Select
    Next
End Sub

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

'===================================================================================================
' DOM 静穏プローブ
'===================================================================================================
'最終 DOM 変化からの経過 ms。プローブが無ければ（新しい document 等）仕込み直して 0 を返す
Private Function DomQuietMs() As Double
    Const JS_READ As String = "(function(){var p=window.__kitProbe;return p?performance.now()-p.last:-1;})()"
    Const JS_INSTALL As String = "(function(){if(window.__kitProbe)return 0;window.__kitProbe={last:performance.now()};" & _
        "new MutationObserver(function(){window.__kitProbe.last=performance.now();})" & _
        ".observe(document.documentElement,{subtree:true,childList:true,attributes:true,characterData:true});return 0;})()"

    Dim v As Variant
    v = m_Ctx.jsEval(JS_READ, IFEXCEPTION:=-1)
    If VarType(v) <> vbDouble Then v = -1
    If v < 0 Then
        m_Ctx.jsEval JS_INSTALL, IFEXCEPTION:=0
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
    v = m_Ctx.jsEval(js, IFEXCEPTION:=False)
    If VarType(v) = vbBoolean Then IsVisible = v
End Function

'===================================================================================================
' 安定待機
'===================================================================================================
'  全信号HIT（+ visXPath 表示）かつ 通信0件 かつ DOM静穏>=minStableMs を満たした状態が、
'  minStableMs 連続したら STABLE。満たさないポーリングが1回でもあれば窓をリセット。
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
