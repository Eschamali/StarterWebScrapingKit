Attribute VB_Name = "Replay_GoogleFlights_exSpa"
'===================================================================================================
' Google Flights 手動操作ログの再現（exCDP_SpaElement / exCDP_SpaIdle 版）
'---------------------------------------------------------------------------------------------------
' 元にしたもの：ForDevelopers\OperationCheck\CDP\TestVBA\総合テスト\GoogleFlights\Replay_GoogleFlights.bas
'   操作の流れ・待つ信号・除外パターン・XPath は元のログ解析をそのまま使い、
'   自作していた要素操作と待機の部分を、拡張クラスに置き換えています。
'
'   元の自作部分                                  → 置き換え先
'   ・FindVisible / ClickVisible(sendClick)       → exCDP_SpaElement.click（UseTrustedClick = True）
'   ・候補を探して sendString                      → exCDP_SpaElement.InputText（入力後に値を照合）
'   ・NetInit / Arm / PumpNetwork                  → exCDP_SpaIdle.Init / ArmNetworkSignal
'   ・DomQuietMs / WaitSettled                     → exCDP_SpaIdle.WaitSettled
'   ・WaitSettled の visXPath                      → exCDP_SpaIdle.ArmVisibilitySignal（操作の「前」に仕掛ける）
'   ・IDLE_IGNORE / TRACK_TYPES                    → exCDP_SpaIdle.AddIgnorePattern / TrackResourceTypes(既定値と同じ)
'
' 操作の流れ（ログの [USER-ACTION] seq 番号）
'   0. https://www.google.com/travel/flights へ遷移                         [NAV-TARGET]
'   1. 券種コンボ(Round trip) → "One way"                                   seq3, seq5
'   2. "Where from?" → ダイアログに入力 → 候補 "Sapporo"                     seq7, (入力), seq9
'   3. "Where to?"   → ダイアログに入力 → 候補 "Paris, France"               seq14, (入力), seq16
'   4. Search → 日付ダイアログ → 出発日のセル → Done                          seq21, seq23, seq24
'   5. Search → 検索結果("Top flights")                                     seq26, seq27
'
' 必要なモジュール：Extensions\MainUnit\CDP\CounterplanForSPA の exCDP_SpaElement.cls / exCDP_SpaIdle.cls / exCDP_SpaHelpers.bas
'
' 注意事項
'   ・実サイトのため、Google 側の画面や通信の変更で動かなくなることがあります。
'     その場合は、各ステップの [WAIT] 行（LastWaitReport）で、何が揃わなかったかを確認してください
'   ・ForAI 側（BiDi版）では、物理クリックで検索結果が「Oops, something went wrong」になった報告があります。
'     同じ症状が出る場合は USE_TRUSTED_CLICK を False にして比較してください
'===================================================================================================
Option Explicit

'---------------------------------------------------------------------------------------------------
' 入力値（元のモジュールと同じ。ログには入力文字列そのものは残らないため、候補の aria-label 等から判断したもの）
'---------------------------------------------------------------------------------------------------
Private Const START_URL          As String = "https://www.google.com/travel/flights?hl=en-US"
Private Const ORIGIN_TEXT        As String = "Sapporo"
Private Const ORIGIN_OPTION      As String = "Sapporo"
Private Const DEST_TEXT          As String = "Paris"
Private Const DEST_OPTION        As String = "Paris, France"
Private Const DEPART_DAYS_AHEAD  As Long = 7            '出発日は今日から何日後か（過去日にならないよう相対指定）

'物理クリックを使うか（元のモジュールは全て sendClick）
Private Const USE_TRUSTED_CLICK  As Boolean = True

'---------------------------------------------------------------------------------------------------
' XPath（元のモジュールと同じ。動的な id は使わず、aria-label / role / data-iso で特定）
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

'アイドル判定から除外する通信（元のモジュールと同じ）
Private Const IDLE_IGNORE As String = "www.google.com/log?|play.google.com/log?|/generate_204|ogads-pa.clients6.google.com/|/FlightsFrontendUi/browserinfo"

'---------------------------------------------------------------------------------------------------
' 状態
'---------------------------------------------------------------------------------------------------
Private m_Ctx       As CDPContext
Private m_Spa       As exCDP_SpaElement
Private m_Idle      As exCDP_SpaIdle
Private m_Log       As String       '実行ログ（RunReplay の戻り値）
Private m_T0        As Double       'ログ表示用の基準時刻(ms)

'===================================================================================================
' メイン
'===================================================================================================
Public Sub ReplayGoogleFlightsSpa()
    Debug.Print RunReplay(keepBrowserOpen:=True)
End Sub

'---------------------------------------------------------------------------------------------------
' 再現を実行し、実行ログを返します（最終行が [RESULT] OK / NG）
'   keepBrowserOpen  True で、終了後もブラウザを開いたままにします（結果の目視確認用）
'---------------------------------------------------------------------------------------------------
Public Function RunReplay(Optional keepBrowserOpen As Boolean = True) As String
    m_Log = vbNullString

    '1. 空のタブで起動し、遷移の「前」に待機クラスを用意する（遷移中の通信も数えるため）
    Set m_Ctx = ShSetting01_StartBrowser.StartCDPModeContext("about:blank")
    m_T0 = CDPHelpers.TimerCounter      '※ TimerCounter はブラウザ起動時の初期化(周波数取得)より後でないと使えない

    Set m_Idle = New exCDP_SpaIdle
    m_Idle.Init m_Ctx
    Dim pat As Variant
    For Each pat In Split(IDLE_IGNORE, "|")
        m_Idle.AddIgnorePattern CStr(pat)
    Next

    Set m_Spa = New exCDP_SpaElement
    m_Spa.Init m_Ctx
    m_Spa.UseTrustedClick = USE_TRUSTED_CLICK
    m_Spa.SearchTimeoutSec = 10
    m_Spa.RaiseOnFailure = False    '失敗はステップごとにログへ残して、そこで打ち切る

    '2. 再現
    If ReplaySteps() Then Note "[RESULT] OK" Else Note "[RESULT] NG"

    '3. 後始末
    If Not keepBrowserOpen Then m_Ctx.ThisCDPBrowser.CloseBrowser
    RunReplay = m_Log
End Function

'===================================================================================================
' 操作手順
'===================================================================================================
Private Function ReplaySteps() As Boolean
    '0. 遷移
    m_Ctx.navigate START_URL
    Settle "0 navigate", 500, 10000

    '1. 券種 → One way
    If Not Act(m_Spa.click(XP_TICKET_TYPE), "1 open ticket type") Then Exit Function
    Settle "1 ticket type list", 500, 10000
    If Not Act(m_Spa.click(XP_ONE_WAY), "1 pick One way") Then Exit Function
    Settle "1 one way", 500, 10000

    '2. 出発地
    If Not SelectPlace(XP_FROM_MAIN, XP_FROM_DIALOG, "Enter your origin", ORIGIN_TEXT, ORIGIN_OPTION, "rpcids=tDoGIe", "2 origin") Then Exit Function

    '3. 目的地（ログでは目的地確定時に tDoGIe と BVAT3 が並んで飛ぶ。BVAT3 を信号にする）
    If Not SelectPlace(XP_TO_MAIN, XP_TO_DIALOG, "Enter your destination", DEST_TEXT, DEST_OPTION, "rpcids=BVAT3", "3 destination") Then Exit Function

    '4. Search（日付未指定なので日付ダイアログが開く）→ 日付 → Done
    Dim xpDate As String
    xpDate = "//div[@role='gridcell'][@data-iso='" & Format$(Date + DEPART_DAYS_AHEAD, "yyyy-mm-dd") & "']//div[@role='button']"

    m_Idle.ArmNetworkSignal "GetCalendarPicker"
    ArmVisible xpDate
    If Not Act(m_Spa.click(XP_SEARCH), "4 search (opens calendar)") Then Exit Function
    Settle "4 calendar", 500, 10000

    If Not Act(m_Spa.click(xpDate), "4 pick date " & Format$(Date + DEPART_DAYS_AHEAD, "yyyy-mm-dd")) Then Exit Function
    Settle "4 date", 500, 10000
    If Not Act(m_Spa.click(XP_DONE), "4 done") Then Exit Function
    Settle "4 done", 500, 10000

    '5. Search → 検索結果（ログでは応答まで約7.5秒。上限30秒）
    m_Idle.ArmNetworkSignal "GetShoppingResults"
    ArmVisible XP_TOP_FLIGHTS
    If Not Act(m_Spa.click(XP_SEARCH), "5 search") Then Exit Function
    Settle "5 results", 500, 30000

    '6. 結果の確認（待機が成功しても、データが正しいとは限らないため）
    Dim topVisible As Boolean
    topVisible = IsVisibleNow(XP_TOP_FLIGHTS)
    Note "[CHECK] 'Top flights' visible=" & topVisible & " | url=" & m_Ctx.Url & " | title=" & m_Ctx.Title
    ReplaySteps = topVisible
End Function

'---------------------------------------------------------------------------------------------------
' 出発地/目的地の共通手順：入力欄クリック → ダイアログで入力 → 候補を選択
'---------------------------------------------------------------------------------------------------
Private Function SelectPlace(xpMain As String, xpDialogInput As String, dialogLabel As String, _
                             placeText As String, optionLabel As String, confirmPattern As String, stepName As String) As Boolean
    Dim xpOption As String
    xpOption = "//div[@role='dialog'][@aria-label='" & dialogLabel & "']//li[@role='option'][@aria-label='" & optionLabel & "']"

    'ダイアログを開く（入力欄の表示を目印にする）
    ArmVisible xpDialogInput
    If Not Act(m_Spa.click(xpMain), stepName & " open dialog") Then Exit Function
    Settle stepName & " dialog", 500, 10000

    '入力 → 候補の表示を待つ（候補を返す H028ib の完了と、目的の候補 li の表示を目印にする）
    m_Idle.ArmNetworkSignal "rpcids=H028ib"
    ArmVisible xpOption
    If Not Act(m_Spa.InputText(xpDialogInput, placeText), stepName & " type '" & placeText & "'") Then Exit Function
    Settle stepName & " suggestions", 500, 10000

    '候補を選択 → 確定時の通信を待つ（初回ログで minStableMs=1000 だった箇所）
    m_Idle.ArmNetworkSignal confirmPattern
    If Not Act(m_Spa.click(xpOption), stepName & " pick '" & optionLabel & "'") Then Exit Function
    Settle stepName & " confirm", 1000, 10000

    SelectPlace = True
End Function

'===================================================================================================
' 小物
'===================================================================================================
'操作の結果をログに残し、そのまま返す
Private Function Act(succeeded As Boolean, stepName As String) As Boolean
    Note "[ACT] " & stepName & " -> " & IIf(succeeded, "OK", "NG")
    Act = succeeded
End Function

'待機して、結果をログに残す（タイムアウトしても続行する。元のモジュールと同じ扱い）
Private Sub Settle(stepName As String, minStableMs As Long, maxWaitMs As Long)
    m_Idle.WaitSettled minStableMs, maxWaitMs
    Note "[WAIT] " & stepName & " | " & m_Idle.LastWaitReport
End Sub

'表示待ちの信号を仕掛ける。既に表示されている場合は、待つ意味がないので仕掛けない
'（exCDP_SpaIdle.ArmVisibilitySignal は、表示済みの要素に仕掛けるとエラー停止するため）
Private Sub ArmVisible(xp As String)
    If IsVisibleNow(xp) Then
        Note "[ARM-SKIP] already visible: " & xp
    Else
        m_Idle.ArmVisibilitySignal xp
    End If
End Sub

'XPath に一致する要素のどれかが、今表示されているか
Private Function IsVisibleNow(xp As String) As Boolean
    Dim js As String, v As Variant
    js = "function(x){" & exCDP_SpaHelpers.VisiblePredicateJs() & exCDP_SpaHelpers.FindElementsJs() & _
         "try{return find(x,false).some(vis);}catch(_){return false;}}"
    v = m_Ctx.jsEval(js, , exCDP_SpaHelpers.JsArgs(xp), False)
    If VarType(v) = vbBoolean Then IsVisibleNow = v
End Function

Private Sub Note(msg As String)
    Dim line As String
    line = "[" & Format$((CDPHelpers.TimerCounter - m_T0) / 1000, "0.00") & "s] " & msg
    Debug.Print line
    m_Log = m_Log & line & vbCrLf
End Sub
