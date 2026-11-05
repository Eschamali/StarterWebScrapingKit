VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_SpaIdle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'                     SPAサイト向けの「操作後の待機」アシストを行います（第2段階）
'
'   SPA には「読み込み完了」の明確な合図がないため、次の状態が揃ったことをもって「落ち着いた」と判断します。
'
'       [通信]  追跡対象(XHR/Fetch/Document)の通信が0件で、最後の通信から minStableMs 以上経過
'       [DOM]   最後のDOM変化から minStableMs 以上経過（open な Shadow DOM 内の変化も含む）
'       [信号]  操作の「前」に仕掛けた信号が全て成立（1回の待機で使い切り）
'                 ・ArmNetworkSignal     … URLに一致する通信が完了した
'                 ・ArmContentSignal     … 既存要素の配下が書き換わった
'                 ・ArmVisibilitySignal  … 要素が表示された
'
'   参考：ForAI\WebDriver-BiDi-for-SeleniumVBA (BiDiCommandWrapper.cls) の SPA 待機、
'         ForDevelopers\OperationCheck\CDP\TestVBA\総合テスト\GoogleFlights\Replay_GoogleFlights.bas
'
'* 必要なモジュール：exCDP_SpaHelpers.bas（exCDP_SpaElement.cls と共通の部品）
'
'* 前提・制約：
'   ・Pipe / WebSocket 接続を想定しています。WebView2 接続では`Network.*`イベントが届かない場合があります
'   ・`Init`で`Network.enable`を行います。ページ遷移の通信も数えるには、遷移より前に`Init`してください
'   ・DOM の監視対象はメインフレームのみです（iframe内の変化は数えません）
'   ・「落ち着いた」は推定であり、アプリ内部の完了を保証するものではありません。
'     重要な操作では、信号を仕掛けた上で、件数や文言の確認も併用してください
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'ブラウザ制御オブジェクト
Private WithEvents ex_CDPContext    As CDPContext   '継承元のタブクラス
Attribute ex_CDPContext.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "exCDP_SpaIdle"
Private Const POLLING_INTERVAL_SEC  As Double = 0.05    '待機ループの間隔(秒)
Private Const REPORT_URL_MAX        As Long = 5         'タイムアウト報告に載せる通信URLの上限数

'基本ステート管理
Private m_LogID                     As String       'ログ識別ID

'動作設定
Private m_DefaultMinStableMs        As Long         '`WaitSettled`省略時の静穏時間(ms)
Private m_DefaultTimeoutMs          As Long         '`WaitSettled`省略時のタイムアウト(ms)
Private m_RaiseOnTimeout            As Boolean      'タイムアウト時に`Err.Raise`で止めるか
Private m_LongRequestThresholdMs    As Long         'これより長く続く通信は、ロングポーリング等とみなして数えない(0で無効)
Private m_TrackTypes                As String       '追跡する resourceType（"|XHR|Fetch|Document|"形式）
Private m_IgnorePatterns            As Collection   '通信の除外URLパターン
Private m_IgnoreSelectors           As Collection   'DOM変化の除外CSSセレクター

'通信の追跡状態
Private m_InflightUrl               As Dictionary   'requestId -> url（追跡対象で未完了のもの）
Private m_InflightStart             As Dictionary   'requestId -> 開始時刻(ms)
Private m_LastNetBusyMs             As Double       '最後に追跡対象の通信が動いた時刻(ms)

'信号の状態
Private m_NetArms                   As Dictionary   'pattern   -> 成立済みか
Private m_NetArmReq                 As Dictionary   'requestId -> pattern（URL一致済み・完了待ち）
Private m_ContentArmed              As Boolean
Private m_ContentHit                As Boolean
Private m_ContentLost               As Boolean      '遷移等で監視が失われた
Private m_ContentXPath              As String
Private m_VisArmed                  As Boolean
Private m_VisSeen                   As Boolean
Private m_VisXPath                  As String
Private m_VisShadow                 As Boolean
Private m_VisNewOnly               As Boolean      '仕掛けた時点の要素を除外するか

'結果
Private m_LastWaitReport            As String



'***************************************************************************************************
'                               ■■■ 設定プロパティ ■■■
'***************************************************************************************************
'* 機能　　：`WaitSettled`で`minStableMs`を省略した際の静穏時間(ms)です（既定：500）
'***************************************************************************************************
Public Property Get DefaultMinStableMs() As Long
    DefaultMinStableMs = m_DefaultMinStableMs
End Property
Public Property Let DefaultMinStableMs(v As Long)
    If v < 0 Then v = 0
    m_DefaultMinStableMs = v
End Property

'***************************************************************************************************
'* 機能　　：`WaitSettled`で`maxWaitMs`を省略した際のタイムアウト(ms)です（既定：10000）
'***************************************************************************************************
Public Property Get DefaultTimeoutMs() As Long
    DefaultTimeoutMs = m_DefaultTimeoutMs
End Property
Public Property Let DefaultTimeoutMs(v As Long)
    If v < 0 Then v = 0
    m_DefaultTimeoutMs = v
End Property

'***************************************************************************************************
'* 機能　　：タイムアウト時に`Err.Raise`で止めるかです（既定：False = WARNログを出して`False`を返す）
'***************************************************************************************************
Public Property Get RaiseOnTimeout() As Boolean
    RaiseOnTimeout = m_RaiseOnTimeout
End Property
Public Property Let RaiseOnTimeout(v As Boolean)
    m_RaiseOnTimeout = v
End Property

'***************************************************************************************************
'* 機能　　：これより長く続いている通信を、ロングポーリング等の常駐通信とみなして数えない閾値(ms)です（既定：0 = 無効）
'---------------------------------------------------------------------------------------------------
'* 注意事項：本当に時間のかかる通信（重い検索APIなど）まで無視してしまうため、
'            まずは`AddIgnorePattern`で個別に除外することをお勧めします
'***************************************************************************************************
Public Property Get LongRequestThresholdMs() As Long
    LongRequestThresholdMs = m_LongRequestThresholdMs
End Property
Public Property Let LongRequestThresholdMs(v As Long)
    If v < 0 Then v = 0
    m_LongRequestThresholdMs = v
End Property

'***************************************************************************************************
'* 機能　　：追跡する通信の種類（CDPの`Network.ResourceType`）をカンマ区切りで指定します（既定："XHR,Fetch,Document"）
'***************************************************************************************************
Public Property Get TrackResourceTypes() As String
    TrackResourceTypes = Replace(Mid$(m_TrackTypes, 2, Len(m_TrackTypes) - 2), "|", ",")
End Property
Public Property Let TrackResourceTypes(v As String)
    m_TrackTypes = "|" & Replace(Replace(v, " ", vbNullString), ",", "|") & "|"
End Property



'***************************************************************************************************
'                             ■■■ 情報取得プロパティ ■■■
'***************************************************************************************************
'* 機能　　：直前の`WaitSettled`の結果報告です（タイムアウト時は、何が揃わなかったかを含みます）
'***************************************************************************************************
Public Property Get LastWaitReport() As String
    LastWaitReport = m_LastWaitReport
End Property

'***************************************************************************************************
'* 機能　　：現在数えている（除外されていない）通信中の件数です
'***************************************************************************************************
Public Property Get InflightCount() As Long
    InflightCount = CountInflight(False)
End Property



'***************************************************************************************************
'                               ■■■ 除外ルール ■■■
'***************************************************************************************************
'* 機能　　：待機判定から除外する通信のURLパターンを追加します
'---------------------------------------------------------------------------------------------------
'* 引数　　：urlPattern   URLの部分一致。`*`で任意の文字列に一致します（例：`/log?`、`*/collect*`）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：解析・広告・ロングポーリング等、常に通信していて判定を妨げるものに使います
'            既に通信中のものにも即座に適用されます
'* 注意事項：広すぎるパターンは、本当に待つべき通信まで除外し、早すぎる完了判定の原因になります
'***************************************************************************************************
Public Sub AddIgnorePattern(urlPattern As String)
    If LenB(Trim$(urlPattern)) = 0 Then Exit Sub
    m_IgnorePatterns.Add urlPattern
End Sub

'***************************************************************************************************
'* 機能　　：待機判定から除外するDOM変化の範囲を、CSSセレクターで追加します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：時計・カウントダウン・カルーセル・広告枠など、常に変化し続ける要素に使います
'            一致する要素自身とその配下の変化を数えません（Shadow DOM 内からはホスト要素も辿ります）
'***************************************************************************************************
Public Sub AddIgnoreSelector(cssSelector As String)
    If LenB(Trim$(cssSelector)) = 0 Then Exit Sub
    m_IgnoreSelectors.Add cssSelector
End Sub

'***************************************************************************************************
'* 機能　　：除外ルールを全て削除します
'***************************************************************************************************
Public Sub ClearIgnoreRules()
    Set m_IgnorePatterns = New Collection
    Set m_IgnoreSelectors = New Collection
End Sub



'***************************************************************************************************
'                     ■■■ 信号（操作の「前」に仕掛ける・1回の待機で使い切り） ■■■
'***************************************************************************************************
'* 機能　　：URLに一致する通信の完了を、次の`WaitSettled`の成立条件に加えます
'---------------------------------------------------------------------------------------------------
'* 引数　　：urlPattern   レスポンスURLの部分一致。`*`使用可（例：`rpcids=BVAT3`、`/api/search*`）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・URLが一致したレスポンスについて、`Network.loadingFinished`が届いた時点で成立します
'            ・ページ遷移をまたいでも有効です（遷移の前に仕掛けることもできます）
'            ・失敗(`loadingFailed`)した通信では成立しません。同じパターンの次の通信を待ちます
'* 注意事項：仕掛ける前に完了した通信は数えません。必ず操作の「前」に呼んでください
'***************************************************************************************************
Public Sub ArmNetworkSignal(urlPattern As String)
    Const FromProcedureName As String = ThisClassName & ".ArmNetworkSignal"


    If LenB(Trim$(urlPattern)) = 0 Then printMsg ERROR_, "urlPattern is empty.", FromProcedureName, doRaiseError:=True
    If m_NetArms.Exists(urlPattern) Then m_NetArms.Remove urlPattern
    m_NetArms.Add urlPattern, False
    printMsg info_, "[ARM] network pattern=" & urlPattern, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：既存要素の配下が書き換わることを、次の`WaitSettled`の成立条件に加えます
'---------------------------------------------------------------------------------------------------
'* 引数　　：xpath              監視したい既存要素（結果一覧の入れ物など）のXPath
'            searchShadow       `True`で open な Shadow-root 内も探索します（XPathは`/`始まり）
'            includeAttributes  `True`で属性の変化も数えます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：「通信は終わったのに、描画がまだ」という隙間を埋めるための信号です
'* 注意事項：・仕掛ける時点で要素が存在している必要があります（無ければエラー停止）。
'              操作後に新しく作られる要素には、その親の要素か`ArmVisibilitySignal`を使ってください
'            ・ページ遷移でページ側の監視が失われるため、遷移をまたぐ用途には使えません
'            ・時計や広告など、無関係に変化し続ける範囲を含む要素は避けてください
'            ・「いったん空にしてから描き直す」UIでは、空にした時点で成立してしまいます。
'              その場合は`ArmVisibilitySignal`の`newOnly:=True`で、新しく描かれた要素を待ってください
'***************************************************************************************************
Public Sub ArmContentSignal(xpath As String, Optional searchShadow As Boolean = False, Optional includeAttributes As Boolean = False)
    Const FromProcedureName As String = ThisClassName & ".ArmContentSignal"


    If searchShadow And Not exCDP_SpaHelpers.IsShadowXPath(xpath) Then
        printMsg ERROR_, "With searchShadow, the XPath must start with '/': " & xpath, FromProcedureName, doRaiseError:=True
    End If

    '1. ページ側のプローブを用意
    Dim fresh As Boolean, quietMs As Double, sigCount As Long, visNow As Long
    If Not PollPage(fresh, quietMs, sigCount, visNow) Then printMsg ERROR_, "The page is not ready for arming a content signal.", FromProcedureName, doRaiseError:=True

    '2. 監視開始
    Dim r As Variant
    r = EvalWithArgs(ArmContentJs(), exCDP_SpaHelpers.JsArgs(xpath, searchShadow, includeAttributes))
    Select Case True
        Case IsNumeric(r) And Not IsEmpty(r)
            If r = 0 Then
                m_ContentArmed = True
                m_ContentHit = False
                m_ContentLost = False
                m_ContentXPath = xpath
                printMsg info_, "[ARM] content xpath=" & xpath, FromProcedureName
                Exit Sub
            ElseIf r = -1 Then
                printMsg ERROR_, "The content-signal target does not exist (it must exist when armed): " & xpath, FromProcedureName, doRaiseError:=True
            End If
    End Select
    printMsg ERROR_, "The content signal could not be armed (invalid XPath?): " & xpath, FromProcedureName, doRaiseError:=True
End Sub

'***************************************************************************************************
'* 機能　　：要素が表示されることを、次の`WaitSettled`の成立条件に加えます
'---------------------------------------------------------------------------------------------------
'* 引数　　：xpath          表示を待ちたい要素のXPath（一致する要素のどれか1つが表示されれば成立）
'            searchShadow   `True`で open な Shadow-root 内も探索します（XPathは`/`始まり）
'            newOnly        `True`で、仕掛けた時点で既に存在していた要素を除外し、新しく作られた要素の表示だけを待ちます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・一度表示されたら、以降は成立したままとします
'            ・ページ遷移をまたいでも有効です（遷移後の要素は全て「新しい要素」として扱います）
'            ・`newOnly:=True`は、「いったん空にしてから描き直す」一覧の再描画待ちに向いています
'* 注意事項：・`newOnly:=False`の場合、仕掛ける時点で既に表示されていると、待つ意味がないためエラー停止します
'            ・`newOnly:=True`は要素の作り直しを前提とします。同じ要素の表示/非表示を切り替えるだけのUIでは成立しません
'***************************************************************************************************
Public Sub ArmVisibilitySignal(xpath As String, Optional searchShadow As Boolean = False, Optional newOnly As Boolean = False)
    Const FromProcedureName As String = ThisClassName & ".ArmVisibilitySignal"


    If searchShadow And Not exCDP_SpaHelpers.IsShadowXPath(xpath) Then
        printMsg ERROR_, "With searchShadow, the XPath must start with '/': " & xpath, FromProcedureName, doRaiseError:=True
    End If

    '1. 仕掛けた上で、現時点の表示状態を確認
    m_VisArmed = True
    m_VisSeen = False
    m_VisXPath = xpath
    m_VisShadow = searchShadow
    m_VisNewOnly = False

    Dim fresh As Boolean, quietMs As Double, sigCount As Long, visNow As Long
    If Not PollPage(fresh, quietMs, sigCount, visNow) Then
        If newOnly Then
            ClearVisibility
            printMsg ERROR_, "The page is not ready for arming a visibility signal with newOnly.", FromProcedureName, doRaiseError:=True
        End If
    ElseIf visNow = -2 Then
        ClearVisibility
        printMsg ERROR_, "The XPath could not be evaluated (syntax error?): " & xpath, FromProcedureName, doRaiseError:=True
    ElseIf newOnly Then
        '2. 今ある要素を「古い要素」として記録
        Dim r As Variant
        r = EvalWithArgs(SnapshotOldJs(), exCDP_SpaHelpers.JsArgs(xpath, searchShadow, Empty))
        If IsEmpty(r) Or Not IsNumeric(r) Then r = -1
        If r <> 0 Then
            ClearVisibility
            printMsg ERROR_, "The existing elements could not be recorded: " & xpath, FromProcedureName, doRaiseError:=True
        End If
        m_VisNewOnly = True
    ElseIf visNow = 1 Then
        ClearVisibility
        printMsg ERROR_, "The element is already visible, so a visibility signal cannot tell when the action completes. Use newOnly:=True if the element is re-created, or choose another signal: " & xpath, FromProcedureName, doRaiseError:=True
    End If

    printMsg info_, "[ARM] visibility xpath=" & xpath, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：仕掛けた信号を全て解除します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`WaitSettled`の終了時にも自動で呼ばれます
'***************************************************************************************************
Public Sub DisarmAll()
    m_NetArms.RemoveAll
    m_NetArmReq.RemoveAll

    If m_ContentArmed Or m_VisNewOnly Then
        On Error Resume Next    'ページが既に無い場合は、何もしなくてよい
        ex_CDPContext.jsEval DisarmPageJs(), IFEXCEPTION:=0
        On Error GoTo 0
    End If
    m_ContentArmed = False
    m_ContentHit = False
    m_ContentLost = False
    m_ContentXPath = vbNullString

    ClearVisibility
End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************
'* 機能　　：ページが落ち着くまで待機します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：落ち着いた場合は`True`、タイムアウトした場合は`False`（`RaiseOnTimeout = True`ならエラー停止）
'* 引数　　：minStableMs   通信・DOMそれぞれに求める静穏時間(ms)。省略時は`DefaultMinStableMs`
'            maxWaitMs     タイムアウト(ms)。省略時は`DefaultTimeoutMs`
'---------------------------------------------------------------------------------------------------
'* 機能説明：次の全てが揃った時点で完了とします
'              ・追跡対象の通信が0件（除外ルール・`LongRequestThresholdMs`適用後）
'              ・最後の通信から minStableMs 以上経過
'              ・最後のDOM変化から minStableMs 以上経過
'              ・仕掛けた信号が全て成立
'            静穏時間は、この待機を開始してから観測した分だけを数えます。そのため最短でも minStableMs かかります
'            （操作の直後、効果が現れる前に「元から静かだった」として完了するのを防ぐためです）
'            終了時（成功・タイムアウトとも）に、仕掛けた信号は解除されます
'* 注意事項：操作から効果（通信やDOM変化）が現れるまで minStableMs 以上かかるUIでは、信号を併用してください
'            タイムアウト時は`LastWaitReport`に、揃わなかった条件と通信中のURLが入ります。
'            除外ルールの見直しにお使いください
'***************************************************************************************************
Public Function WaitSettled(Optional ByVal minStableMs As Long = -1, Optional ByVal maxWaitMs As Long = -1) As Boolean
    Const FromProcedureName As String = ThisClassName & ".WaitSettled"


    If minStableMs < 0 Then minStableMs = m_DefaultMinStableMs
    If maxWaitMs < 0 Then maxWaitMs = m_DefaultTimeoutMs

    Dim t0 As Double: t0 = CDPHelpers.TimerCounter
    Dim pageOK As Boolean, fresh As Boolean, domQuietMs As Double, sigCount As Long, visNow As Long
    Dim inflight As Long, netQuietMs As Double, elapsedMs As Double

    Do
        '1. 届いているイベントを処理（`ex_CDPContext_ExEvent`で通信状態が更新されます）
        TakeEvents

        '2. ページ側の状態を読む
        pageOK = PollPage(fresh, domQuietMs, sigCount, visNow)
        If pageOK Then
            If fresh And m_ContentArmed And Not m_ContentHit And Not m_ContentLost Then
                m_ContentLost = True
                printMsg WARN_, "The page was replaced while a content signal was armed; the signal can no longer be observed: " & m_ContentXPath, FromProcedureName
            End If
            If m_ContentArmed And sigCount > 0 Then m_ContentHit = True
            If m_VisArmed And visNow = 1 Then m_VisSeen = True
        End If

        '3. 通信の状態
        inflight = CountInflight(True)
        netQuietMs = CDPHelpers.TimerCounter - m_LastNetBusyMs

        '   静穏は、この待機の開始以降に観測した分だけ数える
        '   （操作の直後、効果が現れる前に「元から静か」で完了してしまうのを防ぐため）
        elapsedMs = CDPHelpers.TimerCounter - t0
        If netQuietMs > elapsedMs Then netQuietMs = elapsedMs
        If domQuietMs > elapsedMs Then domQuietMs = elapsedMs

        '4. 判定
        If pageOK And inflight = 0 And netQuietMs >= minStableMs And domQuietMs >= minStableMs And AllSignalsMet() Then
            m_LastWaitReport = "STABLE | elapsedMs=" & Format$(CDPHelpers.TimerCounter - t0, "0") & _
                               " | netQuietMs=" & Format$(netQuietMs, "0") & " | domQuietMs=" & Format$(domQuietMs, "0")
            printMsg info_, "[WAIT-END] " & m_LastWaitReport, FromProcedureName
            WaitSettled = True
            Exit Do
        End If

        '5. タイムアウト
        If CDPHelpers.TimerCounter - t0 > maxWaitMs Then
            m_LastWaitReport = "TIMEOUT(" & maxWaitMs & "ms) | page=" & IIf(pageOK, "ok", "unavailable") & _
                               " | inflight=" & inflight & " | netQuietMs=" & Format$(netQuietMs, "0") & _
                               " | domQuietMs=" & Format$(domQuietMs, "0") & PendingSummary()
            printMsg WARN_, "[WAIT-END] " & m_LastWaitReport, FromProcedureName
            Exit Do
        End If

        CDPHelpers.Sleep POLLING_INTERVAL_SEC
    Loop

    DisarmAll
    If Not WaitSettled And m_RaiseOnTimeout Then printMsg ERROR_, m_LastWaitReport, FromProcedureName, doRaiseError:=True
End Function

'***************************************************************************************************
'* 機能　　：通信の追跡状態を初期化します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：完了イベントが届かず、通信中のまま残ってしまった場合の逃げ道です
'***************************************************************************************************
Public Sub ResetNetwork()
    m_InflightUrl.RemoveAll
    m_InflightStart.RemoveAll
    m_NetArmReq.RemoveAll
    m_LastNetBusyMs = CDPHelpers.TimerCounter
End Sub



'***************************************************************************************************
'                      ■■■ メインとなる機能拡張部分 - イベント処理 ■■■
'***************************************************************************************************
Private Sub ex_CDPContext_ExEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPContext_ExEvent"


    '1. 対象外のイベントは、パースせずに即抜け（`Network.*`は大量に来るため）
    Select Case methodName
        Case "Network.requestWillBeSent", "Network.responseReceived", "Network.loadingFinished", "Network.loadingFailed"
        Case Else
            Exit Sub
    End Select

    '2. 必要なイベントだけパース
    Dim params As BiDiCDPJConv
    Set params = BiDiCDPJConv.Parse(RawJson).NodeKey("params")

    Dim requestId As String
    requestId = params.StringKey("requestId")

    Select Case methodName
        Case "Network.requestWillBeSent"
            '追跡対象の種類で、除外されていなければ数える（リダイレクトは同じrequestIdで再度来るため、未登録時のみ）
            Dim resourceType As String
            If params.ExistsKey("type") Then resourceType = params.StringKey("type")
            If InStr(1, m_TrackTypes, "|" & resourceType & "|", vbTextCompare) > 0 Then
                If Not m_InflightUrl.Exists(requestId) Then
                    m_InflightUrl.Add requestId, params.NodeKey("request").StringKey("url")
                    m_InflightStart.Add requestId, CDPHelpers.TimerCounter
                End If
                m_LastNetBusyMs = CDPHelpers.TimerCounter
            End If

        Case "Network.responseReceived"
            '信号のURL照合
            If m_NetArms.Count > 0 And Not m_NetArmReq.Exists(requestId) Then
                Dim reqUrl As String, k As Variant
                reqUrl = params.NodeKey("response").StringKey("url")
                For Each k In m_NetArms.keys
                    If Not m_NetArms(k) Then
                        If MatchPattern(reqUrl, CStr(k)) Then m_NetArmReq.Add requestId, CStr(k): Exit For
                    End If
                Next
            End If

        Case "Network.loadingFinished"
            If m_NetArmReq.Exists(requestId) Then
                m_NetArms(m_NetArmReq(requestId)) = True
                printMsg info_, "[HIT] network pattern=" & m_NetArmReq(requestId), FromProcedureName
                m_NetArmReq.Remove requestId
            End If
            RemoveInflight requestId

        Case "Network.loadingFailed"
            If m_NetArmReq.Exists(requestId) Then
                printMsg WARN_, "[FAIL] network pattern=" & m_NetArmReq(requestId) & " | " & params.StringKey("errorText") & " (waiting for the next matching response)", FromProcedureName
                m_NetArmReq.Remove requestId
            End If
            RemoveInflight requestId
    End Select
End Sub

Private Sub RemoveInflight(requestId As String)
    If Not m_InflightUrl.Exists(requestId) Then Exit Sub
    m_InflightUrl.Remove requestId
    m_InflightStart.Remove requestId
    m_LastNetBusyMs = CDPHelpers.TimerCounter
End Sub



'***************************************************************************************************
'                           ■■■ 判定の部品 ■■■
'***************************************************************************************************
'* 機能　　：数えるべき通信中の件数を返します
'---------------------------------------------------------------------------------------------------
'* 引数　　：markBusy   `True`で、1件以上あれば「最後に通信が動いた時刻」を今にします
'***************************************************************************************************
Private Function CountInflight(markBusy As Boolean) As Long
    Dim k As Variant, nowMs As Double
    nowMs = CDPHelpers.TimerCounter

    For Each k In m_InflightUrl.keys
        If Not IsIgnoredUrl(CStr(m_InflightUrl(k))) Then
            If m_LongRequestThresholdMs = 0 Or nowMs - m_InflightStart(k) < m_LongRequestThresholdMs Then
                CountInflight = CountInflight + 1
            End If
        End If
    Next

    If markBusy And CountInflight > 0 Then m_LastNetBusyMs = nowMs
End Function

Private Function AllSignalsMet() As Boolean
    Dim k As Variant
    For Each k In m_NetArms.keys
        If Not m_NetArms(k) Then Exit Function
    Next
    If m_ContentArmed And Not m_ContentHit Then Exit Function
    If m_VisArmed And Not m_VisSeen Then Exit Function
    AllSignalsMet = True
End Function

'***************************************************************************************************
'* 機能　　：タイムアウト時の報告用に、揃わなかった条件を列挙します
'***************************************************************************************************
Private Function PendingSummary() As String
    Dim s As String, k As Variant, listed As Long

    For Each k In m_NetArms.keys
        If Not m_NetArms(k) Then s = s & " | net-signal-miss=" & k
    Next
    If m_ContentArmed And Not m_ContentHit Then s = s & " | content-signal-" & IIf(m_ContentLost, "lost", "miss") & "=" & m_ContentXPath
    If m_VisArmed And Not m_VisSeen Then s = s & " | visibility-signal-miss=" & m_VisXPath

    For Each k In m_InflightUrl.keys
        If Not IsIgnoredUrl(CStr(m_InflightUrl(k))) Then
            listed = listed + 1
            If listed <= REPORT_URL_MAX Then
                s = s & vbCrLf & "    inflight(" & Format$(CDPHelpers.TimerCounter - m_InflightStart(k), "0") & "ms): " & Left$(m_InflightUrl(k), 200)
            End If
        End If
    Next
    If listed > REPORT_URL_MAX Then s = s & vbCrLf & "    ... and " & (listed - REPORT_URL_MAX) & " more"

    PendingSummary = s
End Function

Private Function IsIgnoredUrl(reqUrl As String) As Boolean
    Dim v As Variant
    For Each v In m_IgnorePatterns
        If MatchPattern(reqUrl, CStr(v)) Then IsIgnoredUrl = True: Exit Function
    Next
End Function

'***************************************************************************************************
'* 機能　　：URLのパターン照合（部分一致。`*`は任意の文字列。大文字小文字は区別しない）
'***************************************************************************************************
Private Function MatchPattern(reqUrl As String, pattern As String) As Boolean
    Dim parts() As String, i As Long, pos As Long, hit As Long

    If InStr(pattern, "*") = 0 Then
        MatchPattern = (InStr(1, reqUrl, pattern, vbTextCompare) > 0)
        Exit Function
    End If

    parts = Split(pattern, "*")
    pos = 1
    For i = 0 To UBound(parts)
        If LenB(parts(i)) Then
            hit = InStr(pos, reqUrl, parts(i), vbTextCompare)
            If hit = 0 Then Exit Function
            pos = hit + Len(parts(i))
        End If
    Next
    MatchPattern = True
End Function

'***************************************************************************************************
'* 機能　　：ページ側の状態を1回読みます（プローブが無ければ仕込みます）
'---------------------------------------------------------------------------------------------------
'* 返り値  ：読めた場合は`True`。ページ遷移中などで読めなかった場合は`False`
'* 引数　　：fresh       プローブを今仕込んだ（=新しいページ）か
'            quietMs     最後のDOM変化からの経過ms（今仕込んだ場合は0）
'            sigCount    content信号の変化数（未設定・消失時は-1）
'            visNow      visibility信号の要素が今表示されているか（1/0、XPath評価失敗時は-2）
'***************************************************************************************************
Private Function PollPage(ByRef fresh As Boolean, ByRef quietMs As Double, ByRef sigCount As Long, ByRef visNow As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".PollPage"
    Dim r As Variant, parts() As String

    fresh = False: quietMs = 0: sigCount = -1: visNow = 0
    r = EvalWithArgs(PollJs(), exCDP_SpaHelpers.JsArgs(m_IgnoreSelectors, m_VisXPath, m_VisShadow, m_VisNewOnly))
    If VarType(r) = vbError Then
        printMsg WARN_, "The page-side probe script raised an exception.", FromProcedureName
        Exit Function
    End If
    If VarType(r) <> vbString Then Exit Function

    parts = Split(r, ";")
    If UBound(parts) <> 3 Then Exit Function
    fresh = (parts(0) = "F")
    quietMs = CDbl(parts(1))
    sigCount = CLng(parts(2))
    visNow = CLng(parts(3))
    PollPage = True
End Function

Private Sub ClearVisibility()
    m_VisArmed = False
    m_VisSeen = False
    m_VisXPath = vbNullString
    m_VisShadow = False
    m_VisNewOnly = False
End Sub



'***************************************************************************************************
'                           ■■■ ブラウザ側で動かすJavaScript ■■■
'***************************************************************************************************
'* 機能　　：ページ側のプローブ（仕込み＋状態読み取り）です（引数：除外セレクター配列, 表示待ちXPath, shadow）
'---------------------------------------------------------------------------------------------------
'* 返り値  ："<F|K>;<DOM静穏ms>;<content信号の変化数 or -1>;<表示 1/0/-2>"
'* 詳細説明：・`window.__exSpaIdle`が無ければ（=新しいページ）、MutationObserver を仕込んで"F"を返します
'            ・open な Shadow-root も、見つけ次第監視対象に加えます
'            ・除外セレクターに一致する要素（Shadow DOM 内からはホスト要素も辿る）の変化は数えません
'            ・探索(`find`)と可視判定(`vis`)は`exCDP_SpaHelpers`の共通部品です。
'              `find`はプローブ(`p.find`)にも保存し、信号用のJavaScriptから再利用します
'***************************************************************************************************
Private Function PollJs() As String
    Dim js As String
    js = "function(ign,visXp,visShadow,visNewOnly){" & exCDP_SpaHelpers.VisiblePredicateJs() & exCDP_SpaHelpers.FindElementsJs()
    js = js & "let p=window.__exSpaIdle,fresh=false;"
    js = js & "if(!p){fresh=true;p=window.__exSpaIdle={last:performance.now(),ign:[],seen:new WeakSet(),sig:0,sigArmed:false,sigObs:null,find:find};"
    js = js & "const skip=t=>{let e=t&&(t.nodeType===1?t:t.parentElement);if(!e||!p.ign.length)return false;"
    js = js & "while(e){for(const s of p.ign){try{if(e.closest(s))return true;}catch(_){}}const rt=e.getRootNode();e=(rt&&rt.host)?rt.host:null;}return false;};"
    js = js & "const cb=ms=>{let hit=false;for(const m of ms){if(!skip(m.target))hit=true;for(const n of m.addedNodes){if(n.nodeType===1)p.scan(n);}}if(hit)p.last=performance.now();};"
    js = js & "p.watch=r=>{if(!r||p.seen.has(r))return;p.seen.add(r);try{new MutationObserver(cb).observe(r,{subtree:true,childList:true,attributes:true,characterData:true});}catch(_){}p.scan(r);};"
    js = js & "p.scan=n=>{if(n.shadowRoot)p.watch(n.shadowRoot);if(n.querySelectorAll){for(const c of n.querySelectorAll('*')){if(c.shadowRoot)p.watch(c.shadowRoot);}}};"
    js = js & "p.watch(document);}"
    js = js & "p.ign=Array.isArray(ign)?ign:[];"
    js = js & "const q=fresh?0:Math.floor(performance.now()-p.last);"
    js = js & "let v=0;if(visXp){try{v=find(visXp,visShadow).some(e=>vis(e)&&!(visNewOnly&&p.visOld&&p.visOld.has(e)))?1:0;}catch(_){v=-2;}}"
    js = js & "return (fresh?'F':'K')+';'+q+';'+(p.sigArmed?p.sig:-1)+';'+v;}"
    PollJs = js
End Function

'***************************************************************************************************
'* 機能　　：newOnly 用に、今ある一致要素を「古い要素」として記録します（引数：XPath, shadow, 未使用）
'---------------------------------------------------------------------------------------------------
'* 返り値  ：0（成功） / -2（XPath評価失敗） / -3（プローブなし）
'***************************************************************************************************
Private Function SnapshotOldJs() As String
    SnapshotOldJs = "function(xp,shadow){const p=window.__exSpaIdle;if(!p)return -3;try{p.visOld=new WeakSet(p.find(xp,shadow));}catch(_){return -2;}return 0;}"
End Function

'***************************************************************************************************
'* 機能　　：content信号の監視を開始します（引数：XPath, shadow, 属性も数えるか）
'---------------------------------------------------------------------------------------------------
'* 返り値  ：0（成功） / -1（要素なし） / -2（XPath評価失敗） / -3（プローブなし）
'***************************************************************************************************
Private Function ArmContentJs() As String
    Dim js As String
    js = "function(xp,shadow,attrs){const p=window.__exSpaIdle;if(!p)return -3;"
    js = js & "let n;try{n=p.find(xp,shadow)[0];}catch(_){return -2;}if(!n)return -1;"
    js = js & "if(p.sigObs)p.sigObs.disconnect();p.sig=0;p.sigArmed=true;"
    js = js & "const opt={subtree:true,childList:true,characterData:true,attributes:!!attrs};"
    js = js & "p.sigObs=new MutationObserver(ms=>{p.sig+=ms.length;});p.sigObs.observe(n,opt);"
    js = js & "if(n.shadowRoot)p.sigObs.observe(n.shadowRoot,opt);"
    js = js & "for(const c of n.querySelectorAll('*')){if(c.shadowRoot)p.sigObs.observe(c.shadowRoot,opt);}"
    js = js & "return 0;}"
    ArmContentJs = js
End Function

'***************************************************************************************************
'* 機能　　：ページ側の信号用の状態（content信号の監視・newOnlyの古い要素一覧）を解除します
'***************************************************************************************************
Private Function DisarmPageJs() As String
    DisarmPageJs = "(function(){const p=window.__exSpaIdle;if(p){if(p.sigObs)p.sigObs.disconnect();p.sigObs=null;p.sigArmed=false;p.sig=0;p.visOld=null;}return 0;})()"
End Function



'***************************************************************************************************
'                           ■■■ 小物 ■■■
'***************************************************************************************************
'* 機能　　：引数付きのJavaScriptを、要素指定なしで実行します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：実行結果。ページ遷移等で実行場所が失われていた場合は`Empty`
'***************************************************************************************************
Private Function EvalWithArgs(js As String, args As Variant) As Variant
    On Error GoTo Trap

    EvalWithArgs = ex_CDPContext.jsEval(js, , args)
    Exit Function

Trap:
    If exCDP_SpaHelpers.IsStaleError(Err.Description) Then Exit Function
    Err.Raise Err.Number, Err.Source, Err.Description
End Function



'***************************************************************************************************
'                           ■■■ `CDPContext`直通メソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Const LogFileExtension As String = ".cdp.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, CDPHelpers.currentLog.logPath
End Sub

Private Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    ex_CDPContext.ThisCDPBrowser.TakeEvents StopPipeError, destruction
End Sub



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　  ：この拡張機能Classを動かすのに必要なオブジェクトを渡し、通信の監視を開始します
'---------------------------------------------------------------------------------------------------
'* 注意事項：ページ遷移の通信も数えたい場合は、遷移より前に呼んでください
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    Set ex_CDPContext = Inheritance
    ex_CDPContext.ExecuteCDP "Network.enable"
    m_LastNetBusyMs = CDPHelpers.TimerCounter
End Sub

'***************************************************************************************************
'* 機能　　：初期設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    m_DefaultMinStableMs = 500
    m_DefaultTimeoutMs = 10000
    m_RaiseOnTimeout = False
    m_LongRequestThresholdMs = 0
    m_TrackTypes = "|XHR|Fetch|Document|"

    Set m_IgnorePatterns = New Collection
    Set m_IgnoreSelectors = New Collection
    Set m_InflightUrl = New Dictionary
    Set m_InflightStart = New Dictionary
    Set m_NetArms = New Dictionary
    Set m_NetArmReq = New Dictionary

    m_LogID = "SIID" & Format(Rnd * 1000, "000")
End Sub
