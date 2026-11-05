VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_SpaElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'                         SPAサイト向けの要素操作アシストを行います
'
'   「どの要素を、今、操作してよいか」の判断と防御をこのClassが担い、
'   実際の操作（クリック・入力・選択）そのものは`CDPElement.cls`に任せます。
'
'       [取得] 可視要素を優先 / Shadow DOM を再帰探索
'       [準備] isConnected確認・スクロール・フォーカス先の特定・ヒットテスト
'       [操作] → CDPElement (SimpleClick / sendClick / sendString / selected ...)
'       [検証] 入力値の照合・選択結果の確認
'       [再試行] 要素が差し替わっていた(stale)ら、同じXPathで取り直して再実行
'
'   参考：ForAI\WebDriver-BiDi-for-SeleniumVBA (BiDiCommandWrapper.cls) の要素操作周り
'
'* 必要なモジュール：exCDP_SpaHelpers.bas（exCDP_SpaIdle.cls と共通の部品）
'
'* 操作後の待機：
'   ・`AutoWait`に`exCDP_SpaIdle`を設定すると、操作が成功するたびに`WaitSettled`で待機します
'   ・未設定の場合は待機しません。呼び出し側で`exCDP_SpaIdle.WaitSettled`等により待ってください
'
'* 範囲外：
'   ・対象はメインフレームのみです（iframe内は未対応）
'   ・closed な Shadow-root 内は探索できません（open のみ）
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'ブラウザ制御オブジェクト
Private ex_CDPContext               As CDPContext   '継承元のタブクラス

'内部使用用
Private Const ThisClassName         As String = "exCDP_SpaElement"
Private Const VISIBLE_CANDIDATE_MAX As Long = 50    '可視判定を行う候補の上限数
Private Const POLLING_INTERVAL_SEC  As Double = 0.1 '要素検索のポーリング間隔(秒)
Private Const STALE_RETRY_WAIT_SEC  As Double = 0.2 'stale検知後、再取得までの待ち時間(秒)

'基本ステート管理
Private m_LogID                     As String       'ログ識別ID

'動作設定
Private m_SearchTimeoutSec          As Double       '要素が見つかるまで待つ秒数
Private m_MaxStaleRetries           As Long         'stale時の再試行回数
Private m_PreferVisible             As Boolean      '複数ヒット時に可視要素を優先するか
Private m_SearchShadow              As Boolean      '既定でShadow DOMも探索するか
Private m_UseTrustedClick           As Boolean      'クリックを物理クリック(`sendClick`)で行うか
Private m_AllowSyntheticFallback    As Boolean      '物理クリック先が覆われていた際、JSクリックへ退避するか
Private m_UserGesture               As Boolean      'CDPElementへ`userGesture`を付与するか
Private m_RaiseOnFailure            As Boolean      '失敗時に`Err.Raise`で止めるか
Private m_AutoWait                  As exCDP_SpaIdle '操作成功後に自動で待機する場合の待機クラス



'***************************************************************************************************
'                               ■■■ 設定プロパティ ■■■
'***************************************************************************************************
'* 機能　　：要素が見つかるまで待つ秒数です（既定：10秒）
'***************************************************************************************************
Public Property Get SearchTimeoutSec() As Double
    SearchTimeoutSec = m_SearchTimeoutSec
End Property
Public Property Let SearchTimeoutSec(v As Double)
    If v < 0 Then v = 0
    m_SearchTimeoutSec = v
End Property

'***************************************************************************************************
'* 機能　　：操作直前に要素が差し替わっていた(stale)場合の再試行回数です（既定：3回）
'***************************************************************************************************
Public Property Get MaxStaleRetries() As Long
    MaxStaleRetries = m_MaxStaleRetries
End Property
Public Property Let MaxStaleRetries(v As Long)
    If v < 0 Then v = 0
    m_MaxStaleRetries = v
End Property

'***************************************************************************************************
'* 機能　　：XPathが複数ヒットした際、最初の「可視」要素を選ぶかです（既定：True）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：可視要素が1つもない場合は、文書順の先頭を選びます
'            可視の定義：DOMに接続済み・サイズあり・visibility:visible・inert祖先なし（opacityは問いません）
'***************************************************************************************************
Public Property Get PreferVisible() As Boolean
    PreferVisible = m_PreferVisible
End Property
Public Property Let PreferVisible(v As Boolean)
    m_PreferVisible = v
End Property

'***************************************************************************************************
'* 機能　　：各メソッドで`searchShadow`を省略した際に、Shadow DOM も探索するかの既定値です（既定：False）
'***************************************************************************************************
Public Property Get DefaultSearchShadow() As Boolean
    DefaultSearchShadow = m_SearchShadow
End Property
Public Property Let DefaultSearchShadow(v As Boolean)
    m_SearchShadow = v
End Property

'***************************************************************************************************
'* 機能　　：クリックを物理クリック(`CDPElement.sendClick`)で行うかです（既定：False = `SimpleClick`）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：物理クリックの前に、要素の中心が最前面にあるか（他の要素に覆われていないか）を確認します
'* 注意事項：サイトによっては物理クリックで挙動が変わる場合があります。比較の上でご利用ください
'***************************************************************************************************
Public Property Get UseTrustedClick() As Boolean
    UseTrustedClick = m_UseTrustedClick
End Property
Public Property Let UseTrustedClick(v As Boolean)
    m_UseTrustedClick = v
End Property

'***************************************************************************************************
'* 機能　　：物理クリック先が他の要素に覆われていた場合、JSクリック(`SimpleClick`)へ退避するかです（既定：True）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`False`の場合は、覆われていた時点で失敗扱いにします
'***************************************************************************************************
Public Property Get AllowSyntheticFallback() As Boolean
    AllowSyntheticFallback = m_AllowSyntheticFallback
End Property
Public Property Let AllowSyntheticFallback(v As Boolean)
    m_AllowSyntheticFallback = v
End Property

'***************************************************************************************************
'* 機能　　：操作を委譲する`CDPElement`へ`SetOptionUserGesture`を付与するかです（既定：False）
'***************************************************************************************************
Public Property Get UseUserGesture() As Boolean
    UseUserGesture = m_UserGesture
End Property
Public Property Let UseUserGesture(v As Boolean)
    m_UserGesture = v
End Property

'***************************************************************************************************
'* 機能　　：失敗時に`Err.Raise`で止めるかです（既定：True）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`False`の場合は、ERRORログを出した上で、各メソッドが`False`を返します
'***************************************************************************************************
Public Property Get RaiseOnFailure() As Boolean
    RaiseOnFailure = m_RaiseOnFailure
End Property
Public Property Let RaiseOnFailure(v As Boolean)
    m_RaiseOnFailure = v
End Property

'***************************************************************************************************
'* 機能　　：操作（click / InputText / SelectOption）が成功するたびに、自動で`WaitSettled`する待機クラスです（既定：Nothing = 待機しない）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・信号（`ArmNetworkSignal`等）は、操作の前に`exCDP_SpaIdle`側へ仕掛けてください。この自動待機で消費されます
'            ・待機の結果は、操作の戻り値には反映しません。`exCDP_SpaIdle.LastWaitReport`で確認してください
'              （タイムアウトで止めたい場合は、`exCDP_SpaIdle.RaiseOnTimeout = True`）
'***************************************************************************************************
Public Property Get AutoWait() As exCDP_SpaIdle
    Set AutoWait = m_AutoWait
End Property
Public Property Set AutoWait(v As exCDP_SpaIdle)
    Set m_AutoWait = v
End Property



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************
'* 機能　　：XPathで要素を検索し、`CDPElement`として返却します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：`CDPElement`。見つからない場合は、`isExist = False`の`CDPElement`
'* 引数　　：xpath          検索したいXPath
'            searchShadow   `True`で open な Shadow-root 内も再帰的に探索します。省略時は`DefaultSearchShadow`プロパティに従います
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・`SearchTimeoutSec`の間、見つかるまでポーリングします
'            ・`CDPContext.getElementByXPath`との違いは「可視要素の優先」と「Shadow DOMの再帰探索」です
'* 注意事項：・`searchShadow:=True`の場合、XPathは`/`始まり（例：`//button[@id='save']`）で指定してください
'            ・返される`CDPElement`は検索JSを持たないため、`onExist`等のポーリング系メソッドは使えません
'***************************************************************************************************
Public Function FindElement(xpath As String, Optional searchShadow As Variant) As CDPElement
    Const FromProcedureName As String = ThisClassName & ".FindElement"


    Dim objectId As String
    objectId = PollResolve(xpath, UseShadow(searchShadow, xpath, FromProcedureName), FromProcedureName)

    Set FindElement = New CDPElement
    If LenB(objectId) Then
        FindElement.Init ex_CDPContext, vbString, objectId
    Else
        FindElement.Init ex_CDPContext, vbNull, vbNullString
        printMsg WARN_, "Element not found within " & m_SearchTimeoutSec & "s: " & xpath, FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：要素をクリックします
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功可否
'* 引数　　：xpath          クリックしたい要素のXPath
'            searchShadow   `FindElement`と同じ
'---------------------------------------------------------------------------------------------------
'* 機能説明：1. 要素を取得（可視優先）
'            2. 接続確認・中央へスクロール・フォーカス
'            3. `UseTrustedClick = True`の場合は、ヒットテストの上で`sendClick`
'               それ以外は`SimpleClick`
'            ※ 2. の時点で要素が差し替わっていたら、1. からやり直します
'* 注意事項：クリック自体の実行後は再試行しません（二重クリックを防ぐため）
'***************************************************************************************************
Public Function click(xpath As String, Optional searchShadow As Variant) As Boolean
    Const FromProcedureName As String = ThisClassName & ".click"


    Dim shadow As Boolean: shadow = UseShadow(searchShadow, xpath, FromProcedureName)
    Dim attempt As Long, stepResult As String
    Dim el As CDPElement

    For attempt = 0 To m_MaxStaleRetries
        '1. 取得
        Set el = ResolveRequired(xpath, shadow, FromProcedureName)
        If el Is Nothing Then Exit Function

        '2-3. 準備とクリック
        stepResult = ClickResolved(el, FromProcedureName)
        Select Case stepResult
            Case "OK"
                click = True
                printMsg info_, "Clicked: " & xpath, FromProcedureName
                AfterAction
                Exit Function
            Case "FAIL"
                Exit Function
        End Select

        '4. staleなので、取り直して再試行
        RetryNotice attempt, xpath, FromProcedureName
    Next

    Fail "The element kept being replaced (stale) and could not be clicked: " & xpath, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：テキストを入力します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功可否
'* 引数　　：xpath          入力したい要素のXPath。クリックすると本物の入力欄が現れるタイプの場合は、その「入口」の要素でもOKです
'            textValue      入力する文字列。空文字の場合はクリアします
'            verify         `True`で、入力後の値が`textValue`と一致するか照合します
'            searchShadow   `FindElement`と同じ
'---------------------------------------------------------------------------------------------------
'* 機能説明：1. 要素を取得（可視優先）し、中央へスクロール・クリック・フォーカス
'            2. 実際にフォーカスが乗ったテキスト欄を、Shadow DOM越しに`document.activeElement`から特定します
'               （同じ要素であり続けることを確認。最大5秒）
'            3. 特定したテキスト欄を`CDPElement`化し、`sendString`（空文字なら`clearValue`）に任せます
'            4. `verify = True`なら値を照合します
'            ※ 途中で要素が差し替わっていたら、1. からやり直します（入力は、クリア→再入力なので再実行しても結果は同じです）
'* 注意事項：・入力内容が自動整形される欄（金額のカンマ区切り等）は、照合で不一致になります。その場合は`verify:=False`にしてください
'            ・`CDPElement.sendString`は入力文字列をログに出力します。パスワード等の入力時はログ設定にご注意ください
'***************************************************************************************************
Public Function InputText(xpath As String, textValue As String, Optional verify As Boolean = True, Optional searchShadow As Variant) As Boolean
    Const FromProcedureName As String = ThisClassName & ".InputText"


    Dim shadow As Boolean: shadow = UseShadow(searchShadow, xpath, FromProcedureName)
    Dim attempt As Long, stepResult As String, verdict As String
    Dim el As CDPElement, textField As CDPElement

    For attempt = 0 To m_MaxStaleRetries
        '1. 取得と準備
        Set el = ResolveRequired(xpath, shadow, FromProcedureName)
        If el Is Nothing Then Exit Function
        el.SetOptionUserGesture = m_UserGesture
        stepResult = RunElementJs(el, PrepInputJs())

        '2. 実際にフォーカスが乗ったテキスト欄を特定
        If stepResult = "OK" Then Set textField = ResolveActiveTextTarget(el, stepResult)

        If stepResult = "OK" Then
            If textField Is Nothing Then
                Fail "No editable text field received focus after clicking: " & xpath & _
                     " (visibilityState/activeElement=" & RunElementJs(el, ActiveElementSummaryJs()) & ")", FromProcedureName
                Exit Function
            End If

            '3. 入力をCDPElementへ委譲
            stepResult = TypeInto(textField, textValue, FromProcedureName)
        End If

        '4. 照合
        If stepResult = "OK" Then
            If Not verify Then
                InputText = True
                printMsg info_, "Typed " & Len(textValue) & " chars (not verified): " & xpath, FromProcedureName
                AfterAction
                Exit Function
            End If

            verdict = RunElementJs(textField, VerifyValueJs(), exCDP_SpaHelpers.JsArgs(textValue))
            Select Case True
                Case verdict = "1"
                    InputText = True
                    printMsg info_, "Typed and verified " & Len(textValue) & " chars: " & xpath, FromProcedureName
                    AfterAction
                    Exit Function
                Case Left$(verdict, 2) = "0:"
                    Fail "Input value mismatch. expected length=" & Len(textValue) & ", actual length=" & Mid$(verdict, 3) & _
                         ". If the field reformats its value, use verify:=False. XPath: " & xpath, FromProcedureName
                    Exit Function
            End Select
            '"STALE"(入力後に差し替わった)は、再試行へ
        ElseIf stepResult = "FAIL" Then
            Exit Function
        End If

        '5. staleなので、取り直して再試行
        RetryNotice attempt, xpath, FromProcedureName
    Next

    Fail "The element kept being replaced (stale) and the text could not be entered: " & xpath, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：選択肢を選びます。ネイティブの`<select>`と、独自ドロップダウン(`role=option`)の両方に対応します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功可否
'* 引数　　：xpath          `<select>`、または独自ドロップダウンを開く要素のXPath
'            optionText     選びたい選択肢の表示テキスト（前後・連続空白は正規化して比較）
'            byValue        `True`で、`<select>`の`value`属性で照合します（独自ドロップダウンでは無視）
'            searchShadow   `FindElement`と同じ。独自ドロップダウンの選択肢探索にも適用します
'---------------------------------------------------------------------------------------------------
'* 機能説明：・`<select>`の場合：選択肢の存在を確認してから、`CDPElement.selected`に任せ、選択結果を確認します
'            ・それ以外の場合：要素をクリックして開き、「可視」かつ「テキストが一致」する`role=option`が
'              ちょうど1つになるのを待ってから、それをクリックします
'* 注意事項：独自ドロップダウンが開かない場合は、`UseTrustedClick = True`もお試しください
'***************************************************************************************************
Public Function SelectOption(xpath As String, optionText As String, Optional byValue As Boolean = False, Optional searchShadow As Variant) As Boolean
    Const FromProcedureName As String = ThisClassName & ".SelectOption"


    Dim shadow As Boolean: shadow = UseShadow(searchShadow, xpath, FromProcedureName)
    Dim attempt As Long, stepResult As String, kind As String
    Dim el As CDPElement
    Dim idx As Variant

    For attempt = 0 To m_MaxStaleRetries
        '1. 取得と種類判定
        Set el = ResolveRequired(xpath, shadow, FromProcedureName)
        If el Is Nothing Then Exit Function
        kind = RunElementJs(el, "function(){ if(!this.isConnected) return 'STALE'; return this.tagName==='SELECT' ? 'NATIVE' : 'CUSTOM'; }")

        Select Case kind
            Case "NATIVE"
                '2-1. 選択肢の存在確認
                idx = RunElementJs(el, FindOptionIndexJs(), exCDP_SpaHelpers.JsArgs(optionText, byValue), -2)
                If idx = -1 Then
                    Fail "Option not found in <select>: " & optionText & " (" & xpath & ")", FromProcedureName
                    Exit Function
                End If

                '2-2. 選択をCDPElementへ委譲し、結果を確認
                If idx >= 0 Then
                    If RunElementJs(el, PrepClickJs()) = "OK" Then
                        el.selected = CStr(idx)
                        If RunElementJs(el, "function(){ return this.isConnected ? String(this.selectedIndex) : 'STALE'; }") = CStr(idx) Then
                            SelectOption = True
                            printMsg info_, "Selected option [" & idx & "] " & optionText & ": " & xpath, FromProcedureName
                            AfterAction
                            Exit Function
                        End If
                    End If
                End If

            Case "CUSTOM"
                '3-1. ドロップダウンを開く
                stepResult = ClickResolved(el, FromProcedureName)
                If stepResult = "FAIL" Then Exit Function

                '3-2. 選択肢をクリック
                If stepResult = "OK" Then
                    If ClickRoleOption(optionText, shadow, FromProcedureName) Then
                        SelectOption = True
                        printMsg info_, "Selected option " & optionText & ": " & xpath, FromProcedureName
                        AfterAction
                    End If
                    Exit Function
                End If
        End Select

        '4. staleなので、取り直して再試行
        RetryNotice attempt, xpath, FromProcedureName
    Next

    Fail "The element kept being replaced (stale) and the option could not be selected: " & xpath, FromProcedureName
End Function



'***************************************************************************************************
'                          ■■■ 各ステップの内部処理 ■■■
'***************************************************************************************************
'* 機能　　：取得済み要素に対して、準備→クリックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ："OK" / "STALE"（取り直しが必要） / "FAIL"（失敗報告済み）
'***************************************************************************************************
Private Function ClickResolved(el As CDPElement, From As String) As String
    el.SetOptionUserGesture = m_UserGesture

    '1. 物理クリックの場合は、ヒットテストで最前面か確認
    If m_UseTrustedClick Then
        Dim verdict As String
        verdict = RunElementJs(el, HitTestJs())

        Select Case True
            Case verdict = "STALE"
                ClickResolved = "STALE"
                Exit Function

            Case verdict = "hit=1"
                el.sendClick
                ClickResolved = "OK"
                Exit Function

            Case Left$(verdict, 6) = "hit=0;"
                '覆われている場合に座標クリックすると、手前の要素を押してしまうため退避する
                If Not m_AllowSyntheticFallback Then
                    Fail "[TRUSTED-HITTEST] The element center is not topmost and synthetic fallback is disabled. (" & Mid$(verdict, 7) & ")", From
                    ClickResolved = "FAIL"
                    Exit Function
                End If
                printMsg WARN_, "[TRUSTED-HITTEST] hit=0 | " & Mid$(verdict, 7) & " | outcome=synthetic-fallback", From

            Case Else
                Fail "[TRUSTED-HITTEST] Indeterminate hit-test result: " & verdict, From
                ClickResolved = "FAIL"
                Exit Function
        End Select
    End If

    '2. 準備（接続確認・中央へスクロール・フォーカス）
    If RunElementJs(el, PrepClickJs()) <> "OK" Then
        ClickResolved = "STALE"
        Exit Function
    End If

    '3. クリックをCDPElementへ委譲
    '※ 準備とクリックの間は1往復分の隙間があります。この間に差し替わると、切り離された要素へのクリックとなり、何も起きません
    el.SimpleClick
    ClickResolved = "OK"
End Function

'***************************************************************************************************
'* 機能　　：特定済みのテキスト欄へ入力します
'---------------------------------------------------------------------------------------------------
'* 返り値  ："OK" / "STALE" / "FAIL"
'***************************************************************************************************
Private Function TypeInto(textField As CDPElement, textValue As String, From As String) As String
    textField.SetOptionUserGesture = m_UserGesture

    '1. input/textarea か、contenteditable かを判定
    Dim kind As String
    kind = RunElementJs(textField, "function(){ if(!this.isConnected) return 'STALE'; return (this.tagName==='INPUT'||this.tagName==='TEXTAREA') ? 'FIELD' : 'EDITABLE'; }")
    If kind = "STALE" Then TypeInto = "STALE": Exit Function

    '2. 入力をCDPElementへ委譲
    If kind = "FIELD" Then
        If LenB(textValue) = 0 Then textField.clearValue Else textField.sendString textValue
    Else
        'contenteditable：中身を全選択して、選択範囲ごと置き換える
        textField.selectText
        If LenB(textValue) = 0 Then
            RunElementJs textField, "function(){ document.execCommand('delete'); return 'OK'; }"
        Else
            textField.sendString textValue
        End If
    End If

    TypeInto = "OK"
End Function

'***************************************************************************************************
'* 機能　　：クリック後、実際にフォーカスが乗ったテキスト欄を`CDPElement`として返します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：`CDPElement`。見つからない場合は`Nothing`
'* 引数　　：el       クリックした要素
'            status   "OK" / "STALE" を返します
'***************************************************************************************************
Private Function ResolveActiveTextTarget(el As CDPElement, ByRef stepResult As String) As CDPElement
    On Error GoTo Trap

    Dim r As Variant
    r = el.jsEval(ActiveTextTargetJs(), , Null, awaitPromise:=True)
    stepResult = "OK"

    If VarType(r) = vbString Then
        Set ResolveActiveTextTarget = New CDPElement
        ResolveActiveTextTarget.Init ex_CDPContext, vbString, CStr(r)
    End If
    Exit Function

Trap:
    If exCDP_SpaHelpers.IsStaleError(Err.Description) Then stepResult = "STALE": Exit Function
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'***************************************************************************************************
'* 機能　　：可視・テキスト一致の`role=option`がちょうど1つになるのを待って、クリックします
'***************************************************************************************************
Private Function ClickRoleOption(optionText As String, shadow As Boolean, From As String) As Boolean
    Dim attempt As Long, stepResult As String, lastCount As Double
    Dim r As Variant, t0 As Double
    Dim opt As CDPElement

    For attempt = 0 To m_MaxStaleRetries
        '1. ちょうど1つ見つかるまでポーリング
        Set opt = Nothing
        t0 = CDPHelpers.TimerCounter
        Do
            r = EvalWithArgs(FindRoleOptionJs(), exCDP_SpaHelpers.JsArgs(optionText, shadow))
            If VarType(r) = vbString Then
                Set opt = New CDPElement
                opt.Init ex_CDPContext, vbString, CStr(r)
                Exit Do
            End If
            If IsNumeric(r) Then lastCount = r
            If CDPHelpers.TimerCounter - t0 > m_SearchTimeoutSec * 1000 Then Exit Do
            TakeEvents
            CDPHelpers.Sleep POLLING_INTERVAL_SEC
        Loop

        If opt Is Nothing Then
            If lastCount > 1 Then
                Fail "Multiple visible role=option elements match '" & optionText & "' (" & lastCount & "). Narrow down the dropdown first.", From
            Else
                Fail "No visible role=option matched '" & optionText & "' within " & m_SearchTimeoutSec & "s. The dropdown may not have opened.", From
            End If
            Exit Function
        End If

        '2. クリック
        stepResult = ClickResolved(opt, From)
        If stepResult = "OK" Then ClickRoleOption = True: Exit Function
        If stepResult = "FAIL" Then Exit Function

        RetryNotice attempt, "role=option '" & optionText & "'", From
    Next

    Fail "The option kept being replaced (stale) and could not be clicked: " & optionText, From
End Function



'***************************************************************************************************
'                           ■■■ 要素の取得 ■■■
'***************************************************************************************************
'* 機能　　：`SearchTimeoutSec`の間ポーリングし、見つからなければ失敗報告して`Nothing`を返します
'***************************************************************************************************
Private Function ResolveRequired(xpath As String, shadow As Boolean, From As String) As CDPElement
    Dim objectId As String
    objectId = PollResolve(xpath, shadow, From)

    If LenB(objectId) = 0 Then
        Fail "Element not found within " & m_SearchTimeoutSec & "s: " & xpath, From
        Exit Function
    End If

    Set ResolveRequired = New CDPElement
    ResolveRequired.Init ex_CDPContext, vbString, objectId
End Function

'***************************************************************************************************
'* 機能　　：見つかるまで`ResolveOnce`を繰り返します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：`objectId`。タイムアウト時は空文字
'***************************************************************************************************
Private Function PollResolve(xpath As String, shadow As Boolean, From As String) As String
    Dim t0 As Double: t0 = CDPHelpers.TimerCounter
    Dim objectId As String

    Do
        objectId = ResolveOnce(xpath, shadow, From)
        If LenB(objectId) Then PollResolve = objectId: Exit Function
        If CDPHelpers.TimerCounter - t0 > m_SearchTimeoutSec * 1000 Then Exit Function

        TakeEvents
        CDPHelpers.Sleep POLLING_INTERVAL_SEC
    Loop
End Function

'***************************************************************************************************
'* 機能　　：1回だけ検索します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：`objectId`。見つからない場合や、ページ遷移中で検索できなかった場合は空文字
'* 注意事項：XPathの構文エラーは、設定に関わらずエラーで停止します
'***************************************************************************************************
Private Function ResolveOnce(xpath As String, shadow As Boolean, From As String) As String
    Dim r As Variant
    r = EvalWithArgs(ResolveJs(), exCDP_SpaHelpers.JsArgs(xpath, shadow, m_PreferVisible))

    Select Case VarType(r)
        Case vbString
            ResolveOnce = r
        Case vbError
            printMsg ERROR_, "The XPath could not be evaluated (syntax error?): " & xpath, From, doRaiseError:=True
    End Select
End Function

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
'* 機能　　：要素に対してJavaScriptを実行します。要素が失われていた場合は`stale`として扱います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：実行結果。要素(objectId)が失われていた場合や、JavaScript例外時は`ifStale`
'***************************************************************************************************
Private Function RunElementJs(el As CDPElement, js As String, Optional args As Variant, Optional ifStale As Variant = "STALE") As Variant
    On Error GoTo Trap

    RunElementJs = el.jsEval(js, args, ifStale)
    Exit Function

Trap:
    If exCDP_SpaHelpers.IsStaleError(Err.Description) Then RunElementJs = ifStale: Exit Function
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'***************************************************************************************************
'* 機能　　：`searchShadow`引数を解決し、Shadow探索時のXPath書式を確認します
'***************************************************************************************************
Private Function UseShadow(searchShadow As Variant, xpath As String, From As String) As Boolean
    If IsMissing(searchShadow) Then UseShadow = m_SearchShadow Else UseShadow = CBool(searchShadow)

    If UseShadow And Not exCDP_SpaHelpers.IsShadowXPath(xpath) Then
        printMsg ERROR_, "With searchShadow, the XPath must start with '/' (e.g. //button[@id='save']): " & xpath, From, doRaiseError:=True
    End If
End Function



'***************************************************************************************************
'                           ■■■ ブラウザ側で動かすJavaScript ■■■
'***************************************************************************************************
'* 機能　　：XPathで要素を1つ選んで返します（引数：xp, shadow, preferVisible）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・探索(`find`)と可視判定(`vis`)は`exCDP_SpaHelpers`の共通部品です
'            ・preferVisible=true の場合、先頭から`VISIBLE_CANDIDATE_MAX`件のうち最初の可視要素を選びます
'***************************************************************************************************
Private Function ResolveJs() As String
    Dim js As String
    js = "function(xp,shadow,preferVisible){" & exCDP_SpaHelpers.VisiblePredicateJs() & exCDP_SpaHelpers.FindElementsJs()
    js = js & "const all=find(xp,shadow);if(!all.length)return null;"
    js = js & "if(preferVisible){const n=Math.min(all.length," & VISIBLE_CANDIDATE_MAX & ");for(let i=0;i<n;i++){if(vis(all[i]))return all[i];}}"
    js = js & "return all[0];}"
    ResolveJs = js
End Function

'***************************************************************************************************
'* 機能　　：クリック前の準備（接続確認・中央へスクロール・フォーカス）
'***************************************************************************************************
Private Function PrepClickJs() As String
    Dim js As String
    js = "function(){try{if(!this.isConnected)return 'STALE';"
    js = js & "this.scrollIntoView({block:'center',inline:'center'});"
    js = js & "if(typeof this.focus==='function')this.focus({preventScroll:true});"
    js = js & "return this.isConnected?'OK':'STALE';}catch(e){return 'STALE';}}"
    PrepClickJs = js
End Function

'***************************************************************************************************
'* 機能　　：入力前の準備（接続確認・中央へスクロール・クリック・フォーカス）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：クリックで本物の入力欄が現れるタイプのUIに対応するため、クリックまで行います
'***************************************************************************************************
Private Function PrepInputJs() As String
    Dim js As String
    js = "function(){try{if(!this.isConnected)return 'STALE';"
    js = js & "this.scrollIntoView({block:'center',inline:'center'});"
    js = js & "if(!this.isConnected)return 'STALE';this.click();"
    js = js & "if(this.isConnected&&typeof this.focus==='function')this.focus({preventScroll:true});"
    js = js & "return 'OK';}catch(e){return 'STALE';}}"
    PrepInputJs = js
End Function

'***************************************************************************************************
'* 機能　　：実際にフォーカスが乗ったテキスト欄を返します（非同期・最大5秒）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・`document.activeElement`を Shadow-root 越しに辿ります
'            ・同じ要素を2回以上観測し、かつ50ms以上アクティブであり続ければ確定とします
'              （回数だけで判定しないのは、非表示のタブでタイマーが約1秒間隔に間引かれても成立させるためです）
'            ・見つからない場合、クリックした要素自体がテキスト欄ならそれを、そうでなければ null を返します
'***************************************************************************************************
Private Function ActiveTextTargetJs() As String
    Dim js As String
    js = "async function(){const el=this;"
    js = js & "const deep=()=>{let a=document.activeElement;while(a&&a.shadowRoot&&a.shadowRoot.activeElement)a=a.shadowRoot.activeElement;return a;};"
    js = js & "const ok=a=>!!(a&&a.isConnected&&!a.disabled&&!a.readOnly&&(a.tagName==='TEXTAREA'||"
    js = js & "(a.tagName==='INPUT'&&/^(text|search|url|tel|email|password|number)?$/i.test(a.getAttribute('type')||''))||a.isContentEditable));"
    js = js & "let prev=null,since=0,seen=0;"
    js = js & "for(const t0=performance.now();performance.now()-t0<5000;){const a=deep(),now=performance.now();"
    js = js & "if(ok(a)){if(a===prev){if(++seen>=2&&now-since>=50)return a;}else{prev=a;since=now;seen=1;}}else{prev=null;seen=0;}"
    js = js & "await new Promise(r=>setTimeout(r,20));}"
    js = js & "return ok(el)?el:null;}"
    ActiveTextTargetJs = js
End Function

'***************************************************************************************************
'* 機能　　：失敗時の診断用に、ページの可視状態と現在のアクティブ要素を返します
'---------------------------------------------------------------------------------------------------
'* 返り値  ："<visibilityState>/<タグ名>#<id>"
'***************************************************************************************************
Private Function ActiveElementSummaryJs() As String
    Dim js As String
    js = "function(){let a=document.activeElement;while(a&&a.shadowRoot&&a.shadowRoot.activeElement)a=a.shadowRoot.activeElement;"
    js = js & "return document.visibilityState+'/'+(a?a.tagName+(a.id?'#'+a.id:''):'none');}"
    ActiveElementSummaryJs = js
End Function

'***************************************************************************************************
'* 機能　　：入力後の値を照合します（引数：期待値）
'---------------------------------------------------------------------------------------------------
'* 返り値  ："1"（一致） / "0:<実際の文字数>"（不一致） / "STALE"（入力後に切り離された）
'* 詳細説明：contenteditable は innerText の末尾空白・改行を除いて比較します
'***************************************************************************************************
Private Function VerifyValueJs() As String
    Dim js As String
    js = "function(v){if(!this.isConnected)return 'STALE';"
    js = js & "const field=(this.tagName==='INPUT'||this.tagName==='TEXTAREA');"
    js = js & "const cur=String((field?this.value:this.innerText)||'');"
    js = js & "const n=s=>String(s).replace(/\u00A0/g,' ').replace(/\s+$/,'');"
    js = js & "const same=field?(cur===String(v)):(n(cur)===n(v));"
    js = js & "return same?'1':'0:'+cur.length;}"
    VerifyValueJs = js
End Function

'***************************************************************************************************
'* 機能　　：`<select>`内の選択肢のindexを返します（引数：テキスト or 値, byValue）
'---------------------------------------------------------------------------------------------------
'* 返り値  ：index / -1（なし） / -2（切り離されていた）
'***************************************************************************************************
Private Function FindOptionIndexJs() As String
    Dim js As String
    js = "function(t,byValue){if(!this.isConnected)return -2;"
    js = js & "const n=s=>String(s).replace(/[\s\u00A0]+/g,' ').trim(),k=n(t);"
    js = js & "for(let i=0;i<this.options.length;i++){const o=this.options[i];"
    js = js & "if(byValue?(o.value===String(t)):(n(o.text)===k))return i;}return -1;}"
    FindOptionIndexJs = js
End Function

'***************************************************************************************************
'* 機能　　：可視・テキスト一致の`role=option`を探します（引数：テキスト, shadow）
'---------------------------------------------------------------------------------------------------
'* 返り値  ：ちょうど1つなら、その要素。それ以外は一致数（0 or 2以上）
'***************************************************************************************************
Private Function FindRoleOptionJs() As String
    Dim js As String
    js = "function(t,shadow){" & exCDP_SpaHelpers.VisiblePredicateJs() & exCDP_SpaHelpers.FindElementsJs()
    js = js & "const n=s=>String(s||'').replace(/[\s\u00A0]+/g,' ').trim(),k=n(t);"
    js = js & "const hits=[];for(const r of (shadow?roots():[document])){for(const o of r.querySelectorAll('[role=""option""]')){if(vis(o)&&n(o.textContent)===k)hits.push(o);}}"
    js = js & "return hits.length===1?hits[0]:hits.length;}"
    FindRoleOptionJs = js
End Function

'***************************************************************************************************
'* 機能　　：物理クリック前のヒットテストです
'---------------------------------------------------------------------------------------------------
'* 返り値  ："hit=1" / "hit=0;<理由>[;<手前の要素>]" / "STALE"
'* 詳細説明：`CDPElement.sendClick`と同じ`scrollIntoView(false)`・同じ中心座標の計算で判定するため、
'            判定した位置と実際にクリックする位置が一致します
'***************************************************************************************************
Private Function HitTestJs() As String
    Dim js As String
    js = "function(){if(!this.isConnected)return 'STALE';this.scrollIntoView(false);"
    js = js & "const t=this.getClientRects()[0];if(!t)return 'hit=0;zero-box';"
    js = js & "const l=Math.max(0,Math.min(t.x,t.x+t.width)),r=Math.min(window.innerWidth,Math.max(t.x,t.x+t.width));"
    js = js & "const u=Math.max(0,Math.min(t.y,t.y+t.height)),d=Math.min(window.innerHeight,Math.max(t.y,t.y+t.height));"
    js = js & "if(r<=l||d<=u)return 'hit=0;out-of-view';"
    js = js & "const cx=l+((r-l)>>1),cy=u+((d-u)>>1);"
    js = js & "let top=document.elementFromPoint(cx,cy),g=0;"
    js = js & "while(top&&top.shadowRoot&&g++<10){const s=top.shadowRoot.elementFromPoint(cx,cy);if(!s||s===top)break;top=s;}"
    js = js & "if(!top)return 'hit=0;no-element';"
    js = js & "let x=top,h=0;while(x&&h++<50){if(x===this||this.contains(x))return 'hit=1';"
    js = js & "const rt=x.getRootNode?x.getRootNode():null;if(rt&&rt.host){x=rt.host;continue;}break;}"
    js = js & "const c=(typeof top.className==='string')?top.className.trim().split(/\s+/).slice(0,2).join('.'):'';"
    js = js & "return 'hit=0;covered;'+top.tagName+(top.id?'#'+top.id:'')+(c?'.'+c:'');}"
    HitTestJs = js
End Function



'***************************************************************************************************
'                           ■■■ 小物 ■■■
'***************************************************************************************************
'* 機能　　：stale検知時の再試行ログと待機です
'***************************************************************************************************
Private Sub RetryNotice(attempt As Long, what As String, From As String)
    If attempt >= m_MaxStaleRetries Then Exit Sub
    printMsg WARN_, "The element was replaced (stale). Re-resolving " & (attempt + 1) & "/" & m_MaxStaleRetries & ": " & what, From
    TakeEvents
    CDPHelpers.Sleep STALE_RETRY_WAIT_SEC
End Sub

'***************************************************************************************************
'* 機能　　：操作成功後の自動待機です（`AutoWait`が設定されている場合のみ）
'***************************************************************************************************
Private Sub AfterAction()
    If Not m_AutoWait Is Nothing Then m_AutoWait.WaitSettled
End Sub

'***************************************************************************************************
'* 機能　　：失敗を報告します。`RaiseOnFailure = True`ならエラーで停止します
'***************************************************************************************************
Private Sub Fail(strMsg As String, From As String)
    printMsg ERROR_, strMsg, From, doRaiseError:=m_RaiseOnFailure
End Sub



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
'* 機能　  ：この拡張機能Classを動かすのに必要なオブジェクトを渡します
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    Set ex_CDPContext = Inheritance
End Sub

'***************************************************************************************************
'* 機能　　：初期設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    m_SearchTimeoutSec = 10
    m_MaxStaleRetries = 3
    m_PreferVisible = True
    m_SearchShadow = False
    m_UseTrustedClick = False
    m_AllowSyntheticFallback = True
    m_UserGesture = False
    m_RaiseOnFailure = True

    m_LogID = "SPID" & Format(Rnd * 1000, "000")
End Sub
