Attribute VB_Name = "Demo_SpaIdle"
'***************************************************************************************************
'                 exCDP_SpaIdle 拡張（SPA向け待機・第2段階） - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：SPAで待機が難しい場面を再現したテストページで、`exCDP_SpaIdle`の動作を確認します
'---------------------------------------------------------------------------------------------------
'* 対応拡張  ：Extensions\MainUnit\CDP\CounterplanForSPA\exCDP_SpaIdle.cls
'              （操作には exCDP_SpaElement.cls を、共通部品として exCDP_SpaHelpers.bas を使います）
'* テストHTML：Extensions\OperationCheck\TestHtml\Test_CounterplanForSPA\idle.html
'---------------------------------------------------------------------------------------------------
'* `exCDP_SpaIdle.cls` の公開API：
'   ・WaitSettled([minStableMs], [maxWaitMs])                       … Boolean（タイムアウト時の詳細は LastWaitReport）
'   ・ArmNetworkSignal(urlPattern)                                  … 一致する通信の完了を待つ
'   ・ArmContentSignal(xpath, [searchShadow], [includeAttributes])  … 既存要素の配下の書き換えを待つ
'   ・ArmVisibilitySignal(xpath, [searchShadow], [newOnly])          … 要素の表示を待つ（newOnly: 新しく作られた要素だけ）
'   ・DisarmAll / AddIgnorePattern / AddIgnoreSelector / ClearIgnoreRules / ResetNetwork
'   ・設定：DefaultMinStableMs / DefaultTimeoutMs / RaiseOnTimeout / LongRequestThresholdMs / TrackResourceTypes
'   ・exCDP_SpaElement.AutoWait に設定すると、操作が成功するたびに自動で WaitSettled します
'---------------------------------------------------------------------------------------------------
'* 使い方の基本（arm-then-act）：
'       idle.ArmContentSignal "//ul[@id='results']"   '1. 操作の「前」に、完了の目印を仕掛ける
'       spa.click "//button[@id='search']"            '2. 操作する
'       idle.WaitSettled                               '3. 目印の成立 + 通信とDOMの静穏を待つ
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・WORKSPACE_PATH をご自身の環境に合わせて設定してください
'   ・I4 は TEST-NET-1(192.0.2.1、どこにも繋がらない文書用アドレス) へ接続を試み、終わらない通信を再現します
'***************************************************************************************************
Option Explicit



'ワークスペースパス
'※ StarterWebScrapingKit のルートフォルダを入力してください
Private Const WORKSPACE_PATH As String = ""

Private Const TEST_HTML As String = "\Extensions\OperationCheck\TestHtml\Test_CounterplanForSPA\idle.html"



'***************************************************************************************************
'                 ■■■ Demo 01：全シナリオを実行 ■■■
'***************************************************************************************************
Sub Demo_SpaIdle_01_全シナリオ()
    Dim summary As String
    summary = RunAllIdleScenarios(WORKSPACE_PATH)

    Debug.Print summary
    MsgBox summary, IIf(InStr(summary, "×") > 0, vbExclamation, vbInformation), "exCDP_SpaIdle"
End Sub

'***************************************************************************************************
'* 機能　　：全シナリオを実行し、結果の一覧を返します
'---------------------------------------------------------------------------------------------------
'* 引数　　：workspacePath   StarterWebScrapingKit のルートフォルダ
'* 返り値  ：結果の一覧（各行が ○ / × で始まります）
'***************************************************************************************************
Public Function RunAllIdleScenarios(workspacePath As String) As String
    '1. 空のタブで起動し、遷移の「前」に待機クラスを用意する（遷移中の通信も数えるため）
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("about:blank")
    browserTab.show

    Dim idle As New exCDP_SpaIdle
    idle.Init browserTab
    idle.DefaultTimeoutMs = 8000

    Dim spa As New exCDP_SpaElement
    spa.Init browserTab
    spa.SearchTimeoutSec = 5

    Dim r As New Collection
    Dim ok As Boolean, ok2 As Boolean, report As String

    '--- I0：ページ遷移後に落ち着くこと ---
    browserTab.navigate TestPageUrl(workspacePath)
    ok = idle.WaitSettled
    r.Add Judge("I0 遷移後の待機", CStr(ok), "True")

    '--- I1：応答と描画の間に隙間がある。content信号で描画まで待つこと ---
    idle.ArmContentSignal "//ul[@id='i1list']"
    spa.click "//button[@id='i1btn']"
    ok = idle.WaitSettled
    r.Add Judge("I1 content信号で描画完了まで待つ", CStr(ok) & "/" & JsText(browserTab, "document.querySelectorAll('#i1list li').length"), "True/3")

    '--- I2：network信号で、特定の通信の完了まで待つこと ---
    idle.ArmNetworkSignal "data:application/json"
    spa.click "//button[@id='i2btn']"
    ok = idle.WaitSettled
    r.Add Judge("I2 network信号で通信完了まで待つ", CStr(ok) & "/" & JsText(browserTab, "document.getElementById('i2').textContent"), "True/loaded")

    '--- I2b：来ない通信を待つとタイムアウトし、報告に残ること ---
    idle.ArmNetworkSignal "never-matching-pattern"
    ok = idle.WaitSettled(maxWaitMs:=1500)
    r.Add Judge("I2b 来ない信号はタイムアウトして報告", CStr(ok) & "/" & CStr(InStr(idle.LastWaitReport, "net-signal-miss=never-matching-pattern") > 0), "False/True")

    '--- I3：visibility信号で、表示されるまで待つこと ---
    idle.ArmVisibilitySignal "//div[@id='i3panel']"
    spa.click "//button[@id='i3btn']"
    ok = idle.WaitSettled
    r.Add Judge("I3 visibility信号で表示まで待つ", CStr(ok) & "/" & JsText(browserTab, "String(!document.getElementById('i3panel').hidden)"), "True/true")

    '--- I4：終わらない通信はタイムアウトの原因として報告され、除外すれば落ち着くこと ---
    spa.click "//button[@id='i4btn']"
    ok = idle.WaitSettled(maxWaitMs:=2000)
    report = idle.LastWaitReport
    idle.AddIgnorePattern "192.0.2.1"
    ok2 = idle.WaitSettled
    r.Add Judge("I4 終わらない通信 → 報告 → 除外で解決", CStr(ok) & "/" & CStr(InStr(report, "192.0.2.1") > 0) & "/" & CStr(ok2), "False/True/True")

    '--- I5：AutoWait。操作するだけで、描画完了まで待ってから戻ること ---
    '    2回目のクリックでは一覧を「いったん空にしてから描き直す」ため、content信号は空にした時点で成立してしまう。
    '    そこで、新しく作られた行(li)の表示を待つ visibility信号(newOnly) を使う
    Set spa.AutoWait = idle
    idle.ArmVisibilitySignal "//ul[@id='i1list']/li", newOnly:=True
    spa.click "//button[@id='i1btn']"
    Set spa.AutoWait = Nothing
    r.Add Judge("I5 AutoWait（操作直後に結果が揃っている）", _
                JsText(browserTab, "document.querySelectorAll('#i1list li').length") & "/" & JsText(browserTab, "document.getElementById('i1status').textContent"), "3/done")

    '--- I6：変化し続ける時計は、除外セレクターで判定から外せること ---
    '    タブが hidden だとブラウザがタイマーを約1秒間隔に間引くため、時計の更新の合間に静穏が成立してしまう。
    '    そのため「除外前は落ち着かない」の確認は、visible の時だけ合否に含める
    Dim visState As String
    spa.click "//button[@id='i6btn']"
    ok = idle.WaitSettled(maxWaitMs:=2000)
    visState = JsText(browserTab, "document.visibilityState")
    idle.AddIgnoreSelector "#clock"
    ok2 = idle.WaitSettled
    If visState = "visible" Then
        r.Add Judge("I6 時計 → 落ち着かない → 除外セレクターで解決", CStr(ok) & "/" & CStr(ok2), "False/True")
    Else
        r.Add Judge("I6 時計 → 除外セレクターで解決", CStr(ok2), "True")
        r.Add "  参考：visibilityState=" & visState & " のため、除外前の結果(" & ok & ")は判定対象外です"
    End If

    '2. 後始末
    browserTab.ThisCDPBrowser.CloseBrowser

    '3. 一覧化
    Dim line As Variant, okCount As Long, total As Long
    For Each line In r
        RunAllIdleScenarios = RunAllIdleScenarios & line & vbCrLf
        If Left$(line, 1) = "○" Then okCount = okCount + 1
        If Left$(line, 1) = "○" Or Left$(line, 1) = "×" Then total = total + 1
    Next
    RunAllIdleScenarios = RunAllIdleScenarios & "---" & vbCrLf & okCount & " / " & total & " passed"
End Function



'***************************************************************************************************
'                                   ■■■ 補助 ■■■
'***************************************************************************************************
Private Function TestPageUrl(workspacePath As String) As String
    If LenB(workspacePath) = 0 Then Err.Raise 5, "Demo_SpaIdle", "WORKSPACE_PATH を設定してください。"
    TestPageUrl = "file:///" & Replace(workspacePath & TEST_HTML, "\", "/")
End Function

Private Function JsText(browserTab As CDPContext, expression As String) As String
    JsText = CStr(browserTab.jsEval(expression))
End Function

Private Function Judge(caseTitle As String, actual As String, expected As String) As String
    Judge = IIf(actual = expected, "○ ", "× ") & caseTitle & "  (expected=" & expected & ", actual=" & actual & ")"
End Function
