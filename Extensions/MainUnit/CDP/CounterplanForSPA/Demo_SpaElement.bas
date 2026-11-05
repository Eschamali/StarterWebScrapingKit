Attribute VB_Name = "Demo_SpaElement"
'***************************************************************************************************
'                 exCDP_SpaElement 拡張（SPA向け要素操作・第1段階） - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：SPAサイトでよくある「要素操作の罠」を再現したテストページで、`exCDP_SpaElement`の動作を確認します
'---------------------------------------------------------------------------------------------------
'* 対応拡張  ：Extensions\MainUnit\CDP\CounterplanForSPA\exCDP_SpaElement.cls
'              （共通部品の exCDP_SpaHelpers.bas も必要です）
'* テストHTML：Extensions\OperationCheck\TestHtml\Test_CounterplanForSPA\index.html
'---------------------------------------------------------------------------------------------------
'* 役割分担：
'   ・exCDP_SpaElement … 取得（可視優先・Shadow再帰）/ 準備 / 検証 / stale時の再試行
'   ・CDPElement       … 実際の操作（SimpleClick / sendClick / sendString / clearValue / selected）
'---------------------------------------------------------------------------------------------------
'* `exCDP_SpaElement.cls` の公開API：
'   ・FindElement(xpath, [searchShadow])                         … CDPElement を返す（見つからなければ isExist=False）
'   ・click(xpath, [searchShadow])                               … Boolean
'   ・InputText(xpath, textValue, [verify=True], [searchShadow])      … Boolean
'   ・SelectOption(xpath, optionText, [byValue], [searchShadow]) … Boolean（<select> と role=option の両対応）
'   ・設定：SearchTimeoutSec / MaxStaleRetries / PreferVisible / DefaultSearchShadow / UseTrustedClick /
'           AllowSyntheticFallback / UseUserGesture / RaiseOnFailure / AutoWait
'---------------------------------------------------------------------------------------------------
'* 検証方針：
'   成否は、テストページ上の結果欄（#r1 等）に「アプリ側の処理が実際に動いた結果」が出ているかで判定します
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・WORKSPACE_PATH をご自身の環境に合わせて設定してください
'   ・第1段階のため、操作後の待機（SPAアイドル判定）は行いません
'***************************************************************************************************
Option Explicit



'ワークスペースパス
'※ StarterWebScrapingKit のルートフォルダを入力してください
Private Const WORKSPACE_PATH As String = ""

Private Const TEST_HTML As String = "\Extensions\OperationCheck\TestHtml\Test_CounterplanForSPA\index.html"



'***************************************************************************************************
'                 ■■■ Demo 01：全シナリオを exCDP_SpaElement で実行 ■■■
'***************************************************************************************************
'* 機能　　：テストページの全シナリオを実行し、結果を一覧表示します
'***************************************************************************************************
Sub Demo_SpaElement_01_全シナリオ()
    Dim summary As String
    summary = RunAllScenarios(WORKSPACE_PATH)

    Debug.Print summary
    MsgBox summary, IIf(InStr(summary, "×") > 0, vbExclamation, vbInformation), "exCDP_SpaElement"
End Sub

'***************************************************************************************************
'* 機能　　：全シナリオを実行し、結果の一覧を返します
'---------------------------------------------------------------------------------------------------
'* 引数　　：workspacePath   StarterWebScrapingKit のルートフォルダ
'* 返り値  ：結果の一覧（各行が ○ / × で始まります）
'***************************************************************************************************
Public Function RunAllScenarios(workspacePath As String) As String
    '1. テストページを開く
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext(TestPageUrl(workspacePath))
    browserTab.show

    Dim spa As New exCDP_SpaElement
    spa.Init browserTab
    spa.SearchTimeoutSec = 5

    Dim r As New Collection

    '--- S1：非表示の複製。可視の方が押されること ---
    spa.click "//button[@class='save']"
    r.Add Judge("S1 非表示の複製を避けてクリック", ResultOf(browserTab, "r1"), "visible")

    '--- S2：フォーカス時に差し替わるボタン。取り直して新しい方が押されること ---
    spa.click "//button[@id='rerender']"
    r.Add Judge("S2 stale を検知して取り直し", ResultOf(browserTab, "r2"), "new")

    '--- S3：クリックで本物の入力欄に差し替わるUI。差し替わった先へ入力されること ---
    spa.InputText "//div[@id='fakeSearch']", "Sapporo"
    r.Add Judge("S3 フォーカス先の入力欄へ入力", ResultOf(browserTab, "r3"), "Sapporo")

    '--- S4：input イベントでのみモデルが更新される入力欄 ---
    spa.InputText "//input[@id='controlled']", "Paris"
    spa.click "//button[@class='save']"    'blur させて、モデルの値が書き戻されても残ることを確認
    r.Add Judge("S4 モデルまで値が届く", ResultOf(browserTab, "r4") & "/" & spa.FindElement("//input[@id='controlled']").value, "Paris/Paris")

    '--- S5：オーバーレイに覆われたボタン。ヒットテストで検知し、ボタン自体が押されること ---
    spa.UseTrustedClick = True
    spa.click "//button[@id='covered']"
    spa.UseTrustedClick = False
    r.Add Judge("S5 覆われたボタン（物理クリック→退避）", ResultOf(browserTab, "r5"), "button")

    '--- S6：独自ドロップダウン ---
    spa.SelectOption "//div[@id='combo']", "Tokyo"
    r.Add Judge("S6 独自ドロップダウン", ResultOf(browserTab, "r6"), "Tokyo")

    '--- S7：ネイティブ select（表示テキスト / value）---
    spa.SelectOption "//select[@id='sel']", "France"
    r.Add Judge("S7 select（テキスト指定）", ResultOf(browserTab, "r7"), "fr")
    spa.SelectOption "//select[@id='sel']", "jp", byValue:=True
    r.Add Judge("S7 select（value指定）", ResultOf(browserTab, "r7"), "jp")

    '--- S8：Shadow DOM ---
    spa.click "//button[@id='shadowBtn']", searchShadow:=True
    r.Add Judge("S8 Shadow DOM 内のクリック", ResultOf(browserTab, "r8"), "clicked")
    spa.InputText "//input[@id='shadowInput']", "hello", searchShadow:=True
    r.Add Judge("S8 Shadow DOM 内の入力", ResultOf(browserTab, "r9"), "hello")

    '--- S10：自動整形される入力欄。照合で不一致を検出し、verify:=False なら通ること ---
    Dim mismatchDetected As Boolean
    spa.RaiseOnFailure = False
    mismatchDetected = Not spa.InputText("//input[@id='amount']", "1234567")
    spa.RaiseOnFailure = True
    r.Add Judge("S10 自動整形による不一致を検出", CStr(mismatchDetected), "True")
    r.Add Judge("S10 verify:=False なら成功扱い", CStr(spa.InputText("//input[@id='amount']", "7654321", verify:=False)) & "/" & ResultOf(browserTab, "r10"), "True/7,654,321")

    '2. 後始末
    browserTab.ThisCDPBrowser.CloseBrowser

    '3. 一覧化
    Dim line As Variant, okCount As Long
    For Each line In r
        RunAllScenarios = RunAllScenarios & line & vbCrLf
        If Left$(line, 1) = "○" Then okCount = okCount + 1
    Next
    RunAllScenarios = RunAllScenarios & "---" & vbCrLf & okCount & " / " & r.Count & " passed"
End Function



'***************************************************************************************************
'          ■■■ Demo 02：比較用。CDPElement だけで同じ操作をすると、どうなるか ■■■
'***************************************************************************************************
'* 機能　　：`CDPContext.getElementByXPath` + `CDPElement` だけで S1 / S5 を操作し、罠にかかる様子を確認します
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - S1：先頭（非表示）のボタンが押され、結果が "hidden" になる
'   - S5：物理クリック(`sendClick`)が手前のオーバーレイに当たり、結果が "overlay" になる
'   ※ どちらも CDPElement の不具合ではなく、「要素が正しく、今操作してよい」という前提が SPA では崩れる例です
'***************************************************************************************************
Sub Demo_SpaElement_02_CDPElementだけの場合()
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext(TestPageUrl(WORKSPACE_PATH))
    browserTab.show

    browserTab.getElementByXPath("//button[@class='save']").SimpleClick
    Debug.Print "[Demo02] S1 結果: " & ResultOf(browserTab, "r1") & "（exCDP_SpaElement なら visible）"

    browserTab.getElementByXPath("//button[@id='covered']").sendClick
    Debug.Print "[Demo02] S5 結果: " & ResultOf(browserTab, "r5") & "（exCDP_SpaElement なら button）"

    MsgBox "イミディエイトウィンドウで結果を確認してください。", vbInformation
    browserTab.ThisCDPBrowser.CloseBrowser
End Sub



'***************************************************************************************************
'                                   ■■■ 補助 ■■■
'***************************************************************************************************
Private Function TestPageUrl(workspacePath As String) As String
    If LenB(workspacePath) = 0 Then Err.Raise 5, "Demo_SpaElement", "WORKSPACE_PATH を設定してください。"
    TestPageUrl = "file:///" & Replace(workspacePath & TEST_HTML, "\", "/")
End Function

Private Function ResultOf(browserTab As CDPContext, id As String) As String
    ResultOf = browserTab.getElementByID(id).innerText
End Function

Private Function Judge(caseTitle As String, actual As String, expected As String) As String
    Judge = IIf(actual = expected, "○ ", "× ") & caseTitle & "  (expected=" & expected & ", actual=" & actual & ")"
End Function
