Attribute VB_Name = "Test_BiDiElement"
'==============================================================================
' PageElement.cls（BiDiモード）動作確認 ＋ WebDriverBiDiContext の検索口テスト
' ・HTML: ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html（CDP版と共用）
' ・WORKSPACE_PATH をルートに設定してから RunAll_BiDiElement_Tests を実行
' 　※空の場合は、このブックの1つ上の階層をルートとみなします（build\配下に置いてる想定）
' ・CDPコマンド依存だったメソッド（sendClick / sendKey / sendString / sendHover / SetFileInputFiles / GetShadowRoot）は、BiDiコマンドでの動作を確認します
'==============================================================================
Option Explicit

Private passCount As Long
Private failCount As Long
Private failNames As String

'ワークスペースパス（StarterWebScrapingKit ルート）
Private Const WORKSPACE_PATH As String = ""

Private Function EOk() As String
    EOk = WorksheetFunction.Unichar(9989)
End Function

Private Function WorkspacePath() As String
    If LenB(WORKSPACE_PATH) Then
        WorkspacePath = WORKSPACE_PATH
    Else
        WorkspacePath = Left$(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") - 1)
    End If
End Function

'==============================================================================
' エントリ
'==============================================================================
Public Sub RunAll_BiDiElement_Tests()
    Dim br As WebDriverBiDiContext: Set br = ShSetting01_StartBrowser.StartBiDiModeContext

    br.navigate "file:///" & Replace(WorkspacePath & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html", "\", "/")
    br.Wait

    passCount = 0: failCount = 0: failNames = vbNullString

    PrintHeader "PageElement(BiDiモード) 検証テスト 開始"

    Test01_ContextSearch br
    Test02_Value_ClearValue br
    Test03_InnerText_InnerHTML br
    Test04_Checked br
    Test05_Selected_SetSelection br
    Test06_Click_SimpleClick_FireEvent br
    Test07_GetAttribute_SetAttribute br
    Test08_Focus_SelectText br
    Test09_Submit br
    Test10_Traversal_Parent_Siblings_FirstChild br
    Test11_GetChildren_ElementsByQuery_ElementsByXPath br
    Test12_ElementByID_Query_XPath_Scoped br
    Test13_IsExist_IfExist_OnExist_OnExistNot br
    Test14_IFrame_GetTab br
    Test15_Diagnostics_And_Options br
    Test16_SendClick_SendKey br
    Test17_SendHover_HoverReveal br
    Test18_SetFileInputFiles br
    Test19_ShadowRoot_Open_Closed br

    PrintHeader "テスト結果: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)

    br.jsEval "updateStatus('s-summary','PASS=" & passCount & " FAIL=" & failCount & " " & EOk() & "', true)"

    br.ThisWebDriverBiDiMode.CloseBrowser

    '自動実行時に結果を判別できるよう、失敗があればエラーで返す
    If failCount > 0 Then Err.Raise vbObjectError + 513, "Test_BiDiElement", "FAIL=" & failCount & " / PASS=" & passCount & failNames
End Sub

'==============================================================================
' ① WebDriverBiDiContext の検索口（document起点）
'==============================================================================
Private Sub Test01_ContextSearch(br As WebDriverBiDiContext)
    PrintSection "① WebDriverBiDiContext 検索口"

    Dim byId As PageElement: Set byId = br.getElementByID("testInput")
    AssertTrue "getElementByID: isExist", byId.isExist
    AssertTrue "getElementByID: BiDiモードであること", byId.isBiDi

    AssertEq "getElementByQuery", br.getElementByQuery("#collectionList [data-n='2']").innerText, "B"
    AssertEq "getElementByXPath", br.getElementByXPath("//ul[@id='collectionList']/li[3]").innerText, "C"
    AssertEq "getElementByXPath(\区切り補正)", br.getElementByXPath("\\ul[@id='collectionList']\li[1]").innerText, "A"

    Dim byQueryAll As Collection: Set byQueryAll = br.getElementsByQuery("#collectionList .list-item")
    AssertEq "getElementsByQueryの件数", CStr(byQueryAll.Count), "5"
    AssertEq "getElementsByQuery[5]のdata-n", byQueryAll(5).getAttribute("data-n"), "5"

    Dim byXPathAll As Collection: Set byXPathAll = br.getElementsByXPath("//ul[@id='collectionList']/li")
    AssertEq "getElementsByXPathの件数", CStr(byXPathAll.Count), "5"

    '存在しないIDは、未検出(vbNull)となること
    '※検索口は検索文字列をそのまま埋め込むため、`"`を含む値は対象外(CDPContextと同仕様)
    AssertFalse "存在しないIDは未検出となること", br.getElementByID("doesNotExist12345").isExist
    AssertTrue "存在しない要素の複数検索は空のCollection", (br.getElementsByQuery(".doesNotExist12345").Count = 0)

    br.jsEval "updateStatus('s-ce01','① 完了 " & EOk() & " | BiDi検索口', true)"
End Sub

'==============================================================================
' ② value / sendString / clearValue
'==============================================================================
Private Sub Test02_Value_ClearValue(br As WebDriverBiDiContext)
    PrintSection "② value / sendString / clearValue"

    Dim el As PageElement: Set el = br.getElementByID("testInput")

    el.value = "Hello VBA"
    AssertEq "value(Let→Get)", el.value, "Hello VBA"

    el.value = "改行" & vbLf & "や""引用符""を含む値"
    AssertEq "value(特殊文字)", el.value, "改行" & "や""引用符""を含む値"

    'sendString（input.performActions で1文字ずつ入力）
    el.sendString "Real Key Input"
    AssertEq "sendString後のvalue", el.value, "Real Key Input"

    '日本語・サロゲートペア(絵文字)を含む文字列
    el.sendString "日本語" & ChrW(&HD83D) & ChrW(&HDE00)
    AssertEq "sendString(日本語・絵文字)後のvalue", el.value, "日本語" & ChrW(&HD83D) & ChrW(&HDE00)

    el.clearValue
    AssertEq "clearValue後のvalue", el.value, ""
End Sub

'==============================================================================
' ③ innerText / innerHTML
'==============================================================================
Private Sub Test03_InnerText_InnerHTML(br As WebDriverBiDiContext)
    PrintSection "③ innerText / innerHTML"

    Dim txtEl As PageElement: Set txtEl = br.getElementByID("testTextDiv")
    AssertEq "innerText初期値", txtEl.innerText, "initial text"
    txtEl.innerText = "changed text"
    AssertEq "innerText変更後", txtEl.innerText, "changed text"

    Dim htmlEl As PageElement: Set htmlEl = br.getElementByID("testHtmlDiv")
    AssertContains "innerHTML初期値", htmlEl.innerHTML, "bold"
    htmlEl.innerHTML = "<i>italic</i>"
    AssertContains "innerHTML変更後", htmlEl.innerHTML, "italic"
End Sub

'==============================================================================
' ④ checked（Boolean引数の変換確認）
'==============================================================================
Private Sub Test04_Checked(br As WebDriverBiDiContext)
    PrintSection "④ checked"

    Dim cb As PageElement: Set cb = br.getElementByID("testCheckbox")

    cb.checked = True
    AssertTrue "checked=True設定後", cb.checked

    cb.checked = False
    AssertFalse "checked=False設定後", cb.checked
End Sub

'==============================================================================
' ⑤ selected / setSelection
'==============================================================================
Private Sub Test05_Selected_SetSelection(br As WebDriverBiDiContext)
    PrintSection "⑤ selected / setSelection"

    Dim sel As PageElement: Set sel = br.getElementByID("testSelect")

    sel.setSelection "opt2"
    AssertEq "setSelection後のvalue", sel.value, "opt2"

    sel.selected = "0"
    AssertEq "selected=0後のvalue", sel.value, "opt1"

    AssertNotEmpty "selected(Get)", CStr(sel.selected)
End Sub

'==============================================================================
' ⑥ click / SimpleClick / fireEvent
'==============================================================================
Private Sub Test06_Click_SimpleClick_FireEvent(br As WebDriverBiDiContext)
    PrintSection "⑥ click / SimpleClick / fireEvent"

    Dim btn As PageElement: Set btn = br.getElementByID("testClickBtn")

    AssertEq "初期クリック回数", btn.getAttribute("data-clickcount"), "0"

    btn.click
    AssertEq "click後のクリック回数", btn.getAttribute("data-clickcount"), "1"

    btn.SimpleClick
    AssertEq "SimpleClick後のクリック回数", btn.getAttribute("data-clickcount"), "2"

    btn.fireEvent "customtestevent"
    AssertEq "fireEvent後のcustomfired", btn.getAttribute("data-customfired"), "true"
End Sub

'==============================================================================
' ⑦ getAttribute / setAttribute（引数2個の変換確認）
'==============================================================================
Private Sub Test07_GetAttribute_SetAttribute(br As WebDriverBiDiContext)
    PrintSection "⑦ getAttribute / setAttribute"

    Dim el As PageElement: Set el = br.getElementByID("testAttrTarget")

    AssertEq "getAttribute初期値", el.getAttribute("data-foo"), "bar"

    el.setAttribute "data-foo", "baz"
    AssertEq "setAttribute後のgetAttribute", el.getAttribute("data-foo"), "baz"
End Sub

'==============================================================================
' ⑧ focus / selectText
'==============================================================================
Private Sub Test08_Focus_SelectText(br As WebDriverBiDiContext)
    PrintSection "⑧ focus / selectText"

    Dim fi As PageElement: Set fi = br.getElementByID("testFocusInput")
    fi.focus
    AssertEq "focus後のdocument.activeElement.id", CStr(br.jsEval("document.activeElement.id")), "testFocusInput"

    Dim st As PageElement: Set st = br.getElementByID("testSelectTextTarget")
    st.selectText
    AssertEq "selectText後の選択文字列", CStr(br.jsEval("window.getSelection().toString()")), "Select this whole text"
End Sub

'==============================================================================
' ⑨ submit
'==============================================================================
Private Sub Test09_Submit(br As WebDriverBiDiContext)
    PrintSection "⑨ submit"

    br.getElementByID("testSubmitInput").submit
    AssertEq "submit呼び出し後のdata-submitted", br.getElementByID("testForm").getAttribute("data-submitted"), "true"
End Sub

'==============================================================================
' ⑩ getParent / getNextSibling / getPrevSibling / getFirstChild
'==============================================================================
Private Sub Test10_Traversal_Parent_Siblings_FirstChild(br As WebDriverBiDiContext)
    PrintSection "⑩ getParent / getNextSibling / getPrevSibling / getFirstChild"

    Dim parentEl As PageElement: Set parentEl = br.getElementByID("traversalParent")
    Dim child1 As PageElement: Set child1 = parentEl.getFirstChild
    AssertEq "getFirstChildのid", child1.getAttribute("id"), "child1"
    AssertTrue "相対検索の子もBiDiモードであること", child1.isBiDi

    Dim child2 As PageElement: Set child2 = child1.getNextSibling
    AssertEq "getNextSiblingのid", child2.getAttribute("id"), "child2"

    AssertEq "getPrevSiblingのid", br.getElementByID("child3").getPrevSibling.getAttribute("id"), "child2"
    AssertEq "getParentのid", child2.getParent.getAttribute("id"), "traversalParent"
End Sub

'==============================================================================
' ⑪ getChildren / getElementsByQuery / getElementsByXPath（要素スコープ）
'==============================================================================
Private Sub Test11_GetChildren_ElementsByQuery_ElementsByXPath(br As WebDriverBiDiContext)
    PrintSection "⑪ getChildren / getElementsByQuery / getElementsByXPath"

    Dim ul As PageElement: Set ul = br.getElementByID("collectionList")

    AssertEq "getChildrenの件数", CStr(ul.getChildren.Count), "5"

    Dim byQuery As Collection: Set byQuery = ul.getElementsByQuery(".list-item")
    AssertEq "getElementsByQueryの件数", CStr(byQuery.Count), "5"
    AssertEq "getElementsByQuery[1]のdata-n", byQuery(1).getAttribute("data-n"), "1"

    AssertEq "getElementsByXPathの件数", CStr(ul.getElementsByXPath("li").Count), "5"
End Sub

'==============================================================================
' ⑫ getElementByID / getElementByQuery / getElementByXPath（要素スコープ）
'==============================================================================
Private Sub Test12_ElementByID_Query_XPath_Scoped(br As WebDriverBiDiContext)
    PrintSection "⑫ getElementByID / getElementByQuery / getElementByXPath（スコープ検索）"

    Dim ul As PageElement: Set ul = br.getElementByID("collectionList")

    AssertEq "getElementByID(スコープ)の内容", ul.getElementByID("collectionItem3").innerText, "C"
    AssertEq "getElementByQuery(スコープ)の内容", ul.getElementByQuery("[data-n='2']").innerText, "B"
    AssertEq "getElementByXPath(スコープ)の内容", ul.getElementByXPath("li[3]").innerText, "C"
End Sub

'==============================================================================
' ⑬ isExist / ifExist / onExist / onExistNot
'==============================================================================
Private Sub Test13_IsExist_IfExist_OnExist_OnExistNot(br As WebDriverBiDiContext)
    PrintSection "⑬ isExist / ifExist / onExist / onExistNot"

    Dim dyn As PageElement: Set dyn = br.getElementByID("dynamicElement")
    AssertFalse "追加前のisExist", dyn.isExist

    Dim errBefore As Long
    On Error Resume Next
    Err.Clear
    dyn.ifExist.focus
    errBefore = Err.Number
    On Error GoTo 0
    AssertTrue "ifExistチェーン（未存在）でエラーが起きないこと", (errBefore = 0)

    'HTML側で1秒後に要素を追加 → onExist（親なし＝document起点の再検索）で検知
    br.getElementByID("btnAddDynamic").click
    Dim dyn2 As PageElement: Set dyn2 = br.getElementByID("dynamicElement").onExist(timeOutInSeconds:=5)
    AssertTrue "追加後のonExist成功", Not (dyn2 Is Nothing)
    If Not (dyn2 Is Nothing) Then AssertTrue "onExist後のisExist", dyn2.isExist

    '相対検索(引数あり)のonExist：親要素起点 + SearchArgsの再利用
    Dim scoped As PageElement: Set scoped = br.getElementByQuery("body").getElementByID("dynamicElement").onExist(timeOutInSeconds:=5)
    AssertTrue "相対検索のonExist成功", Not (scoped Is Nothing)

    br.getElementByID("btnRemoveDynamic").click
    If Not (dyn2 Is Nothing) Then AssertTrue "削除後のonExistNot成功", dyn2.onExistNot(timeOutInSeconds:=5)

    Dim ghost As PageElement
    Set ghost = br.getElementByID("doesNotExist12345").onExist(timeOutInSeconds:=1, raiseTimeoutError:=False)
    AssertTrue "存在しない要素のonExistタイムアウトでNothingが返ること", (ghost Is Nothing)
End Sub

'==============================================================================
' ⑭ iframe（BiDiでは getIFrame 非対応 → getTab で browsingContext として取得）
'==============================================================================
Private Sub Test14_IFrame_GetTab(br As WebDriverBiDiContext)
    PrintSection "⑭ iframe（getTab）"

    CDPHelpers.Sleep 0.5

    '1. getIFrame は、BiDiモードでは代替手段を案内するエラーで止まること
    Dim errNum As Long, unusedFrame As PageElement
    On Error Resume Next
    Err.Clear
    Set unusedFrame = br.getElementByID("testIFrame").getIFrame
    errNum = Err.Number
    On Error GoTo 0
    AssertTrue "getIFrameがBiDiモードではエラーになること", (errNum <> 0)

    '2. iframe は独立した browsingContext のため、getTab(maxDepth:=1) で取得し、その中で検索できること
    '※srcdoc の iframe なので、URLは about:srcdoc になる
    Dim iframeCtx As WebDriverBiDiContext
    Set iframeCtx = br.ThisWebDriverBiDiMode.getTab("about:srcdoc", maxDepth:=1, doRetrySecond:=3)
    AssertTrue "getTab(maxDepth:=1)でiframeのcontextが取れること", Not (iframeCtx Is Nothing)
    If Not (iframeCtx Is Nothing) Then AssertEq "iframe内要素のinnerText", iframeCtx.getElementByID("iframeTarget").innerText, "iframe content"
End Sub

'==============================================================================
' ⑮ 診断プロパティ / 実行オプション
'==============================================================================
Private Sub Test15_Diagnostics_And_Options(br As WebDriverBiDiContext)
    PrintSection "⑮ 診断プロパティ / 実行オプション"

    Dim el As PageElement: Set el = br.getElementByID("testInput")

    AssertContains "UseSearchJSに検索構文が含まれること", el.UseSearchJS, "getElementById"
    AssertTrue "varResultがvbStringであること", (el.varResult = vbString)
    AssertNotEmpty "CurrentObjectId(script.handle)が取得できていること", el.CurrentObjectId

    el.ExposeDevTools "__vbaTestExposed"
    AssertEq "ExposeDevTools後にwindowへ公開されていること", CStr(br.jsEval("typeof window.__vbaTestExposed !== 'undefined' ? 'yes' : 'no'")), "yes"

    'BiDiモードでは`script.LocalValue`形式(type付き)で引数を渡す
    Dim args(1) As Dictionary
    Set args(0) = New Dictionary: args(0).Add "type", "number": args(0).Add "value", 40
    Set args(1) = New Dictionary: args(1).Add "type", "number": args(1).Add "value", 2
    AssertEq "数値引数の受け渡し", CStr(el.jsEval("function(a, b){ return a + b }", args)), "42"

    Dim nullArg(0) As Dictionary
    Set nullArg(0) = New Dictionary: nullArg(0).Add "type", "null"
    AssertEq "Null引数の受け渡し", CStr(el.jsEval("function(v){ return v === null ? 'null' : 'x' }", nullArg)), "null"

    el.SetOptionRunAsync = True
    Dim asyncResult As Variant
    asyncResult = el.jsEval("function(){ return 1 }")
    el.SetOptionRunAsync = False
    AssertTrue "SetOptionRunAsync=True時、戻り値が非同期コマンドIDであること", (IsNumeric(asyncResult) And CLng(asyncResult) > 0)

    el.SetOptionStopException = True
    Dim exNumber As Long
    On Error Resume Next
    Err.Clear
    el.jsEval "function(){ throw new Error('BiDiElement-test'); }"
    exNumber = Err.Number
    On Error GoTo 0
    el.SetOptionStopException = False
    AssertTrue "SetOptionStopException=True時、JS例外がVBAエラーになること", (exNumber <> 0)

    Dim errUserGesture As Long
    On Error Resume Next
    Err.Clear
    el.SetOptionUserGesture = True
    el.jsEval "function(){ return 1 }"
    el.SetOptionUserGesture = False
    errUserGesture = Err.Number
    On Error GoTo 0
    AssertTrue "SetOptionUserGestureの設定・解除でエラーが起きないこと", (errUserGesture = 0)
End Sub

'==============================================================================
' ⑯ sendClick / sendKey（input.performActions）
'==============================================================================
Private Sub Test16_SendClick_SendKey(br As WebDriverBiDiContext)
    PrintSection "⑯ sendClick / sendKey"

    Dim scBtn As PageElement: Set scBtn = br.getElementByID("testSendClickBtn")
    AssertEq "sendClick前のクリック回数", scBtn.getAttribute("data-clickcount"), "0"
    scBtn.sendClick
    AssertEq "sendClick後のクリック回数", scBtn.getAttribute("data-clickcount"), "1"

    Dim ki As PageElement: Set ki = br.getElementByID("testKeyInput")
    ki.sendKey keyEnter
    AssertEq "sendKey(keyEnter)後のkeyCode", ki.getAttribute("data-lastkeycode"), "13"

    ki.sendKey keyBackspace
    AssertEq "sendKey(keyBackspace)後のkeyCode", ki.getAttribute("data-lastkeycode"), "8"
End Sub

'==============================================================================
' ⑰ sendHover / ホバーで出現するボタン（input.performActions）
'==============================================================================
Private Sub Test17_SendHover_HoverReveal(br As WebDriverBiDiContext)
    PrintSection "⑰ sendHover / ホバーで出現するボタン"

    Dim hv As PageElement: Set hv = br.getElementByID("testHoverTarget")
    AssertEq "sendHover前のhovered状態", hv.getAttribute("data-hovered"), "false"
    hv.sendHover
    AssertEq "sendHover後のhovered状態", hv.getAttribute("data-hovered"), "true"
    AssertEq "sendHover後のhover回数", hv.getAttribute("data-hovercount"), "1"

    Dim container As PageElement: Set container = br.getElementByID("hoverRevealContainer")
    Dim revealBtn As PageElement: Set revealBtn = br.getElementByID("hoverRevealBtn")
    AssertEq "ホバー前のクリック回数", revealBtn.getAttribute("data-clickcount"), "0"
    container.sendHover
    Dim waitResult As String: waitResult = WaitUntilClickable(revealBtn)
    AssertTrue "ホバー後、ボタンがクリック可能になること(DOM安定待機)" & IIf(waitResult = "OK", "", " / " & waitResult), (waitResult = "OK")
    revealBtn.sendClick
    AssertEq "ホバー→sendClick後のクリック回数", revealBtn.getAttribute("data-clickcount"), "1"
End Sub

'==============================================================================
' ⑱ SetFileInputFiles（input.setFiles）
'==============================================================================
Private Sub Test18_SetFileInputFiles(br As WebDriverBiDiContext)
    PrintSection "⑱ SetFileInputFiles"

    Dim files As New Collection
    files.Add WorkspacePath & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html"

    Dim fileInput As PageElement: Set fileInput = br.getElementByID("testFileInput")
    fileInput.SetFileInputFiles files

    AssertEq "SetFileInputFiles後のfiles[0].name", CStr(fileInput.jsEval("function(){ return this.files.length > 0 ? this.files[0].name : '' }")), "CDPElementTest.html"
End Sub

'==============================================================================
' ⑲ GetShadowRoot（open / closed、includeShadowTree:="all"）／ GetShadowRoots は BiDi 非対応
'==============================================================================
Private Sub Test19_ShadowRoot_Open_Closed(br As WebDriverBiDiContext)
    PrintSection "⑲ GetShadowRoot（open / closed）"

    Dim rootOpen As PageElement: Set rootOpen = br.getElementByID("shadowHostOpen").GetShadowRoot()
    AssertTrue "GetShadowRoot(open)が取得できること", Not (rootOpen Is Nothing)
    If Not (rootOpen Is Nothing) Then AssertEq "shadow(open)内要素の内容", rootOpen.getElementByQuery(".shadow-content").innerText, "Shadow content (open)"

    Dim rootClosed As PageElement: Set rootClosed = br.getElementByID("shadowHostClosed").GetShadowRoot()
    AssertTrue "GetShadowRoot(closed)が取得できること", Not (rootClosed Is Nothing)
    If Not (rootClosed Is Nothing) Then AssertEq "shadow(closed)内要素の内容", rootClosed.getElementByQuery(".shadow-content").innerText, "Shadow content (closed)"

    'Index 1以上は、BiDiでは対象外(Nothing)となること
    AssertTrue "GetShadowRoot(Index:=1)はNothing", (br.getElementByID("shadowHostOpen").GetShadowRoot(1) Is Nothing)

    'GetShadowRoots は、BiDiモードでは代替手段を案内するエラーで止まること
    Dim errNum As Long, unusedRoots As Collection
    On Error Resume Next
    Err.Clear
    Set unusedRoots = br.getElementByID("shadowHostClosed").GetShadowRoots()
    errNum = Err.Number
    On Error GoTo 0
    AssertTrue "GetShadowRootsがBiDiモードではエラーになること", (errNum <> 0)

    'Shadow-rootがない要素では、Nothingが返ること　※<input>等はブラウザ組み込みのShadow-rootを持つため、divで確認
    AssertTrue "Shadow-rootがない要素ではNothing", (br.getElementByID("testTextDiv").GetShadowRoot() Is Nothing)
End Sub

'==============================================================================
' ヘルパー
'==============================================================================
Private Sub RecordFail(Name As String)
    failCount = failCount + 1
    failNames = failNames & vbLf & "  - " & Name
End Sub

'==============================================================================
' DOMが落ち着いた(要素がクリック可能になった)かの待機
'  ・実行中のアニメーション/トランジションがない
'  ・visibility:visible かつ opacity:1
'  ・要素の中心座標の最前面が、その要素自身(Shadow DOM内は、そのShadow-root側で判定)
'  上記を満たした時点で即 "OK" を返す。上限時間内に満たさない場合は、失敗時の状態(診断情報)を返す(固定秒数の待機ではない)
'  診断情報：ページの表示/フォーカス状態、親要素の:hover状態、ボタンのスタイル、各中心座標の最前面要素 等
'  【メモ】失敗時の診断で visibilityState=hidden かつ runningAnimations>0 のまま opacity=0 の場合は、ホバーは届いているが、
'          ブラウザの画面が他のウィンドウの裏に隠れていて、アニメーションの描画が止まっている状態(環境要因・テスト対象の不具合ではない)
'          → 起動時の追加引数に `--disable-backgrounding-occluded-windows` `--disable-renderer-backgrounding` を付けると安定する
'            (2026/10/16 xlflow経由の5回連続実行で、引数なし：15回中9回失敗 → 引数あり：15回すべて合格)
'==============================================================================
Private Function WaitUntilClickable(el As PageElement, Optional timeoutMs As Long = 5000) As String
    Dim waitJS As String
    waitJS = "function(){" & _
             "  const el = this, root = el.getRootNode(), parent = el.parentElement, deadline = performance.now() + " & timeoutMs & ";" & _
             "  const desc = n => !n ? 'null' : (n.tagName || n.nodeName) + (n.id ? '#' + n.id : '') + (n.className && typeof n.className === 'string' ? '.' + n.className.trim().replace(/\s+/g, '.') : '');" & _
             "  const center = n => { const r = n.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; };" & _
             "  const hitAt = n => { const c = center(n); return root.elementFromPoint(c[0], c[1]); };" & _
             "  const running = () => el.getAnimations().filter(a => a.playState === 'running').length;" & _
             "  const diag = () => { const s = getComputedStyle(el), r = el.getBoundingClientRect(); return [" & _
             "    'visibilityState=' + document.visibilityState, 'hasFocus=' + document.hasFocus()," & _
             "    'viewport=' + window.innerWidth + 'x' + window.innerHeight," & _
             "    'parentHover=' + (parent ? parent.matches(':hover') : 'n/a'), 'parentClass=' + (parent ? parent.className : 'n/a')," & _
             "    'btnVisibility=' + s.visibility, 'btnOpacity=' + s.opacity, 'runningAnimations=' + running()," & _
             "    'btnRect=' + [r.left, r.top, r.width, r.height].map(Math.round).join(',')," & _
             "    'hitAtBtn=' + desc(hitAt(el)), 'hitAtParent=' + (parent ? desc(hitAt(parent)) : 'n/a')" & _
             "  ].join(' | '); };" & _
             "  return new Promise(resolve => {" & _
             "    const check = () => {" & _
             "      const s = getComputedStyle(el), hit = hitAt(el);" & _
             "      const ok = running() === 0 && s.visibility === 'visible' && Number(s.opacity) >= 1 && !!hit && (hit === el || el.contains(hit));" & _
             "      if (ok) { resolve('OK'); } else if (performance.now() > deadline) { resolve(diag()); } else { setTimeout(check, 16); }" & _
             "    };" & _
             "    check();" & _
             "  });" & _
             "}"

    WaitUntilClickable = CStr(el.jsEval(waitJS, IFEXCEPTION:="exception while waiting", awaitPromise:=True))
End Function

Private Sub AssertEq(Name As String, actual As Variant, expected As Variant)
    If CStr(actual) = CStr(expected) Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → """ & CStr(actual) & """"
    Else
        RecordFail Name & " 期待:""" & CStr(expected) & """ 実際:""" & CStr(actual) & """"
        Debug.Print "  FAIL | " & Name & " 期待:""" & CStr(expected) & """ 実際:""" & CStr(actual) & """"
    End If
End Sub

Private Sub AssertNotEmpty(Name As String, actual As String)
    If Len(actual) > 0 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → NOT EMPTY"
    Else
        RecordFail Name
        Debug.Print "  FAIL | " & Name & " が空です"
    End If
End Sub

Private Sub AssertContains(Name As String, actual As String, substring As String)
    If InStr(actual, substring) > 0 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → """ & substring & """ を含む"
    Else
        RecordFail Name
        Debug.Print "  FAIL | " & Name & " に """ & substring & """ が含まれません。実際:""" & actual & """"
    End If
End Sub

Private Sub AssertTrue(Name As String, actual As Boolean)
    If actual Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → True"
    Else
        RecordFail Name
        Debug.Print "  FAIL | " & Name & " が True になりませんでした"
    End If
End Sub

Private Sub AssertFalse(Name As String, actual As Boolean)
    If Not actual Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → False"
    Else
        RecordFail Name
        Debug.Print "  FAIL | " & Name & " が False になりませんでした"
    End If
End Sub

Private Sub PrintHeader(msg As String)
    Debug.Print ""
    Debug.Print String(70, "=")
    Debug.Print "  " & msg
    Debug.Print String(70, "=")
End Sub

Private Sub PrintSection(msg As String)
    Debug.Print ""
    Debug.Print "  ── " & msg & " ──"
End Sub
