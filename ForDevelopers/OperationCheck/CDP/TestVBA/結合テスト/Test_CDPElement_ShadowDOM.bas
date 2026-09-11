Attribute VB_Name = "Test_CDPElement_ShadowDOM"
'==============================================================================
' CDPElement.cls 動作確認（ShadowDOM-Open/Close内操作編・Test_CDPElement.basの亜種）
' ・流れ：
'     TOP層 → getElementByXPathでShadow-Root手前の要素(host)を発見
'           → GetShadowRootでShadow-Root(open/close)内に侵入（通り抜けフープ）
'           → 侵入先のShadow-Root内で、TOP層版(Test_CDPElement.bas)と同じテストケースを実施
' ・ShadowDOM圏内では document.evaluate によるXPath検索が機能しないため、
'   getElementByXPath / getElementsByXPath 関連のテストケースはこのファイルでは除外しています。
' ・open/closedどちらのShadow-Rootも、GetShadowRootが返すCDPElementの扱いは同一（CDPの
'   DOM.describeNode(pierce:=True)経由で侵入するため、closed特有の分岐は不要）。
' ・HTML: ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html
' ・実行前に CDP が起動可、WORKSPACE_PATH をルートに設定してから RunAll_CDPElement_ShadowDOM_Tests を実行
'==============================================================================
Option Explicit

Private passCount As Long
Private failCount As Long

'ワークスペースパス（StarterWebScrapingKit ルート）
Private Const WORKSPACE_PATH As String = ""

Private Function EOk() As String
    EOk = WorksheetFunction.Unichar(9989)
End Function

'==============================================================================
' エントリ
'==============================================================================
Public Sub RunAll_CDPElement_ShadowDOM_Tests()
    Dim br As CDPContext: Set br = ShSetting01_StartBrowser.StartCDPModeContext

    br.navigate "file:///" & Replace(WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html", "\", "/")
    br.wait

    passCount = 0: failCount = 0

    PrintHeader "CDPElement ShadowDOM内操作 検証テスト 開始"

    'TOP層：Shadow-Root手前の要素(host)をXPathで発見 → GetShadowRootで侵入
    Dim hostOpen As CDPElement: Set hostOpen = br.getElementByXPath("//*[@id='shadowHostOpen']")
    Dim rootOpen As CDPElement: Set rootOpen = hostOpen.GetShadowRoot()
    AssertTrue "GetShadowRoot(open)への侵入成功", Not (rootOpen Is Nothing)

    Dim hostClosed As CDPElement: Set hostClosed = br.getElementByXPath("//*[@id='shadowHostClosed']")
    Dim rootClosed As CDPElement: Set rootClosed = hostClosed.GetShadowRoot()
    AssertTrue "GetShadowRoot(closed)への侵入成功", Not (rootClosed Is Nothing)

    Test01_Value_SendString_ClearValue br, rootOpen, rootClosed
    Test02_InnerText_InnerHTML br, rootOpen, rootClosed
    Test03_Checked br, rootOpen, rootClosed
    Test04_Selected_SetSelection br, rootOpen, rootClosed
    Test05_Click_SimpleClick_FireEvent br, rootOpen, rootClosed
    Test06_SendClick_SendKey br, rootOpen, rootClosed
    Test07_GetAttribute_SetAttribute br, rootOpen, rootClosed
    Test08_Focus_SelectText br, rootOpen, rootClosed
    Test09_Submit br, rootOpen, rootClosed
    Test10_Traversal_Parent_Siblings_FirstChild br, rootOpen, rootClosed
    Test11_GetChildren_ElementsByQuery br, rootOpen, rootClosed
    Test12_ElementByID_Query_Scoped br, rootOpen, rootClosed
    Test13_IsExist_IfExist_OnExist_OnExistNot br, rootOpen, rootClosed
    Test14_SetFileInputFiles br, rootOpen, rootClosed
    Test15_Diagnostics_And_Options br, rootOpen, rootClosed
    Test16_SendHover br, rootOpen, rootClosed
    Test17_HoverReveal_Click br, rootOpen, rootClosed

    PrintHeader "テスト結果: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)

    br.jsEval "updateStatus('s-sce-summary','PASS=" & passCount & " FAIL=" & failCount & " " & EOk() & "', true)", StopApiError:=True

    br.ThisCDPBrowser.quit
End Sub

'==============================================================================
' SD① value / sendString / clearValue
'==============================================================================
Private Sub Test01_Value_SendString_ClearValue(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD① value / sendString / clearValue"

    RunValueTest rootOpen, "open"
    RunValueTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce01','SD① 完了 " & EOk() & " | value / sendString / clearValue（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunValueTest(root As CDPElement, modeLabel As String)
    Dim el As CDPElement: Set el = root.getElementByID("sInput")

    el.value = "Hello VBA"
    AssertEq "[" & modeLabel & "] value(Let→Get)", el.value, "Hello VBA"

    el.sendString "Real Key Input"
    AssertEq "[" & modeLabel & "] sendString後のvalue", el.value, "Real Key Input"

    el.clearValue
    AssertEq "[" & modeLabel & "] clearValue後のvalue", el.value, ""
End Sub

'==============================================================================
' SD② innerText / innerHTML
'==============================================================================
Private Sub Test02_InnerText_InnerHTML(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD② innerText / innerHTML"

    RunInnerTextHtmlTest rootOpen, "open"
    RunInnerTextHtmlTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce02','SD② 完了 " & EOk() & " | innerText / innerHTML（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunInnerTextHtmlTest(root As CDPElement, modeLabel As String)
    Dim txtEl As CDPElement: Set txtEl = root.getElementByID("sTextDiv")
    AssertEq "[" & modeLabel & "] innerText初期値", txtEl.innerText, "initial text"
    txtEl.innerText = "changed text"
    AssertEq "[" & modeLabel & "] innerText変更後", txtEl.innerText, "changed text"

    Dim htmlEl As CDPElement: Set htmlEl = root.getElementByID("sHtmlDiv")
    AssertContains "[" & modeLabel & "] innerHTML初期値", htmlEl.innerHTML, "bold"
    htmlEl.innerHTML = "<i>italic</i>"
    AssertContains "[" & modeLabel & "] innerHTML変更後", htmlEl.innerHTML, "italic"
End Sub

'==============================================================================
' SD③ checked（チェックボックス）
'==============================================================================
Private Sub Test03_Checked(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD③ checked"

    RunCheckedTest rootOpen, "open"
    RunCheckedTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce03','SD③ 完了 " & EOk() & " | checked（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunCheckedTest(root As CDPElement, modeLabel As String)
    Dim cb As CDPElement: Set cb = root.getElementByID("sCheckbox")

    cb.checked = True
    AssertTrue "[" & modeLabel & "] checked=True設定後", cb.checked

    cb.checked = False
    AssertFalse "[" & modeLabel & "] checked=False設定後", cb.checked
End Sub

'==============================================================================
' SD④ selected / setSelection（セレクトボックス）
'==============================================================================
Private Sub Test04_Selected_SetSelection(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD④ selected / setSelection"

    RunSelectedTest rootOpen, "open"
    RunSelectedTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce04','SD④ 完了 " & EOk() & " | selected / setSelection（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunSelectedTest(root As CDPElement, modeLabel As String)
    Dim sel As CDPElement: Set sel = root.getElementByID("sSelect")

    '`setSelection`はvalue属性で指定して選択する
    sel.setSelection "opt2"
    AssertEq "[" & modeLabel & "] setSelection後のvalue", sel.value, "opt2"

    '`selected`Letは、selectedIndexで指定して選択する（0始まり）
    sel.selected = "0"
    AssertEq "[" & modeLabel & "] selected=0後のvalue", sel.value, "opt1"

    '`selected`Getは、選択中option要素のobjectIdを返す（文字列取得できていることの確認）
    AssertNotEmpty "[" & modeLabel & "] selected（選択中option要素のobjectId）", CStr(sel.selected)
End Sub

'==============================================================================
' SD⑤ click / SimpleClick / fireEvent
'==============================================================================
Private Sub Test05_Click_SimpleClick_FireEvent(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑤ click / SimpleClick / fireEvent"

    RunClickTest rootOpen, "open"
    RunClickTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce05','SD⑤ 完了 " & EOk() & " | click / SimpleClick / fireEvent（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunClickTest(root As CDPElement, modeLabel As String)
    Dim btn As CDPElement: Set btn = root.getElementByID("sClickBtn")

    AssertEq "[" & modeLabel & "] 初期クリック回数", btn.getAttribute("data-clickcount"), "0"

    btn.click
    AssertEq "[" & modeLabel & "] click後のクリック回数", btn.getAttribute("data-clickcount"), "1"

    btn.SimpleClick
    AssertEq "[" & modeLabel & "] SimpleClick後のクリック回数", btn.getAttribute("data-clickcount"), "2"

    btn.fireEvent "customtestevent"
    AssertEq "[" & modeLabel & "] fireEvent後のcustomfired", btn.getAttribute("data-customfired"), "true"
End Sub

'==============================================================================
' SD⑥ sendClick / sendKey
'==============================================================================
Private Sub Test06_SendClick_SendKey(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑥ sendClick / sendKey"

    RunSendClickKeyTest rootOpen, "open"
    RunSendClickKeyTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce06','SD⑥ 完了 " & EOk() & " | sendClick / sendKey（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunSendClickKeyTest(root As CDPElement, modeLabel As String)
    Dim scBtn As CDPElement: Set scBtn = root.getElementByID("sSendClickBtn")
    AssertEq "[" & modeLabel & "] sendClick前のクリック回数", scBtn.getAttribute("data-clickcount"), "0"
    scBtn.sendClick
    AssertEq "[" & modeLabel & "] sendClick後のクリック回数", scBtn.getAttribute("data-clickcount"), "1"

    Dim ki As CDPElement: Set ki = root.getElementByID("sKeyInput")
    ki.sendKey keyEnter
    AssertEq "[" & modeLabel & "] sendKey(keyEnter)後のkeyCode", ki.getAttribute("data-lastkeycode"), "13"

    ki.sendKey keyBackspace
    AssertEq "[" & modeLabel & "] sendKey(keyBackspace)後のkeyCode", ki.getAttribute("data-lastkeycode"), "8"
End Sub

'==============================================================================
' SD⑦ getAttribute / setAttribute
'==============================================================================
Private Sub Test07_GetAttribute_SetAttribute(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑦ getAttribute / setAttribute"

    RunAttributeTest rootOpen, "open"
    RunAttributeTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce07','SD⑦ 完了 " & EOk() & " | getAttribute / setAttribute（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunAttributeTest(root As CDPElement, modeLabel As String)
    Dim el As CDPElement: Set el = root.getElementByID("sAttrTarget")

    AssertEq "[" & modeLabel & "] getAttribute初期値", el.getAttribute("data-foo"), "bar"

    el.setAttribute "data-foo", "baz"
    AssertEq "[" & modeLabel & "] setAttribute後のgetAttribute", el.getAttribute("data-foo"), "baz"
End Sub

'==============================================================================
' SD⑧ focus / selectText
'==============================================================================
Private Sub Test08_Focus_SelectText(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑧ focus / selectText"

    RunFocusSelectTextTest rootOpen, "open"
    RunFocusSelectTextTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce08','SD⑧ 完了 " & EOk() & " | focus / selectText（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunFocusSelectTextTest(root As CDPElement, modeLabel As String)
    Dim fi As CDPElement: Set fi = root.getElementByID("sFocusInput")
    fi.focus

    '`document.activeElement`はShadow-Root内までは追えない（hostが返る）ため、
    'Shadow-Root自身の`activeElement`（`this`＝Shadow-Root）で確認する
    Dim activeId As Variant
    activeId = root.jsEval("function(){ return this.activeElement ? this.activeElement.id : '' }", StopApiError:=False)
    AssertEq "[" & modeLabel & "] focus後のShadowRoot.activeElement.id", CStr(activeId), "sFocusInput"

    Dim st As CDPElement: Set st = root.getElementByID("sSelectTextTarget")
    st.selectText
    Dim selectedText As Variant
    selectedText = root.jsEval("function(){ return window.getSelection().toString() }", StopApiError:=False)
    AssertEq "[" & modeLabel & "] selectText後の選択文字列", CStr(selectedText), "Select this whole text"
End Sub

'==============================================================================
' SD⑨ submit
'==============================================================================
Private Sub Test09_Submit(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑨ submit"

    RunSubmitTest rootOpen, "open"
    RunSubmitTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce09','SD⑨ 完了 " & EOk() & " | submit（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunSubmitTest(root As CDPElement, modeLabel As String)
    '`submit`は`this.form.submit()`を呼ぶため、通常はページ遷移が発生する。
    'テストページを壊さず検証するため、HTML側で`<form>`インスタンスの`submit`メソッドを
    'JSでオーバーライドし、呼び出しがあったことだけを記録するようにしている（TOP層版と同じ仕掛け）。
    Dim el As CDPElement: Set el = root.getElementByID("sSubmitInput")
    el.submit

    Dim frm As CDPElement: Set frm = root.getElementByID("sForm")
    AssertEq "[" & modeLabel & "] submit呼び出し後のdata-submitted", frm.getAttribute("data-submitted"), "true"
End Sub

'==============================================================================
' SD⑩ DOM階層: getParent / getNextSibling / getPrevSibling / getFirstChild
'==============================================================================
Private Sub Test10_Traversal_Parent_Siblings_FirstChild(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑩ getParent / getNextSibling / getPrevSibling / getFirstChild"

    RunTraversalTest rootOpen, "open"
    RunTraversalTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce10','SD⑩ 完了 " & EOk() & " | getParent / getNextSibling / getPrevSibling / getFirstChild（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunTraversalTest(root As CDPElement, modeLabel As String)
    Dim parentEl As CDPElement: Set parentEl = root.getElementByID("sTraversalParent")
    Dim child1 As CDPElement: Set child1 = parentEl.getFirstChild
    AssertEq "[" & modeLabel & "] getFirstChildのid", child1.getAttribute("id"), "sChild1"

    Dim child2 As CDPElement: Set child2 = child1.getNextSibling
    AssertEq "[" & modeLabel & "] getNextSiblingのid", child2.getAttribute("id"), "sChild2"

    Dim child3 As CDPElement: Set child3 = root.getElementByID("sChild3")
    Dim backToChild2 As CDPElement: Set backToChild2 = child3.getPrevSibling
    AssertEq "[" & modeLabel & "] getPrevSiblingのid", backToChild2.getAttribute("id"), "sChild2"

    Dim parentBack As CDPElement: Set parentBack = child2.getParent
    AssertEq "[" & modeLabel & "] getParentのid", parentBack.getAttribute("id"), "sTraversalParent"
End Sub

'==============================================================================
' SD⑪ getChildren / getElementsByQuery（要素スコープ、複数取得）
' ※ShadowDOM圏内ではdocument.evaluateが機能しないため、getElementsByXPathは対象外
'==============================================================================
Private Sub Test11_GetChildren_ElementsByQuery(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑪ getChildren / getElementsByQuery（XPathは対象外）"

    RunCollectionTest rootOpen, "open"
    RunCollectionTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce11','SD⑪ 完了 " & EOk() & " | getChildren / getElementsByQuery（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunCollectionTest(root As CDPElement, modeLabel As String)
    Dim ul As CDPElement: Set ul = root.getElementByID("sCollectionList")

    Dim children As Collection: Set children = ul.getChildren
    AssertEq "[" & modeLabel & "] getChildrenの件数", CStr(children.Count), "5"

    Dim byQuery As Collection: Set byQuery = ul.getElementsByQuery(".list-item")
    AssertEq "[" & modeLabel & "] getElementsByQueryの件数", CStr(byQuery.Count), "5"
    AssertEq "[" & modeLabel & "] getElementsByQuery[1]のdata-n", byQuery(1).getAttribute("data-n"), "1"
End Sub

'==============================================================================
' SD⑫ getElementByID / getElementByQuery（要素スコープ、単一取得）
' ※ShadowDOM圏内ではdocument.evaluateが機能しないため、getElementByXPathは対象外
'==============================================================================
Private Sub Test12_ElementByID_Query_Scoped(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑫ getElementByID / getElementByQuery（スコープ検索、XPathは対象外）"

    RunScopedSearchTest rootOpen, "open"
    RunScopedSearchTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce12','SD⑫ 完了 " & EOk() & " | getElementByID / getElementByQuery（スコープ・open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunScopedSearchTest(root As CDPElement, modeLabel As String)
    Dim ul As CDPElement: Set ul = root.getElementByID("sCollectionList")

    Dim byId As CDPElement: Set byId = ul.getElementByID("sCollectionItem3")
    AssertEq "[" & modeLabel & "] getElementByID(スコープ)の内容", byId.innerText, "C"

    Dim byQuery As CDPElement: Set byQuery = ul.getElementByQuery("[data-n='2']")
    AssertEq "[" & modeLabel & "] getElementByQuery(スコープ)の内容", byQuery.innerText, "B"
End Sub

'==============================================================================
' SD⑬ isExist / ifExist / onExist / onExistNot
'==============================================================================
Private Sub Test13_IsExist_IfExist_OnExist_OnExistNot(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑬ isExist / ifExist / onExist / onExistNot"

    RunExistenceTest rootOpen, "open"
    RunExistenceTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce13','SD⑬ 完了 " & EOk() & " | isExist / ifExist / onExist / onExistNot（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunExistenceTest(root As CDPElement, modeLabel As String)
    '1. 追加前は存在しないこと
    Dim dyn As CDPElement: Set dyn = root.getElementByID("sDynamicElement")
    AssertFalse "[" & modeLabel & "] 追加前のisExist", dyn.isExist

    '2. ifExistチェーンは、存在しない要素へのメソッド呼び出しをエラーなく無視すること
    Dim errBefore As Long
    On Error Resume Next
    Err.Clear
    dyn.ifExist.focus
    errBefore = Err.Number
    On Error GoTo 0
    AssertTrue "[" & modeLabel & "] ifExistチェーン（未存在）でエラーが起きないこと", (errBefore = 0)

    '3. 追加ボタンを押す（HTML側で1秒後に要素を追加）→ onExistでポーリング検知できること
    root.getElementByID("sBtnAddDynamic").click
    Dim dyn2 As CDPElement: Set dyn2 = root.getElementByID("sDynamicElement")
    Set dyn2 = dyn2.onExist(timeOutInSeconds:=5)
    AssertTrue "[" & modeLabel & "] 追加後のonExist成功", Not (dyn2 Is Nothing)
    If Not (dyn2 Is Nothing) Then AssertTrue "[" & modeLabel & "] onExist後のisExist", dyn2.isExist

    '4. 削除ボタンを押す（HTML側で1秒後に要素を削除）→ onExistNotで消滅検知できること
    root.getElementByID("sBtnRemoveDynamic").click
    AssertTrue "[" & modeLabel & "] 削除後のonExistNot成功", dyn2.onExistNot(timeOutInSeconds:=5)

    '5. 存在し得ないID指定時、onExist(raiseTimeoutError:=False)がタイムアウトでNothingを返すこと
    Dim ghost As CDPElement
    Set ghost = root.getElementByID("doesNotExist12345").onExist(timeOutInSeconds:=1, raiseTimeoutError:=False)
    AssertTrue "[" & modeLabel & "] 存在しない要素のonExistタイムアウトでNothingが返ること", (ghost Is Nothing)
End Sub

'==============================================================================
' SD⑭ SetFileInputFiles
'==============================================================================
Private Sub Test14_SetFileInputFiles(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑭ SetFileInputFiles"

    RunFileInputTest rootOpen, "open"
    RunFileInputTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce14','SD⑭ 完了 " & EOk() & " | SetFileInputFiles（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunFileInputTest(root As CDPElement, modeLabel As String)
    Dim dummyFile As String
    dummyFile = WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html"

    Dim files As New Collection
    files.Add dummyFile

    Dim fileInput As CDPElement: Set fileInput = root.getElementByID("sFileInput")
    fileInput.SetFileInputFiles files

    Dim FileName As Variant
    FileName = fileInput.jsEval("function(){ return this.files.length > 0 ? this.files[0].name : '' }")
    AssertEq "[" & modeLabel & "] SetFileInputFiles後のfiles[0].name", CStr(FileName), "CDPElementTest.html"
End Sub

'==============================================================================
' SD⑮ 診断プロパティ（UseSearchJS / varResult / CurrentObjectId / ExposeDevTools）
'     と実行オプション（SetOptionStopException / SetOptionRunAsyncCDP / SetOptionUserGesture）
'==============================================================================
Private Sub Test15_Diagnostics_And_Options(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑮ 診断プロパティ / 実行オプション"

    RunDiagnosticsTest rootOpen, "open"
    RunDiagnosticsTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce15','SD⑮ 完了 " & EOk() & " | 診断プロパティ / 実行オプション（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunDiagnosticsTest(root As CDPElement, modeLabel As String)
    Dim el As CDPElement: Set el = root.getElementByID("sInput")

    'TOP層版と異なり、Shadow-Root圏内のスコープ検索は`document.getElementById`ではなく
    '`querySelector` + `CSS.escape`で行われる（CDPElement.getElementByID内部の仕様）
    AssertContains "[" & modeLabel & "] UseSearchJSがquerySelectorベースであること", el.UseSearchJS, "querySelector"
    AssertTrue "[" & modeLabel & "] varResultがvbStringであること", (el.varResult = vbString)
    AssertNotEmpty "[" & modeLabel & "] CurrentObjectIdが取得できていること", el.CurrentObjectId

    el.ExposeDevTools "__vbaShadowTestExposed_" & modeLabel
    Dim exposedCheck As Variant
    exposedCheck = root.jsEval("function(){ return typeof window['__vbaShadowTestExposed_" & modeLabel & "'] !== 'undefined' ? 'yes' : 'no' }", StopApiError:=False)
    AssertEq "[" & modeLabel & "] ExposeDevTools後にwindowへ公開されていること", CStr(exposedCheck), "yes"

    'SetOptionRunAsyncCDP: Trueにすると、jsEvalの戻り値が結果値ではなく非同期コマンドIDになる
    el.SetOptionRunAsyncCDP = True
    Dim asyncResult As Variant
    asyncResult = el.jsEval("function(){ return 1 }")
    AssertTrue "[" & modeLabel & "] SetOptionRunAsyncCDP=True時、戻り値が数値の非同期コマンドIDであること", (IsNumeric(asyncResult) And CLng(asyncResult) > 0)
    el.SetOptionRunAsyncCDP = False

    'SetOptionStopException: Trueにすると、JS例外発生時にVBA側でErr.Raiseされる
    '※エラートラップ：`クラスモジュールで中断`では止まります。設定で変更願います
    el.SetOptionStopException = True
    Dim exNumber As Long
    On Error Resume Next
    Err.Clear
    el.jsEval "function(){ throw new Error('CDPElement-ShadowDOM-test'); }"
    exNumber = Err.Number
    On Error GoTo 0
    el.SetOptionStopException = False
    AssertTrue "[" & modeLabel & "] SetOptionStopException=True時、JS例外がVBAエラーになること", (exNumber <> 0)

    'SetOptionUserGesture: エラーなく設定・使用・解除できること
    Dim errUserGesture As Long
    On Error Resume Next
    Err.Clear
    el.SetOptionUserGesture = True
    el.jsEval "function(){ return 1 }"
    el.SetOptionUserGesture = False
    errUserGesture = Err.Number
    On Error GoTo 0
    AssertTrue "[" & modeLabel & "] SetOptionUserGestureの設定・解除でエラーが起きないこと", (errUserGesture = 0)
End Sub

'==============================================================================
' SD⑯ sendHover
'==============================================================================
Private Sub Test16_SendHover(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑯ sendHover"

    RunSendHoverTest rootOpen, "open"
    RunSendHoverTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce16','SD⑯ 完了 " & EOk() & " | sendHover（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunSendHoverTest(root As CDPElement, modeLabel As String)
    Dim hv As CDPElement: Set hv = root.getElementByID("sHoverTarget")
    AssertEq "[" & modeLabel & "] sendHover前のhovered状態", hv.getAttribute("data-hovered"), "false"
    AssertEq "[" & modeLabel & "] sendHover前のhover回数", hv.getAttribute("data-hovercount"), "0"

    hv.sendHover
    AssertEq "[" & modeLabel & "] sendHover後のhovered状態", hv.getAttribute("data-hovered"), "true"
    AssertEq "[" & modeLabel & "] sendHover後のhover回数", hv.getAttribute("data-hovercount"), "1"
End Sub

'==============================================================================
' SD⑰ ホバーで出現するボタン（sendHover → sendClick）
'==============================================================================
Private Sub Test17_HoverReveal_Click(br As CDPContext, rootOpen As CDPElement, rootClosed As CDPElement)
    PrintSection "SD⑰ ホバーで出現するボタン"

    RunHoverRevealTest rootOpen, "open"
    RunHoverRevealTest rootClosed, "closed"

    br.jsEval "updateStatus('s-sce17','SD⑰ 完了 " & EOk() & " | ホバーで出現するボタン（open+closed）', true)", StopApiError:=False
End Sub

Private Sub RunHoverRevealTest(root As CDPElement, modeLabel As String)
    Dim container As CDPElement: Set container = root.getElementByID("sHoverRevealContainer")
    Dim revealBtn As CDPElement: Set revealBtn = root.getElementByID("sHoverRevealBtn")

    AssertEq "[" & modeLabel & "] ホバー前のクリック回数", revealBtn.getAttribute("data-clickcount"), "0"

    container.sendHover
    revealBtn.sendClick
    AssertEq "[" & modeLabel & "] ホバー→sendClick後のクリック回数", revealBtn.getAttribute("data-clickcount"), "1"
End Sub

'==============================================================================
' ヘルパー
'==============================================================================
Private Sub AssertEq(Name As String, actual As Variant, expected As Variant)
    If CStr(actual) = CStr(expected) Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → """ & CStr(actual) & """"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | " & Name & " 期待:""" & CStr(expected) & """ 実際:""" & CStr(actual) & """"
    End If
End Sub

Private Sub AssertNotEmpty(Name As String, actual As String)
    If Len(actual) > 0 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → NOT EMPTY"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | " & Name & " が空です"
    End If
End Sub

Private Sub AssertContains(Name As String, actual As String, substring As String)
    If InStr(actual, substring) > 0 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → """ & substring & """ を含む"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | " & Name & " に """ & substring & """ が含まれません。実際:""" & actual & """"
    End If
End Sub

Private Sub AssertTrue(Name As String, actual As Boolean)
    If actual Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → True"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | " & Name & " が True になりませんでした"
    End If
End Sub

Private Sub AssertFalse(Name As String, actual As Boolean)
    If Not actual Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → False"
    Else
        failCount = failCount + 1
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
