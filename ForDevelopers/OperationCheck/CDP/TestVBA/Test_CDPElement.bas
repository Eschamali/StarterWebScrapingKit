Attribute VB_Name = "Test_CDPElement"
'==============================================================================
' CDPElement.cls 動作確認（要素操作クラスの単体テスト）
' ・HTML: ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html
' ・実行前に CDP が起動可、WORKSPACE_PATH をルートに設定してから RunAll_CDPElement_Tests を実行
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
Public Sub RunAll_CDPElement_Tests()
    Dim br As CDPContext: Set br = ShSetting01_StartBrowser.StartCDPModeContext

    br.navigate "file:///" & Replace(WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html", "\", "/")
    br.wait

    passCount = 0: failCount = 0

    PrintHeader "CDPElement 検証テスト 開始"

    Test01_Value_SendString_ClearValue br
    Test02_InnerText_InnerHTML br
    Test03_Checked br
    Test04_Selected_SetSelection br
    Test05_Click_SimpleClick_FireEvent br
    Test06_SendClick_SendKey br
    Test07_GetAttribute_SetAttribute br
    Test08_Focus_SelectText br
    Test09_Submit br
    Test10_Traversal_Parent_Siblings_FirstChild br
    Test11_GetChildren_ElementsByQuery_ElementsByXPath br
    Test12_ElementByID_Query_XPath_Scoped br
    Test13_IsExist_IfExist_OnExist_OnExistNot br
    Test14_GetIFrame br
    Test15_ShadowRoot_Open_Closed br
    Test16_SetFileInputFiles br
    Test17_Diagnostics_And_Options br
    Test18_SendHover br

    PrintHeader "テスト結果: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)

    br.jsEval "updateStatus('s-summary','PASS=" & passCount & " FAIL=" & failCount & " " & EOk() & "', true)", StopApiError:=True

    br.InheritanceCDPBrowser.quit
End Sub

'==============================================================================
' ① value / sendString / clearValue
'==============================================================================
Private Sub Test01_Value_SendString_ClearValue(br As CDPContext)
    PrintSection "① value / sendString / clearValue"

    Dim el As CDPElement: Set el = br.getElementByID("testInput")

    el.value = "Hello VBA"
    AssertEq "value(Let→Get)", el.value, "Hello VBA"

    el.sendString "Real Key Input"
    AssertEq "sendString後のvalue", el.value, "Real Key Input"

    el.clearValue
    AssertEq "clearValue後のvalue", el.value, ""

    br.jsEval "updateStatus('s-ce01','① 完了 " & EOk() & " | value / sendString / clearValue', true)", StopApiError:=False
End Sub

'==============================================================================
' ② innerText / innerHTML
'==============================================================================
Private Sub Test02_InnerText_InnerHTML(br As CDPContext)
    PrintSection "② innerText / innerHTML"

    Dim txtEl As CDPElement: Set txtEl = br.getElementByID("testTextDiv")
    AssertEq "innerText初期値", txtEl.innerText, "initial text"
    txtEl.innerText = "changed text"
    AssertEq "innerText変更後", txtEl.innerText, "changed text"

    Dim htmlEl As CDPElement: Set htmlEl = br.getElementByID("testHtmlDiv")
    AssertContains "innerHTML初期値", htmlEl.innerHTML, "bold"
    htmlEl.innerHTML = "<i>italic</i>"
    AssertContains "innerHTML変更後", htmlEl.innerHTML, "italic"

    br.jsEval "updateStatus('s-ce02','② 完了 " & EOk() & " | innerText / innerHTML', true)", StopApiError:=False
End Sub

'==============================================================================
' ③ checked（チェックボックス）
'==============================================================================
Private Sub Test03_Checked(br As CDPContext)
    PrintSection "③ checked"

    Dim cb As CDPElement: Set cb = br.getElementByID("testCheckbox")

    cb.checked = True
    AssertTrue "checked=True設定後", cb.checked

    cb.checked = False
    AssertFalse "checked=False設定後", cb.checked

    br.jsEval "updateStatus('s-ce03','③ 完了 " & EOk() & " | checked', true)", StopApiError:=False
End Sub

'==============================================================================
' ④ selected / setSelection（セレクトボックス）
'==============================================================================
Private Sub Test04_Selected_SetSelection(br As CDPContext)
    PrintSection "④ selected / setSelection"

    Dim sel As CDPElement: Set sel = br.getElementByID("testSelect")

    '`setSelection`はvalue属性で指定して選択する
    sel.setSelection "opt2"
    AssertEq "setSelection後のvalue", sel.value, "opt2"

    '`selected`Letは、selectedIndexで指定して選択する（0始まり）
    sel.selected = "0"
    AssertEq "selected=0後のvalue", sel.value, "opt1"

    '`selected`Getは、選択中option要素のobjectIdを返す（文字列取得できていることの確認）
    AssertNotEmpty "selected（選択中option要素のobjectId）", CStr(sel.selected)

    br.jsEval "updateStatus('s-ce04','④ 完了 " & EOk() & " | selected / setSelection', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑤ click / SimpleClick / fireEvent
'==============================================================================
Private Sub Test05_Click_SimpleClick_FireEvent(br As CDPContext)
    PrintSection "⑤ click / SimpleClick / fireEvent"

    Dim btn As CDPElement: Set btn = br.getElementByID("testClickBtn")

    AssertEq "初期クリック回数", btn.getAttribute("data-clickcount"), "0"

    btn.click
    AssertEq "click後のクリック回数", btn.getAttribute("data-clickcount"), "1"

    btn.SimpleClick
    AssertEq "SimpleClick後のクリック回数", btn.getAttribute("data-clickcount"), "2"

    btn.fireEvent "customtestevent"
    AssertEq "fireEvent後のcustomfired", btn.getAttribute("data-customfired"), "true"

    br.jsEval "updateStatus('s-ce05','⑤ 完了 " & EOk() & " | click / SimpleClick / fireEvent', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑥ sendClick / sendKey
'==============================================================================
Private Sub Test06_SendClick_SendKey(br As CDPContext)
    PrintSection "⑥ sendClick / sendKey"

    Dim scBtn As CDPElement: Set scBtn = br.getElementByID("testSendClickBtn")
    AssertEq "sendClick前のクリック回数", scBtn.getAttribute("data-clickcount"), "0"
    scBtn.sendClick
    AssertEq "sendClick後のクリック回数", scBtn.getAttribute("data-clickcount"), "1"

    Dim ki As CDPElement: Set ki = br.getElementByID("testKeyInput")
    ki.sendKey keyEnter
    AssertEq "sendKey(keyEnter)後のkeyCode", ki.getAttribute("data-lastkeycode"), "13"

    ki.sendKey keyBackspace
    AssertEq "sendKey(keyBackspace)後のkeyCode", ki.getAttribute("data-lastkeycode"), "8"

    br.jsEval "updateStatus('s-ce06','⑥ 完了 " & EOk() & " | sendClick / sendKey', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑦ getAttribute / setAttribute
'==============================================================================
Private Sub Test07_GetAttribute_SetAttribute(br As CDPContext)
    PrintSection "⑦ getAttribute / setAttribute"

    Dim el As CDPElement: Set el = br.getElementByID("testAttrTarget")

    AssertEq "getAttribute初期値", el.getAttribute("data-foo"), "bar"

    el.setAttribute "data-foo", "baz"
    AssertEq "setAttribute後のgetAttribute", el.getAttribute("data-foo"), "baz"

    br.jsEval "updateStatus('s-ce07','⑦ 完了 " & EOk() & " | getAttribute / setAttribute', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑧ focus / selectText
'==============================================================================
Private Sub Test08_Focus_SelectText(br As CDPContext)
    PrintSection "⑧ focus / selectText"

    Dim fi As CDPElement: Set fi = br.getElementByID("testFocusInput")
    fi.focus
    Dim activeId As Variant
    activeId = br.jsEval("document.activeElement.id", StopApiError:=False)
    AssertEq "focus後のdocument.activeElement.id", CStr(activeId), "testFocusInput"

    Dim st As CDPElement: Set st = br.getElementByID("testSelectTextTarget")
    st.selectText
    Dim selectedText As Variant
    selectedText = br.jsEval("window.getSelection().toString()", StopApiError:=False)
    AssertEq "selectText後の選択文字列", CStr(selectedText), "Select this whole text"

    br.jsEval "updateStatus('s-ce08','⑧ 完了 " & EOk() & " | focus / selectText', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑨ submit
'==============================================================================
Private Sub Test09_Submit(br As CDPContext)
    PrintSection "⑨ submit"

    '`submit`は`this.form.submit()`を呼ぶため、通常はページ遷移が発生する。
    'テストページを壊さず検証するため、HTML側で`<form>`インスタンスの`submit`メソッドを
    'JSでオーバーライドし、呼び出しがあったことだけを記録するようにしている。
    Dim el As CDPElement: Set el = br.getElementByID("testSubmitInput")
    el.submit

    Dim frm As CDPElement: Set frm = br.getElementByID("testForm")
    AssertEq "submit呼び出し後のdata-submitted", frm.getAttribute("data-submitted"), "true"

    br.jsEval "updateStatus('s-ce09','⑨ 完了 " & EOk() & " | submit', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑩ DOM階層: getParent / getNextSibling / getPrevSibling / getFirstChild
'==============================================================================
Private Sub Test10_Traversal_Parent_Siblings_FirstChild(br As CDPContext)
    PrintSection "⑩ getParent / getNextSibling / getPrevSibling / getFirstChild"

    Dim parentEl As CDPElement: Set parentEl = br.getElementByID("traversalParent")
    Dim child1 As CDPElement: Set child1 = parentEl.getFirstChild
    AssertEq "getFirstChildのid", child1.getAttribute("id"), "child1"

    Dim child2 As CDPElement: Set child2 = child1.getNextSibling
    AssertEq "getNextSiblingのid", child2.getAttribute("id"), "child2"

    Dim child3 As CDPElement: Set child3 = br.getElementByID("child3")
    Dim backToChild2 As CDPElement: Set backToChild2 = child3.getPrevSibling
    AssertEq "getPrevSiblingのid", backToChild2.getAttribute("id"), "child2"

    Dim parentBack As CDPElement: Set parentBack = child2.getParent
    AssertEq "getParentのid", parentBack.getAttribute("id"), "traversalParent"

    br.jsEval "updateStatus('s-ce10','⑩ 完了 " & EOk() & " | getParent / getNextSibling / getPrevSibling / getFirstChild', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑪ getChildren / getElementsByQuery / getElementsByXPath（要素スコープ、複数取得）
'==============================================================================
Private Sub Test11_GetChildren_ElementsByQuery_ElementsByXPath(br As CDPContext)
    PrintSection "⑪ getChildren / getElementsByQuery / getElementsByXPath"

    Dim ul As CDPElement: Set ul = br.getElementByID("collectionList")

    Dim children As Collection: Set children = ul.getChildren
    AssertEq "getChildrenの件数", CStr(children.Count), "5"

    Dim byQuery As Collection: Set byQuery = ul.getElementsByQuery(".list-item")
    AssertEq "getElementsByQueryの件数", CStr(byQuery.Count), "5"
    AssertEq "getElementsByQuery[1]のdata-n", byQuery(1).getAttribute("data-n"), "1"

    Dim byXPath As Collection: Set byXPath = ul.getElementsByXPath("li")
    AssertEq "getElementsByXPathの件数", CStr(byXPath.Count), "5"

    br.jsEval "updateStatus('s-ce11','⑪ 完了 " & EOk() & " | getChildren / getElementsByQuery / getElementsByXPath', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑫ getElementByID / getElementByQuery / getElementByXPath（要素スコープ、単一取得）
'==============================================================================
Private Sub Test12_ElementByID_Query_XPath_Scoped(br As CDPContext)
    PrintSection "⑫ getElementByID / getElementByQuery / getElementByXPath（スコープ検索）"

    Dim ul As CDPElement: Set ul = br.getElementByID("collectionList")

    Dim byId As CDPElement: Set byId = ul.getElementByID("collectionItem3")
    AssertEq "getElementByID(スコープ)の内容", byId.innerText, "C"

    Dim byQuery As CDPElement: Set byQuery = ul.getElementByQuery("[data-n='2']")
    AssertEq "getElementByQuery(スコープ)の内容", byQuery.innerText, "B"

    Dim byXPath As CDPElement: Set byXPath = ul.getElementByXPath("li[3]")
    AssertEq "getElementByXPath(スコープ)の内容", byXPath.innerText, "C"

    br.jsEval "updateStatus('s-ce12','⑫ 完了 " & EOk() & " | getElementByID / getElementByQuery / getElementByXPath（スコープ）', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑬ isExist / ifExist / onExist / onExistNot
'==============================================================================
Private Sub Test13_IsExist_IfExist_OnExist_OnExistNot(br As CDPContext)
    PrintSection "⑬ isExist / ifExist / onExist / onExistNot"

    '1. 追加前は存在しないこと
    Dim dyn As CDPElement: Set dyn = br.getElementByID("dynamicElement")
    AssertFalse "追加前のisExist", dyn.isExist

    '2. ifExistチェーンは、存在しない要素へのメソッド呼び出しをエラーなく無視すること
    Dim errBefore As Long
    On Error Resume Next
    Err.Clear
    dyn.ifExist.focus
    errBefore = Err.Number
    On Error GoTo 0
    AssertTrue "ifExistチェーン（未存在）でエラーが起きないこと", (errBefore = 0)

    '3. 追加ボタンを押す（HTML側で1秒後に要素を追加）→ onExistでポーリング検知できること
    br.getElementByID("btnAddDynamic").click
    Dim dyn2 As CDPElement: Set dyn2 = br.getElementByID("dynamicElement")
    Set dyn2 = dyn2.onExist(timeOutInSeconds:=5)
    AssertTrue "追加後のonExist成功", Not (dyn2 Is Nothing)
    If Not (dyn2 Is Nothing) Then AssertTrue "onExist後のisExist", dyn2.isExist

    '4. 削除ボタンを押す（HTML側で1秒後に要素を削除）→ onExistNotで消滅検知できること
    br.getElementByID("btnRemoveDynamic").click
    AssertTrue "削除後のonExistNot成功", dyn2.onExistNot(timeOutInSeconds:=5)

    '5. 存在し得ないID指定時、onExist(raiseTimeoutError:=False)がタイムアウトでNothingを返すこと
    Dim ghost As CDPElement
    Set ghost = br.getElementByID("doesNotExist12345").onExist(timeOutInSeconds:=1, raiseTimeoutError:=False)
    AssertTrue "存在しない要素のonExistタイムアウトでNothingが返ること", (ghost Is Nothing)

    br.jsEval "updateStatus('s-ce13','⑬ 完了 " & EOk() & " | isExist / ifExist / onExist / onExistNot', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑭ getIFrame（同一オリジンiframe）
'==============================================================================
Private Sub Test14_GetIFrame(br As CDPContext)
    PrintSection "⑭ getIFrame"

    br.InheritanceCDPBrowser.sleep 0.5

    Dim iframeDoc As CDPElement: Set iframeDoc = br.getElementByID("testIFrame").getIFrame
    Dim iframeTarget As CDPElement: Set iframeTarget = iframeDoc.getElementByID("iframeTarget")
    AssertEq "iframe内要素のinnerText", iframeTarget.innerText, "iframe content"

    br.jsEval "updateStatus('s-ce14','⑭ 完了 " & EOk() & " | getIFrame', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑮ GetShadowRoot / GetShadowRoots（open / closed 両モード）
'==============================================================================
Private Sub Test15_ShadowRoot_Open_Closed(br As CDPContext)
    PrintSection "⑮ GetShadowRoot / GetShadowRoots（open / closed）"

    Dim hostOpen As CDPElement: Set hostOpen = br.getElementByID("shadowHostOpen")
    Dim rootOpen As CDPElement: Set rootOpen = hostOpen.GetShadowRoot()
    AssertTrue "GetShadowRoot(open)が取得できること", Not (rootOpen Is Nothing)
    If Not (rootOpen Is Nothing) Then
        AssertEq "shadow(open)内要素の内容", rootOpen.getElementByQuery(".shadow-content").innerText, "Shadow content (open)"
    End If
    br.jsEval "updateStatus('s-ce15a','⑮-a 完了 " & EOk() & " | GetShadowRoot(open)', true)", StopApiError:=False

    Dim hostClosed As CDPElement: Set hostClosed = br.getElementByID("shadowHostClosed")
    Dim rootsClosed As Collection: Set rootsClosed = hostClosed.GetShadowRoots()
    AssertTrue "GetShadowRoots(closed)が取得できること", Not (rootsClosed Is Nothing)
    If Not (rootsClosed Is Nothing) Then
        AssertEq "GetShadowRoots(closed)の件数", CStr(rootsClosed.Count), "1"
        AssertEq "shadow(closed)内要素の内容", rootsClosed(1).getElementByQuery(".shadow-content").innerText, "Shadow content (closed)"
    End If
    br.jsEval "updateStatus('s-ce15b','⑮-b 完了 " & EOk() & " | GetShadowRoots(closed)', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑯ SetFileInputFiles
'==============================================================================
Private Sub Test16_SetFileInputFiles(br As CDPContext)
    PrintSection "⑯ SetFileInputFiles"

    Dim dummyFile As String
    dummyFile = WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html"

    Dim files As New Collection
    files.Add dummyFile

    Dim fileInput As CDPElement: Set fileInput = br.getElementByID("testFileInput")
    fileInput.SetFileInputFiles files

    Dim FileName As Variant
    FileName = fileInput.jsEval("function(){ return this.files.length > 0 ? this.files[0].name : '' }")
    AssertEq "SetFileInputFiles後のfiles[0].name", CStr(FileName), "CDPElementTest.html"

    br.jsEval "updateStatus('s-ce16','⑯ 完了 " & EOk() & " | SetFileInputFiles', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑰ 診断プロパティ（UseSearchJS / varResult / CurrentObjectId / ExposeDevTools）
'    と実行オプション（SetOptionStopException / SetOptionRunAsyncCDP / SetOptionUserGesture）
'==============================================================================
Private Sub Test17_Diagnostics_And_Options(br As CDPContext)
    PrintSection "⑰ 診断プロパティ / 実行オプション"

    Dim el As CDPElement: Set el = br.getElementByID("testInput")

    AssertContains "UseSearchJSに検索構文が含まれること", el.UseSearchJS, "getElementById"
    AssertTrue "varResultがvbStringであること", (el.varResult = vbString)
    AssertNotEmpty "CurrentObjectIdが取得できていること", el.CurrentObjectId

    el.ExposeDevTools "__vbaTestExposed"
    Dim exposedCheck As Variant
    exposedCheck = br.jsEval("typeof window.__vbaTestExposed !== 'undefined' ? 'yes' : 'no'", StopApiError:=False)
    AssertEq "ExposeDevTools後にwindowへ公開されていること", CStr(exposedCheck), "yes"

    'SetOptionRunAsyncCDP: Trueにすると、jsEvalの戻り値が結果値ではなく非同期コマンドIDになる
    el.SetOptionRunAsyncCDP = True
    Dim asyncResult As Variant
    asyncResult = el.jsEval("function(){ return 1 }")
    AssertTrue "SetOptionRunAsyncCDP=True時、戻り値が数値の非同期コマンドIDであること", (IsNumeric(asyncResult) And CLng(asyncResult) > 0)
    el.SetOptionRunAsyncCDP = False

    'SetOptionStopException: Trueにすると、JS例外発生時にVBA側でErr.Raiseされる
    '※エラートラップ：`クラスモジュールで中断`では止まります。設定で変更願います
    el.SetOptionStopException = True
    Dim exNumber As Long
    On Error Resume Next
    Err.Clear
    el.jsEval "function(){ throw new Error('CDPElement-test'); }"
    exNumber = Err.Number
    On Error GoTo 0
    el.SetOptionStopException = False
    AssertTrue "SetOptionStopException=True時、JS例外がVBAエラーになること", (exNumber <> 0)

    'SetOptionUserGesture: エラーなく設定・使用・解除できること
    Dim errUserGesture As Long
    On Error Resume Next
    Err.Clear
    el.SetOptionUserGesture = True
    el.jsEval "function(){ return 1 }"
    el.SetOptionUserGesture = False
    errUserGesture = Err.Number
    On Error GoTo 0
    AssertTrue "SetOptionUserGestureの設定・解除でエラーが起きないこと", (errUserGesture = 0)

    br.jsEval "updateStatus('s-ce17','⑰ 完了 " & EOk() & " | 診断プロパティ / 実行オプション', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑱ sendHover
'==============================================================================
Private Sub Test18_SendHover(br As CDPContext)
    PrintSection "⑱ sendHover"

    Dim hv As CDPElement: Set hv = br.getElementByID("testHoverTarget")
    AssertEq "sendHover前のhovered状態", hv.getAttribute("data-hovered"), "false"
    AssertEq "sendHover前のhover回数", hv.getAttribute("data-hovercount"), "0"

    hv.sendHover
    AssertEq "sendHover後のhovered状態", hv.getAttribute("data-hovered"), "true"
    AssertEq "sendHover後のhover回数", hv.getAttribute("data-hovercount"), "1"

    br.jsEval "updateStatus('s-ce18','⑱ 完了 " & EOk() & " | sendHover', true)", StopApiError:=False
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
