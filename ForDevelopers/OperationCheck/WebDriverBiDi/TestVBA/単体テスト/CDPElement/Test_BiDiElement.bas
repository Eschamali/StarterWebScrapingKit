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
    Test14_GetIFrame br
    Test15_Diagnostics_And_Options br
    Test16_UnsupportedMethods br

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
' ② value / clearValue
'==============================================================================
Private Sub Test02_Value_ClearValue(br As WebDriverBiDiContext)
    PrintSection "② value / clearValue"

    Dim el As CDPElement: Set el = br.getElementByID("testInput")

    el.value = "Hello VBA"
    AssertEq "value(Let→Get)", el.value, "Hello VBA"

    el.value = "改行" & vbLf & "や""引用符""を含む値"
    AssertEq "value(特殊文字)", el.value, "改行" & "や""引用符""を含む値"

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
' ⑭ getIFrame（同一オリジンiframe）
'==============================================================================
Private Sub Test14_GetIFrame(br As WebDriverBiDiContext)
    PrintSection "⑭ getIFrame"

    CDPHelpers.Sleep 0.5

    Dim iframeDoc As CDPElement: Set iframeDoc = br.getElementByID("testIFrame").getIFrame
    AssertEq "iframe内要素のinnerText", iframeDoc.getElementByID("iframeTarget").innerText, "iframe content"
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

    el.SetOptionRunAsyncCDP = True
    Dim asyncResult As Variant
    asyncResult = el.jsEval("function(){ return 1 }")
    el.SetOptionRunAsyncCDP = False
    AssertTrue "SetOptionRunAsyncCDP=True時、戻り値が非同期コマンドIDであること", (IsNumeric(asyncResult) And CLng(asyncResult) > 0)

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
' ⑯ BiDiモード未対応メソッドが、エラーで止まること
'==============================================================================
Private Sub Test16_UnsupportedMethods(br As WebDriverBiDiContext)
    PrintSection "⑯ BiDiモード未対応メソッド"

    Dim el As CDPElement: Set el = br.getElementByID("testSendClickBtn")
    Dim errNum As Long

    On Error Resume Next
    Err.Clear
    el.sendClick
    errNum = Err.Number
    On Error GoTo 0
    AssertTrue "sendClickが未対応エラーになること", (errNum <> 0)
    AssertEq "sendClickでCDPが送信されていないこと", el.getAttribute("data-clickcount"), "0"
End Sub

'==============================================================================
' ヘルパー
'==============================================================================
Private Sub RecordFail(Name As String)
    failCount = failCount + 1
    failNames = failNames & vbLf & "  - " & Name
End Sub

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
