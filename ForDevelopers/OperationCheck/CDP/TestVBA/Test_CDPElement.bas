Attribute VB_Name = "Test_CDPElement"
Option Explicit

'==============================================================================
' CDPElement 全機能テストモジュール
' ・テスト用HTMLファイル: ForDevelopers\OperationCheck\TestHtml\Test_CDPElement\CDPElementTest.html
' ・テスト実行前に CDPBrowser を開いて当該ページが表示されている状態にしてください
'==============================================================================


' ── テスト結果カウンタ
Private passCount As Long
Private failCount As Long
Private results As Collection

'ワークスペースパス
'※StarterWebScrapingKitのルートフォルダ を入力してください
Private Const WORKSPACE_PATH As String = ""

' ブラウザ updateStatus 等へ渡すチェック（モジュール保存で絵文字が化ける場合の代替・U+2705）
Private Function ECheck() As String
    ECheck = WorksheetFunction.Unichar(9989)
End Function

'==============================================================================
' Main: 全テスト実行
'==============================================================================
Public Sub RunAll_CDPElement_Tests()
    Dim br As CDPBrowser: Set br = 設定シートからのCDP起動

    '--- ブラウザ起動 & HTMLページへナビゲート ---
    br.navigate "file:///" & Replace(WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\CDPElementTest.html", "\", "/")
    br.wait

    passCount = 0
    failCount = 0
    Set results = New Collection

    PrintHeader "CDPElement 全機能テスト 開始"

    ' ─── 各テスト実行 ───
    Test01_Value br
    Test02_innerText br
    Test03_innerHTML br
    Test04_checked br
    Test05_selected br
    Test06_click br
    Test07_Attribute br
    Test08_focus_selectText br
    Test09_sendKey br
    Test10_submit br
    Test11_Traversal br
    Test12_Collections br
    Test13_isExist_onExist br
    Test14_getIFrame br

    '--- 最終サマリー ---
    PrintHeader "テスト完了: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)

    br.quit
End Sub

'==============================================================================
' ① value GET/LET / sendString / clearValue
'==============================================================================
Private Sub Test01_Value(br As CDPBrowser)
    PrintSection "① value / sendString / clearValue"
    Dim el As CDPElement

    ' --- value LET ---
    Set el = br.getElementByID("testInput")
    el.value = "VBAから設定した値"
    Dim got As String: got = el.value
    AssertEq "value LET→GET", got, "VBAから設定した値"
    br.jsEval "updateStatus('s-value','value LET: ' + document.getElementById('testInput').value, true)"

    ' --- clearValue ---
    el.clearValue
    AssertEq "clearValue後のvalue", el.value, ""
    br.jsEval "updateStatus('s-value','clearValue後: ' + document.getElementById('testInput').value, true)"

    ' --- sendString ---
    el.sendString "sendStringで入力"
    AssertEq "sendString後のvalue", el.value, "sendStringで入力"
    br.jsEval "updateStatus('s-value','sendString: ' + document.getElementById('testInput').value, true)"

    ' --- varPath / varResult (Basic Properties) ---
    AssertNotEmpty "varPath プロパティ取得", el.varPath
    AssertEq "varResult プロパティ取得(vbString=8)", CStr(el.varResult), "8"
End Sub

'==============================================================================
' ② innerText GET/LET
'==============================================================================
Private Sub Test02_innerText(br As CDPBrowser)
    PrintSection "② innerText GET/LET"
    Dim el As CDPElement
    Set el = br.getElementByID("testInnerText")

    ' GET
    Dim orig As String: orig = el.innerText
    AssertNotEmpty "innerText GET", orig

    ' LET（クォート/特殊文字含む）
    el.innerText = "VBAから設定: 「引用符」と 'アポ' テスト & 記号"
    AssertEq "innerText LET→GET", el.innerText, "VBAから設定: 「引用符」と 'アポ' テスト & 記号"
    br.jsEval "updateStatus('s-innertext', document.getElementById('testInnerText').innerText, true)"
End Sub

'==============================================================================
' ③ innerHTML GET/LET
'==============================================================================
Private Sub Test03_innerHTML(br As CDPBrowser)
    PrintSection "③ innerHTML GET/LET"
    Dim el As CDPElement
    Set el = br.getElementByID("testInnerHTML")

    ' GET
    Dim orig As String: orig = el.innerHTML
    AssertNotEmpty "innerHTML GET", orig

    ' LET
    el.innerHTML = "<span style='color:#6c63ff'>" & ECheck() & " VBAから設定した innerHTML</span>"
    AssertContains "innerHTML LET→GET", el.innerHTML, "VBAから設定"
    br.jsEval "updateStatus('s-innerhtml', 'innerHTML 更新済み " & ECheck() & "', true)"
End Sub

'==============================================================================
' ④ checked GET/LET
'==============================================================================
Private Sub Test04_checked(br As CDPBrowser)
    PrintSection "④ checked GET/LET"
    Dim el As CDPElement
    Set el = br.getElementByID("testCheckbox")

    ' LET = True
    el.checked = True
    AssertEq "checked LET=True", CStr(el.checked), "True"

    ' LET = False
    el.checked = False
    AssertEq "checked LET=False", CStr(el.checked), "False"
    br.jsEval "updateStatus('s-checked','checked テスト完了 " & ECheck() & "', true)"
End Sub

'==============================================================================
' ⑤ selected / setSelection
'==============================================================================
Private Sub Test05_selected(br As CDPBrowser)
    PrintSection "⑤ selected / setSelection"
    Dim el As CDPElement
    Set el = br.getElementByID("testSelect")

    ' selected LET = index
    el.selected = "1"
    Dim selVal As String: selVal = el.selected
    AssertNotEmpty "selected LET=1 → GET", selVal
    br.jsEval "updateStatus('s-selected','selected=' & document.getElementById('testSelect').selectedIndex, true)"

    ' setSelection (option value)
    el.setSelection "opt-c"
    selVal = el.selected
    AssertNotEmpty "setSelection(opt-c) → GET", selVal
    br.jsEval "updateStatus('s-selected','setSelection後 idx=' & document.getElementById('testSelect').selectedIndex, true)"
End Sub

'==============================================================================
' ⑥ click() / fireEvent()
'==============================================================================
Private Sub Test06_click(br As CDPBrowser)
    PrintSection "⑥ click / fireEvent"
    Dim el As CDPElement
    Set el = br.getElementByID("testButton")

    ' click
    el.click isLoading
    AssertPass "click() 実行"

    ' fireEvent
    el.fireEvent "click"
    AssertPass "fireEvent('click') 実行"

    ' sendClick
    Dim sendClickEl As CDPElement
    Set sendClickEl = br.getElementByID("testSendClickBtn")
    sendClickEl.sendClick
    AssertPass "sendClick() 実行"

    br.jsEval "updateStatus('s-click','click/fireEvent/sendClick テスト完了 " & ECheck() & "', true)"
End Sub

'==============================================================================
' ⑦ getAttribute / setAttribute
'==============================================================================
Private Sub Test07_Attribute(br As CDPBrowser)
    PrintSection "⑦ getAttribute / setAttribute"
    Dim el As CDPElement
    Set el = br.getElementByID("testAttr")

    ' getAttribute
    Dim attrVal As String: attrVal = el.getAttribute("data-custom")
    AssertEq "getAttribute(data-custom)", attrVal, "original-attr"

    ' setAttribute
    el.setAttribute "data-custom", "VBAから変更した属性値"
    AssertEq "setAttribute→getAttribute", el.getAttribute("data-custom"), "VBAから変更した属性値"
    br.jsEval "updateStatus('s-attr', document.getElementById('testAttr').dataset.custom, true)"
End Sub

'==============================================================================
' ⑧ focus / selectText
'==============================================================================
Private Sub Test08_focus_selectText(br As CDPBrowser)
    PrintSection "⑧ focus / selectText"

    ' focus
    Dim el As CDPElement
    Set el = br.getElementByID("testFocusInput")
    el.focus
    AssertPass "focus() 実行"

    ' selectText
    Dim ta As CDPElement
    Set ta = br.getElementByID("testTextArea")
    ta.selectText
    AssertPass "selectText() 実行"
    br.jsEval "updateStatus('s-focus','focus/selectText テスト完了 " & ECheck() & "', true)"
End Sub

'==============================================================================
' ⑨ sendKey
'==============================================================================
Private Sub Test09_sendKey(br As CDPBrowser)
    PrintSection "⑨ sendKey"

    ' Field A にフォーカス→ Tab で Field B へ移動
    Dim el1 As CDPElement
    Set el1 = br.getElementByID("keyInput1")
    el1.sendString "Field_A_Updated"
    el1.sendKey keyTab
    AssertPass "sendKey(Tab) 実行"
    br.jsEval "updateStatus('s-sendkey','sendKey(Tab) テスト完了 " & ECheck() & "', true)"

    ' Field B で Backspace 1回
    Dim el2 As CDPElement
    Set el2 = br.getElementByID("keyInput2")
    el2.sendKey keyBackspace
    AssertPass "sendKey(Backspace) 実行"
End Sub

'==============================================================================
' ⑩ submit
'==============================================================================
Private Sub Test10_submit(br As CDPBrowser)
    PrintSection "⑩ submit"
    Dim formEl As CDPElement
    Set formEl = br.getElementByID("testForm")
    formEl.submit isLoading
    AssertPass "submit メソッド 実行 (form要素に対して)"
    br.jsEval "updateStatus('s-submit','submit テスト完了 " & ECheck() & "', true)"
End Sub

'==============================================================================
' ⑪ トラバーサル: getParent / getNextSibling / getPrevSibling / getFirstChild
'==============================================================================
Private Sub Test11_Traversal(br As CDPBrowser)
    PrintSection "⑪ トラバーサル"

    Dim child2 As CDPElement
    Set child2 = br.getElementByID("traversalChild2")

    ' getParent
    Dim parent As CDPElement
    Set parent = child2.getParent()
    AssertEq "getParent → id", parent.getAttribute("id"), "traversalParent"

    ' getNextSibling
    Dim nextEl As CDPElement
    Set nextEl = child2.getNextSibling()
    AssertEq "getNextSibling → id", nextEl.getAttribute("id"), "traversalChild3"

    ' getPrevSibling
    Dim prevEl As CDPElement
    Set prevEl = child2.getPrevSibling()
    AssertEq "getPrevSibling → id", prevEl.getAttribute("id"), "traversalChild1"

    ' getFirstChild (parentから)
    Dim firstEl As CDPElement
    Set firstEl = parent.getFirstChild()
    AssertEq "getFirstChild → id", firstEl.getAttribute("id"), "traversalChild1"

    br.jsEval "updateStatus('s-traversal','トラバーサル テスト完了 " & ECheck() & "', true)"
End Sub

'==============================================================================
' ⑫ コレクション: getChildren / getElementsByQuery / getElementsByXPath
'==============================================================================
Private Sub Test12_Collections(br As CDPBrowser)
    PrintSection "⑫ コレクション"
    Dim ulEl As CDPElement
    Set ulEl = br.getElementByID("collection-list")

    ' getChildren
    Dim children As Collection
    Set children = ulEl.getChildren()
    AssertEq "getChildren → count", CStr(children.Count), "5"
    AssertEq "getChildren(1) data-n", children(1).getAttribute("data-n"), "1"
    AssertEq "getChildren(5) data-n", children(5).getAttribute("data-n"), "5"

    ' getElementsByQuery (from CDPBrowser)
    Dim Items As Collection
    Set Items = br.getElementsByQuery(".list-item")
    AssertEq "getElementsByQuery → count", CStr(Items.Count), "5"
    AssertEq "getElementsByQuery(3) データ", Items(3).getAttribute("data-n"), "3"

    ' getElementByXPath
    Dim xpEl As CDPElement
    Set xpEl = br.getElementByXPath("//li[@data-n='4']")
    AssertEq "getElementByXPath data-n=4", xpEl.getAttribute("data-n"), "4"

    ' getElementsByXPath
    Dim xpItems As Collection
    Set xpItems = br.getElementsByXPath("//li[contains(@class,'list-item')]")
    AssertEq "getElementsByXPath → count", CStr(xpItems.Count), "5"

    ' CDPElement内部 getElementByQuery
    Dim li2 As CDPElement
    Set li2 = ulEl.getElementByQuery("[data-n='2']")
    AssertEq "el.getElementByQuery → data-n", li2.getAttribute("data-n"), "2"

    ' CDPElement内部 getElementsByQuery
    Dim innerItems As Collection
    Set innerItems = ulEl.getElementsByQuery(".list-item")
    AssertEq "el.getElementsByQuery → count", CStr(innerItems.Count), "5"

    ' CDPElement内部 getElementByID / getElementByXPath / getElementsByXPath
    Dim innerIdEl As CDPElement
    ' ulEl doesn't have elements with ID inside it in the original HTML, but we can test from the body or form
    Dim testPanel As CDPElement: Set testPanel = br.getElementByID("test-panel")
    Set innerIdEl = testPanel.getElementByID("collection-list")
    AssertEq "el.getElementByID → id", innerIdEl.getAttribute("id"), "collection-list"

    Dim innerXPEl As CDPElement
    Set innerXPEl = ulEl.getElementByXPath(".//li[@data-n='5']")
    AssertEq "el.getElementByXPath → data-n", innerXPEl.getAttribute("data-n"), "5"

    Dim innerXPItems As Collection
    Set innerXPItems = ulEl.getElementsByXPath(".//li")
    AssertEq "el.getElementsByXPath → count", CStr(innerXPItems.Count), "5"

    br.jsEval "updateStatus('s-collection','コレクション テスト完了 " & ECheck() & "', true)"
End Sub

'==============================================================================
' ⑬ isExist / ifExist / onExist / onExistNot
'==============================================================================
Private Sub Test13_isExist_onExist(br As CDPBrowser)
    PrintSection "⑬ isExist / ifExist / onExist / onExistNot"

    ' isExist: 存在する要素
    Dim el As CDPElement
    Set el = br.getElementByID("testInput")
    AssertEq "testInput.isExist", CStr(el.isExist), "True"

    ' isExist: 存在しない要素
    Dim ghost As CDPElement
    Set ghost = br.getElementByID("nonExistentElement12345")
    AssertEq "nonExistent.isExist", CStr(ghost.isExist), "False"

    ' ifExist → 存在する場合のみ実行（チェーン確認）
    br.getElementByID("testInput").ifExist.focus
    AssertPass "ifExist.focus 実行（スキップなし）"

    ' ifExist → 存在しない場合はスキップ
    br.getElementByID("nonExistent99").ifExist.click isLoading
    AssertPass "ifExist.click スキップ確認（エラーなし）"

    ' onExist: 動的要素を追加してから待機
    br.jsEval "setTimeout(function(){ addDynamic() }, 800)"
    Dim dynEl As CDPElement
    Set dynEl = br.getElementByXPath("//div[@id='dynamicElement']")
    dynEl.onExist timeOutInSeconds:=5
    AssertEq "onExist 後にisExist", CStr(dynEl.isExist), "True"
    br.jsEval "updateStatus('s-exist','onExist 待機成功 " & ECheck() & "', true)"

    ' onExistNot: 削除してから待機
    br.jsEval "setTimeout(function(){ removeDynamic() }, 800)"
    dynEl.onExistNot timeOutInSeconds:=5
    AssertPass "onExistNot 完了"
    br.jsEval "updateStatus('s-exist','onExistNot 待機成功 " & ECheck() & "', true)"
End Sub

'==============================================================================
' ⑭ getIFrame
'==============================================================================
Private Sub Test14_getIFrame(br As CDPBrowser)
    PrintSection "⑭ getIFrame"
    Dim iframeEl As CDPElement
    Set iframeEl = br.getElementByID("testIFrame")

    Dim frameDoc As CDPElement
    Set frameDoc = iframeEl.getIFrame()
    AssertEq "getIFrame.isExist", CStr(frameDoc.isExist), "True"

    ' iFrame内の要素取得
    Dim innerEl As CDPElement
    Set innerEl = frameDoc.getElementByID("iframeContent")
    If innerEl.isExist Then
        Dim txt As String: txt = innerEl.innerText
        AssertNotEmpty "iFrame内 innerText", txt
        innerEl.innerText = "VBAからiFrameを変更 " & ECheck()
        AssertPass "iFrame内 innerText LET"
    Else
        AssertFail "iFrame内の要素が見つかりませんでした"
    End If
    br.jsEval "updateStatus('s-iframe','getIFrame テスト完了 " & ECheck() & "', true)"
End Sub

'==============================================================================
' ヘルパー: アサーション
'==============================================================================
Private Sub AssertEq(testName As String, actual As String, expected As String)
    If actual = expected Then
        passCount = passCount + 1
        Debug.Print "  ? PASS | " & testName & " → """ & actual & """"
    Else
        failCount = failCount + 1
        Debug.Print "  ? FAIL | " & testName & " → 期待値:""" & expected & """ 実際:""" & actual & """"
    End If
End Sub

Private Sub AssertNotEmpty(testName As String, actual As String)
    If Len(Trim(actual)) > 0 Then
        passCount = passCount + 1
        Debug.Print "  ? PASS | " & testName & " → NOT EMPTY (""" & Left(actual, 40) & """)"
    Else
        failCount = failCount + 1
        Debug.Print "  ? FAIL | " & testName & " → 空文字が返ってきました"
    End If
End Sub

Private Sub AssertContains(testName As String, actual As String, substr As String)
    If InStr(actual, substr) > 0 Then
        passCount = passCount + 1
        Debug.Print "  ? PASS | " & testName & " → contains """ & substr & """"
    Else
        failCount = failCount + 1
        Debug.Print "  ? FAIL | " & testName & " → """ & substr & """ が見つかりません"
    End If
End Sub

Private Sub AssertPass(testName As String)
    passCount = passCount + 1
    Debug.Print "  ? PASS | " & testName
End Sub

Private Sub AssertFail(testName As String)
    failCount = failCount + 1
    Debug.Print "  ? FAIL | " & testName
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
