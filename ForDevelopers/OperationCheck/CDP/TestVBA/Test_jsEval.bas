Attribute VB_Name = "Test_jsEval"
'==============================================================================
' CDPBrowser.jsEval 動作確認（Runtime.evaluate / Runtime.callFunctionOn）
' ・HTML: ForDevelopers\OperationCheck\TestHtml\Test_jsEval\Test_jsEval.html
' ・実行前に CDP を起動し、WORKSPACE_PATH をルートに設定してから RunAll_jsEval_Tests を実行
'==============================================================================
Option Explicit

Private passCount As Long
Private failCount As Long

'ワークスペースパス（StarterWebScrapingKit ルート）
Private Const WORKSPACE_PATH As String = ""

Private Function EOk() As String
    EOk = WorksheetFunction.Unichar(9989)
End Function

'--- CDP Runtime.callFunctionOn 用: arguments 配列の要素 ---
Private Function ArgVal(v As Variant) As Scripting.Dictionary
    Dim d As New Scripting.Dictionary
    d.Add "value", v
    Set ArgVal = d
End Function

' returnByValue の JSON 配列は、環境により Scripting.Dictionary（キー "0"…）または
' Collection（1 始まり）として返る。Collection に Exists は無いため TypeName で分岐する。
Private Function DictArrItem(d As Object, idx As Long) As Variant
    Dim ks As String

    Select Case TypeName(d)
        Case "Collection"
            DictArrItem = d.Item(idx + 1)
        Case "Dictionary"
            ks = CStr(idx)
            If d.Exists(ks) Then
                DictArrItem = d(ks)
            ElseIf d.Exists(idx) Then
                DictArrItem = d(idx)
            Else
                DictArrItem = Empty
            End If
        Case Else
            Err.Raise vbObjectError + 513, "Test_jsEval.DictArrItem", "配列要素の型が未対応です: " & TypeName(d)
    End Select
End Function

' serializationOptions: deep 時の DeepSerializedValue（type + value のペア配列）と、
' 通常の入れ子 Dictionary の両方から子を辿る。
Private Function DsvChildObject(ByVal parent As Object, ByVal Key As String) As Object
    If parent Is Nothing Then Exit Function
    On Error Resume Next
    If parent.Exists(Key) Then
        Set DsvChildObject = parent(Key)
        If Err.Number = 0 Then Exit Function
    End If
    Err.Clear
    If Not parent.Exists("type") Then Exit Function
    If parent("type") <> "object" Then Exit Function
    If Not parent.Exists("value") Then Exit Function
    Dim pairs As Object
    Set pairs = parent("value")
    Dim pair As Variant
    If TypeName(pairs) = "Collection" Then
        For Each pair In pairs
            If TypeName(pair) = "Collection" Then
                If pair(1) = Key Then
                    Set DsvChildObject = pair(2)
                    Exit Function
                End If
            End If
        Next
    ElseIf TypeName(pairs) = "Dictionary" Then
        Dim pk As Variant
        For Each pk In pairs
            Set pair = pairs(pk)
            If TypeName(pair) = "Collection" Then
                If pair(1) = Key Then
                    Set DsvChildObject = pair(2)
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Private Function DsvNodeAsDouble(ByVal Node As Variant) As Double
    If IsObject(Node) Then
        Dim o As Object
        Set o = Node
        If o Is Nothing Then Err.Raise 5
        If o.Exists("type") And o("type") = "number" And o.Exists("value") Then
            DsvNodeAsDouble = CDbl(o("value"))
            Exit Function
        End If
    End If
    DsvNodeAsDouble = CDbl(Node)
End Function

Private Function DsvGetPropertyNumber(ByVal parent As Object, ByVal Key As String) As Double
    Dim ch As Variant
    On Error Resume Next
    If parent.Exists(Key) Then
        ch = parent(Key)
        DsvGetPropertyNumber = DsvNodeAsDouble(ch)
        Exit Function
    End If
    Err.Clear
    Dim o As Object
    Set o = DsvChildObject(parent, Key)
    DsvGetPropertyNumber = DsvNodeAsDouble(o)
End Function

'==============================================================================
' エントリ
'==============================================================================
Public Sub RunAll_jsEval_Tests()
    Dim br As CDPBrowser: Set br = 設定シートからのCDP起動

    br.navigate "file:///" & Replace(WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_jsEval\Test_jsEval.html", "\", "/")
    br.wait

    passCount = 0: failCount = 0

    PrintHeader "jsEval 検証テスト 開始"

    Test01_Evaluate_primitives br
    Test02_Evaluate_returnByValue_object_and_array br
    Test03_Evaluate_undefined_null br
    Test04_Evaluate_unicode br
    Test05_Evaluate_promise_br br
    Test06_callFunctionOn_get_objectId br
    Test07_callFunctionOn_no_args br
    Test08_callFunctionOn_many_args br
    Test09_callFunctionOn_apostrophe_string br
    Test10_callFunctionOn_nested br
    Test11_exception_stopException_off br
    Test12_exception_IFEXCEPTION br
    Test13_long_string br
    Test14_contextId_isolatedWorld br
    Test15_serializationOptions_deep br
    Test16_RunAsyncCDP_alert br

    PrintHeader "テスト完了: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)

    br.jsEval "updateStatus('s-summary','PASS=" & passCount & " FAIL=" & failCount & " " & EOk() & "', true)", dbgMsg:=False

    br.quit
End Sub

'==============================================================================
' ① Runtime.evaluate - プリミティブ
'==============================================================================
Private Sub Test01_Evaluate_primitives(br As CDPBrowser)
    PrintSection "① evaluate - 数値・文字列・真偽"

    Dim v As Variant

    v = br.jsEval("2 + 40", dbgMsg:=False)
    AssertEq "数値 42", v, 42#

    v = br.jsEval("'hello-jsEval'", dbgMsg:=False)
    AssertEq "文字列", CStr(v), "hello-jsEval"

    v = br.jsEval("true", dbgMsg:=False)
    AssertEq "真偽 True", CStr(v), "True"

    v = br.jsEval("false", dbgMsg:=False)
    AssertEq "真偽 False", CStr(v), "False"

    br.jsEval "updateStatus('s-js01','① 完了 " & EOk() & " | 数値・文字列・真偽', true)", dbgMsg:=False
End Sub

'==============================================================================
' ② returnByValue - オブジェクト・配列（Dictionary）
'==============================================================================
Private Sub Test02_Evaluate_returnByValue_object_and_array(br As CDPBrowser)
    PrintSection "② evaluate - returnByValue オブジェクト / 配列"

    Dim o As Object, a As Object

    Set o = br.jsEval("({ alpha: 1, beta: 'z', gamma: true })", returnByValue:=True, dbgMsg:=False)
    AssertEq "obj.alpha", CDbl(o("alpha")), 1#
    AssertEq "obj.beta", CStr(o("beta")), "z"
    AssertEq "obj.gamma", CStr(o("gamma")), "True"

    Set a = br.jsEval("[10, 20, 30]", returnByValue:=True, dbgMsg:=False)
    AssertEq "arr[0]", CDbl(DictArrItem(a, 0)), 10#
    AssertEq "arr[1]", CDbl(DictArrItem(a, 1)), 20#
    AssertEq "arr[2]", CDbl(DictArrItem(a, 2)), 30#

    Set o = br.jsEval("window.__JSEVAL_GLOBAL", returnByValue:=True, dbgMsg:=False)
    AssertEq "global.num", CDbl(o("num")), 7#
    AssertEq "global.text", CStr(o("text")), "グローバル文字列"

    br.jsEval "updateStatus('s-js02','② 完了 " & EOk() & " | obj / arr / global', true)", dbgMsg:=False
End Sub

'==============================================================================
' ③ undefined / null
'==============================================================================
Private Sub Test03_Evaluate_undefined_null(br As CDPBrowser)
    PrintSection "③ evaluate - undefined / null"

    Dim v As Variant

    v = br.jsEval("void 0", returnByValue:=True, dbgMsg:=False)
    If IsEmpty(v) Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | undefined → Empty"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | undefined 期待 Empty 実際: " & TypeName(v) & " " & VarType(v)
    End If

    v = br.jsEval("null", returnByValue:=True, dbgMsg:=False)
    If IsNull(v) Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | null → Null"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | null 期待 Null"
    End If

    br.jsEval "updateStatus('s-js03','③ 完了 " & EOk() & " | Empty / Null', true)", dbgMsg:=False
End Sub

'==============================================================================
' ④ Unicode（日本語）
'==============================================================================
Private Sub Test04_Evaluate_unicode(br As CDPBrowser)
    PrintSection "④ evaluate - Unicode"

    Dim v As Variant
    v = br.jsEval("'" & "日本語_VBA連結" & "'", dbgMsg:=False)
    AssertEq "日本語リテラル", CStr(v), "日本語_VBA連結"

    br.jsEval "updateStatus('s-js04','④ 完了 " & EOk() & " | 日本語リテラル', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑤ awaitPromise
'==============================================================================
Private Sub Test05_Evaluate_promise_br(br As CDPBrowser)
    PrintSection "⑤ evaluate - awaitPromise"

    Dim v As Variant
    v = br.jsEval("Promise.resolve(123)", awaitPromise:=True, returnByValue:=True, dbgMsg:=False)
    AssertEq "Promise.resolve(123)", CDbl(v), 123#

    br.jsEval "updateStatus('s-js05','⑤ 完了 " & EOk() & " | Promise.resolve(123)', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑥ objectId 取得（returnByValue:=False）
'==============================================================================
Private Sub Test06_callFunctionOn_get_objectId(br As CDPBrowser)
    PrintSection "⑥ objectId 取得"

    Dim oid As Variant
    oid = br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, dbgMsg:=False)

    If VarType(oid) = vbString And Len(CStr(oid)) > 0 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | objectId 文字列取得 (先頭数文字): " & Left$(CStr(oid), 24) & "..."
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | objectId が取得できません"
    End If

    br.jsEval "updateStatus('s-js06','⑥ 完了 " & EOk() & " | objectId 文字列取得', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑦ callFunctionOn - 引数なし・this
'==============================================================================
Private Sub Test07_callFunctionOn_no_args(br As CDPBrowser)
    PrintSection "⑦ callFunctionOn - 引数なし"

    Dim oid As String
    oid = CStr(br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, dbgMsg:=False))

    Dim v As Variant
    v = br.jsEval("function(){ return this.id }", objectId:=oid, returnByValue:=True, dbgMsg:=False)
    AssertEq "this.id", CStr(v), "jseval-box"

    v = br.jsEval("function(){ return this.dataset.tag }", objectId:=oid, returnByValue:=True, dbgMsg:=False)
    AssertEq "dataset.tag", CStr(v), "jseval-data"

    br.jsEval "updateStatus('s-js07','⑦ 完了 " & EOk() & " | this.id / dataset', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑧ 多引数（10個）
'==============================================================================
Private Sub Test08_callFunctionOn_many_args(br As CDPBrowser)
    PrintSection "⑧ callFunctionOn - 多引数"

    Dim oid As String
    oid = CStr(br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, dbgMsg:=False))

    Dim args As New Collection
    Dim i As Long
    For i = 1 To 10
        args.Add ArgVal(i)
    Next i

    Dim v As Variant
    v = br.jsEval("function(a,b,c,d,e,f,g,h,i,j){ return a+b+c+d+e+f+g+h+i+j }", objectId:=oid, objectArguments:=args, returnByValue:=True, dbgMsg:=False)
    AssertEq "1..10 の和", CDbl(v), 55#

    br.jsEval "updateStatus('s-js08','⑧ 完了 " & EOk() & " | 10 引数 → 和 55', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑨ シングルクォートを含む文字列（objectArguments）
'==============================================================================
Private Sub Test09_callFunctionOn_apostrophe_string(br As CDPBrowser)
    PrintSection "⑨ callFunctionOn - アポストロフィ文字列"

    Dim oid As String
    oid = CStr(br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, dbgMsg:=False))

    Dim args As New Collection
    args.Add ArgVal("It's " & "OK " & "日本語")

    Dim v As Variant
    v = br.jsEval("function(s){ return 'ECHO:' + s }", objectId:=oid, objectArguments:=args, returnByValue:=True, dbgMsg:=False)
    AssertEq "エコー", CStr(v), "ECHO:It's OK 日本語"

    br.jsEval "updateStatus('s-js09','⑨ 完了 " & EOk() & " | objectArguments（引用符含む）エコー OK', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑩ 子要素参照（ネスト）
'==============================================================================
Private Sub Test10_callFunctionOn_nested(br As CDPBrowser)
    PrintSection "⑩ callFunctionOn - 子要素"

    Dim oid As String
    oid = CStr(br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, dbgMsg:=False))

    Dim v As Variant
    v = br.jsEval("function(){ return this.querySelector('#jseval-target').textContent }", objectId:=oid, returnByValue:=True, dbgMsg:=False)
    AssertNotEmpty "子 span.textContent", CStr(v)

    v = br.jsEval("function(){ return this.querySelector('#jseval-input').value }", objectId:=oid, returnByValue:=True, dbgMsg:=False)
    AssertEq "input.value 初期", CStr(v), "初期値"

    br.jsEval "updateStatus('s-js10','⑩ 完了 " & EOk() & " | 子 span / input', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑪ JS 例外（StopException 省略 = False → CVErr または Error 型）
'==============================================================================
Private Sub Test11_exception_stopException_off(br As CDPBrowser)
    PrintSection "⑪ 例外 - StopException=False"

    Dim r As Variant
    r = br.jsEval("(function(){ throw new Error('jsEval-test'); })()", StopException:=False, dbgMsg:=False)

    If IsError(r) Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | 例外時 IsError(CVErr)"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | 例外時に Error 型でない: " & TypeName(r)
    End If

    br.jsEval "updateStatus('s-js11','⑪ 完了 " & EOk() & " | 例外 → IsError', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑫ IFEXCEPTION
'==============================================================================
Private Sub Test12_exception_IFEXCEPTION(br As CDPBrowser)
    PrintSection "⑫ 例外 - IFEXCEPTION"

    Dim r As Variant
    r = br.jsEval("(function(){ throw new Error('x'); })()", StopException:=False, IFEXCEPTION:="fallback-ok", dbgMsg:=False)

    AssertEq "IFEXCEPTION 文字列", CStr(r), "fallback-ok"

    br.jsEval "updateStatus('s-js12','⑫ 完了 " & EOk() & " | IFEXCEPTION フォールバック', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑬ 長い文字列（返却値サイズ）
'==============================================================================
Private Sub Test13_long_string(br As CDPBrowser)
    PrintSection "⑬ 長い文字列 returnByValue"

    Dim v As Variant
    v = br.jsEval("'x'.repeat(2500)", returnByValue:=True, dbgMsg:=False)

    If Len(CStr(v)) = 2500 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | 長さ 2500"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | 長さ期待 2500 実際 " & Len(CStr(v))
    End If

    br.jsEval "updateStatus('s-js13','⑬ 完了 " & EOk() & " | 長さ 2500 文字', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑭ contextId ? Page.createIsolatedWorld の executionContextId で evaluate
'==============================================================================
Private Sub Test14_contextId_isolatedWorld(br As CDPBrowser)
    PrintSection "⑭ contextId ? createIsolatedWorld"

    On Error GoTo Test14_Err

    br.invokeMethod "Page.enable", Nothing

    Dim ftRes As Scripting.Dictionary
    Set ftRes = br.invokeMethod("Page.getFrameTree", Nothing)

    Dim rootFrameId As String
    rootFrameId = CStr(ftRes("frameTree")("frame")("id"))

    Dim pCW As New Scripting.Dictionary
    pCW.Add "frameId", rootFrameId
    pCW.Add "worldName", "jsEvalTestIsolated"

    Dim cwRes As Scripting.Dictionary
    Set cwRes = br.invokeMethod("Page.createIsolatedWorld", pCW)

    Dim execCtx As Long
    execCtx = CLng(cwRes("executionContextId"))

    Dim v As Variant
    v = br.jsEval("window.__JSEVAL_ISO = 'ctx-ok'; window.__JSEVAL_ISO", contextId:=execCtx, returnByValue:=True, dbgMsg:=False)
    AssertEq "isolated で代入→取得", CStr(v), "ctx-ok"

    Dim vMain As Variant
    vMain = br.jsEval("window.__JSEVAL_ISO", returnByValue:=True, dbgMsg:=False)
    If IsEmpty(vMain) Or VarType(vMain) = vbNull Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | メイン context では __JSEVAL_ISO 未定義（Empty/Null）"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | メイン context に隔離値が見えている: " & CStr(vMain)
    End If

    br.jsEval "updateStatus('s-js14','⑭ 完了 " & EOk() & " | contextId=" & CStr(execCtx) & "', true)", dbgMsg:=False
    Exit Sub

Test14_Err:
    failCount = failCount + 1
    Debug.Print "  FAIL | ⑭ " & Err.Description
    On Error Resume Next
    br.jsEval "updateStatus('s-js14','⑭ FAIL', false)", dbgMsg:=False
End Sub

'==============================================================================
' ⑮ serializationOptions ? deep（deepSerializedValue 優先）
'==============================================================================
Private Sub Test15_serializationOptions_deep(br As CDPBrowser)
    PrintSection "⑮ serializationOptions ? deep"

    Dim serOpts As New Scripting.Dictionary
    serOpts.Add "serialization", "deep"
    serOpts.Add "maxDepth", 8

    Dim resObj As Object
    Set resObj = br.jsEval("({ top: 1, nest: { mid: 2, deep: { leaf: 3 } } })", returnByValue:=True, serializationOptions:=serOpts, dbgMsg:=False)

    If resObj Is Nothing Then
        failCount = failCount + 1
        Debug.Print "  FAIL | ⑮ 結果が Nothing"
    Else
        On Error GoTo Test15_Err
        Dim nest As Object, deep As Object
        Set nest = DsvChildObject(resObj, "nest")
        Set deep = DsvChildObject(nest, "deep")
        Dim nestMid As Double, leafVal As Double
        nestMid = DsvGetPropertyNumber(nest, "mid")
        leafVal = DsvGetPropertyNumber(deep, "leaf")
        If nestMid = 2# And leafVal = 3# Then
            passCount = passCount + 1
            Debug.Print "  " & EOk() & " PASS | ⑮ deep ネスト nest.mid=2 leaf=3 Type=" & TypeName(resObj)
        Else
            failCount = failCount + 1
            Debug.Print "  FAIL | ⑮ ネスト値 nest.mid=" & nestMid & " leaf=" & leafVal
        End If
        GoTo Test15_Done
Test15_Err:
        failCount = failCount + 1
        Debug.Print "  FAIL | ⑮ " & Err.Description
Test15_Done:
    End If

    br.jsEval "updateStatus('s-js15','⑮ 完了 " & EOk() & " | serialization=deep', true)", dbgMsg:=False
End Sub

'==============================================================================
' ⑯ RunAsyncCDP ? alert（Demo_CDP.TestAlert と同系）
'==============================================================================
Private Sub Test16_RunAsyncCDP_alert(br As CDPBrowser)
    PrintSection "⑯ RunAsyncCDP ? alert"

    On Error GoTo Test16_Err

    br.invokeMethod "Page.enable", Nothing

    Dim oid As Variant
    oid = br.jsEval("document.getElementById('btn-async-alert')", returnByValue:=False, dbgMsg:=False)

    If VarType(oid) <> vbString Or Len(oid) = 0 Then
        failCount = failCount + 1
        Debug.Print "  FAIL | ⑯ ボタン objectId 取得失敗"
        br.jsEval "updateStatus('s-js16','⑯ FAIL ボタンなし', false)", dbgMsg:=False
        Exit Sub
    End If

    Dim asyncCmdId As Variant
    asyncCmdId = br.jsEval("function(){ this.click(); }", CStr(oid), RunAsyncCDP:=True, dbgMsg:=False)

    If IsNumeric(asyncCmdId) And CLng(asyncCmdId) > 0 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | RunAsyncCDP 戻り値 ID=" & CStr(asyncCmdId)
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | RunAsyncCDP 戻り値: " & CStr(asyncCmdId) & " VarType=" & VarType(asyncCmdId)
    End If

    Set br.BrowserEvents = New Dictionary

    Const evName As String = "Page.javascriptDialogOpening"
    Dim i As Long
    Dim found As Boolean
    For i = 1 To 100
        br.TakeEvents
        If br.BrowserEvents("EventMethods").Exists(evName) Then
            found = True
            Exit For
        End If
        br.sleep 0.05
    Next i

    If found Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & evName & " を検知"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | ダイアログイベントがタイムアウト"
    End If

    Dim pDlg As New Scripting.Dictionary
    pDlg.Add "accept", True
    br.invokeMethod "Page.handleJavaScriptDialog", pDlg

    Set br.BrowserEvents = Nothing

    br.jsEval "updateStatus('s-js16','⑯ 完了 " & EOk() & " | Async+alert+handleDialog', true)", dbgMsg:=False
    Exit Sub

Test16_Err:
    failCount = failCount + 1
    Debug.Print "  FAIL | ⑯ " & Err.Description
    On Error Resume Next
    Set br.BrowserEvents = Nothing
    br.jsEval "updateStatus('s-js16','⑯ FAIL', false)", dbgMsg:=False
End Sub

'==============================================================================
' ヘルパー
'==============================================================================
Private Sub AssertEq(Name As String, actual As Variant, expected As Variant)
    If VarType(actual) = vbDouble Or VarType(expected) = vbDouble Then
        If CDbl(actual) = CDbl(expected) Then
            passCount = passCount + 1
            Debug.Print "  " & EOk() & " PASS | " & Name & " → " & CStr(actual)
            Exit Sub
        End If
    ElseIf CStr(actual) = CStr(expected) Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → """ & CStr(actual) & """"
        Exit Sub
    End If
    failCount = failCount + 1
    Debug.Print "  FAIL | " & Name & " 期待:""" & CStr(expected) & """ 実際:""" & CStr(actual) & """"
End Sub

Private Sub AssertNotEmpty(Name As String, actual As String)
    If Len(Trim(actual)) > 0 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & Name & " → NOT EMPTY"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | " & Name & " が空"
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
