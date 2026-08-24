Attribute VB_Name = "Test_jsEval"
'==============================================================================
' CDPContext.jsEval 動作確認（Runtime.evaluate / Runtime.callFunctionOn）
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
Private Function ArgVal(v As Variant) As Dictionary
    Dim d As New Dictionary
    d.Add "value", v
    Set ArgVal = d
End Function

' returnByValue の JSON 配列は、環境により Scripting.Dictionary（キー "0"…）または
' Collection（1 始まり）として返る。Collection に Exists は無いため TypeName で分岐する。
Private Function DictArrItem(d As Object, idx As Long) As Variant
    Dim ks As String

    ' TypeName で新しいクラス名を判定に追加します
    Select Case TypeName(d)
        Case "BiDiCDPJson", "JSON"
            ' JSON.cls/BiDiCDPJson は一貫して 0 始まりです
            ' ExistsIndex で範囲内かチェックし、ValueAt で値を取り出します
            If d.ExistsIndex(idx) Then
                DictArrItem = d.ValueAt(idx)
            Else
                DictArrItem = Empty
            End If

        Case "Collection"
            ' Collection は 1 始まりなので +1 が必要です
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
Private Function DsvChildObject(ByVal Parent As BiDiCDPJson, ByVal Key As String) As BiDiCDPJson
    If Parent Is Nothing Then Exit Function

    ' 1. 通常の入れ子構造（キーが直接存在する場合）
    ' ExistsKey で爆速判定し、NodeKey で軽量ノードを返します
    If Parent.ExistsKey(Key) Then
        Set DsvChildObject = Parent.NodeKey(Key)
        Exit Function
    End If

    ' 2. DeepSerializedValue (DSV) 構造の判定
    ' {"type": "object", "value": [ ["k1", v1], ["k2", v2] ]} のような形を想定
    If Parent.StringKey("type") = "object" Then
        Dim pairs As BiDiCDPJson: Set pairs = Parent.NodeKey("value")
        
        ' value が配列であることを確認してループ
        If pairs.IsArray Then
            Dim i As Long
            Dim pair As BiDiCDPJson
            
            ' 0始まりのインデックスでループを回します
            For i = 0 To pairs.Count - 1
                Set pair = pairs.NodeIndex(i) ' 1つのペア [key, value] を取得
                
                ' pair(0) が Key と一致するか判定
                If pair.StringAt(0) = Key Then
                    ' pair(1) をオブジェクト（ノード）として返却
                    Set DsvChildObject = pair.NodeAt(1)
                    Exit Function
                End If
            Next i
        End If
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

Private Function DsvGetPropertyNumber(ByVal Parent As Object, ByVal Key As String) As Double
    Dim ch As Variant
    On Error Resume Next
    If Parent.Exists(Key) Then
        ch = Parent(Key)
        DsvGetPropertyNumber = DsvNodeAsDouble(ch)
        Exit Function
    End If
    Err.Clear
    Dim o As Object
    Set o = DsvChildObject(Parent, Key)
    DsvGetPropertyNumber = DsvNodeAsDouble(o)
End Function

'==============================================================================
' エントリ
'==============================================================================
Public Sub RunAll_jsEval_Tests()
    Dim br As CDPContext: Set br = ShSetting01_StartBrowser.StartCDPModeContext

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

    br.jsEval "updateStatus('s-summary','PASS=" & passCount & " FAIL=" & failCount & " " & EOk() & "', true)", StopApiError:=True

    br.InheritanceCDPBrowser.quit
End Sub

'==============================================================================
' ① Runtime.evaluate - プリミティブ
'==============================================================================
Private Sub Test01_Evaluate_primitives(br As CDPContext)
    PrintSection "① evaluate - 数値・文字列・真偽"

    Dim v As Variant

    v = br.jsEval("2 + 40", StopApiError:=False)
    AssertEq "数値 42", v, 42#

    v = br.jsEval("'hello-jsEval'", StopApiError:=False)
    AssertEq "文字列", CStr(v), "hello-jsEval"

    v = br.jsEval("true", StopApiError:=False)
    AssertEq "真偽 True", CStr(v), "True"

    v = br.jsEval("false", StopApiError:=False)
    AssertEq "真偽 False", CStr(v), "False"

    br.jsEval "updateStatus('s-js01','① 完了 " & EOk() & " | 数値・文字列・真偽', true)", StopApiError:=False
End Sub

'==============================================================================
' ② returnByValue - オブジェクト・配列（Dictionary）
'==============================================================================
Private Sub Test02_Evaluate_returnByValue_object_and_array(br As CDPContext)
    PrintSection "② evaluate - returnByValue オブジェクト / 配列"

    Dim o As Object, a As Object

    Set o = br.jsEval("({ alpha: 1, beta: 'z', gamma: true })", returnByValue:=True, StopApiError:=False)
    AssertEq "obj.alpha", CDbl(o("alpha")), 1#
    AssertEq "obj.beta", CStr(o("beta")), "z"
    AssertEq "obj.gamma", CStr(o("gamma")), "True"

    Set a = br.jsEval("[10, 20, 30]", returnByValue:=True, StopApiError:=False)
    AssertEq "arr[0]", CDbl(DictArrItem(a, 0)), 10#
    AssertEq "arr[1]", CDbl(DictArrItem(a, 1)), 20#
    AssertEq "arr[2]", CDbl(DictArrItem(a, 2)), 30#

    Set o = br.jsEval("window.__JSEVAL_GLOBAL", returnByValue:=True, StopApiError:=False)
    AssertEq "global.num", CDbl(o("num")), 7#
    AssertEq "global.text", CStr(o("text")), "グローバル文字列"

    br.jsEval "updateStatus('s-js02','② 完了 " & EOk() & " | obj / arr / global', true)", StopApiError:=False
End Sub

'==============================================================================
' ③ undefined / null
'==============================================================================
Private Sub Test03_Evaluate_undefined_null(br As CDPContext)
    PrintSection "③ evaluate - undefined / null"

    Dim v As Variant

    v = br.jsEval("void 0", returnByValue:=True, StopApiError:=False)
    If IsEmpty(v) Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | undefined → Empty"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | undefined 期待 Empty 実際: " & TypeName(v) & " " & VarType(v)
    End If

    v = br.jsEval("null", returnByValue:=True, StopApiError:=False)
    If IsNull(v) Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | null → Null"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | null 期待 Null"
    End If

    br.jsEval "updateStatus('s-js03','③ 完了 " & EOk() & " | Empty / Null', true)", StopApiError:=False
End Sub

'==============================================================================
' ④ Unicode（日本語）
'==============================================================================
Private Sub Test04_Evaluate_unicode(br As CDPContext)
    PrintSection "④ evaluate - Unicode"

    Dim v As Variant
    v = br.jsEval("'" & "日本語_VBA連結" & "'", StopApiError:=False)
    AssertEq "日本語リテラル", CStr(v), "日本語_VBA連結"

    br.jsEval "updateStatus('s-js04','④ 完了 " & EOk() & " | 日本語リテラル', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑤ awaitPromise
'==============================================================================
Private Sub Test05_Evaluate_promise_br(br As CDPContext)
    PrintSection "⑤ evaluate - awaitPromise"

    Dim v As Variant
    v = br.jsEval("Promise.resolve(123)", awaitPromise:=True, returnByValue:=True, StopApiError:=False)
    AssertEq "Promise.resolve(123)", CDbl(v), 123#

    br.jsEval "updateStatus('s-js05','⑤ 完了 " & EOk() & " | Promise.resolve(123)', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑥ objectId 取得（returnByValue:=False）
'==============================================================================
Private Sub Test06_callFunctionOn_get_objectId(br As CDPContext)
    PrintSection "⑥ objectId 取得"

    Dim oid As Variant
    oid = br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, StopApiError:=False)

    If VarType(oid) = vbString And Len(CStr(oid)) > 0 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | objectId 文字列取得 (先頭数文字): " & Left$(CStr(oid), 24) & "..."
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | objectId が取得できません"
    End If

    br.jsEval "updateStatus('s-js06','⑥ 完了 " & EOk() & " | objectId 文字列取得', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑦ callFunctionOn - 引数なし・this
'==============================================================================
Private Sub Test07_callFunctionOn_no_args(br As CDPContext)
    PrintSection "⑦ callFunctionOn - 引数なし"

    Dim oid As String
    oid = CStr(br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, StopApiError:=False))

    Dim v As Variant
    v = br.jsEval("function(){ return this.id }", objectId:=oid, returnByValue:=True, StopApiError:=False)
    AssertEq "this.id", CStr(v), "jseval-box"

    v = br.jsEval("function(){ return this.dataset.tag }", objectId:=oid, returnByValue:=True, StopApiError:=False)
    AssertEq "dataset.tag", CStr(v), "jseval-data"

    br.jsEval "updateStatus('s-js07','⑦ 完了 " & EOk() & " | this.id / dataset', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑧ 多引数（10個）
'==============================================================================
Private Sub Test08_callFunctionOn_many_args(br As CDPContext)
    PrintSection "⑧ callFunctionOn - 多引数"

    Dim oid As String
    oid = CStr(br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, StopApiError:=False))

    Dim args(1 To 10) As Variant
    Dim i As Long
    For i = 1 To 10
        Set args(i) = ArgVal(i)
    Next i

    Dim v As Variant
    v = br.jsEval("function(a,b,c,d,e,f,g,h,i,j){ return a+b+c+d+e+f+g+h+i+j }", objectId:=oid, objectArguments:=args, returnByValue:=True, StopApiError:=False)
    AssertEq "1..10 の和", CDbl(v), 55#

    br.jsEval "updateStatus('s-js08','⑧ 完了 " & EOk() & " | 10 引数 → 和 55', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑨ シングルクォートを含む文字列（objectArguments）
'==============================================================================
Private Sub Test09_callFunctionOn_apostrophe_string(br As CDPContext)
    PrintSection "⑨ callFunctionOn - アポストロフィ文字列"

    Dim oid As String
    oid = CStr(br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, StopApiError:=False))

    Dim arg As Dictionary
    Set arg = ArgVal("It's " & "OK " & "日本語")

    Dim v As Variant
    v = br.jsEval("function(s){ return 'ECHO:' + s }", objectId:=oid, objectArguments:=Array(arg), returnByValue:=True, StopApiError:=False)
    AssertEq "エコー", CStr(v), "ECHO:It's OK 日本語"

    br.jsEval "updateStatus('s-js09','⑨ 完了 " & EOk() & " | objectArguments（引用符含む）エコー OK', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑩ 子要素参照（ネスト）
'==============================================================================
Private Sub Test10_callFunctionOn_nested(br As CDPContext)
    PrintSection "⑩ callFunctionOn - 子要素"

    Dim oid As String
    oid = CStr(br.jsEval("document.getElementById('jseval-box')", returnByValue:=False, StopApiError:=False))

    Dim v As Variant
    v = br.jsEval("function(){ return this.querySelector('#jseval-target').textContent }", objectId:=oid, returnByValue:=True, StopApiError:=False)
    AssertNotEmpty "子 span.textContent", CStr(v)

    v = br.jsEval("function(){ return this.querySelector('#jseval-input').value }", objectId:=oid, returnByValue:=True, StopApiError:=False)
    AssertEq "input.value 初期", CStr(v), "初期値"

    br.jsEval "updateStatus('s-js10','⑩ 完了 " & EOk() & " | 子 span / input', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑪ JS 例外（StopException 省略 = False → CVErr または Error 型）
'==============================================================================
Private Sub Test11_exception_stopException_off(br As CDPContext)
    PrintSection "⑪ 例外 - StopException=False"

    Dim r As Variant
    r = br.jsEval("(function(){ throw new Error('jsEval-test'); })()", StopException:=False, StopApiError:=False)

    If IsError(r) Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | 例外時 IsError(CVErr)"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | 例外時に Error 型でない: " & TypeName(r)
    End If

    br.jsEval "updateStatus('s-js11','⑪ 完了 " & EOk() & " | 例外 → IsError', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑫ IFEXCEPTION
'==============================================================================
Private Sub Test12_exception_IFEXCEPTION(br As CDPContext)
    PrintSection "⑫ 例外 - IFEXCEPTION"

    Dim r As Variant
    r = br.jsEval("(function(){ throw new Error('x'); })()", StopException:=False, IFEXCEPTION:="fallback-ok", StopApiError:=False)

    AssertEq "IFEXCEPTION 文字列", CStr(r), "fallback-ok"

    br.jsEval "updateStatus('s-js12','⑫ 完了 " & EOk() & " | IFEXCEPTION フォールバック', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑬ 長い文字列（返却値サイズ）
'==============================================================================
Private Sub Test13_long_string(br As CDPContext)
    PrintSection "⑬ 長い文字列 returnByValue"

    Dim v As Variant
    v = br.jsEval("'x'.repeat(2500)", returnByValue:=True, StopApiError:=False)

    If Len(CStr(v)) = 2500 Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | 長さ 2500"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | 長さ期待 2500 実際 " & Len(CStr(v))
    End If

    br.jsEval "updateStatus('s-js13','⑬ 完了 " & EOk() & " | 長さ 2500 文字', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑭ contextId ? Page.createIsolatedWorld の executionContextId で evaluate
'==============================================================================
Private Sub Test14_contextId_isolatedWorld(br As CDPContext)
    PrintSection "⑭ contextId ? createIsolatedWorld"

    On Error GoTo Test14_Err

    br.ExecuteCDP "Page.enable", Nothing

    Dim ftRes As BiDiCDPJson
    Set ftRes = br.ExecuteCDP("Page.getFrameTree", Nothing)

    Dim rootFrameId As String
    rootFrameId = CStr(ftRes("frameTree")("frame")("id"))

    Dim pCW As New Dictionary
    pCW.Add "frameId", rootFrameId
    pCW.Add "worldName", "jsEvalTestIsolated"

    Dim cwRes As BiDiCDPJson
    Set cwRes = br.ExecuteCDP("Page.createIsolatedWorld", pCW)

    Dim execCtx As Long
    execCtx = CLng(cwRes("executionContextId"))

    Dim v As Variant
    v = br.jsEval("window.__JSEVAL_ISO = 'ctx-ok'; window.__JSEVAL_ISO", contextId:=execCtx, returnByValue:=True, StopApiError:=False)
    AssertEq "isolated で代入→取得", CStr(v), "ctx-ok"

    Dim vMain As Variant
    vMain = br.jsEval("window.__JSEVAL_ISO", returnByValue:=True, StopApiError:=False)
    If IsEmpty(vMain) Or VarType(vMain) = vbNull Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | メイン context では __JSEVAL_ISO 未定義（Empty/Null）"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | メイン context に隔離値が見えている: " & CStr(vMain)
    End If

    br.jsEval "updateStatus('s-js14','⑭ 完了 " & EOk() & " | contextId=" & CStr(execCtx) & "', true)", StopApiError:=False
    Exit Sub

Test14_Err:
    failCount = failCount + 1
    Debug.Print "  FAIL | ⑭ " & Err.Description
    On Error Resume Next
    br.jsEval "updateStatus('s-js14','⑭ FAIL', false)", StopApiError:=False
End Sub

'==============================================================================
' ⑮ serializationOptions ? deep（deepSerializedValue 優先）
'==============================================================================
Private Sub Test15_serializationOptions_deep(br As CDPContext)
    PrintSection "⑮ serializationOptions ? deep"

    Dim serOpts As New Dictionary
    serOpts.Add "serialization", "deep"
    serOpts.Add "maxDepth", 8

    Dim resObj As Object
    Set resObj = br.jsEval("({ top: 1, nest: { mid: 2, deep: { leaf: 3 } } })", returnByValue:=True, serializationOptions:=serOpts, StopApiError:=False)

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

    br.jsEval "updateStatus('s-js15','⑮ 完了 " & EOk() & " | serialization=deep', true)", StopApiError:=False
End Sub

'==============================================================================
' ⑯ RunAsyncCDP ? alert（Demo_CDP.TestAlert と同系）
'==============================================================================
Private Sub Test16_RunAsyncCDP_alert(br As CDPContext)
    PrintSection "⑯ RunAsyncCDP ? alert"

    On Error GoTo Test16_Err

    br.ExecuteCDP "Page.enable", Nothing

    Dim oid As Variant
    oid = br.jsEval("document.getElementById('btn-async-alert')", returnByValue:=False, StopApiError:=False)

    If VarType(oid) <> vbString Or Len(oid) = 0 Then
        failCount = failCount + 1
        Debug.Print "  FAIL | ⑯ ボタン objectId 取得失敗"
        br.jsEval "updateStatus('s-js16','⑯ FAIL ボタンなし', false)", StopApiError:=False
        Exit Sub
    End If

    Dim asyncCmdId As Variant
    asyncCmdId = br.jsEval("function(){ this.click(); }", CStr(oid), RunAsyncCDP:=True, StopApiError:=False)

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
        br.InheritanceCDPBrowser.TakeEvents
        If br.BrowserEvents("EventMethods").Exists(evName) Then
            found = True
            Exit For
        End If
        br.InheritanceCDPBrowser.sleep 0.05
    Next i

    If found Then
        passCount = passCount + 1
        Debug.Print "  " & EOk() & " PASS | " & evName & " を検知"
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL | ダイアログイベントがタイムアウト"
    End If

    Dim pDlg As New Dictionary
    pDlg.Add "accept", True
    br.ExecuteCDP "Page.handleJavaScriptDialog", pDlg

    Set br.BrowserEvents = Nothing

    br.jsEval "updateStatus('s-js16','⑯ 完了 " & EOk() & " | Async+alert+handleDialog', true)", StopApiError:=False
    Exit Sub

Test16_Err:
    failCount = failCount + 1
    Debug.Print "  FAIL | ⑯ " & Err.Description
    On Error Resume Next
    Set br.BrowserEvents = Nothing
    br.jsEval "updateStatus('s-js16','⑯ FAIL', false)", StopApiError:=False
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
