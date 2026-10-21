Attribute VB_Name = "Test_JsEvalResult_BiDi"
'==============================================================================
' JsEvalResult.cls 動作確認（BiDi版）
' ・WebDriverBiDiContext.jsEval(returnByDetails:=True) で返る JsEvalResult を、
'   script.RemoteValue の type ごとに受け取り、`Stop`で止めます
' ・止まったら、ローカルウィンドウで`r`の中身を確認してください（イミディエイトにも要約を出力します）
' ・BiDiにはCDPのsubtypeがない代わりに、`Ownership`(resultOwnership)で返り方が変わるため、
'   ・Ownership:=False … value で返る（④まで）
'   ・Ownership:=True  … handle / sharedId で返る（⑥）
'   の両方を確認します
' ・参考：https://w3c.github.io/webdriver-bidi/#type-script-RemoteValue
'==============================================================================
Option Explicit

'==============================================================================
' エントリ
'==============================================================================
Public Sub RunAll_JsEvalResult_BiDi()
    Dim br As WebDriverBiDiContext: Set br = ShSetting01_StartBrowser.StartBiDiModeContext

    PrintHeader "JsEvalResult(BiDi) type 網羅確認 開始"

    Test01_Primitive br
    Test02_NumberSpecial br
    Test03_NoValueTypes br
    Test04_ValueNodeTypes br
    Test05_NonSpecTypes br
    Test06_Ownership br
    Test07_SerializationOptions br
    Test08_AwaitPromise br
    Test09_CallFunction br
    Test10_Exception br

    PrintHeader "JsEvalResult(BiDi) 網羅確認 終了"

    br.ThisWebDriverBiDiMode.CloseBrowser
End Sub

'==============================================================================
' ① プリミティブ型（undefined / null / string / number / boolean / bigint）
'==============================================================================
Private Sub Test01_Primitive(br As WebDriverBiDiContext)
    PrintSection "① プリミティブ型"
    Dim r As JsEvalResult

    'type: undefined
    Set r = br.jsEval("undefined", returnByDetails:=True)
    PrintSummary "undefined", r
    Stop

    'type: null（BiDiは、CDPと違い独立した型）
    Set r = br.jsEval("null", returnByDetails:=True)
    PrintSummary "null", r
    Stop

    'type: string
    Set r = br.jsEval("'こんにちは VBA'", returnByDetails:=True)
    PrintSummary "string", r
    Stop

    'type: number
    Set r = br.jsEval("42.5", returnByDetails:=True)
    PrintSummary "number", r
    Stop

    'type: boolean
    Set r = br.jsEval("true", returnByDetails:=True)
    PrintSummary "boolean", r
    Stop

    'type: bigint（BiDiは、valueに文字列で入る）
    Set r = br.jsEval("12345678901234567890n", returnByDetails:=True)
    PrintSummary "bigint", r
    Stop
End Sub

'==============================================================================
' ② number の特殊値（NaN / -0 / Infinity / -Infinity）
'    ※BiDiは、valueに文字列で入るため、#NUM! 判定の分岐を確認
'==============================================================================
Private Sub Test02_NumberSpecial(br As WebDriverBiDiContext)
    PrintSection "② number（特殊値）"
    Dim r As JsEvalResult

    Set r = br.jsEval("NaN", returnByDetails:=True)
    PrintSummary "number(NaN)", r
    Stop

    Set r = br.jsEval("-0", returnByDetails:=True)
    PrintSummary "number(-0)", r
    Stop

    Set r = br.jsEval("Infinity", returnByDetails:=True)
    PrintSummary "number(Infinity)", r
    Stop

    Set r = br.jsEval("-Infinity", returnByDetails:=True)
    PrintSummary "number(-Infinity)", r
    Stop
End Sub

'==============================================================================
' ③ value を持たない型（Ownership:=False → value は Nothing）
'==============================================================================
Private Sub Test03_NoValueTypes(br As WebDriverBiDiContext)
    PrintSection "③ value を持たない型"
    Dim r As JsEvalResult

    'type: symbol
    Set r = br.jsEval("Symbol('vba')", returnByDetails:=True)
    PrintSummary "symbol", r
    Stop

    'type: function
    Set r = br.jsEval("(function add(a, b){ return a + b })", returnByDetails:=True)
    PrintSummary "function", r
    Stop

    'type: weakmap
    Set r = br.jsEval("new WeakMap()", returnByDetails:=True)
    PrintSummary "weakmap", r
    Stop

    'type: weakset
    Set r = br.jsEval("new WeakSet()", returnByDetails:=True)
    PrintSummary "weakset", r
    Stop

    'type: generator
    Set r = br.jsEval("(function* (){ yield 1 })()", returnByDetails:=True)
    PrintSummary "generator", r
    Stop

    'type: error（例外ではなく、Errorオブジェクトを値として返す）
    Set r = br.jsEval("new TypeError('vba error')", returnByDetails:=True)
    PrintSummary "error", r
    Stop

    'type: proxy
    Set r = br.jsEval("new Proxy({}, {})", returnByDetails:=True)
    PrintSummary "proxy", r
    Stop

    'type: promise（awaitPromise:=False なので、Promiseそのもの）
    Set r = br.jsEval("Promise.resolve(1)", returnByDetails:=True)
    PrintSummary "promise", r
    Stop

    'type: typedarray
    Set r = br.jsEval("new Uint8Array([1, 2, 3])", returnByDetails:=True)
    PrintSummary "typedarray", r
    Stop

    'type: arraybuffer
    Set r = br.jsEval("new ArrayBuffer(8)", returnByDetails:=True)
    PrintSummary "arraybuffer", r
    Stop
End Sub

'==============================================================================
' ④ value をNode(BiDiCDPJson)で持つ型 / date（Ownership:=False）
'==============================================================================
Private Sub Test04_ValueNodeTypes(br As WebDriverBiDiContext)
    PrintSection "④ value を持つ型"
    Dim r As JsEvalResult

    'type: array
    Set r = br.jsEval("[1, 'two', true]", returnByDetails:=True)
    PrintSummary "array", r
    Stop

    'type: object（value は [[key, RemoteValue], ...] 形式）
    Set r = br.jsEval("({ a: 1, b: 'x' })", returnByDetails:=True)
    PrintSummary "object", r
    Stop

    'type: map
    Set r = br.jsEval("new Map([['k', 1]])", returnByDetails:=True)
    PrintSummary "map", r
    Stop

    'type: set
    Set r = br.jsEval("new Set([1, 2])", returnByDetails:=True)
    PrintSummary "set", r
    Stop

    'type: regexp（value は { pattern, flags }）
    Set r = br.jsEval("/ab+c/gi", returnByDetails:=True)
    PrintSummary "regexp", r
    Stop

    'type: date（ISO文字列 → VBAのDate型に変換される想定）
    Set r = br.jsEval("new Date(Date.UTC(2026, 9, 15, 3, 34, 56))", returnByDetails:=True)
    PrintSummary "date", r
    Stop

    'type: node（Ownership:=False でも sharedId は付く）
    Set r = br.jsEval("document.body", returnByDetails:=True)
    PrintSummary "node", r
    Stop

    'type: nodelist
    Set r = br.jsEval("document.querySelectorAll('*')", returnByDetails:=True)
    PrintSummary "nodelist", r
    Stop

    'type: htmlcollection
    Set r = br.jsEval("document.getElementsByTagName('*')", returnByDetails:=True)
    PrintSummary "htmlcollection", r
    Stop

    'type: window（value は { context }）
    Set r = br.jsEval("window", returnByDetails:=True)
    PrintSummary "window", r
    Stop
End Sub

'==============================================================================
' ⑤ 参考：BiDi仕様に無い型（CDPではsubtypeがあったもの）
'    ※マッパー側で何の型に寄せられるかの確認用。未知の型の場合は、JsEvalResult側でエラー停止する作りのため、エラー内容も出力します
'==============================================================================
Private Sub Test05_NonSpecTypes(br As WebDriverBiDiContext)
    PrintSection "⑤ 参考：BiDi仕様外の型"

    TryNonSpec br, "iterator(Map)", "new Map([['k', 1]]).entries()"
    Stop

    TryNonSpec br, "dataview", "new DataView(new ArrayBuffer(8))"
    Stop

    TryNonSpec br, "webassemblymemory", "new WebAssembly.Memory({ initial: 1 })"
    Stop

    TryNonSpec br, "trustedtype", "trustedTypes.createPolicy('vba' + Date.now(), { createHTML: s => s }).createHTML('<b>x</b>')"
    Stop
End Sub

'==============================================================================
' ⑥ Ownership:=True（resultOwnership:"root"）→ handle で返る（nodeは sharedId が優先）
'==============================================================================
Private Sub Test06_Ownership(br As WebDriverBiDiContext)
    PrintSection "⑥ Ownership:=True"
    Dim r As JsEvalResult

    'object → handle
    Set r = br.jsEval("({ a: 1, b: 'x' })", Ownership:=True, returnByDetails:=True)
    PrintSummary "object(Ownership)", r
    Stop

    'array → handle
    Set r = br.jsEval("[1, 'two', true]", Ownership:=True, returnByDetails:=True)
    PrintSummary "array(Ownership)", r
    Stop

    'node → handle と sharedId の両方（value は sharedId が優先）
    Set r = br.jsEval("document.body", Ownership:=True, returnByDetails:=True)
    PrintSummary "node(Ownership)", r
    Stop

    'nodelist → handle
    Set r = br.jsEval("document.querySelectorAll('*')", Ownership:=True, returnByDetails:=True)
    PrintSummary "nodelist(Ownership)", r
    Stop

    'function → handle
    Set r = br.jsEval("(function add(a, b){ return a + b })", Ownership:=True, returnByDetails:=True)
    PrintSummary "function(Ownership)", r
    Stop

    'promise → handle
    Set r = br.jsEval("Promise.resolve(1)", Ownership:=True, returnByDetails:=True)
    PrintSummary "promise(Ownership)", r
    Stop

    'プリミティブは Ownership:=True でも handle は付かない
    Set r = br.jsEval("'primitive'", Ownership:=True, returnByDetails:=True)
    PrintSummary "string(Ownership)", r
    Stop
End Sub

'==============================================================================
' ⑦ serializationOptions（maxObjectDepth / maxDomDepth / includeShadowTree）
'==============================================================================
Private Sub Test07_SerializationOptions(br As WebDriverBiDiContext)
    PrintSection "⑦ serializationOptions"
    Dim r As JsEvalResult
    Dim opt As Dictionary

    'maxObjectDepth:=0 → ネストしたobjectの中身が省略される
    Set opt = New Dictionary: opt.Add "maxObjectDepth", 0
    Set r = br.jsEval("({ a: 1, nest: { b: 2 } })", serializationOptions:=opt, returnByDetails:=True)
    PrintSummary "object(maxObjectDepth=0)", r
    Stop

    'maxDomDepth:=1 → nodeのchildrenまで含まれる
    Set opt = New Dictionary: opt.Add "maxDomDepth", 1
    Set r = br.jsEval("document.body", serializationOptions:=opt, returnByDetails:=True)
    PrintSummary "node(maxDomDepth=1)", r
    Stop

    'includeShadowTree:="all" → closed の Shadow-root も含まれる
    Set opt = New Dictionary: opt.Add "maxDomDepth", 1: opt.Add "includeShadowTree", "all"
    Set r = br.jsEval("(() => { const h = document.createElement('div'); h.attachShadow({ mode: 'closed' }).innerHTML = '<b>x</b>'; return h })()", serializationOptions:=opt, returnByDetails:=True)
    PrintSummary "node(includeShadowTree=all, closed)", r
    Stop
End Sub

'==============================================================================
' ⑧ awaitPromise:=True
'==============================================================================
Private Sub Test08_AwaitPromise(br As WebDriverBiDiContext)
    PrintSection "⑧ awaitPromise"
    Dim r As JsEvalResult

    Set r = br.jsEval("new Promise(res => setTimeout(() => res(7), 100))", awaitPromise:=True, returnByDetails:=True)
    PrintSummary "promise(awaitPromise)→number", r
    Stop
End Sub

'==============================================================================
' ⑨ script.callFunction 経由（scriptSharedId / objectArguments）
'==============================================================================
Private Sub Test09_CallFunction(br As WebDriverBiDiContext)
    PrintSection "⑨ script.callFunction"
    Dim r As JsEvalResult

    'scriptSharedId を this にして実行（簡易モードで得た sharedId を流用）
    Dim bodySharedId As String: bodySharedId = br.jsEval("document.body")
    Set r = br.jsEval("function(){ return this.tagName }", scriptSharedId:=bodySharedId, returnByDetails:=True)
    PrintSummary "callFunction(scriptSharedId)→string", r
    Stop

    'objectArguments（script.LocalValue形式）のみ指定
    Dim args(0) As Dictionary
    Set args(0) = New Dictionary: args(0).Add "type", "number": args(0).Add "value", 20
    Set r = br.jsEval("function(n){ return { doubled: n * 2 } }", objectArguments:=args, returnByDetails:=True)
    PrintSummary "callFunction(objectArguments)→object", r
    Stop

    'sharedId を引数(SharedReference)として渡す
    Dim nodeResult As JsEvalResult: Set nodeResult = br.jsEval("document.body", returnByDetails:=True)
    Dim refArgs(0) As Dictionary
    Set refArgs(0) = New Dictionary: refArgs(0).Add "sharedId", nodeResult.sharedId
    Set r = br.jsEval("function(el){ return el.tagName }", objectArguments:=refArgs, returnByDetails:=True)
    PrintSummary "callFunction(sharedId引数)→string", r
    Stop
End Sub

'==============================================================================
' ⑩ 例外（exceptionDetails）
'    ※詳細モードでは、例外時も`JsEvalResult`(type=exception)で返ります
'==============================================================================
Private Sub Test10_Exception(br As WebDriverBiDiContext)
    PrintSection "⑩ 例外"
    Dim r As JsEvalResult

    'throw new Error
    Set r = br.jsEval("throw new Error('vba exception')", returnByDetails:=True)
    PrintSummary "exception(Error)", r
    PrintException r
    Stop

    'throw プリミティブ
    Set r = br.jsEval("throw 'just a string'", returnByDetails:=True)
    PrintSummary "exception(string)", r
    PrintException r
    Stop

    'Promise の reject（awaitPromise:=True）
    Set r = br.jsEval("Promise.reject(new RangeError('rejected'))", awaitPromise:=True, returnByDetails:=True)
    PrintSummary "exception(reject)", r
    PrintException r
    Stop

    '構文エラー
    Set r = br.jsEval("function (", returnByDetails:=True)
    PrintSummary "exception(SyntaxError)", r
    PrintException r
    Stop

    '既定メンバー(value)経由のLet代入でも、従来通りIsErrorで判定できること
    Dim v As Variant
    v = br.jsEval("throw new Error('let')", returnByDetails:=True)
    Debug.Print "  [exception(Let代入)] TypeName=" & TypeName(v) & " / IsError=" & IsError(v)
    Stop

    '詳細モード + IFEXCEPTION 指定時：JsEvalResult で返り、value(既定メンバー)が代替値になること
    Set r = br.jsEval("throw new Error('ifexception')", IFEXCEPTION:="代替値", returnByDetails:=True)
    PrintSummary "exception(IFEXCEPTION:文字列)", r
    PrintException r
    Stop

    '詳細モード + IFEXCEPTION にオブジェクトを指定：value が Set で返ること
    Dim altObj As New Collection: altObj.Add "alt"
    Set r = br.jsEval("throw new Error('ifexception-obj')", IFEXCEPTION:=altObj, returnByDetails:=True)
    PrintSummary "exception(IFEXCEPTION:オブジェクト)", r
    Debug.Print "     value Is altObj=" & (r.value Is altObj)
    Stop
End Sub

'==============================================================================
' ヘルパー
'==============================================================================
'* BiDi仕様外の型を、エラーを捕まえつつ実行します
Private Sub TryNonSpec(br As WebDriverBiDiContext, caseName As String, js As String)
    Dim r As JsEvalResult
    On Error Resume Next
    Set r = br.jsEval(js, returnByDetails:=True)
    If Err.Number <> 0 Then
        Debug.Print "  [" & caseName & "] ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    PrintSummary caseName, r
End Sub

'* JsEvalResult の主要プロパティをイミディエイトに要約出力します
'  ※BiDiの例外時は結果本体がないため、handle / sharedId 等は出しません
Private Sub PrintSummary(caseName As String, r As JsEvalResult)
    Dim valueText As String
    If r.ResultTypeForVBA = vbObject Then
        valueText = "[" & TypeName(r.value) & "]"
    ElseIf IsError(r.value) Then
        valueText = "[" & CStr(r.value) & "]"
    ElseIf IsNull(r.value) Then
        valueText = "[Null]"
    ElseIf IsEmpty(r.value) Then
        valueText = "[Empty]"
    Else
        valueText = CStr(r.value)
    End If

    Debug.Print "  [" & caseName & "]"
    Debug.Print "     type=" & r.ResultType & " / VarTypeVBA=" & r.ResultTypeForVBA & " / isException=" & r.isException
    Debug.Print "     value=" & valueText
    If Not r.isException Then Debug.Print "     handle=" & r.handle & " / sharedId=" & r.sharedId & " / internalId=" & r.internalId
End Sub

'* 例外情報(exceptionDetails)と、投げられた値(ExceptionResultObject)を要約出力します
Private Sub PrintException(r As JsEvalResult)
    If Not r.isException Then Exit Sub

    Debug.Print "     exceptionText=" & r.ExceptionText & " / line=" & r.ExceptionLineNumber & " / column=" & r.ExceptionColumnNumber

    Dim thrown As JsEvalResult: Set thrown = r.ExceptionResultObject
    If thrown Is Nothing Then Debug.Print "     thrown=(Nothing)": Exit Sub
    Dim thrownValue As String
    If thrown.ResultTypeForVBA = vbObject Then thrownValue = "[" & TypeName(thrown.value) & "]" Else thrownValue = CStr(thrown.value)
    Debug.Print "     thrown: type=" & thrown.ResultType & " / value=" & thrownValue
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
