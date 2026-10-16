Attribute VB_Name = "Test_JsEvalResult_CDP"
'==============================================================================
' JsEvalResult.cls 動作確認（CDP版）
' ・CDPContext.jsEval(returnByDetails:=True) で返る JsEvalResult を、
'   Runtime.RemoteObject の type / subtype の組み合わせごとに受け取り、`Stop`で止めます
' ・止まったら、ローカルウィンドウで`r`の中身を確認してください（イミディエイトにも要約を出力します）
' ・参考：https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#type-RemoteObject
'==============================================================================
Option Explicit

'==============================================================================
' エントリ
'==============================================================================
Public Sub RunAll_JsEvalResult_CDP()
    Dim br As CDPContext: Set br = ShSetting01_StartBrowser.StartCDPModeContext

    PrintHeader "JsEvalResult(CDP) type / subtype 網羅確認 開始"

    Test01_Primitive br
    Test02_NumberUnserializable br
    Test03_NonObjectSpecial br
    Test04_ObjectSubtypes br
    Test05_ReturnByValueAndDeep br
    Test06_CallFunctionOn br
    Test07_Exception br

    PrintHeader "JsEvalResult(CDP) 網羅確認 終了"

    br.ThisCDPBrowser.CloseBrowser
End Sub

'==============================================================================
' ① プリミティブ型（type: string / number / boolean / undefined / bigint）
'==============================================================================
Private Sub Test01_Primitive(br As CDPContext)
    PrintSection "① プリミティブ型"
    Dim r As JsEvalResult

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

    'type: undefined
    Set r = br.jsEval("undefined", returnByDetails:=True)
    PrintSummary "undefined", r
    Stop

    'type: bigint（unserializableValue で返る）
    Set r = br.jsEval("12345678901234567890n", returnByDetails:=True)
    PrintSummary "bigint", r
    Stop
End Sub

'==============================================================================
' ② number の unserializableValue（NaN / Infinity / -Infinity / -0）
'    ※subtypeではないが、value分析の分岐が変わるため個別に確認
'==============================================================================
Private Sub Test02_NumberUnserializable(br As CDPContext)
    PrintSection "② number（unserializableValue）"
    Dim r As JsEvalResult

    Set r = br.jsEval("NaN", returnByDetails:=True)
    PrintSummary "number(NaN)", r
    Stop

    Set r = br.jsEval("Infinity", returnByDetails:=True)
    PrintSummary "number(Infinity)", r
    Stop

    Set r = br.jsEval("-Infinity", returnByDetails:=True)
    PrintSummary "number(-Infinity)", r
    Stop

    Set r = br.jsEval("-0", returnByDetails:=True)
    PrintSummary "number(-0)", r
    Stop
End Sub

'==============================================================================
' ③ object以外の参照型（type: symbol / function）
'==============================================================================
Private Sub Test03_NonObjectSpecial(br As CDPContext)
    PrintSection "③ symbol / function"
    Dim r As JsEvalResult

    'type: symbol
    Set r = br.jsEval("Symbol('vba')", returnByDetails:=True)
    PrintSummary "symbol", r
    Stop

    'type: function
    Set r = br.jsEval("(function add(a, b){ return a + b })", returnByDetails:=True)
    PrintSummary "function", r
    Stop
End Sub

'==============================================================================
' ④ type: object の subtype 網羅（returnByValue:=False 既定 → objectId で返る）
'==============================================================================
Private Sub Test04_ObjectSubtypes(br As CDPContext)
    PrintSection "④ object / subtype"
    Dim r As JsEvalResult

    'subtype: なし（素のオブジェクト）
    Set r = br.jsEval("({ a: 1, b: 'x' })", returnByDetails:=True)
    PrintSummary "object(subtypeなし)", r
    Stop

    'subtype: array
    Set r = br.jsEval("[1, 'two', true]", returnByDetails:=True)
    PrintSummary "object/array", r
    Stop

    'subtype: null（objectIdは付かず、value分析側を通る）
    Set r = br.jsEval("null", returnByDetails:=True)
    PrintSummary "object/null", r
    Stop

    'subtype: node
    Set r = br.jsEval("document.body", returnByDetails:=True)
    PrintSummary "object/node", r
    Stop

    'subtype: regexp
    Set r = br.jsEval("/ab+c/gi", returnByDetails:=True)
    PrintSummary "object/regexp", r
    Stop

    'subtype: date
    Set r = br.jsEval("new Date(2026, 9, 15, 12, 34, 56)", returnByDetails:=True)
    PrintSummary "object/date", r
    Stop

    'subtype: map
    Set r = br.jsEval("new Map([['k', 1]])", returnByDetails:=True)
    PrintSummary "object/map", r
    Stop

    'subtype: set
    Set r = br.jsEval("new Set([1, 2])", returnByDetails:=True)
    PrintSummary "object/set", r
    Stop

    'subtype: weakmap
    Set r = br.jsEval("new WeakMap()", returnByDetails:=True)
    PrintSummary "object/weakmap", r
    Stop

    'subtype: weakset
    Set r = br.jsEval("new WeakSet()", returnByDetails:=True)
    PrintSummary "object/weakset", r
    Stop

    'subtype: iterator（Map/Setのイテレーター）
    '※`Array Iterator`は、V8上ではsubtypeが付かない(素のobject扱い)ため、Mapの方で確認
    Set r = br.jsEval("new Map([['k', 1]]).entries()", returnByDetails:=True)
    PrintSummary "object/iterator", r
    Stop

    '参考：Array Iterator（subtypeなしになることの確認）
    Set r = br.jsEval("[1, 2][Symbol.iterator]()", returnByDetails:=True)
    PrintSummary "object(Array Iterator)", r
    Stop

    'subtype: generator
    Set r = br.jsEval("(function* (){ yield 1 })()", returnByDetails:=True)
    PrintSummary "object/generator", r
    Stop

    'subtype: error（例外ではなく、Errorオブジェクトを値として返す）
    Set r = br.jsEval("new TypeError('vba error')", returnByDetails:=True)
    PrintSummary "object/error", r
    Stop

    'subtype: proxy
    Set r = br.jsEval("new Proxy({}, {})", returnByDetails:=True)
    PrintSummary "object/proxy", r
    Stop

    'subtype: promise（awaitPromise:=False なので、Promiseそのもの）
    Set r = br.jsEval("Promise.resolve(1)", returnByDetails:=True)
    PrintSummary "object/promise", r
    Stop

    'subtype: typedarray
    Set r = br.jsEval("new Uint8Array([1, 2, 3])", returnByDetails:=True)
    PrintSummary "object/typedarray", r
    Stop

    'subtype: arraybuffer
    Set r = br.jsEval("new ArrayBuffer(8)", returnByDetails:=True)
    PrintSummary "object/arraybuffer", r
    Stop

    'subtype: dataview
    Set r = br.jsEval("new DataView(new ArrayBuffer(8))", returnByDetails:=True)
    PrintSummary "object/dataview", r
    Stop

    'subtype: webassemblymemory
    Set r = br.jsEval("new WebAssembly.Memory({ initial: 1 })", returnByDetails:=True)
    PrintSummary "object/webassemblymemory", r
    Stop

    'subtype: trustedtype（TrustedHTML）
    '※ポリシー名の重複を避けるため、時刻を付与
    Set r = br.jsEval("trustedTypes.createPolicy('vba' + Date.now(), { createHTML: s => s }).createHTML('<b>x</b>')", returnByDetails:=True)
    PrintSummary "object/trustedtype", r
    Stop

    'subtype: wasmvalue
    '※WebAssemblyのデバッグ(Debugger.evaluateOnCallFrame)時のみ出現するため、jsEvalでは再現できません
End Sub

'==============================================================================
' ⑤ returnByValue:=True / serializationOptions(deep) / awaitPromise:=True
'    ※objectId を持たない場合の value 分析・deepSerializedValue 分岐の確認
'==============================================================================
Private Sub Test05_ReturnByValueAndDeep(br As CDPContext)
    PrintSection "⑤ returnByValue / deep / awaitPromise"
    Dim r As JsEvalResult

    'object(subtypeなし) + returnByValue:=True → value(Node)で返る
    Set r = br.jsEval("({ a: 1, b: 'x' })", returnByValue:=True, returnByDetails:=True)
    PrintSummary "object(returnByValue)", r
    Stop

    'object/array + returnByValue:=True
    Set r = br.jsEval("[1, 'two', true]", returnByValue:=True, returnByDetails:=True)
    PrintSummary "object/array(returnByValue)", r
    Stop

    'object/date + returnByValue:=True（JSON化されると、どう返るか）
    Set r = br.jsEval("new Date(2026, 9, 15, 12, 34, 56)", returnByValue:=True, returnByDetails:=True)
    PrintSummary "object/date(returnByValue)", r
    Stop

    'function + returnByValue:=True（{"type":"function","value":{}} になる想定）
    Set r = br.jsEval("(function add(a, b){ return a + b })", returnByValue:=True, returnByDetails:=True)
    PrintSummary "function(returnByValue)", r
    Stop

    'serializationOptions(deep) → deepSerializedValue が最優先で返る
    Dim deepOpt As New Dictionary
    deepOpt.Add "serialization", "deep"
    Set r = br.jsEval("({ a: 1, list: [1, 2], when: new Date(2026, 9, 15), m: new Map([['k', 1]]) })", serializationOptions:=deepOpt, returnByDetails:=True)
    PrintSummary "object(deep)", r
    Stop

    'serializationOptions(deep) + node
    Set r = br.jsEval("document.body", serializationOptions:=deepOpt, returnByDetails:=True)
    PrintSummary "object/node(deep)", r
    Stop

    'promise + awaitPromise:=True → 解決後の値(number)で返る
    Set r = br.jsEval("new Promise(res => setTimeout(() => res(7), 100))", awaitPromise:=True, returnByDetails:=True)
    PrintSummary "promise(awaitPromise)→number", r
    Stop
End Sub

'==============================================================================
' ⑥ Runtime.callFunctionOn 経由（objectId / objectArguments 指定時）
'==============================================================================
Private Sub Test06_CallFunctionOn(br As CDPContext)
    PrintSection "⑥ Runtime.callFunctionOn"
    Dim r As JsEvalResult

    'objectId を this にして実行
    Dim bodyId As String: bodyId = br.jsEval("document.body")
    Set r = br.jsEval("function(){ return this.tagName }", objectId:=bodyId, returnByDetails:=True)
    PrintSummary "callFunctionOn(objectId)→string", r
    Stop

    'objectArguments のみ指定（this なし）
    Dim args(0) As Dictionary
    Set args(0) = New Dictionary: args(0).Add "value", 20
    Set r = br.jsEval("function(n){ return { doubled: n * 2 } }", objectArguments:=args, returnByDetails:=True)
    PrintSummary "callFunctionOn(objectArguments)→object", r
    Stop
End Sub

'==============================================================================
' ⑦ 例外（exceptionDetails）
'    ※詳細モードでは、例外時も`JsEvalResult`(type=exception)で返ります
'==============================================================================
Private Sub Test07_Exception(br As CDPContext)
    PrintSection "⑦ 例外"
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
'* JsEvalResult の主要プロパティをイミディエイトに要約出力します
'  ※`rawValue`等、キーがない場合にエラーとなり得るものは出しません。ローカルウィンドウで確認してください
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
    Debug.Print "     type=" & r.ResultType & " / subtype=" & r.ResultSubtype & " / VarTypeVBA=" & r.ResultTypeForVBA
    Debug.Print "     value=" & valueText
    Debug.Print "     className=" & r.ClassName & " / description=" & r.Description
    Debug.Print "     objectId=" & r.objectId & " / unserializableValue=" & r.UnserializableValue & " / isException=" & r.isException
End Sub

'* 例外情報(exceptionDetails)と、投げられた値(ExceptionResultObject)を要約出力します
Private Sub PrintException(r As JsEvalResult)
    If Not r.isException Then Exit Sub

    Debug.Print "     exceptionText=" & r.ExceptionText & " / line=" & r.ExceptionLineNumber & " / column=" & r.ExceptionColumnNumber

    Dim thrown As JsEvalResult: Set thrown = r.ExceptionResultObject
    If thrown Is Nothing Then Debug.Print "     thrown=(Nothing)": Exit Sub
    Debug.Print "     thrown: type=" & thrown.ResultType & " / subtype=" & thrown.ResultSubtype & " / className=" & thrown.ClassName & " / description=" & Replace(thrown.Description, vbLf, " ")
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
