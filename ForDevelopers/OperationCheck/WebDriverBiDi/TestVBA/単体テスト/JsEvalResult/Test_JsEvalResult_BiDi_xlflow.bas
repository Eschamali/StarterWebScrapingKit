Attribute VB_Name = "Test_JsEvalResult_BiDi_xlflow"
'==============================================================================
' JsEvalResult.cls 自動テスト（BiDi版 / xlflow test 用）
' ・Test_JsEvalResult_BiDi（Stopで目視確認する版）と同じケースを、期待値で自動判定します
' ・BeforeAll でBiDiブラウザを起動し、AfterAll で閉じます
' ・実行例：xlflow test --module Test_JsEvalResult_BiDi_xlflow --session --json
'==============================================================================
Option Explicit

Private br As WebDriverBiDiContext



'==============================================================================
' ライフサイクル
'==============================================================================
Public Sub BeforeAll()
    Set br = ShSetting01_StartBrowser.StartBiDiModeContext
End Sub

Public Sub AfterAll()
    If Not br Is Nothing Then br.ThisWebDriverBiDiMode.CloseBrowser
    Set br = Nothing
End Sub



'==============================================================================
' ① プリミティブ型
'==============================================================================
Public Sub Test_Primitive_Undefined()
    Dim r As JsEvalResult: Set r = Details("undefined")
    CheckType r, "undefined", vbEmpty
    AssertSame Empty, r.value, "value"
End Sub

Public Sub Test_Primitive_Null()
    Dim r As JsEvalResult: Set r = Details("null")
    CheckType r, "null", vbNull
    AssertSame Null, r.value, "value"
End Sub

Public Sub Test_Primitive_String()
    Dim r As JsEvalResult: Set r = Details("'hello VBA'")
    CheckType r, "string", vbString
    AssertSame "hello VBA", r.value, "value"
    AssertSame "", r.handle, "handle"
End Sub

Public Sub Test_Primitive_Number()
    Dim r As JsEvalResult: Set r = Details("42.5")
    CheckType r, "number", vbDouble
    AssertSame 42.5, r.value, "value"
End Sub

Public Sub Test_Primitive_Boolean()
    Dim r As JsEvalResult: Set r = Details("true")
    CheckType r, "boolean", vbBoolean
    AssertSame True, r.value, "value"
End Sub

Public Sub Test_Primitive_BigInt()
    'BiDiは`n`なしの数字文字列で返る
    Dim r As JsEvalResult: Set r = Details("12345678901234567890n")
    CheckType r, "bigint", vbString
    AssertSame "12345678901234567890", r.value, "value"
End Sub



'==============================================================================
' ② number の特殊値 → #NUM!
'==============================================================================
Public Sub Test_Number_NaN()
    CheckSpecialNumber "NaN"
End Sub

Public Sub Test_Number_MinusZero()
    CheckSpecialNumber "-0"
End Sub

Public Sub Test_Number_Infinity()
    CheckSpecialNumber "Infinity"
End Sub

Public Sub Test_Number_MinusInfinity()
    CheckSpecialNumber "-Infinity"
End Sub



'==============================================================================
' ③ value を持たない型（Ownership:=False → value は Nothing）
'==============================================================================
Public Sub Test_NoValue_Symbol()
    CheckNoValue "symbol", "Symbol('vba')"
End Sub

Public Sub Test_NoValue_Function()
    CheckNoValue "function", "(function add(a, b){ return a + b })"
End Sub

Public Sub Test_NoValue_WeakMap()
    CheckNoValue "weakmap", "new WeakMap()"
End Sub

Public Sub Test_NoValue_WeakSet()
    CheckNoValue "weakset", "new WeakSet()"
End Sub

Public Sub Test_NoValue_Generator()
    CheckNoValue "generator", "(function* (){ yield 1 })()"
End Sub

Public Sub Test_NoValue_Error()
    CheckNoValue "error", "new TypeError('vba error')"
End Sub

Public Sub Test_NoValue_Proxy()
    CheckNoValue "proxy", "new Proxy({}, {})"
End Sub

Public Sub Test_NoValue_Promise()
    CheckNoValue "promise", "Promise.resolve(1)"
End Sub

Public Sub Test_NoValue_TypedArray()
    CheckNoValue "typedarray", "new Uint8Array([1, 2, 3])"
End Sub

Public Sub Test_NoValue_ArrayBuffer()
    CheckNoValue "arraybuffer", "new ArrayBuffer(8)"
End Sub



'==============================================================================
' ④ value を持つ型（Ownership:=False）
'==============================================================================
Public Sub Test_Value_Array()
    CheckJsonValue "array", "[1, 'two', true]"
End Sub

Public Sub Test_Value_Object()
    CheckJsonValue "object", "({ a: 1, b: 'x' })"
End Sub

Public Sub Test_Value_Map()
    CheckJsonValue "map", "new Map([['k', 1]])"
End Sub

Public Sub Test_Value_Set()
    CheckJsonValue "set", "new Set([1, 2])"
End Sub

Public Sub Test_Value_RegExp()
    CheckJsonValue "regexp", "/ab+c/gi"
End Sub

Public Sub Test_Value_Window()
    CheckJsonValue "window", "window"
End Sub

Public Sub Test_Value_NodeList()
    Dim r As JsEvalResult: Set r = Details("document.querySelectorAll('*')")
    CheckType r, "nodelist", vbObject
End Sub

Public Sub Test_Value_HTMLCollection()
    Dim r As JsEvalResult: Set r = Details("document.getElementsByTagName('*')")
    CheckType r, "htmlcollection", vbObject
End Sub

Public Sub Test_Value_Date()
    'UTCで作り、UtcConverter.ParseIsoと同じ値(ローカル時刻)になること
    Dim r As JsEvalResult: Set r = Details("new Date(Date.UTC(2026, 9, 15, 3, 34, 56))")
    CheckType r, "date", vbDate
    AssertSame UtcConverter.ParseIso("2026-10-15T03:34:56.000Z"), r.value, "value"
End Sub

Public Sub Test_Value_Node()
    'nodeは Ownership:=False だと、value はNode(BiDiCDPJson)。sharedId はプロパティから取得できる
    Dim r As JsEvalResult: Set r = Details("document.body")
    CheckType r, "node", vbObject
    AssertTrue Len(r.sharedId) > 0, "sharedIdがあること"
    AssertSame "", r.handle, "handle"
    AssertSame "BiDiCDPJson", TypeName(r.value), "TypeName(value)"
End Sub



'==============================================================================
' ⑤ 参考：BiDi仕様に無い型（エラーにならず、何らかのtypeで返ること）
'==============================================================================
Public Sub Test_NonSpec_MapIterator()
    CheckNonSpec "new Map([['k', 1]]).entries()"
End Sub

Public Sub Test_NonSpec_DataView()
    CheckNonSpec "new DataView(new ArrayBuffer(8))"
End Sub

Public Sub Test_NonSpec_WebAssemblyMemory()
    CheckNonSpec "new WebAssembly.Memory({ initial: 1 })"
End Sub

Public Sub Test_NonSpec_TrustedType()
    CheckNonSpec "trustedTypes.createPolicy('vba' + Date.now(), { createHTML: s => s }).createHTML('<b>x</b>')"
End Sub



'==============================================================================
' ⑥ Ownership:=True → handle で返る
'==============================================================================
Public Sub Test_Ownership_Object()
    CheckHandle "object", "({ a: 1, b: 'x' })"
End Sub

Public Sub Test_Ownership_Array()
    CheckHandle "array", "[1, 'two', true]"
End Sub

Public Sub Test_Ownership_NodeList()
    CheckHandle "nodelist", "document.querySelectorAll('*')"
End Sub

Public Sub Test_Ownership_Function()
    CheckHandle "function", "(function add(a, b){ return a + b })"
End Sub

Public Sub Test_Ownership_Promise()
    CheckHandle "promise", "Promise.resolve(1)"
End Sub

Public Sub Test_Ownership_Node_HandlePreferred()
    'nodeは handle と sharedId の両方が付くが、value は handle が優先
    Dim r As JsEvalResult: Set r = CheckHandle("node", "document.body")
    AssertTrue Len(r.sharedId) > 0, "sharedIdもあること"
End Sub

Public Sub Test_Ownership_Primitive_NoHandle()
    Dim r As JsEvalResult
    Set r = br.jsEval("'primitive'", Ownership:=True, returnByDetails:=True)
    CheckType r, "string", vbString
    AssertSame "", r.handle, "handle"
    AssertSame "primitive", r.value, "value"
End Sub



'==============================================================================
' ⑦ serializationOptions
'==============================================================================
Public Sub Test_Serialization_MaxDomDepth()
    Dim opt As New Dictionary: opt.Add "maxDomDepth", 1
    Dim r As JsEvalResult
    Set r = br.jsEval("document.body", serializationOptions:=opt, returnByDetails:=True)
    CheckType r, "node", vbObject
    AssertSame "BiDiCDPJson", TypeName(r.rawValue), "TypeName(rawValue)"
    AssertTrue r.rawValue.ExistsKey("children"), "childrenが含まれること"
End Sub

Public Sub Test_Serialization_IncludeShadowTree_Closed()
    Dim opt As New Dictionary: opt.Add "maxDomDepth", 1: opt.Add "includeShadowTree", "all"
    Dim r As JsEvalResult
    Set r = br.jsEval("(() => { const h = document.createElement('div'); h.attachShadow({ mode: 'closed' }).innerHTML = '<b>x</b>'; return h })()", serializationOptions:=opt, returnByDetails:=True)
    CheckType r, "node", vbObject
    AssertTrue Not r.rawValue.NodeKey("shadowRoot") Is Nothing, "closedのshadowRootが含まれること"
End Sub



'==============================================================================
' ⑧ awaitPromise:=True
'==============================================================================
Public Sub Test_AwaitPromise_Number()
    Dim r As JsEvalResult
    Set r = br.jsEval("new Promise(res => setTimeout(() => res(7), 100))", awaitPromise:=True, returnByDetails:=True)
    CheckType r, "number", vbDouble
    AssertSame 7#, r.value, "value"
End Sub



'==============================================================================
' ⑨ script.callFunction 経由
'==============================================================================
Public Sub Test_CallFunction_ScriptHandle()
    Dim bodyHandle As String: bodyHandle = br.jsEval("document.body", Ownership:=True)
    Dim r As JsEvalResult
    Set r = br.jsEval("function(){ return this.tagName }", scriptHandle:=bodyHandle, returnByDetails:=True)
    CheckType r, "string", vbString
    AssertSame "BODY", r.value, "value"
End Sub

Public Sub Test_CallFunction_ObjectArguments()
    Dim args(0) As Dictionary
    Set args(0) = New Dictionary: args(0).Add "type", "number": args(0).Add "value", 20
    Dim r As JsEvalResult
    Set r = br.jsEval("function(n){ return n * 2 }", objectArguments:=args, returnByDetails:=True)
    CheckType r, "number", vbDouble
    AssertSame 40#, r.value, "value"
End Sub

Public Sub Test_CallFunction_SharedIdArgument()
    Dim nodeResult As JsEvalResult: Set nodeResult = Details("document.body")
    Dim args(0) As Dictionary
    Set args(0) = New Dictionary: args(0).Add "sharedId", nodeResult.sharedId
    Dim r As JsEvalResult
    Set r = br.jsEval("function(el){ return el.tagName }", objectArguments:=args, returnByDetails:=True)
    CheckType r, "string", vbString
    AssertSame "BODY", r.value, "value"
End Sub



'==============================================================================
' ⑩ 例外
'==============================================================================
Public Sub Test_Exception_ThrowError()
    Dim r As JsEvalResult: Set r = Details("throw new Error('vba exception')")
    CheckException r
    AssertSame "error", r.ExceptionResultObject.ResultType, "thrown.type"
End Sub

Public Sub Test_Exception_ThrowString()
    Dim r As JsEvalResult: Set r = Details("throw 'just a string'")
    CheckException r
    Dim thrown As JsEvalResult: Set thrown = r.ExceptionResultObject
    AssertSame "string", thrown.ResultType, "thrown.type"
    AssertSame "just a string", thrown.value, "thrown.value"
End Sub

Public Sub Test_Exception_PromiseReject()
    Dim r As JsEvalResult
    Set r = br.jsEval("Promise.reject(new RangeError('rejected'))", awaitPromise:=True, returnByDetails:=True)
    CheckException r
    AssertSame "error", r.ExceptionResultObject.ResultType, "thrown.type"
End Sub

Public Sub Test_Exception_SyntaxError()
    Dim r As JsEvalResult: Set r = Details("function (")
    CheckException r
    AssertSame "error", r.ExceptionResultObject.ResultType, "thrown.type"
End Sub

Public Sub Test_Exception_LetAssign_IsError()
    Dim v As Variant
    v = br.jsEval("throw new Error('let')", returnByDetails:=True)
    AssertTrue IsError(v), "Let代入でCVErrになること"
    AssertSame CVErr(xlErrValue), v, "value"
End Sub

Public Sub Test_Exception_IFEXCEPTION_String()
    Dim r As JsEvalResult
    Set r = br.jsEval("throw new Error('alt')", IFEXCEPTION:="ALT", returnByDetails:=True)
    AssertTrue r.isException, "isException"
    AssertSame "exception", r.ResultType, "type"
    AssertSame CLng(vbString), CLng(r.ResultTypeForVBA), "ResultTypeForVBA"
    AssertSame "ALT", r.value, "value"
End Sub

Public Sub Test_Exception_IFEXCEPTION_Object()
    Dim altObj As New Collection: altObj.Add "alt"
    Dim r As JsEvalResult
    Set r = br.jsEval("throw new Error('alt-obj')", IFEXCEPTION:=altObj, returnByDetails:=True)
    AssertTrue r.isException, "isException"
    AssertSame CLng(vbObject), CLng(r.ResultTypeForVBA), "ResultTypeForVBA"
    AssertTrue r.value Is altObj, "value Is altObj"
End Sub



'==============================================================================
' ⑪ 簡易モード（既定インスタンス使い回し）の回帰確認
'==============================================================================
Public Sub Test_SimpleMode_ValueOnly()
    AssertSame 42.5, br.jsEval("42.5"), "number"
    AssertTrue Len(br.jsEval("document.body", Ownership:=True)) > 0, "handle文字列"
    AssertTrue IsError(br.jsEval("throw 1")), "例外はCVErr"
End Sub

Public Sub Test_SimpleMode_ObjectThenIFEXCEPTION()
    '直前にオブジェクト結果(vbObject)を残した状態で、IFEXCEPTION(文字列)付きの例外を実行してもエラーにならないこと
    Dim o As Object
    Set o = br.jsEval("({ a: 1 })")
    AssertSame "ALT", br.jsEval("throw new Error('x')", IFEXCEPTION:="ALT"), "IFEXCEPTION"
End Sub



'==============================================================================
' ヘルパー
'==============================================================================
Private Function Details(js As String) As JsEvalResult
    Set Details = br.jsEval(js, returnByDetails:=True)
End Function

Private Sub CheckType(r As JsEvalResult, expType As String, expVarType As VbVarType)
    AssertTrue Not r Is Nothing, "JsEvalResultが返ること"
    AssertSame expType, r.ResultType, "ResultType"
    AssertSame "", r.ResultSubtype, "ResultSubtype(BiDiは常に空)"
    AssertSame CLng(expVarType), CLng(r.ResultTypeForVBA), "ResultTypeForVBA"
    AssertFalse r.isException, "isException"
End Sub

Private Sub CheckSpecialNumber(js As String)
    Dim r As JsEvalResult: Set r = Details(js)
    CheckType r, "number", vbError
    AssertSame CVErr(xlErrNum), r.value, "value"
End Sub

Private Sub CheckNoValue(expType As String, js As String)
    Dim r As JsEvalResult: Set r = Details(js)
    CheckType r, expType, vbObject
    AssertTrue r.value Is Nothing, "value Is Nothing"
    AssertSame "", r.handle, "handle"
End Sub

Private Sub CheckJsonValue(expType As String, js As String)
    Dim r As JsEvalResult: Set r = Details(js)
    CheckType r, expType, vbObject
    AssertSame "BiDiCDPJson", TypeName(r.value), "TypeName(value)"
End Sub

Private Function CheckHandle(expType As String, js As String) As JsEvalResult
    Dim r As JsEvalResult
    Set r = br.jsEval(js, Ownership:=True, returnByDetails:=True)
    CheckType r, expType, vbDataObject
    AssertTrue Len(r.handle) > 0, "handleがあること"
    AssertSame r.handle, r.value, "value = handle"
    Set CheckHandle = r
End Function

Private Sub CheckNonSpec(js As String)
    Dim r As JsEvalResult: Set r = Details(js)
    AssertTrue Not r Is Nothing, "JsEvalResultが返ること"
    AssertFalse r.isException, "isException"
    AssertTrue Len(r.ResultType) > 0, "ResultTypeがあること"
End Sub

Private Sub CheckException(r As JsEvalResult)
    AssertTrue Not r Is Nothing, "JsEvalResultが返ること"
    AssertTrue r.isException, "isException"
    AssertSame "exception", r.ResultType, "ResultType"
    AssertSame CLng(vbError), CLng(r.ResultTypeForVBA), "ResultTypeForVBA"
    AssertSame CVErr(xlErrValue), r.value, "value"
    AssertTrue Len(r.ExceptionText) > 0, "ExceptionText"
    AssertTrue Not r.rawException Is Nothing, "rawException"
    AssertTrue Not r.ExceptionResultObject Is Nothing, "ExceptionResultObject"
End Sub

'VarTypeまで含めて一致を確認します（Null/Empty/エラー値も比較可能）
Private Sub AssertSame(expected As Variant, actual As Variant, label As String)
    If VarType(expected) <> VarType(actual) Then Fail label & ": VarType expected " & VarType(expected) & " but got " & VarType(actual) & " (" & Describe(actual) & ")"
    Select Case VarType(expected)
        Case vbNull, vbEmpty
            Exit Sub
        Case vbError
            If CStr(expected) <> CStr(actual) Then Fail label & ": expected " & CStr(expected) & " but got " & CStr(actual)
        Case Else
            If expected <> actual Then Fail label & ": expected <" & CStr(expected) & "> but got <" & CStr(actual) & ">"
    End Select
End Sub

Private Sub AssertTrue(condition As Boolean, label As String)
    If Not condition Then Fail label & ": expected True"
End Sub

Private Sub AssertFalse(condition As Boolean, label As String)
    If condition Then Fail label & ": expected False"
End Sub

Private Function Describe(v As Variant) As String
    If IsObject(v) Then Describe = TypeName(v): Exit Function
    Select Case VarType(v)
        Case vbNull: Describe = "Null"
        Case vbEmpty: Describe = "Empty"
        Case Else: Describe = CStr(v)
    End Select
End Function

Private Sub Fail(msg As String)
    Err.Raise vbObjectError + 1000, "Test_JsEvalResult_BiDi_xlflow", msg
End Sub
