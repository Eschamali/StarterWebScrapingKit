Attribute VB_Name = "Test_JsEvalResult_CDP_Forxlflow"
'==============================================================================
' JsEvalResult.cls 自動テスト（CDP版 / xlflow test 用）
' ・Test_JsEvalResult_CDP（Stopで目視確認する版）と同じケースを、期待値で自動判定します
' ・BeforeAll でCDPブラウザを起動し、AfterAll で閉じます
' ・実行例：xlflow test --module Test_JsEvalResult_CDP_Forxlflow --session --json
'==============================================================================
Option Explicit

Private br As CDPContext



'==============================================================================
' ライフサイクル
'==============================================================================
Public Sub BeforeAll()
    Set br = ShSetting01_StartBrowser.StartCDPModeContext
End Sub

Public Sub AfterAll()
    If Not br Is Nothing Then br.ThisCDPBrowser.CloseBrowser
    Set br = Nothing
End Sub



'==============================================================================
' ① プリミティブ型
'==============================================================================
Public Sub Test_Primitive_String()
    Dim r As JsEvalResult: Set r = Details("'hello VBA'")
    CheckType r, "string", "", vbString
    AssertSame "hello VBA", r.value, "value"
    AssertSame "", r.objectId, "objectId"
End Sub

Public Sub Test_Primitive_Number()
    Dim r As JsEvalResult: Set r = Details("42.5")
    CheckType r, "number", "", vbDouble
    AssertSame 42.5, r.value, "value"
    AssertSame "42.5", r.Description, "description"
End Sub

Public Sub Test_Primitive_Boolean()
    Dim r As JsEvalResult: Set r = Details("true")
    CheckType r, "boolean", "", vbBoolean
    AssertSame True, r.value, "value"
End Sub

Public Sub Test_Primitive_Undefined()
    Dim r As JsEvalResult: Set r = Details("undefined")
    CheckType r, "undefined", "", vbEmpty
    AssertSame Empty, r.value, "value"
End Sub

Public Sub Test_Primitive_BigInt()
    Dim r As JsEvalResult: Set r = Details("12345678901234567890n")
    CheckType r, "bigint", "", vbString
    AssertSame "12345678901234567890n", r.value, "value"
    AssertSame "12345678901234567890n", r.UnserializableValue, "unserializableValue"
End Sub



'==============================================================================
' ② number の unserializableValue → #NUM!
'==============================================================================
Public Sub Test_Number_NaN()
    CheckUnserializableNumber "NaN"
End Sub

Public Sub Test_Number_Infinity()
    CheckUnserializableNumber "Infinity"
End Sub

Public Sub Test_Number_MinusInfinity()
    CheckUnserializableNumber "-Infinity"
End Sub

Public Sub Test_Number_MinusZero()
    CheckUnserializableNumber "-0"
End Sub



'==============================================================================
' ③ symbol / function
'==============================================================================
Public Sub Test_Symbol()
    Dim r As JsEvalResult: Set r = Details("Symbol('vba')")
    CheckObjectId r, "symbol", "", "", "Symbol(vba)"
End Sub

Public Sub Test_Function()
    Dim r As JsEvalResult: Set r = Details("(function add(a, b){ return a + b })")
    CheckObjectId r, "function", "", "Function", "function add(a, b){ return a + b }"
End Sub



'==============================================================================
' ④ type: object の subtype（objectId で返る）
'==============================================================================
Public Sub Test_Object_Plain()
    CheckObjectId Details("({ a: 1, b: 'x' })"), "object", "", "Object", "Object"
End Sub

Public Sub Test_Object_Array()
    CheckObjectId Details("[1, 'two', true]"), "object", "array", "Array", "Array(3)"
End Sub

Public Sub Test_Object_Null()
    Dim r As JsEvalResult: Set r = Details("null")
    CheckType r, "object", "null", vbNull
    AssertSame Null, r.value, "value"
    AssertSame "", r.objectId, "objectId"
End Sub

Public Sub Test_Object_Node()
    CheckObjectId Details("document.body"), "object", "node", "HTMLBodyElement", "body"
End Sub

Public Sub Test_Object_RegExp()
    CheckObjectId Details("/ab+c/gi"), "object", "regexp", "RegExp", "/ab+c/gi"
End Sub

Public Sub Test_Object_Date()
    Dim r As JsEvalResult: Set r = Details("new Date(2026, 9, 15, 12, 34, 56)")
    CheckObjectId r, "object", "date", "Date"
    AssertTrue Left$(r.Description, 24) = "Thu Oct 15 2026 12:34:56", "description: " & r.Description
End Sub

Public Sub Test_Object_Map()
    CheckObjectId Details("new Map([['k', 1]])"), "object", "map", "Map", "Map(1)"
End Sub

Public Sub Test_Object_Set()
    CheckObjectId Details("new Set([1, 2])"), "object", "set", "Set", "Set(2)"
End Sub

Public Sub Test_Object_WeakMap()
    CheckObjectId Details("new WeakMap()"), "object", "weakmap", "WeakMap", "WeakMap"
End Sub

Public Sub Test_Object_WeakSet()
    CheckObjectId Details("new WeakSet()"), "object", "weakset", "WeakSet", "WeakSet"
End Sub

Public Sub Test_Object_Iterator()
    CheckObjectId Details("new Map([['k', 1]]).entries()"), "object", "iterator"
End Sub

Public Sub Test_Object_ArrayIterator_NoSubtype()
    CheckObjectId Details("[1, 2][Symbol.iterator]()"), "object", "", "Array Iterator", "Array Iterator"
End Sub

Public Sub Test_Object_Generator()
    CheckObjectId Details("(function* (){ yield 1 })()"), "object", "generator", "Generator", "Generator"
End Sub

Public Sub Test_Object_Error()
    Dim r As JsEvalResult: Set r = Details("new TypeError('vba error')")
    CheckObjectId r, "object", "error", "TypeError"
    AssertTrue Left$(r.Description, 20) = "TypeError: vba error", "description: " & r.Description
    AssertFalse r.isException, "値として返したErrorは例外ではない"
End Sub

Public Sub Test_Object_Proxy()
    CheckObjectId Details("new Proxy({}, {})"), "object", "proxy", "Object", "Proxy(Object)"
End Sub

Public Sub Test_Object_Promise()
    CheckObjectId Details("Promise.resolve(1)"), "object", "promise", "Promise", "Promise"
End Sub

Public Sub Test_Object_TypedArray()
    CheckObjectId Details("new Uint8Array([1, 2, 3])"), "object", "typedarray", "Uint8Array", "Uint8Array(3)"
End Sub

Public Sub Test_Object_ArrayBuffer()
    CheckObjectId Details("new ArrayBuffer(8)"), "object", "arraybuffer", "ArrayBuffer", "ArrayBuffer(8)"
End Sub

Public Sub Test_Object_DataView()
    CheckObjectId Details("new DataView(new ArrayBuffer(8))"), "object", "dataview", "DataView", "DataView(8)"
End Sub

Public Sub Test_Object_WebAssemblyMemory()
    CheckObjectId Details("new WebAssembly.Memory({ initial: 1 })"), "object", "webassemblymemory", "Memory", "Memory(1)"
End Sub

Public Sub Test_Object_TrustedType()
    CheckObjectId Details("trustedTypes.createPolicy('vba' + Date.now(), { createHTML: s => s }).createHTML('<b>x</b>')"), "object", "trustedtype", "TrustedHTML", "<b>x</b>"
End Sub



'==============================================================================
' ⑤ returnByValue / deep / awaitPromise
'==============================================================================
Public Sub Test_ReturnByValue_Object()
    Dim r As JsEvalResult: Set r = br.jsEval("({ a: 1, b: 'x' })", returnByValue:=True, returnByDetails:=True)
    CheckType r, "object", "", vbObject
    AssertSame "BiDiCDPJConv", TypeName(r.value), "TypeName(value)"
    AssertSame "", r.objectId, "objectId"
End Sub

Public Sub Test_ReturnByValue_Array()
    Dim r As JsEvalResult: Set r = br.jsEval("[1, 'two', true]", returnByValue:=True, returnByDetails:=True)
    CheckType r, "object", "", vbObject
    AssertSame "BiDiCDPJConv", TypeName(r.value), "TypeName(value)"
End Sub

Public Sub Test_ReturnByValue_Function()
    Dim r As JsEvalResult: Set r = br.jsEval("(function add(a, b){ return a + b })", returnByValue:=True, returnByDetails:=True)
    CheckType r, "function", "", vbObject
End Sub

Public Sub Test_Deep_Object()
    Dim r As JsEvalResult: Set r = br.jsEval("({ a: 1, list: [1, 2] })", serializationOptions:=DeepOption(), returnByDetails:=True)
    CheckType r, "object", "", vbObject
    AssertSame "BiDiCDPJConv", TypeName(r.value), "TypeName(value)"
    AssertTrue Len(r.objectId) > 0, "deepでもobjectIdは付く"
    AssertTrue Not r.DeepSerializedValue Is Nothing, "DeepSerializedValue"
End Sub

Public Sub Test_Deep_Node()
    Dim r As JsEvalResult: Set r = br.jsEval("document.body", serializationOptions:=DeepOption(), returnByDetails:=True)
    CheckType r, "object", "node", vbObject
    AssertSame "HTMLBodyElement", r.ClassName, "className"
End Sub

Public Sub Test_AwaitPromise_Number()
    Dim r As JsEvalResult: Set r = br.jsEval("new Promise(res => setTimeout(() => res(7), 100))", awaitPromise:=True, returnByDetails:=True)
    CheckType r, "number", "", vbDouble
    AssertSame 7#, r.value, "value"
End Sub



'==============================================================================
' ⑥ Runtime.callFunctionOn 経由
'==============================================================================
Public Sub Test_CallFunctionOn_ObjectId()
    Dim bodyId As String: bodyId = br.jsEval("document.body")
    Dim r As JsEvalResult: Set r = br.jsEval("function(){ return this.tagName }", objectId:=bodyId, returnByDetails:=True)
    CheckType r, "string", "", vbString
    AssertSame "BODY", r.value, "value"
End Sub

Public Sub Test_CallFunctionOn_ObjectArguments()
    Dim args(0) As Dictionary
    Set args(0) = New Dictionary: args(0).Add "value", 20
    Dim r As JsEvalResult: Set r = br.jsEval("function(n){ return { doubled: n * 2 } }", objectArguments:=args, returnByDetails:=True)
    CheckObjectId r, "object", "", "Object", "Object"
End Sub



'==============================================================================
' ⑦ 例外
'==============================================================================
Public Sub Test_Exception_ThrowError()
    Dim r As JsEvalResult: Set r = Details("throw new Error('vba exception')")
    CheckException r
    Dim thrown As JsEvalResult: Set thrown = r.ExceptionResultObject
    AssertSame "object", thrown.ResultType, "thrown.type"
    AssertSame "error", thrown.ResultSubtype, "thrown.subtype"
    AssertSame "Error", thrown.ClassName, "thrown.className"
    AssertTrue Left$(thrown.Description, 20) = "Error: vba exception", "thrown.description: " & thrown.Description
End Sub

Public Sub Test_Exception_ThrowString()
    Dim r As JsEvalResult: Set r = Details("throw 'just a string'")
    CheckException r
    Dim thrown As JsEvalResult: Set thrown = r.ExceptionResultObject
    AssertSame "string", thrown.ResultType, "thrown.type"
    AssertSame "just a string", thrown.value, "thrown.value"
End Sub

Public Sub Test_Exception_PromiseReject()
    Dim r As JsEvalResult: Set r = br.jsEval("Promise.reject(new RangeError('rejected'))", awaitPromise:=True, returnByDetails:=True)
    CheckException r
    AssertSame "RangeError", r.ExceptionResultObject.ClassName, "thrown.className"
End Sub

Public Sub Test_Exception_SyntaxError()
    Dim r As JsEvalResult: Set r = Details("function (")
    CheckException r
    AssertSame "SyntaxError", r.ExceptionResultObject.ClassName, "thrown.className"
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
    AssertSame vbString, r.ResultTypeForVBA, "ResultTypeForVBA"
    AssertSame "ALT", r.value, "value"
End Sub

Public Sub Test_Exception_IFEXCEPTION_Object()
    Dim altObj As New Collection: altObj.Add "alt"
    Dim r As JsEvalResult
    Set r = br.jsEval("throw new Error('alt-obj')", IFEXCEPTION:=altObj, returnByDetails:=True)
    AssertTrue r.isException, "isException"
    AssertSame vbObject, r.ResultTypeForVBA, "ResultTypeForVBA"
    AssertTrue r.value Is altObj, "value Is altObj"
End Sub



'==============================================================================
' ⑧ 簡易モード（既定インスタンス使い回し）の回帰確認
'==============================================================================
Public Sub Test_SimpleMode_ValueOnly()
    AssertSame 42.5, br.jsEval("42.5"), "number"
    AssertTrue Len(br.jsEval("document.body")) > 0, "objectId文字列"
    AssertTrue IsError(br.jsEval("throw 1")), "例外はCVErr"
End Sub

Public Sub Test_SimpleMode_ObjectThenIFEXCEPTION()
    '直前にオブジェクト結果(vbObject)を残した状態で、IFEXCEPTION(文字列)付きの例外を実行してもエラーにならないこと
    Dim o As Object
    Set o = br.jsEval("({ a: 1 })", returnByValue:=True)
    AssertSame "ALT", br.jsEval("throw new Error('x')", IFEXCEPTION:="ALT"), "IFEXCEPTION"
End Sub



'==============================================================================
' ヘルパー
'==============================================================================
Private Function Details(js As String) As JsEvalResult
    Set Details = br.jsEval(js, returnByDetails:=True)
End Function

Private Function DeepOption() As Dictionary
    Set DeepOption = New Dictionary
    DeepOption.Add "serialization", "deep"
End Function

Private Sub CheckType(r As JsEvalResult, expType As String, expSubtype As String, expVarType As VbVarType)
    AssertTrue Not r Is Nothing, "JsEvalResultが返ること"
    AssertSame expType, r.ResultType, "ResultType"
    AssertSame expSubtype, r.ResultSubtype, "ResultSubtype"
    AssertSame CLng(expVarType), CLng(r.ResultTypeForVBA), "ResultTypeForVBA"
    AssertFalse r.isException, "isException"
End Sub

Private Sub CheckObjectId(r As JsEvalResult, expType As String, expSubtype As String, Optional expClassName As Variant, Optional expDescription As Variant)
    CheckType r, expType, expSubtype, vbDataObject
    AssertTrue Len(r.objectId) > 0, "objectIdがあること"
    AssertSame r.objectId, r.value, "value = objectId"
    If Not IsMissing(expClassName) Then AssertSame expClassName, r.ClassName, "className"
    If Not IsMissing(expDescription) Then AssertSame expDescription, r.Description, "description"
End Sub

Private Sub CheckUnserializableNumber(js As String)
    Dim r As JsEvalResult: Set r = Details(js)
    CheckType r, "number", "", vbError
    AssertSame CVErr(xlErrNum), r.value, "value"
    AssertSame js, r.UnserializableValue, "unserializableValue"
End Sub

Private Sub CheckException(r As JsEvalResult)
    AssertTrue Not r Is Nothing, "JsEvalResultが返ること"
    AssertTrue r.isException, "isException"
    AssertSame "exception", r.ResultType, "ResultType"
    AssertSame "", r.ResultSubtype, "ResultSubtype"
    AssertSame CLng(vbError), CLng(r.ResultTypeForVBA), "ResultTypeForVBA"
    AssertSame CVErr(xlErrValue), r.value, "value"
    AssertTrue Len(r.ExceptionText) > 0, "ExceptionText"
    AssertTrue Not r.rawException Is Nothing, "rawException"
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
    Err.Raise vbObjectError + 1000, "Test_JsEvalResult_CDP_Auto", msg
End Sub
