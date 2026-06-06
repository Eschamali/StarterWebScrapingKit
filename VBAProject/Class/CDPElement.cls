VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
' CDP HTML Element Class
'------------------------------------------------------------------------------------------------
' Author(s)   :
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       24/10/25 Daniel Polak: Dim of variables and added Option Explicit, removed extraneous blank lines
'                              getPrevSibling: fixed bug where wrong assignment led to no return result
'                              setAttribute: changed Function to Sub
'       02/03/23 Long Vh: - Added the element class for easy object reference.
'       13/11/23 Long Vh: - Added getChildren & property let innerHTML. Updated .value to work
'                           with TextArea node type.
'       23/11/23 Long Vh: - Added raiseTimeoutError option for .onExist and .onExistNot
'       12/12/23 Long Vh: - Added checkExist to only execute things if True as well as ifExist
'       12/07/24 Long Vh: - Added the sendKey method.
' Description :
'       This class helps to simplify HTML element object referencing task using CDP. This shall
'       enables the development of 'getElementBy' methods that are commonly used for automating
'       browser sessions. Without this class, all interaction with the HTML elements will have
'       to be parsed via the jsEval method which can be counter-intuitive.
'================================================================================================

Option Explicit

Private Const ThisClassName As String = "CDPElement"

Private main As CDPContext
Private varJS As String                  ' DOM検索式（ポーリング用 & objectId 未取得時の lazy 取得用）
Private elemVarType As Long              ' VBA VarType定数: vbNull=1: 未発見 / vbString=8: objectId取得済み / vbEmpty=0: 未評価
Private elementObjectId As String        ' CDPのobjectId（StrPtr=0:未取得 / StrPtr≠0:取得済み）
Private checkExist As Boolean            ' Needed for the isExist check
Private Const printDbgMsg = True

Public Enum keyboardCode    'Used for .sendKey operation in the browser class and element class

    keyEnter = 13
    keyDelete = 46
    keyTab = 9
    keyEsc = 27
    keyBackspace = 8

End Enum



'***************************************************************************************************
'                                   ■■■ 初期化処理 ■■■
'***************************************************************************************************
'* 機能　　：アクティブな`CDPBrowser`を継承させ、`jsEval`で実行した結果の一部を設定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：targetObjMain      継承したい`CDPBrowser`そのもの
'            varJS_             継承直前に実行したJavaScriptコードそのもの。ポーリング用に使用します。
'            varType_           JavaScript実行結果の型をVBA型に変換した際の状況(https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/vartype-function)
'            objectId_          `Runtime.callFunctionOn`で使用する`objectId`
'---------------------------------------------------------------------------------------------------
'* 詳細説明：要素がなく、`objectId`が返せない場合は、`objectId_`は、`StrPtr`チェックで0になります
'* 注意事項：・このヘルパークラスを使用する際は、`objectId`を返すようなJavaScriptコードとして実行すること。要素操作準備なら`objectId`でほぼほぼ返ると思いますが念押しです。
'            　`objectId_`は、「vbNull(1)=未発見 / vbString(8)=objectId取得済み / vbEmpty(0)=未評価」の3パターンを想定してます。
'            ・現状複数取得系(getElementsByXPath等)の場合、`varJS_`は機能しません。
'***************************************************************************************************
Public Sub Init(targetObjMain As CDPContext, Optional varJS_ As String, Optional varType_ As Long, Optional objectId_ As String)
    Set main = targetObjMain
    varJS = varJS_
    elemVarType = varType_
    If StrPtr(objectId_) <> 0 Then elementObjectId = objectId_  '空文字も含め「代入された」ことを記録
End Sub



'***************************************************************************************************
'                                ■■■ パラメーター確認 ■■■
'***************************************************************************************************
Property Get varPath() As String
'----------------------------------------------------------------------------------
' Retrieve the path of this variable initialized. The same name can also be used
' directly in the browser console window if the element is assigned successfully.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

    varPath = varJS

End Property

Property Get varResult() As Long
'----------------------------------------------------------------------------------
' Retrieve the VarType code of the result returned by getElement method.
' vbNull (1): element was not found (JS null)
' vbString (8): element was found (objectId obtained)
' vbEmpty (0): not yet evaluated (e.g. initialized via varJS only)
' Updated: 27/04/2023
'----------------------------------------------------------------------------------

    varResult = elemVarType

End Property

'***************************************************************************************************
'* 機能　　：このClassで保持中の`ObjectId`を、引数名として、DevToolsコンソールでも扱えるようにします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：なんかうまくいかない時、これを利用することで、DevToolsコンソール上でのデバッグが可能になります。
'* 注意事項：サイトによってはこの不自然なグローバル変数を検知するスクレイピング対策があります。ご利用時は注意してください。
'***************************************************************************************************
Public Sub ExposeDevTools(ByVal varName As String)
    If elementObjectId = "" Then Exit Sub

    Dim params As New Dictionary
    Dim arg As New Dictionary
    Dim args As New Collection

    ' 1. 関数に渡す引数の設定 (ユーザーが指定した変数名)
    arg.Add "value", varName
    args.Add arg

    ' 2. 実行
    main.jsEval "function(name) { window[name] = this; console.log('VBA: Exposed as ' + name); }", getObjectId, objectArguments:=args
End Sub

'***************************************************************************************************
'* 機能　　：このClassで保持中の`ObjectId`をそのまま、取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：このClassで保持中の`ObjectId`
'* 引数　　：refresh      継承直前に実行したJavaScriptコードそのものを再度実行し、ObjectIdを再取得します。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・StrPtr(elementObjectId) ≠ 0 → 取得済みのobjectIdを即返す
'            ・StrPtr(elementObjectId) = 0  → varJSをruntime.evaluateで評価してobjectIdを遅延取得(Lazy Init)
'
'* 注意事項：返り値が空文字の場合は、要素がnull(未発見)を意味します。
'***************************************************************************************************
Property Get getObjectId(Optional refresh As Boolean) As String

   '取得済みであれば即返す
    If StrPtr(elementObjectId) <> 0 And Not refresh Then
        getObjectId = elementObjectId
        Exit Property
    End If

   'varJSから遅延取得
    Dim res As Variant
    res = main.jsEval(varJS, returnByValue:=False, dbgMsg:=False)
    If VarType(res) = vbString Then
        elementObjectId = CStr(res)
    Else
        elementObjectId = ""    '空文字を代入することで「取得試行済み」扱い（StrPtr≠0になる）
    End If
    getObjectId = elementObjectId

End Property

'========================================================================================================
' ELEMENT PROPERTY INTERACTIONS
'========================================================================================================

Property Get value() As String
'----------------------------------------------------------
' Get the value property of the element in the HTML.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    value = main.jsEval("function(){ return this.value }", objectId:=getObjectId(), dbgMsg:=printDbgMsg)

End Property

Property Let value(strTextVal As String)
'-----------------------------------------------------------------------------------------------
' Change the value property of the element in the HTML. Enhanced to work even with React fields.
' For React fields, CDP need to input value to the element prototype in order to trigger the
' React event handler. For further references refer below:
' https://stackoverflow.com/a/46012210
' https://hustle.bizongo.in/simulate-react-on-change-on-controlled-components-baa336920e04
' Updated: 28/04/23 - React fields value inputing enhancement
'          12/12/23 - added checkExist
'-----------------------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".value"


    Dim e1 As String
    Dim isReact As Boolean

   'Gracefully exit the value input if the element is found not exist under isExist
    If checkExist Then If Not isExist Then Exit Property

   'Reactフィールド検出（thisコンテキスト内でObject.entriesを実行）
    e1 = main.jsEval("function(){ try { return Object.entries(this)[0][0] } catch(e) { return '' } }", objectId:=getObjectId(), dbgMsg:=False)
    If InStr(e1, "react") Or InStr(e1, "Tracker") Or InStr(e1, "State") Then isReact = True
    If isReact Then printMsg WARN_, "FYI: this is a REACT field as the first method is " & e1, FromProcedureName

   'objectArgumentsで安全に値を渡す（エスケープ不要）
    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", strTextVal
    args.Add arg

   'Perform inputing using all the methods
    If Not isReact Then
        main.jsEval "function(v){ this.value = v }", objectId:=getObjectId(), objectArguments:=args, dbgMsg:=printDbgMsg
    Else
        main.jsEval "function(v){ Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(this, v) }", objectId:=getObjectId(), objectArguments:=args
        main.jsEval "function(v){ Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set.call(this, v) }", objectId:=getObjectId(), objectArguments:=args
    End If

   'Most input elements need this
    fireEvent "input", printDbg:=False

End Property

Property Get innerText() As String
'----------------------------------------------------------
' Get the innerText property of the element in the HTML.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    innerText = main.jsEval("function(){ return this.innerText }", objectId:=getObjectId(), dbgMsg:=printDbgMsg)

End Property

Property Let innerText(strTextVal As String)
'----------------------------------------------------------
' Change the innerText property of the element in the HTML.
' Note: objectArgumentsで安全に値を渡すのでエスケープ不要
' Updated: 02/03/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    main.jsEval "function(){ this.scrollIntoView(false) }", objectId:=getObjectId(), dbgMsg:=False

    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", strTextVal
    args.Add arg
    main.jsEval "function(v){ this.innerText = v }", objectId:=getObjectId(), objectArguments:=args, dbgMsg:=printDbgMsg

End Property

Property Get innerHTML() As String
'----------------------------------------------------------
' Get the innerHTML property of the element in the HTML.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    innerHTML = main.jsEval("function(){ return this.innerHTML }", objectId:=getObjectId(), dbgMsg:=printDbgMsg)

End Property

Property Let innerHTML(strTextVal As String)
'----------------------------------------------------------
' Change the innerHTML property of the element.
' Note: objectArgumentsで安全に値を渡すのでエスケープ不要
' Updated: 13/11/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    main.jsEval "function(){ this.scrollIntoView(false) }", objectId:=getObjectId(), dbgMsg:=False

    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", strTextVal
    args.Add arg
    main.jsEval "function(v){ this.innerHTML = v }", objectId:=getObjectId(), objectArguments:=args, dbgMsg:=printDbgMsg

End Property

Property Let checked(boolChecked As Boolean)
'----------------------------------------------------------
' Change the checked value of a check box element
' Updated: 04/12/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", CBool(boolChecked)
    args.Add arg
    main.jsEval "function(v){ this.checked = v }", objectId:=getObjectId(), objectArguments:=args, dbgMsg:=False

End Property

Property Get checked() As Boolean
'----------------------------------------------------------
' Change the checked value of a check box element
' Updated: 04/12/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    checked = CBool(main.jsEval("function(){ return this.checked }", objectId:=getObjectId(), dbgMsg:=printDbgMsg))

End Property

Property Let selected(selectedOption As String)
'----------------------------------------------------------
' Change the selected value of an html select element type.
' eg. br.getElementbyId("list").selected = 2
' Updated: 06/03/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", selectedOption
    args.Add arg
    main.jsEval "function(v){ this.selectedIndex = v }", objectId:=getObjectId(), objectArguments:=args, dbgMsg:=False
    fireEvent "change"

End Property

Property Get selected() As String
'----------------------------------------------------------
' Get the selected option. Only returns the first option.
' Updated: 06/03/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    selected = main.jsEval("function(){ return this.selectedOptions[0] }", objectId:=getObjectId(), dbgMsg:=printDbgMsg)

End Property

'========================================================================================================
' ELEMENT OBJECT INTERACTIONS
'========================================================================================================

Public Function click(Optional till As ReadyState = isComplete) As Boolean
'----------------------------------------------------------
' Perform a standard click on the element. Remember to
' perform wait if necessary for the page to be refreshed.
' Updated: 12/12/2023
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

   'Scroll into the view for better visibility of the automation
    main.jsEval "function(){ this.scrollIntoView(false) }", objectId:=getObjectId(), dbgMsg:=False

   'Some fields require the browser to manually focus before allowing
   'the click action, such as the input type 'file' field.
    Dim params As New Scripting.Dictionary
    params("text") = ""
    ExecuteCDP "Input.insertText", params

   'Perform click
    main.jsEval "function(){ this.dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true})) }", objectId:=getObjectId(), dbgMsg:=printDbgMsg
    main.wait till:=till
    click = True

End Function

Public Function submit(Optional till As ReadyState = isComplete) As Boolean
'----------------------------------------------------------
' Submit a form element.
' Updated: 12/12/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function
    main.jsEval "function(){ this.form.submit() }", objectId:=getObjectId(), dbgMsg:=printDbgMsg
    main.wait till:=till
    submit = True

End Function

Public Function getAttribute(strAttributeName As String) As String
'----------------------------------------------------------
' Get the value of an element attribute based on its name.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function
    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", strAttributeName
    args.Add arg
    getAttribute = main.jsEval("function(name){ return this.getAttribute(name) || '' }", objectId:=getObjectId(), objectArguments:=args, dbgMsg:=printDbgMsg)

End Function

Public Function setAttribute(strAttributeName As String, strValue As String)
'----------------------------------------------------------
' Set value to an element attribute based on its name.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function
    Dim args2 As New Collection
    Dim a1 As New Scripting.Dictionary: a1.Add "value", strAttributeName: args2.Add a1
    Dim a2 As New Scripting.Dictionary: a2.Add "value", strValue: args2.Add a2
    main.jsEval "function(name, val){ this.setAttribute(name, val) }", objectId:=getObjectId(), objectArguments:=args2, dbgMsg:=printDbgMsg

End Function

Public Function setSelection(strOptionName As String)
'----------------------------------------------------------
' Set option value to a select element.
' Updated: 06/03/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function
    main.jsEval "function(){ this.scrollIntoView(false) }", objectId:=getObjectId(), dbgMsg:=False
    Me.value = strOptionName
    fireEvent "change"

End Function

Public Function selectText()
'----------------------------------------------------------
' Select all the text held by the element.
' Updated: 18/09/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function
    main.jsEval "function(){ let r = document.createRange(); let s = window.getSelection(); r.selectNodeContents(this); s.removeAllRanges(); s.addRange(r) }", objectId:=getObjectId()

End Function

Public Function focus()
'----------------------------------------------------------
' Set focus to an element.
' Updated: 18/09/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function
    main.jsEval "function(){ this.focus() }", objectId:=getObjectId(), dbgMsg:=False

End Function

Public Function fireEvent(strEventName As String, Optional till As ReadyState = isComplete, Optional printDbg As Boolean = printDbgMsg) As Boolean
'------------------------------------------------------------
' Dispatch the custom event on the element. Unlike the IE
' DOM event name, the argument strEventName for JS should
' not start with "on" such as "onblur" but only "blur".
' Updated: 12/12/2023
'------------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function

   'Auto format some common event name mistakes
    strEventName = Replace(strEventName, "on", "")  'eg. "onchange" -> "change" - correct JS event name

   'Fire the event; simulated=true is in case form is REACT 15.5 or lower
    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", strEventName
    args.Add arg
    main.jsEval "function(ev){ this.dispatchEvent(new Event(ev, { bubbles: true, simulated: true })) }", objectId:=getObjectId(), objectArguments:=args, dbgMsg:=printDbg

   'Wait till applied
    main.wait till:=till, dbgState:=printDbg
    fireEvent = True

End Function

Public Function sendString(textToSend As String) As Boolean
'----------------------------------------------------------------
' Attempt to directly send the text string to the element via
' backdoor interaction with CDP instead of manipulating the DOM.
' Updated: 28/04/2023
'----------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".sendString"


    If checkExist Then If Not isExist Then Exit Function

   'First clear the current text of the field
    main.jsEval "function(){ this.value = '' }", objectId:=getObjectId(), dbgMsg:=False

   'getObjectId()で取得済みのobjectIdをそのまま使用（Runtime.evaluateの二度取り不要）
    Dim oid As String: oid = getObjectId()
    If Left(oid, 1) <> "-" Then oid = "+" & oid   'DOM.focusハング回避用符号補完

   'Set focus on the node
    Dim params As New Scripting.Dictionary
    params("objectId") = oid
    ExecuteCDP "DOM.focus", params

   'Sendkeys to the node
    params.RemoveAll
    params("text") = textToSend
    ExecuteCDP "Input.insertText", params

    sendString = True
    printMsg info_, "Executed sendString(""" & textToSend & """) successfully on objectId: " & getObjectId(), FromProcedureName

End Function

Public Function sendClick() As Boolean
'----------------------------------------------------------------
' Attempt to simulate physical left mouse click input instead of
' DOM manipulation with JavaScript dispatchEvent method.
'----------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".sendClick"

    Dim x As Double, y As Double

    If checkExist Then If Not isExist Then Exit Function

   'Viewport 内に収めてから座標取得（CDP の x/y は CSS ピクセル・数値必須）
    main.jsEval "function(){ this.scrollIntoView(false) }", objectId:=getObjectId(), dbgMsg:=False

    x = CDbl(main.jsEval("function(){ return this.getBoundingClientRect().x }", objectId:=getObjectId()))
    y = CDbl(main.jsEval("function(){ return this.getBoundingClientRect().y }", objectId:=getObjectId()))

    Dim params As New Scripting.Dictionary
    params("type") = "mousePressed"
    params("x") = x
    params("y") = y
    params("button") = "left"
    params("clickCount") = 1
    ExecuteCDP "Input.dispatchMouseEvent", params

    params.RemoveAll
    params("type") = "mouseReleased"
    params("x") = x
    params("y") = y
    params("button") = "left"
    params("clickCount") = 1
    ExecuteCDP "Input.dispatchMouseEvent", params

    sendClick = True
    printMsg info_, "Executed sendClick successfully on objectId: " & getObjectId(), FromProcedureName

End Function

Public Function sendKey(Key As keyboardCode, Optional altKey As Boolean = False) As Boolean
'----------------------------------------------------------------
' Attempt to directly send a unique keyboard input to the element.
' Updated: 12/07/2024
'----------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".sendKey"


    If checkExist Then If Not isExist Then Exit Function

   'getObjectId()で取得済みのobjectIdをそのまま使用
    Dim oid As String: oid = getObjectId()
    If Left(oid, 1) <> "-" Then oid = "+" & oid   'DOM.focusハング回避用符号補完

   'Set focus on the node
    Dim params As New Scripting.Dictionary
    params("objectId") = oid
    ExecuteCDP "DOM.focus", params

   'Send key down event to the browser
    params.RemoveAll
    params("type") = "keyDown"
    params("windowsVirtualKeyCode") = Key
    params("altKey") = altKey
    ExecuteCDP "Input.dispatchKeyEvent", params

   'Send key up event to the browser to release the key
    params.RemoveAll
    params("type") = "keyUp"
    params("windowsVirtualKeyCode") = Key
    params("altKey") = False
    ExecuteCDP "Input.dispatchKeyEvent", params

    sendKey = True
    printMsg info_, "Executed sendKey successfully on objectId: " & getObjectId(), FromProcedureName

End Function

Public Function clearValue() As Boolean
'----------------------------------------------------------------
' Clear the value currently had by the input element
'----------------------------------------------------------------

    Dim e1 As String
    Dim isReact As Boolean

    If checkExist Then If Not isExist Then Exit Function

   'Scroll into view
    main.jsEval "function(){ this.scrollIntoView(false) }", objectId:=getObjectId(), dbgMsg:=False

   'Reactフィールド検出
    e1 = main.jsEval("function(){ try { return Object.entries(this)[0][0] } catch(e) { return '' } }", objectId:=getObjectId(), dbgMsg:=False)
    If InStr(e1, "react") Or InStr(e1, "Tracker") Or InStr(e1, "State") Then isReact = True

   'Perform clearing using all the methods
    If Not isReact Then
        main.jsEval "function(){ this.value = '' }", objectId:=getObjectId(), dbgMsg:=printDbgMsg
    Else
        main.jsEval "function(){ Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(this, '') }", objectId:=getObjectId()
        main.jsEval "function(){ Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set.call(this, '') }", objectId:=getObjectId()
    End If

   'Most input elements need this
    fireEvent "input", printDbg:=False
    clearValue = True

End Function

Public Function isExist() As Boolean
'----------------------------------------------------------
' Easy way to check if the element is existing within the
' current page being referred to by the CDP object.
' Updated: 06/06/2023
'----------------------------------------------------------

    If elemVarType = vbNull Then isExist = False Else isExist = True

End Function

Public Function ifExist() As CDPElement
'----------------------------------------------------------
' Easy way to check if the element is existing within the
' current page being referred to by the CDP object. When the
' element does not exist, it returns an empty CDPElement
' object with checkExist tag = False, effectively skipping
' all subsequent functions related to it.
' Updated: 12/12/2023
'----------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".ifExist"


    checkExist = True
    Set ifExist = Me
    printMsg info_, "ifExist returned " & isExist & " on [" & varJS & "]." & IIf(Not isExist, " Subsequent actions on this element are disregarded.", ""), FromProcedureName

End Function

Public Function onExist(Optional timeOutInSeconds As Double = 30, Optional raiseTimeoutError As Boolean = True) As CDPElement
'----------------------------------------------------------
' Wait until the element applying this method comes into
' existance on the HTML page.
' Updated: 23/11/2023
'----------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".onExist"


    Dim timeStart As Double

    timeStart = TimerCounter
    Do While elemVarType = vbNull
        If TimerCounter - timeStart > timeOutInSeconds * 1000 Then
            printMsg ERROR_, "onExists timeout on varJS = " & varJS, FromProcedureName, doRaiseError:=raiseTimeoutError
            Exit Function
        End If
        elemVarType = VarType(main.jsEval(varJS, dbgMsg:=False))
        sleep 0.1
    Loop

    Set onExist = Me
    printMsg info_, "onExists satisfied: [" & varJS & "] is now available for interaction", FromProcedureName

End Function

Public Function onExistNot(Optional timeOutInSeconds As Double = 30, Optional raiseTimeoutError As Boolean = True) As Boolean
'----------------------------------------------------------
' Wait until the element applying this method comes into
' non-existance on the HTML page. Unlike onExists, this
' does not return anything as the element would be out of
' existence hence no further interactions should be needed.
' Updated: 23/11/2023 - make function returns True
'----------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".onExistNot"


    Dim timeStart As Double

    timeStart = TimerCounter
    Do While elemVarType <> vbNull
        If TimerCounter - timeStart > timeOutInSeconds * 1000 Then
            printMsg ERROR_, "onExistsNot timeout on varJS = " & varJS, FromProcedureName, doRaiseError:=raiseTimeoutError
            Exit Function
        End If
        elemVarType = VarType(main.jsEval(varJS, dbgMsg:=False))
        sleep 0.1
    Loop

    onExistNot = True
    printMsg info_, "onExistsNot satisfied: [" & varJS & "] is no longer available for interaction", FromProcedureName

End Function

'========================================================================================================
' NODE-TREE TRAVERSAL METHODS
'========================================================================================================

Public Function getParent() As CDPElement
'----------------------------------------------------------
' Get the parent node of the current element.
' Updated: 12/12/23
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Get the object and assign to the new ID in the browser
    Dim result As Variant
    result = main.jsEval("function(){ return this.parentElement }", objectId:=getObjectId(), returnByValue:=False, dbgMsg:=printDbgMsg)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getParent = New CDPElement
    getParent.Init main, , VarType(result), oid

End Function

Public Function getNextSibling() As CDPElement
'----------------------------------------------------------
' Get the next sibling node of the current element.
' Updated: 12/12/23
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Get the object and assign to the new ID in the browser
    Dim result As Variant
    result = main.jsEval("function(){ return this.nextElementSibling }", objectId:=getObjectId(), returnByValue:=False, dbgMsg:=printDbgMsg)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getNextSibling = New CDPElement
    getNextSibling.Init main, , VarType(result), oid

End Function

Public Function getPrevSibling() As CDPElement
'----------------------------------------------------------
' Get the previous sibling node of the current element.
' Updated: 24/10/2025: Daniel Polak - fixed bug where wrong assignment led to no return result
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Get the object and assign to the new ID in the browser
    Dim result As Variant
    result = main.jsEval("function(){ return this.previousElementSibling }", objectId:=getObjectId(), returnByValue:=False, dbgMsg:=printDbgMsg)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getPrevSibling = New CDPElement
    getPrevSibling.Init main, , VarType(result), oid

End Function

Public Function getFirstChild() As CDPElement
'----------------------------------------------------------
' Get the next sibling node of the current element.
' Updated: 12/12/2023
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Get the object and assign to the new ID in the browser
    Dim result As Variant
    result = main.jsEval("function(){ return this.firstElementChild }", objectId:=getObjectId(), returnByValue:=False, dbgMsg:=printDbgMsg)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getFirstChild = New CDPElement
    getFirstChild.Init main, , VarType(result), oid

End Function

Public Function getChildren() As Collection
'----------------------------------------------------------------------------------
' Obtain a collection of elements which are the child of the current element.
' This enables multiple getElementBy methods to be stringed together, ie. the varID
' of the current element is employed instead of 'document'.
' Updated: 10/11/2023
'----------------------------------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist And Not isExist Then
        Set getChildren = New Collection 'set an empty collection
        Exit Function
    End If

   'Get children HTMLCollection objectId via callFunctionOn
    Dim collResult As Variant
    collResult = main.jsEval("function(){ return this.children }", objectId:=getObjectId(), returnByValue:=False)
    If VarType(collResult) <> vbString Then Exit Function
    Dim collOid As String: collOid = CStr(collResult)

   'Get length via callFunctionOn
    Dim elements As Collection: Set elements = New Collection
    Dim length_ As Long, N As Long
    length_ = main.jsEval("function(){ return this.length }", objectId:=collOid)

   'Get each item's objectId via callFunctionOn
    Dim tElement As CDPElement
    Dim itemArgs As Collection, itemArg As Scripting.Dictionary
    Dim itemResult As Variant, itemOid As String
    For N = 0 To length_ - 1
        Set itemArgs = New Collection
        Set itemArg = New Scripting.Dictionary
        itemArg.Add "value", N
        itemArgs.Add itemArg
        itemResult = main.jsEval("function(n){ return this.item(n) }", objectId:=collOid, objectArguments:=itemArgs, returnByValue:=False)
        itemOid = "": If VarType(itemResult) = vbString Then itemOid = CStr(itemResult)
        Set tElement = New CDPElement
        tElement.Init main, , VarType(itemResult), itemOid
        elements.Add tElement
    Next

    Set getChildren = elements

End Function

Public Function getElementByID(strID As String) As CDPElement
'----------------------------------------------------------------------------------
' Obtain an element that is the child of the current element with the defined HTML
' ID. This enables multiple getElementBy methods to be stringed together. Thus for
' this to work, the varID of the current element is employed instead of 'document'.
' Updated: 12/12/2023
'----------------------------------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'getElementById は Document のみ。要素配下の ID 探索は querySelector + CSS.escape を使用
    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", strID
    args.Add arg
    Dim result As Variant
    result = main.jsEval("function(id){ return this.querySelector('#' + CSS.escape(id)) }", objectId:=getObjectId(), objectArguments:=args, returnByValue:=False)
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByID = New CDPElement
    getElementByID.Init main, , VarType(result), oid

End Function

Public Function getElementByQuery(strQuery As String) As CDPElement
'----------------------------------------------------------------------------------
' Obtain an element that is the child of the current element using QuerySelector.
' This enables multiple getElementBy methods to be stringed together. Thus for
' this to work, the varID of the current element is employed instead of 'document'.
' Updated: 12/12/2023
' Examples ------------------------------------------------------------------------
' eg.1 "#abc": get the first element with id attribute = abc
' eg.2 "tb.xyz": get the first element with tag tb and with class xyz
' eg.3 "a[name='abc']": get the first element with tag a and attribute name = abc
' eg.4 "div > p": get the first element with tag 'p' and has parent with tag 'div'
' eg.5 "[a1='abc'][a2='xyz']": get element with attribute a1 = abc and a2 = 123
' eg.6 "[a1='abc'] [a2='xyz']": get element with attribute a2 = xyz and is a child
' of an element with attribute a1 = abc
'----------------------------------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Get element by Query using callFunctionOn on this element
    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", strQuery
    args.Add arg
    Dim result As Variant
    result = main.jsEval("function(sel){ return this.querySelector(sel) }", objectId:=getObjectId(), objectArguments:=args, returnByValue:=False)
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByQuery = New CDPElement
    getElementByQuery.Init main, , VarType(result), oid

End Function

Public Function getElementsByQuery(strQuery As String) As Collection
'----------------------------------------------------------------------------------
' Obtain a collection of elements which are the child of the current element using
' QuerySelectorAll. This enables multiple getElementBy methods to be stringed
' together, ie. the varID of the current element is employed instead of 'document'.
' Updated: 12/12/2023
'----------------------------------------------------------------------------------

   'Get NodeList objectId via callFunctionOn on this element
    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", strQuery
    args.Add arg
    Dim collResult As Variant
    collResult = main.jsEval("function(sel){ return this.querySelectorAll(sel) }", objectId:=getObjectId(), objectArguments:=args, returnByValue:=False)
    If VarType(collResult) <> vbString Then Exit Function
    Dim collOid As String: collOid = CStr(collResult)

   'Get length and each item's objectId
    Dim elements As Collection: Set elements = New Collection
    Dim length_ As Long, N As Long
    Dim tElement As CDPElement
    Dim itemArgs As Collection, itemArg As Scripting.Dictionary
    Dim itemResult As Variant, itemOid As String
    length_ = main.jsEval("function(){ return this.length }", objectId:=collOid)
    For N = 0 To length_ - 1
        Set itemArgs = New Collection
        Set itemArg = New Scripting.Dictionary
        itemArg.Add "value", N
        itemArgs.Add itemArg
        itemResult = main.jsEval("function(n){ return this.item(n) }", objectId:=collOid, objectArguments:=itemArgs, returnByValue:=False)
        itemOid = "": If VarType(itemResult) = vbString Then itemOid = CStr(itemResult)
        Set tElement = New CDPElement
        tElement.Init main, , VarType(itemResult), itemOid
        elements.Add tElement
    Next

    Set getElementsByQuery = elements

End Function

Public Function getElementByXPath(strXPath As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class from XPath. The element class is init in CDPElement module.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' XPath Syntax help: https://www.w3schools.com/xml/xpath_syntax.asp
' Updated: 12/12/2023
' Examples -----------------------------------------------------------------------------
' eg.1 "//b[text()='T-BILL CURVE']": tag "b" with innertext to be 'T-BILL CURVE'
' eg.2 "//input[@name='username']": tag "input" & attribute "name" to be 'username'
' eg.3 "//a[@a1='abc' and @a2='xyz']": tag "a" with attribute a1 = abc and a2 = xyz
' eg.4 "//div[@id='abc']/div": get first div child of the div with id = abc
' eg.5 "//div[@id='abc'].//div": get first div descendant of the div with id = abc
' eg.6 "div": get first div child from the prescribed contextNode (default = 'document')
'---------------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".getElementByXPath"


   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Verify double slash syntax
    strXPath = Replace(strXPath, "\", "/")

   'Check if strXPath has invalid // syntax
    If Left(strXPath, 2) = "//" Then
        printMsg WARN_, "Note: '//' has been removed from the next XPath call string """ & strXPath & """ as it is an invalid syntax for the contextNode", FromProcedureName
        strXPath = Right(strXPath, Len(strXPath) - 2)
    End If

   'Get element by XPath using callFunctionOn (this element as contextNode)
    Dim args As New Collection
    Dim arg As New Scripting.Dictionary
    arg.Add "value", strXPath
    args.Add arg
    Dim result As Variant
    result = main.jsEval("function(xp){ return document.evaluate(xp, this, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue }", objectId:=getObjectId(), objectArguments:=args, returnByValue:=False)
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByXPath = New CDPElement
    getElementByXPath.Init main, , VarType(result), oid

End Function

Public Function getElementsByXPath(strXPath As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by XPath. For further info on the particular XPath
' syntax used for retrieving the elements, refer to Mozilla doc below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 12/12/2023
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".getElementsByXPath"


    Dim result As Variant

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Verify double slash syntax
    strXPath = Replace(strXPath, "\", "/")

   'Check if strXPath has invalid // syntax
    If Left(strXPath, 2) = "//" Then
        printMsg WARN_, "Note: '//' has been removed from the next XPath call string """ & strXPath & """ as it is an invalid syntax for the contextNode", FromProcedureName
        strXPath = Right(strXPath, Len(strXPath) - 2)
    End If

   'Get XPathResult objectId (ORDERED_NODE_SNAPSHOT_TYPE) using callFunctionOn (this as contextNode)
    Dim xpArgs As New Collection
    Dim xpArg As New Scripting.Dictionary
    xpArg.Add "value", strXPath
    xpArgs.Add xpArg
    Dim collResult As Variant
    collResult = main.jsEval("function(xp){ return document.evaluate(xp, this, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null) }", objectId:=getObjectId(), objectArguments:=xpArgs, returnByValue:=False)
    If VarType(collResult) <> vbString Then Exit Function
    Dim collOid As String: collOid = CStr(collResult)

   'Get snapshotLength and each item's objectId
    Dim elements As Collection: Set elements = New Collection
    Dim length_ As Long, N As Long
    Dim tElement As CDPElement
    Dim itemArgs As Collection, itemArg As Scripting.Dictionary
    Dim itemResult As Variant, itemOid As String
    length_ = main.jsEval("function(){ return this.snapshotLength }", objectId:=collOid)
    For N = 0 To length_ - 1
        Set itemArgs = New Collection
        Set itemArg = New Scripting.Dictionary
        itemArg.Add "value", N
        itemArgs.Add itemArg
        itemResult = main.jsEval("function(n){ return this.snapshotItem(n) }", objectId:=collOid, objectArguments:=itemArgs, returnByValue:=False)
        itemOid = "": If VarType(itemResult) = vbString Then itemOid = CStr(itemResult)
        Set tElement = New CDPElement
        tElement.Init main, , VarType(itemResult), itemOid
        elements.Add tElement
    Next

    Set getElementsByXPath = elements

End Function

Public Function getIFrame() As CDPElement
'----------------------------------------------------------------------------------
' Obtain the iFrame (inline frame) document of an iFrame HTML element. Some
' websites still use iFrame for form interactions with users so this method will be
' useful to quickly access the document object of the iFrame. iFrame cannot be
' directly interacted with getElement methods but must be accessed via its document
' object using .contentDocument. If an iFrame has a source, it is likely loaded not
' on the same domain. As a result, only indirect automation to the iFrame document
' may work. For further reference, visit w3 info:
' https://www.w3schools.com/jsref/prop_frame_contentdocument.asp
' Updated: 12/12/2023
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".getIFrame"


   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Check if the iFrame is hosted on a different domain
    Dim src As String: src = getAttribute("src")
    If InStr(src, "http") Then printMsg WARN_, "This iFrame may not be hosted in the same domain. src=""" & src & """", FromProcedureName

   'Get contentDocument objectId via callFunctionOn
    Dim result As Variant
    result = main.jsEval("function(){ return this.contentDocument }", objectId:=getObjectId(), returnByValue:=False)
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getIFrame = New CDPElement
    getIFrame.Init main, , VarType(result), oid

End Function

'***************************************************************************************************
'* 機能　　：現在の`ObjectID`配下のShadow-root(open/close)での、`ObjectID`の`CDPElement`を取得します
'---------------------------------------------------------------------------------------------------
'* 引数    ：Index  複数の`Shadow-root`があった際のIndex値。
'* 返り値　：`CDPElement`オブジェクト
'---------------------------------------------------------------------------------------------------
'* 機能説明：1. `getElementByXPath`等で、Shadow-root(open/close)手前の要素`CDPElement`を取得します
'            2. このプロシージャを実行して、Shadow-root(open/close)の中に、入り込みます
'            3. 再度、`getElementByXPath`等で、Shadow-root(open/close)内の要素を検索して、目的の操作を行います
'
'* 注意事項：・あくまでも、今の`CDPElement`直下の子Shadow-Root取得のみ対応します。事前にその手前での要素を取得の上、ご利用ください。
'            ・複数ある場合は、次項の`GetShadowRoots`をご利用ください。ただし、場所が固定されてるのがわかってる場合は、引数にて調整可能です
'            ・失敗時は、`nothing`となります。これを目印に、存在確認を行ってください
'***************************************************************************************************
Public Function GetShadowRoot(Optional Index As Long = 1) As CDPElement
    Const FromProcedureName As String = ThisClassName & ".GetShadowRoot"
    Dim paramsCDP As Dictionary, ResultCDP As Dictionary


    '1. `objectId`に対して、直下のNode情報を取得
    Set paramsCDP = New Dictionary
    paramsCDP.Add "objectId", elementObjectId
    paramsCDP.Add "depth", 0        'Shadow-root(open/close)手前前提により、`0`としておく
    paramsCDP.Add "pierce", True    'Shadow-root(open/close)に入るため、`True`にしておく
    Set ResultCDP = ExecuteCDP("DOM.describeNode", paramsCDP)

    '2. `shadowRoots`の存在確認
    If Not ResultCDP.Exists("node") Then printMsg WARN_, "No child nodes were found in the current `CDPElement`.", FromProcedureName: Exit Function
    If Not ResultCDP("node").Exists("shadowRoots") Then printMsg WARN_, "The `Shadow-Root` was not found in the child nodes of the current `CDPElement`.", FromProcedureName: Exit Function

    '3. 先頭の`shadowRoots`の`backendNodeId`を取得
    Dim shadowRootNodeId As Long
    shadowRootNodeId = ResultCDP("node")("shadowRoots")(Index)("backendNodeId")
    printMsg Debug_, "`Shadow-Root` was found. backendNodeId: " & shadowRootNodeId, FromProcedureName

    '4. 得た`backendNodeId`を`CDPElement`で使える`ObjectID`に変換
    paramsCDP.RemoveAll
    paramsCDP.Add "backendNodeId", shadowRootNodeId
    Set ResultCDP = ExecuteCDP("DOM.resolveNode", paramsCDP)

    '5. `Runtime.RemoteObject`の存在確認
    If Not ResultCDP.Exists("object") Then printMsg WARN_, "The existence of `Runtime.RemoteObject` could not be confirmed.", FromProcedureName: Exit Function
    If Not ResultCDP("object").Exists("objectId") Then printMsg WARN_, "Failed to retrieve the `ObjectID` in `Runtime.RemoteObject`.", FromProcedureName: Exit Function

    '6. `objectId`を取得
    Dim shadowRootObjectID As String
    shadowRootObjectID = ResultCDP("object")("objectId")
    printMsg Debug_, "`Successfully infiltrated `Shadow-Root`. objectId: " & shadowRootObjectID, FromProcedureName

    '7. `CDPElement`に必要な情報を継承して返却
    Set GetShadowRoot = New CDPElement
    GetShadowRoot.Init main, "", VarType(shadowRootObjectID), shadowRootObjectID
End Function

'***************************************************************************************************
'* 機能　　：`GetShadowRoot`プロシージャの複数取得版です。2つ以上になっちゃうときにご利用ください
'---------------------------------------------------------------------------------------------------
'* 返り値　：Collectionオブジェクト。中身は、`CDPElement`オブジェクトが詰まってます
'***************************************************************************************************
Public Function GetShadowRoots() As Collection
    Const FromProcedureName As String = ThisClassName & ".GetShadowRoots"
    Dim paramsCDP As Dictionary, ResultCDP As Dictionary


    '1. `objectId`に対して、直下のNode情報を取得
    Set paramsCDP = New Dictionary
    paramsCDP.Add "objectId", elementObjectId
    paramsCDP.Add "depth", 0        'Shadow-root(open/close)手前前提により、`0`としておく
    paramsCDP.Add "pierce", True    'Shadow-root(open/close)に入るため、`True`にしておく
    Set ResultCDP = ExecuteCDP("DOM.describeNode", paramsCDP)

    '2. `shadowRoots`の存在確認
    If Not ResultCDP.Exists("node") Then printMsg WARN_, "No child nodes were found in the current `CDPElement`.", FromProcedureName: Exit Function
    If Not ResultCDP("node").Exists("shadowRoots") Then printMsg WARN_, "The `Shadow-Root` was not found in the child nodes of the current `CDPElement`.", FromProcedureName: Exit Function

    '3. 各`shadowRoots`の`backendNodeId`を取得していく
    Set GetShadowRoots = New Collection
    printMsg Debug_, ResultCDP("node")("shadowRoots").Count & " `Shadow-Root` instances were found.", FromProcedureName
    Dim shadowRootsNodes
    For Each shadowRootsNodes In ResultCDP("node")("shadowRoots")
        '4. 各`shadowRoots`の`backendNodeId`を取得
        Dim shadowRootNodeId As Long
        shadowRootNodeId = shadowRootsNodes("backendNodeId")
        printMsg Debug_, "`Shadow-Root` - backendNodeId: " & shadowRootNodeId, FromProcedureName

        '5. 得た`backendNodeId`を`CDPElement`で使える`ObjectID`に変換
        paramsCDP.RemoveAll
        paramsCDP.Add "backendNodeId", shadowRootNodeId
        Set ResultCDP = ExecuteCDP("DOM.resolveNode", paramsCDP)

        '6. `Runtime.RemoteObject`の存在確認
        If Not ResultCDP.Exists("object") Then printMsg WARN_, "The existence of `Runtime.RemoteObject` could not be confirmed.", FromProcedureName: Exit Function
        If Not ResultCDP("object").Exists("objectId") Then printMsg WARN_, "Failed to retrieve the `ObjectID` in `Runtime.RemoteObject`.", FromProcedureName: Exit Function

        '7. `objectId`を取得
        Dim shadowRootObjectID As String
        shadowRootObjectID = ResultCDP("object")("objectId")
        printMsg Debug_, "`Successfully infiltrated `Shadow-Root`. objectId: " & shadowRootObjectID, FromProcedureName

        '8. `CDPElement`に必要な情報を継承して返却
        Dim GetShadowRoot As CDPElement: Set GetShadowRoot = New CDPElement
        GetShadowRoot.Init main, "", VarType(shadowRootObjectID), shadowRootObjectID

        '9. `Collection`に追加
        GetShadowRoots.Add GetShadowRoot
    Next
End Function



'***************************************************************************************************
'                             ■■■ CDPContext.cls 直通メソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    main.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Sub sleep(Optional seconds As Double = 0.5)
    main.InheritanceCDPBrowser.sleep seconds
End Sub

Private Function TimerCounter() As Double
    TimerCounter = main.InheritanceCDPBrowser.TimerCounter
End Function

Private Function ExecuteCDP(methodName As String, Optional params As Scripting.Dictionary, Optional StopCDPError As Boolean = True) As Scripting.Dictionary
    Set ExecuteCDP = main.ExecuteCDP(methodName, params, StopCDPError)
End Function



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
Private Sub Class_Terminate()
    '不要になった objectId をリリースします(あれば)
    If elementObjectId <> "" Then
        Dim paramsCDP As New Dictionary
        paramsCDP.Add "objectId", elementObjectId
        ExecuteCDP "Runtime.releaseObject", paramsCDP, False
    End If
End Sub
