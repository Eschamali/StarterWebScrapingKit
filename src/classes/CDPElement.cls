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
'       JavaScript操作の基となる`objectId`を使用した便利な様々な要素操作をまとめた物です。
'  主に`querySelector`や`Xpath`、`ID`といった検索で得た要素(objectId)に対する操作アシストを担います。
'
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



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const ThisClassName As String = "CDPElement"    'ログのトレース用

Public Enum keyboardCode    'Used for .sendKey operation in the browser class and element class
    keyEnter = 13
    keyDelete = 46
    keyTab = 9
    keyEsc = 27
    keyBackspace = 8
End Enum



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'コンテキスト(タブ)制御クラスオブジェクト
Private main As CDPContext

'JavaScript検索結果情報
Private Type ElementInfos
    objectId        As String       ' 検索ヒット時の`ObjectId`
    ResultVarType   As VbVarType    ' 検索ヒット時の返り値型(https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/vartype-function)
    SearchJS        As String       ' DOM検索で使ったJavaScript構文（ポーリング用）
    SearchArgs()    As Dictionary   ' DOM検索で使ったJavaScript構文の引数情報（ポーリング用）
End Type
Private CurrentElement  As ElementInfos     '`CDPContext.cls`または、このClass内で相対検索した際の検索結果情報用
Private ParentElement   As CDPElement       ' このClass内で相対検索した際の親の検索結果情報 ※デバッグ時、ツリー情報で辿れるメリットがあるので、`CDPElement.cls`本体を採用

'簡易check系
Private checkExist      As Boolean      ' チェーンメソッドを利用した簡易存在チェック

'ログ設定
Private currentLog As StateLog

'各種便利なJavaScript実行オプション
Private Type callFunctionOnParameters
    StopException    As Boolean         'JavaScript実行時、エラーで止めるか？   ※基本は開発中のみ`True`にすること
    RunAsyncCDP      As Boolean         '非同期実行
    userGesture      As Boolean         '人間が押したかのように偽装するか？
End Type
Private RunOption As callFunctionOnParameters



'***************************************************************************************************
'                                   ■■■ 継承処理 ■■■
'***************************************************************************************************
'* 機能　　：`CDPContext.cls`で検索実行したJavaScript結果を格納します
'---------------------------------------------------------------------------------------------------
'* 引数　　：InheritanceCDPContext  継承したい`CDPContext`そのもの
'            argResultVarType       検索実行結果の`VarType` ※`ObjectID`の場合は、`vbString`として扱います。
'            objectId_              検索実行結果時の`ObjectId`
'            JavaScriptStr          検索実行したJavaScriptコードそのもの。ポーリング用に使用します。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：要素がなく、`objectId`が返せない場合は、`objectId_`は、`StrPtr`チェックで0になります
'* 注意事項：・このヘルパークラスを使用する際は、`objectId`を返すようなJavaScriptコードとして実行すること。要素操作準備なら`objectId`でほぼほぼ返ると思いますが念押しです。
'            ・`JavaScriptStr`は、以下のシチュエーションの場合、無視され、ポーリング系のメソッドは使用できなくなります
'            　・複数取得系(getElementsByXPath等)
'              ・`GetShadowRoot`メソッドの使用
'***************************************************************************************************
Public Sub Init(InheritanceCDPContext As CDPContext, argResultVarType As VbVarType, objectid_ As String, Optional JavaScriptStr As String)
    'コンテキスト(タブ)情報を継承
    Set main = InheritanceCDPContext
    CDPHelpers.InheritLogSettings currentLog, InheritanceCDPContext.InheritanceCDPCore.LogSettings

    With CurrentElement
        '実行結果の型タイプ
        .ResultVarType = argResultVarType

        '検索実行時のJavaScript構文
        .SearchJS = JavaScriptStr

        '実行結果の`objectId`格納
        .objectId = objectid_
    End With
End Sub

Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    currentLog.LogID = "ELID" & Format(Rnd * 1000, "000")
End Sub



'***************************************************************************************************
'                                ■■■ パラメーター確認 ■■■
'***************************************************************************************************
'* 機能　　：この`CDPElement.cls`が作られた際の検索JavaScript構文を返却します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：要素の割り当てが成功した場合、このJavaScript構文をブラウザのコンソールウィンドウに直接入力することもできます
'***************************************************************************************************
Property Get UseSearchJS() As String
    UseSearchJS = CurrentElement.SearchJS
End Property

'***************************************************************************************************
'* 機能　　：この`CDPElement.cls`が作られた際の検索結果の`VarType`値を返します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：基本は`vbNull`or`vbString`を想定してます
'***************************************************************************************************
Property Get varResult() As VbVarType
    varResult = CurrentElement.ResultVarType
End Property

'***************************************************************************************************
'* 機能　　：このClassで保持中の`ObjectId`を、引数名として、DevToolsコンソールでも扱えるようにします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：なんかうまくいかない時、これを利用することで、DevToolsコンソール上でのデバッグが可能になります。
'* 注意事項：サイトによってはこの不自然なグローバル変数を検知するスクレイピング対策があります。ご利用時は注意してください。
'***************************************************************************************************
Public Sub ExposeDevTools(varName As String)
    Const FromProcedureName As String = ThisClassName & ".ExposeDevTools"


    ' 1. 関数に渡す引数の設定 (ユーザーが指定した変数名)
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", varName
    Set args(0) = arg

    ' 2. 実行
    jsEval "function(name) { window[name] = this; console.log('VBA: Exposed as ' + name); }", args
    printMsg info_, "In the DevTools console tab, `" & varName & "` was assigned.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：このClassで保持中の`ObjectId`をそのまま、取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：このClassで保持中の`ObjectId`
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`CDPElement`では足りないより高度な操作(CDPContext - jsEval等)をしたい場合にどうぞ
'* 注意事項：空の場合は、検索に引っかからなかったことを意味します。別途`onExist`メソッドで更新して下さい
'***************************************************************************************************
Property Get CurrentObjectId() As String
    CurrentObjectId = CurrentElement.objectId
End Property



'***************************************************************************************************
'                                   ■■■ 要素操作関連 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] クリック・送信
'---------------------------------------------------------------------------------------------------

'***************************************************************************************************
'* 機能　　：JavaScriptでのシンプルなクリックコマンドを実行します
'***************************************************************************************************
Public Function SimpleClick() As Boolean

    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function


    'シンプルなクリックコマンドを実行
    jsEval "function(){ this.click() }"

    '成功サインを返す
    SimpleClick = True

End Function

Public Function click() As Boolean
'----------------------------------------------------------
' Perform a standard click on the element. Remember to
' perform wait if necessary for the page to be refreshed.
' Updated: 12/12/2023
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function


   'Scroll into the view for better visibility of the automation
    jsEval "function(){ this.scrollIntoView(false) }"

   'Some fields require the browser to manually focus before allowing
   'the click action, such as the input type 'file' field.
    Dim params As New Dictionary
    params("text") = ""
    ExecuteCDP "Input.insertText", params

   'Perform click
    jsEval "function(){ this.dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true})) }"
    click = True

End Function

Public Function submit() As Boolean
'----------------------------------------------------------
' Submit a form element.
' Updated: 12/12/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function


    jsEval "function(){ this.form.submit() }"
    submit = True

End Function

Public Function sendClick() As Boolean
'----------------------------------------------------------------
' Attempt to simulate physical left mouse click input instead of
' DOM manipulation with JavaScript dispatchEvent method.
'----------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".sendClick"


    If checkExist Then If Not isExist Then Exit Function


    'Viewport 内に収めてから座標取得（CDP の x/y は CSS ピクセル・数値必須）
    Dim x As Double, y As Double
    jsEval "function(){ this.scrollIntoView(false) }"

    '中心座標として返却
    x = jsEval("function(){ const r = this.getBoundingClientRect(); return r.left + (r.width / 2); }")
    y = jsEval("function(){ const r = this.getBoundingClientRect(); return r.top + (r.height / 2); }")

    Dim params As New Dictionary
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
    printMsg info_, "Executed sendClick successfully on objectId: " & CurrentElement.objectId, FromProcedureName

End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 入力・クリア
'---------------------------------------------------------------------------------------------------
Public Function sendKey(Key As keyboardCode, Optional altKey As Boolean = False) As Boolean
'----------------------------------------------------------------
' Attempt to directly send a unique keyboard input to the element.
' Updated: 12/07/2024
'----------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".sendKey"


    If checkExist Then If Not isExist Then Exit Function


    'Set focus on the node
    Dim params As New Dictionary
    params("objectId") = CurrentElement.objectId
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
    printMsg info_, "Executed sendKey successfully on objectId: " & CurrentElement.objectId, FromProcedureName

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
    jsEval "function(){ this.value = '' }"

    'Set focus on the node
    Dim params As New Dictionary
    params("objectId") = CurrentElement.objectId
    ExecuteCDP "DOM.focus", params

    'Sendkeys to the node
    params.RemoveAll
    params("text") = textToSend
    ExecuteCDP "Input.insertText", params

    sendString = True
    printMsg info_, "Executed sendString(""" & textToSend & """) successfully on objectId: " & CurrentElement.objectId, FromProcedureName

End Function

Public Function clearValue() As Boolean
'----------------------------------------------------------------
' Clear the value currently had by the input element
'----------------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function


    Dim e1 As String
    Dim isReact As Boolean

    'Scroll into view
    jsEval "function(){ this.scrollIntoView(false) }"

    'Reactフィールド検出
    e1 = jsEval("function(){ try { return Object.entries(this)[0][0] } catch(e) { return '' } }")
    If InStr(e1, "react") Or InStr(e1, "Tracker") Or InStr(e1, "State") Then isReact = True

    'Perform clearing using all the methods
    If Not isReact Then
        jsEval "function(){ this.value = '' }"
    Else
        jsEval "function(){ Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(this, '') }"
        jsEval "function(){ Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set.call(this, '') }"
    End If

   'Most input elements need this
    fireEvent "input"
    clearValue = True

End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] フォーム系
'---------------------------------------------------------------------------------------------------
Property Let checked(boolChecked As Boolean)
'----------------------------------------------------------
' Change the checked value of a check box element
' Updated: 04/12/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", CBool(boolChecked)
    Set args(0) = arg
    jsEval "function(v){ this.checked = v }", args

End Property

Property Get checked() As Boolean
'----------------------------------------------------------
' Change the checked value of a check box element
' Updated: 04/12/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    checked = jsEval("function(){ return this.checked }")

End Property

Property Let selected(selectedOption As String)
'----------------------------------------------------------
' Change the selected value of an html select element type.
' eg. br.getElementbyId("list").selected = 2
' Updated: 06/03/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", selectedOption
    Set args(0) = arg
    jsEval "function(v){ this.selectedIndex = v }", args
    fireEvent "change"

End Property

Property Get selected() As String
'----------------------------------------------------------
' Get the selected option. Only returns the first option.
' Updated: 06/03/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    selected = jsEval("function(){ return this.selectedOptions[0] }")

End Property

Public Function setSelection(strOptionName As String)
'----------------------------------------------------------
' Set option value to a select element.
' Updated: 06/03/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function


    jsEval "function(){ this.scrollIntoView(false) }"
    Me.value = strOptionName
    fireEvent "change"

End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 属性操作
'---------------------------------------------------------------------------------------------------
Public Function getAttribute(strAttributeName As String) As String
'----------------------------------------------------------
' Get the value of an element attribute based on its name.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function


    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strAttributeName
    Set args(0) = arg
    getAttribute = jsEval("function(name){ return this.getAttribute(name) || '' }", args)

End Function

Public Function setAttribute(strAttributeName As String, strValue As String)
'----------------------------------------------------------
' Set value to an element attribute based on its name.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function


    Dim args2(1) As Dictionary
    Dim a1 As New Dictionary: a1.Add "value", strAttributeName:   Set args2(0) = a1
    Dim a2 As New Dictionary: a2.Add "value", strValue:           Set args2(1) = a2
    jsEval "function(name, val){ this.setAttribute(name, val) }", args2

End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 要素本体の内容・値
'---------------------------------------------------------------------------------------------------
Property Get value() As String
'----------------------------------------------------------
' Get the value property of the element in the HTML.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    value = jsEval("function(){ return this.value }")

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


    'Gracefully exit the value input if the element is found not exist under isExist
    If checkExist Then If Not isExist Then Exit Property


    Dim e1 As String
    Dim isReact As Boolean

    'Reactフィールド検出（thisコンテキスト内でObject.entriesを実行）
    e1 = jsEval("function(){ try { return Object.entries(this)[0][0] } catch(e) { return '' } }")
    If InStr(e1, "react") Or InStr(e1, "Tracker") Or InStr(e1, "State") Then isReact = True
    If isReact Then printMsg WARN_, "FYI: this is a REACT field as the first method is " & e1, FromProcedureName

    'objectArgumentsで安全に値を渡す（エスケープ不要）
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strTextVal
    Set args(0) = arg

    'Perform inputing using all the methods
    If Not isReact Then
        jsEval "function(v){ this.value = v }", args
    Else
        jsEval "function(v){ Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(this, v) }", args
        jsEval "function(v){ Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set.call(this, v) }", args
    End If

    'Most input elements need this
    fireEvent "input"

End Property

Property Get innerText() As String
'----------------------------------------------------------
' Get the innerText property of the element in the HTML.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    innerText = jsEval("function(){ return this.innerText }")

End Property

Property Let innerText(strTextVal As String)
'----------------------------------------------------------
' Change the innerText property of the element in the HTML.
' Note: objectArgumentsで安全に値を渡すのでエスケープ不要
' Updated: 02/03/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    jsEval "function(){ this.scrollIntoView(false) }"

    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strTextVal
    Set args(0) = arg
    jsEval "function(v){ this.innerText = v }", args

End Property

Property Get innerHTML() As String
'----------------------------------------------------------
' Get the innerHTML property of the element in the HTML.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    innerHTML = jsEval("function(){ return this.innerHTML }")

End Property

Property Let innerHTML(strTextVal As String)
'----------------------------------------------------------
' Change the innerHTML property of the element.
' Note: objectArgumentsで安全に値を渡すのでエスケープ不要
' Updated: 13/11/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    jsEval "function(){ this.scrollIntoView(false) }"

    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strTextVal
    Set args(0) = arg
    jsEval "function(v){ this.innerHTML = v }", args

End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] フォーカス・選択・イベント
'---------------------------------------------------------------------------------------------------
Public Function selectText()
'----------------------------------------------------------
' Select all the text held by the element.
' Updated: 18/09/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function


    jsEval "function(){ let r = document.createRange(); let s = window.getSelection(); r.selectNodeContents(this); s.removeAllRanges(); s.addRange(r) }"

End Function

Public Function focus()
'----------------------------------------------------------
' Set focus to an element.
' Updated: 18/09/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function


    jsEval "function(){ this.focus() }"

End Function

Public Function fireEvent(strEventName As String) As Boolean
'------------------------------------------------------------
' Dispatch the custom event on the element.
'------------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function


   'Fire the event; simulated=true is in case form is REACT 15.5 or lower
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strEventName
    Set args(0) = arg
    jsEval "function(ev){ this.dispatchEvent(new Event(ev, { bubbles: true, simulated: true })) }", args

    fireEvent = True

End Function

'***************************************************************************************************
'* 機能　　：要素に対する物理的なマウスホバーのシミュレートを行います
'***************************************************************************************************
Public Function sendHover() As Boolean
    Const FromProcedureName As String = ThisClassName & ".sendHover"


    If checkExist Then If Not isExist Then Exit Function


    'Viewport 内に収めてから座標取得（CDP の x/y は CSS ピクセル・数値必須）
    Dim x As Double, y As Double
    jsEval "function(){ this.scrollIntoView(false) }"

    '中心座標として返却
    x = jsEval("function(){ const r = this.getBoundingClientRect(); return r.left + (r.width / 2); }")
    y = jsEval("function(){ const r = this.getBoundingClientRect(); return r.top + (r.height / 2); }")

    '物理マウスをその要素の中心座標へ移動（これでホバー状態になる）
    Dim params As New Dictionary
    params("type") = "mouseMoved"
    params("x") = x
    params("y") = y
    ExecuteCDP "Input.dispatchMouseEvent", params

    sendHover = True
    printMsg info_, "Executed sendHover successfully on objectId: " & CurrentElement.objectId, FromProcedureName

End Function



'***************************************************************************************************
'                                   ■■■ ポーリング関連 ■■■
'***************************************************************************************************
'* 機能　　：各種要素操作が可能か判定します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：操作可能論理値
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`CDPContext-jsEval`によるXpath等で検索し存在すれば、`Runtime.RemoteObject`の`objectId`を返す仕組みを取ってます。
'            検索で見つからない場合は、`Null`で返す仕組みを取ってます。
'            この仕様を利用して、`vbNull`判定を利用することで、操作可能かの判定を素早く行っています。
'***************************************************************************************************
Public Function isExist() As Boolean
    isExist = Not (CurrentElement.ResultVarType = vbNull)
End Function

'***************************************************************************************************
'* 機能　　：メソッドチェーンを利用した存在チェック後の操作を提供します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：このClass自身
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・When the element does not exist, it returns an empty CDPElement
'               object with checkExist tag = False, effectively skipping
'               all subsequent functions related to it.
'
'            ・わざわざ、`if X.isExist then X.focus`といったことをせず、`X.ifExist.focus`のように簡略化できます。お好みで
'***************************************************************************************************
Public Function ifExist() As CDPElement
    Const FromProcedureName As String = ThisClassName & ".ifExist"


    checkExist = True
    Set ifExist = Me
    printMsg info_, "ifExist returned " & isExist & " on [" & CurrentElement.SearchJS & "]." & IIf(Not isExist, " Subsequent actions on this element are disregarded.", ""), FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`CDPContext.cls`で検索実行したJavaScriptを再実行し、有効なObjectIDが来るまで実行し続けます
'---------------------------------------------------------------------------------------------------
'* 返り値  ：見つかれば、`CDPElement`オブジェクトが返り、見つからない場合は、`nothing`です
'* 引数　　：timeOutInSeconds   検索し続ける秒数
'            raiseTimeoutError  タイムアウトしたら、エラーで止めるか？
'***************************************************************************************************
Public Function onExist(Optional timeOutInSeconds As Double = 30, Optional raiseTimeoutError As Boolean = True) As CDPElement
    Const FromProcedureName As String = ThisClassName & ".onExist"


    'カウントスタート
    Dim timeStart As Double: timeStart = CDPHelpers.TimerCounter

    '`vbNull`以外になるまで続ける
    '※このメソッド実行時点で既にある場合は、ループ探索には入りません
    Do While CurrentElement.ResultVarType = vbNull
        '1. タイムアウトチェック
        If CDPHelpers.TimerCounter - timeStart > timeOutInSeconds * 1000 Then
            printMsg ERROR_, "onExists timeout on SearchJS = " & CurrentElement.SearchJS, FromProcedureName, doRaiseError:=raiseTimeoutError
            Exit Function
        End If

        '2. 親情報に応じて検索方法を分岐
        Dim resultSearch As Variant
        If ParentElement Is Nothing Then
            '2-1. `CDPContext.cls`から検索した場合は、親情報なしなので、検索文を再実行
            resultSearch = main.jsEval(CurrentElement.SearchJS)
        Else
            '2-1. 引数情報に応じて、分岐
            '※`IsMissing`を機能させるため
            '判定の使用については右記にて→https://www.limecode.jp/entry/utility/check-redim-dynamic-array
            If Not Not CurrentElement.SearchArgs Then
                '引数情報があるので、引数付きで実行
                resultSearch = main.jsEval(CurrentElement.SearchJS, ParentElement.CurrentObjectId, CurrentElement.SearchArgs)
            Else
                resultSearch = main.jsEval(CurrentElement.SearchJS, ParentElement.CurrentObjectId)
            End If
        End If

        '3. 結果を判定。文字型なら、ObjectIDを確保
        CurrentElement.ResultVarType = VarType(resultSearch)
        If CurrentElement.ResultVarType = vbString Then CurrentElement.objectId = resultSearch Else CDPHelpers.Sleep 0.1
    Loop

    '自身を返却
    Set onExist = Me
    printMsg info_, "onExists satisfied: [" & CurrentElement.SearchJS & "] is now available for interaction", FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`CDPContext.cls`で検索実行したJavaScriptを再実行し、要素がnullになまで実行し続けます
'---------------------------------------------------------------------------------------------------
'* 返り値  ：なくなれば`True`、指定秒経ってもまだある場合は、`False`
'* 引数　　：timeOutInSeconds   検索し続ける秒数
'***************************************************************************************************
Public Function onExistNot(Optional timeOutInSeconds As Double = 30) As Boolean
    Const FromProcedureName As String = ThisClassName & ".onExistNot"


    'カウントスタート
    Dim timeStart As Double: timeStart = CDPHelpers.TimerCounter

    '`vbNull`になるまで続ける
    Do Until CurrentElement.ResultVarType = vbNull
        '1. タイムアウトになったら抜ける
        If CDPHelpers.TimerCounter - timeStart > timeOutInSeconds * 1000 Then printMsg WARN_, "onExistsNot timeout on SearchJS = " & CurrentElement.SearchJS, FromProcedureName: Exit Function

        '2. 検索を再実行
        Dim resultSearch As Boolean
        resultSearch = jsEval("function(){ return this.isConnected; }")

        '3. 結果を判定。存在しなかったら、ObjectIDをクリア
        CurrentElement.ResultVarType = IIf(resultSearch, vbBoolean, vbNull)
        If CurrentElement.ResultVarType = vbNull Then CurrentElement.objectId = vbNullString Else CDPHelpers.Sleep 0.1
    Loop

    '消えたサインを返却
    onExistNot = True
    printMsg info_, "onExistsNot satisfied: [" & CurrentElement.SearchJS & "] is no longer available for interaction", FromProcedureName

End Function



'***************************************************************************************************
'             ■■■ 今のNODE(CDPElement.cls)を基準にした相対検索 METHODS ■■■
'***************************************************************************************************
'* 機能　　：現在の要素の親ノードを取得します
'***************************************************************************************************
Public Function getParent() As CDPElement
    Const ParentSearchJS As String = "function(){ return this.parentElement }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If


    '1. 現在の要素の親ノードを取得するJavaScriptコードを実行
    Dim resultSearch As Variant
    resultSearch = jsEval(ParentSearchJS)

    '2. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '3. 結果を`CDPElement.cls`に格納
    Set getParent = New CDPElement
    getParent.Init main, JsResultVarType, objectId, ParentSearchJS

    '4. 元の親`CDPElement.cls`を適用
    Set getParent.InitParentCDPElement = Me
End Function

'***************************************************************************************************
'* 機能　　：現在の要素の次の兄弟ノードを取得します
'***************************************************************************************************
Public Function getNextSibling() As CDPElement
    Const NextSiblingSearchJS As String = "function(){ return this.nextElementSibling }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getNextSibling = Me
            Exit Function
        End If
    End If


    '1. 現在の要素の次の兄弟ノードを取得するJavaScriptコードを生成
    Dim resultSearch As Variant
    resultSearch = jsEval(NextSiblingSearchJS)

    '2. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '3. 結果を`CDPElement.cls`に格納
    Set getNextSibling = New CDPElement
    getNextSibling.Init main, JsResultVarType, objectId, NextSiblingSearchJS

    '4. 元の親`CDPElement.cls`を適用
    Set getNextSibling.InitParentCDPElement = Me
End Function

'***************************************************************************************************
'* 機能　　：現在の要素の前の兄弟ノードを取得します
'***************************************************************************************************
Public Function getPrevSibling() As CDPElement
    Const PrevSiblingSearchJS As String = "function(){ return this.previousElementSibling }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getPrevSibling = Me
            Exit Function
        End If
    End If


    '1. 現在の要素の前の兄弟ノードを取得するJavaScriptコードを生成
    Dim resultSearch As Variant
    resultSearch = jsEval(PrevSiblingSearchJS)

    '2. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '3. 結果を`CDPElement.cls`に格納
    Set getPrevSibling = New CDPElement
    getPrevSibling.Init main, JsResultVarType, objectId, PrevSiblingSearchJS

    '4. 元の親`CDPElement.cls`を適用
    Set getPrevSibling.InitParentCDPElement = Me
End Function

'***************************************************************************************************
'* 機能　　：現在の要素の先頭の兄弟ノードを取得します
'***************************************************************************************************
Public Function getFirstChild() As CDPElement
    Const FirstChildSearchJS As String = "function(){ return this.firstElementChild }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getFirstChild = Me
            Exit Function
        End If
    End If


    '1. 現在の要素の先頭の兄弟ノードを取得するJavaScriptコードを生成
    Dim resultSearch As Variant
    resultSearch = jsEval(FirstChildSearchJS)

    '2. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '3. 結果を`CDPElement.cls`に格納
    Set getFirstChild = New CDPElement
    getFirstChild.Init main, JsResultVarType, objectId, FirstChildSearchJS

    '4. 元の親`CDPElement.cls`を適用
    Set getFirstChild.InitParentCDPElement = Me
End Function

'***************************************************************************************************
'* 機能　　：現在の要素の子要素のコレクションを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：子要素のコレクション
'---------------------------------------------------------------------------------------------------
'* 詳細説明：これにより、複数の getElementBy メソッドを連結して使用できます
'* 注意事項：複数検索の場合、親情報を引き継がない関係上、返される`CDPElement.cls`でのポーリング関連のメソッドはご利用できません
'***************************************************************************************************
Public Function getChildren() As Collection
    Const ChildrenSearchJS As String = "function(){ return this.children }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist And Not isExist Then Exit Function


    '1. 現在の要素の子要素を取得するJavaScriptコードを生成
    Dim collResult As Variant
    collResult = jsEval(ChildrenSearchJS)

    '2. 結果を判定。文字型なら、ObjectIDを確保
    Dim collOid As String
    If VarType(collResult) = vbString Then
        collOid = collResult
    Else
        'ない場合は、即抜け
        Exit Function
    End If

    '3. Get length via callFunctionOn
    Set getChildren = New Collection
    Dim length_ As Long, N As Long
    length_ = main.jsEval("function(){ return this.length }", collOid)

    '4. Get each item's objectId via callFunctionOn
    Dim tElement As CDPElement
    Dim itemArgs(0) As Dictionary, itemArg As Dictionary
    Dim itemResult As Variant, itemObjectID As String, itemResultVarType As VbVarType
    For N = 0 To length_ - 1
        '4-1. `callFunctionOn`に渡す引数を初期化
        Set itemArg = New Dictionary

        '4-2. `callFunctionOn`に渡す引数をセット
        itemArg.Add "value", N
        Set itemArgs(0) = itemArg

        '4-3. 引数付きで実行
        itemResult = main.jsEval("function(n){ return this.item(n) }", collOid, itemArgs)

        '4-4. 念のため、有効な`ObjectID`かチェック
        itemResultVarType = VarType(itemResult)
        If itemResultVarType = vbString Then
            itemObjectID = itemResult

            '4-5. 結果を`CDPElement.cls`に格納
            Set tElement = New CDPElement
            tElement.Init main, itemResultVarType, itemObjectID
            getChildren.Add tElement
        End If
    Next
End Function

'***************************************************************************************************
'* 機能　　：現在の要素配下での、指定ID属性値を持つ子要素を取得します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：主に、`ShadowRoot`配下でのID検索で役立ちます
'***************************************************************************************************
Public Function getElementByID(strID As String) As CDPElement
    Const ElementByIDSearchJS As String = "function(id){ return this.querySelector('#' + CSS.escape(id)) }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getElementByID = Me
            Exit Function
        End If
    End If


    '1. `callFunctionOn`に渡す引数を準備
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strID
    Set args(0) = arg

    '2. 現在の要素配下での、指定ID属性値を持つ子要素を取得するJavaScriptコードを生成
    '※getElementById は Document のみ。要素配下の ID 探索は querySelector + CSS.escape を使用。これにより、`ShadowRoot`内検索も可能になります
    Dim resultSearch As Variant
    resultSearch = jsEval(ElementByIDSearchJS, args)

    '3. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '4. 結果を`CDPElement.cls`に格納
    Set getElementByID = New CDPElement
    getElementByID.Init main, JsResultVarType, objectId, ElementByIDSearchJS

    '5. 元の親`CDPElement.cls`を適用
    Set getElementByID.InitParentCDPElement(args) = Me
End Function

'***************************************************************************************************
'* 機能　　：QuerySelector を使用して、現在の要素の子要素を取得します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Examples
'             eg.1 "#abc": get the first element with id attribute = abc
'             eg.2 "tb.xyz": get the first element with tag tb and with class xyz
'             eg.3 "a[name='abc']": get the first element with tag a and attribute name = abc
'             eg.4 "div > p": get the first element with tag 'p' and has parent with tag 'div'
'             eg.5 "[a1='abc'][a2='xyz']": get element with attribute a1 = abc and a2 = 123
'             eg.6 "[a1='abc'] [a2='xyz']": get element with attribute a2 = xyz and is a child
'              of an element with attribute a1 = abc
'***************************************************************************************************
Public Function getElementByQuery(strQuery As String) As CDPElement
    Const ElementByQuerySearchJS As String = "function(sel){ return this.querySelector(sel) }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getElementByQuery = Me
            Exit Function
        End If
    End If


    '1. `callFunctionOn`に渡す引数を準備
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strQuery
    Set args(0) = arg

    '2. QuerySelector を使用した現在の要素の子要素を取得するJavaScriptコードを生成
    Dim resultSearch As Variant
    resultSearch = jsEval(ElementByQuerySearchJS, args)

    '3. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '4. 結果を`CDPElement.cls`に格納
    Set getElementByQuery = New CDPElement
    getElementByQuery.Init main, JsResultVarType, objectId, ElementByQuerySearchJS

    '5. 元の親`CDPElement.cls`を適用
    Set getElementByQuery.InitParentCDPElement(args) = Me
End Function

'***************************************************************************************************
'* 機能　　：QuerySelectorAll を使用して、現在の要素の子要素のコレクションを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：子要素のコレクション。1個もない場合は、`Nothing`
'* 引数　　：strQuery   検索したい子要素クエリセレクター
'---------------------------------------------------------------------------------------------------
'* 注意事項：複数検索の場合、親情報を引き継がない関係上、返される`CDPElement.cls`でのポーリング関連のメソッドはご利用できません
'***************************************************************************************************
Public Function getElementsByQuery(strQuery As String) As Collection
    Const ElementsByQuerySearchJS As String = "function(sel){ return this.querySelectorAll(sel) }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist And Not isExist Then Exit Function


    '1. `callFunctionOn`に渡す引数を準備
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strQuery
    Set args(0) = arg

    '2. QuerySelectorAll を使用して、現在の要素の子要素のコレクションを取得するJavaScriptコードを生成
    Dim collResult As Variant
    collResult = jsEval(ElementsByQuerySearchJS, args)

    '3. 結果を判定。文字型なら、ObjectIDを確保
    Dim collOid As String
    If VarType(collResult) = vbString Then
        collOid = collResult
    Else
        'ない場合は、即抜け
        Exit Function
    End If

    '3. Get length and each item's objectId
    Set getElementsByQuery = New Collection
    Dim length_ As Long, N As Long
    length_ = main.jsEval("function(){ return this.length }", collOid)

    '4. 前述の実行結果の数分、ループで順次取得していく
    Dim tElement As CDPElement
    Dim itemArgs(0) As Dictionary, itemArg As Dictionary
    Dim itemResult As Variant, itemObjectID As String, itemResultVarType As VbVarType
    For N = 0 To length_ - 1
        '4-1. `callFunctionOn`に渡す引数を初期化
        Set itemArg = New Dictionary

        '4-2. `callFunctionOn`に渡す引数をセット
        itemArg.Add "value", N
        Set itemArgs(0) = itemArg

        '4-3. 引数付きで実行
        itemResult = main.jsEval("function(n){ return this.item(n) }", collOid, itemArgs)

        '4-4. 念のため、有効な`ObjectID`かチェック
        itemResultVarType = VarType(itemResult)
        If itemResultVarType = vbString Then
            itemObjectID = itemResult

            '4-5. 結果を`CDPElement.cls`に格納
            Set tElement = New CDPElement
            tElement.Init main, itemResultVarType, itemObjectID
            getElementsByQuery.Add tElement
        End If
    Next
End Function

'***************************************************************************************************
'* 機能　　：現在の要素配下での、指定Xpath持つ子要素を取得します
'---------------------------------------------------------------------------------------------------
'* 引数　　：strXpath   検索したい子要素Xpath
'---------------------------------------------------------------------------------------------------
'* 詳細説明：The element class is init in CDPElement module.
'            The element class has useful properties and methods similar to the HTML element
'            object in previous IE automation.
'            XPath Syntax help: https://www.w3schools.com/xml/xpath_syntax.asp
' Examples -----------------------------------------------------------------------------
' eg.1 "//b[text()='T-BILL CURVE']": tag "b" with innertext to be 'T-BILL CURVE'
' eg.2 "//input[@name='username']": tag "input" & attribute "name" to be 'username'
' eg.3 "//a[@a1='abc' and @a2='xyz']": tag "a" with attribute a1 = abc and a2 = xyz
' eg.4 "//div[@id='abc']/div": get first div child of the div with id = abc
' eg.5 "//div[@id='abc'].//div": get first div descendant of the div with id = abc
' eg.6 "div": get first div child from the prescribed contextNode (default = 'document')
'---------------------------------------------------------------------------------------
'***************************************************************************************************
Public Function getElementByXPath(strXPath As String) As CDPElement
    Const FromProcedureName         As String = ThisClassName & ".getElementByXPath"
    Const ElementByXPathSearchJS    As String = "function(xp){ return document.evaluate(xp, this, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getElementByXPath = Me
            Exit Function
        End If
    End If


    '1. Verify double slash syntax
    'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")

    '2. Check if strXPath has invalid // syntax
    If Left(strXPath, 2) = "//" Then
        printMsg WARN_, "Note: '//' has been removed from the next XPath call string """ & strXPath & """ as it is an invalid syntax for the contextNode", FromProcedureName
        strXPath = Right(strXPath, Len(strXPath) - 2)
    End If

    '3. `callFunctionOn`に渡す引数を準備
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strXPath
    Set args(0) = arg

    '4. 現在の要素配下での､指定Xpath持つ子要素を取得するJavaScriptコードを生成
    Dim resultSearch As Variant
    resultSearch = jsEval(ElementByXPathSearchJS, args)

    '5. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '6. 結果を`CDPElement.cls`に格納
    Set getElementByXPath = New CDPElement
    getElementByXPath.Init main, JsResultVarType, objectId, ElementByXPathSearchJS

    '7. 元の親`CDPElement.cls`を適用
    Set getElementByXPath.InitParentCDPElement(args) = Me
End Function

'***************************************************************************************************
'* 機能　　：XPath を使用して要素のコレクションを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：子要素のコレクション。1個もない場合は、`Nothing`
'* 引数　　：strXpath   検索したい子要素Xpath
'---------------------------------------------------------------------------------------------------
'* 詳細説明：syntax used for retrieving the elements, refer to Mozilla doc below:
'             https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
'* 注意事項：・index for accessing the collection members start at one (not zero).
'            ・複数検索の場合、親情報を引き継がない関係上、返される`CDPElement.cls`でのポーリング関連のメソッドはご利用できません
'***************************************************************************************************
Public Function getElementsByXPath(strXPath As String) As Collection
    Const FromProcedureName As String = ThisClassName & ".getElementsByXPath"
    Const ChildrenSearchJS  As String = "function(xp){ return document.evaluate(xp, this, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null) }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist And Not isExist Then Exit Function


    '1. Verify double slash syntax
    strXPath = Replace(strXPath, "\", "/")

    '2.Check if strXPath has invalid // syntax
    If Left(strXPath, 2) = "//" Then
        printMsg WARN_, "Note: '//' has been removed from the next XPath call string """ & strXPath & """ as it is an invalid syntax for the contextNode", FromProcedureName
        strXPath = Right(strXPath, Len(strXPath) - 2)
    End If

    '3. `callFunctionOn`に渡す引数を準備
    Dim xpArgs(0) As Dictionary
    Dim xpArg As New Dictionary
    xpArg.Add "value", strXPath
    Set xpArgs(0) = xpArg

    '4. Get XPathResult objectId (ORDERED_NODE_SNAPSHOT_TYPE) using callFunctionOn (this as contextNode)
    Dim collResult As Variant
    collResult = jsEval(ChildrenSearchJS, xpArgs)

    '5. 結果を判定。文字型なら、ObjectIDを確保
    Dim collOid As String
    If VarType(collResult) = vbString Then
        collOid = collResult
    Else
        'ない場合は、即抜け
        Exit Function
    End If

    '6. Get snapshotLength and each item's objectId
    Set getElementsByXPath = New Collection
    Dim length_ As Long, N As Long
    length_ = main.jsEval("function(){ return this.snapshotLength }", collOid)

    '7. 前述の実行結果の数分、ループで順次取得していく
    Dim tElement As CDPElement
    Dim itemArgs(0) As Dictionary, itemArg As Dictionary
    Dim itemResult As Variant, itemObjectID As String, itemResultVarType As VbVarType
    For N = 0 To length_ - 1
        '7-1. `callFunctionOn`に渡す引数を初期化
        Set itemArg = New Dictionary

        '7-2. `callFunctionOn`に渡す引数をセット
        itemArg.Add "value", N
        Set itemArgs(0) = itemArg

        '7-3. 引数付きで実行
        itemResult = main.jsEval("function(n){ return this.snapshotItem(n) }", collOid, itemArgs)

        '7-4. 念のため、有効な`ObjectID`かチェック
        itemResultVarType = VarType(itemResult)
        If itemResultVarType = vbString Then
            itemObjectID = itemResult

            '7-5. 結果を`CDPElement.cls`に格納
            Set tElement = New CDPElement
            tElement.Init main, itemResultVarType, itemObjectID
            getElementsByXPath.Add tElement
        End If
    Next
End Function

'***************************************************************************************************
'* 機能　　：相対検索時の親となる`CDPElement.cls`を適用します
'---------------------------------------------------------------------------------------------------
'* 注意事項：相対検索の場合、検索文字列が「this. ...」という書き方になるため、検索情報となる引数が必要です
'***************************************************************************************************
Property Set InitParentCDPElement(Optional Arguments As Variant, Parent As CDPElement)
    Set ParentElement = Parent
    If Not IsMissing(Arguments) Then CurrentElement.SearchArgs = Arguments
End Property



'***************************************************************************************************
'                               ■■■ その他便利メソッド集 ■■■
'***************************************************************************************************
'* 機能　　：iFrame HTML要素のiFrame（インラインフレーム）ドキュメントを取得します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Some websites still use iFrame for form interactions with users so this method will be
'             useful to quickly access the document object of the iFrame. iFrame cannot be
'             directly interacted with getElement methods but must be accessed via its document
'             object using .contentDocument. If an iFrame has a source, it is likely loaded not
'             on the same domain. As a result, only indirect automation to the iFrame document
'             may work. For further reference, visit w3 info:
'              https://www.w3schools.com/jsref/prop_frame_contentdocument.asp
'
'* 注意事項：iframeのURLが親サイトと1文字でも違えば（例：Yahooに貼ってあるGoogleの広告など）、このメソッドは機能しません
'            その場合は、`CDPBrowser - getTab`か`CDPContext - getIFrameContextID → getIFrame`からご利用ください
'***************************************************************************************************
Public Function getIFrame() As CDPElement
    Const FromProcedureName As String = ThisClassName & ".getIFrame"
    Const getIFrameSearchJS As String = "function(){ return this.contentDocument }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getIFrame = Me
            Exit Function
        End If
    End If


    '1. Check if the iFrame is hosted on a different domain
    Dim src As String: src = getAttribute("src")
    If InStr(src, "http") Then printMsg WARN_, "This iFrame may not be hosted in the same domain. src=""" & src & """", FromProcedureName

    '2. Get contentDocument objectId via callFunctionOn
    Dim resultSearch As Variant
    resultSearch = jsEval(getIFrameSearchJS)

    '3. 結果を判定。文字型なら、ObjectIDを確保
    Dim objectId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then objectId = resultSearch

    '4. 結果を`CDPElement.cls`に格納
    Set getIFrame = New CDPElement
    getIFrame.Init main, JsResultVarType, objectId, getIFrameSearchJS

    '5. 元の親`CDPElement.cls`を適用
    Set getIFrame.InitParentCDPElement = Me
End Function

'***************************************************************************************************
'* 機能　　：現在の`ObjectID`配下のShadow-root(open/close)での、`ObjectID`の`CDPElement`を取得します
'---------------------------------------------------------------------------------------------------
'* 引数    ：Index  複数の`Shadow-root`があった際のIndex値。    ※先頭は、0始まり基準
'* 返り値　：`CDPElement`オブジェクト
'---------------------------------------------------------------------------------------------------
'* 機能説明：1. `getElementByXPath`等で、Shadow-root(open/close)手前の要素`CDPElement`を取得します
'            2. このプロシージャを実行して、Shadow-root(open/close)の中に、入り込みます
'            3. 再度、`getElementByQuery`等で、Shadow-root(open/close)内の要素を検索して、目的の操作を行います
'
'* 注意事項：・あくまでも、今の`CDPElement`直下の子Shadow-Root取得のみ対応します。事前にその手前での要素を取得の上、ご利用ください。
'            ・複数ある場合は、次項の`GetShadowRoots`をご利用ください。ただし、場所が固定されてるのがわかってる場合は、引数にて調整可能です
'            ・失敗時は、`nothing`となります。これを目印に、存在確認を行ってください
'            ・CDPによるShadowDOM侵入の都合上、親情報を引き継がないため、返される`CDPElement.cls`でのポーリング関連のメソッドはご利用できません
'***************************************************************************************************
Public Function GetShadowRoot(Optional Index As Long) As CDPElement
    Const FromProcedureName As String = ThisClassName & ".GetShadowRoot"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set GetShadowRoot = Me
            Exit Function
        End If
    End If


    '1. `objectId`に対して、直下のNode情報を取得
    Dim paramsCDP As Dictionary, ResultCDP As BiDiCDPJson
    Set paramsCDP = New Dictionary
    paramsCDP.Add "objectId", CurrentElement.objectId
    paramsCDP.Add "depth", 0        'Shadow-root(open/close)手前前提により、`0`としておく
    paramsCDP.Add "pierce", True    'Shadow-root(open/close)に入るため、`True`にしておく
    Set ResultCDP = ExecuteCDP("DOM.describeNode", paramsCDP)

    '2. `shadowRoots`の存在確認
    If Not ResultCDP.ExistsKey("node") Then printMsg WARN_, "No child nodes were found in the current `CDPElement`.", FromProcedureName: Exit Function
    If Not ResultCDP.NodeKey("node").ExistsKey("shadowRoots") Then printMsg WARN_, "The `Shadow-Root` was not found in the child nodes of the current `CDPElement`.", FromProcedureName: Exit Function

    '3. 指定Index圏内か？
    Dim Nodes_shadowRoots As BiDiCDPJson: Set Nodes_shadowRoots = ResultCDP("node").NodeKey("shadowRoots")
    If Nodes_shadowRoots.Count <= Index Then printMsg WARN_, "shadowRoots-Index out of range.", FromProcedureName: Exit Function

    '4. 指定位置の`shadowRoots`の`backendNodeId`を取得
    Dim shadowRootNodeId As Long
    shadowRootNodeId = Nodes_shadowRoots.NodeIndex(Index).NumberKey("backendNodeId")
    printMsg Debug_, "`Shadow-Root` was found. backendNodeId: " & shadowRootNodeId, FromProcedureName

    '5. 得た`backendNodeId`を`CDPElement`で使える`ObjectID`に変換
    paramsCDP.RemoveAll
    paramsCDP.Add "backendNodeId", shadowRootNodeId
    Set ResultCDP = ExecuteCDP("DOM.resolveNode", paramsCDP)

    '6. `Runtime.RemoteObject`の存在確認
    If Not ResultCDP.ExistsKey("object") Then printMsg WARN_, "The existence of `Runtime.RemoteObject` could not be confirmed.", FromProcedureName: Exit Function
    If Not ResultCDP.NodeKey("object").ExistsKey("objectId") Then printMsg WARN_, "Failed to retrieve the `ObjectID` in `Runtime.RemoteObject`.", FromProcedureName: Exit Function

    '7. `objectId`を取得
    Dim shadowRootObjectID As String
    shadowRootObjectID = ResultCDP.NodeKey("object").StringKey("objectId")
    printMsg info_, "`Successfully infiltrated `Shadow-Root`. objectId: " & shadowRootObjectID, FromProcedureName

    '8. `CDPElement`に必要な情報を継承して返却
    Set GetShadowRoot = New CDPElement
    GetShadowRoot.Init main, vbString, shadowRootObjectID   '※前述で、`Exists`チェックしてるので、`vbString`直代入する
End Function

'***************************************************************************************************
'* 機能　　：`GetShadowRoot`プロシージャの複数取得版です。2つ以上操作する際にどうぞ
'---------------------------------------------------------------------------------------------------
'* 返り値　：Collectionオブジェクト。中身は、`CDPElement`オブジェクトが詰まってます
'***************************************************************************************************
Public Function GetShadowRoots() As Collection
    Const FromProcedureName As String = ThisClassName & ".GetShadowRoots"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist And Not isExist Then Exit Function


    '1. `objectId`に対して、直下のNode情報を取得
    Dim paramsCDP As Dictionary, ResultCDP As BiDiCDPJson
    Set paramsCDP = New Dictionary
    paramsCDP.Add "objectId", CurrentElement.objectId
    paramsCDP.Add "depth", 0        'Shadow-root(open/close)手前前提により、`0`としておく
    paramsCDP.Add "pierce", True    'Shadow-root(open/close)に入るため、`True`にしておく
    Set ResultCDP = ExecuteCDP("DOM.describeNode", paramsCDP)

    '2. `shadowRoots`の存在確認
    If Not ResultCDP.ExistsKey("node") Then printMsg WARN_, "No child nodes were found in the current `CDPElement`.", FromProcedureName: Exit Function
    If Not ResultCDP.NodeKey("node").ExistsKey("shadowRoots") Then printMsg WARN_, "The `Shadow-Root` was not found in the child nodes of the current `CDPElement`.", FromProcedureName: Exit Function

    '3. `shadowRoots`の最初の`shadowRoots`要素のトークンIDを取得
    Dim Nodes_shadowRoots As BiDiCDPJson: Set Nodes_shadowRoots = ResultCDP.NodeKey("node").NodeKey("shadowRoots")
    Dim tId_shadowRoots As Long: tId_shadowRoots = Nodes_shadowRoots.FirstChildToken

    '4. 探索準備
    If tId_shadowRoots > 0 Then Set GetShadowRoots = New Collection
    printMsg Debug_, Nodes_shadowRoots.Count & " `Shadow-Root` instances were found.", FromProcedureName

    '5. 探索開始
    Do While tId_shadowRoots > 0
        '5-1. 各shadowRootの `backendNodeId` を直接取得
        ' NodeIndex(i) でオブジェクトを作らず、IDから直接数値を抜き出すのが最速です
        Dim shadowRootNodeId As Long
        shadowRootNodeId = Nodes_shadowRoots.TokenNumber(tId_shadowRoots, "backendNodeId")
        printMsg Debug_, "`Shadow-Root` - backendNodeId: " & shadowRootNodeId, FromProcedureName

        '5-2. 得た`backendNodeId`を`CDPElement`で使える`ObjectID`に変換
        Dim ResultRemoteObject As BiDiCDPJson
        paramsCDP.RemoveAll
        paramsCDP.Add "backendNodeId", shadowRootNodeId
        Set ResultRemoteObject = ExecuteCDP("DOM.resolveNode", paramsCDP).NodeKey("object")

        '5-3. 念のため、`objectId`の存在確認。ドキュメント上では、`optional`扱いのため
        If Not ResultRemoteObject.ExistsKey("objectId") Then printMsg WARN_, "Failed to retrieve the `ObjectID` in `Runtime.RemoteObject`.", FromProcedureName: Exit Function

        '5-4. 結果から、ObjectIDを取得
        Dim shadowRootObjectID As String
        shadowRootObjectID = ResultRemoteObject.StringKey("objectId")
        printMsg Debug_, "`Successfully infiltrated `Shadow-Root`. objectId: " & shadowRootObjectID, FromProcedureName

        '5-5. `CDPElement`に必要な情報を継承
        Dim GetShadowRoot As CDPElement: Set GetShadowRoot = New CDPElement
        GetShadowRoot.Init main, vbString, shadowRootObjectID '※前述で、`Exists`チェックしてるので、`vbString`直代入する

        '5-6. できあがった`CDPElement`をCollectionに蓄積
        GetShadowRoots.Add GetShadowRoot

        '5-7. 次の要素（隣のトークン）へ移動
        tId_shadowRoots = Nodes_shadowRoots.NextToken(tId_shadowRoots)
    Loop
End Function

'***************************************************************************************************
'* 機能　　：見つけた「<input type="file">」の`objectId`に対して、ファイル添付を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：files  ファイルパスが詰まってるCollectionを指定します
'***************************************************************************************************
Public Sub SetFileInputFiles(files As Collection)
    Const FromProcedureName As String = ThisClassName & ".setFileInputFiles"


    '1. 未登録の場合は、即抜け
    If files Is Nothing Then printMsg WARN_, "The file list is undefined.", FromProcedureName: Exit Sub
    If files.Count = 0 Then printMsg WARN_, "No files have been registered.", FromProcedureName: Exit Sub

    '2. パラメーターを組み立てて送信
    Dim params As New Dictionary
    params.Add "files", files
    params.Add "objectId", CurrentElement.objectId
    ExecuteCDP "DOM.setFileInputFiles", params

    '3. 別途ログ化
    If currentLog.seeRawSendMsgDbg Then
        printMsg info_, "Attachment complete: " & WebJsonConverter.serialize(files), FromProcedureName
    Else
        printMsg info_, "Attachment complete.", FromProcedureName
    End If
End Sub



'***************************************************************************************************
'                           ■■■ CDPHelpers ラッパーメソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub



'***************************************************************************************************
'                             ■■■ CDPContext.cls 直通メソッド ■■■
'***************************************************************************************************
Private Function ExecuteCDP(methodName As String, Optional params As Dictionary, Optional StopCDPError As Boolean = True) As BiDiCDPJson
    Set ExecuteCDP = main.ExecuteCDP(methodName, params, StopCDPError)
End Function

'***************************************************************************************************
'* 機能　　：見つけた`objectId`に対して、高度なことをしたい方向け
'---------------------------------------------------------------------------------------------------
'* 注意事項：・当然、`objectId`は固定してます
'            ・一部引数は、プロパティメソッドから変更する必要があります
'            ・Class内でもこれを経由して使用しています
'***************************************************************************************************
Public Function jsEval(JavaScriptStr As String, Optional objectArguments As Variant, Optional IFEXCEPTION As Variant, _
                       Optional returnByValue As Boolean, _
                       Optional awaitPromise As Boolean, _
                       Optional serializationOptions As Dictionary, _
                       Optional generatePreview As Boolean, _
                       Optional StopApiError As Boolean = True) As Variant
    Const FromProcedureName As String = ThisClassName & ".jsEval"


    '1. `objectId`がない場合は、エラー停止
    If StrPtr(CurrentElement.objectId) = 0 Then printMsg ERROR_, "This class does not have an ObjectID assigned.", FromProcedureName, doRaiseError:=True: Exit Function

    '2. 実行
    SetVar(jsEval) = main.jsEval(JavaScriptStr, CurrentElement.objectId, objectArguments, IFEXCEPTION, , _
                                    returnByValue, _
                                    awaitPromise, _
                                    RunOption.userGesture, _
                                    True, _
                                    serializationOptions, _
                                    generatePreview, _
                                    RunOption.RunAsyncCDP, _
                                    RunOption.StopException, _
                                    StopApiError)
End Function

'***************************************************************************************************
'* 機能　　：JavaScript実行結果にて`SET`or`LET`判定によるヘルパーメソッド
'***************************************************************************************************
Private Property Let SetVar(ByRef outVariable As Variant, ByRef inExpression As Variant)
    'IsObject で判定し、`Set`or`Let`で反映させる
    '※特定のオブジェクト(`vbDataObject`)では`IsObject`が`False`になるが、そういったのは返さないので省略
    If IsObject(inExpression) Then
        Set outVariable = inExpression
    Else
        Let outVariable = inExpression
    End If
End Property



'***************************************************************************************************
'                             ■■■ JavaScript実行設定 ■■■
'***************************************************************************************************
'* 機能　　：`True`で、JavaScriptコード内で例外が起こった場合、Err.Raiseで止まります。基本は、開発時のみ使用する想定です
'***************************************************************************************************
Property Let SetOptionStopException(v As Boolean)
    RunOption.StopException = v
End Property

'***************************************************************************************************
'* 機能　　：Trueで、結果を待たないCDPで実行します。
'---------------------------------------------------------------------------------------------------
'* 注意事項：必要な場面で使い終わったら必ず、`False`に戻すこと。面倒ごとになります
'***************************************************************************************************
Property Let SetOptionRunAsyncCDP(v As Boolean)
    RunOption.RunAsyncCDP = v
End Property

'***************************************************************************************************
'* 機能　　：True で「人間がクリックやキーボード操作をした結果」としてブラウザに偽装します。スクレピング対策が施されてる場合は、`True`にしておくといいかもしれません
'***************************************************************************************************
Property Let SetOptionUserGesture(v As Boolean)
    RunOption.userGesture = v
End Property



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
Private Sub Class_Terminate()
    '不要になった objectId をリリースします(あれば)
    '※毎回の同期解放はパフォーマンスが落ちるのでここは、非同期解放
    If StrPtr(CurrentElement.objectId) Then
        Dim paramsCDP As New Dictionary
        paramsCDP.Add "objectId", CurrentElement.objectId
        main.ExecuteCDPAsync "Runtime.releaseObject", paramsCDP, False
    End If
End Sub
