VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverBiDiElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
' WebDriver BiDi HTML Element Class
'------------------------------------------------------------------------------------------------
'       JavaScript操作の基となる`script.SharedId`を使用した便利な様々な要素操作をまとめた物です。
'  主に`querySelector`や`Xpath`、`ID`といった検索で得た要素(sharedId)に対する操作アシストを担います。
'  ※`CDPElement.cls`のWebDriver BiDi版です。CDP版との違いは以下の通りです
'    ・`getIFrame`は非搭載です。iframeは独立した`browsingContext`のため、`WebDriverBiDiMode - getTab(URL, maxDepth:=1)`で取得してください
'    ・`GetShadowRoots`は非搭載です。BiDiでは要素ごとに1つのShadow-rootのみ取得できるため、`GetShadowRoot`をご利用ください
'    ・`keyboardCode`/`NodeNodeType`列挙型は、`CDPElement.cls`側で定義しています
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
'       This class helps to simplify HTML element object referencing task using WebDriver BiDi.
'       This shall enables the development of 'getElementBy' methods that are commonly used for
'       automating browser sessions. Without this class, all interaction with the HTML elements
'       will have to be parsed via the jsEval method which can be counter-intuitive.
'================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const ThisClassName As String = "WebDriverBiDiElement"    'ログのトレース用



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'コンテキスト(タブ)制御クラスオブジェクト
Private main As WebDriverBiDiContext

'JavaScript検索結果情報
Private Type ElementInfos
    sharedId        As String       ' 検索ヒット時の`script.SharedId`
    ResultVarType   As VbVarType    ' 検索ヒット時の返り値型(https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/vartype-function)
    SearchJS        As String       ' DOM検索で使ったJavaScript構文（ポーリング用）
    SearchArgs()    As Dictionary   ' DOM検索で使ったJavaScript構文の引数情報（ポーリング用）
End Type
Private CurrentElement  As ElementInfos     '`WebDriverBiDiContext.cls`または、このClass内で相対検索した際の検索結果情報用
Private ParentElement   As WebDriverBiDiElement ' このClass内で相対検索した際の親の検索結果情報 ※デバッグ時、ツリー情報で辿れるメリットがあるので、`WebDriverBiDiElement.cls`本体を採用

'簡易check系
Private checkExist      As Boolean      ' チェーンメソッドを利用した簡易存在チェック

'ログ設定
Private m_LogID As String

'各種便利なJavaScript実行オプション
Private Type callFunctionParameters
    StopException    As Boolean         'JavaScript実行時、エラーで止めるか？   ※基本は開発中のみ`True`にすること
    RunAsyncBiDi     As Boolean         '非同期実行
    userGesture      As Boolean         '人間が押したかのように偽装するか？
End Type
Private RunOption As callFunctionParameters



'***************************************************************************************************
'                                   ■■■ 継承処理 ■■■
'***************************************************************************************************
'* 機能　　：`WebDriverBiDiContext.cls`で検索実行したJavaScript結果を格納します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ThisBiDiContext        継承したい`WebDriverBiDiContext`そのもの
'            argResultVarType       検索実行結果の`VarType` ※`script.SharedId`の場合は、`vbString`として扱います。
'            sharedId_              検索実行結果時の`script.SharedId`
'            JavaScriptStr          検索実行したJavaScriptコードそのもの。ポーリング用に使用します。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・`sharedId`は、`WebDriverBiDiContext.jsEval`でnodeを返した際に得られる物を想定しています
'            ・要素がなく、`sharedId`が返せない場合は、`sharedId_`は、`StrPtr`チェックで0になります
'* 注意事項：・このヘルパークラスを使用する際は、nodeを返すようなJavaScriptコードとして実行すること。要素操作準備なら`script.SharedId`でほぼほぼ返ると思いますが念押しです。
'            ・`JavaScriptStr`は、以下のシチュエーションの場合、無視され、ポーリング系のメソッドは使用できなくなります
'            　・複数取得系(getElementsByXPath等)
'              ・`GetShadowRoot`メソッドの使用
'***************************************************************************************************
Public Sub Init(ThisBiDiContext As WebDriverBiDiContext, argResultVarType As VbVarType, sharedId_ As String, Optional JavaScriptStr As String)
    'コンテキスト(タブ)情報を継承
    Set main = ThisBiDiContext

    With CurrentElement
        '実行結果の型タイプ
        .ResultVarType = argResultVarType

        '検索実行時のJavaScript構文
        .SearchJS = JavaScriptStr

        '実行結果の`script.SharedId`格納
        .sharedId = sharedId_
    End With
End Sub

Private Sub Class_Initialize()
    'ログID設定
    m_LogID = "ELID" & Format(Rnd * 1000, "000")
End Sub



'***************************************************************************************************
'                                ■■■ パラメーター確認 ■■■
'***************************************************************************************************
'* 機能　　：この`WebDriverBiDiElement.cls`が作られた際の検索JavaScript構文を返却します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：要素の割り当てが成功した場合、このJavaScript構文をブラウザのコンソールウィンドウに直接入力することもできます
'***************************************************************************************************
Property Get UseSearchJS() As String
    UseSearchJS = CurrentElement.SearchJS
End Property

'***************************************************************************************************
'* 機能　　：この`WebDriverBiDiElement.cls`が作られた際の検索結果の`VarType`値を返します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：基本は`vbNull`or`vbString`を想定してます
'***************************************************************************************************
Property Get varResult() As VbVarType
    varResult = CurrentElement.ResultVarType
End Property

'***************************************************************************************************
'* 機能　　：このClassで保持中の`script.SharedId`を、引数名として、DevToolsコンソールでも扱えるようにします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：なんかうまくいかない時、これを利用することで、DevToolsコンソール上でのデバッグが可能になります。
'* 注意事項：サイトによってはこの不自然なグローバル変数を検知するスクレイピング対策があります。ご利用時は注意してください。
'***************************************************************************************************
Public Sub ExposeDevTools(varName As String)
    Const FromProcedureName As String = ThisClassName & ".ExposeDevTools"


    ' 1. 関数に渡す引数の設定 (ユーザーが指定した変数名)
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", varName: arg.Add "type", "string"
    Set args(0) = arg

    ' 2. 実行
    jsEval "function(name) { window[name] = this; console.log('VBA: Exposed as ' + name); }", args
    printMsg info_, "In the DevTools console tab, `" & varName & "` was assigned.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：このClassで保持中の`script.SharedId`をそのまま、取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：このClassで保持中の`script.SharedId`
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`WebDriverBiDiElement`では足りないより高度な操作(WebDriverBiDiContext - jsEval の`scriptSharedId`等)をしたい場合にどうぞ
'* 注意事項：・空の場合は、検索に引っかからなかったことを意味します。別途`onExist`メソッドで更新して下さい
'            ・`CDPElement`と同じ使い勝手にするため、プロパティ名は`CurrentObjectId`のままにしています
'***************************************************************************************************
Property Get CurrentElementId() As String
    CurrentElementId = CurrentElement.sharedId
End Property

'***************************************************************************************************
'* 機能　　：このClassで保持中の`script.SharedId`の`nodeType`値を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：このClassで保持中の`script.SharedId`の`nodeType`値
'***************************************************************************************************
Property Get CurrentNodeType() As NodeNodeType
    CurrentNodeType = jsEval("function(){ return this.nodeType; }")
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
    '`Input.insertText`相当がないため、JavaScriptの`focus`で代用
    jsEval "function(){ this.focus() }"

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


    'Viewport 内に収める
    jsEval "function(){ this.scrollIntoView(false) }"

    '要素の中心へ移動 → 左ボタン押下 → 解放　※`origin`に要素を指定すると、座標計算は不要
    Dim pointerActions As New Collection
    pointerActions.Add PointerMoveToThis()
    pointerActions.Add InputAction("pointerDown", "button", 0)
    pointerActions.Add InputAction("pointerUp", "button", 0)
    PerformActions "pointer", pointerActions

    sendClick = True
    printMsg info_, "Executed sendClick successfully on sharedId: " & CurrentElement.sharedId, FromProcedureName

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


    'フォーカスはJavaScriptで当て、キーは`input.performActions`で押下 → 解放
    jsEval "function(){ this.focus() }"

    Dim keyValue As String: keyValue = ConvertKeyValue(Key)
    Dim keyActions As New Collection
    If altKey Then keyActions.Add InputAction("keyDown", "value", ChrW(&HE00A))   'Alt
    keyActions.Add InputAction("keyDown", "value", keyValue)
    keyActions.Add InputAction("keyUp", "value", keyValue)
    If altKey Then keyActions.Add InputAction("keyUp", "value", ChrW(&HE00A))     'Alt
    PerformActions "key", keyActions

    sendKey = True
    printMsg info_, "Executed sendKey successfully on sharedId: " & CurrentElement.sharedId, FromProcedureName

End Function

Public Function sendString(textToSend As String) As Boolean
'----------------------------------------------------------------
' Attempt to directly send the text string to the element via
' backdoor interaction with WebDriver BiDi instead of manipulating the DOM.
' Updated: 28/04/2023
'----------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".sendString"


    If checkExist Then If Not isExist Then Exit Function


    'First clear the current text of the field
    jsEval "function(){ this.value = '' }"

    '`Input.insertText`相当がないため、1文字ずつキーを押下 → 解放する(全文字を1回の`input.performActions`で送信)
    jsEval "function(){ this.focus() }"

    Dim keyActions As New Collection
    Dim i As Long, oneChar As String
    i = 1
    Do While i <= Len(textToSend)
        '1. 1文字取り出す　※サロゲートペア(上位)の場合は、下位とセットで1文字とする
        oneChar = Mid$(textToSend, i, 1)
        If (AscW(oneChar) And &HFC00&) = &HD800& Then oneChar = Mid$(textToSend, i, 2)

        '2. 押下 → 解放
        keyActions.Add InputAction("keyDown", "value", oneChar)
        keyActions.Add InputAction("keyUp", "value", oneChar)
        i = i + Len(oneChar)
    Loop
    If keyActions.Count > 0 Then PerformActions "key", keyActions

    sendString = True
    printMsg info_, "Executed sendString(""" & textToSend & """) successfully on sharedId: " & CurrentElement.sharedId, FromProcedureName

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
    arg.Add "value", CBool(boolChecked): arg.Add "type", "boolean"
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

'***************************************************************************************************
'* 機能　　：<select> タグの項目を選択します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Value    option の value 属性値で指定する場合（例: Value:="13"）
'  　　　　　Label    画面に見えている表示文字で指定する場合（例: Label:="東京都"）
'  　　　　　Index    0始まりの番号で指定する場合（例: Index:=1）
'---------------------------------------------------------------------------------------------------
'* 機能説明：名前付き引数なしの1つの場合は、`Value`側として扱います
'* 注意事項：どれか【1つだけ】を名前付き引数で指定してください（重複指定はエラーになります）
'***************************************************************************************************
Public Sub SelectOption(Optional value As String, Optional Label As String, Optional Index As Long = -1)
    Const FromProcedureName As String = ThisClassName & ".SelectOption"


    If checkExist Then If Not isExist Then Exit Sub


    ' --- 1. 【実質1択チェック】指定された個数を数える ---
    Dim targetVal As String, searchMode As SelectOptionMode
    Dim specifiedCount As Long
    If StrPtr(value) Then specifiedCount = specifiedCount + 1:  targetVal = value: searchMode = SelectOptionMode.seValue
    If StrPtr(Label) Then specifiedCount = specifiedCount + 1:  targetVal = Label: searchMode = SelectOptionMode.seLabel
    If Index > -1 Then specifiedCount = specifiedCount + 1:     targetVal = Index: searchMode = SelectOptionMode.seIndex

    If specifiedCount > 1 Then printMsg ERROR_, "This procedure accepts only one argument.", FromProcedureName, doRaiseError:=True: Exit Sub
    If specifiedCount = 0 Then printMsg ERROR_, "Please specify one of the arguments.", FromProcedureName, doRaiseError:=True: Exit Sub


    ' --- 2. 実行用 JavaScript を組み立てる ---
    '2-1. 反映タイプに応じて、Selectに反映させるJavaScript文を生成
    '※存在しない物を選ぶを例外エラーで止まります
    Dim UseJs As String
    Select Case searchMode
        Case SelectOptionMode.seValue
            ' value 属性値で探して選択する
            UseJs = "function(val) {" & _
                    "  const opt = Array.from(this.options).find(o => o.value === String(val));" & _
                    "  if (!opt) throw new Error('Option with value \' + val + \' not found.');" & _
                    "  opt.selected = true;" & _
                    "}"

        Case SelectOptionMode.seLabel
            ' 表示文字（前後の空白トリム付き）で探して選択する
            UseJs = "function(txt) {" & _
                    "  const opt = Array.from(this.options).find(o => o.text.trim() === String(txt));" & _
                    "  if (!opt) throw new Error('Option with label \' + txt + \' not found.');" & _
                    "  opt.selected = true;" & _
                    "}"

        Case SelectOptionMode.seIndex
            ' 0始まりのインデックス番号で直接選択する
            UseJs = "function(idx) {" & _
                    "  const i = parseInt(idx, 10);" & _
                    "  if (i < 0 || i >= this.options.length) throw new Error('Index ' + i + ' is out of bounds (length: ' + this.options.length + ').');" & _
                    "  this.selectedIndex = i;" & _
                    "}"
    End Select

    '2-2. 引数部分を組み立て
    Dim args(0) As Dictionary
    Dim arg As New Dictionary: arg.Add "value", targetVal: arg.Add "type", "string": Set args(0) = arg

    '3. 反映
    '※存在しない選択肢の場合は、エラー停止します
    '※BiDiの例外情報(`script.ExceptionDetails`)は、`text`にメッセージが入ります
    If IsError(jsEval(UseJs, args)) Then printMsg ERROR_, "No candidates were found for the specified SelectOption. > " & main.LastJavaScriptException.StringKey("text"), FromProcedureName

    '4. SPAにも対応すべく、`change`イベントを発火
    fireEvent "change"
End Sub

'***************************************************************************************************
'* 機能　　：<select> タグで、現在選択中の項目を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：mode に応じた値　※未選択の場合は、seValue/seLabel：空文字、seIndex：-1
'* 引数　　：mode     seValue：option の value 属性値
'  　　　　　         seLabel：画面に見えている表示文字（前後の空白トリム）
'  　　　　　         seIndex：0始まりの番号
'---------------------------------------------------------------------------------------------------
'* 機能説明：`SelectOption`で選んだ項目を、同じ基準(Value/Label/Index)で確認できます
'* 注意事項：複数選択(multiple)の場合は、先頭の選択項目を返します
'***************************************************************************************************
Public Function getSelection(mode As SelectOptionMode) As Variant
    '※JavaScript側の`m`は、`SelectOptionMode`の値(0:seValue / 1:seLabel / 2:seIndex)
    Const GetSelectionJS As String = "function(m) {" & _
                                     "  if (m === 2) return this.selectedIndex;" & _
                                     "  const opt = this.selectedOptions[0];" & _
                                     "  if (!opt) return '';" & _
                                     "  return (m === 1) ? opt.text.trim() : opt.value;" & _
                                     "}"


    If checkExist Then If Not isExist Then Exit Function


    '1. 引数部分を組み立て
    Dim args(0) As Dictionary
    Dim arg As New Dictionary: arg.Add "value", mode: arg.Add "type", "number": Set args(0) = arg

    '2. 取得して、Indexは数値、それ以外は文字列で返却
    Dim result As Variant
    result = jsEval(GetSelectionJS, args)
    If mode = SelectOptionMode.seIndex Then getSelection = CLng(result) Else getSelection = CStr(result)
End Function

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
Property Get PageAttribute(strAttributeName As String) As String
'----------------------------------------------------------
' Get the value of an element attribute based on its name.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strAttributeName: arg.Add "type", "string"
    Set args(0) = arg
    PageAttribute = jsEval("function(name){ return this.getAttribute(name) || '' }", args)

End Property

Property Let PageAttribute(strAttributeName As String, strValue As String)
'----------------------------------------------------------
' Set value to an element attribute based on its name.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property


    Dim args2(1) As Dictionary
    Dim a1 As New Dictionary: a1.Add "value", strAttributeName: a1.Add "type", "string":   Set args2(0) = a1
    Dim a2 As New Dictionary: a2.Add "value", strValue: a2.Add "type", "string":           Set args2(1) = a2
    jsEval "function(name, val){ this.setAttribute(name, val) }", args2

End Property

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
' For React fields, we need to input value to the element prototype in order to trigger the
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
    arg.Add "value", strTextVal: arg.Add "type", "string"
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
    arg.Add "value", strTextVal: arg.Add "type", "string"
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
    arg.Add "value", strTextVal: arg.Add "type", "string"
    Set args(0) = arg
    jsEval "function(v){ this.innerHTML = v }", args

End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] フォーカス・選択・イベント
'---------------------------------------------------------------------------------------------------
Public Sub selectText()
'----------------------------------------------------------
' Select all the text held by the element.
' Updated: 18/09/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Sub


    jsEval "function(){ let r = document.createRange(); let s = window.getSelection(); r.selectNodeContents(this); s.removeAllRanges(); s.addRange(r) }"

End Sub

Public Sub focus()
'----------------------------------------------------------
' Set focus to an element.
' Updated: 18/09/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Sub


    jsEval "function(){ this.focus() }"

End Sub

Public Function fireEvent(strEventName As String) As Boolean
'------------------------------------------------------------
' Dispatch the custom event on the element.
'------------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function


   'Fire the event; simulated=true is in case form is REACT 15.5 or lower
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strEventName: arg.Add "type", "string"
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


    'Viewport 内に収める
    jsEval "function(){ this.scrollIntoView(false) }"

    '要素の中心へ移動するだけ（これでホバー状態になる）
    Dim pointerActions As New Collection
    pointerActions.Add PointerMoveToThis()
    PerformActions "pointer", pointerActions

    sendHover = True
    printMsg info_, "Executed sendHover successfully on sharedId: " & CurrentElement.sharedId, FromProcedureName

End Function



'***************************************************************************************************
'                                   ■■■ ポーリング関連 ■■■
'***************************************************************************************************
'* 機能　　：各種要素操作が可能か判定します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：操作可能論理値
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`WebDriverBiDiContext-jsEval`によるXpath等で検索し存在すれば、nodeの`script.SharedId`を返す仕組みを取ってます。
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
'* 詳細説明：・When the element does not exist, it returns an empty WebDriverBiDiElement
'               object with checkExist tag = False, effectively skipping
'               all subsequent functions related to it.
'
'            ・わざわざ、`if X.isExist then X.focus`といったことをせず、`X.ifExist.focus`のように簡略化できます。お好みで
'***************************************************************************************************
Public Function ifExist() As WebDriverBiDiElement
    Const FromProcedureName As String = ThisClassName & ".ifExist"


    checkExist = True
    Set ifExist = Me
    printMsg info_, "ifExist returned " & isExist & " on [" & CurrentElement.SearchJS & "]." & IIf(Not isExist, " Subsequent actions on this element are disregarded.", ""), FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`WebDriverBiDiContext.cls`で検索実行したJavaScriptを再実行し、有効な`script.SharedId`が来るまで実行し続けます
'---------------------------------------------------------------------------------------------------
'* 返り値  ：見つかれば、`WebDriverBiDiElement`オブジェクトが返り、見つからない場合は、`nothing`です
'* 引数　　：timeOutInSeconds   検索し続ける秒数
'            raiseTimeoutError  タイムアウトしたら、エラーで止めるか？
'***************************************************************************************************
Public Function onExist(Optional timeOutInSeconds As Double = 30, Optional raiseTimeoutError As Boolean = True) As WebDriverBiDiElement
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
            '2-1. `WebDriverBiDiContext.cls`から検索した場合は、親情報なしなので、検索文を再実行
            resultSearch = main.jsEval(CurrentElement.SearchJS)
        Else
            '2-1. 引数情報に応じて、分岐
            '※`IsMissing`を機能させるため
            '判定の使用については右記にて→https://www.limecode.jp/entry/utility/check-redim-dynamic-array
            If Not Not CurrentElement.SearchArgs Then
                '引数情報があるので、引数付きで実行
                resultSearch = main.jsEval(CurrentElement.SearchJS, ParentElement.CurrentElementId, CurrentElement.SearchArgs)
            Else
                resultSearch = main.jsEval(CurrentElement.SearchJS, ParentElement.CurrentElementId)
            End If
        End If

        '3. 結果を判定。文字型なら、`script.SharedId`を確保
        CurrentElement.ResultVarType = VarType(resultSearch)
        If CurrentElement.ResultVarType = vbString Then CurrentElement.sharedId = resultSearch Else CDPHelpers.Sleep 0.1
    Loop

    '自身を返却
    Set onExist = Me
    printMsg info_, "onExists satisfied: [" & CurrentElement.SearchJS & "] is now available for interaction", FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`WebDriverBiDiContext.cls`で検索実行したJavaScriptを再実行し、要素がnullになまで実行し続けます
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

        '3. 結果を判定。存在しなかったら、`script.SharedId`をクリア
        CurrentElement.ResultVarType = IIf(resultSearch, vbBoolean, vbNull)
        If CurrentElement.ResultVarType = vbNull Then CurrentElement.sharedId = vbNullString Else CDPHelpers.Sleep 0.1
    Loop

    '消えたサインを返却
    onExistNot = True
    printMsg info_, "onExistsNot satisfied: [" & CurrentElement.SearchJS & "] is no longer available for interaction", FromProcedureName

End Function



'***************************************************************************************************
'             ■■■ 今のNODE(WebDriverBiDiElement.cls)を基準にした相対検索 METHODS ■■■
'***************************************************************************************************
'* 機能　　：現在の要素の親ノードを取得します
'***************************************************************************************************
Public Function getParent() As WebDriverBiDiElement
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

    '2. 結果を判定。文字型なら、`script.SharedId`を確保
    Dim sharedId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then sharedId = resultSearch

    '3. 結果を`WebDriverBiDiElement.cls`に格納
    Set getParent = New WebDriverBiDiElement
    getParent.Init main, JsResultVarType, sharedId, ParentSearchJS

    '4. 元の親`WebDriverBiDiElement.cls`を適用
    Set getParent.InitParentWebDriverBiDiElement = Me
End Function

'***************************************************************************************************
'* 機能　　：現在の要素の次の兄弟ノードを取得します
'***************************************************************************************************
Public Function getNextSibling() As WebDriverBiDiElement
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

    '2. 結果を判定。文字型なら、`script.SharedId`を確保
    Dim sharedId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then sharedId = resultSearch

    '3. 結果を`WebDriverBiDiElement.cls`に格納
    Set getNextSibling = New WebDriverBiDiElement
    getNextSibling.Init main, JsResultVarType, sharedId, NextSiblingSearchJS

    '4. 元の親`WebDriverBiDiElement.cls`を適用
    Set getNextSibling.InitParentWebDriverBiDiElement = Me
End Function

'***************************************************************************************************
'* 機能　　：現在の要素の前の兄弟ノードを取得します
'***************************************************************************************************
Public Function getPrevSibling() As WebDriverBiDiElement
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

    '2. 結果を判定。文字型なら、`script.SharedId`を確保
    Dim sharedId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then sharedId = resultSearch

    '3. 結果を`WebDriverBiDiElement.cls`に格納
    Set getPrevSibling = New WebDriverBiDiElement
    getPrevSibling.Init main, JsResultVarType, sharedId, PrevSiblingSearchJS

    '4. 元の親`WebDriverBiDiElement.cls`を適用
    Set getPrevSibling.InitParentWebDriverBiDiElement = Me
End Function

'***************************************************************************************************
'* 機能　　：現在の要素の先頭の兄弟ノードを取得します
'***************************************************************************************************
Public Function getFirstChild() As WebDriverBiDiElement
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

    '2. 結果を判定。文字型なら、`script.SharedId`を確保
    Dim sharedId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then sharedId = resultSearch

    '3. 結果を`WebDriverBiDiElement.cls`に格納
    Set getFirstChild = New WebDriverBiDiElement
    getFirstChild.Init main, JsResultVarType, sharedId, FirstChildSearchJS

    '4. 元の親`WebDriverBiDiElement.cls`を適用
    Set getFirstChild.InitParentWebDriverBiDiElement = Me
End Function

'***************************************************************************************************
'* 機能　　：現在の要素の子要素のコレクションを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：子要素のコレクション
'---------------------------------------------------------------------------------------------------
'* 詳細説明：これにより、複数の getElementBy メソッドを連結して使用できます
'* 注意事項：複数検索の場合、親情報を引き継がない関係上、返される`WebDriverBiDiElement.cls`でのポーリング関連のメソッドはご利用できません
'***************************************************************************************************
Public Function getChildren() As Collection
    Const ChildrenSearchJS As String = "function(){ return this.children }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist And Not isExist Then Exit Function


    '1. 現在の要素の子要素を取得するJavaScriptコードを生成
    Dim collResult As Variant
    collResult = jsEval(ChildrenSearchJS, Ownership:=True)   'コレクションは`script.Handle`で受け取る

    '2. 結果を判定。文字型なら、コレクションの`script.Handle`を確保
    Dim collHandle As String
    If VarType(collResult) = vbString Then
        collHandle = collResult
    Else
        'ない場合は、即抜け
        Exit Function
    End If

    '3. Get length via script.callFunction
    Set getChildren = New Collection
    Dim length_ As Long, N As Long
    '※BiDiの`scriptSharedId`は要素専用のため、コレクション(`script.Handle`)は引数側で渡す
    Dim collRef As New Dictionary, collArgs(0) As Dictionary
    collRef.Add "handle", collHandle
    Set collArgs(0) = collRef
    length_ = main.jsEval("function(c){ return c.length }", , collArgs)

    '4. Get each item's sharedId via script.callFunction
    Dim tElement As WebDriverBiDiElement
    Dim itemArgs(1) As Dictionary, itemArg As Dictionary
    Set itemArgs(0) = collRef   '第1引数は、コレクション固定
    Dim itemResult As Variant, itemSharedId As String, itemResultVarType As VbVarType
    For N = 0 To length_ - 1
        '4-1. `script.callFunction`に渡す引数を初期化
        Set itemArg = New Dictionary

        '4-2. `script.callFunction`に渡す引数をセット
        itemArg.Add "value", N: itemArg.Add "type", "number"
        Set itemArgs(1) = itemArg

        '4-3. 引数付きで実行
        itemResult = main.jsEval("function(c, n){ return c.item(n) }", , itemArgs)   'nodeは`script.SharedId`で返る

        '4-4. 念のため、有効な`script.SharedId`かチェック
        itemResultVarType = VarType(itemResult)
        If itemResultVarType = vbString Then
            itemSharedId = itemResult

            '4-5. 結果を`WebDriverBiDiElement.cls`に格納
            Set tElement = New WebDriverBiDiElement
            tElement.Init main, itemResultVarType, itemSharedId
            getChildren.Add tElement
        End If
    Next
End Function

'***************************************************************************************************
'* 機能　　：現在の要素配下での、指定ID属性値を持つ子要素を取得します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：主に、`ShadowRoot`配下でのID検索で役立ちます
'***************************************************************************************************
Public Function getElementByID(strID As String) As WebDriverBiDiElement
    Const ElementByIDSearchJS As String = "function(id){ return this.querySelector('#' + CSS.escape(id)) }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getElementByID = Me
            Exit Function
        End If
    End If


    '1. `script.callFunction`に渡す引数を準備
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strID: arg.Add "type", "string"
    Set args(0) = arg

    '2. 現在の要素配下での、指定ID属性値を持つ子要素を取得するJavaScriptコードを生成
    '※getElementById は Document のみ。要素配下の ID 探索は querySelector + CSS.escape を使用。これにより、`ShadowRoot`内検索も可能になります
    Dim resultSearch As Variant
    resultSearch = jsEval(ElementByIDSearchJS, args)

    '3. 結果を判定。文字型なら、`script.SharedId`を確保
    Dim sharedId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then sharedId = resultSearch

    '4. 結果を`WebDriverBiDiElement.cls`に格納
    Set getElementByID = New WebDriverBiDiElement
    getElementByID.Init main, JsResultVarType, sharedId, ElementByIDSearchJS

    '5. 元の親`WebDriverBiDiElement.cls`を適用
    Set getElementByID.InitParentWebDriverBiDiElement(args) = Me
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
Public Function getElementByQuery(strQuery As String) As WebDriverBiDiElement
    Const ElementByQuerySearchJS As String = "function(sel){ return this.querySelector(sel) }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getElementByQuery = Me
            Exit Function
        End If
    End If


    '1. `script.callFunction`に渡す引数を準備
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strQuery: arg.Add "type", "string"
    Set args(0) = arg

    '2. QuerySelector を使用した現在の要素の子要素を取得するJavaScriptコードを生成
    Dim resultSearch As Variant
    resultSearch = jsEval(ElementByQuerySearchJS, args)

    '3. 結果を判定。文字型なら、`script.SharedId`を確保
    Dim sharedId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then sharedId = resultSearch

    '4. 結果を`WebDriverBiDiElement.cls`に格納
    Set getElementByQuery = New WebDriverBiDiElement
    getElementByQuery.Init main, JsResultVarType, sharedId, ElementByQuerySearchJS

    '5. 元の親`WebDriverBiDiElement.cls`を適用
    Set getElementByQuery.InitParentWebDriverBiDiElement(args) = Me
End Function

'***************************************************************************************************
'* 機能　　：QuerySelectorAll を使用して、現在の要素の子要素のコレクションを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：子要素のコレクション。1個もない場合は、`Nothing`
'* 引数　　：strQuery   検索したい子要素クエリセレクター
'---------------------------------------------------------------------------------------------------
'* 注意事項：複数検索の場合、親情報を引き継がない関係上、返される`WebDriverBiDiElement.cls`でのポーリング関連のメソッドはご利用できません
'***************************************************************************************************
Public Function getElementsByQuery(strQuery As String) As Collection
    Const ElementsByQuerySearchJS As String = "function(sel){ return this.querySelectorAll(sel) }"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist And Not isExist Then Exit Function


    '1. `script.callFunction`に渡す引数を準備
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strQuery: arg.Add "type", "string"
    Set args(0) = arg

    '2. QuerySelectorAll を使用して、現在の要素の子要素のコレクションを取得するJavaScriptコードを生成
    Dim collResult As Variant
    collResult = jsEval(ElementsByQuerySearchJS, args, Ownership:=True)   'コレクションは`script.Handle`で受け取る

    '3. 結果を判定。文字型なら、コレクションの`script.Handle`を確保
    Dim collHandle As String
    If VarType(collResult) = vbString Then
        collHandle = collResult
    Else
        'ない場合は、即抜け
        Exit Function
    End If

    '3. Get length and each item's sharedId
    Set getElementsByQuery = New Collection
    Dim length_ As Long, N As Long
    '※BiDiの`scriptSharedId`は要素専用のため、コレクション(`script.Handle`)は引数側で渡す
    Dim collRef As New Dictionary, collArgs(0) As Dictionary
    collRef.Add "handle", collHandle
    Set collArgs(0) = collRef
    length_ = main.jsEval("function(c){ return c.length }", , collArgs)

    '4. 前述の実行結果の数分、ループで順次取得していく
    Dim tElement As WebDriverBiDiElement
    Dim itemArgs(1) As Dictionary, itemArg As Dictionary
    Set itemArgs(0) = collRef   '第1引数は、コレクション固定
    Dim itemResult As Variant, itemSharedId As String, itemResultVarType As VbVarType
    For N = 0 To length_ - 1
        '4-1. `script.callFunction`に渡す引数を初期化
        Set itemArg = New Dictionary

        '4-2. `script.callFunction`に渡す引数をセット
        itemArg.Add "value", N: itemArg.Add "type", "number"
        Set itemArgs(1) = itemArg

        '4-3. 引数付きで実行
        itemResult = main.jsEval("function(c, n){ return c.item(n) }", , itemArgs)   'nodeは`script.SharedId`で返る

        '4-4. 念のため、有効な`script.SharedId`かチェック
        itemResultVarType = VarType(itemResult)
        If itemResultVarType = vbString Then
            itemSharedId = itemResult

            '4-5. 結果を`WebDriverBiDiElement.cls`に格納
            Set tElement = New WebDriverBiDiElement
            tElement.Init main, itemResultVarType, itemSharedId
            getElementsByQuery.Add tElement
        End If
    Next
End Function

'***************************************************************************************************
'* 機能　　：現在の要素配下での、指定Xpath持つ子要素を取得します
'---------------------------------------------------------------------------------------------------
'* 引数　　：strXpath   検索したい子要素Xpath
'---------------------------------------------------------------------------------------------------
'* 詳細説明：The element class is init in WebDriverBiDiElement module.
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
' eg.7 ".//button[@id='abc']": get first button descendant of the Shadow-root (`GetShadowRoot`で取得した`WebDriverBiDiElement`から実行)
'---------------------------------------------------------------------------------------
'* 注意事項：・この`WebDriverBiDiElement`での検索は、あくまでも「現在の要素を起点とした相対検索」です。TOP(document)からの検索は、`WebDriverBiDiContext.getElementByXPath`をご利用ください
'            ・先頭が`/`や`//`の場合は、先頭に`.`を自動付与します（例：`//div` → `.//div`＝配下の子孫）
'            ・`GetShadowRoot`で得た Shadow-root の`WebDriverBiDiElement`から実行する場合は、`./○○`(直下) や `.//○○`(配下の子孫) のように、先頭に`.`を付けてください
'              `.`無しの相対指定（eg.6の`div`等）は、Shadow-root内では意図通りに探せません
'            ・Shadow-root は XPath の起点にできないため、内部では最初の子ノードを起点に、先頭へ`.`を足して`..`(＝Shadow-root)として評価しています
'            ・複数取得版の`getElementsByXPath`も、同じ規則で Shadow-root 内の検索に対応しています
'            ・子要素が1つもない特殊なShadowDomには対応できません。その場合は従来通り、CSSルートで我慢してください
'***************************************************************************************************
Public Function getElementByXPath(strXPath As String) As WebDriverBiDiElement
    Const FromProcedureName         As String = ThisClassName & ".getElementByXPath"
    Const ElementByXPathSearchJS    As String = "function(xp){" & _
                                                "  var c = (this.nodeType === Node.DOCUMENT_FRAGMENT_NODE) ? this.firstChild : this;" & _
                                                "  var x = (this.nodeType === 11) ? '.' + xp : xp;" & _
                                                "  return document.evaluate(x, c, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue" & _
                                                "}"

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

    '2. 先頭が`/`の場合は、相対検索になるよう先頭に`.`を付与
    If Left$(strXPath, 1) = "/" Then
        printMsg WARN_, "Note: '.' has been prepended to the next XPath call string """ & strXPath & """ because `WebDriverBiDiElement` searches relative to the current element. Use `WebDriverBiDiContext` to search from the top of the document.", FromProcedureName
        strXPath = "." & strXPath
    End If

    '3. `script.callFunction`に渡す引数を準備
    Dim args(0) As Dictionary
    Dim arg As New Dictionary
    arg.Add "value", strXPath: arg.Add "type", "string"
    Set args(0) = arg

    '4. 現在の要素配下での､指定Xpath持つ子要素を取得するJavaScriptコードを生成
    Dim resultSearch As Variant
    resultSearch = jsEval(ElementByXPathSearchJS, args)

    '5. 結果を判定。文字型なら、`script.SharedId`を確保
    Dim sharedId As String, JsResultVarType As VbVarType
    JsResultVarType = VarType(resultSearch)
    If JsResultVarType = vbString Then sharedId = resultSearch

    '6. 結果を`WebDriverBiDiElement.cls`に格納
    Set getElementByXPath = New WebDriverBiDiElement
    getElementByXPath.Init main, JsResultVarType, sharedId, ElementByXPathSearchJS

    '7. 元の親`WebDriverBiDiElement.cls`を適用
    Set getElementByXPath.InitParentWebDriverBiDiElement(args) = Me
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
'            ・複数検索の場合、親情報を引き継がない関係上、返される`WebDriverBiDiElement.cls`でのポーリング関連のメソッドはご利用できません
'            ・`getElementByXPath`と同じく、現在の要素を起点とした相対検索です。先頭が`/`や`//`の場合は、先頭に`.`を自動付与します
'            ・`GetShadowRoot`で得た Shadow-root の`WebDriverBiDiElement`から実行する場合は、`./○○`(直下) や `.//○○`(配下の子孫) のように、先頭に`.`を付けてください
'              `.`無しの相対指定は、Shadow-root内では意図通りに探せません。詳細は`getElementByXPath`の注意事項を参照してください
'***************************************************************************************************
Public Function getElementsByXPath(strXPath As String) As Collection
    Const FromProcedureName As String = ThisClassName & ".getElementsByXPath"
    Const ChildrenSearchJS  As String = "function(xp){" & _
                                        "  var c = (this.nodeType === Node.DOCUMENT_FRAGMENT_NODE) ? this.firstChild : this;" & _
                                        "  var x = (this.nodeType === Node.DOCUMENT_FRAGMENT_NODE) ? '.' + xp : xp;" & _
                                        "  return document.evaluate(x, c, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)" & _
                                        "}"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist And Not isExist Then Exit Function


    '1. Verify double slash syntax
    strXPath = Replace(strXPath, "\", "/")

    '2. 先頭が`/`の場合は、相対検索になるよう先頭に`.`を付与
    If Left$(strXPath, 1) = "/" Then
        printMsg WARN_, "Note: '.' has been prepended to the next XPath call string """ & strXPath & """ because `WebDriverBiDiElement` searches relative to the current element. Use `WebDriverBiDiContext` to search from the top of the document.", FromProcedureName
        strXPath = "." & strXPath
    End If

    '3. `script.callFunction`に渡す引数を準備
    Dim xpArgs(0) As Dictionary
    Dim xpArg As New Dictionary
    xpArg.Add "value", strXPath: xpArg.Add "type", "string"
    Set xpArgs(0) = xpArg

    '4. Get XPathResult handle (ORDERED_NODE_SNAPSHOT_TYPE) using script.callFunction (this as contextNode)
    Dim collResult As Variant
    collResult = jsEval(ChildrenSearchJS, xpArgs, Ownership:=True)   'コレクションは`script.Handle`で受け取る

    '5. 結果を判定。文字型なら、コレクションの`script.Handle`を確保
    Dim collHandle As String
    If VarType(collResult) = vbString Then
        collHandle = collResult
    Else
        'ない場合は、即抜け
        Exit Function
    End If

    '6. Get snapshotLength and each item's sharedId
    Set getElementsByXPath = New Collection
    Dim length_ As Long, N As Long
    '※BiDiの`scriptSharedId`は要素専用のため、コレクション(`script.Handle`)は引数側で渡す
    Dim collRef As New Dictionary, collArgs(0) As Dictionary
    collRef.Add "handle", collHandle
    Set collArgs(0) = collRef
    length_ = main.jsEval("function(c){ return c.snapshotLength }", , collArgs)

    '7. 前述の実行結果の数分、ループで順次取得していく
    Dim tElement As WebDriverBiDiElement
    Dim itemArgs(1) As Dictionary, itemArg As Dictionary
    Set itemArgs(0) = collRef   '第1引数は、コレクション固定
    Dim itemResult As Variant, itemSharedId As String, itemResultVarType As VbVarType
    For N = 0 To length_ - 1
        '7-1. `script.callFunction`に渡す引数を初期化
        Set itemArg = New Dictionary

        '7-2. `script.callFunction`に渡す引数をセット
        itemArg.Add "value", N: itemArg.Add "type", "number"
        Set itemArgs(1) = itemArg

        '7-3. 引数付きで実行
        itemResult = main.jsEval("function(c, n){ return c.snapshotItem(n) }", , itemArgs)   'nodeは`script.SharedId`で返る

        '7-4. 念のため、有効な`script.SharedId`かチェック
        itemResultVarType = VarType(itemResult)
        If itemResultVarType = vbString Then
            itemSharedId = itemResult

            '7-5. 結果を`WebDriverBiDiElement.cls`に格納
            Set tElement = New WebDriverBiDiElement
            tElement.Init main, itemResultVarType, itemSharedId
            getElementsByXPath.Add tElement
        End If
    Next
End Function

'***************************************************************************************************
'* 機能　　：相対検索時の親となる`WebDriverBiDiElement.cls`を適用します
'---------------------------------------------------------------------------------------------------
'* 注意事項：相対検索の場合、検索文字列が「this. ...」という書き方になるため、検索情報となる引数が必要です
'***************************************************************************************************
Property Set InitParentWebDriverBiDiElement(Optional Arguments As Variant, Parent As WebDriverBiDiElement)
    Set ParentElement = Parent
    If Not IsMissing(Arguments) Then CurrentElement.SearchArgs = Arguments
End Property



'***************************************************************************************************
'                               ■■■ その他便利メソッド集 ■■■
'***************************************************************************************************
'* 機能　　：現在の`script.SharedId`配下のShadow-root(open/close)での、`script.SharedId`の`WebDriverBiDiElement`を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`WebDriverBiDiElement`オブジェクト
'---------------------------------------------------------------------------------------------------
'* 機能説明：1. `getElementByXPath`等で、Shadow-root(open/close)手前の要素`WebDriverBiDiElement`を取得します
'            2. このプロシージャを実行して、Shadow-root(open/close)の中に、入り込みます
'            3. 再度、`getElementByQuery`等で、Shadow-root(open/close)内の要素を検索して、目的の操作を行います
'
'* 注意事項：・あくまでも、今の`WebDriverBiDiElement`直下の子Shadow-Root取得のみ対応します。事前にその手前での要素を取得の上、ご利用ください。
'            ・BiDiでは要素ごとに1つのShadow-rootのみ取得できるため、`CDPElement`のような`Index`引数や複数取得版(`GetShadowRoots`)はありません
'            ・失敗時は、`nothing`となります。これを目印に、存在確認を行ってください
'            ・親情報を引き継がないため、返される`WebDriverBiDiElement.cls`でのポーリング関連のメソッドはご利用できません
'***************************************************************************************************
Public Function GetShadowRoot() As WebDriverBiDiElement
    Const FromProcedureName As String = ThisClassName & ".GetShadowRoot"


    'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set GetShadowRoot = Me
            Exit Function
        End If
    End If


    '1. 要素自身を、Shadow-rootを含めてシリアライズ　※`includeShadowTree:="all"`で、closedも含まれる
    Dim serializeOption As New Dictionary
    serializeOption.Add "maxDomDepth", 1
    serializeOption.Add "includeShadowTree", "all"

    Dim nodeResult As JsEvalResult
    Set nodeResult = jsEval("function(){ return this }", serializationOptions:=serializeOption, returnByDetails:=True)
    If nodeResult.isException Or nodeResult.ResultType <> "node" Then printMsg WARN_, "The current `WebDriverBiDiElement` could not be serialized as a node.", FromProcedureName: Exit Function

    '2. `shadowRoot`の存在確認　※ない場合は`null`
    Dim NodeValue As BiDiCDPJConv: Set NodeValue = nodeResult.rawValue
    Dim shadowRootNode As Variant
    If NodeValue.ExistsKey("shadowRoot") Then CDPHelpers.SetVar(shadowRootNode) = NodeValue.Item("shadowRoot")
    If Not IsObject(shadowRootNode) Then printMsg WARN_, "The `Shadow-Root` was not found in the current `WebDriverBiDiElement`.", FromProcedureName: Exit Function

    '3. `script.SharedId`を取得
    Dim shadowRootSharedId As String
    shadowRootSharedId = shadowRootNode.StringKey("sharedId")
    printMsg info_, "`Successfully infiltrated `Shadow-Root`. sharedId: " & shadowRootSharedId, FromProcedureName

    '4. `WebDriverBiDiElement`に必要な情報を継承して返却
    Set GetShadowRoot = New WebDriverBiDiElement
    GetShadowRoot.Init main, vbString, shadowRootSharedId   '※前述で、存在チェックしてるので、`vbString`直代入する
End Function

'***************************************************************************************************
'* 機能　　：見つけた「<input type="file">」の`script.SharedId`に対して、ファイル添付を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：files  ファイルパスが詰まってるCollectionを指定します
'***************************************************************************************************
Public Sub SetFileInputFiles(files As Collection)
    Const FromProcedureName As String = ThisClassName & ".setFileInputFiles"


    '1. 未登録の場合は、即抜け
    If files Is Nothing Then printMsg WARN_, "The file list is undefined.", FromProcedureName: Exit Sub
    If files.Count = 0 Then printMsg WARN_, "No files have been registered.", FromProcedureName: Exit Sub

    '2. パラメーターを組み立てて送信　※要素は`script.SharedReference`として指定
    Dim params As New Dictionary
    params.Add "files", files
    params.Add "element", SharedReference()
    ExecuteBiDi "input.setFiles", params

    '3. 別途ログ化
    If CDPHelpers.currentLog.seeRawSendMsgDbg Then
        printMsg info_, "Attachment complete: " & WebJsonConverter.serialize(files), FromProcedureName
    Else
        printMsg info_, "Attachment complete.", FromProcedureName
    End If
End Sub



'***************************************************************************************************
'                           ■■■ CDPHelpers ラッパーメソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".bidi.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, CDPHelpers.currentLog.logPath
End Sub



'***************************************************************************************************
'                         ■■■ WebDriverBiDiContext.cls 直通メソッド ■■■
'***************************************************************************************************
'* 注意事項：`context`は、`WebDriverBiDiContext.ExecuteBiDi`側で付与されます
'***************************************************************************************************
Private Function ExecuteBiDi(methodName As String, Optional params As Dictionary, Optional StopBiDiError As Boolean = True) As BiDiCDPJConv
    Set ExecuteBiDi = main.ExecuteBiDi(methodName, params, StopBiDiError)
End Function

'***************************************************************************************************
'* 機能　　：見つけた`script.SharedId`に対して、高度なことをしたい方向け
'---------------------------------------------------------------------------------------------------
'* 引数　　：objectArguments    `script.LocalValue`形式(`{"type": ..., "value": ...}`)で組み立ててください。変換は行いません
'            Ownership          `WebDriverBiDiContext.jsEval`の`Ownership`と同じです
'                               `False`(既定)：nodeは`script.SharedId`、それ以外のオブジェクトは値(シリアライズ結果)で返ります
'                               `True`      ：nodeは`script.SharedId`、それ以外のオブジェクトは`script.Handle`で返ります
'            returnByDetails    `True`で、`JsEvalResult`(JavaScript実行結果の詳細オブジェクト)として返却します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・当然、`script.SharedId`は固定してます
'            ・一部引数は、プロパティメソッドから変更する必要があります
'            ・Class内でもこれを経由して使用しています
'***************************************************************************************************
Public Function jsEval(JavaScriptStr As String, Optional objectArguments As Variant, Optional IFEXCEPTION As Variant, _
                       Optional Ownership As Boolean, _
                       Optional awaitPromise As Boolean, _
                       Optional serializationOptions As Dictionary, _
                       Optional returnByDetails As Boolean, _
                       Optional StopApiError As Boolean = True) As Variant
    Const FromProcedureName As String = ThisClassName & ".jsEval"


    '1. `script.SharedId`がない場合は、エラー停止
    If StrPtr(CurrentElement.sharedId) = 0 Then printMsg ERROR_, "This class does not have an sharedId assigned.", FromProcedureName, doRaiseError:=True: Exit Function

    '2. 実行
    CDPHelpers.SetVar(jsEval) = main.jsEval(JavaScriptStr, CurrentElement.sharedId, objectArguments, IFEXCEPTION, , _
                                    Ownership, _
                                    awaitPromise, _
                                    RunOption.userGesture, _
                                    serializationOptions, _
                                    returnByDetails, _
                                    RunOption.RunAsyncBiDi, _
                                    RunOption.StopException, _
                                    StopApiError)
End Function

'***************************************************************************************************
'* 機能　　：このClassの要素を指す`script.SharedReference`(`{"sharedId": ...}`)を返します
'***************************************************************************************************
Private Function SharedReference() As Dictionary
    Set SharedReference = New Dictionary
    SharedReference.Add "sharedId", CurrentElement.sharedId
End Function

'***************************************************************************************************
'* 機能　　：`input.performActions`の1アクション(`{"type": ..., 任意の項目: 値}`)を返します
'---------------------------------------------------------------------------------------------------
'* 引数　　：actionType     "pointerDown" / "keyDown" 等
'            itemName       追加する項目名(`button` / `value`等)。省略時は`type`のみ
'            itemValue      追加する項目の値
'***************************************************************************************************
Private Function InputAction(actionType As String, Optional itemName As String, Optional itemValue As Variant) As Dictionary
    Set InputAction = New Dictionary
    InputAction.Add "type", actionType
    If LenB(itemName) Then InputAction.Add itemName, itemValue
End Function

'***************************************************************************************************
'* 機能　　：このClassの要素の中心へ、マウスを移動させるアクションを返します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`origin`に要素を指定すると、ブラウザ側で要素の表示中心が基準になるため、座標計算は不要です
'* 注意事項：要素が画面外の場合はエラーになるため、事前に`scrollIntoView`しておくこと
'***************************************************************************************************
Private Function PointerMoveToThis() As Dictionary
    Dim pointerOrigin As New Dictionary
    pointerOrigin.Add "type", "element"
    pointerOrigin.Add "element", SharedReference()

    Set PointerMoveToThis = InputAction("pointerMove")
    PointerMoveToThis.Add "x", 0
    PointerMoveToThis.Add "y", 0
    PointerMoveToThis.Add "origin", pointerOrigin
End Function

'***************************************************************************************************
'* 機能　　：1つの入力元(マウス or キーボード)として、アクション群を`input.performActions`で実行します
'---------------------------------------------------------------------------------------------------
'* 引数　　：sourceType     "pointer"：マウス / "key"：キーボード
'            actions        `InputAction`等で作ったアクションのCollection
'***************************************************************************************************
Private Sub PerformActions(sourceType As String, actions As Collection)
    '1. 入力元を組み立て
    Dim inputSource As New Dictionary
    inputSource.Add "type", sourceType
    inputSource.Add "id", "vba-" & sourceType
    If sourceType = "pointer" Then
        Dim pointerParameters As New Dictionary
        pointerParameters.Add "pointerType", "mouse"
        inputSource.Add "parameters", pointerParameters
    End If
    inputSource.Add "actions", actions

    '2. 実行
    Dim inputSources As New Collection
    inputSources.Add inputSource
    Dim params As New Dictionary
    params.Add "actions", inputSources
    ExecuteBiDi "input.performActions", params
End Sub

'***************************************************************************************************
'* 機能　　：`keyboardCode`(Windows仮想キーコード)を、WebDriverのキー値に変換します
'---------------------------------------------------------------------------------------------------
'* 参考URL ：https://w3c.github.io/webdriver/#keyboard-actions
'***************************************************************************************************
Private Function ConvertKeyValue(Key As keyboardCode) As String
    Select Case Key
        Case keyEnter:      ConvertKeyValue = ChrW(&HE007)
        Case keyDelete:     ConvertKeyValue = ChrW(&HE017)
        Case keyTab:        ConvertKeyValue = ChrW(&HE004)
        Case keyEsc:        ConvertKeyValue = ChrW(&HE00C)
        Case keyBackspace:  ConvertKeyValue = ChrW(&HE003)
        Case Else:          ConvertKeyValue = ChrW(Key)
    End Select
End Function



'***************************************************************************************************
'                             ■■■ JavaScript実行設定 ■■■
'***************************************************************************************************
'* 機能　　：`True`で、JavaScriptコード内で例外が起こった場合、Err.Raiseで止まります。基本は、開発時のみ使用する想定です
'***************************************************************************************************
Property Let SetOptionStopException(v As Boolean)
    RunOption.StopException = v
End Property

'***************************************************************************************************
'* 機能　　：Trueで、結果を待たないBiDiで実行します。
'---------------------------------------------------------------------------------------------------
'* 注意事項：必要な場面で使い終わったら必ず、`False`に戻すこと。面倒ごとになります
'***************************************************************************************************
Property Let SetOptionRunAsyncBiDi(v As Boolean)
    RunOption.RunAsyncBiDi = v
End Property

'***************************************************************************************************
'* 機能　　：True で「人間がクリックやキーボード操作をした結果」としてブラウザに偽装します。スクレピング対策が施されてる場合は、`True`にしておくといいかもしれません
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`script.callFunction`の`userActivation`として扱います
'***************************************************************************************************
Property Let SetOptionUserGesture(v As Boolean)
    RunOption.userGesture = v
End Property



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
Private Sub Class_Terminate()
    '`script.SharedId`はブラウザ側で保持されないため、`CDPElement`のような解放処理は不要です
End Sub
