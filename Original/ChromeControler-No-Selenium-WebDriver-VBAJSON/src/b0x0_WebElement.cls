VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "b0x0_WebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWebElement

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As LongPtr
        

Private json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos
Private CDP_ As a4_ExecuteCDP
Private events_ As a5_CDPEventHandler
Private Helper_ As a6_ExecuteHelperFunction

Private objectId_ As String
Private findInfo_ As Collection
Private nodeId_ As Long
Private attributes_ As Object
Private frameId_ As String

Private xpath_ As String
'

Public Sub Init(json As a2_JSONHandler, _
                       basicInfo As a3_BasicInfos, _
                       ExecCDP As a4_ExecuteCDP, _
                       events As a5_CDPEventHandler, _
                       execHelper As a6_ExecuteHelperFunction, _
                       objectId As String, _
                       findInfo As Collection, _
                       Optional attributes As Object = Nothing, _
                       Optional FrameId As String)
                       
    Set json_ = json
    Set basicInfo_ = basicInfo
    Set CDP_ = ExecCDP
    Set events_ = events
    Set Helper_ = execHelper
    
    objectId_ = objectId
    Set findInfo_ = findInfo
    Set attributes_ = attributes
    frameId_ = FrameId
    
End Sub


Public Property Get objectId() As String
    objectId = objectId_
End Property
Public Property Get FrameId() As String
    FrameId = frameId_
End Property


''CallJavaScriptFunction呼出しの共通関数
'戻り値が"Cannot find context with specified id"となる場合、ObjectIdの再取得が必要なため共通化
Private Function CallFunction(functionCode As String, args As Variant) As String
    Dim res As String
    res = CDP_.CallJavaScriptFunction(functionCode, objectId_, args)
    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CDP_.CallJavaScriptFunction(functionCode, objectId_, args)
    End If
    CallFunction = res
End Function
Private Function CallFunction_ArgMultiple(functionCode As String, ParamArray args() As Variant) As String
    Dim res As String
    res = CDP_.CallJavaScriptFunction_ArgMultiple(functionCode, objectId_, args)
    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CDP_.CallJavaScriptFunction_ArgMultiple(functionCode, objectId_, args)
    End If
    CallFunction_ArgMultiple = res
End Function
Public Function ReGetObjectId()
    Dim Codes As Variant
    Select Case findInfo_(1)
      Case "id"
        Codes = Helper_.GetIdCodes(findInfo_(2))
      Case "xpath"
        Codes = Helper_.GetXpathCodes(findInfo_(2))
      Case "name"
        Codes = Helper_.GetNameCodes(findInfo_(2))
      Case "tag"
        Codes = Helper_.GetTagCodes(findInfo_(2), findInfo_(3))
    End Select
    objectId_ = CDP_.GetObjectIdOfElementBy(findInfo_(2), basicInfo_, Codes)
End Function








'***********************************************
'Public
'チェックボックス要素についてChecked属性の設定
Private Property Let IWebElement_Checked(RHS As Boolean)
    Dim res As String
    'checkedプロパティを持っているか事前にチェック。持っていない場合はじく。
    res = CallFunction("function(){return this.checked;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim boolString As String
    If RHS Then
        boolString = "true"
    Else
        boolString = "false"
    End If

    res = CallFunction("function(value){ this.checked = value;}", boolString)
Exit Property

quit:
Helper_.DisplayErrorAndEnd _
    "<WebElement：Checkedプロパティ>" & vbCrLf & _
    "このElementはCheckedプロパティを持ちません。" & vbCrLf & _
    "※CheckedはCheckBoxで使用可能" & vbCrLf
End Property
'Public
'チェックボックス要素についてチェック状態取得
Private Property Get IWebElement_Checked() As Boolean
    Dim res As String
    res = CallFunction("function(){return this.checked;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Checked As Boolean
    Checked = json_.GetValue(res, "result", "result", "value")
    IWebElement_Checked = Checked
Exit Property

quit:
Helper_.DisplayErrorAndEnd "このElementはCheckedプロパティを持ちません。" & vbCrLf & _
                                            "※CheckedはCheckBoxで使用可能" & vbCrLf
End Property





















'Public
'要素の状態(有効or無効)を取得
Private Property Get IWebElement_IsEnabled() As Boolean
    Dim res As String
    res = CallFunction("function(){return this.disabled;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Enabled As Boolean
    Enabled = json_.GetValue(res, "result", "result", "value")
    IWebElement_IsEnabled = Not Enabled
Exit Property
    
quit:
Helper_.DisplayErrorAndEnd "このElementはEnabledプロパティを持ちません。" & vbCrLf
End Property

'エレメントをViewPort中心になるところまでスクロールし、エレメントの中心xy座標に左クリックイベント発行
Private Sub IWebElement_DispatchCDPMouseEvent(action As CDP_Mouse, Optional point As ClickPoint = ClickPoint.Center)
  
    ' 画面外にいるとクリックできないため、要素の位置を確認
    Dim contentBox As Object: Set contentBox = GetContentBox()
    Dim viewPortHeight As Long: viewPortHeight = GetViewPortHeight()
    '画面の中心に要素がくるように、Scroll
'    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, 0, contentBox("TopLeftY") - (viewPortHeight / 2)
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, 0, contentBox("TopLeftY") - (contentBox("Height") / 2)
    Sleep 200
    'そのうえで、画面内の位置として再度、位置を取得。
    Set contentBox = GetContentBox()
    
    Dim x As Long, y As Long
    Select Case point
        Case ClickPoint.Center
            x = contentBox("CenterX")
            y = contentBox("CenterY")
        Case ClickPoint.TopLeft
            x = contentBox("TopLeftX") + 2
            y = contentBox("TopLeftY") + 2
        Case ClickPoint.TopRight
            x = contentBox("TopRightX") - 2
            y = contentBox("TopRightY") + 2
        Case ClickPoint.BottomLeft
            x = contentBox("BottomLeftX") + 2
            y = contentBox("BottomLeftY") - 2
        Case ClickPoint.BottomRight
            x = contentBox("BottomRightX") - 2
            y = contentBox("BottomRightY") - 2
    End Select

    
    Select Case action
        Case CDP_Mouse.LeftClick_XY
            CDPClick "left", x, y, 1
        Case CDP_Mouse.LeftDoubleClick_XY
            CDPClick "left", x, y, 2
        Case CDP_Mouse.RightClick_XY
            CDPClick "right", x, y, 1
        Case CDP_Mouse.ScrollToThisElement
            'この要素位置まで移動するだけで何もしない
    End Select
End Sub


'Private Function IsRequireScrollingToThisElement(action As CDP_Mouse, Optional x_pixel As Long = 0, Optional y_pixel As Long = 0) As Boolean
'    If action = CDP_Mouse.Back Or action = CDP_Mouse.Forward _
'    Or action = CDP_Mouse.Moved Or action = CDP_Mouse.MouseWheel Then
'        IsRequireScrollingToThisElement = False
'    Else
'        If x_pixel = 0 And y_pixel = 0 Then
'            IsRequireScrollingToThisElement = True
'        Else
'            IsRequireScrollingToThisElement = False
'        End If
'    End If
'End Function

'CDPMethodを呼出す場合に、引数にObjectIdを使用して呼出すMethodの共通管理のための関数
'戻り値が"Cannot find context with specified id"となる場合に、ObjectIdの再取得が必要になるため共通化
'現在FocusとGetGetBoxModelの呼び出しのみ
Private Function CallObjectIdUsingMethod(methodName As String) As String
    Dim res As String
    res = CallByName(CDP_, methodName, VbMethod, objectId_)

    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CallByName(CDP_, methodName, VbMethod, objectId_)
    End If
    CallObjectIdUsingMethod = res
End Function

'Boxmodelのcontent領域の4つの頂点のx,y,width,height,center位置のxyを辞書として取得
Private Function GetContentBox() As Variant
    Dim res As String: res = CallObjectIdUsingMethod("GetBoxModel")
    Dim quadXY As Variant: quadXY = json_.GetValue(res, "result", "model", "content")
    quadXY = Split(quadXY, ",")
    Dim posName As Variant: posName = "TopLeftX,TopLeftY,TopRightX,TopRightY,BottomRightX,BottomRightY,BottomLeftX,BottomLeftY"
    posName = Split(posName, ",")
    
    Dim contentBox As Object: Set contentBox = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(posName) To UBound(posName)
        contentBox.Add posName(i), quadXY(i)
    Next
    
    Dim width As Long: width = json_.GetValue(res, "result", "model", "width")
    Dim height As Long: height = json_.GetValue(res, "result", "model", "height")
    contentBox.Add "Width", width
    contentBox.Add "Height", height
    contentBox.Add "CenterX", quadXY(0) + (width / 2)
    contentBox.Add "CenterY", quadXY(1) + (height / 2)
    Set GetContentBox = contentBox
End Function
Private Function GetViewPortHeight() As Long
    Dim res As String
    res = CallFunction("function(){return window.innerHeight;}", "")
    res = json_.GetValue(res, "result", "result", "value")
    GetViewPortHeight = Int(res)
End Function

Private Sub CDPClick(button As String, x As Long, y As Long, clickNum As Long)
    Dim result As String
    Dim i As Long
    For i = 1 To clickNum
        result = CDP_.DispatchMouseEvent("mousePressed", x, y, button, 1)
        result = CDP_.DispatchMouseEvent("mouseReleased", x, y, button, 1)
    Next
End Sub

'Public
'CDPによるキーボードイベント発行
Private Sub IWebElement_DispatchCDPKeyEvent(action As CDP_Keyboard, KeyCode As Key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False)
    Dim actionStr As String
    Select Case action
        Case CDP_Keyboard.KeyDown_
            actionStr = "keyDown"
        Case CDP_Keyboard.KeyUp_
            actionStr = "keyUp"
        Case CDP_Keyboard.RawKeyDown
            actionStr = "rawKeyDown"
        Case CDP_Keyboard.Char
            actionStr = "char"
    End Select
    
    Dim modifiers As Long
    If alt Then modifiers = 1
    If ctrl Then modifiers = modifiers + 2
    If shift Then modifiers = modifiers + 8
    
    Dim result As String
    result = CDP_.DispatchKeyEvent(actionStr, modifiers, KeyCode)
End Sub


'Public
'javaScriptによってクリック実行
Private Sub IWebElement_Click()
    CallObjectIdUsingMethod "Focus"

    Dim res As String
    res = CallFunction("function(){this.click();}", "")
    
    If InStr(res, "this.click is not a function") Then
        res = CallFunction("function(){this.dispatchEvent(new MouseEvent( 'click',{bubbles: true} ));}", "")
'        res = CallFunction("function(value){" & _
                                    "this.dispatchEvent(new CustomEvent('click', {bubbles: true}));" & _
                                    "}", "")
    End If
    
    If InStr(res, "this.dispatchEvent is not a function") Then
        FocusAndEnterKey
    End If
    
    If IsPageChanging Then WaitPageChangeToComplete
    If IsIsolatedFrameChanging Then WaitIsolatedFrameChangeToComplete
End Sub
Private Sub FocusAndEnterKey()
    CallObjectIdUsingMethod "Focus"
    CDP_.DispatchKeyEvent "keyDown", 0, 13
    CDP_.DispatchKeyEvent "keyUp", 0, 13
End Sub


Private Function IsPageChanging() As Boolean
    CDP_.PageEnable '一回無害なメソッドを投げて、イベント発生状況更新
    
    If events_.IsCurrentTargetScheduledNavigation _
    Or events_.IsCurrentTargetStartedLoading Then
        IsPageChanging = True
    Else
        IsPageChanging = False
    End If
End Function
Private Sub WaitPageChangeToComplete()
    
    'TargetNavigatedは返って来ないケースがある？ので繰り返し回数に制限をかけてループ
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = events_.IsCurrentTargetNavigated
'                                Or　chrome_.CDPEvents.IsCurrentTargetClearedScheduledNavigation
        If isEndChange Then Exit Do
        
        CDP_.RuntimeEnable
        Sleep 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 4)
    
    events_.IsCurrentTargetScheduledNavigation = False
    events_.IsCurrentTargetStartedLoading = False
    events_.IsCurrentTargetClearedScheduledNavigation = False
    events_.IsCurrentTargetNavigated = False
End Sub
Private Function IsIsolatedFrameChanging() As Boolean
    CDP_.DOMEnable '一回無害なメソッド(frameなのでDOM系)を投げて、イベント発生状況更新
    
    If events_.IsIsolatedFrameScheduledNavigation _
    Or events_.IsIsolatedFrameStartedLoading Then
        IsIsolatedFrameChanging = True
    Else
        IsIsolatedFrameChanging = False
    End If
End Function
Private Sub WaitIsolatedFrameChangeToComplete()
    
    'FrameNavigatedは返って来ないケースがある？ので繰り返し回数に制限をかけてループ
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = events_.IsIsolatedFrameNavigated
'                               Or chrome_.CDPEvents.IsIsolatedFrameClearedScheduledNavigation
        If isEndChange Then Exit Do
        
        CDP_.DOMEnable
        CDP_.RuntimeEnable
        Sleep 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 4)
    
    events_.IsIsolatedFrameScheduledNavigation = False
    events_.IsIsolatedFrameStartedLoading = False
    events_.IsIsolatedFrameClearedScheduledNavigation = False
    events_.IsIsolatedFrameNavigated = False
End Sub




'Public
'要素のValue属性に文字列を設定
'非推奨（本来メソッド名をSetValueと命名すべきであった。→機能は全く同じでSetValueメソッドを新たに用意。過去バージョンとの互換性のために残されているメソッド）
Private Sub IWebElement_SetText(text As String)
    FocusAndEnterKey
    Dim encodingText As String: encodingText = json_.EncodeURL(text)
    CallFunction "function(value){value = decodeURI(value);this.value = value;}", encodingText
    
    CallFunction Helper_.GetDispatchInputEventCode, encodingText
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

'Public
'要素のValue属性に文字列を設定
Private Sub IWebElement_SetValue(value As String)
    FocusAndEnterKey
    Dim encodingText As String: encodingText = json_.EncodeURL(value)
    CallFunction "function(value){value = decodeURI(value);this.value = value;}", encodingText
    
    CallFunction Helper_.GetDispatchInputEventCode, encodingText
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

'Public
'要素のTextContentに文字列を設定
Private Sub IWebElement_SetTextContent(text As String)
    CallObjectIdUsingMethod "Focus"
    Dim encodingText As String: encodingText = json_.EncodeURL(text)
    CallFunction "function(value){value = decodeURI(value);this.textContent  = value;}", encodingText
    
    CallFunction Helper_.GetDispatchInputEventCode, encodingText
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

'Public
'要素のValue属性にPropertySetter経由で文字列を設定。
Private Function IWebElement_SetValueViaPropertySetter(value As String) As String
    FocusAndEnterKey
    Dim encodingText As String: encodingText = json_.EncodeURL(value)
    CallFunction "function(value){value = decodeURI(value);Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(this, value);}", encodingText
                                                                                    
    CallFunction Helper_.GetDispatchInputEventCode, ""
    CallFunction Helper_.GetDispatchChangeEventCode, ""
End Function

'Public
'SelectBoxの選択肢を文字列で選択
Private Sub IWebElement_SelectItemInSelectBoxByText(text As String)
    Dim encodingText As String: encodingText = json_.EncodeURL(text)
    
    Dim res As String: res = CallFunction(Helper_.GetSelectItemCode, encodingText)
    
    Dim IsSelect As Boolean: IsSelect = json_.GetValue(res, "result", "result", "value")
    If IsSelect = False Then
        MsgBox "<SelectItemInSelectBoxByText>" & vbCrLf & _
                    "「" & text & "」は選択肢の中に存在しません。" & vbCrLf & "処理を終了します。"
        End
    End If
    
'    CallFunction "function(value){this.onchange();}", encodingText
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

'Public
Private Function IWebElement_FindElementByName(name As String) As IWebElement
    'name指定に日本語が含まれるとエラーになるため要エンコード
    name = json_.EncodeURL(name)
    
    Dim Codes As Variant: Codes = Helper_.GetNameCodes(name)
    Dim funcCode As String: funcCode = Codes(1)
    Dim res As String: res = CallFunction(funcCode, name)
    
    Dim objectId As String
    objectId = json_.GetValue(res, "result", "result", "objectId")
    If objectId = "" Then GoTo quit
    
    Dim findInfoInElem As Collection
    Set findInfoInElem = Helper_.MakeFindInfo("name", name)
    
    Set IWebElement_FindElementByName = _
        CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & name
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Name", name)
End Function

'Public
Private Function IWebElement_FindElementsByTag(tagName As String) As IWebElements
    Dim elemsCount As Long
    elemsCount = GetElementsCount(tagName)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    Dim index As Long
    For index = 0 To elemsCount - 1
        elements.Add CreateElementsItemByIndex(tagName, index)
    Next
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebElement_FindElementsByTag = elems
Exit Function
quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & tagName
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Tag", tagName, "○○")
End Function
Private Function GetElementsCount(tagName As String) As Long
    Dim funcCode As String, res As String
    funcCode = "function(tagName){return  this.getElementsByTagName('" & tagName & "').length;}"
    res = CallFunction(funcCode, tagName)
    GetElementsCount = json_.GetValue(res, "result", "result", "value")
End Function
Private Function CreateElementsItemByIndex(tagName As String, index As Long) As IWebElement
    Dim funcCode As String, res As String
    funcCode = "function(tagName){return  this.getElementsByTagName(tagName)[" & index & "];}"
    res = CallFunction(funcCode, tagName)
    
    Dim objectId As String
    objectId = json_.GetValue(res, "result", "result", "objectId")
    Dim findInfoInElem As Collection: Set findInfoInElem = Helper_.MakeFindInfo("tag", tagName, index)

    Set CreateElementsItemByIndex = _
        CDP_.MakeElementObject(objectId, findInfoInElem, basicInfo_, events_, Helper_)
End Function

'Public
'エレメントのTextContentを取得
Private Function IWebElement_GetTextContent() As String
    Dim funcCode As String, res As String
    funcCode = "function(){return this.textContent.replace(/\\s/g, ' ');}"
    res = CallFunction(funcCode, "")
    
    IWebElement_GetTextContent = json_.GetValue(res, "result", "result", "value")
End Function

'Public
'セレクトボックスの選択中オプションのValueを取得
Private Function IWebElement_GetSelectedValue() As String
    Dim funcCode As String, res As String
    funcCode = "function(){return this.value.replace(/\\s/g, ' ');}"
    res = CallFunction(funcCode, "")

    IWebElement_GetSelectedValue = json_.GetValue(res, "result", "result", "value")
End Function

'Public
'セレクトボックスの選択中オプションのTextContentを取得
Private Function IWebElement_GetSelectedTextContent() As String
    Dim funcCode As String, res As String
    funcCode = Helper_.GetSelectedTextContentCode
    res = CallFunction(funcCode, "")

    IWebElement_GetSelectedTextContent = json_.GetValue(res, "result", "result", "value")
End Function

'Public
'type=fileのInputBoxへファイル設定
Private Sub IWebElement_SetFileToFileSelectInputButton(filePath As String)
'    filePath = Replace(filePath, "\", "\\")
'    filePath = Json_.EncodeURL(filePath)
'    filePath = Replace(filePath, "%5C", "\")
'    chrome_.SetFileInputFiles objectId_,
'CDPMethodで設定しようとしても、ファイルパスに日本語が含まれる場合、
'エンコードして送信はできるが、クローム側でデコードしてくれない為、没
    Dim res As String: res = CallObjectIdUsingMethod("GetBoxModel")
    Dim contentBox As Object: Set contentBox = GetContentBox(res)
    
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, 0, CLng(contentBox("CenterY")) - 50
    
    IWebElement_DispatchCDPMouseEvent LeftClick
    Sleep 1000
    ControlFileSelectDialog filePath
    Sleep 1000
End Sub
'ファイル選択ダイアログ操作
Private Sub ControlFileSelectDialog(filePath As String)
    Dim DialogHwnd As LongPtr: DialogHwnd = FindWindow(vbNullString, "開く")
    Dim InputHwnd As Long
    InputHwnd = FindWindowEx(DialogHwnd, 0&, "ComboBoxEx32", "")
    InputHwnd = FindWindowEx(InputHwnd, 0&, "ComboBox", "")
    InputHwnd = FindWindowEx(InputHwnd, 0&, "Edit", "")
    Call SendMessage(InputHwnd, &HC, 0, ByVal filePath)
 
    Dim ButtonHwnd As Long
    ButtonHwnd = FindWindowEx(DialogHwnd, 0&, "Button", "開く(&O)")
    Call SendMessage(ButtonHwnd, &HF5, 0, 0&)
End Sub

'Public
'クリック後に制御を返さないアラートダイアログ出現時に使用するアラート処理メソッド
Private Sub IWebElement_ClickAndThenAlertDialogErase()
    CallObjectIdUsingMethod "Focus"
    Dim res As String
    res = CallFunction("function(){this.click();}", "")
    Sleep 1000
    
    CDP_.HandleAlertDialog
End Sub

'Public
'elementの属性値を取得（属性が存在しない場合は””を返却）
Private Function IWebElement_GetAttribute(attributeName As String) As String
    Dim res As String
    res = CallFunction("function(name){return this.getAttribute(name);}", attributeName)
    If InStr(res, """value"":null") > 0 Then
        IWebElement_GetAttribute = ""
    Else
        IWebElement_GetAttribute = json_.GetValue(res, "result", "result", "value")
    End If
End Function

'Public
'Elementのclass属性に値を設定
Private Sub IWebElement_SetClassName(className As String)
    Dim funcCode As String, res As String
    funcCode = "function(className){this.className = '" & className & "';}"
    res = CallFunction(funcCode, className)
    res = json_.GetValue(res, "result", "result", "value")
End Sub


Private Function IWebElement_SetAttribute(attributeName As String, value As Variant) As String
    Dim funcCode As String, res As String
    ' 引数は一つしか渡せないため、↓このValueはオブジェクトとして渡し、コード内部ではオブジェクトから値を取り出すことになる
    funcCode = "function(value){let a = value.arg1;let b = value.arg2;console.log(a,b); this.setAttribute( a , b);}"
    ' 呼出し時にはオブジェクトに設定するプロパティ名,値・・・とセットで渡す
    res = CallFunction_ArgMultiple(funcCode, "arg1", attributeName, "arg2", value)
    IWebElement_SetAttribute = json_.GetValue(res, "result", "result")
    
    ' res = CallFunction(funcCode, "{""arg1"", " & attributeName & ", ""arg2""," & value & "}")
    ' CallFunctionを呼出すときに第二引数に↑のようにJSオブジェクト文字列を渡す形式にすれば、CallFunction_ArgMultipleメソッドは不要。
    ' 作った後に気づいた。
End Function


'Public
Private Sub IWebElement_DispatchJSChangeEvent()
    CallFunction Helper_.GetDispatchChangeEventCode, ""
End Sub

'Public
Private Sub IWebElement_DispatchJSInputEvent()
    CallFunction Helper_.GetDispatchInputEventCode, ""
End Sub

'Public
Private Sub IWebElement_DispatchJSFocusEvent()
    CallFunction Helper_.GetDispatchFocusEventCode, ""
End Sub

'Public
Private Sub IWebElement_DispatchJSMouseEvent(action As JS_Mouse)
    Dim jsCode As String
    jsCode = Helper_.GetDispatchMouseEventCode(action)
    CallFunction jsCode, ""
End Sub

' 「keyCode:」　押されたキーのキーコード番号 （Enter: 13 、Down：40 など）
' 「altKey:」 Alt の同時押しは true, 「shiftKey:」 Shift の同時押しは true,「ctrlKey:」 Ctrl の同時押しは true
Private Sub IWebElement_DispatchJSKeyboardEvent(action As JS_Keyboard, KeyCode As Key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False)
    Dim jsCode As String
    jsCode = Helper_.GetDispatchKeyboardEventCode(action, KeyCode, alt, shift, ctrl)
    CallFunction jsCode, ""
End Sub


Private Function IWebElement_GetProperty(PropertyName As String) As Variant
    Dim funcCode As String, res As String
    funcCode = "function(value){return this['" & PropertyName & "'];}"
    res = CallFunction(funcCode, "")
    Dim returnValue As Variant
    returnValue = json_.GetValue(res, "result", "result", "value")
    If IsNumeric(returnValue) Then
        IWebElement_GetProperty = CLng(returnValue)
    ElseIf IsDate(returnValue) Then
         IWebElement_GetProperty = CDate(returnValue)
    ElseIf returnValue = "true" Then
        IWebElement_GetProperty = True
    ElseIf returnValue = "false" Then
        IWebElement_GetProperty = False
    Else
        IWebElement_GetProperty = returnValue
    End If
End Function

Private Function IWebElement_SetProperty(PropertyName As String, value As Variant) As String
    Dim funcCode As String, res As String
    funcCode = "function(value){return this['" & PropertyName & "'] = value ;}"
    res = CallFunction(funcCode, value)
    IWebElement_SetProperty = json_.GetValue(res, "result", "result")
End Function

