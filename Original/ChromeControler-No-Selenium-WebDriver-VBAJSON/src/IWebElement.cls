VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'JSMouseEvnet用
Public Enum JS_Mouse
: LeftDown = 1: LeftPress: Leftup
: Enter: Leave: Move
: Over: Out
: LeftClick: LeftDoublelClick
': Drag: DragStart: DragEnd
': DragEnter: DragExit: DragLeave: DragOver
': Drop
End Enum
'auxclick(マウス中央ボタン)、Scroll(ページや対象がスクロール時)、 wheel（マウスホイール回転時）

'JSKeyboardイベント用
Public Enum JS_Keyboard
    KeyDown = 1
    KeyPress
    KeyUp
End Enum

Public Enum Key
: Tab_Key = 9: Enter_Key = 13: shift_Key = 16: Control_Key = 17: Alt_Key = 18
: Escape_Key = 27: Space_Key = 32
: PageUp_Key = 33: PageDown_Key = 34
: ArrowLeft_Key = 37: ArrowUp_Key = 38: ArrowRight_Key = 39: ArrowDown_Key = 40
: PrintScreen_Key = 44: Delete_Key = 46
: Digit0 = 48: Digit1: Digit2: Digit3: Digit4:
: Digit5: Digit6: Digit7: Digit8: Digit9
: KeyA = 65: KeyB: KeyC: KeyD: KeyE: KeyF: KeyG
: KeyH: KeyI: KeyJ: KeyK: KeyL: KeyM: KeyN: KeyO
: KeyP: KeyQ: KeyR: KeyS: KeyT: KeyU: KeyV: KeyW: KeyX: KeyY: KeyZ
End Enum
' 参考他イベント
' ,Select(input,textareaのテキスト選択時),submit(フォーム送信時),toggle(<details>要素のアコーディオンカーテン）
'

Public Enum CDP_Mouse
     LeftClick_XY = 1
     LeftDoubleClick_XY
     RightClick_XY
     ScrollToThisElement
End Enum

Public Enum CDP_Keyboard
    KeyDown_ = 1
    KeyUp_
    RawKeyDown
    Char
End Enum

Public Enum ClickPoint
    Center = 1
    TopLeft
    TopRight
    BottomLeft
    BottomRight
End Enum

'clickはpressとreleaseの合わせ技
'middle
'     Moved
'     MouseWheel
'     Back
'     Forward

Public Property Get Checked() As Boolean

End Property
Public Property Let Checked(bool As Boolean)

End Property

Public Property Get IsEnabled() As Boolean

End Property

'非推奨（本来メソッド名をSetValueと命名すべきであったため、機能は全く同じでSetValueメソッドを用意。過去バージョンとの互換性のために残しているだけ）
Public Sub SetText(text As String)

End Sub

Public Sub SetValue(value As String)

End Sub

Public Sub SetTextContent(text As String)

End Sub

Public Function SetValueViaPropertySetter(value As String) As String

End Function

Public Sub Click()

End Sub

Public Sub SelectItemInSelectBoxByText(text As String)

End Sub

Public Function FindElementByName(name As String) As IWebElement

End Function

Public Function FindElementsByTag(tagName As String) As IWebElements

End Function

Public Function GetTextContent() As String

End Function

Public Function GetSelectedValue() As String

End Function

Public Function GetSelectedTextContent() As String

End Function

Public Sub SetFileToFileSelectInputButton(filePath As String)

End Sub

Public Sub ClickAndThenAlertDialogErase()

End Sub

Public Function GetAttribute(attributeName As String) As String

End Function

Public Function SetAttribute(attributeName As String, value As Variant) As String
'
End Function

Public Sub SetClassName(className As String)
    
End Sub

Public Sub DispatchCDPMouseEvent(action As CDP_Mouse, Optional point As ClickPoint = ClickPoint.Center)

End Sub
Public Sub DispatchCDPKeyEvent(action As CDP_Keyboard, KeyCode As Key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False)

End Sub

Public Sub DispatchJSInputEvent()

End Sub
Public Sub DispatchJSChangeEvent()

End Sub
Public Sub DispatchJSFocusEvent()

End Sub
Public Sub DispatchJSMouseEvent(action As JS_Mouse)

End Sub
Public Sub DispatchJSKeyboardEvent(action As JS_Keyboard, KeyCode As Key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False)

End Sub

Public Function GetProperty(PropertyName As String) As Variant

End Function

Public Function SetProperty(PropertyName As String, value As Variant) As String
'
End Function
