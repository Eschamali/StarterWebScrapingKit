VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'

Public Enum TargetIs
    html_outerhtml = 1
    html_innerhtml
    body_outerhtml
    body_innerhtml
    body_innertext
End Enum

Public Enum Image
    jpeg = 1
    png
    webp
End Enum

Public Enum WindowSize
    normal = 1
    minimize
    maximize
End Enum

Public Property Get URL() As String

End Property

Public Property Get QueryString() As Object
    
End Property

Public Function GetValueFromQueryString(KeyName As String) As String
    
End Function

Public Sub OpenURL(targetURL As String)

End Sub

Public Function ExecuteScript(script As String) As String

End Function

Public Function FindElementById(Id As String) As IWebElement

End Function

Public Function FindElementByXPath(xPath As String) As IWebElement

End Function

Public Function FindElementByName(name As String) As IWebElement

End Function

Public Function FindElementsByTag(tagName As String) As IWebElements

End Function
Public Function FindElementsByXPath(xPath As String) As IWebElements

End Function

Public Function FindElementsByClassName(className As String) As IWebElements

End Function

Public Function IsExistsElementByXpath(xPath As String) As Boolean

End Function

Public Function IsExistsElementById(Id As String) As Boolean

End Function

Public Function IsExistsElementByName(name As String) As Boolean

End Function

Public Sub WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)

End Sub

Public Sub WaitForElementToAppearById(Id As String, maxWaitTime_sec As Long)

End Sub

Public Sub WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)

End Sub

Public Sub HandleAlertDialog()

End Sub

Public Sub SleepByWinAPI(milliSecond As Long)

End Sub

Public Sub SwitchTabByIndex(index As Long)
    
End Sub

Public Sub SwitchFrameByNameOrURLOrIndex(value As Variant)

End Sub

Public Sub SwitchFrameByIframeElement(IframeElement As IWebElement)

End Sub

Public Sub SwitchFrameToDefault()

End Sub

Public Sub DebugPrintAccessibleFrameInfo()

End Sub

Public Sub CloseWindow()
    
End Sub

Public Sub ScrollWindow(yPixel As Long, Optional xPixel As Long = 0)

End Sub

Public Sub SetForeGroundWindow()

End Sub

'Public Sub AllowPopUpWindowForURL(targetURL As String)
'
'End Sub

Public Function PageSource(target As TargetIs) As String

End Function

Public Sub Back()

End Sub

Public Sub Forward()

End Sub

'Public Sub PointerMove(viewportX As Long, viewportY As Long)
'
'End Sub



Public Function ScreenShotSaveAsFile( _
    Optional folderPath As String = "", _
    Optional fileName As String = "", _
    Optional imageFormat As Image = Image.png) As String

End Function

Public Function ScreenShotPasteToSheet( _
    left As Single, top As Single, Optional width As Single = 0, Optional height As Single = 0, _
    Optional sheetName As String = "", _
    Optional imageFormat As Image = Image.png) As Shape

End Function

Public Sub ShowWindow(size As WindowSize)

End Sub

Public Function GetSecureRandomNumber(count As Long) As Variant
    
End Function

Public Function GetURLSafeRundomString(charNum As Long) As String

End Function

Public Function GetBase64URLEncodedSHA256(str As String) As String

End Function

Public Function EncodeURIConpornent(str As String) As String
    
End Function

Public Function ExecuteCDP(CDPMethodJson As String) As String

End Function
