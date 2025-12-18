VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a6_ExecuteHelperFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private Const MOUSE_EVENT_NAME_NUM = 17
Private Const MOUSE_EVENT_NAMES = _
"mousedown,mousepress,mouseup,mouseenter,mouseleave, mousemove,mouseover,mouseout," & _
"click,dblclick," & _
"drag,dragstart,dragend,dragenter,dragexit,dragleave,dragover,drop"


Private Const KEYBOARD_EVENT_KEY_NUM = 51
'5,4,4,2,10,7,7,12
Private Const KEY_CODES = _
"9,13,16,17,18," & _
"27,32,33,34," & _
"37,38,39,40," & _
"44,46," & _
"48,49,50,51,52,53,54,55,56,57," & _
"65,66,67,68,69,70,71," & _
"72,73,74,75,76,77,78," & _
"79,80,81,82,83,84,85,86,87,88,89,90"

Private Const KEY_NAMES = _
"Tab,Enter,Shift,Control,Alt," & _
"Escape,Space,PageUp,PageDown," & _
"ArrowLeft,ArrowUp,ArrowRight,ArrowDown," & _
"PrintScreen,Delete," & _
"Digit0,Digit1,Digit2,Digit3,Digit4,Digit5,Digit6,Digit7,Digit8,Digit9," & _
"KeyA,KeyB,KeyC,KeyD,KeyE,KeyF,KeyG," & _
"KeyH,KeyI,KeyJ,KeyK,KeyL,KeyM,KeyN," & _
"KeyO,KeyP,KeyQ,KeyR,KeyS,KeyT,KeyU,KeyV,KeyW,KeyX,KeyY,KeyZ"
 

Private mouseEventNames_ As Object
Private keyboardEventNames_ As Object
Private keyNames_ As Object
Private imageFormatNames_ As Object
Private ngCharsAsFileName_ As Object

Private Sub Class_Initialize()
    Set mouseEventNames_ = CreateObject("Scripting.Dictionary")
    Dim mouseEventNamesString As Variant: mouseEventNamesString = Split(MOUSE_EVENT_NAMES, ",")
    Dim num As Long
    For num = 1 To MOUSE_EVENT_NAME_NUM
        mouseEventNames_.Add num, mouseEventNamesString(num - 1)
    Next
    
    Set keyboardEventNames_ = CreateObject("Scripting.Dictionary")
    keyboardEventNames_.Add 1, "keydown"
    keyboardEventNames_.Add 2, "keypress"
    keyboardEventNames_.Add 3, "keyup"
    
    Set keyNames_ = CreateObject("Scripting.Dictionary")
    Dim keyboardEventCodes As Variant: keyboardEventCodes = Split(KEY_CODES, ",")
    Dim KeyNames As Variant: KeyNames = Split(KEY_NAMES, ",")
    Dim i As Long, KeyCode As Long
    Dim KeyName As String
    For i = 0 To KEYBOARD_EVENT_KEY_NUM - 1
        KeyCode = keyboardEventCodes(i)
        KeyName = KeyNames(i)
        keyNames_.Add KeyCode, KeyName
    Next
    
    Set imageFormatNames_ = CreateObject("Scripting.Dictionary")
    imageFormatNames_.Add 1, "jpeg"
    imageFormatNames_.Add 2, "png"
    imageFormatNames_.Add 3, "webp"
    
    Set ngCharsAsFileName_ = CreateObject("Scripting.Dictionary")
    ngCharsAsFileName_.Add "\", ""
    ngCharsAsFileName_.Add "/", ""
    ngCharsAsFileName_.Add ":", ""
    keyboardEventNames_.Add "*", ""
    keyboardEventNames_.Add "?", ""
    keyboardEventNames_.Add """", ""
    keyboardEventNames_.Add ">", ""
    keyboardEventNames_.Add "<", ""
    keyboardEventNames_.Add "|", ""
End Sub

Public Sub SleepByWinAPI(ms As Long)
    Sleep ms
End Sub

''FindByXPathのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetXpathCodes(xPath As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.evaluate(decodeURIComponent('" & xPath & "'), document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    Codes(1) = "function(xpath){return  document.evaluate(decodeURIComponent(xpath), document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
    GetXpathCodes = Codes
End Function

'FindByIdのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetIdCodes(Id As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementById(decodeURIComponent('" & Id & "'));"
    Codes(1) = "function(id){return  document.getElementById(decodeURIComponent(id));}"
    GetIdCodes = Codes
End Function

'FindByNameのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetNameCodes(name As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementsByName(decodeURIComponent('" & name & "'))[0];"
    Codes(1) = "function(name){return  document.getElementsByName(decodeURIComponent(name))[0];}"
    GetNameCodes = Codes
End Function

'FindByTagのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetTagCodes(tagName As String, index As Long) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementsByTagName('" & tagName & "')[" & index & "];"
    Codes(1) = "function(tagName){return  document.getElementsByTagName(tagName)[" & index & "];}"
    GetTagCodes = Codes
End Function

'FindByClassNameのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetClassNameCodes(className As String, index As Long) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementsByClassName('" & className & "')[" & index & "];"
    Codes(1) = "function(className){return  document.getElementsByClassName(className)[" & index & "];}"
    GetClassNameCodes = Codes
End Function

'FindElementsByTagの要素数を取得するJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetTagCountCodes(tagName As String) As Variant
    Dim countCodes(1) As String
    countCodes(0) = "document.getElementsByTagName('" & tagName & "').length;"
    countCodes(1) = "function(tagName){return  document.getElementsByTagName('" & tagName & "').length;}"
    GetTagCountCodes = countCodes
End Function

'FindElementsByClassの要素数を取得するJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetClassNameCountCodes(className As String) As Variant
    Dim countCodes(1) As String
    countCodes(0) = "document.getElementsByClassName('" & className & "').length;"
    countCodes(1) = "function(className){return  document.getElementsByClassName('" & className & "').length;}"
    GetClassNameCountCodes = countCodes
End Function

'FindElementsByXPathの要素数を取得するJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetXPathCountCodes(xPath As String)
    Dim Codes(1) As String
    Codes(0) = "result = document.evaluate(decodeURIComponent('" & xPath & "'), document, null, " & _
                        "XPathResult.ORDERED_NODE_ITERATOR_TYPE, null );" & _
                     "ary = [];" & _
                     "thisNode = result.iterateNext();" & _
                     "while (thisNode) {" & _
                         "ary.push(thisNode);" & _
                         "thisNode = result.iterateNext();" & _
                     "}" & _
                     "ary.length;"
                     
    Codes(1) = "function(xpath){" & _
                            "result = document.evaluate(decodeURIComponent('" & xPath & "'), document, null, " & _
                                "XPathResult.ORDERED_NODE_ITERATOR_TYPE, null );" & _
                            "ary = [];" & _
                            "thisNode = result.iterateNext();" & _
                            "while (thisNode) {" & _
                                "ary.push(thisNode);" & _
                                "thisNode = result.iterateNext();" & _
                            "}" & _
                            "return  ary.length;" & _
                        "}"
                            
    GetXPathCountCodes = Codes
End Function

'FindElementsByXPathの要素をIndexで判断し取得するJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetXPathResultCodes(xPath As String, index As Long)
    Dim Codes(1) As String
    Codes(0) = "result = document.evaluate(decodeURIComponent('" & xPath & "'), document, null, " & _
                        "XPathResult.ORDERED_NODE_ITERATOR_TYPE, null );" & _
                     "count = 0;" & _
                     "thisNode = result.iterateNext();" & _
                     "while (thisNode) {" & _
                         "if(count === " & index & "){" & _
                             "break;" & _
                         "}" & _
                         "thisNode = result.iterateNext();" & _
                         "count ++;" & _
                     "}" & _
                     "thisNode;"
                     
    Codes(1) = "function(xpath){" & _
                            "result = document.evaluate(decodeURIComponent('" & xPath & "'), document, null, " & _
                                "XPathResult.ORDERED_NODE_ITERATOR_TYPE, null );" & _
                            "count = 0;" & _
                            "thisNode = result.iterateNext();" & _
                            "while (thisNode) {" & _
                                "if(count === " & index & "){" & _
                                    "break;" & _
                                "}" & _
                                "thisNode = result.iterateNext();" & _
                                "count ++;" & _
                            "}" & _
                            "return  thisNode;" & _
                        "}"
                            
    GetXPathResultCodes = Codes
End Function


'共通エラー対応。メッセージ表示、かつインターネットハンドルを解放して強制終了とする｡
'※エラー発生時にデバッグモードに入り、ユーザーに任せると解放されない可能性が高いので強制終了とする。
Public Sub DisplayErrorAndEnd(displayText As String)
    displayText = displayText & vbCrLf & "処理を終了します。"
    MsgBox displayText, vbCritical, "エラー"
    End
End Sub

Public Function GetSelectItemCode() As String
    GetSelectItemCode = _
    "function(value){value = decodeURI(value);" & _
        "for(let i = 0; i < this.options.length; i++){" & _
            "let val1 = this.options[i].value.replace(/\\r?\\n/g,'');" & _
            "let val2 = this.options[i].textContent.replace(/\\r?\\n/g,'');" & _
            "console.log('(you)' + value + ':(value)' + val1 + ':(textContent)'+ val2);" & _
            "if(val1 === value || val2 === value){" & _
                "this.options[i].selected = true;" & _
                "return true;" & _
            "}" & _
        "}" & _
    "return false;}"
End Function

Public Function GetSelectedTextContentCode() As String
    GetSelectedTextContentCode = _
    "function(){" & _
        "for(let i = 0; i < this.options.length; i++){" & _
            "if(this.options[i].selected){" & _
                "return this.options[i].textContent.replace(/\\s/g, ' ');" & _
            "}" & _
        "}" & _
    "}"
End Function

Public Function GetDispatchChangeEventCode() As String
    GetDispatchChangeEventCode = _
    "function(){" & _
        "this.dispatchEvent(new CustomEvent('change', {bubbles: true}));" & _
    "}"
End Function

Public Function GetDispatchInputEventCode() As String
    GetDispatchInputEventCode = _
    "function(){" & _
        "this.dispatchEvent(new CustomEvent('input', {bubbles: true}));" & _
    "}"
End Function

Public Function GetDispatchFocusEventCode() As String
    GetDispatchFocusEventCode = _
    "function(){" & _
        "this.dispatchEvent(new CustomEvent('focus', {bubbles: true}));" & _
    "}"
End Function

Public Function GetDispatchMouseEventCode(action As JS_Mouse) As String
    Dim eventName As String: eventName = mouseEventNames_(action)
    
    GetDispatchMouseEventCode = _
    "function(){" & _
        "this.dispatchEvent(new MouseEvent('" & eventName & "', {bubbles: true}));" & _
    "}"
End Function

Public Function GetDispatchKeyboardEventCode(action As JS_Keyboard, KeyCode As Key, Optional alt As Boolean = False, Optional shift As Boolean = False, Optional ctrl As Boolean = False) As String
    Dim eventActionName As String: eventActionName = keyboardEventNames_(action)
    
    Dim KeyName As String: KeyName = keyNames_(KeyCode)
    
    Dim optionArgs As String
    optionArgs = "{keyCode: " & KeyCode & ",key: '" & KeyName & "' ,bubbles: true"
    If alt Then optionArgs = optionArgs & ",altKey: true"
    If shift Then optionArgs = optionArgs & ",shiftKey: true"
    If ctrl Then optionArgs = optionArgs & ",ctrlKey: true "
    optionArgs = optionArgs & "}"
    
    GetDispatchKeyboardEventCode = _
    "function(){" & _
        "this.dispatchEvent(new KeyboardEvent('" & eventActionName & "'," & optionArgs & "));" & _
    "}"
End Function

'Public Function GetDispatchKeyArrowDownEventCode() As String
'    GetDispatchKeyArrowDownEventCode = _
'    "function(){" & _
'        "this.dispatchEvent(new KeyboardEvent('keydown', {keyCode: 40,key: 'ArrowDown' ,bubbles: true}));" & _
'    "}"
'End Function


'elementをFindした方法についての情報を配列にして返す(作成するElementに持たせるため）
Public Function MakeFindInfo(way_string As String, way As String, Optional index As Long = -1) As Collection
    Dim findInfo As Collection: Set findInfo = New Collection
    findInfo.Add way_string
    findInfo.Add way
    findInfo.Add index
    Set MakeFindInfo = findInfo
End Function

'Find系メソッドで要素が見つからなかった場合のエラーメッセージを作成する
Public Function MakeFindErrorMessage(wayString As String, way As String, Optional wait As String = "") As String
    Dim errMsg As String
    errMsg = _
    "＜FindElementBy" & wayString & "＞" & vbCrLf & _
    "指定された以下" & wayString & "の要素は存在しませんでした。" & vbCrLf & _
    way & vbCrLf & "(イミディエイトウィンドウで確認可能)" & vbCrLf & vbCrLf
    
    If wayString = "XPath" Then
        errMsg = errMsg & _
        "※引数のXPathにはChromeの検証ツールで「完全なXPath」が" & vbCrLf & _
        "　 コピペできるので、それを指定してください。" & vbCrLf & vbCrLf
    End If
    
    errMsg = errMsg & _
        "※もし取得したい要素がiframe内の要素である場合、SwitchFrameを" & vbCrLf & _
        "　 使用して、フレームを切り替えてからでないと取得できません。" & vbCrLf & vbCrLf
    
    errMsg = errMsg & _
    "※上記に沿っていても、要素が画面に出現する前だと失敗します。" & vbCrLf
    
    If wait = "" Then
        errMsg = errMsg & _
        "　 必要であればWaitForElementToAppearBy" & wayString & "で" & vbCrLf
    Else
        errMsg = errMsg & _
        "　 必要であればWaitForElementToAppearBy" & wait & "で" & vbCrLf
    End If
    
    errMsg = errMsg & _
    "　 要素の出現を確認する処理を入れましょう。" & vbCrLf
    
    MakeFindErrorMessage = errMsg
End Function

' 列挙体を画像フォーマット文字列へ変換
Public Function GetFormatName(imageFormat As Image) As String
    GetFormatName = imageFormatNames_(imageFormat)
End Function

' Base64文字列のデコードし、binary配列で返す
Public Function DecodeBase64(s As String) As Byte()
    Dim elm As Object: Set elm = CreateObject("MSXML2.DOMDocument").createElement("base64")
    elm.DataType = "bin.base64"
    elm.text = s
    DecodeBase64 = elm.nodeTypedValue
End Function

Public Function IncludeNGCharInFileName(fileName As String) As Boolean
    Dim i As Long
    For i = 1 To Len(fileName)
        If ngCharsAsFileName_.exists(Mid(fileName, i, 1)) Then
            IncludeNGCharInFileName = True
            Exit Function
        End If
    Next
    
    IncludeNGCharInFileName = False
End Function

' 画像ファイル名につけるインクリメント番号を取得
Public Function GetFileNumber(folderPath As String, formatName As String) As Long
    Dim fileName As String
    Dim num As Long: num = 1
    Do
        fileName = folderPath & "\screenshot" & num & "." & formatName
        If Dir(fileName) = "" Then
            GetFileNumber = num
            Exit Function
        Else
            num = num + 1
        End If
    Loop
End Function

Public Function ExistsSheetName(sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = sheetName Then
            ExistsSheetName = True
            Exit Function
        End If
    Next
    ExistsSheetName = False
End Function
