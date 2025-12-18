VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a4_ExecuteCDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private MsgGenerate_ As a4x1_CDPMessageGenerator
Private MsgHandler_ As a4x2_CDPMessageHandler
Private json_ As a2_JSONHandler


Public Sub Init(MsgGenerate As a4x1_CDPMessageGenerator, _
                       Handler As a4x2_CDPMessageHandler, _
                       json As a2_JSONHandler)
                       
    Set MsgGenerate_ = MsgGenerate
    Set MsgHandler_ = Handler
    Set json_ = json
End Sub


''Browser系************************************************
Public Function SessionClose() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.SessionClose
    SessionClose = MsgHandler_.SendAndReceive(sendMsg)
End Function

'************************************************

'DOM(ElementId)系************************************************
'CDPMethodにないのでJavascript実行
Public Function GetDocumentObjectID() As String
    Dim resMsg  As String
    resMsg = ExecuteJavaScript("document;")
    GetDocumentObjectID = json_.GetValue(resMsg, "result", "result", "objectId")
End Function

Public Function GetDocumentNodeID() As Long
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.GetDocumentID
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    GetDocumentNodeID = json_.GetValue(resMsg, "result", "root", "nodeId")
End Function

Private Function ChangeObjectIdToNodeId(objectId As String) As String
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.RequestNode(objectId)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    ChangeObjectIdToNodeId = json_.GetValue(resMsg, "result", "nodeId")
End Function

Private Function ChangeNodeIdToObjectId(nodeId As Long) As String
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.ResolveNode(nodeId)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    ChangeNodeIdToObjectId = json_.GetValue(resMsg, "result", "objectId")
End Function
'************************************************

'DOM(Element情報)系************************************************
Private Function GetAttributes(nodeId As Long) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.GetAttributes(nodeId)
    GetAttributes = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function DescribeNode(Id As Variant) As Variant
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DescribeNode(Id)
    DescribeNode = MsgHandler_.SendAndReceive(sendMsg)
End Function
'************************************************

'Runtime************************************************
'executionContextCreated イベントによる実行コンテキスト生成の報告を有効にします。
'レポートが有効になると、イベントは既存の実行コンテキストごとに直ちに送信されます。
Public Function RuntimeEnable() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.RunTime_Enable()
    RuntimeEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function
'JavaScriptコード実行
Public Function ExecuteJavaScript(script As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.JavaScritpEvaluate(script)
    ExecuteJavaScript = MsgHandler_.SendAndReceive(sendMsg)
End Function

'JavaScript_Function実行
Public Function CallJavaScriptFunction(funcName As String, Id As Variant, args As Variant) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.CallJavaScriptFunction(funcName, Id, args)
    CallJavaScriptFunction = MsgHandler_.SendAndReceive(sendMsg)
End Function

'JavaScript_Function実行(引数複数版）
Public Function CallJavaScriptFunction_ArgMultiple(funcName As String, Id As Variant, ParamArray args() As Variant) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.CallJavaScriptFunction_ArgMultiple(funcName, Id, args)
    CallJavaScriptFunction_ArgMultiple = MsgHandler_.SendAndReceive(sendMsg)
End Function

'************************************************
'Input Domain*******************************************************************
'キーイベントをページにディスパッチする。keyDown , keyUp, rawKeyDown, charのいずれか。
Public Function DispatchKeyEvent(action As String, modifiers As Long, windowsVirtualKeyCode As Long) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DispatchKeyEvent(action, modifiers, windowsVirtualKeyCode)
    DispatchKeyEvent = MsgHandler_.SendAndReceive(sendMsg)
End Function

'マウスイベントをxy座標にディスパッチする。mousePressed, mouseReleased, mouseMoved, mouseWheelのいずれか。
Public Function DispatchMouseEvent(eventType As String, x As Long, y As Long, _
                                                        button As String, clickCount As Long, _
                                                        Optional modifiers As Long = 0, _
                                                        Optional deltaX As Long = 0, _
                                                        Optional deltaY As Long = 0) As String
                                                        
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DispatchMouseEvent( _
                        eventType, x, y, button, clickCount, _
                        modifiers, deltaX, deltaY)
    DispatchMouseEvent = MsgHandler_.SendAndReceive(sendMsg)
End Function

'*******************************************************************

'DOM操作系************************************************
Public Function DOMEnable() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DOM_Enable()
    DOMEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function Focus(objectId As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.Focus(objectId)
    Focus = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function GetBoxModel(objectId As String) As Variant
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.GetBoxModel(objectId)
    GetBoxModel = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function SetFileInputFiles(objectId As Variant, files As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.SetFileInputFiles(objectId, files)
    SetFileInputFiles = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Sub FindElementByQuerySelector(Selector As String)
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.querySelector(Selector)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    
    Dim nodeId As Long
    nodeId = json_.GetValue(resMsg, "result", "nodeId")
    Dim nodeInfoJson As String
    nodeInfoJson = DescribeNode(nodeId)
End Sub

'************************************************

'Input系************************************************
Public Function Click(objId As String) As String
    CallJavaScriptFunction "function(value){this.click();}", objId, ""
End Function
'************************************************

'Target:Page系************************************************
Public Function OpenURL(targetURL As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.URLOpen(targetURL)
    OpenURL = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function PageEnable() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.PageEnable()
    PageEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function GetFrameTree() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.GetFrameTree()
    GetFrameTree = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function CreateIsolatedFrameWorld(FrameId As String) As Long
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.CreateIsolatedFrameWorld(FrameId)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    If InStr(resMsg, "error") > 0 Then
        CreateIsolatedFrameWorld = 0
    Else
        CreateIsolatedFrameWorld = json_.GetValue(resMsg, "result", "executionContextId")
        Sleep 200
    End If
End Function

Public Function SetDiscoverTargets() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.SetDiscoverTargets()
    SetDiscoverTargets = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function CaptureScreenshot(formatName As String) As String
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.CaptureScreenshot(formatName)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)

    CaptureScreenshot = json_.GetValue(resMsg, "result", "data")
End Function

Public Function AttachToTarget(targetId As String) As String
    Dim sendMsg  As String, res As String, SessionId As String
    sendMsg = MsgGenerate_.AttachToTarget(targetId)
    res = MsgHandler_.SendAndReceive(sendMsg)
    SessionId = json_.GetValue(res, "result", "sessionId")
    MsgGenerate_.SessionId = SessionId
    AttachToTarget = res
End Function

Public Function ActivateTarget(targetId As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.ActivateTarget(targetId)
    ActivateTarget = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function DetachFromTarget(targetId As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DetachFromTarget(targetId)
    DetachFromTarget = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function TargetClose(targetId As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.TargetClose(targetId)
    TargetClose = MsgHandler_.SendAndReceive(sendMsg)
End Function
Public Function GetTargetInfo() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.GetTargets()
    GetTargetInfo = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function HandleAlertDialog() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.HandleJavaScriptDialog()
    HandleAlertDialog = MsgHandler_.SendAndReceive(sendMsg)
End Function

    

'自作Function系************************************************

'ページ表示待ちループ
Public Function WaitingForWebPageDisplay()
    Dim javaScriptCode As String, responseMsg As String, state As String
    Do
        Sleep 800
        javaScriptCode = "document.readyState;"
        responseMsg = ExecuteJavaScript(javaScriptCode)
        state = json_.GetValue(responseMsg, "result", "result", "value")
        DoEvents
    Loop Until (state = "complete")
End Function



'指定の方法で要素を探し(最大４秒間探す）、オブジェクトIDを返す
Public Function GetObjectIdOfElementBy(way As String, basicInfo As a3_BasicInfos, Codes As Variant) As String
    
    Dim startTime As Date: startTime = Timer()
    Dim res As String, objectId As String
    Do
        If basicInfo.ExecutionContextId = 0 Then
            res = ExecuteJavaScript((Codes(0)))
        Else
            res = CallJavascriptFunctionWrapper((Codes(1)), basicInfo, way)
        End If

        objectId = json_.GetValue(res, "result", "result", "objectId")
    
        If objectId <> "" Then Exit Do
        
        If CInt((Timer() - startTime)) > 4 Then Exit Do
        
        DoEvents:  Sleep 500
    Loop
    
    GetObjectIdOfElementBy = objectId
End Function

'CallJavaScriptFunctionを呼ぶ場合のラッパ―関数。エラー対応が必須のため
Public Function CallJavascriptFunctionWrapper(funcCode As String, basicInfo As a3_BasicInfos, args As Variant) As String
    Dim res As String
    res = CallJavaScriptFunction(funcCode, basicInfo.ExecutionContextId, args)
    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo.UpdateExecutionContextId Me
        res = CallJavaScriptFunction(funcCode, basicInfo.ExecutionContextId, args)
    End If
    CallJavascriptFunctionWrapper = res
End Function



'エレメントを作成し、戻り値として返却
'エレメント内容：ObjectId,NodeId,Attributes(辞書),フレームの場合のみframeId追加
Public Function MakeElementObject( _
        objectId As String, _
        findInfo As Collection, _
        basicInfo As a3_BasicInfos, _
        events As a5_CDPEventHandler, _
        Helper As a6_ExecuteHelperFunction) As IWebElement
    
    Dim info  As String
    info = DescribeNode(objectId)
    If InStr(info, "attributes") > 0 Then
        Dim attributes As Object
        Set attributes = MakeAttributesDic(json_.GetValue(info, "result", "node", "attributes"))
    End If
    If InStr(info, "frameId") > 0 Then
        Dim FrameId As String
        FrameId = json_.GetValue(info, "result", "node", "frameId")
    End If
    
    Dim element As b0x0_WebElement: Set element = New b0x0_WebElement
    element.Init json_, basicInfo, Me, events, Helper, _
                       objectId, findInfo, attributes, FrameId
    Set MakeElementObject = element
End Function
'elementの属性情報を辞書にして返す。
Private Function MakeAttributesDic(attributes As String) As Object
    Dim ary As Variant: ary = Split(attributes, ",")
    Dim attributesDic As Object: Set attributesDic = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(ary) To UBound(ary) - 1 Step 2
        If attributesDic.exists(ary(i)) = False Then
            attributesDic.Add ary(i), ary(i + 1)
        End If
    Next
    Set MakeAttributesDic = attributesDic
End Function


Public Function GetElementsCountBy(way As String, basicInfo As a3_BasicInfos, Codes As Variant) As Long
    Dim responseMsg As String
    If basicInfo.ExecutionContextId = 0 Then
        responseMsg = ExecuteJavaScript((Codes(0)))
    Else
        responseMsg = CallJavascriptFunctionWrapper((Codes(1)), basicInfo, way)
    End If
    GetElementsCountBy = json_.GetValue(responseMsg, "result", "result", "value")
End Function


'IsExists共通関数
Public Function IsExistsElementBy(way As String, basicInfo As a3_BasicInfos, Codes As Variant) As Boolean

    Dim responseMsg As String
    If basicInfo.ExecutionContextId = 0 Then
        responseMsg = ExecuteJavaScript((Codes(0)))
    Else
        responseMsg = CallJavascriptFunctionWrapper((Codes(1)), basicInfo, way)
    End If
    Dim objectId As String
    objectId = json_.GetValue(responseMsg, "result", "result", "objectId")

    If objectId = "" Then
        IsExistsElementBy = False
    Else
        IsExistsElementBy = True
    End If
End Function

'WaitForElement共通関数
Public Function WaitForElementToAppearBy( _
    way As String, maxWaitTime_sec As Long, basicInfo As a3_BasicInfos, Codes As Variant) As Boolean
    
    Dim startTime As Date: startTime = Timer()
    Do
        If IsExistsElementBy(way, basicInfo, Codes) Then
            WaitForElementToAppearBy = True
            Exit Function
        ElseIf CInt((Timer() - startTime)) > maxWaitTime_sec Then
            WaitForElementToAppearBy = False
            Exit Function
        End If
        DoEvents: Sleep 500
    Loop
End Function


' クエリ文字列部を取得
Public Function GetQueryStringFrom(URL As String) As String
    Dim javaScriptCode As String, responseMsg As String, result As String
    javaScriptCode = "var url = new URL('" & URL & "');" & _
                             "url.search;"
    responseMsg = ExecuteJavaScript(javaScriptCode)
    result = json_.GetValue(responseMsg, "result", "result", "value")
    GetQueryStringFrom = result
End Function


' 暗号論的乱数取得
Public Function GetSecureRondomNumber(count As Long) As String
    Dim javaScriptCode As String, responseMsg As String, result As String
    javaScriptCode = "var crpt =  Array.from( crypto.getRandomValues(new Uint32Array(" & count & ")) );" & _
                             "var s = '';" & _
                             "crpt.forEach(function(value){s += value + ',';});" & _
                             "s.slice(0,-1)"
    responseMsg = ExecuteJavaScript(javaScriptCode)
    result = json_.GetValue(responseMsg, "result", "result", "value")
    GetSecureRondomNumber = result
End Function

Public Function GetBase64EncodedSHA256(str As String) As String
    Dim javaScriptCode As String, responseMsg As String, result As String
    javaScriptCode = "var encoder = new TextEncoder();" & _
                             "var data = encoder.encode('" & str & "');" & _
                             "var base64urlHash = '';" & _
                             "window.crypto.subtle.digest('SHA-256', data).then(hash => {" & _
                             "var bytes = new Uint8Array(hash);" & _
                             "for (let i = 0; i < bytes.byteLength; i++) {base64urlHash += String.fromCharCode(bytes[i]);}" & _
                             "base64urlHash = btoa(base64urlHash);});"
    responseMsg = ExecuteJavaScript(javaScriptCode)
    javaScriptCode = "base64urlHash"
    responseMsg = ExecuteJavaScript(javaScriptCode)
    result = json_.GetValue(responseMsg, "result", "result", "value")
    GetBase64EncodedSHA256 = result
End Function

Public Function ExecuteDirect(CDPMethodJson As String) As String
        Dim sendMsg  As String
        ExecuteDirect = MsgHandler_.SendAndReceive(CDPMethodJson)
End Function
