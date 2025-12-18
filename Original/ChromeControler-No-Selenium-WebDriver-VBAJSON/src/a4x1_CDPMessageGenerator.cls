VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a4x1_CDPMessageGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private id_ As Long
Private sessionId_ As String

Public Property Get Id() As Long
    Id = id_
End Property

Public Property Let SessionId(value As String)
    sessionId_ = value
End Property

''TargetDomain*******************************************************************
'pageDomainの通知を有効にする。
Public Function PageEnable() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.enable""}"
    PageEnable = message
End Function

Public Function SetDiscoverTargets() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.setDiscoverTargets""," & _
                      """params"": {""discover"": true}}"
    SetDiscoverTargets = message
End Function

Public Function AttachToTarget(targetId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.attachToTarget""," & _
                      """params"": {""targetId"": """ & targetId & """}}"
    AttachToTarget = message
End Function

Public Function ActivateTarget(targetId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.activateTarget""," & _
                      """params"": {""targetId"": """ & targetId & """}}"
    ActivateTarget = message
End Function

Public Function DetachFromTarget(targetId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.detachFromTarget""," & _
                      """params"": {""targetId"": """ & targetId & """}}"
    DetachFromTarget = message
End Function

Public Function TargetClose(targetId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.closeTarget""," & _
                      """params"": {""targetId"": """ & targetId & """}}"
    TargetClose = message
End Function
'利用可能なターゲットのリストを取得する。
Public Function GetTargets() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.getTargets""}"
    GetTargets = message
End Function
'*******************************************************************


'BrowserDomain*******************************************************************
Public Function SessionClose() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Browser.close""}"
    SessionClose = message
End Function
'*******************************************************************

'PageDomain*******************************************************************
Public Function URLOpen(URL As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.navigate""," & _
                      """params"": {""url"": """ & URL & """}}"
    URLOpen = message
End Function

Public Function HandleJavaScriptDialog() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.handleJavaScriptDialog""," & _
                     """params"": {""accept"": true}}"
    HandleJavaScriptDialog = message
End Function

Public Function GetFrameTree() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.getFrameTree""}"
    GetFrameTree = message
End Function

Public Function CreateIsolatedFrameWorld(FrameId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.createIsolatedWorld""," & _
                     """params"": {""frameId"": """ & FrameId & """,""grantUniveralAccess"": true}}"
    CreateIsolatedFrameWorld = message
End Function

Public Function CaptureScreenshot(format As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.captureScreenshot""," & _
                     """params"": {""format"": """ & format & """}}"
    CaptureScreenshot = message
End Function

'*******************************************************************
'Input Domain*******************************************************************
'キーイベントをページにディスパッチする。keyDown , keyUp, rawKeyDown, charのいずれか。
Public Function DispatchKeyEvent(action As String, modifiers As Long, windowsVirtualKeyCode As Long) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Input.dispatchKeyEvent""," & _
                      """params"": {""type"":  """ & action & """,""modifiers"": " & modifiers & ",""windowsVirtualKeyCode"": " & windowsVirtualKeyCode & "}}"
    DispatchKeyEvent = message

End Function

'マウスイベントをxy座標にディスパッチする。mousePressed, mouseReleased, mouseMoved, mouseWheelのいずれか。
Public Function DispatchMouseEvent(eventType As String, x As Long, y As Long, _
                                                        button As String, clickCount As Long, _
                                                        Optional modifiers As Long = 0, _
                                                        Optional deltaX As Long = 0, Optional deltaY As Long = 0) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Input.dispatchMouseEvent""," & _
                      """params"": {""type"":  """ & eventType & """,""x"": " & x & ",""y"": " & y & "," & _
                                            """button"": """ & button & """,""clickCount"": " & clickCount
                                            'none, left, middle, right, back, forwardのいずれか
                                            
    'Alt=1, Ctrl=2, Meta/Command=4, Shift=8
    If modifiers <> 0 Then
        message = message & ",""modifiers"": " & modifiers
    End If
    
    If deltaX <> 0 Or deltaY <> 0 Then
        message = message & ",""deltaX"": " & deltaX & ",""deltaY"": " & deltaY
    End If
    
    message = message & "}}"
    DispatchMouseEvent = message
End Function
'*******************************************************************
'DOM Domain*******************************************************************
'Domの有効化
Public Function DOM_Enable()
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.enable""}"
    DOM_Enable = message
End Function

'指定された要素にフォーカスを当てる。
Public Function Focus(objectId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.focus""," & _
                      """params"": {""objectId"":  """ & objectId & """}}"
    Focus = message
End Function

'DocumentObjectのNodeIdを取得
Public Function GetDocumentID() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.getDocument""," & _
                     """params"": {""depth"": " & 1 & "}}"
    GetDocumentID = message
End Function

'CSSセレクタによりNode取得
Public Function querySelector(Selector As String, Optional nodeId As Long = 1) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.querySelector""," & _
                      """params"": {""nodeId"":  " & nodeId & ",""selector"": """ & Selector & """}}"
    querySelector = message
End Function

'objectIdからnodeId取得
Public Function RequestNode(objectId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.requestNode""," & _
                      """params"": {""objectId"":  """ & objectId & """}}"
    RequestNode = message
End Function

'nodeIdからobjectId取得
Public Function ResolveNode(nodeId As Long)
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.resolveNode""," & _
                      """params"": {""nodeId"":  " & nodeId & "}}"
    ResolveNode = message
End Function
'idで指定されたノードの属性情報を取得
Public Function GetAttributes(nodeId As Long)
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.getAttributes""," & _
                      """params"":{""nodeId"":  " & nodeId & "}}"
    GetAttributes = message
End Function

'id で指定されたノード情報を叙述
Public Function DescribeNode(Id As Variant) As Variant
    id_ = id_ + 1
    Dim message As String
    If IsNumeric(Id) Then
        message = "{""id"": " & id_ & ",""method"": ""DOM.describeNode""," & _
                          """params"": {""nodeId"":  " & Id & ",""depth"": 1,""pierce"": false}}"
    Else
        message = "{""id"": " & id_ & ",""method"": ""DOM.describeNode""," & _
                         """params"": {""objectId"":  """ & Id & """,""depth"": 1,""pierce"": false}}"
    End If
    DescribeNode = message
End Function

Public Function RequestChildNodes() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.RequestChildNodes""," & _
                      """params"": {""nodeId"":  " & id_ & ",""depth"":  -1}}"
    RequestChildNodes = message
End Function

'指定された要素にフォーカスを当てる。
Public Function GetBoxModel(objectId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.getBoxModel""," & _
                      """params"": {""objectId"":  """ & objectId & """}}"
    GetBoxModel = message
End Function

'ファイル登録インプットボックスへファイルPath登録
Public Function SetFileInputFiles(objectId As Variant, files As String) As String
    id_ = id_ + 1
    files = "[""" & files & """]"
    
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.setFileInputFiles""," & _
                      """params"": " & _
                     "{""files"": " & files & ",""objectId"": """ & objectId & """}}"
    
    SetFileInputFiles = message
End Function

'RunTime Domain*******************************************************************
Public Function RunTime_Enable()
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Runtime.enable""}"
    RunTime_Enable = message
End Function
Public Function JavaScritpEvaluate(script As String)
    id_ = id_ + 1
    Dim message As String
'    If sessionId_ = "" Then
        message = "{""id"": " & id_ & ",""method"": ""Runtime.evaluate""," & _
                          """params"": {""expression"": """ & script & """,""returnByValue"": false}}"
'    Else
'        message = "{""id"": " & id_ & " ,""sessionId"": """ & sessionId_ & """,""method"": ""Runtime.evaluate""," & _
'                          """params"": {""expression"": """ & script & """,""returnByValue"": false}}"
'    End If
    JavaScritpEvaluate = message
End Function

Public Function CallJavaScriptFunction(funcDeclear As String, objectId As Variant, argments As Variant)
    id_ = id_ + 1
    Dim JavaScriptObjectArgs As String
    
    If argments = "true" Or argments = "false" Then
        JavaScriptObjectArgs = "[{""value"": " & argments & "}]"
    Else
        JavaScriptObjectArgs = "[{""value"": """ & argments & """}]"
    End If
    
    Dim message As String
'    If sessionId_ = "" Then
        message = "{""id"": " & id_ & ",""method"": ""Runtime.callFunctionOn""," & _
                          """params"": " & _
                         "{""functionDeclaration"": """ & funcDeclear & """,""arguments"": " & JavaScriptObjectArgs
'    Else
'        message = "{""id"": " & id_ & " ,""sessionId"": " & sessionId_ & """,""method"": ""Runtime.callFunctionOn""," & _
'                          """params"": " & _
'                         "{""functionDeclaration"": """ & funcDeclear & """,""arguments"": " & JavaScriptObjectArgs
'    End If
    
    If IsNumeric(objectId) Then
        message = message & ",""executionContextId"": " & objectId & "}}"
    Else
        message = message & ",""objectId"": """ & objectId & """}}"
    End If
    CallJavaScriptFunction = message
End Function

Public Function CallJavaScriptFunction_ArgMultiple(funcDeclear As String, objectId As Variant, ParamArray args() As Variant)
    id_ = id_ + 1
    Dim JavaScriptObjectArgs As String: JavaScriptObjectArgs = "[{""value"":{"
    
    Dim arg_arr As Variant
    arg_arr = args(0)(0)
    
    Dim i As Long
    For i = LBound(arg_arr) To UBound(arg_arr)
        If i Mod 2 = 0 Then
                JavaScriptObjectArgs = JavaScriptObjectArgs & """" & arg_arr(i) & """: "
        Else
            If arg_arr(i) = "true" Or arg_arr(i) = "false" Or IsNumeric(arg_arr(i)) Then
                JavaScriptObjectArgs = JavaScriptObjectArgs & arg_arr(i)
            Else
                JavaScriptObjectArgs = JavaScriptObjectArgs & """" & arg_arr(i) & ""","
            End If
        End If
        
        If i = UBound(arg_arr) Then JavaScriptObjectArgs = left(JavaScriptObjectArgs, Len(JavaScriptObjectArgs) - 1)
    Next
    
    JavaScriptObjectArgs = JavaScriptObjectArgs & "}}]"
    
    Dim message As String
        message = "{""id"": " & id_ & ",""method"": ""Runtime.callFunctionOn""," & _
                          """params"": " & _
                         "{""functionDeclaration"": """ & funcDeclear & """,""arguments"": " & JavaScriptObjectArgs

    
    If IsNumeric(objectId) Then
        message = message & ",""executionContextId"": " & objectId & "}}"
    Else
        message = message & ",""objectId"": """ & objectId & """}}"
    End If
    CallJavaScriptFunction_ArgMultiple = message
End Function
'*******************************************************************

