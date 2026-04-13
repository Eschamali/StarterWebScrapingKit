VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_ShadowDOMSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
' CDPBrowser 拡張 - Shadow DOM検索
'
' getElementByDeepCss, getElementsByDeepCss 等を通して Shadow DOM (mode: open) 内部の要素を
' 横断検索するための機能を提供します。
'***************************************************************************************************
Option Explicit

'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

Private Const ThisClassName         As String = "exCDP_ShadowDOMSearch"
Private EnableBrowserEvents         As Boolean

'***************************************************************************************************
'                               ■■■ 公開プロパティ ■■■
'***************************************************************************************************
Public Property Get EnableEvents() As Boolean
    EnableEvents = EnableBrowserEvents
End Property

Public Property Let EnableEvents(ByVal v As Boolean)
    EnableBrowserEvents = v
End Property

'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分編 ■■■
'***************************************************************************************************

'---------------------------------------------------------------------------------------------------
' Shadow DOM を含む CSS セレクタ一意検索
'---------------------------------------------------------------------------------------------------
Public Function getElementByDeepCss(strCss As String) As CDPElement
    Dim varJS As String: Dim result As Variant
    varJS = GetDeepCssCodes(strCss)(0)
    
    result = ex_CDPBrowser.jsEval(varJS, returnByValue:=False)
    
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByDeepCss = New CDPElement
    getElementByDeepCss.Init ex_CDPBrowser, varJS, VarType(result), oid
End Function

'---------------------------------------------------------------------------------------------------
' Shadow DOM を含む CSS セレクタ複数検索 (Collectionを返す)
'---------------------------------------------------------------------------------------------------
Public Function getElementsByDeepCss(strCss As String) As Collection
    Dim collResult As Variant
    Dim jsCode As String
    jsCode = GetDeepCssCountCodes(strCss)(0)
    
    collResult = ex_CDPBrowser.jsEval(jsCode, returnByValue:=True)
    
    Dim length_ As Long
    If VarType(collResult) = vbDouble Or VarType(collResult) = vbLong Or VarType(collResult) = vbInteger Then
        length_ = CLng(collResult)
    Else
        Set getElementsByDeepCss = New Collection
        Exit Function
    End If
    
    Dim elements As Collection: Set elements = New Collection
    Dim N As Long, result As Variant
    Dim oid As String
    
    For N = 0 To length_ - 1
        jsCode = GetDeepCssResultCodes(strCss, N)(0)
        result = ex_CDPBrowser.jsEval(jsCode, returnByValue:=False)
        
        If VarType(result) = vbString Then
            oid = CStr(result)
            Dim elem As CDPElement
            Set elem = New CDPElement
            elem.Init ex_CDPBrowser, jsCode, VarType(result), oid
            elements.Add elem
        End If
    Next N
    
    Set getElementsByDeepCss = elements
End Function

'---------------------------------------------------------------------------------------------------
' Shadow DOM を含む Text 一意検索
'---------------------------------------------------------------------------------------------------
Public Function getElementByDeepText(strText As String, Optional partialMatch As Boolean = False) As CDPElement
    Dim varJS As String: Dim result As Variant
    varJS = GetDeepTextCodes(strText, partialMatch)(0)
    
    result = ex_CDPBrowser.jsEval(varJS, returnByValue:=False)
    
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByDeepText = New CDPElement
    getElementByDeepText.Init ex_CDPBrowser, varJS, VarType(result), oid
End Function

'---------------------------------------------------------------------------------------------------
' Shadow DOM を含む Text 複数検索 (Collectionを返す)
'---------------------------------------------------------------------------------------------------
Public Function getElementsByDeepText(strText As String, Optional partialMatch As Boolean = False) As Collection
    Dim collResult As Variant
    Dim jsCode As String
    jsCode = GetDeepTextCountCodes(strText, partialMatch)(0)
    
    collResult = ex_CDPBrowser.jsEval(jsCode, returnByValue:=True)
    
    Dim length_ As Long
    If VarType(collResult) = vbDouble Or VarType(collResult) = vbLong Or VarType(collResult) = vbInteger Then
        length_ = CLng(collResult)
    Else
        Set getElementsByDeepText = New Collection
        Exit Function
    End If
    
    Dim elements As Collection: Set elements = New Collection
    Dim N As Long, result As Variant
    Dim oid As String
    
    For N = 0 To length_ - 1
        jsCode = GetDeepTextResultCodes(strText, partialMatch, N)(0)
        result = ex_CDPBrowser.jsEval(jsCode, returnByValue:=False)
        
        If VarType(result) = vbString Then
            oid = CStr(result)
            Dim elem As CDPElement
            Set elem = New CDPElement
            elem.Init ex_CDPBrowser, jsCode, VarType(result), oid
            elements.Add elem
        End If
    Next N
    
    Set getElementsByDeepText = elements
End Function


'***************************************************************************************************
'                           ■■■ 内部ヘルパー (JS生成) ■■■
'***************************************************************************************************

Private Function EscapeJSString(str As String) As String
    Dim res As String
    res = Replace(str, "\", "\\")
    res = Replace(res, "'", "\'")
    res = Replace(res, vbCrLf, "\n")
    res = Replace(res, vbCr, "\n")
    res = Replace(res, vbLf, "\n")
    EscapeJSString = res
End Function

Private Function GetDeepCssCodes(cssSelector As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "(function(){" & _
               "var selector='" & EscapeJSString(cssSelector) & "';" & _
               "function findIn(root, selector){" & _
               "if(!root){return null;}" & _
               "if(root.querySelector){" & _
               "var hit=root.querySelector(selector);" & _
               "if(hit){return hit;}" & _
               "}" & _
               "if(!root.querySelectorAll){return null;}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var i=0;i<nodes.length;i++){" & _
               "if(nodes[i].shadowRoot){" & _
               "var subHit=findIn(nodes[i].shadowRoot, selector);" & _
               "if(subHit){return subHit;}" & _
               "}" & _
               "}" & _
               "return null;" & _
               "}" & _
               "return findIn(document, selector);" & _
               "})();"
    GetDeepCssCodes = Codes
End Function

Private Function GetDeepCssCountCodes(cssSelector As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "(function(){" & _
               "var selector='" & EscapeJSString(cssSelector) & "';" & _
               "function collectIn(root, selector, out){" & _
               "if(!root){return;}" & _
               "if(root.querySelectorAll){" & _
               "var hits=root.querySelectorAll(selector);" & _
               "for(var i=0;i<hits.length;i++){" & _
               "out.push(hits[i]);" & _
               "}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var j=0;j<nodes.length;j++){" & _
               "if(nodes[j].shadowRoot){" & _
               "collectIn(nodes[j].shadowRoot, selector, out);" & _
               "}" & _
               "}" & _
               "}" & _
               "}" & _
               "var result=[];" & _
               "collectIn(document, selector, result);" & _
               "return result.length;" & _
               "})();"
    GetDeepCssCountCodes = Codes
End Function

Private Function GetDeepCssResultCodes(cssSelector As String, index As Long) As Variant
    Dim Codes(1) As String
    Codes(0) = "(function(){" & _
               "var selector='" & EscapeJSString(cssSelector) & "';" & _
               "function collectIn(root, selector, out){" & _
               "if(!root){return;}" & _
               "if(root.querySelectorAll){" & _
               "var hits=root.querySelectorAll(selector);" & _
               "for(var i=0;i<hits.length;i++){" & _
               "out.push(hits[i]);" & _
               "}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var j=0;j<nodes.length;j++){" & _
               "if(nodes[j].shadowRoot){" & _
               "collectIn(nodes[j].shadowRoot, selector, out);" & _
               "}" & _
               "}" & _
               "}" & _
               "}" & _
               "var result=[];" & _
               "collectIn(document, selector, result);" & _
               "return (result.length > " & CStr(index) & ") ? result[" & CStr(index) & "] : null;" & _
               "})();"
    GetDeepCssResultCodes = Codes
End Function

Private Function GetDeepTextCodes(text As String, partialMatch As Boolean) As Variant
    Dim Codes(1) As String
    Dim partial As String

    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Codes(0) = "(function(){" & _
               "var target='" & EscapeJSString(text) & "';" & _
               "var partial=" & partial & ";" & _
               "function checkNode(node, target, partial){" & _
               "if(partial){return node.textContent.indexOf(target) !== -1;}" & _
               "else{return node.textContent.trim() === target;}" & _
               "}" & _
               "function findIn(root, target, partial){" & _
               "if(!root){return null;}" & _
               "if(!root.querySelectorAll){return null;}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var i=0;i<nodes.length;i++){" & _
               "var node=nodes[i];" & _
               "if(node.shadowRoot){" & _
               "var subHit=findIn(node.shadowRoot, target, partial);" & _
               "if(subHit){return subHit;}" & _
               "}" & _
               "if(node.children.length === 0 && node.textContent){" & _
               "if(checkNode(node, target, partial)){return node;}" & _
               "}" & _
               "}" & _
               "return null;" & _
               "}" & _
               "return findIn(document, target, partial);" & _
               "})();"
    GetDeepTextCodes = Codes
End Function

Private Function GetDeepTextCountCodes(text As String, partialMatch As Boolean) As Variant
    Dim Codes(1) As String
    Dim partial As String

    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Codes(0) = "(function(){" & _
               "var target='" & EscapeJSString(text) & "';" & _
               "var partial=" & partial & ";" & _
               "function checkNode(node, target, partial){" & _
               "if(partial){return node.textContent.indexOf(target) !== -1;}" & _
               "else{return node.textContent.trim() === target;}" & _
               "}" & _
               "function collectIn(root, target, partial, out){" & _
               "if(!root){return;}" & _
               "if(!root.querySelectorAll){return;}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var i=0;i<nodes.length;i++){" & _
               "var node=nodes[i];" & _
               "if(node.children.length === 0 && node.textContent){" & _
               "if(checkNode(node, target, partial)){out.push(node);}" & _
               "}" & _
               "if(node.shadowRoot){" & _
               "collectIn(node.shadowRoot, target, partial, out);" & _
               "}" & _
               "}" & _
               "}" & _
               "var result=[];" & _
               "collectIn(document, target, partial, result);" & _
               "return result.length;" & _
               "})();"
    GetDeepTextCountCodes = Codes
End Function

Private Function GetDeepTextResultCodes(text As String, partialMatch As Boolean, index As Long) As Variant
    Dim Codes(1) As String
    Dim partial As String

    If partialMatch Then
        partial = "true"
    Else
        partial = "false"
    End If

    Codes(0) = "(function(){" & _
               "var target='" & EscapeJSString(text) & "';" & _
               "var partial=" & partial & ";" & _
               "function checkNode(node, target, partial){" & _
               "if(partial){return node.textContent.indexOf(target) !== -1;}" & _
               "else{return node.textContent.trim() === target;}" & _
               "}" & _
               "function collectIn(root, target, partial, out){" & _
               "if(!root){return;}" & _
               "if(!root.querySelectorAll){return;}" & _
               "var nodes=root.querySelectorAll('*');" & _
               "for(var i=0;i<nodes.length;i++){" & _
               "var node=nodes[i];" & _
               "if(node.children.length === 0 && node.textContent){" & _
               "if(checkNode(node, target, partial)){out.push(node);}" & _
               "}" & _
               "if(node.shadowRoot){" & _
               "collectIn(node.shadowRoot, target, partial, out);" & _
               "}" & _
               "}" & _
               "}" & _
               "var result=[];" & _
               "collectIn(document, target, partial, result);" & _
               "return (result.length > " & CStr(index) & ") ? result[" & CStr(index) & "] : null;" & _
               "})();"
    GetDeepTextResultCodes = Codes
End Function


'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_CDPBrowser.sleep seconds
End Sub

Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set invokeMethod = ex_CDPBrowser.invokeMethod(methodName, params, alwaysBrowserContext, StopError)
End Function

'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPBrowser)
    Set ex_CDPBrowser = Inheritance
End Sub

Private Sub Class_Initialize()
    EnableBrowserEvents = True
End Sub
