VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a2_JSONHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private html As Object
Private Sub Class_Initialize()
    Set html = createObject("htmlfile")
    html.Write "<meta http-equiv='X-UA-Compatible' content='IE=edge' />"

    ' URLエンコーディング
    html.Write _
        "<script>" & vbCrLf & _
            "document.EncodeURL = " & vbCrLf & _
            "function(s){" & vbCrLf & _
            "return encodeURI(s);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    
    html.Write _
        "<script>" & vbCrLf & _
            "document.EncodeURLComponent = " & vbCrLf & _
            "function(s){" & vbCrLf & _
            "return encodeURIComponent(s);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    
    ' URLデコード
    html.Write _
        "<script>" & vbCrLf & _
            "document.DecodeURL = " & vbCrLf & _
            "function(s){" & vbCrLf & _
            "return decodeURI(s);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    
    'URL文字列操作
    html.Write _
        "<script>" & vbCrLf & _
            "document.getParam = " & vbCrLf & _
            "function(s){" & vbCrLf & _
            "return s.search.substring(1);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
        
''JSON操作
'============================
    '配列からオブジェクト取り出す
    html.Write _
        "<script>" & vbCrLf & _
            "document.getObjectFromArrayBySetIndex = " & vbCrLf & _
            "function(targetJson,index){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return JSON.stringify(vals[index]);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    'オブジェクトからオブジェクトを取り出す
    html.Write _
        "<script>" & vbCrLf & _
            "document.getObjectFromObjectBySetKey = " & vbCrLf & _
            "function(targetJson,key){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return JSON.stringify(vals[key]);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    '配列から値を取り出す
    html.Write _
        "<script>" & vbCrLf & _
            "document.getValueFromArrayBySetIndex = " & vbCrLf & _
            "function(targetJson,index){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals[index];" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    'オブジェクトから値を取り出す
    html.Write _
        "<script>" & vbCrLf & _
            "document.getValueFromObjectBySetKey = " & vbCrLf & _
            "function(targetJson,key){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals[key];" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
        
    html.Write _
        "<script>" & vbCrLf & _
            "document.getJsonLength = " & vbCrLf & _
            "function(targetJson){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals.length;" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    
'    'JScriptTypeInfoObject
'    html.Write _
'        "<script>" & vbCrLf & _
'            "document.getJsonObject=" & vbCrLf & _
'            "function (targetJson) {" & vbCrLf & _
'            "return eval('(' + targetJson + ')');" & vbCrLf & _
'            "}" & vbCrLf & _
'        "</script>"
    
    html.Write _
        "<script>" & vbCrLf & _
            "document.getJsonString=" & vbCrLf & _
            "function (targetJson) {" & vbCrLf & _
            "var jsonObj = eval('(' + targetJson + ')');" & vbCrLf & _
            "return JSON.stringify(jsonObj);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
End Sub

Public Function EncodeURL(json As String) As String
    EncodeURL = html.EncodeURL(json)
End Function

Public Function EncodeURLCompornent(URLCompornent As String) As String
    URLCompornent = Replace(URLCompornent, "\", "\\")
    URLCompornent = Replace(URLCompornent, "'", "\'")
    EncodeURLCompornent = html.EncodeURLComponent(URLCompornent)
End Function

Public Function GetURLParamValue(URL As String, paramName As String) As String
    Dim QueryString As String
    
    QueryString = html.getParam(URL)
'     // 「&」で引数を分割して配列に
'    var pair = paramOld.split("&");
'    var i = temp = "";
'    var key = new Array();
'    for (i = 0; i < pair.length; i++){
'      // 配列の値を「=」で分割
'      temp = pair[i].split("=");
'      keyName = temp[0];
'      keyValue = temp[1];
'      // キーと値の連想配列を生成
'      key[keyName] = keyValue;
'    } // _for
'    var path = count = "";
'    path = Boolean(key["path"]);
'    count = parseInt(key["count"]);
'  } // _if_else
    GetURLParamValue = QueryString
End Function

Public Function DecodeURL(json As String) As String
    DecodeURL = html.DecodeURL(json)
End Function

Public Function GetValue(ByVal json As String, ParamArray argsArray() As Variant) As String
    Dim i As Long
    Dim result As String
    For i = LBound(argsArray) To UBound(argsArray)
        If json = "" Then
            GetValue = result
            Exit Function
        End If
        If i <> UBound(argsArray) Then
            If IsNumeric(argsArray(i)) Then
                json = html.getObjectFromArrayBySetIndex(json, argsArray(i))
            Else
                json = html.getObjectFromObjectBySetKey(json, argsArray(i))
            End If
        Else
            If IsNumeric(argsArray(i)) Then
                result = html.getValueFromArrayBySetIndex(json, argsArray(i))
            Else
                result = html.getValueFromObjectBySetKey(json, argsArray(i))
            End If
        End If
    Next
    GetValue = result
End Function

Public Function GetObject(ByVal json As String, ParamArray argsArray() As Variant) As String
    Dim i As Long
    Dim result As String
    For i = LBound(argsArray) To UBound(argsArray)
        If json = "" Then
            GetObject = json
            Exit Function
        End If
     
        If IsNumeric(argsArray(i)) Then
            json = html.getObjectFromArrayBySetIndex(json, argsArray(i))
        Else
            json = html.getObjectFromObjectBySetKey(json, argsArray(i))
        End If
    Next
    GetObject = json
End Function


Public Function GetJsonLength(json As String) As Long
    GetJsonLength = html.GetJsonLength(json)
End Function

'Public Function GetJScriptTypeInfoObject(json As String) As Object
'    Set GetJScriptTypeInfoObject = html.GetJsonObject(json)
'End Function

Public Function GetParsedAndStringifiedJsonString(json As String) As String
    GetParsedAndStringifiedJsonString = html.GetJsonString(json)
End Function

