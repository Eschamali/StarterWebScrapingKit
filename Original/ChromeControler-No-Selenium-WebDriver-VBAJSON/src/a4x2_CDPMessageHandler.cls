VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a4x2_CDPMessageHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private webSocket_ As a1_WebSocketCommunicator
Private json_ As a2_JSONHandler
Private events_ As a5_CDPEventHandler

Public Sub Init(WebSocket As a1_WebSocketCommunicator, _
                        json As a2_JSONHandler, _
                        events As a5_CDPEventHandler)
    Set webSocket_ = WebSocket
    Set json_ = json
    Set events_ = events
End Sub

''message処理送信＆受信
'受信時には、送信に対する受信だけではなく、クロームからのイベント通知も混ざってくる。
'イベント通知の場合、responsにidが含まれないため、idの有無で送信に対する受信なのかイベント通知なのかを判断
'注意:送信に成功しても、もし送信したJsonに不備がある場合、responsにはidが含まれない。
Public Function SendAndReceive(sendMsg As String) As String
    
    Dim sendId As Long: sendId = json_.GetValue(sendMsg, "id")
    
    Debug.Print vbCrLf & ">>>>>>" & "送信内容：" & sendMsg
    If webSocket_.SendMessage(sendMsg) = False Then GoTo quit

    Dim res As String, resId As Variant
    Do
        res = webSocket_.GetMessage
        If res = "受信エラー発生" Then GoTo quit
        
        resId = json_.GetValue(res, "id")
        If resId = "" Then
            If InStr(res, "method") > 0 Then
                events_.GetInfo res
                '制御を奪うダイアログが存在し、そのDialogOpenのイベント通知がくるとその後、ループ受信でとまる。
                '例：クリック送信⇒クリックに対する受信までの間にそのイベント通知が来ると、受信ができずに止まる。
                'そのため、DialogOpenの通知のみ特殊対応。抜けた後にはすぐダイアログ処理する必要あり。
                If events_.DialogInfoDic("IsExistDialog") Then Exit Function
            ElseIf InStr(res, "error") > 0 Then
                GoTo quit
            End If
        End If
    Debug.Print "受信：" & res
    Loop While (sendId <> resId)
    
    SendAndReceive = res
    
Exit Function

quit:
Dim errMsg As String
If res = "" Then
    errMsg = "送信エラー発生：イミディエイト参照"
ElseIf res = "受信エラー発生" Then
    errMsg = "受信エラー発生:イミディエイト参照"
Else
    errMsg = "クロームからのエラー通知あり:イミディエイト参照"
End If
MsgBox errMsg, vbCritical, "エラー"
webSocket_.Http.CloseHInternetHandles
End
End Function


