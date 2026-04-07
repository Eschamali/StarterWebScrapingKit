Attribute VB_Name = "Demo_WebSocketViaNamedPipe"
Option Explicit

Sub testtest()
    Dim WebSocket As New WebSocketViaNamedPipe
    WebSocket.OpenAndConnectNamePipe
End Sub


Sub fwrji()
    Dim WebSocket As New WebSocketViaNamedPipe
    WebSocket.deserialize
    WebSocket.ClosePipeCDP
End Sub
