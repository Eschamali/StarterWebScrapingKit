Attribute VB_Name = "WinHttpMessageBridge"
Option Explicit

Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal msg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Const GWLP_WNDPROC As Long = -4
Public Const WM_APP As Long = &H8000&
Public Const WM_APP_WINHTTP_CALLBACK As Long = WM_APP + 711

Private g_hookHwnd As LongPtr
Private g_prevWndProc As LongPtr
Private g_msgCount As Long
Private g_target As WebSocketHTTPCommunicator

Public Sub InstallWinHttpMessageHook(ByVal targetHwnd As LongPtr, ByVal target As WebSocketHTTPCommunicator)
    If targetHwnd = 0 Then Exit Sub
    If g_hookHwnd = targetHwnd And g_prevWndProc <> 0 Then Exit Sub

    RemoveWinHttpMessageHook
    g_msgCount = 0
    Set g_target = target
    g_hookHwnd = targetHwnd
    g_prevWndProc = SetWindowLongPtr(g_hookHwnd, GWLP_WNDPROC, AddressOf WinHttpBridgeWndProc)
End Sub

Public Sub RemoveWinHttpMessageHook()
    On Error Resume Next
    If g_hookHwnd <> 0 And g_prevWndProc <> 0 Then
        SetWindowLongPtr g_hookHwnd, GWLP_WNDPROC, g_prevWndProc
    End If
    Set g_target = Nothing
    g_hookHwnd = 0
    g_prevWndProc = 0
    On Error GoTo 0
End Sub

Public Function GetWinHttpMessageCount() As Long
    GetWinHttpMessageCount = g_msgCount
End Function

Private Function WinHttpBridgeWndProc(ByVal hWnd As LongPtr, ByVal msg As Long, _
                                      ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If msg = WM_APP_WINHTTP_CALLBACK Then
        g_msgCount = g_msgCount + 1
        If Not (g_target Is Nothing) Then
            g_target.HandlePostedWinHttpCallback CLng(wParam), lParam
        End If
        WinHttpBridgeWndProc = 0
        Exit Function
    End If

    WinHttpBridgeWndProc = CallWindowProc(g_prevWndProc, hWnd, msg, wParam, lParam)
End Function
