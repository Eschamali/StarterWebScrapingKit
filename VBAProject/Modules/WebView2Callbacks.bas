Attribute VB_Name = "WebView2Callbacks"
Option Explicit
'***
' WebView2Callbacks.bas
' VBA の制約上、AddressOf はクラスモジュールで Win32 コールバックとして使用できないため、
' このモジュールに COM コールバックのシムを集約する。
' WebView2Core.cls からの指示で動作し、すべての実ロジックは WebView2Core に委譲する。
'***

' QueryInterface の ppvObject 書き込みに使用する
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal length As Long)
' リサイズ用タイマー
Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
#End If

' グローバル参照
Public g_WebView2Core As WebView2Core

' リサイズタイマー ID（1ms ワンショット）
Public g_ResizeTimerId As LongPtr

'----------------------------------------------------------------------
' GetFuncAddr
'   AddressOf は変数への直接代入不可（VBA言語制約）。
'   「ByVal As LongPtr 引数」として渡すことで関数ポインタを取り出すラッパー。
'----------------------------------------------------------------------
Public Function GetFuncAddr(ByVal pfn As LongPtr) As LongPtr
    GetFuncAddr = pfn
End Function

'----------------------------------------------------------------------
' WV2_ScheduleResize
'   UserForm_Resize から呼ばれ、1ms 後にタイマーコールバックを発火させる。
'   put_Bounds は WM_SIZE ハンドラの外側（TimerProc）で呼ぶ必要があるため。
'----------------------------------------------------------------------
Public Sub WV2_ScheduleResize()
    If g_ResizeTimerId <> 0 Then KillTimer 0, g_ResizeTimerId  ' 既存タイマーをキャンセル
    g_ResizeTimerId = SetTimer(0, 0, 1, GetFuncAddr(AddressOf WV2_ResizeTimerProc))
End Sub

'----------------------------------------------------------------------
' WV2_ResizeTimerProc
'   SetTimer のコールバック。WM_SIZE 処理完了後、Excel のメッセージループ
'   内の DispatchMessage から呼ばれる。このコンテキストは put_Bounds に安全。
'   FinishWebViewSetup の put_Bounds も同様のコンテキストで動作している。
'----------------------------------------------------------------------
Public Function WV2_ResizeTimerProc(ByVal hWndTimer As LongPtr, ByVal uMsg As Long, _
                                     ByVal nIDEvent As LongPtr, ByVal dwTime As Long) As Long
    KillTimer hWndTimer, nIDEvent  ' ワンショット：即キャンセル
    g_ResizeTimerId = 0
    On Error Resume Next
    If Not g_WebView2Core Is Nothing Then g_WebView2Core.DoTimerResize
    WV2_ResizeTimerProc = 0
End Function

'---------------------------------------------------------------------
' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
' vtable[0] QueryInterface / [1] AddRef / [2] Release / [3] Invoke
'---------------------------------------------------------------------
Public Function WV2_EnvCB_QI(ByVal pThis As LongPtr, ByVal riid As LongPtr, ByVal ppvObject As LongPtr) As Long
    If ppvObject <> 0 Then CopyMemory ByVal ppvObject, pThis, LenB(pThis)
    WV2_EnvCB_QI = 0   ' S_OK
End Function
Public Function WV2_EnvCB_AddRef(ByVal pThis As LongPtr) As Long:  WV2_EnvCB_AddRef = 1:  End Function
Public Function WV2_EnvCB_Release(ByVal pThis As LongPtr) As Long: WV2_EnvCB_Release = 1: End Function
Public Function WV2_EnvCB_Invoke(ByVal pThis As LongPtr, ByVal ErrorCode As Long, ByVal pEnv As LongPtr) As Long
    Debug.Print "[WV2] EnvCB_Invoke fired: errorCode=" & ErrorCode & ", pEnv=" & Hex(pEnv)
    On Error Resume Next
    If Not g_WebView2Core Is Nothing Then g_WebView2Core.CB_EnvironmentCreated ErrorCode, pEnv
    WV2_EnvCB_Invoke = 0    ' S_OK
End Function

'---------------------------------------------------------------------
' ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
' vtable[0] QueryInterface / [1] AddRef / [2] Release / [3] Invoke
'---------------------------------------------------------------------
Public Function WV2_CtrlCB_QI(ByVal pThis As LongPtr, ByVal riid As LongPtr, ByVal ppvObject As LongPtr) As Long
    If ppvObject <> 0 Then CopyMemory ByVal ppvObject, pThis, LenB(pThis)
    WV2_CtrlCB_QI = 0   ' S_OK
End Function
Public Function WV2_CtrlCB_AddRef(ByVal pThis As LongPtr) As Long:  WV2_CtrlCB_AddRef = 1:  End Function
Public Function WV2_CtrlCB_Release(ByVal pThis As LongPtr) As Long: WV2_CtrlCB_Release = 1: End Function
Public Function WV2_CtrlCB_Invoke(ByVal pThis As LongPtr, ByVal ErrorCode As Long, ByVal pCtrl As LongPtr) As Long
    Debug.Print "[WV2] CtrlCB_Invoke fired: errorCode=" & ErrorCode & ", pCtrl=" & Hex(pCtrl)
    On Error Resume Next
    If Not g_WebView2Core Is Nothing Then g_WebView2Core.CB_ControllerCreated ErrorCode, pCtrl
    WV2_CtrlCB_Invoke = 0   ' S_OK
End Function

'--- AddressOf テーブルを WebView2Core に渡すヘルパー ----------------
Public Sub WV2_FillFunctionPointers(ByRef envFn() As LongPtr, ByRef ctrlFn() As LongPtr)
    ReDim envFn(0 To 3)
    envFn(0) = GetFuncAddr(AddressOf WV2_EnvCB_QI)
    envFn(1) = GetFuncAddr(AddressOf WV2_EnvCB_AddRef)
    envFn(2) = GetFuncAddr(AddressOf WV2_EnvCB_Release)
    envFn(3) = GetFuncAddr(AddressOf WV2_EnvCB_Invoke)

    ReDim ctrlFn(0 To 3)
    ctrlFn(0) = GetFuncAddr(AddressOf WV2_CtrlCB_QI)
    ctrlFn(1) = GetFuncAddr(AddressOf WV2_CtrlCB_AddRef)
    ctrlFn(2) = GetFuncAddr(AddressOf WV2_CtrlCB_Release)
    ctrlFn(3) = GetFuncAddr(AddressOf WV2_CtrlCB_Invoke)
End Sub
