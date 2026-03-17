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
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
' リサイズ用タイマー
Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
' ole32 ─ get_AdditionalBrowserArguments が返す文字列用（呼び出し側が CoTaskMemFree する）
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cb As LongPtr) As LongPtr
Private Declare PtrSafe Function IsEqualGUID Lib "ole32" (guid1 As Any, guid2 As Any) As Long
#End If

' IID 比較用（QI で EnvironmentOptions2 等の拡張を要求されたら E_NOINTERFACE）
Private Type Guid16
    d(0 To 15) As Byte
End Type
Private IID_EnvOpt  As Guid16  ' {2FDE08A8-1E9A-4766-8C05-95A9CEB9D1C5}
Private IID_EnvOpt2 As Guid16  ' {FF85C98A-1BA7-4A6B-90C8-2B752C89E9E2} EnvironmentOptions2
Private IID_IUnknown As Guid16 ' {00000000-0000-0000-C000-000000000046}

' グローバル参照
Public g_WebView2Core As WebView2Core

' リサイズタイマー ID（1ms ワンショット）
Public g_ResizeTimerId As LongPtr

' CDP CallDevToolsProtocolMethod 完了用（コールバック内は状態＋コピーのみ）
Public g_CDPResultJson  As String
Public g_CDPCompleted   As Boolean
Public g_CDPErrorCode   As Long

' ICoreWebView2EnvironmentOptions 偽装用。追加ブラウザ引数（Initialize で設定し get_ で返す）
Public g_EnvOptions_AdditionalArgs As String

Private Sub InitEnvOptIID()
    If IID_EnvOpt.d(0) = 0 And IID_EnvOpt.d(1) = 0 Then
        ' IID_ICoreWebView2EnvironmentOptions: {2FDE08A8-1E9A-4766-8C05-95A9CEB9D1C5}
        IID_EnvOpt.d(0) = &HA8: IID_EnvOpt.d(1) = &H8: IID_EnvOpt.d(2) = &HDE: IID_EnvOpt.d(3) = &H2F
        IID_EnvOpt.d(4) = &H9A: IID_EnvOpt.d(5) = &H1E: IID_EnvOpt.d(6) = &H66: IID_EnvOpt.d(7) = &H47
        IID_EnvOpt.d(8) = &H8C: IID_EnvOpt.d(9) = &H5: IID_EnvOpt.d(10) = &H95: IID_EnvOpt.d(11) = &HA9
        IID_EnvOpt.d(12) = &HCE: IID_EnvOpt.d(13) = &HB9: IID_EnvOpt.d(14) = &HD1: IID_EnvOpt.d(15) = &HC5
        ' IID_ICoreWebView2EnvironmentOptions2: {FF85C98A-1BA7-4A6B-90C8-2B752C89E9E2}
        IID_EnvOpt2.d(0) = &H8A: IID_EnvOpt2.d(1) = &HC9: IID_EnvOpt2.d(2) = &H85: IID_EnvOpt2.d(3) = &HFF
        IID_EnvOpt2.d(4) = &HA7: IID_EnvOpt2.d(5) = &H1B: IID_EnvOpt2.d(6) = &H6B: IID_EnvOpt2.d(7) = &H4A
        IID_EnvOpt2.d(8) = &H90: IID_EnvOpt2.d(9) = &HC8: IID_EnvOpt2.d(10) = &H2B: IID_EnvOpt2.d(11) = &H75
        IID_EnvOpt2.d(12) = &H2C: IID_EnvOpt2.d(13) = &H89: IID_EnvOpt2.d(14) = &HE9: IID_EnvOpt2.d(15) = &HE2
        ' IID_IUnknown: {00000000-0000-0000-C000-000000000046}
        IID_IUnknown.d(0) = 0: IID_IUnknown.d(1) = 0: IID_IUnknown.d(2) = 0: IID_IUnknown.d(3) = 0
        IID_IUnknown.d(4) = 0: IID_IUnknown.d(5) = 0
        IID_IUnknown.d(6) = &HC0: IID_IUnknown.d(7) = 0
        IID_IUnknown.d(8) = 0: IID_IUnknown.d(9) = 0: IID_IUnknown.d(10) = 0: IID_IUnknown.d(11) = 0
        IID_IUnknown.d(12) = 0: IID_IUnknown.d(13) = 0: IID_IUnknown.d(14) = &H46: IID_IUnknown.d(15) = 0
    End If
End Sub

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

'---------------------------------------------------------------------
' ICoreWebView2EnvironmentOptions 偽装（vtable 0..12、Options2 の ExclusiveUserDataFolderAccess 含む）
' QI: IUnknown / EnvironmentOptions / EnvironmentOptions2 をサポート
'---------------------------------------------------------------------
Public Function WV2_EnvOpt_QI(ByVal pThis As LongPtr, ByVal riid As LongPtr, ByVal ppvObject As LongPtr) As Long
    InitEnvOptIID
    Dim r As Guid16
    CopyMemory r, ByVal riid, 16
    If IsEqualGUID(r, IID_EnvOpt) = 0 And IsEqualGUID(r, IID_EnvOpt2) = 0 And IsEqualGUID(r, IID_IUnknown) = 0 Then
        WV2_EnvOpt_QI = &H80004002  ' E_NOINTERFACE
        Exit Function
    End If
    If ppvObject <> 0 Then CopyMemory ByVal ppvObject, pThis, 8
    WV2_EnvOpt_QI = 0  ' S_OK
End Function
Public Function WV2_EnvOpt_AddRef(ByVal pThis As LongPtr) As Long:  WV2_EnvOpt_AddRef = 1:  End Function
Public Function WV2_EnvOpt_Release(ByVal pThis As LongPtr) As Long:  WV2_EnvOpt_Release = 1:  End Function
' get_AdditionalBrowserArguments(LPWSTR* ppValue) - CoTaskMemAlloc で null 終端付きコピーを返す
Public Function WV2_EnvOpt_get_AdditionalBrowserArguments(ByVal pThis As LongPtr, ByVal ppValue As LongPtr) As Long
    If ppValue = 0 Then WV2_EnvOpt_get_AdditionalBrowserArguments = &H80070057: Exit Function  ' E_INVALIDARG
    Dim s As String: s = g_EnvOptions_AdditionalArgs
    If Len(s) = 0 Then CopyMemory ByVal ppValue, 0, 8: WV2_EnvOpt_get_AdditionalBrowserArguments = 0: Exit Function
    Dim cb As LongPtr: cb = (Len(s) + 1) * 2
    Dim pMem As LongPtr: pMem = CoTaskMemAlloc(cb)
    If pMem = 0 Then WV2_EnvOpt_get_AdditionalBrowserArguments = &H8007000E: Exit Function  ' E_OUTOFMEMORY
    CopyMemory ByVal pMem, ByVal StrPtr(s), CLng(cb)  ' null 終端込みでコピー
    CopyMemory ByVal ppValue, pMem, 8
    WV2_EnvOpt_get_AdditionalBrowserArguments = 0
End Function
' get_Language / get_TargetCompatibleBrowserVersion: NULL を返してデフォルト使用
Public Function WV2_EnvOpt_get_LanguageOrVersion(ByVal pThis As LongPtr, ByVal ppValue As LongPtr) As Long
    If ppValue <> 0 Then CopyMemory ByVal ppValue, 0, 8
    WV2_EnvOpt_get_LanguageOrVersion = 0  ' S_OK
End Function
' get_AllowSingleSignOnUsingOSPrimaryAccount: FALSE を返す
Public Function WV2_EnvOpt_get_AllowSSO(ByVal pThis As LongPtr, ByVal pAllow As LongPtr) As Long
    If pAllow <> 0 Then CopyMemory ByVal pAllow, 0, 4  ' BOOL = FALSE
    WV2_EnvOpt_get_AllowSSO = 0
End Function
Public Function WV2_EnvOpt_Stub(ByVal pThis As LongPtr, ByVal p1 As LongPtr, ByVal p2 As LongPtr) As Long
    WV2_EnvOpt_Stub = 0  ' put_* は S_OK で握りつぶす
End Function
Public Function WV2_EnvOpt_put_AdditionalBrowserArguments(ByVal pThis As LongPtr, ByVal pValue As LongPtr) As Long
    WV2_EnvOpt_put_AdditionalBrowserArguments = 0  ' S_OK（未使用だが一応）
End Function
' ICoreWebView2EnvironmentOptions2: get/put ExclusiveUserDataFolderAccess（vtable 11,12）
Public Function WV2_EnvOpt_get_ExclusiveUserDataFolderAccess(ByVal pThis As LongPtr, ByVal pValue As LongPtr) As Long
    If pValue <> 0 Then CopyMemory ByVal pValue, 0, 4  ' BOOL = FALSE
    WV2_EnvOpt_get_ExclusiveUserDataFolderAccess = 0
End Function
Public Function WV2_EnvOpt_put_ExclusiveUserDataFolderAccess(ByVal pThis As LongPtr, ByVal value As Long) As Long
    WV2_EnvOpt_put_ExclusiveUserDataFolderAccess = 0  ' S_OK
End Function

'---------------------------------------------------------------------
' ICoreWebView2CallDevToolsProtocolMethodCompletedHandler
' Invoke(errorCode, returnObjectAsJson As LPCWSTR) ? コールバック内は状態＋コピーのみ
'---------------------------------------------------------------------
Public Function WV2_CDPCB_QI(ByVal pThis As LongPtr, ByVal riid As LongPtr, ByVal ppvObject As LongPtr) As Long
    If ppvObject <> 0 Then CopyMemory ByVal ppvObject, pThis, LenB(pThis)
    WV2_CDPCB_QI = 0
End Function
Public Function WV2_CDPCB_AddRef(ByVal pThis As LongPtr) As Long:  WV2_CDPCB_AddRef = 1:  End Function
Public Function WV2_CDPCB_Release(ByVal pThis As LongPtr) As Long: WV2_CDPCB_Release = 1: End Function
Public Function WV2_CDPCB_Invoke(ByVal pThis As LongPtr, ByVal ErrorCode As Long, ByVal pResultJson As LongPtr) As Long
    g_CDPErrorCode = ErrorCode
    g_CDPResultJson = PtrToStrW(pResultJson)
    g_CDPCompleted = True
    Debug.Print "[CDP] Invoke: errorCode=" & ErrorCode & " pResult=0x" & Hex(pResultJson) & " len=" & Len(g_CDPResultJson)
    WV2_CDPCB_Invoke = 0
End Function

Private Function PtrToStrW(ByVal pWStr As LongPtr) As String
    If pWStr = 0 Then PtrToStrW = "": Exit Function
    Dim length As Long: length = lstrlenW(pWStr)
    If length <= 0 Then PtrToStrW = "": Exit Function
    Dim buf As String: buf = Space$(length)
    CopyMemory ByVal StrPtr(buf), ByVal pWStr, length * 2
    PtrToStrW = buf
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

Public Sub WV2_FillCDPFunctionPointers(ByRef cdpFn() As LongPtr)
    ReDim cdpFn(0 To 3)
    cdpFn(0) = GetFuncAddr(AddressOf WV2_CDPCB_QI)
    cdpFn(1) = GetFuncAddr(AddressOf WV2_CDPCB_AddRef)
    cdpFn(2) = GetFuncAddr(AddressOf WV2_CDPCB_Release)
    cdpFn(3) = GetFuncAddr(AddressOf WV2_CDPCB_Invoke)
End Sub

' ICoreWebView2EnvironmentOptions 偽装用 vtable（Options + Options2 の 13 エントリ）
Public Sub WV2_FillEnvOptionsFunctionPointers(ByRef fn() As LongPtr)
    ReDim fn(0 To 12)
    fn(0) = GetFuncAddr(AddressOf WV2_EnvOpt_QI)
    fn(1) = GetFuncAddr(AddressOf WV2_EnvOpt_AddRef)
    fn(2) = GetFuncAddr(AddressOf WV2_EnvOpt_Release)
    fn(3) = GetFuncAddr(AddressOf WV2_EnvOpt_get_AdditionalBrowserArguments)
    fn(4) = GetFuncAddr(AddressOf WV2_EnvOpt_put_AdditionalBrowserArguments)
    fn(5) = GetFuncAddr(AddressOf WV2_EnvOpt_get_LanguageOrVersion)   ' get_Language
    fn(6) = GetFuncAddr(AddressOf WV2_EnvOpt_Stub)                    ' put_Language
    fn(7) = GetFuncAddr(AddressOf WV2_EnvOpt_get_LanguageOrVersion)   ' get_TargetCompatibleBrowserVersion
    fn(8) = GetFuncAddr(AddressOf WV2_EnvOpt_Stub)                    ' put_TargetCompatibleBrowserVersion
    fn(9) = GetFuncAddr(AddressOf WV2_EnvOpt_get_AllowSSO)             ' get_AllowSingleSignOn
    fn(10) = GetFuncAddr(AddressOf WV2_EnvOpt_Stub)                   ' put_AllowSingleSignOn
    fn(11) = GetFuncAddr(AddressOf WV2_EnvOpt_get_ExclusiveUserDataFolderAccess)  ' Options2
    fn(12) = GetFuncAddr(AddressOf WV2_EnvOpt_put_ExclusiveUserDataFolderAccess)
End Sub
