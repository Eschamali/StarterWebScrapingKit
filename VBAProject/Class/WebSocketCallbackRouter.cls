VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebSocketCallbackRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'   WebSocket 非同期コールバックの hWnd ルーティング + VirtualAlloc ネイティブ thunk
'   AddressOf / .bas / RtlMoveMemory 不使用（PointerAccessor + VirtualAlloc のみ）
'
'   x64 WinHTTP callback: PostMessageW のみ（ワーカースレッドから VBA 非呼び出し）
'   WndProc フック       : メインスレッド上で vtable 経由 OnBridgeWndProc 呼び出し
'***************************************************************************************************
Option Explicit

Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
    ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" ( _
    ByVal lpModuleName As LongPtr) As LongPtr
Private Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageW" ( _
    lpMsg As msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, _
    ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Private Const PM_REMOVE As Long = 1

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type msg
    hWnd As LongPtr
    message As Long
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    pt As POINTAPI
End Type

Private Const GWLP_WNDPROC As Long = -4
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const MEM_RELEASE As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const WM_APP As Long = &H8000&
Private Const WM_APP_WINHTTP_CALLBACK As Long = WM_APP + 711
Private Const CALLBACK_VTABLE_INDEX As Long = 0
Private Const VTABLE_USER_METHOD_SLOT As Long = 7

Private g_msgCount As Long
Private g_targetByHwnd As Dictionary
Private m_pPostMessageCallbackThunk As LongPtr
Private m_pHookWndProcThunk As LongPtr
Private Const THUNK_BUILD_ID As Long = 2
Private m_thunkBuildId As Long

#If VBA7 = 0 Then
    Private Enum LONG_PTR: [_]: End Enum
#End If

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize As Long = 8
#Else
    Private Const NullPtr As Long = 0&
    Private Const PtrSize As Long = 4
#End If

Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims     As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks    As Long
    pvData    As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type
Private Type PointerAccessor
    arr() As LongPtr
    sa    As SAFEARRAY_1D
End Type

'***************************************************************************************************
' !! vtable 先頭 Public メンバー固定（GetVTableProcPtr index=0）!!
'***************************************************************************************************
Public Function OnBridgeWndProc(ByVal hWnd As LongPtr, ByVal msg As Long, _
                                ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    EnsureHookMaps
    Dim Key As String
    Key = HwndKey(hWnd)

    If msg = WM_APP_WINHTTP_CALLBACK Then g_msgCount = g_msgCount + 1

    If g_targetByHwnd.Exists(Key) Then
        Dim t As WebSocketHTTPCommunicator
        Set t = g_targetByHwnd(Key)
        If Not t Is Nothing Then
            OnBridgeWndProc = t.DispatchHookedMessage(hWnd, msg, wParam, lParam)
            Exit Function
        End If
    End If

    OnBridgeWndProc = 0
End Function

Public Property Get AppWinHttpCallbackMsg() As Long
    AppWinHttpCallbackMsg = WM_APP_WINHTTP_CALLBACK
End Property

Public Function GetWinHttpStatusCallbackPtr(ByVal Target As WebSocketHTTPCommunicator) As LongPtr
#If Win64 Then
    If m_pPostMessageCallbackThunk = 0 Or m_thunkBuildId <> THUNK_BUILD_ID Then
        m_pPostMessageCallbackThunk = BuildPostMessageCallbackThunkX64()
        m_thunkBuildId = THUNK_BUILD_ID
    End If
    GetWinHttpStatusCallbackPtr = m_pPostMessageCallbackThunk
#Else
    If Target Is Nothing Then Exit Function
    GetWinHttpStatusCallbackPtr = BuildWinHttpCallbackJumpThunkX86(Target)
#End If
End Function

Public Function ReadMemLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadMemLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Public Sub InstallWinHttpMessageHook(ByVal targetHwnd As LongPtr, ByVal Target As WebSocketHTTPCommunicator)
    If targetHwnd = 0 Then Exit Sub
    If Target Is Nothing Then Exit Sub

    EnsureHookMaps
    Dim Key As String
    Key = HwndKey(targetHwnd)

    Dim prev As LongPtr

    ' x64/x86 共通: notify hwnd の WndProc を安全サンクへ差し替える。
    ' WinHTTP コールバック本体は worker スレッド→PostMessage サンク（VBA 非呼び出し）。
    ' WM_APP は UI スレッドの DispatchMessage 経由でこの WndProc サンク→OnBridgeWndProc→
    ' DispatchHookedMessage と自動排水される（旧 WebSocketAsyncHelpers と同じ構図）。
    ' これにより DoEvents だけで断片受信が滞りなく届き、isReserved の固着を防ぐ。
    If g_targetByHwnd.Exists(Key) Then
        Dim existing As WebSocketHTTPCommunicator
        Set existing = g_targetByHwnd(Key)
        If Not existing Is Nothing Then prev = existing.prevWndProc
    Else
        Dim safeWndProc As LongPtr
        safeWndProc = GetHookWndProcThunkPtr()
        If safeWndProc = 0 Then Exit Sub
        prev = SetWindowLongPtr(targetHwnd, GWLP_WNDPROC, safeWndProc)
    End If

    Target.RegisterAsHookTarget targetHwnd, prev
    Set g_targetByHwnd(Key) = Target
End Sub

' x64: worker→PostMessage された WM_APP を UI スレッドで PeekMessage 排水（VBA/COM 安全）
Public Sub PumpPostedWinHttpMessages(Optional ByVal targetHwnd As LongPtr = 0)
    On Error Resume Next
    EnsureHookMaps

    If targetHwnd <> 0 Then
        PumpPostedWinHttpMessagesForHwnd targetHwnd
    Else
        Dim k As Variant
        For Each k In g_targetByHwnd.keys
            PumpPostedWinHttpMessagesForHwnd CLngPtr(k)
        Next k
    End If
    On Error GoTo 0
End Sub

Private Sub PumpPostedWinHttpMessagesForHwnd(ByVal targetHwnd As LongPtr)
    Dim msg As msg
    Dim Key As String
    Key = HwndKey(targetHwnd)

    Do While PeekMessage(msg, targetHwnd, WM_APP_WINHTTP_CALLBACK, WM_APP_WINHTTP_CALLBACK, PM_REMOVE) <> 0
        g_msgCount = g_msgCount + 1
        If g_targetByHwnd.Exists(Key) Then
            Dim t As WebSocketHTTPCommunicator
            Set t = g_targetByHwnd(Key)
            If Not t Is Nothing Then t.HandlePostedWinHttpCallback msg.wParam, msg.lParam
        End If
    Loop
End Sub

Public Sub RemoveWinHttpMessageHook(Optional ByVal targetHwnd As LongPtr)
    On Error Resume Next
    EnsureHookMaps

    If targetHwnd = 0 Then
        Dim k As Variant
        For Each k In g_targetByHwnd.keys
            Dim t As WebSocketHTTPCommunicator
            Set t = g_targetByHwnd(k)
            If Not t Is Nothing Then t.RestoreWndProcIfNeeded
        Next k
        g_targetByHwnd.RemoveAll
        On Error GoTo 0
        Exit Sub
    End If

    Dim Key As String: Key = HwndKey(targetHwnd)
    If g_targetByHwnd.Exists(Key) Then
        Dim ti As WebSocketHTTPCommunicator
        Set ti = g_targetByHwnd(Key)
        If Not ti Is Nothing Then ti.RestoreWndProcIfNeeded
        g_targetByHwnd.Remove Key
    End If
    On Error GoTo 0
End Sub

Public Function GetWinHttpMessageCount() As Long
    GetWinHttpMessageCount = g_msgCount
End Function

Private Function GetHookWndProcThunkPtr() As LongPtr
    If m_pHookWndProcThunk <> 0 Then
        GetHookWndProcThunkPtr = m_pHookWndProcThunk
        Exit Function
    End If
    Dim procAddr As LongPtr
    procAddr = GetObjectVTableProc(Me, VTABLE_USER_METHOD_SLOT + CALLBACK_VTABLE_INDEX)
    If procAddr = 0 Then Exit Function
#If Win64 Then
    m_pHookWndProcThunk = BuildHookWndProcThunkX64(Me, procAddr)
#Else
    m_pHookWndProcThunk = BuildHookWndProcThunkX86(Me, procAddr)
#End If
    GetHookWndProcThunkPtr = m_pHookWndProcThunk
End Function

#If Win64 Then
Private Function BuildPostMessageCallbackThunkX64() As LongPtr
    Dim postMessageProc As LongPtr
    postMessageProc = GetPostMessageWProc()
    If postMessageProc = 0 Then Exit Function

    Dim p As Long: p = 0
    Dim bb() As Byte: ReDim bb(0 To 63)
    ' PostMessageW(RDX=notifyHwnd, WM_APP_WINHTTP_CALLBACK, packed(status+len), packed(wsStatus))
    EmitBytes bb, p, &H48, &H89, &HD1                          ' MOV RCX, RDX  (hWnd = dwContext)
    EmitBytes bb, p, &H48, &HC7, &HC2                            ' MOV RDX, imm32
    EmitDword bb, p, WM_APP_WINHTTP_CALLBACK
    EmitBytes bb, p, &H4C, &H8B, &H54, &H24, &H40                  ' MOV R10, [RSP+40h] 第5引数
    EmitBytes bb, p, &H49, &HC1, &HE2, &H20
    EmitBytes bb, p, &H45, &H8B, &HD8
    EmitBytes bb, p, &H4D, &HB, &HD3
    EmitBytes bb, p, &H4D, &H89, &HD0
    EmitBytes bb, p, &H4D, &H85, &HC9
    EmitBytes bb, p, &H74, &H3
    EmitBytes bb, p, &H4D, &H8B, &H9
    EmitBytes bb, p, &H48, &H83, &HEC, &H28
    EmitBytes bb, p, &H48, &HB8
    EmitPtr bb, p, postMessageProc
    EmitBytes bb, p, &HFF, &HD0
    EmitBytes bb, p, &H48, &H83, &HC4, &H28
    EmitBytes bb, p, &HC3

    BuildPostMessageCallbackThunkX64 = VirtualAlloc(0, p, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If BuildPostMessageCallbackThunkX64 = 0 Then Exit Function
    WriteBytesToNative BuildPostMessageCallbackThunkX64, bb, p
End Function

Private Function BuildHookWndProcThunkX64(ByVal o As Object, ByVal comProcAddress As LongPtr) As LongPtr
    Dim lEbMode As LongPtr
    lEbMode = GetEbModeProc()

    BuildHookWndProcThunkX64 = VirtualAlloc(0, 192, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If BuildHookWndProcThunkX64 = 0 Or comProcAddress = 0 Or ObjPtr(o) = 0 Then Exit Function

    Dim p As Long: p = 0
    Dim bb() As Byte: ReDim bb(0 To 191)
    Dim jneAt As Long
    Dim retZeroAt As Long

    EmitBytes bb, p, &H48, &H83, &HEC, &H58
    EmitBytes bb, p, &H48, &H89, &H4C, &H24, &H38
    EmitBytes bb, p, &H48, &H89, &H54, &H24, &H40
    EmitBytes bb, p, &H4C, &H89, &H44, &H24, &H48
    EmitBytes bb, p, &H4C, &H89, &H4C, &H24, &H50

    EmitBytes bb, p, &H48, &HB8
    EmitPtr bb, p, lEbMode
    EmitBytes bb, p, &H48, &H85, &HC0
    EmitBytes bb, p, &H74, &H7
    EmitBytes bb, p, &HFF, &HD0
    EmitBytes bb, p, &H83, &HF8, &H1
    jneAt = p + 1
    EmitBytes bb, p, &H75, &H0

    EmitBytes bb, p, &H48, &HB9
    EmitPtr bb, p, ObjPtr(o)
    EmitBytes bb, p, &H48, &H8B, &H54, &H24, &H38
    EmitBytes bb, p, &H4C, &H8B, &H44, &H24, &H40
    EmitBytes bb, p, &H4C, &H8B, &H4C, &H24, &H48
    EmitBytes bb, p, &H48, &H8B, &H44, &H24, &H50
    EmitBytes bb, p, &H48, &H89, &H44, &H24, &H20
    EmitBytes bb, p, &H48, &HB8
    EmitPtr bb, p, comProcAddress
    EmitBytes bb, p, &HFF, &HD0

    EmitBytes bb, p, &H48, &H83, &HC4, &H58
    EmitBytes bb, p, &HC3

    retZeroAt = p
    EmitBytes bb, p, &H48, &H83, &HC4, &H58
    EmitBytes bb, p, &HC3

    bb(jneAt) = CByte(retZeroAt - (jneAt + 1))

    WriteBytesToNative BuildHookWndProcThunkX64, bb, p
End Function
#Else
Private Function BuildWinHttpCallbackJumpThunkX86(ByVal Target As WebSocketHTTPCommunicator) As LongPtr
    Dim tProcPtr As LongPtr
    tProcPtr = GetObjectVTableProc(Target, VTABLE_USER_METHOD_SLOT)
    If tProcPtr = 0 Then Exit Function

    BuildWinHttpCallbackJumpThunkX86 = VirtualAlloc(0, 32, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If BuildWinHttpCallbackJumpThunkX86 = 0 Then Exit Function

    Dim p As LongPtr: p = BuildWinHttpCallbackJumpThunkX86
    WriteThunkDword p, &H824448B:  p = p + 4
    WriteThunkDword p, &H4244489:  p = p + 4
    WriteThunkDword p, &HB8&:      p = p + 1
    WriteThunkDword p, tProcPtr:   p = p + 4
    WriteThunkDword p, &HE0FF&:    p = p + 2
    WriteThunkDword p, &H14C2&
End Function

Private Function BuildHookWndProcThunkX86(ByVal o As Object, ByVal comProcAddress As LongPtr) As LongPtr
    If comProcAddress = 0 Or ObjPtr(o) = 0 Then Exit Function

    BuildHookWndProcThunkX86 = VirtualAlloc(0, 48, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If BuildHookWndProcThunkX86 = 0 Then Exit Function

    Dim p As LongPtr: p = BuildHookWndProcThunkX86
    WriteThunkDword p, &H55:           p = p + 1
    WriteThunkDword p, &H89E5:         p = p + 2
    WriteThunkDword p, &H83EC04:       p = p + 3
    WriteThunkDword p, &HFF7514:       p = p + 3
    WriteThunkDword p, &HFF7510:       p = p + 3
    WriteThunkDword p, &HFF750C:       p = p + 3
    WriteThunkDword p, &HFF7508:       p = p + 3
    WriteThunkDword p, &H68:            p = p + 1
    WriteThunkDword p, ObjPtr(o):      p = p + 4
    WriteThunkDword p, &HB8:           p = p + 1
    WriteThunkDword p, comProcAddress: p = p + 4
    WriteThunkDword p, &HFFD0:         p = p + 2
    WriteThunkDword p, &H89EC:         p = p + 2
    WriteThunkDword p, &H5D:           p = p + 1
    WriteThunkDword p, &HC21000
End Function
#End If

Private Function GetPostMessageWProc() As LongPtr
    Dim hUser32 As LongPtr
    hUser32 = GetModuleHandleW(StrPtr("user32.dll"))
    If hUser32 = 0 Then Exit Function
    GetPostMessageWProc = GetProcAddress(hUser32, "PostMessageW")
End Function

Private Function GetEbModeProc() As LongPtr
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then
        Err.Clear
        On Error GoTo 0
        Dim hVBE As LongPtr
        hVBE = GetModuleHandle("vbe7.dll")
        If hVBE <> 0 Then GetEbModeProc = GetProcAddress(hVBE, "EbMode")
    End If
    On Error GoTo 0
End Function

Private Function GetObjectVTableProc(ByVal o As Object, ByVal vtableSlot As Long) As LongPtr
    Static pa(0 To 0) As PointerAccessor

    If pa(0).sa.cDims = 0 Then
        With pa(0)
            .sa.cDims = 1
            .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
            .sa.cbElements = PtrSize
            .sa.cLocks = 1
            WritePtrNatively pa, VarPtr(.sa)
        End With
    End If

    With pa(0)
        .sa.pvData = ObjPtr(o)
        .sa.rgsabound0.cElements = 1
        .sa.pvData = .arr(0) + vtableSlot * PtrSize
        .sa.rgsabound0.cElements = 1
        GetObjectVTableProc = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub

Private Sub WriteBytesToNative(ByVal dest As LongPtr, ByRef bb() As Byte, ByVal N As Long)
    Dim i As Long
    For i = 0 To N - 1
        WriteNativeByte dest + i, bb(i)
    Next i
End Sub

Private Sub WriteNativeByte(ByVal addr As LongPtr, ByVal b As Byte)
    Static pa(0 To 0) As PointerAccessor
    With pa(0)
        If .sa.cDims = 0 Then
            .sa.cDims = 1
            .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
            .sa.cbElements = 1
            .sa.cLocks = 1
            WritePtrNatively pa, VarPtr(.sa)
        End If
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = CLng(b)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Sub

Private Sub WriteThunkDword(ByVal addr As LongPtr, ByVal v As LongLong)
    WriteNativeByte addr, CByte(v And &HFF)
    WriteNativeByte addr + 1, CByte((v \ &H100&) And &HFF)
    WriteNativeByte addr + 2, CByte((v \ &H10000) And &HFF)
    WriteNativeByte addr + 3, CByte((v \ &H1000000) And &HFF)
End Sub

#If Win64 Then
Private Sub EmitDword(ByRef bb() As Byte, ByRef p As Long, ByVal v As Long)
    bb(p) = CByte(v And &HFF)
    bb(p + 1) = CByte((v \ &H100&) And &HFF)
    bb(p + 2) = CByte((v \ &H10000) And &HFF)
    bb(p + 3) = CByte((v \ &H1000000) And &HFF)
    p = p + 4
End Sub

Private Sub EmitBytes(ByRef bb() As Byte, ByRef p As Long, ParamArray values() As Variant)
    Dim v As Variant
    For Each v In values
        bb(p) = CByte(v)
        p = p + 1
    Next v
End Sub

Private Sub EmitPtr(ByRef bb() As Byte, ByRef p As Long, ByVal ptrValue As LongPtr)
    Dim i As Long
    For i = 0 To 7
        bb(p + i) = CByte((ptrValue \ (2 ^ (8 * i))) And &HFF)
    Next i
    p = p + 8
End Sub
#End If

Private Sub EnsureHookMaps()
    If g_targetByHwnd Is Nothing Then Set g_targetByHwnd = CreateObject("Scripting.Dictionary")
End Sub

Private Function HwndKey(ByVal hWnd As LongPtr) As String
    HwndKey = CStr(hWnd)
End Function
