Attribute VB_Name = "WebSocketAsyncHelpers"
'***************************************************************************************************
'   WebSocketの非同期コールバックを円滑に処理するためのヘルパー関数群。
'
'   WinHttpSetStatusCallback に渡すコールバック関数ポインタを、
'   VBA-SafeTimer（https://github.com/cristianbuse/VBA-SafeTimer）と同じ手法で生成します。
'
'   SafeTimer との対応：
'       LibTimers.bas::GetTimerProc        → GetWinHttpCallbackProc
'       LibTimers.bas::EntryPoint          → EntryPoint  （EBMode チェック用、空 Sub 固定）
'       LibTimers.bas::DummyASM            → DummyASM    （マシンコード書き込み先、空 Sub 固定）
'
'   SetTimer との違い：
'       ・インスタンス識別引数が x64=R8(3番目) → x64=RDX(2番目) に変わる
'       ・引数が 4個 → 5個 に増えるため、x32 は RETN 0x10 → RETN 0x14 に変わる
'       ・vtable オフセットが 0x40 → 0x38 に変わる（WinHttpCallbackProc がクラス先頭ユーザーメソッド）
'***************************************************************************************************
Option Explicit

Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal msg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Const GWLP_WNDPROC As Long = -4
Public Const WM_APP As Long = &H8000&
Public Const WM_APP_WINHTTP_CALLBACK As Long = WM_APP + 711

Private g_msgCount As Long
Private g_prevWndProcByHwnd As Dictionary 'Scripting.Dictionary: key=Str(hWnd), value=prevWndProc(LongPtr)
Private g_targetByHwnd As Dictionary      'Scripting.Dictionary: key=Str(hWnd), value=WebSocketHTTPCommunicator


#If VBA7 = 0 Then
    Public Enum LongPtr: [_]: End Enum
    Private Enum LONG_PTR: [_]: End Enum
#End If

#Const x64 = Win64
#Const x32 = (x64 = 0)

#If x64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    Private Const NullPtr As Long = 0&
    Private Const PtrSize = 4
#End If

'--- SAFEARRAY 操作用型定義（LibTimers.bas と同一） ---
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

'--- マシンコード注入に使う空 Sub（絶対に移動・変更しないこと） ---
Private Sub EntryPoint(): End Sub   ' EBMode チェックのエントリポイント（VBA が管理）
Private Sub DummyASM():   End Sub   ' マシンコード書き込み先

'--- WndProc用の空 Function（シグネチャを4引数に合わせる） ---
Private Function WndProcEntryPoint(ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr: End Function
Private Function WndProcDummyASM(ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr: End Function
'***************************************************************************************************
'* 機能    ：WinHttpSetStatusCallback に渡せる安全なコールバック関数ポインタを生成します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：コールバック関数ポインタ（DummyASM アドレス）
'* 引数    ：target   コールバックを受け取る WebSocketHTTPCommunicator インスタンス
'---------------------------------------------------------------------------------------------------
'* 仕組み  ：SafeTimer の GetTimerProc と同じ原理。
'            DummyASM 本体領域にマシンコードを直接書き込み、
'            WinHttp から渡される dwContext（= ObjPtr(target)）を RCX/[ESP+04] に置き換えて
'            target.WinHttpCallbackProc（vtable[7]）を呼び出す。
'
'* 注意事項：target.WinHttpCallbackProc は vtable の先頭ユーザーメソッド（7番目、0-indexed）
'            として固定配置されている必要があります。
'***************************************************************************************************
Public Function GetWinHttpCallbackProc(ByVal Target As WebSocketHTTPCommunicator) As LongPtr
    If Target Is Nothing Then Exit Function
    Static pa As PointerAccessor
    Dim aPtr As LongPtr

    'SAFEARRAY を初回のみ初期化
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    'ObjPtr(target) → vtable ポインタを読む
    pa.sa.pvData = ObjPtr(Target)
    pa.sa.rgsabound0.cElements = 1
    pa.sa.pvData = pa.arr(0) + PtrSize * 7
    Dim tProcPtr As LongPtr: tProcPtr = pa.arr(0)   'WebSocketHTTPCommunicator.WinHttpCallbackProc
    Dim postMessageProc As LongPtr: postMessageProc = GetPostMessageWProc()

    'SafeTimer と同様に EntryPoint を返し、Break mode は VBA 側の EBMode 判定に任せる
    GetWinHttpCallbackProc = VBA.Int(AddressOf EntryPoint)
    aPtr = VBA.Int(AddressOf DummyASM)
    pa.sa.pvData = aPtr

#If x64 Then
    '--- x64 マシンコード（57バイト）---
    ' PostMessageW(dwContext(=notifyHwnd), WM_APP_WINHTTP_CALLBACK, packed(status+len), packed(wsStatus))
    If postMessageProc = 0 Then Exit Function
    '48 89 D1           MOV RCX, RDX            ; hwnd = dwContext
    pa.arr(0) = &HD18948
    '48 C7 C2 <imm32>   MOV RDX, WM_APP_WINHTTP_CALLBACK
    pa.sa.pvData = aPtr + 3: pa.arr(0) = &HC2C748
    pa.sa.pvData = aPtr + 6: pa.arr(0) = WM_APP_WINHTTP_CALLBACK
    'wParam に dwInternetStatus(下位32bit) + dwStatusInformationLength(上位32bit) を詰める
    '4C 8B 54 24 28     MOV R10, [RSP+28h]      ; length
    pa.sa.pvData = aPtr + 10: pa.arr(0) = &H24548B4C
    pa.sa.pvData = aPtr + 14: pa.arr(0) = &H28&
    '49 C1 E2 20        SHL R10, 32
    pa.sa.pvData = aPtr + 15: pa.arr(0) = &H20E2C149
    '45 8B D8           MOV R11D, R8D           ; status
    pa.sa.pvData = aPtr + 19: pa.arr(0) = &HD88B45
    '4D 0B D3           OR R10, R11
    pa.sa.pvData = aPtr + 22: pa.arr(0) = &HD30B4D
    '4D 89 D0           MOV R8, R10             ; wParam
    pa.sa.pvData = aPtr + 25: pa.arr(0) = &HD0894D
    'R9 を packed(WS_STATUS) に変換（ポインタ直接渡しを避ける）
    '4D 85 C9           TEST R9, R9
    pa.sa.pvData = aPtr + 28: pa.arr(0) = &HC9854D
    '74 03              JE +3
    pa.sa.pvData = aPtr + 31: pa.arr(0) = &H374&
    '4D 8B 09           MOV R9, [R9]
    pa.sa.pvData = aPtr + 33: pa.arr(0) = &H98B4D
    '48 83 EC 28        SUB RSP, 28h
    pa.sa.pvData = aPtr + 36: pa.arr(0) = &H28EC8348
    '48 B8 <imm64>      MOV RAX, postMessageProc
    pa.sa.pvData = aPtr + 40: pa.arr(0) = &HB848
    pa.sa.pvData = aPtr + 42: pa.arr(0) = postMessageProc
    'FF D0              CALL RAX
    pa.sa.pvData = aPtr + 50: pa.arr(0) = &HD0FF&
    '48 83 C4 28        ADD RSP, 28h
    pa.sa.pvData = aPtr + 52: pa.arr(0) = &H28C48348
    'C3                 RET
    pa.sa.pvData = aPtr + 56: pa.arr(0) = &HC3&
    'EntryPoint 内の call target を DummyASM へ差し替える（Break mode は call をスキップ）
    pa.sa.pvData = GetWinHttpCallbackProc + 55
    pa.arr(0) = aPtr
#Else
    '--- x32 マシンコード（計20バイト）---
    ' WinHttp コールバックシグネチャ（stdcall 5引数）:
    '   [ESP+04]=hInternet, [ESP+08]=dwContext(=ObjPtr), [ESP+0C]=dwInternetStatus,
    '   [ESP+10]=lpvStatusInformation, [ESP+14]=dwStatusInformationLength
    '
    ' 目標：[ESP+08](ObjPtr) を [ESP+04] に上書き → this として WinHttpCallbackProc へジャンプ
    '
    '8B 44 24 08       MOV EAX, [ESP+08]     ; dwContext(ObjPtr) 取得
    pa.arr(0) = &H824448B
    '89 44 24 04       MOV [ESP+04], EAX     ; hInternet 位置に上書き（this として渡す）
    pa.sa.pvData = aPtr + 4:  pa.arr(0) = &H4244489
    'B8 + imm32        MOV EAX, tProcPtr
    pa.sa.pvData = aPtr + 8:  pa.arr(0) = &HB8&
    pa.sa.pvData = aPtr + 9:  pa.arr(0) = tProcPtr
    'FF E0             JMP EAX
    pa.sa.pvData = aPtr + 13: pa.arr(0) = &HE0FF&
    'C2 14 00          RET 0x14              ; 5引数 × 4byte = 0x14（SetTimer は 0x10）
    pa.sa.pvData = aPtr + 15: pa.arr(0) = &H14C2&
    'EntryPoint 内の call target を DummyASM へ差し替える（Break mode は call をスキップ）
    pa.sa.pvData = GetWinHttpCallbackProc + 22
    pa.arr(0) = aPtr
#End If

    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = NullPtr
End Function

Private Function GetPostMessageWProc() As LongPtr
    Dim hUser32 As LongPtr
    hUser32 = GetModuleHandleW(StrPtr("user32.dll"))
    If hUser32 = 0 Then Exit Function
    GetPostMessageWProc = GetProcAddress(hUser32, "PostMessageW")
End Function

'***************************************************************************************************
'* 機能    ：SetWindowLongPtr で安全に使える WndProc サンクを生成します
'---------------------------------------------------------------------------------------------------
'* 仕組み  ：VBAがリセットされるとEBMode=2となり、元のVBA関数へジャンプせずに安全に0を返して終了します。
'            これにより、VBEのリセットボタンを押した際のクラッシュ（0xc0000027）を防ぎます。
'***************************************************************************************************
Public Function GetSafeWndProc(ByVal targetProc As LongPtr) As LongPtr
    Static pa As PointerAccessor
    Dim aPtr As LongPtr

    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    GetSafeWndProc = VBA.Int(AddressOf WndProcEntryPoint)
    aPtr = VBA.Int(AddressOf WndProcDummyASM)
    pa.sa.pvData = aPtr
    pa.sa.rgsabound0.cElements = 1

#If x64 Then
    '--- x64: 引数 RCX, RDX, R8, R9 をそのままにして targetProc へ JMP ---
    ' 48 B8 <imm64>      MOV RAX, targetProc
    pa.arr(0) = &HB848
    pa.sa.pvData = aPtr + 2: pa.arr(0) = targetProc
    ' FF E0              JMP RAX
    pa.sa.pvData = aPtr + 10: pa.arr(0) = &HE0FF&

    ' WndProcEntryPoint 内の EBMode チェック通過後の JMP 先を上書き
    pa.sa.pvData = GetSafeWndProc + 55
    pa.arr(0) = aPtr
#Else
    '--- x32: スタック引数をそのままにして targetProc へ JMP ---
    ' B8 <imm32>         MOV EAX, targetProc
    pa.arr(0) = &HB8&
    pa.sa.pvData = aPtr + 1: pa.arr(0) = targetProc
    ' FF E0              JMP EAX
    pa.sa.pvData = aPtr + 5: pa.arr(0) = &HE0FF&

    ' WndProcEntryPoint 内の EBMode チェック通過後の JMP 先を上書き
    pa.sa.pvData = GetSafeWndProc + 22
    pa.arr(0) = aPtr
#End If

    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = NullPtr
End Function


'***************************************************************************************************
'                                   ■■■ 内部ヘルパー ■■■
'***************************************************************************************************
'* 機能    ：指定アドレスに LongPtr 値を書き込みます（LibTimers.bas::MemLongPtr と同一）
'***************************************************************************************************
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

'***************************************************************************************************
'* 機能    ：指定アドレスから LongPtr 値を読み取ります（CopyMemory 不使用）
'***************************************************************************************************
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

'* 機能    ：型安全にポインタを書き込みます（LibTimers.bas::WritePtrNatively と同一）
'https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub


Public Sub InstallWinHttpMessageHook(ByVal targetHwnd As LongPtr, ByVal Target As WebSocketHTTPCommunicator)
    If targetHwnd = 0 Then Exit Sub
    If Target Is Nothing Then Exit Sub

    EnsureHookMaps
    Dim Key As String
    Key = HwndKey(targetHwnd)

    If Not g_prevWndProcByHwnd.Exists(Key) Then
        Dim safeWndProc As LongPtr
        safeWndProc = GetSafeWndProc(AddressOf WinHttpBridgeWndProc)
        g_prevWndProcByHwnd.Add Key, SetWindowLongPtr(targetHwnd, GWLP_WNDPROC, safeWndProc)
    End If
    Set g_targetByHwnd(Key) = Target
End Sub

Public Sub RemoveWinHttpMessageHook(Optional ByVal targetHwnd As LongPtr)
    On Error Resume Next
    EnsureHookMaps
    Dim Key As String
    Dim k As Variant

    If targetHwnd = 0 Then
        For Each k In g_prevWndProcByHwnd.keys
            If CLngPtr(k) <> 0 And g_prevWndProcByHwnd(k) <> 0 Then
                SetWindowLongPtr CLngPtr(k), GWLP_WNDPROC, g_prevWndProcByHwnd(k)
            End If
        Next k
        g_prevWndProcByHwnd.RemoveAll
        g_targetByHwnd.RemoveAll
        On Error GoTo 0
        Exit Sub
    End If

    Key = HwndKey(targetHwnd)
    If g_prevWndProcByHwnd.Exists(Key) Then
        If targetHwnd <> 0 And g_prevWndProcByHwnd(Key) <> 0 Then
            SetWindowLongPtr targetHwnd, GWLP_WNDPROC, g_prevWndProcByHwnd(Key)
        End If
        g_prevWndProcByHwnd.Remove Key
    End If
    If g_targetByHwnd.Exists(Key) Then
        g_targetByHwnd.Remove Key
    End If
    On Error GoTo 0
End Sub

Public Function GetWinHttpMessageCount() As Long
    GetWinHttpMessageCount = g_msgCount
End Function

Private Function WinHttpBridgeWndProc(ByVal hWnd As LongPtr, ByVal msg As Long, _
                                      ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    EnsureHookMaps
    Dim Key As String
    Key = HwndKey(hWnd)

    If msg = WM_APP_WINHTTP_CALLBACK Then
        g_msgCount = g_msgCount + 1
        If g_targetByHwnd.Exists(Key) Then
            If Not (g_targetByHwnd(Key) Is Nothing) Then
                g_targetByHwnd(Key).HandlePostedWinHttpCallback wParam, lParam
            End If
        End If
        WinHttpBridgeWndProc = 0
        Exit Function
    End If

    If g_prevWndProcByHwnd.Exists(Key) Then
        WinHttpBridgeWndProc = CallWindowProc(g_prevWndProcByHwnd(Key), hWnd, msg, wParam, lParam)
    Else
        WinHttpBridgeWndProc = 0
    End If
End Function

Private Sub EnsureHookMaps()
    If g_prevWndProcByHwnd Is Nothing Then Set g_prevWndProcByHwnd = CreateObject("Scripting.Dictionary")
    If g_targetByHwnd Is Nothing Then Set g_targetByHwnd = CreateObject("Scripting.Dictionary")
End Sub

Private Function HwndKey(ByVal hWnd As LongPtr) As String
    HwndKey = CStr(hWnd)
End Function
