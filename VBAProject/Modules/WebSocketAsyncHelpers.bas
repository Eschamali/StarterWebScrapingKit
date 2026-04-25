Attribute VB_Name = "WebSocketAsyncHelpers"
'***************************************************************************************************
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
    Dim targetObjPtr As LongPtr: targetObjPtr = ObjPtr(Target)
    pa.sa.pvData = pa.arr(0) + PtrSize * 7
    Dim tProcPtr As LongPtr: tProcPtr = pa.arr(0)   'WebSocketHTTPCommunicator.WinHttpCallbackProc

    'DummyASM アドレスを戻り値にセット
    aPtr = VBA.Int(AddressOf DummyASM)
    GetWinHttpCallbackProc = aPtr
    pa.sa.pvData = aPtr

#If x64 Then
    '--- x64 マシンコード（診断最小版: 34バイト）---
    ' this/第1引数/第2引数のみを厳密に詰め替え、残り2引数は 0 で呼ぶ。
    ' まず CallbackHitCount が増えることを優先して確認する。
    '4C 8B D1           MOV R10, RCX            ; hInternet 退避
    pa.arr(0) = &HD18B4C
    '49 89 D3           MOV R11, RDX            ; dwContext 退避（未使用だがレジスタ保存）
    pa.sa.pvData = aPtr + 3: pa.arr(0) = &HD38949
    '48 B9 <imm64>      MOV RCX, targetObjPtr   ; this
    pa.sa.pvData = aPtr + 6: pa.arr(0) = &HB948
    pa.sa.pvData = aPtr + 8: pa.arr(0) = targetObjPtr
    '4C 89 D2           MOV RDX, R10            ; arg1 = hInternet
    pa.sa.pvData = aPtr + 16: pa.arr(0) = &HD2894C
    '4D 89 C0           MOV R8, R8              ; arg2 = dwInternetStatus（NOP的に明示）
    pa.sa.pvData = aPtr + 19: pa.arr(0) = &HC0894D
    '4D 31 C9           XOR R9, R9              ; arg3 = 0
    pa.sa.pvData = aPtr + 22: pa.arr(0) = &HC9314D
    '48 83 EC 28        SUB RSP, 28h
    pa.sa.pvData = aPtr + 25: pa.arr(0) = &H28EC8348
    '48 B8 <imm64>      MOV RAX, tProcPtr
    pa.sa.pvData = aPtr + 29: pa.arr(0) = &HB848
    pa.sa.pvData = aPtr + 31: pa.arr(0) = tProcPtr
    'FF D0              CALL RAX
    pa.sa.pvData = aPtr + 39: pa.arr(0) = &HD0FF&
    '48 83 C4 28        ADD RSP, 28h
    pa.sa.pvData = aPtr + 41: pa.arr(0) = &H28C48348
    'C3                 RET
    pa.sa.pvData = aPtr + 45: pa.arr(0) = &HC3&
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

'* 機能    ：型安全にポインタを書き込みます（LibTimers.bas::WritePtrNatively と同一）
'https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
