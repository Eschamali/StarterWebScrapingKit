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
'* 返り値  ：コールバック関数ポインタ（EntryPoint アドレス）
'* 引数    ：target   コールバックを受け取る WebSocketHTTPCommunicator インスタンス
'---------------------------------------------------------------------------------------------------
'* 仕組み  ：SafeTimer の GetTimerProc と同じ原理。
'            DummyASM 本体領域にマシンコードを直接書き込み、
'            WinHttp から渡される dwContext（= ObjPtr(target)）を RCX/[ESP+04] に置き換えて
'            target.WinHttpCallbackProc（vtable[7]）を呼び出す。
'
'* 注意事項：VBA が Break モード中はコールバック呼び出しをスキップします（EBMode チェック）。
'            target.WinHttpCallbackProc は vtable の先頭ユーザーメソッド（7番目、0-indexed）
'            として固定配置されている必要があります。
'***************************************************************************************************
Public Function GetWinHttpCallbackProc(ByVal target As WebSocketHTTPCommunicator) As LongPtr
    If target Is Nothing Then Exit Function
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
    pa.sa.pvData = ObjPtr(target)
    pa.sa.rgsabound0.cElements = 1

#If x32 Then
    '--- x32: vtable[7] = WinHttpCallbackProc のアドレスを事前取得 ---
    ' (SafeTimer は vtable[8] = TimerProc だったが、本実装は vtable[7] = 先頭ユーザーメソッド)
    pa.sa.pvData = pa.arr(0) + PtrSize * 7
    Dim tProcPtr As Long: tProcPtr = pa.arr(0)   'WebSocketHTTPCommunicator.WinHttpCallbackProc
#End If

    'EntryPoint アドレスを戻り値にセット（EBMode チェック付きトランポリン）
    GetWinHttpCallbackProc = VBA.Int(AddressOf EntryPoint)
    aPtr = VBA.Int(AddressOf DummyASM)
    pa.sa.pvData = aPtr

#If x64 Then
    '--- x64 マシンコード（計34バイト）---
    ' WinHttp コールバックシグネチャ（stdcall 5引数）:
    '   RCX=hInternet, RDX=dwContext(=ObjPtr), R8=dwInternetStatus,
    '   R9=lpvStatusInformation, [RSP+28h]=dwStatusInformationLength
    '
    ' 目標：RDX(ObjPtr) → RCX(this)、残り引数をシフト、vtable[7] を呼ぶ
    '
    ' 注意：[RSP+28h] の読み取りは PUSH/SUB の前に行う（スタックシフト後はオフセットがずれる）
    '
    If (pa.arr(0) And &HFFFFFF) <> &HD18948 Then
        '48 89 D1          MOV RCX, RDX          ; dwContext(ObjPtr) → this
        pa.arr(0) = &HD18948
        '4C 89 C2          MOV RDX, R8           ; dwInternetStatus をシフト
        pa.sa.pvData = aPtr + 3:  pa.arr(0) = &HC2894C
        '4D 89 C8          MOV R8, R9            ; lpvStatusInformation をシフト
        pa.sa.pvData = aPtr + 6:  pa.arr(0) = &HC8894D
        '4C 8B 4C 24       MOV R9, [RSP+28h] の前半4バイト
        pa.sa.pvData = aPtr + 9:  pa.arr(0) = &H244C8B4C
        '28                MOV R9, [RSP+28h] の最終バイト（+ 次命令の開始）
        pa.sa.pvData = aPtr + 13: pa.arr(0) = &H18B4828    '28 48 8B 01
        '                  ↑ 28=MOV R9[RSP+28h]終端, 48 8B 01=MOV RAX,[RCX]の開始
        pa.sa.pvData = aPtr + 17: pa.arr(0) = &H55&        '55       PUSH RBP
        pa.sa.pvData = aPtr + 18: pa.arr(0) = &HEC8B48     '48 8B EC MOV RBP,RSP
        pa.sa.pvData = aPtr + 21: pa.arr(0) = &H28EC8348   '48 83 EC 28  SUB RSP,0x28
        'FF 50 38          CALL [RAX+0x38]       ; vtable[7]=WinHttpCallbackProc (7×8=0x38)
        pa.sa.pvData = aPtr + 25: pa.arr(0) = &H3850FF
        pa.sa.pvData = aPtr + 28: pa.arr(0) = &H28C48348   '48 83 C4 28  ADD RSP,0x28
        pa.sa.pvData = aPtr + 32: pa.arr(0) = &H5D&        '5D       POP RBP
        pa.sa.pvData = aPtr + 33: pa.arr(0) = &HC3&        'C3       RET
    End If
    'DummyASM アドレスを EntryPoint+55 の位置に書き込む（EBMode トランポリン完成）
    pa.sa.pvData = GetWinHttpCallbackProc + 55

#Else
    '--- x32 マシンコード（計20バイト）---
    ' WinHttp コールバックシグネチャ（stdcall 5引数）:
    '   [ESP+04]=hInternet, [ESP+08]=dwContext(=ObjPtr), [ESP+0C]=dwInternetStatus,
    '   [ESP+10]=lpvStatusInformation, [ESP+14]=dwStatusInformationLength
    '
    ' 目標：[ESP+08](ObjPtr) を [ESP+04] に上書き → this として WinHttpCallbackProc へジャンプ
    '
    If pa.arr(0) <> &H824448B Then
        '8B 44 24 08       MOV EAX, [ESP+08]     ; dwContext(ObjPtr) 取得
        pa.arr(0) = &H824448B
        '89 44 24 04       MOV [ESP+04], EAX     ; hInternet 位置に上書き（this として渡す）
        pa.sa.pvData = aPtr + 4:  pa.arr(0) = &H4244489
        'B8                MOV EAX, ...          ; WinHttpCallbackProc アドレスを即値ロード
        pa.sa.pvData = aPtr + 8:  pa.arr(0) = &HB8&
        '                  (WinHttpCallbackProc の絶対アドレス)
        pa.sa.pvData = aPtr + 9:  pa.arr(0) = tProcPtr
        'FF E0             JMP EAX
        pa.sa.pvData = aPtr + 13: pa.arr(0) = &HE0FF&
        'C2 14 00          RET 0x14              ; 5引数 × 4byte = 0x14（SetTimer は 0x10）
        pa.sa.pvData = aPtr + 15: pa.arr(0) = &H14C2&
    End If
    'DummyASM アドレスを EntryPoint+22 の位置に書き込む（EBMode トランポリン完成）
    pa.sa.pvData = GetWinHttpCallbackProc + 22

#End If

    'DummyASM の先頭アドレスを書き込んでトランポリンを完成させる
    pa.arr(0) = aPtr
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
