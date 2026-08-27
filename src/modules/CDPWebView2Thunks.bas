Attribute VB_Name = "CDPWebView2Thunks"
'***************************************************************************************************
'   WebView2用の機械語サンク・vtable呼び出し・SAFEARRAYメモリプリミティブを担う基盤モジュールです。
'
'   出典・移植元：WebView2-For-Excel-VBA プロジェクトの `Wv2Thunks.bas`(第9.16段階)。
'   このモジュールの心臓部(PointerAccessor/SAFEARRAYメモリプリミティブ、機械語サンク生成、
'   スロット管理、Handler_QueryInterface/AddRef/Release、センチネル機構)は、実機検証済みの
'   ロジックをそのまま移植したものであり、バイト列やオフセット値は一切変更していません。
'
'   このプロジェクト向けに追加/変更した点：
'     ・`HandlerKind`を、CDP用の4種類(HK_EnvironmentCompleted/HK_ControllerCompleted/
'       HK_CdpMethodCompleted/HK_CdpEventReceived)に絞った
'     ・`InitIIDTable`に、CallDevToolsProtocolMethodCompletedHandler と
'       DevToolsProtocolEventReceivedEventHandler の実IIDを追加した
'     ・`EnsureWebView2LoaderResolved`(WebView2Loader.dll探索ヘルパー)を新設した
'         → StarterWebScrapingKitのCLAUDE.mdは「外部バイナリの配置」を禁止しているため、
'           このプロジェクト専用のWebView2Loader.dllは同梱しない。代わりに、Excelの
'           Power Query統合アドインに同梱されている実物を実行時に探索してLoadLibraryする。
'     ・UserFormマウスリサイズ用API(GetClientRect/GetAncestor/SetWindowLongPtrW/
'       SetWindowPos)や、この用途で使わないTest_系Subは移植対象外とした
'     ・初期表示タブは犠牲にし、`newtab`からスタートすることで、UserFormなしで一応、可視化状態で制御可。※タブ化はしない
'
'   ★重要(既知の落とし穴、継承不可避)★
'     ・全てのCOMコールバックはこのモジュールの機械語サンクを経由する。VBEでブレーク/
'       ステップ実行するタイミングによっては、Excelがクラッシュする可能性がある
'       (コールバック待ち中はブレークしないこと)
'     ・`WritePtrNatively`の引数は`LONG_PTR`(`LongPtr`ではない)。実行時エラーの原因になる
'***************************************************************************************************
Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "このモジュールは64ビットVBA(x64)が必要です"
#End If



'***************************************************************************************************
'                                   ■■■ SafeArray / PointerAccessor ■■■
'***************************************************************************************************
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/refTypes/discussions/3
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
' --- kernel32: メモリ / モジュール ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long

Private Declare PtrSafe Function VirtualQuery Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByRef lpBuffer As MEMORY_BASIC_INFORMATION, _
    ByVal dwLength As LongPtr) As LongPtr

' --- oleaut32: DispCallFunc ---
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Any, _
    ByRef prgpvarg As Any, _
    ByRef pvargResult As Any) As Long

' --- 文字列ヘルパー用API ---
Private Declare PtrSafe Function lstrlenW Lib "kernel32" ( _
    ByVal lpString As LongPtr) As Long

Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" ( _
    ByVal pv As LongPtr)

Private Declare PtrSafe Function lstrcpyW Lib "kernel32" ( _
    ByVal lpString1 As LongPtr, _
    ByVal lpString2 As LongPtr) As LongPtr

' --- 環境変数API(センチネル機構用) ---
Private Declare PtrSafe Function GetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, _
    ByVal lpBuffer As LongPtr, _
    ByVal nSize As Long) As Long

Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, _
    ByVal lpValue As LongPtr) As Long

' --- RECT構造体(ICoreWebView2Controller::put_Boundsに渡す) ---
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' --- VirtualQuery用構造体(センチネルの健在判定用) ---
Private Type MEMORY_BASIC_INFORMATION
    BaseAddress As LongPtr
    AllocationBase As LongPtr
    AllocationProtect As Long
    pad1 As Long
    RegionSize As LongPtr
    state As Long
    Protect As Long
    Type_ As Long
    pad2 As Long
End Type



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const MEM_COMMIT             As Long = &H1000&
Private Const MEM_RESERVE            As Long = &H2000&
Private Const MEM_RELEASE            As Long = &H8000&    ' サフィックス必須(Integer誤判定の罠回避)
Private Const MEM_FREE               As Long = &H10000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&
Private Const PAGE_READWRITE         As Long = &H4&
Private Const S_OK                   As Long = 0
Private Const E_NOINTERFACE          As Long = &H80004002
Private Const CC_STDCALL             As Long = 4

' --- センチネル機構用定数 ---
Private Const SENTINEL_ENV_NAME    As String = "CDPWV2_VBA_LastRegion"
Private Const SENTINEL_BUFFER_SIZE As Long = 32

' --- スロット / 領域レイアウト定数(移植元のオフセット値をそのまま使用) ---
Private Const LATE_BIND_OFFSET As Long = 55
Private Const STUB_LEN         As Long = 91
Private Const THUNK_LEN        As Long = 74
Private Const THUNK_OFFSET     As Long = 96
Public Const VTABLE_OBJ_OFFSET As Long = 176
Private Const SLOT_SIZE        As Long = 224
Private Const SLOT_COUNT       As Long = 512
Private Const HEADER_SIZE      As Long = 64
Private Const REGION_SIZE      As Long = HEADER_SIZE + SLOT_SIZE * SLOT_COUNT
Private Const THUNK_BUF_SIZE   As Long = 80



'***************************************************************************************************
'                                   ■■■ GUID型 / HandlerKind ■■■
'***************************************************************************************************
Public Type GUID
    data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Enum HandlerKind
    HK_None = 0
    HK_EnvironmentCompleted = 1
    HK_ControllerCompleted = 2
    HK_CdpMethodCompleted = 3    ' ICoreWebView2CallDevToolsProtocolMethodCompletedHandler(通常版/ForSession版で共用)
    HK_CdpEventReceived = 4      ' ICoreWebView2DevToolsProtocolEventReceivedEventHandler(永続)
End Enum



'***************************************************************************************************
'                                   ■■■ 各種モジュール変数 ■■■
'***************************************************************************************************
Private m_pHandler_QI       As LongPtr
Private m_pHandler_AddRef   As LongPtr
Private m_pHandler_Release  As LongPtr

Private m_pRegionBase As LongPtr
Private m_freeHead    As Long
Private m_freeNext()  As Long
Private m_inUse       As Long

Private m_handlers(0 To SLOT_COUNT - 1) As CDPWebView2CallbackHandler
Private m_iidTable(HK_None To HK_CdpEventReceived) As GUID
Private m_iidIUnknown As GUID

Private m_loaderModule As LongPtr   ' EnsureWebView2LoaderResolvedが解決したHMODULE。0なら未解決

' EntryPointスタブのソース(空Sub。AddressOfでVBAランタイム生成のトランポリンを取得するために存在)
Private Sub EntryPoint(): End Sub



'***************************************************************************************************
'                              ■■■ AcquireHandlerFor ■■■
'***************************************************************************************************
'* 機能　　：新規のCOMコールバックハンドラを1個確保します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：初期化済みの`CDPWebView2CallbackHandler`(失敗時`Nothing`)
'* 引数　　：kind    このハンドラの種別
'            owner   コールバック受信先オブジェクト(`Public`メソッドが動的束縛で呼ばれる)
'***************************************************************************************************
Public Function AcquireHandlerFor( _
    ByVal kind As HandlerKind, _
    ByVal owner As Object) As CDPWebView2CallbackHandler

    If m_pRegionBase = 0 Then
        If Not Thunks_Init() Then Exit Function
    End If

    Dim h As CDPWebView2CallbackHandler
    Set h = New CDPWebView2CallbackHandler

    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(h, 7)
    If pHandlerInvoke = 0 Then Exit Function

    Dim pSlot As LongPtr
    pSlot = Thunks_AcquireSlot(h, ObjPtr(h), pHandlerInvoke)
    If pSlot = 0 Then Exit Function

    h.Init kind, owner, pSlot

    Set AcquireHandlerFor = h
End Function



'***************************************************************************************************
'                              ■■■ dcf(汎用vtable呼び出し) ■■■
'***************************************************************************************************
'* 機能　　：DispCallFunc経由で、任意のvtableスロットのCOMメソッドを呼びます
'---------------------------------------------------------------------------------------------------
'* 引数　　：pInterface  COMインターフェースポインタ(this)
'            vtblIndex   vtableのスロット番号(IUnknown 0=QI/1=AddRef/2=Release、以降は宣言順)
'            funcName    デバッグ用(空文字なら出力しない)
'            args        メソッドに渡す可変長引数
'* 返り値  ：呼び出し結果(HRESULTやAddRef/Releaseの戻り値をそのまま返す)
'***************************************************************************************************
Public Function dcf( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal funcName As String, _
    ParamArray args() As Variant) As Long

    If pInterface = 0 Then
        Debug.Print "dcf: null interface - " & funcName
        dcf = &H80004003   ' E_POINTER
        Exit Function
    End If

    Dim argc As Long
    argc = UBound(args) - LBound(args) + 1
    If argc < 0 Then argc = 0

    Dim res As Variant
    Dim hr As Long

    If argc = 0 Then
        hr = DispCallFunc(pInterface, vtblIndex * PtrSize, _
                          CC_STDCALL, vbLong, _
                          0, ByVal 0&, ByVal 0&, res)
    Else
        Dim vt() As Integer
        Dim vp() As LongPtr
        Dim vals() As Variant
        ReDim vt(0 To argc - 1)
        ReDim vp(0 To argc - 1)
        ReDim vals(0 To argc - 1)

        Dim i As Long
        For i = 0 To argc - 1
            vals(i) = args(LBound(args) + i)
            Select Case VarType(vals(i))
                Case vbLong:     vt(i) = vbLong
                Case vbLongLong: vt(i) = vbLongLong
                Case vbDouble:   vt(i) = vbDouble
                Case Else:       vt(i) = vbLongLong
            End Select
            vp(i) = VarPtr(vals(i))
        Next i

        hr = DispCallFunc(pInterface, vtblIndex * PtrSize, _
                          CC_STDCALL, vbLong, _
                          argc, vt(0), vp(0), res)
    End If

    If hr <> 0 Then
        If LenB(funcName) > 0 Then _
            Debug.Print "dcf CALL failed: " & funcName & " hr=&H" & Hex(hr)
        dcf = hr
    Else
        dcf = CLng(res)
    End If
End Function

Public Function ComRelease(ByVal pInterface As LongPtr) As Long
    If pInterface <> 0 Then ComRelease = dcf(pInterface, 2, "Release")
End Function

Public Function ComAddRef(ByVal pInterface As LongPtr) As Long
    If pInterface <> 0 Then ComAddRef = dcf(pInterface, 1, "AddRef")
End Function



'***************************************************************************************************
'                              ■■■ 文字列/プロパティヘルパー ■■■
'***************************************************************************************************
'* 機能　　：LPWSTR(UTF-16終端文字列ポインタ)をVBAのStringに変換します
'---------------------------------------------------------------------------------------------------
'* 注意事項：入力ポインタはCoTaskMemAllocで確保されている前提。呼び出し側が
'            変換後に`CoTaskMemFree`で解放する責任を持つ(`GetStringProperty`は一括で行う)
'***************************************************************************************************
Public Function PtrToString(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    Dim cch As Long
    cch = lstrlenW(p)
    If cch = 0 Then Exit Function
    PtrToString = String$(cch, vbNullChar)
    lstrcpyW StrPtr(PtrToString), p
End Function

'* 機能　　：`HRESULT get_Xxx([out,retval] LPWSTR *value)`形のCOMメソッドを呼び、Stringで返します
Public Function GetStringProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As String

    If pInterface = 0 Then Exit Function

    Dim pStr As LongPtr
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(pStr))
    If hr = 0 And pStr <> 0 Then
        GetStringProperty = PtrToString(pStr)
        CoTaskMemFree pStr
    End If
End Function



'***************************************************************************************************
'                              ■■■ Thunks_Init / AcquireSlot / ReleaseSlot / Shutdown ■■■
'***************************************************************************************************
'* 機能　　：サンク領域をVirtualAllocで確保し、全スロットへスタブをコピーし、
'            フリーリストを初期化します(初回のみ実行)
'***************************************************************************************************
Public Function Thunks_Init() As Boolean

    If m_pRegionBase <> 0 Then
        Thunks_Init = True
        Exit Function
    End If

    ' 前回リセットで回収漏れした領域があれば回収する
    Sentinel_RecoverIfNeeded

    ' EntryPointスタブのソースを実体化
    EntryPoint

    Dim pStubSrc As LongPtr
    pStubSrc = VBA.Int(AddressOf EntryPoint)
    If pStubSrc = 0 Then Exit Function

    m_pHandler_QI = GetAddr(AddressOf Handler_QueryInterface)
    m_pHandler_AddRef = GetAddr(AddressOf Handler_AddRef)
    m_pHandler_Release = GetAddr(AddressOf Handler_Release)
    If m_pHandler_QI = 0 Or m_pHandler_AddRef = 0 Or m_pHandler_Release = 0 Then
        Exit Function
    End If

    m_pRegionBase = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If m_pRegionBase = 0 Then Exit Function

    Sentinel_StorePrevRegion m_pRegionBase

    Dim k As Long
    For k = 0 To HEADER_SIZE - 1 Step 8
        MemLongPtr(m_pRegionBase + k) = 0^
    Next k

    MemLongPtr(m_pRegionBase) = 1^  ' 生存フラグを立てる

    Dim i As Long, pSlot As LongPtr, pVTableObj As LongPtr, pFunctions As LongPtr
    For i = 0 To SLOT_COUNT - 1
        pSlot = SlotAddrAt(i)

        For k = 0 To 95 Step 8
            MemLongPtr(pSlot + k) = ReadLongPtr(pStubSrc + k)
        Next k

        pVTableObj = pSlot + VTABLE_OBJ_OFFSET
        pFunctions = pVTableObj + PtrSize

        MemLongPtr(pVTableObj) = pFunctions
        MemLongPtr(pFunctions + 0 * PtrSize) = m_pHandler_QI
        MemLongPtr(pFunctions + 1 * PtrSize) = m_pHandler_AddRef
        MemLongPtr(pFunctions + 2 * PtrSize) = m_pHandler_Release
        MemLongPtr(pFunctions + 3 * PtrSize) = pSlot
    Next i

    ReDim m_freeNext(0 To SLOT_COUNT - 1)
    For i = 0 To SLOT_COUNT - 2
        m_freeNext(i) = i + 1
    Next i
    m_freeNext(SLOT_COUNT - 1) = -1
    m_freeHead = 0
    m_inUse = 0

    For i = 0 To SLOT_COUNT - 1
        Set m_handlers(i) = Nothing
    Next i

    InitIIDTable

    Thunks_Init = True
End Function

'* 機能　　：フリーリストから空きスロットを1個取得し、サンクを書き込みます
Public Function Thunks_AcquireSlot( _
    ByVal handler As CDPWebView2CallbackHandler, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    If m_pRegionBase = 0 Then Exit Function
    If m_freeHead < 0 Then Exit Function
    If handler Is Nothing Then Exit Function

    Dim idx As Long
    idx = m_freeHead
    m_freeHead = m_freeNext(idx)
    m_freeNext(idx) = -1

    Dim pSlot As LongPtr
    pSlot = SlotAddrAt(idx)

    WriteThunkMachineCode pSlot + THUNK_OFFSET, pSelfObj, pTargetFunc, m_pRegionBase
    MemLongPtr(pSlot + LATE_BIND_OFFSET) = pSlot + THUNK_OFFSET

    Set m_handlers(idx) = handler

    m_inUse = m_inUse + 1
    Thunks_AcquireSlot = pSlot
End Function

'* 機能　　：スロットをフリーリストに返却します
Public Sub Thunks_ReleaseSlot(ByVal pSlot As LongPtr)
    If m_pRegionBase = 0 Then Exit Sub
    If pSlot = 0 Then Exit Sub

    Dim idx As Long
    idx = SlotIndexFromAddr(pSlot)
    If idx < 0 Then Exit Sub
    If m_freeNext(idx) <> -1 Then Exit Sub  ' 既に空き = 二重解放

    Set m_handlers(idx) = Nothing
    m_freeNext(idx) = m_freeHead
    m_freeHead = idx
    m_inUse = m_inUse - 1
End Sub

'* 機能　　：全スロットを解放します(全てのWebView2利用が終わった後、明示的に呼ぶこと)
'* 注意事項：スロットプールはこのモジュール内で全インスタンス共通(グローバル)なので、
'            他の`CDPCoreViaWebView2`インスタンスがまだ使用中の可能性がある間は呼ばないこと
Public Sub Thunks_Shutdown()
    If m_pRegionBase = 0 Then Exit Sub

    MemLongPtr(m_pRegionBase) = 0^

    Dim i As Long
    For i = 0 To SLOT_COUNT - 1
        If Not (m_handlers(i) Is Nothing) Then
            m_handlers(i).ClearOwner
            Set m_handlers(i) = Nothing
        End If
    Next i

    Dim shutFreeResult As Long
    shutFreeResult = VirtualFree(m_pRegionBase, 0, MEM_RELEASE)
    Debug.Print "CDPWebView2Thunks.Thunks_Shutdown: VirtualFree(" & m_pRegionBase & ") returned " & shutFreeResult

    m_pRegionBase = 0
    m_freeHead = -1
    m_inUse = 0
    Erase m_freeNext

    Sentinel_ClearPrevRegion
End Sub

Private Function SlotAddrAt(ByVal idx As Long) As LongPtr
    SlotAddrAt = m_pRegionBase + HEADER_SIZE + CLngLng(idx) * SLOT_SIZE
End Function

Private Function SlotIndexFromAddr(ByVal pSlot As LongPtr) As Long
    SlotIndexFromAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim Offset As LongPtr
    Offset = pSlot - (m_pRegionBase + HEADER_SIZE)
    If Offset < 0 Then Exit Function
    If (Offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(Offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromAddr = idx
End Function

Private Function SlotIndexFromVTableObjAddr(ByVal pVTableObj As LongPtr) As Long
    SlotIndexFromVTableObjAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim Offset As LongPtr
    Offset = pVTableObj - (m_pRegionBase + HEADER_SIZE + VTABLE_OBJ_OFFSET)
    If Offset < 0 Then Exit Function
    If (Offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(Offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromVTableObjAddr = idx
End Function



'***************************************************************************************************
'                              ■■■ サンクの機械語生成 ■■■
'***************************************************************************************************
'* 機能　　：サンクの機械語(有効長74バイト)を指定アドレスに書き込みます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：先頭18バイトで領域の生存フラグをチェックし(倒れていれば即0を返す=EBMode対策)、
'            以降56バイトで「引数を1個右にシフトしてpSelfObjを注入しpTargetFuncを呼ぶ」処理を行う
'***************************************************************************************************
Private Sub WriteThunkMachineCode( _
    ByVal addr As LongPtr, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr, _
    ByVal pRegionBase As LongPtr)

    Dim b(0 To THUNK_BUF_SIZE - 1) As Byte, i As Long
    i = 0

    ' --- 生存フラグチェック(18 bytes) ---
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2               ' mov rax, imm64 (pRegionBase)
    MemLongPtr(VarPtr(b(i))) = pRegionBase: i = i + 8
    b(i) = &H80: b(i + 1) = &H38: b(i + 2) = &H1: i = i + 3 ' cmp byte ptr [rax], 1
    b(i) = &H74: b(i + 1) = &H3: i = i + 2                 ' je +3
    b(i) = &H33: b(i + 1) = &HC0: i = i + 2                ' xor eax, eax
    b(i) = &HC3: i = i + 1                                 ' ret

    ' --- サンク本体(56 bytes) ---
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4    ' sub rsp, 0x38
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3                     ' mov r9, r8
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3                     ' mov r8, rdx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3                     ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2                                      ' mov rcx, imm64 (pSelfObj)
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5  ' lea rax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5  ' mov [rsp+0x20], rax
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2                                      ' mov rax, imm64 (pTargetFunc)
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2                                      ' call rax
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4     ' mov eax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4     ' add rsp, 0x38
    b(i) = &HC3: i = i + 1                                                       ' ret
    b(i) = &HCC: i = i + 1                                                       ' int3 (padding)
    b(i) = &HCC: i = i + 1                                                       ' int3 (padding)

    Dim k As Long
    For k = 0 To THUNK_BUF_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
End Sub

Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal Offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(Offset)))
End Function

Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function

Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function



'***************************************************************************************************
'                              ■■■ IUnknownスタブ群(標準モジュール) ■■■
'***************************************************************************************************
Private Function Handler_QueryInterface( _
    ByVal This As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long

    If riid = 0 Then
        ppvObject = 0
        Handler_QueryInterface = &H80004003   ' E_POINTER
        Exit Function
    End If

    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(This)
    If idx < 0 Then
        ppvObject = 0
        Handler_QueryInterface = E_NOINTERFACE
        Exit Function
    End If
    If m_handlers(idx) Is Nothing Then
        ppvObject = 0
        Handler_QueryInterface = E_NOINTERFACE
        Exit Function
    End If

    If IsEqualGUIDInPlace(riid, m_iidIUnknown) Then
        ppvObject = This
        HandlerAddRefInternal This
        Handler_QueryInterface = S_OK
        Exit Function
    End If

    Dim kind As HandlerKind
    kind = m_handlers(idx).kind
    If IsEqualGUIDInPlace(riid, m_iidTable(kind)) Then
        ppvObject = This
        HandlerAddRefInternal This
        Handler_QueryInterface = S_OK
        Exit Function
    End If

    Dim data1 As Long
    data1 = LongPtrLowDword(ReadLongPtr(riid))
    Debug.Print "  QI rejected [idx=" & idx & " kind=" & kind & _
                "] riid.Data1=&H" & Hex(data1) & " -> E_NOINTERFACE"
    ppvObject = 0
    Handler_QueryInterface = E_NOINTERFACE
End Function

Private Function Handler_AddRef(ByVal This As LongPtr) As Long
    Handler_AddRef = HandlerAddRefInternal(This)
End Function

Private Function Handler_Release(ByVal This As LongPtr) As Long
    Handler_Release = HandlerReleaseInternal(This)
End Function

'* 機能　　：riid(WebView2から渡されるポインタ)と、VBA側GUID変数を比較します
Private Function IsEqualGUIDInPlace(ByVal pRiid As LongPtr, ByRef refGuid As GUID) As Boolean
    Dim pRef As LongPtr
    pRef = VarPtr(refGuid)

    If ReadLongPtr(pRiid) <> ReadLongPtr(pRef) Then Exit Function
    If ReadLongPtr(pRiid + 8) <> ReadLongPtr(pRef + 8) Then Exit Function

    IsEqualGUIDInPlace = True
End Function

Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
    Dim u As LongLong
    u = CLngLng(v) And &HFFFFFFFF^
    If u > &H7FFFFFFF^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
End Function



'***************************************************************************************************
'                              ■■■ IIDテーブル ■■■
'***************************************************************************************************
'* 出典：WebView2.h(公式SDKヘッダ、`WebView2_Vtable etc\build\native\include\WebView2.h`で実測確認済み)
Private Sub InitIIDTable()
    FillGUID m_iidIUnknown, "00000000-0000-0000-C000-000000000046"

    ' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
    FillGUID m_iidTable(HK_EnvironmentCompleted), _
             "4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d"

    ' ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
    FillGUID m_iidTable(HK_ControllerCompleted), _
             "6c4819f3-c9b7-4260-8127-c9f5bde7f68c"

    ' ICoreWebView2CallDevToolsProtocolMethodCompletedHandler(通常版/ForSession版で共用)
    FillGUID m_iidTable(HK_CdpMethodCompleted), _
             "5c4889f0-5ef6-4c5a-952c-d8f1b92d0574"

    ' ICoreWebView2DevToolsProtocolEventReceivedEventHandler
    FillGUID m_iidTable(HK_CdpEventReceived), _
             "e2fda4be-5456-406c-a261-3d452138362c"
End Sub

'* 機能　　："xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"形式の文字列からGUID構造体を埋めます
Public Sub FillGUID(ByRef g As GUID, ByVal s As String)
    g.data1 = HexStrToLong(Mid$(s, 1, 8))
    g.Data2 = HexStrToInt(Mid$(s, 10, 4))
    g.Data3 = HexStrToInt(Mid$(s, 15, 4))

    g.Data4(0) = CByte("&H" & Mid$(s, 20, 2))
    g.Data4(1) = CByte("&H" & Mid$(s, 22, 2))
    g.Data4(2) = CByte("&H" & Mid$(s, 25, 2))
    g.Data4(3) = CByte("&H" & Mid$(s, 27, 2))
    g.Data4(4) = CByte("&H" & Mid$(s, 29, 2))
    g.Data4(5) = CByte("&H" & Mid$(s, 31, 2))
    g.Data4(6) = CByte("&H" & Mid$(s, 33, 2))
    g.Data4(7) = CByte("&H" & Mid$(s, 35, 2))
End Sub

Private Function HexStrToLong(ByVal s As String) As Long
    Dim v As LongLong
    v = CLngLng("&H" & s)
    If v >= &H80000000^ Then
        HexStrToLong = CLng(v - &H100000000^)
    Else
        HexStrToLong = CLng(v)
    End If
End Function

Private Function HexStrToInt(ByVal s As String) As Integer
    Dim v As Long
    v = CLng("&H" & s)
    If v >= &H8000& Then
        HexStrToInt = CInt(v - &H10000)
    Else
        HexStrToInt = CInt(v)
    End If
End Function



'***************************************************************************************************
'                              ■■■ 参照カウント管理 ■■■
'***************************************************************************************************
Private Function HandlerAddRefInternal(ByVal This As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(This)
    If idx < 0 Then Exit Function
    If m_handlers(idx) Is Nothing Then Exit Function

    Dim N As Long
    N = m_handlers(idx).refCount + 1
    m_handlers(idx).refCount = N
    HandlerAddRefInternal = N
End Function

Private Function HandlerReleaseInternal(ByVal This As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(This)
    If idx < 0 Then Exit Function
    If m_handlers(idx) Is Nothing Then Exit Function

    Dim N As Long
    N = m_handlers(idx).refCount - 1
    If N < 0 Then N = 0
    m_handlers(idx).refCount = N

    If N = 0 Then
        Dim h As CDPWebView2CallbackHandler
        Set h = m_handlers(idx)

        Dim pSlot As LongPtr
        pSlot = h.Slot

        h.ClearOwner
        Set m_handlers(idx) = Nothing
        Thunks_ReleaseSlot pSlot
        Set h = Nothing
    End If

    HandlerReleaseInternal = N
End Function



'***************************************************************************************************
'                              ■■■ メモリプリミティブ(PointerAccessor経由) ■■■
'***************************************************************************************************
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub



'***************************************************************************************************
'                              ■■■ センチネル機構(VBAリセット耐性) ■■■
'***************************************************************************************************
'* 機能　　：前回`Thunks_Init`で確保した領域のベースアドレスを、プロセス環境変数から読み出します
Private Function Sentinel_LoadPrevRegion() As LongPtr
    Sentinel_LoadPrevRegion = 0^

    Dim buff As String
    buff = String$(SENTINEL_BUFFER_SIZE, vbNullChar)

    Dim N As Long
    N = GetEnvironmentVariableW(StrPtr(SENTINEL_ENV_NAME), StrPtr(buff), SENTINEL_BUFFER_SIZE)

    If N = 0 Then Exit Function
    If N >= SENTINEL_BUFFER_SIZE Then Exit Function

    Dim s As String
    s = Left$(buff, N)

    On Error Resume Next
    Sentinel_LoadPrevRegion = CLngLng(s)
    On Error GoTo 0
End Function

Private Sub Sentinel_StorePrevRegion(ByVal addr As LongPtr)
    Dim s As String
    s = CStr(addr)
    SetEnvironmentVariableW StrPtr(SENTINEL_ENV_NAME), StrPtr(s)
End Sub

Private Sub Sentinel_ClearPrevRegion()
    SetEnvironmentVariableW StrPtr(SENTINEL_ENV_NAME), 0
End Sub

'* 機能　　：`Thunks_Init`の冒頭で呼ばれ、前回の痕跡があれば旧領域を回収します
'* 詳細説明：VirtualQueryで「領域が健在(MEM_COMMIT・AllocationBase一致・読み書き可能)」と
'            確認できた場合に限り、実際にメモリへアクセスする(健在でない領域への
'            アクセスはアクセス違反=Excel即落ちにつながるため、健在判定を必ず先に行う)
Private Sub Sentinel_RecoverIfNeeded()
    Dim prevBase As LongPtr
    prevBase = Sentinel_LoadPrevRegion()
    If prevBase = 0 Then Exit Sub

    Dim mbi As MEMORY_BASIC_INFORMATION
    Dim qSize As LongPtr
    Dim regionAlive As Boolean
    regionAlive = False

    qSize = VirtualQuery(prevBase, mbi, LenB(mbi))
    If qSize <> 0 Then
        regionAlive = (mbi.state = MEM_COMMIT) _
                      And (mbi.AllocationBase = prevBase) _
                      And IsProtectReadWritable(mbi.Protect)
    End If

    If Not regionAlive Then
        Sentinel_ClearPrevRegion
        Exit Sub
    End If

    MemLongPtr(prevBase) = 0^  ' 生存フラグを倒す(二重防御)

    Dim freeResult As Long
    freeResult = VirtualFree(prevBase, 0, MEM_RELEASE)
    Debug.Print "CDPWebView2Thunks.Sentinel: recovered previous region " & prevBase & _
                " (VirtualFree result=" & freeResult & ")"

    Sentinel_ClearPrevRegion
End Sub

Private Function IsProtectReadWritable(ByVal prot As Long) As Boolean
    If (prot And &H100&) <> 0 Then Exit Function   ' PAGE_GUARD
    Dim base As Long
    base = prot And Not &H100& And Not &H200& And Not &H400&
    Select Case base
        Case PAGE_READWRITE, PAGE_EXECUTE_READWRITE
            IsProtectReadWritable = True
    End Select
End Function
