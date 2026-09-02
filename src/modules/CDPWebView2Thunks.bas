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

Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" ( _
    ByVal cb As LongPtr) As LongPtr

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

' --- EnvOpt(ICoreWebView2EnvironmentOptions)用の複合fakeオブジェクト、メモリレイアウト定数 ---
' `EnvOpt_CreateNative`が確保するブロックは、7つのCOMインターフェース識別(this)を1つの
' ブロックに同居させる(base+Options2/3/5/6/7/8。Options4[CustomSchemeRegistrations]は
' 配列を扱う特殊な形のため対象外)。各`ENVOPT_THISxxx_OFFSET`がそれぞれの識別(this)セルで、
' その中身(vtable配列の先頭アドレス)が対応する`ENVOPT_VTABLE_xxx_OFFSET`を指す。
' `EnvOpt_ResolveBlockBase`はこの関係の逆算で「どのthisで呼ばれたか」からブロック先頭を復元する
Private Const ENVOPT_THISBASE_OFFSET   As Long = 0     ' ICoreWebView2EnvironmentOptions識別(this)
Private Const ENVOPT_THISOPTS6_OFFSET  As Long = 8     ' ICoreWebView2EnvironmentOptions6識別(this)
Private Const ENVOPT_THISOPTS2_OFFSET  As Long = 152   ' ICoreWebView2EnvironmentOptions2識別(this)
Private Const ENVOPT_THISOPTS3_OFFSET  As Long = 160   ' ICoreWebView2EnvironmentOptions3識別(this)
Private Const ENVOPT_THISOPTS5_OFFSET  As Long = 168   ' ICoreWebView2EnvironmentOptions5識別(this)
Private Const ENVOPT_THISOPTS7_OFFSET  As Long = 176   ' ICoreWebView2EnvironmentOptions7識別(this)
Private Const ENVOPT_THISOPTS8_OFFSET  As Long = 184   ' ICoreWebView2EnvironmentOptions8識別(this)
Private Const ENVOPT_REFCOUNT_OFFSET   As Long = 16
Private Const ENVOPT_VTABLE_BASE_OFFSET  As Long = 24   ' 11スロット(IUnknown3+基底8) * 8bytes = 88
Private Const ENVOPT_VTABLE_OPTS6_OFFSET As Long = 112  ' 5スロット(IUnknown3+Options6用2) * 8bytes = 40
Private Const ENVOPT_VTABLE_OPTS2_OFFSET As Long = 192  ' 5スロット(IUnknown3+Options2用2) * 8bytes = 40
Private Const ENVOPT_VTABLE_OPTS3_OFFSET As Long = 232  ' 5スロット(IUnknown3+Options3用2) * 8bytes = 40
Private Const ENVOPT_VTABLE_OPTS5_OFFSET As Long = 272  ' 5スロット(IUnknown3+Options5用2) * 8bytes = 40
Private Const ENVOPT_VTABLE_OPTS7_OFFSET As Long = 312  ' 7スロット(IUnknown3+Options7用4) * 8bytes = 56
Private Const ENVOPT_VTABLE_OPTS8_OFFSET As Long = 368  ' 5スロット(IUnknown3+Options8用2) * 8bytes = 40
Private Const ENVOPT_BLOCK_SIZE        As Long = 512    ' 408byte使用。余裕を持たせて512確保



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

    ' --- EnvOpt(ICoreWebView2EnvironmentOptions/Options6)の各get_/put_専用 ---
    ' 1個のkind = 1個のvtableスロット(=1個のプロパティのget_またはput_)。
    ' `EnvOpt_CreateNative`が、これら1個ずつに専用のスロットを`AcquireHandlerFor`で
    ' 確保し、その`.Slot`(生の呼び出しエントリ)を自前組み立てのvtable配列へ直接埋め込む
    HK_EnvOpt_GetAdditionalBrowserArguments = 5
    HK_EnvOpt_PutAdditionalBrowserArguments = 6
    HK_EnvOpt_GetLanguage = 7
    HK_EnvOpt_PutLanguage = 8
    HK_EnvOpt_GetTargetCompatibleBrowserVersion = 9
    HK_EnvOpt_PutTargetCompatibleBrowserVersion = 10
    HK_EnvOpt_GetAllowSingleSignOnUsingOSPrimaryAccount = 11
    HK_EnvOpt_PutAllowSingleSignOnUsingOSPrimaryAccount = 12
    HK_EnvOpt_GetAreBrowserExtensionsEnabled = 13
    HK_EnvOpt_PutAreBrowserExtensionsEnabled = 14

    HK_AddBrowserExtensionCompleted = 15    ' ICoreWebView2ProfileAddBrowserExtensionCompletedHandler
    HK_GetBrowserExtensionsCompleted = 16   ' ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler
    HK_RemoveBrowserExtensionCompleted = 17 ' ICoreWebView2BrowserExtensionRemoveCompletedHandler

    ' --- EnvOpt(ICoreWebView2EnvironmentOptions2/3/5/7/8)の各get_/put_専用 ---
    HK_EnvOpt_GetExclusiveUserDataFolderAccess = 18
    HK_EnvOpt_PutExclusiveUserDataFolderAccess = 19
    HK_EnvOpt_GetIsCustomCrashReportingEnabled = 20
    HK_EnvOpt_PutIsCustomCrashReportingEnabled = 21
    HK_EnvOpt_GetEnableTrackingPrevention = 22
    HK_EnvOpt_PutEnableTrackingPrevention = 23
    HK_EnvOpt_GetChannelSearchKind = 24
    HK_EnvOpt_PutChannelSearchKind = 25
    HK_EnvOpt_GetReleaseChannels = 26
    HK_EnvOpt_PutReleaseChannels = 27
    HK_EnvOpt_GetScrollBarStyle = 28
    HK_EnvOpt_PutScrollBarStyle = 29
End Enum



'***************************************************************************************************
'                                   ■■■ 各種モジュール変数 ■■■
'***************************************************************************************************
Private m_pHandler_QI       As LongPtr
Private m_pHandler_AddRef   As LongPtr
Private m_pHandler_Release  As LongPtr

' --- EnvOpt(複合fakeオブジェクト)専用のQI/AddRef/Release実アドレス ---
Private m_pEnvOptQI      As LongPtr
Private m_pEnvOptAddRef  As LongPtr
Private m_pEnvOptRelease As LongPtr

Private m_pRegionBase As LongPtr
Private m_freeHead    As Long
Private m_freeNext()  As Long
Private m_inUse       As Long

Private m_handlers(0 To SLOT_COUNT - 1) As CDPWebView2CallbackHandler
Private m_iidTable(HK_None To HK_RemoveBrowserExtensionCompleted) As GUID
Private m_iidIUnknown As GUID

' --- EnvOptが実装するインターフェースのIID ---
Private m_iidEnvOptBase  As GUID   ' ICoreWebView2EnvironmentOptions
Private m_iidEnvOptOpts6 As GUID   ' ICoreWebView2EnvironmentOptions6
Private m_iidEnvOptOpts2 As GUID   ' ICoreWebView2EnvironmentOptions2
Private m_iidEnvOptOpts3 As GUID   ' ICoreWebView2EnvironmentOptions3
Private m_iidEnvOptOpts5 As GUID   ' ICoreWebView2EnvironmentOptions5
Private m_iidEnvOptOpts7 As GUID   ' ICoreWebView2EnvironmentOptions7
Private m_iidEnvOptOpts8 As GUID   ' ICoreWebView2EnvironmentOptions8

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
'                              ■■■ EnvOpt(ICoreWebView2EnvironmentOptions) ■■■
'***************************************************************************************************
'   `ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler`等の「1メソッドだけのfake
'   オブジェクト」とは異なり、`ICoreWebView2EnvironmentOptions`は8個(get/put4組)、
'   `ICoreWebView2EnvironmentOptions6`はさらに2個(get/put1組)のメソッドを持つ**複合**インター
'   フェースで、しかも両者は継承チェーンではなく独立している(`Settings`/`Controller`/`Profile`の
'   ようにQueryInterfaceで1本のポインタに寄せられない)。そのため、1個のfakeオブジェクトに
'   2本のvtable配列(2種類のthis識別)を持たせる自前実装が必要になる。
'
'   ★方式★
'     ・プロパティ1個(get_またはput_1個)につき、`AcquireHandlerFor`で専用スロットを1個確保し、
'       その`.Slot`(生の呼び出しエントリ。`.Slot + VTABLE_OBJ_OFFSET`ではない点に注意)を
'       自前のvtable配列へ直接埋め込む
'     ・QueryInterface/AddRef/Releaseは複合オブジェクト専用の実装(`EnvOpt_QueryInterface`等、
'       本モジュールの標準Function・AddressOf直結)を新設し、共有する
'     ・ブロックレイアウトは固定オフセット(`ENVOPT_*`定数)。thisポインタ(オフセット0/8)から
'       ブロック先頭を逆算できるよう設計してあり、複数インスタンスが同時に存在してもよい
'     ・参照カウントは実装するが、`Release`が0になっても実メモリ解放は行わない(WebView2Loaderが
'       いつまで参照を保持するか保証がないため)。実解放は`EnvOpt_DestroyNative`をVBA側から
'       明示的に呼んだときのみ行う(`CDPWebView2Host.RunWebView2`が、Environment作成の完了
'       待ち後に呼ぶ)
'***************************************************************************************************
'* 機能　　：`ICoreWebView2EnvironmentOptions`(+`Options6`)を実装するfakeオブジェクトを構築します
'---------------------------------------------------------------------------------------------------
'* 引数　　：owner  `EnvOpt_OnGetXxx`/`EnvOpt_OnPutXxx`(全てPublic Sub必須)を持つ
'            `CDPWebView2EnvOptions`インスタンス
'* 返り値  ：`CreateCoreWebView2EnvironmentWithOptions`へそのまま渡せるポインタ(失敗時0)
'***************************************************************************************************
Public Function EnvOpt_CreateNative(ByVal owner As Object) As LongPtr
    Const FromProcedureName As String = "CDPWebView2Thunks.EnvOpt_CreateNative"

    If m_pRegionBase = 0 Then
        If Not Thunks_Init() Then Exit Function
    End If

    Dim blockBase As LongPtr
    blockBase = VirtualAlloc(0, ENVOPT_BLOCK_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
    If blockBase = 0 Then Exit Function

    Dim k As Long
    For k = 0 To ENVOPT_BLOCK_SIZE - 1 Step 8
        MemLongPtr(blockBase + k) = 0^
    Next k

    ' --- プロパティ10個分(get/put4組+get/put1組)の専用スロットを確保 ---
    Dim hGetArgs As CDPWebView2CallbackHandler, hPutArgs As CDPWebView2CallbackHandler
    Dim hGetLang As CDPWebView2CallbackHandler, hPutLang As CDPWebView2CallbackHandler
    Dim hGetVer  As CDPWebView2CallbackHandler, hPutVer  As CDPWebView2CallbackHandler
    Dim hGetSSO  As CDPWebView2CallbackHandler, hPutSSO  As CDPWebView2CallbackHandler
    Dim hGetExt  As CDPWebView2CallbackHandler, hPutExt  As CDPWebView2CallbackHandler

    Set hGetArgs = AcquireHandlerFor(HK_EnvOpt_GetAdditionalBrowserArguments, owner)
    Set hPutArgs = AcquireHandlerFor(HK_EnvOpt_PutAdditionalBrowserArguments, owner)
    Set hGetLang = AcquireHandlerFor(HK_EnvOpt_GetLanguage, owner)
    Set hPutLang = AcquireHandlerFor(HK_EnvOpt_PutLanguage, owner)
    Set hGetVer = AcquireHandlerFor(HK_EnvOpt_GetTargetCompatibleBrowserVersion, owner)
    Set hPutVer = AcquireHandlerFor(HK_EnvOpt_PutTargetCompatibleBrowserVersion, owner)
    Set hGetSSO = AcquireHandlerFor(HK_EnvOpt_GetAllowSingleSignOnUsingOSPrimaryAccount, owner)
    Set hPutSSO = AcquireHandlerFor(HK_EnvOpt_PutAllowSingleSignOnUsingOSPrimaryAccount, owner)
    Set hGetExt = AcquireHandlerFor(HK_EnvOpt_GetAreBrowserExtensionsEnabled, owner)
    Set hPutExt = AcquireHandlerFor(HK_EnvOpt_PutAreBrowserExtensionsEnabled, owner)

    ' --- プロパティ6個分(Options2/3/5/7[2組]/8)の専用スロットを新規確保 ---
    Dim hGetExcl As CDPWebView2CallbackHandler, hPutExcl As CDPWebView2CallbackHandler
    Dim hGetCrash As CDPWebView2CallbackHandler, hPutCrash As CDPWebView2CallbackHandler
    Dim hGetTrack As CDPWebView2CallbackHandler, hPutTrack As CDPWebView2CallbackHandler
    Dim hGetChKind As CDPWebView2CallbackHandler, hPutChKind As CDPWebView2CallbackHandler
    Dim hGetRelCh As CDPWebView2CallbackHandler, hPutRelCh As CDPWebView2CallbackHandler
    Dim hGetScroll As CDPWebView2CallbackHandler, hPutScroll As CDPWebView2CallbackHandler

    Set hGetExcl = AcquireHandlerFor(HK_EnvOpt_GetExclusiveUserDataFolderAccess, owner)
    Set hPutExcl = AcquireHandlerFor(HK_EnvOpt_PutExclusiveUserDataFolderAccess, owner)
    Set hGetCrash = AcquireHandlerFor(HK_EnvOpt_GetIsCustomCrashReportingEnabled, owner)
    Set hPutCrash = AcquireHandlerFor(HK_EnvOpt_PutIsCustomCrashReportingEnabled, owner)
    Set hGetTrack = AcquireHandlerFor(HK_EnvOpt_GetEnableTrackingPrevention, owner)
    Set hPutTrack = AcquireHandlerFor(HK_EnvOpt_PutEnableTrackingPrevention, owner)
    Set hGetChKind = AcquireHandlerFor(HK_EnvOpt_GetChannelSearchKind, owner)
    Set hPutChKind = AcquireHandlerFor(HK_EnvOpt_PutChannelSearchKind, owner)
    Set hGetRelCh = AcquireHandlerFor(HK_EnvOpt_GetReleaseChannels, owner)
    Set hPutRelCh = AcquireHandlerFor(HK_EnvOpt_PutReleaseChannels, owner)
    Set hGetScroll = AcquireHandlerFor(HK_EnvOpt_GetScrollBarStyle, owner)
    Set hPutScroll = AcquireHandlerFor(HK_EnvOpt_PutScrollBarStyle, owner)

    If hGetArgs Is Nothing Or hPutArgs Is Nothing Or hGetLang Is Nothing Or hPutLang Is Nothing _
        Or hGetVer Is Nothing Or hPutVer Is Nothing Or hGetSSO Is Nothing Or hPutSSO Is Nothing _
        Or hGetExt Is Nothing Or hPutExt Is Nothing _
        Or hGetExcl Is Nothing Or hPutExcl Is Nothing Or hGetCrash Is Nothing Or hPutCrash Is Nothing _
        Or hGetTrack Is Nothing Or hPutTrack Is Nothing Or hGetChKind Is Nothing Or hPutChKind Is Nothing _
        Or hGetRelCh Is Nothing Or hPutRelCh Is Nothing Or hGetScroll Is Nothing Or hPutScroll Is Nothing Then

        Debug.Print FromProcedureName & ": ハンドラスロットの確保に失敗しました"
        If Not (hGetArgs Is Nothing) Then Thunks_ReleaseSlot hGetArgs.Slot
        If Not (hPutArgs Is Nothing) Then Thunks_ReleaseSlot hPutArgs.Slot
        If Not (hGetLang Is Nothing) Then Thunks_ReleaseSlot hGetLang.Slot
        If Not (hPutLang Is Nothing) Then Thunks_ReleaseSlot hPutLang.Slot
        If Not (hGetVer Is Nothing) Then Thunks_ReleaseSlot hGetVer.Slot
        If Not (hPutVer Is Nothing) Then Thunks_ReleaseSlot hPutVer.Slot
        If Not (hGetSSO Is Nothing) Then Thunks_ReleaseSlot hGetSSO.Slot
        If Not (hPutSSO Is Nothing) Then Thunks_ReleaseSlot hPutSSO.Slot
        If Not (hGetExt Is Nothing) Then Thunks_ReleaseSlot hGetExt.Slot
        If Not (hPutExt Is Nothing) Then Thunks_ReleaseSlot hPutExt.Slot
        If Not (hGetExcl Is Nothing) Then Thunks_ReleaseSlot hGetExcl.Slot
        If Not (hPutExcl Is Nothing) Then Thunks_ReleaseSlot hPutExcl.Slot
        If Not (hGetCrash Is Nothing) Then Thunks_ReleaseSlot hGetCrash.Slot
        If Not (hPutCrash Is Nothing) Then Thunks_ReleaseSlot hPutCrash.Slot
        If Not (hGetTrack Is Nothing) Then Thunks_ReleaseSlot hGetTrack.Slot
        If Not (hPutTrack Is Nothing) Then Thunks_ReleaseSlot hPutTrack.Slot
        If Not (hGetChKind Is Nothing) Then Thunks_ReleaseSlot hGetChKind.Slot
        If Not (hPutChKind Is Nothing) Then Thunks_ReleaseSlot hPutChKind.Slot
        If Not (hGetRelCh Is Nothing) Then Thunks_ReleaseSlot hGetRelCh.Slot
        If Not (hPutRelCh Is Nothing) Then Thunks_ReleaseSlot hPutRelCh.Slot
        If Not (hGetScroll Is Nothing) Then Thunks_ReleaseSlot hGetScroll.Slot
        If Not (hPutScroll Is Nothing) Then Thunks_ReleaseSlot hPutScroll.Slot
        VirtualFree blockBase, 0, MEM_RELEASE
        Exit Function
    End If

    ' --- thisセル(それぞれのオフセットに、対応するvtable配列の先頭アドレスを書く) ---
    MemLongPtr(blockBase + ENVOPT_THISBASE_OFFSET) = blockBase + ENVOPT_VTABLE_BASE_OFFSET
    MemLongPtr(blockBase + ENVOPT_THISOPTS6_OFFSET) = blockBase + ENVOPT_VTABLE_OPTS6_OFFSET
    MemLongPtr(blockBase + ENVOPT_THISOPTS2_OFFSET) = blockBase + ENVOPT_VTABLE_OPTS2_OFFSET
    MemLongPtr(blockBase + ENVOPT_THISOPTS3_OFFSET) = blockBase + ENVOPT_VTABLE_OPTS3_OFFSET
    MemLongPtr(blockBase + ENVOPT_THISOPTS5_OFFSET) = blockBase + ENVOPT_VTABLE_OPTS5_OFFSET
    MemLongPtr(blockBase + ENVOPT_THISOPTS7_OFFSET) = blockBase + ENVOPT_VTABLE_OPTS7_OFFSET
    MemLongPtr(blockBase + ENVOPT_THISOPTS8_OFFSET) = blockBase + ENVOPT_VTABLE_OPTS8_OFFSET

    ' --- vtable配列(base、11スロット:IUnknown3+get/put4組) ---
    Dim vb As LongPtr: vb = blockBase + ENVOPT_VTABLE_BASE_OFFSET
    MemLongPtr(vb + 0 * PtrSize) = m_pEnvOptQI
    MemLongPtr(vb + 1 * PtrSize) = m_pEnvOptAddRef
    MemLongPtr(vb + 2 * PtrSize) = m_pEnvOptRelease
    MemLongPtr(vb + 3 * PtrSize) = hGetArgs.Slot
    MemLongPtr(vb + 4 * PtrSize) = hPutArgs.Slot
    MemLongPtr(vb + 5 * PtrSize) = hGetLang.Slot
    MemLongPtr(vb + 6 * PtrSize) = hPutLang.Slot
    MemLongPtr(vb + 7 * PtrSize) = hGetVer.Slot
    MemLongPtr(vb + 8 * PtrSize) = hPutVer.Slot
    MemLongPtr(vb + 9 * PtrSize) = hGetSSO.Slot
    MemLongPtr(vb + 10 * PtrSize) = hPutSSO.Slot

    ' --- vtable配列(Options6、5スロット:IUnknown3+get/put1組) ---
    Dim v6 As LongPtr: v6 = blockBase + ENVOPT_VTABLE_OPTS6_OFFSET
    MemLongPtr(v6 + 0 * PtrSize) = m_pEnvOptQI
    MemLongPtr(v6 + 1 * PtrSize) = m_pEnvOptAddRef
    MemLongPtr(v6 + 2 * PtrSize) = m_pEnvOptRelease
    MemLongPtr(v6 + 3 * PtrSize) = hGetExt.Slot
    MemLongPtr(v6 + 4 * PtrSize) = hPutExt.Slot

    ' --- vtable配列(Options2、5スロット:IUnknown3+get/put1組) ---
    Dim v2 As LongPtr: v2 = blockBase + ENVOPT_VTABLE_OPTS2_OFFSET
    MemLongPtr(v2 + 0 * PtrSize) = m_pEnvOptQI
    MemLongPtr(v2 + 1 * PtrSize) = m_pEnvOptAddRef
    MemLongPtr(v2 + 2 * PtrSize) = m_pEnvOptRelease
    MemLongPtr(v2 + 3 * PtrSize) = hGetExcl.Slot
    MemLongPtr(v2 + 4 * PtrSize) = hPutExcl.Slot

    ' --- vtable配列(Options3、5スロット:IUnknown3+get/put1組) ---
    Dim v3 As LongPtr: v3 = blockBase + ENVOPT_VTABLE_OPTS3_OFFSET
    MemLongPtr(v3 + 0 * PtrSize) = m_pEnvOptQI
    MemLongPtr(v3 + 1 * PtrSize) = m_pEnvOptAddRef
    MemLongPtr(v3 + 2 * PtrSize) = m_pEnvOptRelease
    MemLongPtr(v3 + 3 * PtrSize) = hGetCrash.Slot
    MemLongPtr(v3 + 4 * PtrSize) = hPutCrash.Slot

    ' --- vtable配列(Options5、5スロット:IUnknown3+get/put1組) ---
    Dim v5 As LongPtr: v5 = blockBase + ENVOPT_VTABLE_OPTS5_OFFSET
    MemLongPtr(v5 + 0 * PtrSize) = m_pEnvOptQI
    MemLongPtr(v5 + 1 * PtrSize) = m_pEnvOptAddRef
    MemLongPtr(v5 + 2 * PtrSize) = m_pEnvOptRelease
    MemLongPtr(v5 + 3 * PtrSize) = hGetTrack.Slot
    MemLongPtr(v5 + 4 * PtrSize) = hPutTrack.Slot

    ' --- vtable配列(Options7、7スロット:IUnknown3+ChannelSearchKind/ReleaseChannels各get/put) ---
    Dim v7 As LongPtr: v7 = blockBase + ENVOPT_VTABLE_OPTS7_OFFSET
    MemLongPtr(v7 + 0 * PtrSize) = m_pEnvOptQI
    MemLongPtr(v7 + 1 * PtrSize) = m_pEnvOptAddRef
    MemLongPtr(v7 + 2 * PtrSize) = m_pEnvOptRelease
    MemLongPtr(v7 + 3 * PtrSize) = hGetChKind.Slot
    MemLongPtr(v7 + 4 * PtrSize) = hPutChKind.Slot
    MemLongPtr(v7 + 5 * PtrSize) = hGetRelCh.Slot
    MemLongPtr(v7 + 6 * PtrSize) = hPutRelCh.Slot

    ' --- vtable配列(Options8、5スロット:IUnknown3+get/put1組) ---
    Dim v8 As LongPtr: v8 = blockBase + ENVOPT_VTABLE_OPTS8_OFFSET
    MemLongPtr(v8 + 0 * PtrSize) = m_pEnvOptQI
    MemLongPtr(v8 + 1 * PtrSize) = m_pEnvOptAddRef
    MemLongPtr(v8 + 2 * PtrSize) = m_pEnvOptRelease
    MemLongPtr(v8 + 3 * PtrSize) = hGetScroll.Slot
    MemLongPtr(v8 + 4 * PtrSize) = hPutScroll.Slot

    MemLongPtr(blockBase + ENVOPT_REFCOUNT_OFFSET) = 1^

    EnvOpt_CreateNative = blockBase + ENVOPT_THISBASE_OFFSET
End Function

'***************************************************************************************************
'* 機能　　：VBA(所有者)側が持っている分の参照を1つ手放します
'---------------------------------------------------------------------------------------------------
'* 引数　　：pThisBase  `EnvOpt_CreateNative`の返り値(=ブロック先頭アドレス)
'---------------------------------------------------------------------------------------------------
'* 注意事項：★重要★ 実メモリの解放(`EnvOpt_FreeBlock`)は、ここで無条件には行わない。
'            参照カウントが実際に0になった時(=WebView2Loader側も含め、誰も参照していない
'            ことが確定した時)にのみ`EnvOpt_ReleaseInternal`経由で行われる。
'            (以前は「Environment作成の完了待ち後だから、もう誰も見てないはず」という
'            タイミングの推測だけで無条件`VirtualFree`していたが、WebView2Loader内部の
'            Release呼び出しがそれより後にずれ込むケースがあり、解放済みメモリの読み取り
'            [use-after-free]を引き起こすことが実機で確認された。参照カウントの実値だけを
'            根拠にすることで、どちらが最後に手放しても確実に・二重に壊れず解放される)
'***************************************************************************************************
Public Sub EnvOpt_DestroyNative(ByVal pThisBase As LongPtr)
    If pThisBase = 0 Then Exit Sub
    EnvOpt_ReleaseInternal pThisBase   ' レイアウト上、pThisBase = blockBase
End Sub

'***************************************************************************************************
'* 機能　　：`EnvOpt_CreateNative`が確保した全リソース(10個のスロット+ブロック本体)を実際に解放します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`EnvOpt_ReleaseInternal`が参照カウント0を検知した時にのみ呼ぶこと
'***************************************************************************************************
Private Sub EnvOpt_FreeBlock(ByVal blockBase As LongPtr)
    Dim vb As LongPtr: vb = blockBase + ENVOPT_VTABLE_BASE_OFFSET
    Dim v6 As LongPtr: v6 = blockBase + ENVOPT_VTABLE_OPTS6_OFFSET
    Dim v2 As LongPtr: v2 = blockBase + ENVOPT_VTABLE_OPTS2_OFFSET
    Dim v3 As LongPtr: v3 = blockBase + ENVOPT_VTABLE_OPTS3_OFFSET
    Dim v5 As LongPtr: v5 = blockBase + ENVOPT_VTABLE_OPTS5_OFFSET
    Dim v7 As LongPtr: v7 = blockBase + ENVOPT_VTABLE_OPTS7_OFFSET
    Dim v8 As LongPtr: v8 = blockBase + ENVOPT_VTABLE_OPTS8_OFFSET

    Dim i As Long
    For i = 3 To 10
        Thunks_ReleaseSlot ReadLongPtr(vb + i * PtrSize)
    Next i
    For i = 3 To 4
        Thunks_ReleaseSlot ReadLongPtr(v6 + i * PtrSize)
        Thunks_ReleaseSlot ReadLongPtr(v2 + i * PtrSize)
        Thunks_ReleaseSlot ReadLongPtr(v3 + i * PtrSize)
        Thunks_ReleaseSlot ReadLongPtr(v5 + i * PtrSize)
        Thunks_ReleaseSlot ReadLongPtr(v8 + i * PtrSize)
    Next i
    For i = 3 To 6
        Thunks_ReleaseSlot ReadLongPtr(v7 + i * PtrSize)
    Next i

    VirtualFree blockBase, 0, MEM_RELEASE
End Sub

'***************************************************************************************************
'* 機能　　：get_Xxx(LPWSTR* value)へ、文字列値を書き出します
'---------------------------------------------------------------------------------------------------
'* 引数　　：pOut  [out]ポインタ(呼び出し元がCoTaskMemFreeする前提)
'            s     書き出す文字列。空文字なら`nullptr`(未設定扱い)を書く
'***************************************************************************************************
Public Sub EnvOpt_WriteStringOut(ByVal pOut As LongPtr, ByVal s As String)
    If pOut = 0 Then Exit Sub
    If LenB(s) = 0 Then
        MemLongPtr(pOut) = 0^
    Else
        MemLongPtr(pOut) = StringToCoTaskMem(s)
    End If
End Sub

'***************************************************************************************************
'* 機能　　：get_Xxx(BOOL* value)へ、真偽値を書き出します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`BOOL`は4byteだが、この基盤の書き込みプリミティブは8byte単位(`MemLongPtr`)しか
'            持たない。8byte書き込みで隣接メモリを破壊しないよう、まず既存の8byteを読み、
'            上位4byteは元の値を保持したまま、下位4byteだけを差し替えて書き戻す
'***************************************************************************************************
Public Sub EnvOpt_WriteBoolOut(ByVal pOut As LongPtr, ByVal v As Boolean)
    If pOut = 0 Then Exit Sub

    Dim existing As LongLong
    existing = CLngLng(ReadLongPtr(pOut))

    Dim merged As LongLong
    merged = (existing And &HFFFFFFFF00000000^) Or CLngLng(IIf(v, 1, 0))

    MemLongPtr(pOut) = merged
End Sub

'***************************************************************************************************
'* 機能　　：get_Xxx(enumへのポインタ、またはビットマスクのLong* value)へ、Long値を書き出します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`EnvOpt_WriteBoolOut`と同じ「上位4byte保持」ロジックをLong値向けに汎用化したもの
'            (`ChannelSearchKind`/`ReleaseChannels`/`ScrollBarStyle`用)
'***************************************************************************************************
Public Sub EnvOpt_WriteLongOut(ByVal pOut As LongPtr, ByVal v As Long)
    If pOut = 0 Then Exit Sub

    Dim existing As LongLong
    existing = CLngLng(ReadLongPtr(pOut))

    Dim merged As LongLong
    merged = (existing And &HFFFFFFFF00000000^) Or (CLngLng(v) And &HFFFFFFFF^)

    MemLongPtr(pOut) = merged
End Sub

'* 機能　　：VBAの文字列をCoTaskMemAllocされたLPWSTRへ複製します(呼び出し元がCoTaskMemFreeする)
Private Function StringToCoTaskMem(ByVal s As String) As LongPtr
    Dim cb As LongPtr
    cb = CLngLng(Len(s) + 1) * 2

    Dim p As LongPtr
    p = CoTaskMemAlloc(cb)
    If p <> 0 Then lstrcpyW p, StrPtr(s)

    StringToCoTaskMem = p
End Function

'***************************************************************************************************
'* 機能　　：EnvOptの`this`ポインタ(7つのCOMインターフェース識別のいずれか)から、ブロック先頭
'            アドレスを逆算します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：レイアウトが固定なので、「(thisセルのオフセット, 対応vtable配列のオフセット)」の
'            全組み合わせ(7通り)を検算するだけで求まる…はずだったが、実機で誤検知が発生した。
'            `ReadLongPtr(This) = cand + vtblOffset(i)`という1本の等式だけでは、
'            `thisOffset(i) - vtblOffset(i)`の値が別のiと偶然一致してしまうと、本来とは違う
'            iで先にマッチしてしまう(実際に`i=1`[Opts6: 8-112=-104]と`i=4`[Opts5: 168-272=-104]が
'            衝突し、`This`=本物のOpts5識別セルなのに`i=1`の式で誤ってマッチしてしまうバグが
'            実機で確認された)。そのため、候補が見つかった時点で「その候補自身のbase識別セルが、
'            自分自身のvtable配列(base)を指しているか」という独立した等式でも裏付けを取り、
'            本物のブロック先頭であることを二重に保証する
'***************************************************************************************************
Private Function EnvOpt_ResolveBlockBase(ByVal This As LongPtr) As LongPtr
    If This = 0 Then Exit Function

    ' 「(thisセルのオフセット, 対応vtable配列のオフセット)」の全組み合わせ(7通り)を検算する。
    ' 正しい組み合わせだけが「候補ブロック先頭 + vtable配列オフセット」= 「thisの中身」になる
    Dim thisOffsets(0 To 6) As Long, vtblOffsets(0 To 6) As Long
    thisOffsets(0) = ENVOPT_THISBASE_OFFSET:  vtblOffsets(0) = ENVOPT_VTABLE_BASE_OFFSET
    thisOffsets(1) = ENVOPT_THISOPTS6_OFFSET: vtblOffsets(1) = ENVOPT_VTABLE_OPTS6_OFFSET
    thisOffsets(2) = ENVOPT_THISOPTS2_OFFSET: vtblOffsets(2) = ENVOPT_VTABLE_OPTS2_OFFSET
    thisOffsets(3) = ENVOPT_THISOPTS3_OFFSET: vtblOffsets(3) = ENVOPT_VTABLE_OPTS3_OFFSET
    thisOffsets(4) = ENVOPT_THISOPTS5_OFFSET: vtblOffsets(4) = ENVOPT_VTABLE_OPTS5_OFFSET
    thisOffsets(5) = ENVOPT_THISOPTS7_OFFSET: vtblOffsets(5) = ENVOPT_VTABLE_OPTS7_OFFSET
    thisOffsets(6) = ENVOPT_THISOPTS8_OFFSET: vtblOffsets(6) = ENVOPT_VTABLE_OPTS8_OFFSET

    Dim i As Long, cand As LongPtr
    For i = 0 To 6
        cand = This - thisOffsets(i)
        If ReadLongPtr(This) = cand + vtblOffsets(i) Then
            ' ★裏付けチェック★ 候補自身のbase識別セル(cand+0)が、候補自身のbase vtable配列
            ' (cand+24)を指しているか。他のiとの偶然の一致(上記詳細説明参照)を弾くための独立検算
            If ReadLongPtr(cand + ENVOPT_THISBASE_OFFSET) = cand + ENVOPT_VTABLE_BASE_OFFSET Then
                EnvOpt_ResolveBlockBase = cand
                Exit Function
            End If
        End If
    Next i
End Function

Private Function EnvOpt_QueryInterface( _
    ByVal This As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long

    If riid = 0 Then
        ppvObject = 0
        EnvOpt_QueryInterface = &H80004003   ' E_POINTER
        Exit Function
    End If

    Dim blockBase As LongPtr
    blockBase = EnvOpt_ResolveBlockBase(This)
    If blockBase = 0 Then
        ppvObject = 0
        EnvOpt_QueryInterface = E_NOINTERFACE
        Exit Function
    End If

    If IsEqualGUIDInPlace(riid, m_iidIUnknown) Or IsEqualGUIDInPlace(riid, m_iidEnvOptBase) Then
        ppvObject = blockBase + ENVOPT_THISBASE_OFFSET
        EnvOpt_AddRefInternal blockBase
        EnvOpt_QueryInterface = S_OK
        Exit Function
    End If

    If IsEqualGUIDInPlace(riid, m_iidEnvOptOpts6) Then
        ppvObject = blockBase + ENVOPT_THISOPTS6_OFFSET
        EnvOpt_AddRefInternal blockBase
        EnvOpt_QueryInterface = S_OK
        Exit Function
    End If

    If IsEqualGUIDInPlace(riid, m_iidEnvOptOpts2) Then
        ppvObject = blockBase + ENVOPT_THISOPTS2_OFFSET
        EnvOpt_AddRefInternal blockBase
        EnvOpt_QueryInterface = S_OK
        Exit Function
    End If

    If IsEqualGUIDInPlace(riid, m_iidEnvOptOpts3) Then
        ppvObject = blockBase + ENVOPT_THISOPTS3_OFFSET
        EnvOpt_AddRefInternal blockBase
        EnvOpt_QueryInterface = S_OK
        Exit Function
    End If

    If IsEqualGUIDInPlace(riid, m_iidEnvOptOpts5) Then
        ppvObject = blockBase + ENVOPT_THISOPTS5_OFFSET
        EnvOpt_AddRefInternal blockBase
        EnvOpt_QueryInterface = S_OK
        Exit Function
    End If

    If IsEqualGUIDInPlace(riid, m_iidEnvOptOpts7) Then
        ppvObject = blockBase + ENVOPT_THISOPTS7_OFFSET
        EnvOpt_AddRefInternal blockBase
        EnvOpt_QueryInterface = S_OK
        Exit Function
    End If

    If IsEqualGUIDInPlace(riid, m_iidEnvOptOpts8) Then
        ppvObject = blockBase + ENVOPT_THISOPTS8_OFFSET
        EnvOpt_AddRefInternal blockBase
        EnvOpt_QueryInterface = S_OK
        Exit Function
    End If

    ppvObject = 0
    EnvOpt_QueryInterface = E_NOINTERFACE
End Function

Private Function EnvOpt_AddRef(ByVal This As LongPtr) As Long
    EnvOpt_AddRef = EnvOpt_AddRefInternal(EnvOpt_ResolveBlockBase(This))
End Function

'* 機能　　：参照カウントを減らし、実際に0になった時だけ`EnvOpt_FreeBlock`で実メモリ解放する
Private Function EnvOpt_Release(ByVal This As LongPtr) As Long
    EnvOpt_Release = EnvOpt_ReleaseInternal(EnvOpt_ResolveBlockBase(This))
End Function

'* 機能　　：参照カウントを1増やす(`IUnknown::AddRef`はULONG[32bit]契約のため、内部でも
'            `Long`で完結させる。8バイト単位でしか読み書きできない`ReadLongPtr`/`MemLongPtr`
'            との境界だけ、読み取った直後に一度だけ`Long`へ narrow する)
Private Function EnvOpt_AddRefInternal(ByVal blockBase As LongPtr) As Long
    If blockBase = 0 Then Exit Function

    Dim N As LongLong
    N = ReadLongPtr(blockBase + ENVOPT_REFCOUNT_OFFSET) + 1
    MemLongPtr(blockBase + ENVOPT_REFCOUNT_OFFSET) = N

    EnvOpt_AddRefInternal = CLng(N)
End Function

'* 機能　　：参照カウントを1減らす。0になった場合のみ、実メモリ(スロット+ブロック本体)を解放する
'---------------------------------------------------------------------------------------------------
'* 注意事項：`EnvOpt_AddRef`(WebView2Loader経由)と`EnvOpt_DestroyNative`(VBA所有者側)の
'            両方から、同じ1つのカウンタに対して呼ばれる。どちら経由で最後の1つを手放しても、
'            ここで初めて`EnvOpt_FreeBlock`が呼ばれるため、解放タイミングの推測が不要になる
'***************************************************************************************************
Private Function EnvOpt_ReleaseInternal(ByVal blockBase As LongPtr) As Long
    If blockBase = 0 Then Exit Function

    Dim N As LongLong
    N = ReadLongPtr(blockBase + ENVOPT_REFCOUNT_OFFSET) - 1
    If N < 0 Then N = 0
    MemLongPtr(blockBase + ENVOPT_REFCOUNT_OFFSET) = N

    If N = 0 Then EnvOpt_FreeBlock blockBase

    EnvOpt_ReleaseInternal = CLng(N)
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

'* 機能　　：`HRESULT get_Xxx([out,retval] BOOL *value)`形のCOMメソッドを呼び、Booleanで返します
Public Function GetBoolProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As Boolean

    If pInterface = 0 Then Exit Function

    Dim v As Long
    dcf pInterface, vtblIndex, funcName, VarPtr(v)
    GetBoolProperty = (v <> 0)
End Function

'* 機能　　：`HRESULT get_Xxx([out,retval] LONG/INT32/enum *value)`形のCOMメソッドを呼び、Longで返します
Public Function GetLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As Long

    If pInterface = 0 Then Exit Function

    Dim v As Long
    dcf pInterface, vtblIndex, funcName, VarPtr(v)
    GetLongProperty = v
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

    m_pEnvOptQI = GetAddr(AddressOf EnvOpt_QueryInterface)
    m_pEnvOptAddRef = GetAddr(AddressOf EnvOpt_AddRef)
    m_pEnvOptRelease = GetAddr(AddressOf EnvOpt_Release)
    If m_pEnvOptQI = 0 Or m_pEnvOptAddRef = 0 Or m_pEnvOptRelease = 0 Then
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

    ' ICoreWebView2EnvironmentOptions
    FillGUID m_iidEnvOptBase, "2fde08a8-1e9a-4766-8c05-95a9ceb9d1c5"

    ' ICoreWebView2EnvironmentOptions6
    FillGUID m_iidEnvOptOpts6, "57d29cc3-c84f-42a0-b0e2-effbd5e179de"

    ' ICoreWebView2EnvironmentOptions2
    FillGUID m_iidEnvOptOpts2, "ff85c98a-1ba7-4a6b-90c8-2b752c89e9e2"

    ' ICoreWebView2EnvironmentOptions3
    FillGUID m_iidEnvOptOpts3, "4a5c436e-a9e3-4a2e-89c3-910d3513f5cc"

    ' ICoreWebView2EnvironmentOptions5
    FillGUID m_iidEnvOptOpts5, "0ae35d64-c47f-4464-814e-259c345d1501"

    ' ICoreWebView2EnvironmentOptions7
    FillGUID m_iidEnvOptOpts7, "c48d539f-e39f-441c-ae68-1f66e570bdc5"

    ' ICoreWebView2EnvironmentOptions8
    FillGUID m_iidEnvOptOpts8, "7c7ecf51-e918-5caf-853c-e9a2bcc27775"

    ' ICoreWebView2ProfileAddBrowserExtensionCompletedHandler
    FillGUID m_iidTable(HK_AddBrowserExtensionCompleted), _
             "df1aab27-82b9-4ab6-aae8-017a49398c14"

    ' ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler
    FillGUID m_iidTable(HK_GetBrowserExtensionsCompleted), _
             "fce16a1c-f107-4601-8b75-fc4940ae25d0"

    ' ICoreWebView2BrowserExtensionRemoveCompletedHandler
    FillGUID m_iidTable(HK_RemoveBrowserExtensionCompleted), _
             "8e41909a-9b18-4bb1-8cdf-930f467a50be"
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
