VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
' WebView2Core.cls
' TLB不要・参照設定不要・外部ファイル不要の 64bit VBA 純粋実装
' DispCallFunc + vtable偽造 による WebView2 COM 直接制御
'
' 動作要件:
'   - 64bit Excel (VBA7)
'   - Microsoft Edge (Chromium) インストール済み
'   - WebView2Callbacks.bas（AddressOf制約のため必須）
'
' 基本フロー:
'   Initialize → OnEnvironmentCreated（CB） → OnControllerCreated（CB）
'   → Ready イベント → Navigate でページ遷移
'***

' =====================================================================
' 型定義
' =====================================================================
' COM vtable 偽造オブジェクト（16バイト: vtablePtr(8) のみ）
Private Type VtblObj
    pVTable As LongPtr
End Type
' 4エントリの vtable（QI / AddRef / Release / Invoke）
Private Type VtblData4
    fn(0 To 3) As LongPtr
End Type

Private Type RECT_T
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

' =====================================================================
' Win32 API Declares (64bit PtrSafe)
' =====================================================================
#If VBA7 Then

' WebView2Loader.dll ─ Edge に同梱されている DLL を使用
' Edge インストールフォルダ（通常は自動的に PATH に含まれないため、
' Initialize 内で LoadLibrary を用いてフルパスからロードする）
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions _
    Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder          As LongPtr, _
    ByVal environmentOptions      As LongPtr, _
    ByVal handler                 As LongPtr) As Long

' OleAut32 ─ COM vtable メソッドを番号（オフセット）で呼ぶ
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance  As LongPtr, _
    ByVal oVft        As LongPtr, _
    ByVal lCc         As Long, _
    ByVal vtReturn    As Integer, _
    ByVal cActuals    As Long, _
    prgvt             As Any, _
    prgpvarg          As Any, _
    pvargResult       As Variant) As Long

' Kernel32 ─ メモリ操作・DLLロード
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" ( _
    ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr

' User32 ─ ウィンドウ操作
Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal uFlags As Long) As Long

' User32 ─ メッセージポンプ（COM STA コールバックの配送に必要）
Private Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageW" ( _
    lpMsg As Any, ByVal hwnd As LongPtr, _
    ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, _
    ByVal wRemoveMsg As Long) As Long
Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As Any) As Long
Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (lpMsg As Any) As LongPtr

#End If

' =====================================================================
' 定数
' =====================================================================
Private Const CC_STDCALL  As Long = 4
Private Const PTR_SIZE    As Long = 8   ' 64bit ポインタ = 8バイト
Private Const PM_REMOVE   As Long = 1   ' PeekMessage フラグ

' --- ICoreWebView2Environment vtable オフセット (index × PTR_SIZE) ---
Private Const OFF_ENV_CreateController As Long = 3 * 8   ' = 24

' --- ICoreWebView2Controller vtable オフセット ---
Private Const OFF_CTRL_get_CoreWebView2 As Long = 3 * 8  ' = 24
Private Const OFF_CTRL_put_IsVisible    As Long = 5 * 8  ' = 40
Private Const OFF_CTRL_put_Bounds       As Long = 7 * 8  ' = 56 (x64ではポインタ渡し)
Private Const OFF_CTRL_Close            As Long = 25 * 8 ' = 200

' --- ICoreWebView2 vtable オフセット ---
Private Const OFF_WV2_Navigate             As Long = 5 * 8  ' = 40
Private Const OFF_WV2_NavigateToString     As Long = 6 * 8  ' = 48
Private Const OFF_WV2_add_NavCompleted     As Long = 15 * 8 ' = 120
Private Const OFF_WV2_ExecuteScript        As Long = 38 * 8 ' ※要検証

' =====================================================================
' プライベート状態
' =====================================================================
' vtable 偽造オブジェクト（2つ: 環境ハンドラ / コントローラーハンドラ）
Private m_EnvHandlerThis  As VtblObj
Private m_EnvHandlerVT    As VtblData4
Private m_CtrlHandlerThis As VtblObj
Private m_CtrlHandlerVT   As VtblData4

' COM オブジェクトポインタ
Private m_pEnvironment  As LongPtr   ' ICoreWebView2Environment
Private m_pController   As LongPtr   ' ICoreWebView2Controller
Private m_pWebView2     As LongPtr   ' ICoreWebView2

' 親ウィンドウ
Private m_hWndParent As LongPtr

' 表示領域
Private m_Bounds As RECT_T

' 初期遷移 URL
Private m_InitialUrl As String

' エラー情報
Private m_LastErrorCode As Long
Private m_LastErrorDesc As String

' 初期化ステート管理
'   0 = 初期 / 1 = Environment取得済み(Controller未作成) / 2 = Ready
Private m_State As Long

' =====================================================================
' Public Events
' =====================================================================
Public Event Ready()
Public Event NavigationCompleted(ByVal isSuccess As Boolean, ByVal webErrorStatus As Long)
Public Event InitializeFailed(ByVal errorCode As Long, ByVal description As String)

' =====================================================================
' Public Interface
' =====================================================================

'----------------------------------------------------------------------
' Initialize
'   hWndParent : WebView2 を埋め込む親ウィンドウハンドル
'   boundsLeft/Top/Width/Height : 表示領域（ピクセル）
'   initialUrl : 最初に遷移するURL（省略可）
'   戻り値 : True = 初期化コマンド送信成功（Ready イベント待ち）
'----------------------------------------------------------------------
Public Function Initialize( _
    ByVal hWndParent  As LongPtr, _
    ByVal boundsLeft  As Long, _
    ByVal boundsTop   As Long, _
    ByVal boundsWidth As Long, _
    ByVal boundsHeight As Long, _
    Optional ByVal initialUrl As String = "https://www.google.com") As Boolean

    m_hWndParent   = hWndParent
    m_InitialUrl   = initialUrl
    m_Bounds.Left  = boundsLeft
    m_Bounds.Top   = boundsTop
    m_Bounds.Right  = boundsLeft + boundsWidth
    m_Bounds.Bottom = boundsTop + boundsHeight

    ' グローバル参照をセット（コールバックから Me へ戻るため）
    Set g_WebView2Core = Me

    ' vtable を構築
    BuildVtables

    ' WebView2Loader.dll が見つからない環境用に Edge フォルダから明示ロード
    EnsureWebView2LoaderDll

    ' 環境作成開始（非同期; WV2_EnvCB_Invoke が呼ばれると CB_EnvironmentCreated が動く）
    Dim hr As Long
    Dim userDataPath As String
    userDataPath = Environ("LOCALAPPDATA") & "\WebView2_VBA_Profile"

    Dim pUserData   As LongPtr: pUserData = StrPtr(userDataPath)
    Dim pHandler    As LongPtr: pHandler  = VarPtr(m_EnvHandlerThis)

    hr = CreateCoreWebView2EnvironmentWithOptions(0, pUserData, 0, pHandler)

    If hr <> 0 Then
        m_LastErrorCode = hr
        m_LastErrorDesc = "CreateCoreWebView2EnvironmentWithOptions failed: 0x" & Hex(hr)
        RaiseEvent InitializeFailed(hr, m_LastErrorDesc)
        Initialize = False
    Else
        Initialize = True
    End If
End Function

'----------------------------------------------------------------------
' Navigate : 指定 URL へ遷移
'----------------------------------------------------------------------
Public Sub Navigate(ByVal url As String)
    If m_pWebView2 = 0 Then Exit Sub
    Dim hr As Long
    hr = COM_Navigate(m_pWebView2, url)
End Sub

'----------------------------------------------------------------------
' NavigateToString : HTML 文字列を直接描画
'----------------------------------------------------------------------
Public Sub NavigateToString(ByVal htmlContent As String)
    If m_pWebView2 = 0 Then Exit Sub
    COM_DispCallPtr1 m_pWebView2, OFF_WV2_NavigateToString, StrPtr(htmlContent)
End Sub

'----------------------------------------------------------------------
' Resize : 表示領域を変更
'----------------------------------------------------------------------
Public Sub Resize(ByVal newLeft As Long, ByVal newTop As Long, _
                  ByVal newWidth As Long, ByVal newHeight As Long)
    m_Bounds.Left   = newLeft
    m_Bounds.Top    = newTop
    m_Bounds.Right  = newLeft + newWidth
    m_Bounds.Bottom = newTop + newHeight
    If m_pController <> 0 Then COM_SetBounds m_pController, m_Bounds
End Sub

'----------------------------------------------------------------------
' ProcessMessages  ― COM STA コールバック層に必要なメッセージの処理
'   IsReady になるまでの待機ループ内で繰り返し呼ぶ。
'   State 1（Environment取得済み）を検出したら Controller 作成を開始する。
'----------------------------------------------------------------------
Public Sub ProcessMessages(Optional ByVal iterations As Long = 100)
    Dim msg(0 To 5) As LongPtr
    Dim i As Long
    For i = 1 To iterations
        If PeekMessage(msg(0), 0, 0, 0, PM_REMOVE) Then
            TranslateMessage msg(0)
            DispatchMessage msg(0)
        Else
            Exit For
        End If
    Next i
    DoEvents

    ' State 1: Environment → Controller 作成
    If m_State = 1 Then BeginControllerCreation
    ' State 2: Controller → 表示・WebView2取得
    If m_State = 2 Then FinishWebViewSetup
End Sub

Public Sub SetVisible(ByVal isVisible As Boolean)
    If m_pController = 0 Then Exit Sub
    Dim visVal As Long: visVal = IIf(isVisible, 1, 0)
    COM_DispCallLong1 m_pController, OFF_CTRL_put_IsVisible, visVal
End Sub

'----------------------------------------------------------------------
' Quit : リソース解放
'----------------------------------------------------------------------
Public Sub Quit()
    On Error Resume Next
    If m_pController <> 0 Then
        Dim retV As Variant
        DispCallFunc m_pController, OFF_CTRL_Close, CC_STDCALL, vbLong, 0, 0, 0, retV
    End If
    m_pWebView2     = 0
    m_pController   = 0
    m_pEnvironment  = 0
    Set g_WebView2Core = Nothing
End Sub

'----------------------------------------------------------------------
' LastError : 直前のエラー情報を返す
'----------------------------------------------------------------------
Public Property Get LastErrorCode() As Long:        LastErrorCode = m_LastErrorCode:   End Property
Public Property Get LastErrorDescription() As String: LastErrorDescription = m_LastErrorDesc: End Property
Public Property Get IsReady() As Boolean:           IsReady = (m_pWebView2 <> 0):     End Property

' =====================================================================
' Friend（Callbacks から呼ばれる）
' =====================================================================

'----------------------------------------------------------------------
' CB_EnvironmentCreated
'   WV2_EnvCB_Invoke → ここへ委譲。
'   ※ コールバック内で CreateCoreWebView2Controller を呼ぶと再入問題が起きるため、
'     pEnv を保存して m_State=1 をセットするだけにする。
'     実際の Controller 作成は ProcessMessages → BeginControllerCreation で行う。
'----------------------------------------------------------------------
Friend Sub CB_EnvironmentCreated(ByVal errorCode As Long, ByVal pEnv As LongPtr)
    Debug.Print "[WV2] CB_EnvironmentCreated: errorCode=" & errorCode & ", pEnv=" & Hex(pEnv)
    If errorCode <> 0 Or pEnv = 0 Then
        m_LastErrorCode = errorCode
        m_LastErrorDesc = "Environment creation failed: 0x" & Hex(errorCode)
        RaiseEvent InitializeFailed(errorCode, m_LastErrorDesc)
        Exit Sub
    End If
    COM_AddRef pEnv
    m_pEnvironment = pEnv
    m_State = 1     ' → BeginControllerCreation へのシグナル
End Sub

'----------------------------------------------------------------------
' BeginControllerCreation（ProcessMessages から呼ばれる）
' 注: DispCallFunc の prgpvarg は VARIANTARG*(16byte構造体) へのポインタ配列
'----------------------------------------------------------------------
Private Sub BeginControllerCreation()
    m_State = 9     ' 「作成中」マーク（二重呼び出し防止）

    ' --- 診断ログ ---
    Dim pCtrlHandler As LongPtr: pCtrlHandler = VarPtr(m_CtrlHandlerThis)
    Debug.Print "[WV2] BeginControllerCreation"
    Debug.Print "  hWnd          = 0x" & Hex(m_hWndParent)
    Debug.Print "  pEnvironment  = 0x" & Hex(m_pEnvironment)
    Debug.Print "  pCtrlHandler  = 0x" & Hex(pCtrlHandler)
    Debug.Print "  vtable ptr    = 0x" & Hex(m_CtrlHandlerThis.pVTable)

    ' --- VARIANT 構造体ベースで引数を渡す ---
    ' VARIANTARG: vt(2)+予約(6)+value(8) = 16 bytes
    ' DispCallFunc が *prgpvarg[i] の offset+8 から値を読む場合に対応
    Dim vHwnd(0 To 1) As LongPtr    ' [0]=vt+pad(8bytes), [1]=value(8bytes)
    Dim vHandler(0 To 1) As LongPtr

    ' vt フィールド (下位2バイト) に vbLongLong=20 をセット、残りは0
    Dim vtLongLong As Integer: vtLongLong = 20  ' VT_I8
    CopyMemory vHwnd(0), vtLongLong, 2
    vHwnd(1) = m_hWndParent

    CopyMemory vHandler(0), vtLongLong, 2
    vHandler(1) = pCtrlHandler

    Dim argT(1) As Integer
    Dim argP(1) As LongPtr
    argT(0) = vbLongLong: argP(0) = VarPtr(vHwnd(0))
    argT(1) = vbLongLong: argP(1) = VarPtr(vHandler(0))

    Dim retV As Variant
    Dim hr As Long
    hr = DispCallFunc(m_pEnvironment, OFF_ENV_CreateController, CC_STDCALL, vbLong, 2, argT(0), argP(0), retV)
    Debug.Print "  hr(DispCallFunc)=" & Hex(hr) & "  result(HRESULT)=" & Hex(CLng(retV))
End Sub

'----------------------------------------------------------------------
' CB_ControllerCreated
'   WV2_CtrlCB_Invoke → ここへ委譲。
'   ※ コールバック内での COM 呼び出しは VBE7.DLL クラッシュの原因になるため、
'     pCtrl を保存して m_State=2 をセットするだけにする。
'     実際の表示処理は ProcessMessages → FinishWebViewSetup で行う。
'----------------------------------------------------------------------
Friend Sub CB_ControllerCreated(ByVal errorCode As Long, ByVal pCtrl As LongPtr)
    Debug.Print "[WV2] CB_ControllerCreated: errorCode=" & errorCode & ", pCtrl=" & Hex(pCtrl)
    If errorCode <> 0 Or pCtrl = 0 Then
        m_LastErrorCode = errorCode
        m_LastErrorDesc = "Controller creation failed: 0x" & Hex(errorCode)
        RaiseEvent InitializeFailed(errorCode, m_LastErrorDesc)
        Exit Sub
    End If
    ' AddRef（0引数 DispCallFunc）はコールバック内でも安全（山引数ありの呼び出しがクラッシュするのでこちらだけ残す）
    COM_AddRef pCtrl
    m_pController = pCtrl
    m_State = 2
End Sub

'----------------------------------------------------------------------
' FinishWebViewSetup（ProcessMessages から呼ばれる）
'   BeginControllerCreation と同じ VARIANT構造体ベースの引数渡しで COM を呼ぶ。
'----------------------------------------------------------------------
Private Sub FinishWebViewSetup()
    m_State = 99
    Debug.Print "[WV2] FinishWebViewSetup: pController=0x" & Hex(m_pController)

    ' VARIANTARGバッファ: [0]=vt+予約(8bytes) [1]=value(8bytes)
    Dim vtII As Integer: vtII = 20  ' VT_I8
    Dim retV As Variant

    ' --- put_Bounds (RECT ポインタ渡し) ---
    Dim vBounds(0 To 1) As LongPtr
    CopyMemory vBounds(0), vtII, 2
    Dim pBoundsArg As LongPtr: pBoundsArg = VarPtr(m_Bounds)
    vBounds(1) = pBoundsArg
    Dim aT1(0) As Integer: aT1(0) = 20
    Dim aP1(0) As LongPtr: aP1(0) = VarPtr(vBounds(0))
    DispCallFunc m_pController, OFF_CTRL_put_Bounds, CC_STDCALL, vbLong, 1, aT1(0), aP1(0), retV
    Debug.Print "  put_Bounds HRESULT=0x" & Hex(CLng(retV))

    ' --- put_IsVisible (BOOL=1) ---
    Dim vVis(0 To 1) As LongPtr
    CopyMemory vVis(0), vtII, 2
    vVis(1) = 1
    Dim aT2(0) As Integer: aT2(0) = 20
    Dim aP2(0) As LongPtr: aP2(0) = VarPtr(vVis(0))
    DispCallFunc m_pController, OFF_CTRL_put_IsVisible, CC_STDCALL, vbLong, 1, aT2(0), aP2(0), retV
    Debug.Print "  put_IsVisible HRESULT=0x" & Hex(CLng(retV))

    ' --- get_CoreWebView2 (出力: ICoreWebView2*) ---
    Dim pWV2 As LongPtr
    Dim vWV2(0 To 1) As LongPtr
    CopyMemory vWV2(0), vtII, 2
    vWV2(1) = VarPtr(pWV2)
    Dim aT3(0) As Integer: aT3(0) = 20
    Dim aP3(0) As LongPtr: aP3(0) = VarPtr(vWV2(0))
    DispCallFunc m_pController, OFF_CTRL_get_CoreWebView2, CC_STDCALL, vbLong, 1, aT3(0), aP3(0), retV
    m_pWebView2 = pWV2
    Debug.Print "  get_CoreWebView2 HRESULT=0x" & Hex(CLng(retV)) & ", pWV2=0x" & Hex(m_pWebView2)

    ' Ready 通知
    RaiseEvent Ready
    If Len(m_InitialUrl) > 0 Then Navigate m_InitialUrl
End Sub

' =====================================================================
' Private ─ vtable 構築
' =====================================================================
Private Sub BuildVtables()
    ' コールバック関数アドレスを Callbacks モジュールから取得
    Dim envFn()  As LongPtr
    Dim ctrlFn() As LongPtr
    WV2_FillFunctionPointers envFn, ctrlFn

    ' 環境ハンドラ vtable
    m_EnvHandlerVT.fn(0) = envFn(0)
    m_EnvHandlerVT.fn(1) = envFn(1)
    m_EnvHandlerVT.fn(2) = envFn(2)
    m_EnvHandlerVT.fn(3) = envFn(3)
    m_EnvHandlerThis.pVTable = VarPtr(m_EnvHandlerVT.fn(0))

    ' コントローラーハンドラ vtable
    m_CtrlHandlerVT.fn(0) = ctrlFn(0)
    m_CtrlHandlerVT.fn(1) = ctrlFn(1)
    m_CtrlHandlerVT.fn(2) = ctrlFn(2)
    m_CtrlHandlerVT.fn(3) = ctrlFn(3)
    m_CtrlHandlerThis.pVTable = VarPtr(m_CtrlHandlerVT.fn(0))
End Sub

' =====================================================================
' Private ─ COM ヘルパー
' =====================================================================

' AddRef（参照カウンタ増加）
Private Sub COM_AddRef(ByVal pObj As LongPtr)
    Dim retV As Variant
    DispCallFunc pObj, 1 * PTR_SIZE, CC_STDCALL, vbLong, 0, 0, 0, retV
End Sub

' Navigate(url: LPCWSTR)
Private Function COM_Navigate(ByVal pWV2 As LongPtr, ByVal url As String) As Long
    Dim pUrl As LongPtr: pUrl = StrPtr(url)
    Dim argT(0) As Integer: argT(0) = vbLongLong
    Dim argP(0) As LongPtr: argP(0) = VarPtr(pUrl)
    Dim retV As Variant
    DispCallFunc pWV2, OFF_WV2_Navigate, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
    COM_Navigate = CLng(retV)
End Function

' put_Bounds(RECT)  ─ x64 ABI では 16バイト構造体はポインタ渡し
Private Sub COM_SetBounds(ByVal pCtrl As LongPtr, ByRef bounds As RECT_T)
    Dim pBounds As LongPtr: pBounds = VarPtr(bounds)
    Dim argT(0) As Integer: argT(0) = vbLongLong
    Dim argP(0) As LongPtr: argP(0) = VarPtr(pBounds)
    Dim retV As Variant
    DispCallFunc pCtrl, OFF_CTRL_put_Bounds, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
End Sub

' 引数が LongPtr 1個の汎用DispCallFunc
Private Sub COM_DispCallPtr1(ByVal pObj As LongPtr, ByVal offset As Long, ByVal arg0 As LongPtr)
    Dim pArg As LongPtr: pArg = arg0
    Dim argT(0) As Integer: argT(0) = vbLongLong
    Dim argP(0) As LongPtr: argP(0) = VarPtr(pArg)
    Dim retV As Variant
    DispCallFunc pObj, offset, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
End Sub

' 引数が Long 1個の汎用DispCallFunc（put_IsVisible など）
Private Sub COM_DispCallLong1(ByVal pObj As LongPtr, ByVal offset As Long, ByVal arg0 As Long)
    Dim argT(0) As Integer: argT(0) = vbLong
    Dim argP(0) As LongPtr: argP(0) = VarPtr(arg0)
    Dim retV As Variant
    DispCallFunc pObj, offset, CC_STDCALL, vbLong, 1, argT(0), argP(0), retV
End Sub

' =====================================================================
' Private ─ WebView2Loader.dll の検索とロード
' =====================================================================
Private Sub EnsureWebView2LoaderDll()
    ' Edge インストール先から WebView2Loader.dll をロードする
    ' 既に PATH に入っていれば CreateCoreWebView2EnvironmentWithOptions がそのまま動く。
    ' 入っていない場合、以下の典型的パスを順に試みる。
    Dim candidates(2) As String
    candidates(0) = "C:\Program Files (x86)\Microsoft\EdgeWebView\Application\"
    candidates(1) = "C:\Program Files\Microsoft\EdgeWebView\Application\"
    candidates(2) = Environ("LOCALAPPDATA") & "\Microsoft\Edge\Application\"

    Dim i As Integer
    For i = 0 To UBound(candidates)
        Dim dllPath As String
        dllPath = FindLatestEdgeDll(candidates(i))
        If Len(dllPath) > 0 Then
            LoadLibrary StrPtr(dllPath)
            Exit For
        End If
    Next i
End Sub

Private Function FindLatestEdgeDll(ByVal basePath As String) As String
    ' basePath\<バージョン>\WebView2Loader.dll を Dir で探す（最初に見つかったもの）
    If Len(basePath) = 0 Then Exit Function
    Dim ver As String: ver = Dir(basePath, vbDirectory)
    Do While Len(ver) > 0
        If ver <> "." And ver <> ".." Then
            Dim candidate As String
            candidate = basePath & ver & "\WebView2Loader.dll"
            If Len(Dir(candidate)) > 0 Then
                FindLatestEdgeDll = candidate
                Exit Function
            End If
        End If
        ver = Dir()
    Loop
End Function

' =====================================================================
' Class lifecycle
' =====================================================================
Private Sub Class_Initialize()
    m_pEnvironment  = 0
    m_pController   = 0
    m_pWebView2     = 0
    m_LastErrorCode = 0
    m_State         = 0
End Sub

Private Sub Class_Terminate()
    Quit
End Sub
