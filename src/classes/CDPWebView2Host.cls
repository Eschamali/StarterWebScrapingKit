VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPWebView2Host"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   WebView2のEnvironment→Controller→ICoreWebView2生成だけを担うホスティング専用Classです。
'
'   `CDPCoreViaWebView2.cls`とは責務を分離しています(WebSocket版が1ハンドルで済むのに対し、
'   WebView2はForm/Env/Controller/Viewという生成順序・Releaseチェーンが伴うため、
'   CDP送受信ロジックと混ぜるとClass_Terminateの順序管理が事故りやすいためです)。
'
'   ホスト用の親ウィンドウは、UserForm(.frm/.frx)を使わず、システム標準の"Static"
'   ウィンドウクラスで非表示ネイティブウィンドウを都度生成する方式を採用しています
'   (このプロジェクトはCDP専用でWebView2を可視化する必要がなく、.frxバイナリを
'   手作業で正しく生成するのは事故のリスクが高いため)。
'
'   ★注意(既知の落とし穴、継承不可避)★
'     コールバック待ち中(このClassのInit実行中?完了まで)にVBEでブレーク/ステップ
'     実行すると、Excelがクラッシュする可能性があります(サンク技法そのものの制約)。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
' --- WebView2Loader.dll ---
'   Excel(Power Query)に同梱されている実物を`EnsureWebView2LoaderResolved`で
'   事前にLoadLibraryしておくことで、この`Declare`はその既読み込み済みモジュールを
'   ベース名で再利用する(Win32ローダーの標準動作)。
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)



#If Win64 Then
    ' OK
#Else
    #Error "このクラスは64ビットVBA(x64)が必要です"
#End If

Public Enum WebView2HostState
    Whs_Created = 0
    Whs_Waiting = 1
    Whs_Ready = 2
    Whs_Failed = 3
End Enum

Private Const ClassNameStatic As String = "Static"

Private m_hostHwnd     As LongPtr
Private m_pEnvironment As LongPtr
Private m_pController  As LongPtr
Private m_pView        As LongPtr

Private m_envHandler  As CDPWebView2CallbackHandler
Private m_ctrlHandler As CDPWebView2CallbackHandler

Private m_envState  As WebView2HostState
Private m_ctrlState As WebView2HostState
Private m_lastError As Long
Private m_closed    As Boolean



'***************************************************************************************************
'                              ■■■ 初期化/後始末 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    m_envState = Whs_Created
    m_ctrlState = Whs_Created
End Sub

Private Sub Class_Terminate()
    If Not m_closed Then CLOSE_
End Sub

'***************************************************************************************************
'* 機能　　：ICoreWebView2Environment→ICoreWebView2Controller→ICoreWebView2の生成を、
'            完了まで同期的に待って実行します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功可否
'* 引数　　：userDataPath   WebView2のユーザーデータフォルダ(省略時は%TEMP%\CDPWebView2)
'***************************************************************************************************
Public Function Init(Optional userDataPath As String) As Boolean
    Const FromProcedureName As String = "CDPWebView2Host.Init"

    If m_envState <> Whs_Created Then
        Debug.Print FromProcedureName & ": 二重呼び出しは不可です"
        Exit Function
    End If

    '1. WebView2Loader.dllの解決
    If Not CDPWebView2Thunks.EnsureWebView2LoaderResolved() Then
        Debug.Print FromProcedureName & ": WebView2Loader.dllが見つかりませんでした。" & _
                    "Power Query for Excelが導入されているか確認してください。"
        m_envState = Whs_Failed
        Exit Function
    End If

    '2. Excelのハンドルを直接使用する
    m_hostHwnd = Application.hWnd

    '3. Environment作成
    Dim Path As String
    Path = userDataPath
    If LenB(Path) = 0 Then Path = Environ$("TEMP") & "\CDPWebView2"
    EnsureFolderExists Path

    Set m_envHandler = CDPWebView2Thunks.AcquireHandlerFor(HK_EnvironmentCompleted, Me)
    If m_envHandler Is Nothing Then
        Debug.Print FromProcedureName & ": Environment用ハンドラのAcquireに失敗しました"
        m_envState = Whs_Failed
        Exit Function
    End If

    Dim pEnvHandlerObj As LongPtr
    pEnvHandlerObj = m_envHandler.Slot + CDPWebView2Thunks.VTABLE_OBJ_OFFSET

    '重要: 状態をAPI呼び出しの「前」にWaitingへ。WebView2Loaderが内部で同期的に
    'コールバックを発火させるケースがあるため(実機で確認済みの罠)
    m_envState = Whs_Waiting

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(Path), 0, pEnvHandlerObj)
    If hr <> 0 Then
        Debug.Print FromProcedureName & ": CreateCoreWebView2EnvironmentWithOptions失敗 hr=&H" & Hex(hr)
        m_envState = Whs_Failed
        m_lastError = hr
        Exit Function
    End If

    If Not WaitFor(m_envState) Then
        Debug.Print FromProcedureName & ": Environment完了待ちがタイムアウトしました"
        Exit Function
    End If
    If m_envState <> Whs_Ready Then Exit Function

    '4. Controller作成
    Set m_ctrlHandler = CDPWebView2Thunks.AcquireHandlerFor(HK_ControllerCompleted, Me)
    If m_ctrlHandler Is Nothing Then
        Debug.Print FromProcedureName & ": Controller用ハンドラのAcquireに失敗しました"
        m_ctrlState = Whs_Failed
        Exit Function
    End If

    Dim pCtrlHandlerObj As LongPtr
    pCtrlHandlerObj = m_ctrlHandler.Slot + CDPWebView2Thunks.VTABLE_OBJ_OFFSET

    m_ctrlState = Whs_Waiting

    hr = CDPWebView2Thunks.dcf(m_pEnvironment, 3, "Env.CreateCoreWebView2Controller", m_hostHwnd, pCtrlHandlerObj)
    If hr <> 0 Then
        Debug.Print FromProcedureName & ": CreateCoreWebView2Controller失敗 hr=&H" & Hex(hr)
        m_ctrlState = Whs_Failed
        m_lastError = hr
        Exit Function
    End If

    If Not WaitFor(m_ctrlState) Then
        Debug.Print FromProcedureName & ": Controller完了待ちがタイムアウトしました"
        Exit Function
    End If
    If m_ctrlState <> Whs_Ready Then Exit Function

    '5. ICoreWebView2を取得(Controller::get_CoreWebView2、vtable index 25)
    Dim pv As LongPtr
    hr = CDPWebView2Thunks.dcf(m_pController, 25, "Ctrl.get_CoreWebView2", VarPtr(pv))
    If hr <> 0 Or pv = 0 Then
        Debug.Print FromProcedureName & ": get_CoreWebView2失敗 hr=&H" & Hex(hr)
        m_lastError = hr
        Exit Function
    End If
    m_pView = pv

    '6. CDP専用のため、非表示・最小サイズのまま維持する
    Dim rc As CDPWebView2Thunks.RECT
    rc.Left = 0: rc.Top = 0: rc.Right = 4: rc.Bottom = 4
    CDPWebView2Thunks.dcf m_pController, 6, "Ctrl.put_Bounds", VarPtr(rc)          ' put_Bounds
    CDPWebView2Thunks.dcf m_pController, 4, "Ctrl.put_IsVisible", CLng(0)          ' put_IsVisible(False)

    Init = True
End Function

'* 機能　　：`m_envState`/`m_ctrlState`がWhs_Waitingを脱するまでDoEventsで待ちます
Private Function WaitFor(ByRef state As WebView2HostState) As Boolean
    Dim tick As Long
    Do While state = Whs_Waiting And tick < 1000    '無難に1s待機
        DoEvents
        sleep2 1
        tick = tick + 1
    Loop
    WaitFor = (state <> Whs_Waiting)
End Function

Private Sub EnsureFolderExists(ByVal Path As String)
    If LenB(Path) = 0 Then Exit Sub
    Dim parentDir As String
    Dim slashPos As Long
    slashPos = InStrRev(Path, "\")
    If slashPos > 0 Then
        parentDir = Left$(Path, slashPos - 1)
        If LenB(parentDir) > 0 Then
            If Dir(parentDir, vbDirectory) = "" Then MkDir parentDir
        End If
    End If
    If Dir(Path, vbDirectory) = "" Then MkDir Path
End Sub



'***************************************************************************************************
'                              ■■■ 完了コールバック(Public Sub必須) ■■■
'***************************************************************************************************
Public Sub OnEnvironmentCompleted(ByVal ErrorCode As Long, ByVal pEnvironment As LongPtr)
    m_lastError = ErrorCode
    If ErrorCode = 0 And pEnvironment <> 0 Then
        m_pEnvironment = pEnvironment
        CDPWebView2Thunks.ComAddRef m_pEnvironment
        m_envState = Whs_Ready
    Else
        m_pEnvironment = 0
        m_envState = Whs_Failed
    End If
    Set m_envHandler = Nothing   ' 1ショットハンドラパターン
End Sub

Public Sub OnControllerCompleted(ByVal ErrorCode As Long, ByVal pController As LongPtr)
    m_lastError = ErrorCode
    If ErrorCode = 0 And pController <> 0 Then
        m_pController = pController
        CDPWebView2Thunks.ComAddRef m_pController
        m_ctrlState = Whs_Ready
    Else
        m_pController = 0
        m_ctrlState = Whs_Failed
    End If
    Set m_ctrlHandler = Nothing  ' 1ショットハンドラパターン
End Sub



'***************************************************************************************************
'                              ■■■ 明示的な後始末 ■■■
'***************************************************************************************************
'* 機能　　：Controller::Close → 各種ポインタのRelease → ホストウィンドウの破棄を行います
'* 注意事項：ここでは`CDPWebView2Thunks.Thunks_Shutdown`は呼ばない。あちらはこのモジュール内で
'            全`CDPWebView2Host`インスタンス共通(グローバル)のスロットプールを解体するため、
'            他のインスタンスがまだ使用中の可能性がある間に呼ぶと事故になる
'***************************************************************************************************
Public Sub CLOSE_()
    If m_closed Then Exit Sub

    If m_pController <> 0 Then
        CDPWebView2Thunks.dcf m_pController, 24, "Ctrl.Close"
    End If
    If m_pView <> 0 Then
        CDPWebView2Thunks.ComRelease m_pView
        m_pView = 0
    End If
    If m_pController <> 0 Then
        CDPWebView2Thunks.ComRelease m_pController
        m_pController = 0
    End If
    If m_pEnvironment <> 0 Then
        CDPWebView2Thunks.ComRelease m_pEnvironment
        m_pEnvironment = 0
    End If

    m_closed = True
End Sub



'***************************************************************************************************
'                              ■■■ プロパティ群 ■■■
'***************************************************************************************************
Public Property Get ViewPtr() As LongPtr
    ViewPtr = m_pView
End Property

Public Property Get IsReady() As Boolean
    IsReady = (Not m_closed) And (m_pView <> 0)
End Property

Public Property Get LastError() As Long
    LastError = m_lastError
End Property
