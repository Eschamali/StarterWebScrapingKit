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
'   ホスト用の親ウィンドウは、利用者が使いたいハンドルを指定する方式です。
'
'   ★注意(既知の落とし穴、継承不可避)★
'     コールバック待ち中(このClassのInit実行中?完了まで)にVBEでブレーク/ステップ
'     実行すると、Excelがクラッシュする可能性があります(サンク技法そのものの制約)。
'***************************************************************************************************
Option Explicit



#If Win64 Then
    ' OK
#Else
    #Error "このクラスは64ビットVBA(x64)が必要です"
#End If



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

'DJP 20241105 to solve problems with terminating sessions on terminal server
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" ( _
    ) As Long

Private Declare PtrSafe Function ProcessIdToSessionId Lib "kernel32" ( _
    ByVal dwProcessId As Long, _
    ByRef pSessionId As Long) As Long



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Public Enum WebView2HostState
    Whs_Created = 0
    Whs_Waiting = 1
    Whs_Ready = 2
    Whs_Failed = 3
End Enum


'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
Private m_hostHwnd      As LongPtr
Private m_pEnvironment  As LongPtr
Private m_pController   As LongPtr
Private m_pView         As LongPtr
Private m_pSettings     As LongPtr
Private m_settingsLevel As Long   ' `m_pSettings`が実際にQueryInterfaceで辿り着けた`ICoreWebView2SettingsN`のN(基底のみなら1)
Private m_controllerLevel As Long ' `m_pController`が実際にQueryInterfaceで辿り着けた`ICoreWebView2ControllerN`のN(基底のみなら1)

Private m_envHandler    As CDPWebView2CallbackHandler
Private m_ctrlHandler   As CDPWebView2CallbackHandler

Private m_envState      As WebView2HostState
Private m_ctrlState     As WebView2HostState

Private m_lastError     As Long
Private m_closed        As Boolean

Private WebView2UserDataDirPath As String



'***************************************************************************************************
'                                   ■■■ 起動前準備 ■■■
'***************************************************************************************************
'* 機能　　：このClassが機能する必要な情報を渡します
'---------------------------------------------------------------------------------------------------
'* 引数    ：UserName   WebView2のユーザーデータフォルダ名
'            AttachHwnd WebView2を貼り付けるハンドル
'***************************************************************************************************
Public Sub InitInfo(UserName As String, AttachHwnd As LongPtr)
    WebView2UserDataDirPath = Environ("LOCALAPPDATA") & "\Microsoft\Office\16.0\Wef\webview2\VBA\" & UserName   'Officeのルールに従うようなパスにしておく
    m_hostHwnd = AttachHwnd
End Sub

'***************************************************************************************************
'* 機能　　：ICoreWebView2Environment→ICoreWebView2Controller→ICoreWebView2の生成を、
'            完了まで同期的に待って実行します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功可否
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に`InitInfo`にて、情報を渡すこと
'***************************************************************************************************
Public Function RunWebView2() As Boolean
    Const FromProcedureName As String = "CDPWebView2Host.RunWebView2"


    '1. 2重インスタンス保持はさせない
    If m_envState <> Whs_Created Then
        Debug.Print FromProcedureName & ": 二重呼び出しは不可です"
        Exit Function
    End If

    '2. Environment作成
    Set m_envHandler = CDPWebView2Thunks.AcquireHandlerFor(HK_EnvironmentCompleted, Me)
    If m_envHandler Is Nothing Then
        Debug.Print FromProcedureName & ": Environment用ハンドラのAcquireに失敗しました"
        m_envState = Whs_Failed
        Exit Function
    End If

    Dim pEnvHandlerObj As LongPtr
    pEnvHandlerObj = m_envHandler.Slot + CDPWebView2Thunks.VTABLE_OBJ_OFFSET

    '3. 重要: 状態をAPI呼び出しの「前」にWaitingへ。WebView2Loaderが内部で同期的に
    'コールバックを発火させるケースがあるため(実機で確認済みの罠)
    m_envState = Whs_Waiting

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(WebView2UserDataDirPath), 0, pEnvHandlerObj)
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

    '6. ICoreWebView2Settingsを取得(View::get_Settings、vtable index 3)
    '   ★毎回取得/解放していると呼び出しごとにQI相当のオーバーヘッドが発生するため、
    '   ここで1回だけ取得してキャッシュし、`CLOSE_`でまとめて解放する方式に統一する。
    '   取得に失敗しても(古いWebView2 Runtime等)Init自体は成功扱いとし、`m_pSettings=0`の
    '   ままにする(DevToolsEnabled等の各プロパティは0ガードで静かに無視するだけなので安全)。
    Dim ps As LongPtr
    hr = CDPWebView2Thunks.dcf(m_pView, 3, "View.get_Settings", VarPtr(ps))
    If hr = 0 And ps <> 0 Then
        m_pSettings = ps
        m_settingsLevel = 1
        DetectSettingsLevel
    End If

    '7. 成功サインを返す
    RunWebView2 = True
End Function

'***************************************************************************************************
'* 機能　　：`m_pSettings`が実際にQueryInterfaceで辿り着ける`ICoreWebView2SettingsN`の最大のNを
'            判定し、`m_pSettings`/`m_settingsLevel`を更新します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`ICoreWebView2Settings2`~`9`は(`EnvironmentOptions2`~`8`のような並列独立継承ではなく)
'            `Settings2 : Settings`→`Settings3 : Settings2`→…という一本の継承チェーンなので、
'            1個のCOMオブジェクトは実質1枚の(どんどん後ろへ延びていく)vtableしか持たない。
'            そのため、対応している最大のバージョンへ1回QueryInterceすれば、そのポインタ1本だけで
'            基底からそのバージョンまでの全プロパティに(スロット番号さえ合わせれば)アクセスできる。
'            高い方から順に試し、最初に成功したものを採用する(全滅時は基底のまま=level1)。
'***************************************************************************************************
Private Sub DetectSettingsLevel()
    Dim iids(1 To 8) As String
    iids(1) = "0528a73b-e92d-49f4-927a-e547dddaa37d"   ' ICoreWebView2Settings9
    iids(2) = "9e6b0e8f-86ad-4e81-8147-a9b5edb68650"   ' ICoreWebView2Settings8
    iids(3) = "488dc902-35ef-42d2-bc7d-94b65c4bc49c"   ' ICoreWebView2Settings7
    iids(4) = "11cb3acd-9bc8-43b8-83bf-f40753714f87"   ' ICoreWebView2Settings6
    iids(5) = "183e7052-1d03-43a0-ab99-98e043b66b39"   ' ICoreWebView2Settings5
    iids(6) = "cb56846c-4168-4d53-b04f-03b6d6796ff2"   ' ICoreWebView2Settings4
    iids(7) = "fdb5ab74-af33-4854-84f0-0a631deb5eba"   ' ICoreWebView2Settings3
    iids(8) = "ee9a0f68-f46c-4e32-ac23-ef8cac224d2a"   ' ICoreWebView2Settings2

    Dim levels(1 To 8) As Long
    levels(1) = 9: levels(2) = 8: levels(3) = 7: levels(4) = 6
    levels(5) = 5: levels(6) = 4: levels(7) = 3: levels(8) = 2

    Dim i As Long, iid As CDPWebView2Thunks.GUID, pTry As LongPtr, hrQ As Long
    For i = 1 To 8
        CDPWebView2Thunks.FillGUID iid, iids(i)
        hrQ = CDPWebView2Thunks.dcf(m_pSettings, 0, "Settings.QueryInterface", VarPtr(iid), VarPtr(pTry))
        If hrQ = 0 And pTry <> 0 Then
            CDPWebView2Thunks.ComRelease m_pSettings   ' 基底(または下位)の分の参照は不要になるので解放
            m_pSettings = pTry
            m_settingsLevel = levels(i)
            Exit Sub
        End If
    Next i
End Sub

'***************************************************************************************************
'* 機能　　：`m_pController`が実際にQueryInterfaceで辿り着ける`ICoreWebView2ControllerN`の最大の
'            Nを判定し、`m_pController`/`m_controllerLevel`を更新します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`ICoreWebView2Controller2`~`4`も`Settings`系と同様、一本の継承チェーン
'            (`Controller2 : Controller`→`Controller3 : Controller2`→…)なので、対応してる
'            最大のバージョンへ1回QueryInterceすれば、そのポインタ1本だけで基底からそのバージョン
'            までの全プロパティ/メソッドに(スロット番号さえ合わせれば)アクセスできる
'***************************************************************************************************
Private Sub DetectControllerLevel()
    Dim iids(1 To 3) As String
    iids(1) = "97d418d5-a426-4e49-a151-e1a10f327d9e"   ' ICoreWebView2Controller4
    iids(2) = "f9614724-5d2b-41dc-aef7-73d62b51543b"   ' ICoreWebView2Controller3
    iids(3) = "c979903e-d4ca-4228-92eb-47ee3fa96eab"   ' ICoreWebView2Controller2

    Dim levels(1 To 3) As Long
    levels(1) = 4: levels(2) = 3: levels(3) = 2

    Dim i As Long, iid As CDPWebView2Thunks.GUID, pTry As LongPtr, hrQ As Long
    For i = 1 To 3
        CDPWebView2Thunks.FillGUID iid, iids(i)
        hrQ = CDPWebView2Thunks.dcf(m_pController, 0, "Ctrl.QueryInterface", VarPtr(iid), VarPtr(pTry))
        If hrQ = 0 And pTry <> 0 Then
            CDPWebView2Thunks.ComRelease m_pController   ' 基底(または下位)の分の参照は不要になるので解放
            m_pController = pTry
            m_controllerLevel = levels(i)
            Exit Sub
        End If
    Next i
    m_controllerLevel = 1
End Sub

'* 機能　　：`m_envState`/`m_ctrlState`がWhs_Waitingを脱するまでDoEventsで待ちます
Private Function WaitFor(ByRef state As WebView2HostState) As Boolean
    Const MaxTick As Long = 100 '無難に10.0s待機


    Dim tick As Long
    Do While state = Whs_Waiting And tick < MaxTick
        DoEvents
        CDPHelpers.Sleep 0.1
        tick = tick + 1
    Loop
    WaitFor = (state <> Whs_Waiting)
End Function



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
DetectControllerLevel
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
    If m_pSettings <> 0 Then
        CDPWebView2Thunks.ComRelease m_pSettings
        m_pSettings = 0
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
'* 機能　　：新しいWebView2セッションを開始する前にクリーンアップを行います
'---------------------------------------------------------------------------------------------------
'* 機能説明：・基本は対象WebView2の`user-data-dir`に属するプロセスのみを終了します
'            ・コマンドライン引数に加えて、ターミナルサーバー上で (このユーザーに属さない) セッションを閉じる際の問題も解決しています
'***************************************************************************************************
Public Sub cleanUpSessions()
    '1. 今、Excelが動いてるプロセスIDとそのログインユーザーのセッションIDを取得します
    Dim lProcessId As Long, lSessionId As Long
    lProcessId = GetCurrentProcessId()          'Get the current(Excel) process ID
    ProcessIdToSessionId lProcessId, lSessionId 'Get the Session ID for the current(Excel) process

    '2. WebView2のexe名 ※1つしかないので固定値
    Const exeName As String = "msedgewebview2.exe"

    '3. 起動中のプロセスを巡る
    Dim proc As Object
    On Error Resume Next
    For Each proc In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process Where SessionId = " & lSessionId)
        '4. 起動中のブラウザ名であるか確認
        If LCase$(proc.Name) = LCase$(exeName) Then
            '5. Killタイプ判定
            If InStr(1, proc.CommandLine, "--user-data-dir=""" & getUserDataDirPath & Chr(34)) > 0 Then
                '6. 対象ユーザー名のWebView2なら、Killします
                proc.Terminate
            End If
        End If
        DoEvents
    Next
    On Error GoTo 0
    Err.Clear

End Sub



'***************************************************************************************************
'                                       ■■■ プロパティ群 ■■■
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

Public Property Get LastErrorMessage() As String
    LastErrorMessage = WinApiError.GetMessage(m_lastError, "WebView2Loader")
End Property

'***************************************************************************************************
'* 機能　　：実際の`--user-data-dir`情報を返します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：WebView2に渡す情報は`WebView2UserDataDirPath`ですが実際に作られるパスはその末尾に"\EBWebView"が付与されるためその補正を考慮します
'***************************************************************************************************
Public Property Get getUserDataDirPath() As String
    getUserDataDirPath = WebView2UserDataDirPath & "\EBWebView"
End Property



'***************************************************************************************************
'                               ■■■ ICoreWebView2Controller 一式 ■■■
'***************************************************************************************************
'* 機能　　：リサイズ処理を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：WV2_Right      横の長さ
'            WV2_Bottom     縦の長さ
'            WV2_Top        上端からのオフセット値
'            WV2_Left       左端からのオフセット値
'***************************************************************************************************
Public Sub Resize(WV2_Right As Long, WV2_Bottom As Long, WV2_Top As Long, WV2_Left As Long)
    '1. ひとまとまりにする
    Dim rc As CDPWebView2Thunks.RECT
    rc.Left = WV2_Left: rc.Top = WV2_Top: rc.Right = WV2_Right: rc.Bottom = WV2_Bottom

    '2. リサイズ
    CDPWebView2Thunks.dcf m_pController, 6, "Ctrl.put_Bounds", VarPtr(rc)       ' put_Bounds
End Sub

'***************************************************************************************************
'* 機能　　：表示/非表示切り替え
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Web描画の表示/非表示
'***************************************************************************************************
Property Let Visible(v As Boolean)
    CDPWebView2Thunks.dcf m_pController, 4, "Ctrl.put_IsVisible", CLng(v * -1)  ' put_IsVisible
End Property

'***************************************************************************************************
'* 機能　　：ブラウザUIレベルのズーム倍率を設定します(CDPの`Emulation.setPageScaleFactor`とは別物)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  ズーム倍率(1.0が等倍)
'***************************************************************************************************
Property Let ZoomFactor(v As Double)
    CDPWebView2Thunks.dcf m_pController, 8, "Ctrl.put_ZoomFactor", v
End Property

'***************************************************************************************************
'* 機能　　：表示位置/サイズとズーム倍率を、1回のレイアウト計算でまとめて設定します
'---------------------------------------------------------------------------------------------------
'* 引数    ：WV2_Right/WV2_Bottom/WV2_Top/WV2_Left  `Resize`と同じ
'            zoomFactor                              ズーム倍率(1.0が等倍)
'---------------------------------------------------------------------------------------------------
'* 注意事項：`Resize`+`ZoomFactor`を別々に呼ぶより、こちらでまとめた方が描画の無駄が少ない
'***************************************************************************************************
Public Sub SetBoundsAndZoomFactor(WV2_Right As Long, WV2_Bottom As Long, WV2_Top As Long, WV2_Left As Long, ZoomFactor As Double)
    Dim rc As CDPWebView2Thunks.RECT
    rc.Left = WV2_Left: rc.Top = WV2_Top: rc.Right = WV2_Right: rc.Bottom = WV2_Bottom
    CDPWebView2Thunks.dcf m_pController, 11, "Ctrl.SetBoundsAndZoomFactor", VarPtr(rc), ZoomFactor
End Sub

'***************************************************************************************************
'* 機能　　：フォーカスをWebView2側へプログラム的に移動します
'---------------------------------------------------------------------------------------------------
'* 引数    ：reason  `COREWEBVIEW2_MOVE_FOCUS_REASON`(Long。0=Programmatic, 1=Next, 2=Previous。
'            `CDPCoreViaWebView2`側の`WV2MoveFocusReason`Enum使用を推奨)
'***************************************************************************************************
Public Sub MoveFocus(reason As Long)
    CDPWebView2Thunks.dcf m_pController, 12, "Ctrl.MoveFocus", reason
End Sub

'***************************************************************************************************
'* 機能　　：親ウィンドウの位置が動いたことをWebView2へ明示的に通知します
'---------------------------------------------------------------------------------------------------
'* 注意事項：ポップアップ/ドロップダウン等、親ウィンドウの画面上の絶対位置に依存する描画がある場合に
'            呼ぶためのものです。このプロジェクトの非表示運用では基本的に不要です
'***************************************************************************************************
Public Sub NotifyParentWindowPositionChanged()
    CDPWebView2Thunks.dcf m_pController, 23, "Ctrl.NotifyParentWindowPositionChanged"
End Sub

'***************************************************************************************************
'* 機能　　：WebView2のページ読み込み前に表示される既定の背景色を設定します(Controller2以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：a/r/g/b  各0~255。`a`(アルファ値)は現状255(不透明)のみサポート
'---------------------------------------------------------------------------------------------------
'* 注意事項：`COREWEBVIEW2_COLOR`はA/R/G/B各1バイトの4バイト構造体で、x64呼び出し規約上
'            ポインタ経由ではなく32bit整数1個としてレジスタに直接乗る(メモリ上A,R,G,Bの順に
'            並ぶリトルエンディアンの4バイト値として組み立てて渡す)。4引数全てが対等な色成分で
'            あり「代入値」の形にそぐわないため、`Property Let`ではなく`Sub`にしている
'***************************************************************************************************
Public Sub SetDefaultBackgroundColor(a As Byte, r As Byte, g As Byte, b As Byte)
    If m_controllerLevel < 2 Then Exit Sub
    Dim packed As Long
    packed = CLng(a) Or (CLng(r) * &H100) Or (CLng(g) * &H10000) Or (CLng(b) * &H1000000)
    CDPWebView2Thunks.dcf m_pController, 27, "Ctrl2.put_DefaultBackgroundColor", packed
End Sub

'***************************************************************************************************
'* 機能　　：DPIスケールを手動指定します(Controller3以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  DPIスケール値(例:1.0=100%、1.5=150%)
'***************************************************************************************************
Property Let RasterizationScale(v As Double)
    If m_controllerLevel < 3 Then Exit Property
    CDPWebView2Thunks.dcf m_pController, 29, "Ctrl3.put_RasterizationScale", v
End Property

'***************************************************************************************************
'* 機能　　：モニターのDPIスケール変更を自動検知するか切り替えます(Controller3以降、既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで自動検知、Falseで`RasterizationScale`を手動管理する
'***************************************************************************************************
Property Let ShouldDetectMonitorScaleChanges(v As Boolean)
    If m_controllerLevel < 3 Then Exit Property
    CDPWebView2Thunks.dcf m_pController, 31, "Ctrl3.put_ShouldDetectMonitorScaleChanges", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：`Bounds`(位置/サイズ)の解釈方式を切り替えます(Controller3以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  `COREWEBVIEW2_BOUNDS_MODE`(Long。0=USE_RAW_PIXELS(既定)、1=USE_RASTERIZATION_SCALE。
'            `CDPCoreViaWebView2`側の`WV2BoundsMode`Enum使用を推奨)
'***************************************************************************************************
Property Let BoundsMode(v As Long)
    If m_controllerLevel < 3 Then Exit Property
    CDPWebView2Thunks.dcf m_pController, 35, "Ctrl3.put_BoundsMode", v
End Property

'***************************************************************************************************
'* 機能　　：外部(エクスプローラー等)からのドラッグ&ドロップ受け入れ可否を切り替えます
'            (Controller4以降、既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let AllowExternalDrop(v As Boolean)
    If m_controllerLevel < 4 Then Exit Property
    CDPWebView2Thunks.dcf m_pController, 37, "Ctrl4.put_AllowExternalDrop", CLng(v * -1)
End Property



'***************************************************************************************************
'                               ■■■ ICoreWebView2Settings 一式 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Settings
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：DevTools(F12キー・右クリック「検証」から起動するWebView2内蔵のDevToolsウィンドウ)の
'            起動可否を切り替えます
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'---------------------------------------------------------------------------------------------------
'* 注意事項：`ICoreWebView2Settings::put_AreDevToolsEnabled`はブラウザ上のDevTools UIのみを
'            制御するものであり、このプロジェクトが依存するCDPリモートデバッグ
'            (`--remote-debugging-pipe`)経由の自動操作には一切影響しません
'***************************************************************************************************
Property Let DevToolsEnabled(v As Boolean)
    If m_pSettings = 0 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 12, "Settings.put_AreDevToolsEnabled", CLng(v * -1)   ' put_AreDevToolsEnabled
End Property

'***************************************************************************************************
'* 機能　　：既定の右クリックコンテキストメニューの表示可否を切り替えます
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let ContextMenuEnabled(v As Boolean)
    If m_pSettings = 0 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 14, "Settings.put_AreDefaultContextMenusEnabled", CLng(v * -1)   ' put_AreDefaultContextMenusEnabled
End Property

'--- 以下、`ICoreWebView2Settings`(基底、level1)の残りのプロパティ ---
'***************************************************************************************************
'* 機能　　：ページ内のJavaScript実行可否を切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let ScriptEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 4, "Settings.put_IsScriptEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：`postMessage`/`WebMessageReceived`によるWebページ⇔ホストアプリ間の通信可否を
'            切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let WebMessageEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 6, "Settings.put_IsWebMessageEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：`alert`/`confirm`/`prompt`/離脱確認等、既定のJavaScriptダイアログ表示可否を
'            切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'---------------------------------------------------------------------------------------------------
'* 注意事項：Falseにすると、ダイアログは表示されず既定値で即座に処理されます
'***************************************************************************************************
Property Let DefaultScriptDialogsEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 8, "Settings.put_AreDefaultScriptDialogsEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：リンクにマウスオーバーした際のURLプレビュー等、ステータスバーの表示可否を
'            切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let StatusBarEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 10, "Settings.put_IsStatusBarEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：`AddHostObjectToScript`で追加したホストオブジェクトへ、スクリプト側からアクセスできるか
'            切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let HostObjectsAllowed(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 16, "Settings.put_AreHostObjectsAllowed", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：Ctrl+ホイール/ピンチ/キーボードショートカット等によるユーザー操作でのズーム可否を
'            切り替えます(既定True)
'***************************************************************************************************
Property Let ZoomControlEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 18, "Settings.put_IsZoomControlEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：ページ読み込み失敗時等に表示される、WebView2内蔵のエラーページ表示可否を
'            切り替えます(既定True)
'***************************************************************************************************
Property Let BuiltInErrorPageEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 20, "Settings.put_IsBuiltInErrorPageEnabled", CLng(v * -1)
End Property

'--- 以下、`ICoreWebView2Settings2`~`9`(古いWebView2 Runtimeでは`m_settingsLevel`不足によりNo-op) ---
'***************************************************************************************************
'* 機能　　：リクエストヘッダーおよび`navigator.userAgent`に使われるUser-Agent文字列を上書きします
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  上書きするUser-Agent文字列
'---------------------------------------------------------------------------------------------------
'* 注意事項：`ICoreWebView2EnvironmentOptions2`とは別物(こちらは実行時にいつでも変更可能)
'***************************************************************************************************
Property Let UserAgentOverride(v As String)
    If m_settingsLevel < 2 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 22, "Settings2.put_UserAgent", StrPtr(v)
End Property

'***************************************************************************************************
'* 機能　　：`Ctrl+F`/`F3`/`F5`/`F6`/`F7`/`F12`/`Ctrl+P`等、ブラウザ的なアクセラレータキーの
'            有効可否を切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'---------------------------------------------------------------------------------------------------
'* 注意事項：`DevToolsEnabled=False`と組み合わせても、`F12`キー自体を無効化したい場合はこちらも
'            合わせて`False`にする必要があります(`DevToolsEnabled`はDevTools機能自体のON/OFF、
'            こちらはキー入力の割り当て自体のON/OFF)
'***************************************************************************************************
Property Let BrowserAcceleratorKeysEnabled(v As Boolean)
    If m_settingsLevel < 3 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 24, "Settings3.put_AreBrowserAcceleratorKeysEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：パスワードの自動保存機能(保存確認ダイアログ含む)の有効可否を切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let PasswordAutosaveEnabled(v As Boolean)
    If m_settingsLevel < 4 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 26, "Settings4.put_IsPasswordAutosaveEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：住所/フォーム入力内容等の一般的なオートフィル機能(パスワード以外)の有効可否を
'            切り替えます(既定True)
'***************************************************************************************************
Property Let GeneralAutofillEnabled(v As Boolean)
    If m_settingsLevel < 4 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 28, "Settings4.put_IsGeneralAutofillEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：タッチ操作によるピンチズームジェスチャーの有効可否を切り替えます(既定True)
'***************************************************************************************************
Property Let PinchZoomEnabled(v As Boolean)
    If m_settingsLevel < 5 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 30, "Settings5.put_IsPinchZoomEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：左右スワイプによる進む/戻るナビゲーションジェスチャーの有効可否を切り替えます(既定True)
'***************************************************************************************************
Property Let SwipeNavigationEnabled(v As Boolean)
    If m_settingsLevel < 6 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 32, "Settings6.put_IsSwipeNavigationEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：PDFビューア内蔵ツールバーの、非表示にしたいボタンをビットマスクで指定します
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  `COREWEBVIEW2_PDF_TOOLBAR_ITEMS`のビットマスク(Long)。既定は`0`(全表示)
'            例: `SAVE(&H1) Or PRINT(&H2)`で保存・印刷ボタンを非表示にする、等
'            (SAVE=&H1, PRINT=&H2, SAVE_AS=&H4, ZOOM_IN=&H8, ZOOM_OUT=&H10, ROTATE=&H20,
'             FIT_PAGE=&H40, PAGE_LAYOUT=&H80, BOOKMARKS=&H100, PAGE_SELECTOR=&H200,
'             SEARCH=&H400, FULL_SCREEN=&H800, MORE_SETTINGS=&H1000)
'***************************************************************************************************
Property Let HiddenPdfToolbarItems(v As Long)
    If m_settingsLevel < 7 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 34, "Settings7.put_HiddenPdfToolbarItems", v
End Property

'***************************************************************************************************
'* 機能　　：SmartScreen(フィッシング/マルウェア報告済みサイトの検知)の要否を切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで有効、Falseで無効
'---------------------------------------------------------------------------------------------------
'* 注意事項：同じ`user-data-dir`を共有する全`ICoreWebView2`にまたがる設定です。いずれか1つでも
'            Trueなら全体でSmartScreenが有効になり、全てFalseの場合のみ無効になります
'***************************************************************************************************
Property Let ReputationCheckingRequired(v As Boolean)
    If m_settingsLevel < 8 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 36, "Settings8.put_IsReputationCheckingRequired", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：Webページ側でCSSの`app-region: drag`/`no-drag`を使った独自タイトルバー領域(ドラッグで
'            ウィンドウ自体を移動できる領域)の指定を有効にするか切り替えます(既定False)
'***************************************************************************************************
Property Let NonClientRegionSupportEnabled(v As Boolean)
    If m_settingsLevel < 9 Then Exit Property
    CDPWebView2Thunks.dcf m_pSettings, 38, "Settings9.put_IsNonClientRegionSupportEnabled", CLng(v * -1)
End Property


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
