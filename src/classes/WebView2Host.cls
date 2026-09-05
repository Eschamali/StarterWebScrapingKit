VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2Host"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   WebView2のEnvironment→Controller→ICoreWebView2 + α生成を担うホスティング専用Classです。
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
    ByVal EnvironmentOptions As LongPtr, _
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
'完了コールバック管理用
Private Enum WebView2HostState
    Whs_Created = 0
    Whs_Waiting = 1
    Whs_Ready = 2
    Whs_Failed = 3
End Enum



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'メイン制御ポインター
Private m_pView         As LongPtr  'ICoreWebView2
Private m_pController   As LongPtr  'ICoreWebView2Controller
Private m_pEnvironment  As LongPtr  'ICoreWebView2Environment
Private m_pSettings     As LongPtr  'ICoreWebView2Settings
Private m_pProfile      As LongPtr  'ICoreWebView2Profile

'たどり着けた階層レベル
Private m_viewLevel         As Long '`m_pView`      が実際にQueryInterfaceで辿り着けた`ICoreWebView2_N`のN          (基底のみなら1)
Private m_controllerLevel   As Long '`m_pController`が実際にQueryInterfaceで辿り着けた`ICoreWebView2ControllerN`のN (基底のみなら1)
Private m_environmentLevel  As Long '`m_pEnvironment`が実際にQueryInterfaceで辿り着けた`ICoreWebView2EnvironmentN`のN(基底のみなら1。_11までしか使わないためそこまでしか探らない)
Private m_settingsLevel     As Long '`m_pSettings`  が実際にQueryInterfaceで辿り着けた`ICoreWebView2SettingsN`のN   (基底のみなら1)
Private m_profileLevel      As Long '`m_pProfile`   が実際にQueryInterfaceで辿り着けた`ICoreWebView2ProfileN`のN    (基底のみなら1、未取得なら0)

'==============================================COM管理==============================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' [ SECTION ] 必須系COM一式定義
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'---------------------------------------------------------------------------------------------------
'* 機能　　：`ICoreWebView2Environment`の**作成**1回分の状態+ハンドラをまとめます
'---------------------------------------------------------------------------------------------------
'* 注意事項：結果は`m_pEnvironment`(別枠のメイン制御ポインター)に入るため、専用フィールドは無し
'            (`WebView2RemoveExtState`と同じ、state+handlerのみの型)
'---------------------------------------------------------------------------------------------------
Private Type WebView2EnvState
    state   As WebView2HostState        '完了待機
    handler As WebView2CallbackHandler
End Type

'---------------------------------------------------------------------------------------------------
'* 機能　　：`ICoreWebView2Controller`の**作成**1回分の状態+ハンドラをまとめます
'---------------------------------------------------------------------------------------------------
'* 注意事項：結果は`m_pController`(別枠のメイン制御ポインター)に入るため、専用フィールドは無し
'---------------------------------------------------------------------------------------------------
Private Type WebView2CtrlState
    state   As WebView2HostState        '完了待機
    handler As WebView2CallbackHandler
End Type

'---------------------------------------------------------------------------------------------------
'* 機能　　：WebView2生成に必須の非同期処理(Environment/Controller)を、操作(env/ctrl)ごとに
'            まとめます
'---------------------------------------------------------------------------------------------------
'* 注意事項：これが無いと何も始まらない、という意味で「必須機能」。`WebView2ExtensionAsyncState`と
'            同じ「操作ごとにstate+handler[+結果]をまとめる」構成に揃えている
'---------------------------------------------------------------------------------------------------
Private Type WebView2CoreAsyncState
    env  As WebView2EnvState
    ctrl As WebView2CtrlState
End Type

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' [ SECTION ] 拡張機能系COM一式定義
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'---------------------------------------------------------------------------------------------------
'* 機能　　：拡張機能の**インストール**(`AddBrowserExtension`)1回分の状態+ハンドラ+結果をまとめます
'---------------------------------------------------------------------------------------------------
Private Type WebView2AddExtState
    state    As WebView2HostState      '完了待機
    handler  As WebView2CallbackHandler
    resultId As String                 '結果(`OnAddBrowserExtensionCompleted`が書き込む)
End Type

'---------------------------------------------------------------------------------------------------
'* 機能　　：インストール済み拡張機能の**一覧取得**(`GetBrowserExtensions`)1回分の状態+ハンドラ+結果を
'            まとめます
'---------------------------------------------------------------------------------------------------
'* 注意事項：結果は「拡張機能ID(String) → ICoreWebView2BrowserExtensionへの新規参照(LongPtr)」の
'            Dictionaryで保持する(`RemoveBrowserExtension`が対象を特定した後、使い終わり次第
'            `ReleaseGetExtPtrs`で解放すること)
'---------------------------------------------------------------------------------------------------
Private Type WebView2GetExtState
    state   As WebView2HostState        '完了待機
    handler As WebView2CallbackHandler
    Items   As Dictionary               '結果(`OnGetBrowserExtensionsCompleted`が書き込む)
End Type

'---------------------------------------------------------------------------------------------------
'* 機能　　：拡張機能の**アンインストール**(`ICoreWebView2BrowserExtension::Remove`)1回分の
'            状態+ハンドラをまとめます(結果は成功可否[state]のみのため、専用フィールドは無し)
Private Type WebView2RemoveExtState
    state   As WebView2HostState        '完了待機
    handler As WebView2CallbackHandler
End Type

'* 機能　　：拡張機能管理(ICoreWebView2Profile7)関連の非同期処理を、操作(add/get/remove)ごとに
'            まとめます
'---------------------------------------------------------------------------------------------------
'* 注意事項：CDPの`Extensions.loadUnpacked`がWebView2越しには使えないための代替として、
'            後から追加した機能一式(拡張機能を一切使わないなら出番がない)
'---------------------------------------------------------------------------------------------------
Private Type WebView2ExtensionAsyncState
    addExt    As WebView2AddExtState
    getExt    As WebView2GetExtState
    removeExt As WebView2RemoveExtState
End Type

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' [ SECTION ] 使用するCOM一式変数
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private m_core As WebView2CoreAsyncState        '必須機能 ※これがないと始まらない
Private m_ext  As WebView2ExtensionAsyncState   '拡張機能系
'===================================================================================================

'状態管理
Private m_lastError     As Long
Private m_closed        As Boolean

'活動場所
Private m_hostHwnd              As LongPtr
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
'                                 ■■■ WebView2起動フロー ■■■
'***************************************************************************************************
'* 機能　　：ICoreWebView2Environment→ICoreWebView2Controller→ICoreWebView2の生成を、
'            完了まで同期的に待って実行します
'---------------------------------------------------------------------------------------------------
'* 引数    ：EnvironmentOptions  `ICoreWebView2EnvironmentOptions`系のネイティブfakeオブジェクトへの
'            ポインタ(`WebView2EnvOptions.BuildNativePointer`の戻り値)。未使用なら0を渡す
'* 返り値  ：成功可否
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に`InitInfo`にて、情報を渡すこと。`EnvironmentOptions`のネイティブリソース解放は
'            呼び出し元の責務(このFunctionが戻った時点で、WebView2Loaderは読み終えている)
'***************************************************************************************************
Public Function RunWebView2(ByVal EnvironmentOptions As LongPtr) As Boolean
    Const FromProcedureName As String = "WebView2Host.RunWebView2"


    '1. 2重インスタンス保持はさせない
    If m_core.env.state <> Whs_Created Then
        Debug.Print FromProcedureName & ": 二重呼び出しは不可です"
        Exit Function
    End If


    '--------------------------------- 必須コントロールポインター取得フェーズ ---------------------------------


    '2. `CreateCoreWebView2EnvironmentWithOptions`の第4引数に渡すための、`ICoreWebView2Environment`用COMコールバックオブジェクトを生成する
    Set m_core.env.handler = WebView2Thunks.AcquireHandlerFor(HK_EnvironmentCompleted, Me)
    If m_core.env.handler Is Nothing Then
        Debug.Print FromProcedureName & ": Environment用ハンドラのAcquireに失敗しました"
        m_core.env.state = Whs_Failed
        Exit Function
    End If

    Dim pEnvHandlerObj As LongPtr
    pEnvHandlerObj = m_core.env.handler.Slot + WebView2Thunks.VTABLE_OBJ_OFFSET

    '3. 重要: 状態をAPI呼び出しの「前」にWaitingへ。WebView2Loaderが内部で同期的に
    'コールバックを発火させるケースがあるため(実機で確認済みの罠)
    m_core.env.state = Whs_Waiting

    '4. 2で作った「COMコールバックオブジェクト」と事前準備した`EnvironmentOptions`ポインター、`WebView2UserDataDirPath`を渡して、WebView2を呼び出す
    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(WebView2UserDataDirPath), EnvironmentOptions, pEnvHandlerObj)
    If hr <> 0 Then
        Debug.Print FromProcedureName & ": CreateCoreWebView2EnvironmentWithOptions失敗 hr=&H" & Hex(hr)
        m_core.env.state = Whs_Failed
        m_lastError = hr
        Exit Function
    End If

    '5. `WebView2Loader`の処理が終わるまで、待機
    '※終了すると、COMコールバックにより`OnEnvironmentCompleted`プロシージャにワープします
    If Not WaitFor(m_core.env.state) Then
        Debug.Print FromProcedureName & ": Environment完了待ちがタイムアウトしました"
        Exit Function
    End If
    If m_core.env.state <> Whs_Ready Then Exit Function

    '5-1. ICoreWebView2Environmentのどの最新番号(_11まで)まで対応してるか探る
    DetectEnvironmentLevel

    '6. `ICoreWebView2Controller`用COMコールバックオブジェクトを生成する
    Set m_core.ctrl.handler = WebView2Thunks.AcquireHandlerFor(HK_ControllerCompleted, Me)
    If m_core.ctrl.handler Is Nothing Then
        Debug.Print FromProcedureName & ": Controller用ハンドラのAcquireに失敗しました"
        m_core.ctrl.state = Whs_Failed
        Exit Function
    End If

    Dim pCtrlHandlerObj As LongPtr
    pCtrlHandlerObj = m_core.ctrl.handler.Slot + WebView2Thunks.VTABLE_OBJ_OFFSET

    '7. 3と同様、処理前に待機ステータスをセット
    m_core.ctrl.state = Whs_Waiting

    '8. DispCallFunc経由で、`ICoreWebView2Controller`オブジェクトを作成依頼する
    hr = WebView2Thunks.dcf(m_pEnvironment, 3, "Env.CreateCoreWebView2Controller", m_hostHwnd, pCtrlHandlerObj)
    If hr <> 0 Then
        Debug.Print FromProcedureName & ": CreateCoreWebView2Controller失敗 hr=&H" & Hex(hr)
        m_core.ctrl.state = Whs_Failed
        m_lastError = hr
        Exit Function
    End If

    '9. 8の処理が終わるまで、待機
    '※終了すると、COMコールバックにより`OnControllerCompleted`プロシージャにワープします
    If Not WaitFor(m_core.ctrl.state) Then
        Debug.Print FromProcedureName & ": Controller完了待ちがタイムアウトしました"
        Exit Function
    End If
    If m_core.ctrl.state <> Whs_Ready Then Exit Function

    '10. ICoreWebView2ポインターを取得(Controller::get_CoreWebView2、vtable index 25)
    Dim pv As LongPtr
    hr = WebView2Thunks.dcf(m_pController, 25, "Ctrl.get_CoreWebView2", VarPtr(pv))
    If hr <> 0 Or pv = 0 Then
        Debug.Print FromProcedureName & ": get_CoreWebView2失敗 hr=&H" & Hex(hr)
        m_lastError = hr
        Exit Function
    End If
    m_pView = pv

    '11. ICoreWebView2のどの最新番号まで対応してるか探る
    DetectViewLevel


    '--------------------------------- WebView2設定用ポインター取得フェーズ ---------------------------------


    '12. ICoreWebView2Settingsポインターを取得(View::get_Settings、vtable index 3)
    '   ★毎回取得/解放していると呼び出しごとにQI相当のオーバーヘッドが発生するため、
    '   ここで1回だけ取得してキャッシュし、`CLOSE_`でまとめて解放する方式に統一する。
    '   取得に失敗しても(古いWebView2 Runtime等)Init自体は成功扱いとし、`m_pSettings=0`の
    '   ままにする(DevToolsEnabled等の各プロパティは0ガードで静かに無視するだけなので安全)。
    Dim ps As LongPtr
    hr = WebView2Thunks.dcf(m_pView, 3, "View.get_Settings", VarPtr(ps))
    If hr = 0 And ps <> 0 Then
        m_pSettings = ps

        '13. ICoreWebView2Settingsのどの最新番号まで対応してるか探る
        DetectSettingsLevel
    End If

    '14. ICoreWebView2Profileポインターを取得(View→ICoreWebView2_13へQueryInterface→get_Profile、
    '   vtable index 105)。取得に失敗しても(古いWebView2 Runtime等)Init自体は成功扱いとし、
    '   `m_pProfile=0`のままにする(各プロパティは`m_profileLevel`ガードで静かに無視されるだけ)
    Dim p13 As LongPtr
    Dim iid13 As WebView2Thunks.GUID

    '14-1. ICoreWebView2Profileポインターは、ICoreWebView2_13の所に所属する`get_Profile`でないと取れないのでまずそこのポインターを取得
    WebView2Thunks.FillGUID iid13, "f75f09a8-667e-4983-88d6-c8773f315e84"   ' ICoreWebView2_13
    hr = WebView2Thunks.dcf(m_pView, 0, "View.QueryInterface(ICoreWebView2_13)", VarPtr(iid13), VarPtr(p13))
    If hr = 0 And p13 <> 0 Then
        '14-2. 存在したら実際に、ICoreWebView2Profileポインターを取得していく
        Dim pProfile As LongPtr
        hr = WebView2Thunks.dcf(p13, 105, "View13.get_Profile", VarPtr(pProfile))
        WebView2Thunks.ComRelease p13   ' get_Profileを呼ぶためだけに必要だったので、以後は不要
        If hr = 0 And pProfile <> 0 Then
            m_pProfile = pProfile

            '14-3. ICoreWebView2Profileのどの最新番号まで対応してるか探る
            DetectProfileLevel
        End If
    End If

    '15. 成功サインを返す
    RunWebView2 = True
End Function

'***************************************************************************************************
'* 機能　　：`m_core.env.state`/`m_core.ctrl.state`がWhs_Waitingを脱するまでDoEventsで待ちます
'***************************************************************************************************
Private Function WaitFor(ByRef state As WebView2HostState) As Boolean
    '1. 今の経過時間を記録し、取り出し開始
    Dim timerStart As Double: timerStart = CDPHelpers.TimerCounter

    '2. 30s間、メッセージ取り出しを行う
    Do
        DoEvents
    Loop While state = Whs_Waiting And CDPHelpers.TimerCounter - timerStart < TIMEOUT_DEFAULT * 1000

    '3. 最終結果を返す
    WaitFor = (state <> Whs_Waiting)
End Function



'***************************************************************************************************
'                          ■■■ 最新インターフェース番号チェック ■■■
'***************************************************************************************************
'* 機能　　：`pInterface`が実際にQueryInterfaceで辿り着ける最大のインターフェースバージョンNを判定し、
'            `pInterface`をそのバージョンのポインタへ差し替えます
'---------------------------------------------------------------------------------------------------
'* 引数    ：pInterface  probe対象のポインタ(ByRef。成功時はその場で最大バージョンのポインタに
'                        差し替わり、元(下位バージョン)の参照は解放される)
'            iids        高いバージョン→低いバージョンの順に並べたIID文字列配列(1 To N)。
'                        `iids(N)`が対応する最低バージョンは常に2という前提で、最高位のバージョン
'                        番号は`UBound(iids) + 1`として逆算する
'            funcName    `dcf`呼び出し時のトレース用ラベル(省略可)
'* 返り値  ：実際に辿り着けた最大のバージョン番号(全滅時は基底のまま1)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`ICoreWebView2`本体/`Settings`/`Controller`/`Profile`は、いずれも`SettingsN : Settings
'            (N-1)`のような一本の継承チェーンで出来ているため、対応している最大のバージョンへ
'            1回QueryInterfaceすれば、そのポインタ1本だけで基底からそのバージョンまでの全メンバーに
'            (スロット番号さえ合わせれば)アクセスできる。高い方から順に試し、最初に成功したものを
'            採用する
'***************************************************************************************************
Private Function DetectLevel(ByRef pInterface As LongPtr, iids() As String, Optional ByVal funcName As String = "QueryInterface") As Long
    Dim highestLevel As Long
    highestLevel = UBound(iids) + 1

    Dim i As Long, iid As WebView2Thunks.GUID, pTry As LongPtr, hrQ As Long
    For i = 1 To UBound(iids)
        WebView2Thunks.FillGUID iid, iids(i)
        hrQ = WebView2Thunks.dcf(pInterface, 0, funcName, VarPtr(iid), VarPtr(pTry))
        If hrQ = 0 And pTry <> 0 Then
            WebView2Thunks.ComRelease pInterface   ' 基底(または下位)の分の参照は不要になるので解放
            pInterface = pTry
            DetectLevel = highestLevel - i + 1
            Exit Function
        End If
    Next i

    DetectLevel = 1   ' 全滅時は基底のまま
End Function

'***************************************************************************************************
'* 機能　　：`m_pView`の`ICoreWebView2_N`最新対応バージョンを判定し、`m_viewLevel`を更新します
'***************************************************************************************************
Private Sub DetectViewLevel()
    Dim iids(1 To 28) As String
    iids(1) = "650f917e-cc16-5c8c-9287-b0bbc8a4ab2c"    ' ICoreWebView2_29
    iids(2) = "62e50381-5bf5-51a8-aae0-f20a3a9c8a90"    ' ICoreWebView2_28
    iids(3) = "00fbe33b-8c07-517c-aa23-0ddd4b5f6fa0"    ' ICoreWebView2_27
    iids(4) = "806268b8-f897-5685-88e5-c45fca0b1a48"    ' ICoreWebView2_26
    iids(5) = "b5a86092-df50-5b4f-a17b-6c8f8b40b771"    ' ICoreWebView2_25
    iids(6) = "39a7ad55-4287-5cc1-88a1-c6f458593824"    ' ICoreWebView2_24
    iids(7) = "508f0db5-90c4-5872-90a7-267a91377502"    ' ICoreWebView2_23
    iids(8) = "db75dfc7-a857-4632-a398-6969dde26c0a"    ' ICoreWebView2_22
    iids(9) = "c4980dea-587b-43b9-8143-3ef3bf552d95"    ' ICoreWebView2_21
    iids(10) = "b4bc1926-7305-11ee-b962-0242ac120002"   ' ICoreWebView2_20
    iids(11) = "6921f954-79b0-437f-a997-c85811897c68"   ' ICoreWebView2_19
    iids(12) = "7a626017-28be-49b2-b865-3ba2b3522d90"   ' ICoreWebView2_18
    iids(13) = "702e75d4-fd44-434d-9d70-1a68a6b1192a"   ' ICoreWebView2_17
    iids(14) = "0eb34dc9-9f91-41e1-8639-95cd5943906b"   ' ICoreWebView2_16
    iids(15) = "517b2d1d-7dae-4a66-a4f4-10352ffb9518"   ' ICoreWebView2_15
    iids(16) = "6daa4f10-4a90-4753-8898-77c5df534165"   ' ICoreWebView2_14
    iids(17) = "f75f09a8-667e-4983-88d6-c8773f315e84"   ' ICoreWebView2_13
    iids(18) = "35d69927-bcfa-4566-9349-6b3e0d154cac"   ' ICoreWebView2_12
    iids(19) = "0be78e56-c193-4051-b943-23b460c08bdb"   ' ICoreWebView2_11
    iids(20) = "b1690564-6f5a-4983-8e48-31d1143fecdb"   ' ICoreWebView2_10
    iids(21) = "4d7b2eab-9fdc-468d-b998-a9260b5ed651"   ' ICoreWebView2_9
    iids(22) = "e9632730-6e1e-43ab-b7b8-7b2c9e62e094"   ' ICoreWebView2_8
    iids(23) = "79c24d83-09a3-45ae-9418-487f32a58740"   ' ICoreWebView2_7
    iids(24) = "499aadac-d92c-4589-8a75-111bfc167795"   ' ICoreWebView2_6
    iids(25) = "bedb11b8-d63c-11eb-b8bc-0242ac130003"   ' ICoreWebView2_5
    iids(26) = "20d02d59-6df2-42dc-bd06-f98a694b1302"   ' ICoreWebView2_4
    iids(27) = "a0d6df20-3b92-416d-aa0c-437a9c727857"   ' ICoreWebView2_3
    iids(28) = "9e8f0cf8-e670-4b5e-b2bc-73e061e3184c"   ' ICoreWebView2_2

    m_viewLevel = DetectLevel(m_pView, iids, "View.QueryInterface")
End Sub

'***************************************************************************************************
'* 機能　　：`m_pSettings`の`ICoreWebView2SettingsN`最新対応バージョンを判定し、`m_settingsLevel`を更新します
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

    m_settingsLevel = DetectLevel(m_pSettings, iids, "Settings.QueryInterface")
End Sub

'***************************************************************************************************
'* 機能　　：`m_pController`の`ICoreWebView2ControllerN`最新対応バージョンを判定し、`m_controllerLevel`を更新します
'***************************************************************************************************
Private Sub DetectControllerLevel()
    Dim iids(1 To 3) As String
    iids(1) = "97d418d5-a426-4e49-a151-e1a10f327d9e"   ' ICoreWebView2Controller4
    iids(2) = "f9614724-5d2b-41dc-aef7-73d62b51543b"   ' ICoreWebView2Controller3
    iids(3) = "c979903e-d4ca-4228-92eb-47ee3fa96eab"   ' ICoreWebView2Controller2

    m_controllerLevel = DetectLevel(m_pController, iids, "Ctrl.QueryInterface")
End Sub

'***************************************************************************************************
'* 機能　　：`m_pProfile`の`ICoreWebView2ProfileN`最新対応バージョンを判定し、`m_profileLevel`を更新します
'***************************************************************************************************
Private Sub DetectProfileLevel()
    Dim iids(1 To 8) As String
    iids(1) = "40c764d0-12fc-5d23-816d-971b353174c1"   ' ICoreWebView2Profile9
    iids(2) = "fbf70c2f-eb1f-4383-85a0-163e92044011"   ' ICoreWebView2Profile8
    iids(3) = "7b4c7906-a1aa-4cb4-b723-db09f813d541"   ' ICoreWebView2Profile7
    iids(4) = "bd82fa6a-1d65-4c33-b2b4-0393020cc61b"   ' ICoreWebView2Profile6
    iids(5) = "2ee5b76e-6e80-4df2-bcd3-d4ec3340a01b"   ' ICoreWebView2Profile5
    iids(6) = "8f4ae680-192e-4ec8-833a-21cfadaef628"   ' ICoreWebView2Profile4
    iids(7) = "b188e659-5685-4e05-bdba-fc640e0f1992"   ' ICoreWebView2Profile3
    iids(8) = "fa740d4b-5eae-4344-a8ad-74be31925397"   ' ICoreWebView2Profile2

    m_profileLevel = DetectLevel(m_pProfile, iids, "Profile.QueryInterface")
End Sub

'***************************************************************************************************
'* 機能　　：`m_pEnvironment`の`ICoreWebView2EnvironmentN`最新対応バージョンを判定し、
'            `m_environmentLevel`を更新します
'---------------------------------------------------------------------------------------------------
'* 注意事項：現状`_11`(`FailureReportFolderPath`)までしか消費していないため、探索もそこまでに留める
'***************************************************************************************************
Private Sub DetectEnvironmentLevel()
    Dim iids(1 To 10) As String
    iids(1) = "f0913dc6-a0ec-42ef-9805-91dff3a2966a"    ' ICoreWebView2Environment11
    iids(2) = "ee0eb9df-6f12-46ce-b53f-3f47b9c928e0"    ' ICoreWebView2Environment10
    iids(3) = "f06f41bf-4b5a-49d8-b9f6-fa16cd29f274"    ' ICoreWebView2Environment9
    iids(4) = "d6eb91dd-c3d2-45e5-bd29-6dc2bc4de9cf"    ' ICoreWebView2Environment8
    iids(5) = "43c22296-3bbd-43a4-9c00-5c0df6dd29a2"    ' ICoreWebView2Environment7
    iids(6) = "e59ee362-acbd-4857-9a8e-d3644d9459a9"    ' ICoreWebView2Environment6
    iids(7) = "319e423d-e0d7-4b8d-9254-ae9475de9b17"    ' ICoreWebView2Environment5
    iids(8) = "20944379-6dcf-41d6-a0a0-abc0fc50de0d"    ' ICoreWebView2Environment4
    iids(9) = "80a22ae3-be7c-4ce2-afe1-5a50056cdeeb"    ' ICoreWebView2Environment3
    iids(10) = "41f3632b-5ef4-404f-ad82-2d606c5a9a21"   ' ICoreWebView2Environment2

    m_environmentLevel = DetectLevel(m_pEnvironment, iids, "Env.QueryInterface")
End Sub



'***************************************************************************************************
'                                 ■■■ 拡張機能 操作一式 ■■■
'***************************************************************************************************
'* 注意事項：`ICoreWebView2Profile7`/`ICoreWebView2BrowserExtension`をまたぐ複合フロー(要求発行→
'            完了待ち→結果の取り出し)をここに集約する。個々のインターフェース呼び出し自体
'            (バージョンチェック/Vtable発行/ハンドラ準備)は上の`ICoreWebView2Profile7`セクション内の
'            `IssueXxx`系プロシージャに任せ、ここでは`WaitFor`による待機と結果の取り出しのみを行う
'***************************************************************************************************
'* 機能　　：指定フォルダの未パック拡張機能を、プロファイルへインストールします(Profile7以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：extensionFolderPath  拡張機能の`manifest.json`を含むトップフォルダの絶対パス
'* 返り値  ：インストールされた拡張機能のID(失敗時は空文字。詳細は`LastError`/`LastErrorMessage`)
'***************************************************************************************************
Public Function AddBrowserExtension(ByVal extensionFolderPath As String) As String
    Const FromProcedureName As String = "WebView2Host.AddBrowserExtension"


    '1. 予め定義した拡張機能インストールプロシージャに処理を委託
    If Not IssueAddBrowserExtension(extensionFolderPath) Then Exit Function

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ2: 非同期処理が終わるまで「郵便ポストを見ながら」待つ！】
    '-----------------------------------------------------------------------------------------------
    ' ★ここがポイント！
    ' C++側のインストール作業は裏で進むので、終わるまでWaitForで待機する。
    ' (WaitForの中でメッセージループが回ることで、完了時のWindowsメッセージを開封され、対象の「完了コールバック」プロシージャへワープする)
    If Not WaitFor(m_ext.addExt.state) Then
        Debug.Print FromProcedureName & ": 完了待ちがタイムアウトしました"
        Exit Function
    End If

    '3. インストールIDを返す
    AddBrowserExtension = m_ext.addExt.resultId
End Function

'***************************************************************************************************
'* 機能　　：`GetBrowserExtensions`要求の発行から完了待ちまでを行い、結果を`m_ext.getExt.items`へ
'            格納します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功可否
'* 注意事項：`m_ext.getExt.items`の各値は`ICoreWebView2BrowserExtension`への新規参照であり、
'            呼び出し元が`ReleaseGetExtPtrs`で解放する責任を持つ
'***************************************************************************************************
Private Function RunGetBrowserExtensions() As Boolean
    Const FromProcedureName As String = "WebView2Host.RunGetBrowserExtensions"

    
    '1. 予め定義した拡張機能インストールプロシージャに処理を委託
    If Not IssueGetBrowserExtensions Then Exit Function

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ2: 自前await(WaitFor)で待つ！ ── 裏で起きる大スペクタクル】
    '-----------------------------------------------------------------------------------------------
    ' ここでメッセージループを回しながら、C++の返事を待つ！
    '
    ' ★この待機中に、今日の前半の「あのコード」が裏で発火します！↓
    '   ① C++からポストに手紙が届く！
    '   ② 配達員が`WebView2Host.OnGetBrowserExtensionsCompleted`を呼び出す！
    '   ③ 3番(get_Count)で個数を調べてループ！
    '   ④ 4番(GetValueAtIndex)で拡張機能オブジェクトを取り出す！
    '   ⑤ 3番(get_Id)でID文字列を引っこ抜いて、ステップ1で作った m_ext.getExt.Items に Add！
    '   ⑥ ぜんぶ終わったら state を「準備完了(Whs_Ready)」に書き換える！
    '
    If Not WaitFor(m_ext.getExt.state) Then
        Debug.Print FromProcedureName & ": 完了待ちがタイムアウトしました"
        Exit Function
    End If

    '3. 成功有無を返す
    RunGetBrowserExtensions = (m_ext.getExt.state = Whs_Ready)
End Function

'***************************************************************************************************
'* 機能　　：インストール済みの拡張機能を、ID/名前/有効状態付きで一覧取得します(Profile7以降)
'---------------------------------------------------------------------------------------------------
'* 返り値  ：各要素が`Dictionary`(キー:"ID"/"Name"/"IsEnabled")の`Collection`。
'            0件/失敗時は空の`Collection`
'***************************************************************************************************
Public Function GetBrowserExtensionIds() As Collection
    Dim results As New Collection

    If Not RunGetBrowserExtensions() Then
        ReleaseGetExtPtrs
        Set GetBrowserExtensionIds = results
        Exit Function
    End If

    Dim Key As Variant, pExt As LongPtr, info As Dictionary
    For Each Key In m_ext.getExt.Items.keys
        pExt = m_ext.getExt.Items(Key)

        Set info = New Dictionary
        info.Add "ID", CStr(Key)
        info.Add "Name", WebView2Thunks.GetStringProperty(pExt, 4, "BrowserExtension.get_Name")
        info.Add "IsEnabled", WebView2Thunks.GetBoolProperty(pExt, 6, "BrowserExtension.get_IsEnabled")
        results.Add info
    Next Key

    Set GetBrowserExtensionIds = results
    ReleaseGetExtPtrs
End Function

'***************************************************************************************************
'* 機能　　：指定IDの拡張機能をアンインストールします(Profile7以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：extensionId  `AddBrowserExtension`の戻り値、または`GetBrowserExtensionIds`の各要素の
'            "ID"キーで得られる拡張機能ID
'* 返り値  ：成功可否
'---------------------------------------------------------------------------------------------------
'* 詳細説明：WebView2にはID指定の直接アンインストールAPIが無いため、`GetBrowserExtensions`で
'            一覧を取得して対象を特定してから`ICoreWebView2BrowserExtension::Remove`を呼ぶ、
'            という2段構えになる(CDPの`Extensions.uninstall`はWebView2では使えないため代替不可)
'***************************************************************************************************
Public Function RemoveBrowserExtension(ByVal extensionId As String) As Boolean
    Const FromProcedureName As String = "WebView2Host.RemoveBrowserExtension"


    '-----------------------------------------------------------------------------------------------
    ' 【第1段: マンション(拡張機能)の住人を全員ロビーに集める(全件取得)】
    '-----------------------------------------------------------------------------------------------
    ' 1-1. 全拡張機能の一覧を取得する(裏でGetBrowserExtensions → WaitForが走る)
    If Not RunGetBrowserExtensions() Then
        ' 失敗した場合は、途中で確保されたポインタたちを安全に全解放してお掃除
        ReleaseGetExtPtrs
        Exit Function
    End If

    ' 1-2. 取得した一覧の中に、消したいIDの住人が存在するかチェック！
    If Not m_ext.getExt.Items.Exists(extensionId) Then
        Debug.Print FromProcedureName & ": 指定IDの拡張機能が見つかりません: " & extensionId

        ' いなかったので、集めた住人ポインタたちを全員解放してお見送り
        ReleaseGetExtPtrs
        Exit Function
    End If

    '-----------------------------------------------------------------------------------------------
    ' 【第2段の手前: お目当ての住人を確保し、無関係な住人を全員解放する】
    '-----------------------------------------------------------------------------------------------
    ' 2-1. 消したい対象のオブジェクトポインタ(pTarget)をキープ！
    Dim pTarget As LongPtr: pTarget = m_ext.getExt.Items(extensionId)

    ' 2-2. ★超重要(メモリ管理):
    '      対象(extensionId)「以外」の無関係なポインタはもう用済みなので、
    '      この瞬間に全員 Release してメモリから解放する！(メモリリーク防止)
    ReleaseGetExtPtrs extensionId

    '-----------------------------------------------------------------------------------------------
    ' 【第2段: 本人に「退去届」を出させる(Remove実行)】
    '-----------------------------------------------------------------------------------------------
    ' 3-1. 対象オブジェクト(pTarget)の Remove メソッドを非同期で叩く！
    '      (内部で HK_RemoveBrowserExtensionCompleted の配達員がセットされる)
    Dim issued As Boolean
    issued = IssueRemoveBrowserExtension(pTarget)

    ' 3-2. ★超神業(メモリ管理):
    '      GetValueAtIndex で取得した時に増えていた参照カウント(+1)を、ここで手放す(-1)！
    '      ※Remove処理自体は裏のC++側でキープされて継続するので、
    '        VBA側の手元の参照はここで解放してしまって全く問題ない。
    WebView2Thunks.ComRelease pTarget
    
    ' 発行自体に失敗していたら終了
    If Not issued Then Exit Function

    '-----------------------------------------------------------------------------------------------
    ' 【第2段の完了待ち: 退去が完了するまで自前await(WaitFor)で待つ】
    '-----------------------------------------------------------------------------------------------
    ' 削除完了コールバックが届いて、state が Whs_Ready になるまでメッセージループ待機！
    If Not WaitFor(m_ext.removeExt.state) Then
        Debug.Print FromProcedureName & ": 完了待ちがタイムアウトしました"
        Exit Function
    End If

    '-----------------------------------------------------------------------------------------------
    ' 【ゴール: 成功判定を返却】
    '-----------------------------------------------------------------------------------------------
    RemoveBrowserExtension = (m_ext.removeExt.state = Whs_Ready)
End Function

'***************************************************************************************************
'* 機能　　：`m_ext.getExt.items`の参照を解放します
'---------------------------------------------------------------------------------------------------
'* 引数    ：keepId  この拡張機能ID分の参照だけは解放せず残す(既定は空文字で全解放)
'***************************************************************************************************
Private Sub ReleaseGetExtPtrs(Optional ByVal keepId As String)
    If m_ext.getExt.Items Is Nothing Then Exit Sub

    Dim k As Variant
    For Each k In m_ext.getExt.Items.keys
        If CStr(k) <> keepId Then
            Dim p As LongPtr
            p = m_ext.getExt.Items(k)
            If p <> 0 Then WebView2Thunks.ComRelease p
        End If
    Next k
End Sub



'***************************************************************************************************
'                              ■■■ 完了コールバック(Public Sub必須) ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 必須系完了コールバック
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：`RunWebView2`の`CreateCoreWebView2EnvironmentWithOptions`の処理を終えると呼び出されます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Step5 の待機中に、COMコールバックによって、呼ばれる想定です
'***************************************************************************************************
Public Sub OnEnvironmentCompleted(ByVal ErrorCode As Long, ByVal pEnvironment As LongPtr)
    '1. 完了コールバック全般の最後のエラーコードを記録
    m_lastError = ErrorCode

    '2. エラーチェック
    If ErrorCode = 0 And pEnvironment <> 0 Then
        '2-1. ICoreWebView2Environment ポインターを取得
        m_pEnvironment = pEnvironment

        '2-2. VBA用として今後も使うので、保持する合図を出して、参照カウントUPする
        WebView2Thunks.ComAddRef m_pEnvironment

        '2-3. 準備完了statusに更新
        m_core.env.state = Whs_Ready
    Else
        '2-1. 失敗として記述
        m_pEnvironment = 0
        m_core.env.state = Whs_Failed
    End If

    '3. もう使わないので解放。1ショットハンドラパターン
    Set m_core.env.handler = Nothing
End Sub

'***************************************************************************************************
'* 機能　　：`RunWebView2`の`ICoreWebView2Controller`オブジェクトの作成が終わると呼び出されます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・Step9 の待機の間に来る想定です
'            ・ついでに、`ICoreWebView2Controller`の最新番号も調査します
'***************************************************************************************************
Public Sub OnControllerCompleted(ByVal ErrorCode As Long, ByVal pController As LongPtr)
    '1. 完了コールバック全般の最後のエラーコードを記録
    m_lastError = ErrorCode

    '2. エラーチェック
    If ErrorCode = 0 And pController <> 0 Then
        '2-1. `ICoreWebView2Controller`ポインターを取得
        m_pController = pController

        '2-2. VBA用として今後も使うので、保持する合図を出して、参照カウントUPする
        WebView2Thunks.ComAddRef m_pController

        '2-3. `ICoreWebView2Controller`の最新番号を調査
        DetectControllerLevel

        '2-4. 準備完了statusに更新
        m_core.ctrl.state = Whs_Ready
    Else
        '2-1. 失敗として記述
        m_pController = 0
        m_core.ctrl.state = Whs_Failed
    End If

    '3. もう使わないので解放。1ショットハンドラパターン
    Set m_core.ctrl.handler = Nothing
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 拡張機能系完了コールバック
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：`AddBrowserExtension`の完了コールバックです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`pResult`(ICoreWebView2BrowserExtension*)は借用ポインタ(このコールバックの間だけ
'            有効。WebView2側が呼び出し後にReleaseする)。`get_Id`(vtableスロット3)を読むだけで
'            AddRef/Releaseは行わない
'***************************************************************************************************
Public Sub OnAddBrowserExtensionCompleted(ByVal ErrorCode As Long, ByVal pResult As LongPtr)
    '1. 完了コールバック全般の最後のエラーコードを記録
    m_lastError = ErrorCode

    '2. エラーチェック
    If ErrorCode = 0 And pResult <> 0 Then
        '2-1. 結果ポインターから`ICoreWebView2BrowserExtension`Vtableの`get_Id`を読みに行って、読み込んだ拡張機能IDを取り出す
        m_ext.addExt.resultId = WebView2Thunks.GetStringProperty(pResult, 3, "BrowserExtension.get_Id")

        '2-2. 準備完了statusに更新
        m_ext.addExt.state = Whs_Ready
    Else
        '2-1. 失敗として記述
        m_ext.addExt.resultId = vbNullString    '空で上書き。前回値を残さないように
        m_ext.addExt.state = Whs_Failed
    End If

    '3. もう使わないので解放。1ショットハンドラパターン
    Set m_ext.addExt.handler = Nothing
End Sub

'***************************************************************************************************
'* 機能　　：`GetBrowserExtensions`の完了コールバックです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`pResult`(ICoreWebView2BrowserExtensionList*)は借用ポインタ。一方、
'            `GetValueAtIndex`で得る各`ICoreWebView2BrowserExtension`は[retval]の新規参照なので、
'            `m_ext.getExt.items`に保持したまま返し、解放責任は呼び出し元(`ReleaseGetExtPtrs`)に委ねる
'***************************************************************************************************
Public Sub OnGetBrowserExtensionsCompleted(ByVal ErrorCode As Long, ByVal pResult As LongPtr)
    '1. 完了コールバック全般の最後のエラーコードを記録
    m_lastError = ErrorCode

    '2. エラーチェック
    If ErrorCode = 0 And pResult <> 0 Then
        '2-1. 結果ポインターから`ICoreWebView2BrowserExtensionList`Vtableの`get_Count`を読みに行って、インストール済み拡張機能ID数を取り出す
        Dim cnt As Long
        Dim hr As Long
        hr = WebView2Thunks.dcf(pResult, 3, "BrowserExtensionList.get_Count", VarPtr(cnt))
        If hr = 0 And cnt > 0 Then
            '2-2. 見つけた数分、ループで、拡張機能IDを取り出していく
            Dim i As Long
            For i = 0 To cnt - 1
                '2-3. 配列番号を基に、`ICoreWebView2BrowserExtensionList`Vtableの`GetValueAtIndex`から、`ICoreWebView2BrowserExtension`というオブジェクトポインターを取り出す
                Dim pExt As LongPtr
                hr = WebView2Thunks.dcf(pResult, 4, "BrowserExtensionList.GetValueAtIndex", CLng(i), VarPtr(pExt))
                If hr = 0 And pExt <> 0 Then
                    '2-4. 拡張機能ID(2-3のオブジェクトポインターから、文字列取り出し)と、2-3のオブジェクトポインターを紐づける
                    m_ext.getExt.Items.Add WebView2Thunks.GetStringProperty(pExt, 3, "BrowserExtension.get_Id"), pExt
                End If
            Next i
        End If

        '2-5. 必要なループを終えたら、準備完了statusに更新
        m_ext.getExt.state = Whs_Ready
    Else
        '2-1. 失敗として更新
        m_ext.getExt.state = Whs_Failed
    End If

    '3. もう使わないので解放。1ショットハンドラパターン
    Set m_ext.getExt.handler = Nothing
End Sub

'***************************************************************************************************
'* 機能　　：`ICoreWebView2BrowserExtension::Remove`の完了コールバックです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：SDKシグネチャが`Invoke(HRESULT errorCode)`のみ(結果オブジェクトが無い)ため、
'            引数は`ErrorCode`だけになる
'***************************************************************************************************
Public Sub OnRemoveBrowserExtensionCompleted(ByVal ErrorCode As Long)
    '1. 完了コールバック全般の最後のエラーコードを記録
    m_lastError = ErrorCode

    '2. エラーコードに応じたstatusに更新
    m_ext.removeExt.state = IIf(ErrorCode = 0, Whs_Ready, Whs_Failed)

    '3. もう使わないので解放。1ショットハンドラパターン
    Set m_ext.removeExt.handler = Nothing
End Sub



'***************************************************************************************************
'                              ■■■ 明示的な後始末 ■■■
'***************************************************************************************************
'* 機能　　：Controller::Close → 各種ポインタのRelease → ホストウィンドウの破棄を行います
'* 注意事項：ここでは`WebView2Thunks.Thunks_Shutdown`は呼ばない。あちらはこのモジュール内で
'            全`WebView2Host`インスタンス共通(グローバル)のスロットプールを解体するため、
'            他のインスタンスがまだ使用中の可能性がある間に呼ぶと事故になる
'***************************************************************************************************
Public Sub CLOSE_()
    If m_closed Then Exit Sub

    If m_pController <> 0 Then
        WebView2Thunks.dcf m_pController, 24, "Ctrl.Close"
    End If
    If m_pSettings <> 0 Then
        WebView2Thunks.ComRelease m_pSettings
        m_pSettings = 0
    End If
    If m_pProfile <> 0 Then
        WebView2Thunks.ComRelease m_pProfile
        m_pProfile = 0
    End If
    If m_pView <> 0 Then
        WebView2Thunks.ComRelease m_pView
        m_pView = 0
    End If
    If m_pController <> 0 Then
        WebView2Thunks.ComRelease m_pController
        m_pController = 0
    End If
    If m_pEnvironment <> 0 Then
        WebView2Thunks.ComRelease m_pEnvironment
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
'                              ■■■ Environment拡張 一式 ■■■
'***************************************************************************************************
'* 機能　　：使用中のWebView2ランタイム(Edge)のバージョン文字列を取得します
'***************************************************************************************************
Property Get BrowserVersionString() As String
    BrowserVersionString = WebView2Thunks.GetStringProperty(m_pEnvironment, 5, "Env.get_BrowserVersionString")
End Property

'***************************************************************************************************
'* 機能　　：実際に使用されているユーザーデータフォルダの絶対パスを取得します(ICoreWebView2Environment7以降)
'---------------------------------------------------------------------------------------------------
'* 注意事項：`getUserDataDirPath`(VBA側で組み立てたパス文字列をそのまま返すだけ)とは別に、
'            ランタイムが実際に認識している値をCOM経由で正式に取得します
'***************************************************************************************************
Property Get userDataFolder() As String
    If m_environmentLevel < 7 Then Exit Property
    userDataFolder = WebView2Thunks.GetStringProperty(m_pEnvironment, 15, "Env7.get_UserDataFolder")
End Property

'***************************************************************************************************
'* 機能　　：クラッシュ時のminidumpファイルが書き出されるフォルダのパスを取得します
'            (ICoreWebView2Environment11以降)
'***************************************************************************************************
Property Get FailureReportFolderPath() As String
    If m_environmentLevel < 11 Then Exit Property
    FailureReportFolderPath = WebView2Thunks.GetStringProperty(m_pEnvironment, 23, "Env11.get_FailureReportFolderPath")
End Property



'***************************************************************************************************
'                               ■■■ ICoreWebView2Controller 一式 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Controller
'---------------------------------------------------------------------------------------------------
'* 機能　　：リサイズ処理を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：WV2_Right      横の長さ
'            WV2_Bottom     縦の長さ
'            WV2_Top        上端からのオフセット値
'            WV2_Left       左端からのオフセット値
'***************************************************************************************************
Public Sub Resize(WV2_Right As Long, WV2_Bottom As Long, WV2_Top As Long, WV2_Left As Long)
    '1. ひとまとまりにする
    Dim rc As WebView2Thunks.RECT
    rc.Left = WV2_Left: rc.Top = WV2_Top: rc.Right = WV2_Right: rc.Bottom = WV2_Bottom

    '2. リサイズ
    WebView2Thunks.dcf m_pController, 6, "Ctrl.put_Bounds", VarPtr(rc)       ' put_Bounds
End Sub

'***************************************************************************************************
'* 機能　　：表示/非表示切り替え
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Web描画の表示/非表示
'***************************************************************************************************
Property Let Visible(v As Boolean)
    WebView2Thunks.dcf m_pController, 4, "Ctrl.put_IsVisible", CLng(v * -1)  ' put_IsVisible
End Property

'***************************************************************************************************
'* 機能　　：ブラウザUIレベルのズーム倍率を設定します(CDPの`Emulation.setPageScaleFactor`とは別物)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  ズーム倍率(1.0が等倍)
'***************************************************************************************************
Property Let ZoomFactor(v As Double)
    WebView2Thunks.dcf m_pController, 8, "Ctrl.put_ZoomFactor", v
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
    Dim rc As WebView2Thunks.RECT
    rc.Left = WV2_Left: rc.Top = WV2_Top: rc.Right = WV2_Right: rc.Bottom = WV2_Bottom
    WebView2Thunks.dcf m_pController, 11, "Ctrl.SetBoundsAndZoomFactor", VarPtr(rc), ZoomFactor
End Sub

'***************************************************************************************************
'* 機能　　：フォーカスをWebView2側へプログラム的に移動します
'---------------------------------------------------------------------------------------------------
'* 引数    ：reason  `COREWEBVIEW2_MOVE_FOCUS_REASON`(Long。0=Programmatic, 1=Next, 2=Previous。
'            `CDPCoreViaWebView2`側の`WV2MoveFocusReason`Enum使用を推奨)
'***************************************************************************************************
Public Sub MoveFocus(reason As Long)
    WebView2Thunks.dcf m_pController, 12, "Ctrl.MoveFocus", reason
End Sub

'***************************************************************************************************
'* 機能　　：親ウィンドウの位置が動いたことをWebView2へ明示的に通知します
'---------------------------------------------------------------------------------------------------
'* 注意事項：ポップアップ/ドロップダウン等、親ウィンドウの画面上の絶対位置に依存する描画がある場合に
'            呼ぶためのものです。このプロジェクトの非表示運用では基本的に不要です
'***************************************************************************************************
Public Sub NotifyParentWindowPositionChanged()
    WebView2Thunks.dcf m_pController, 23, "Ctrl.NotifyParentWindowPositionChanged"
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Controller2
'---------------------------------------------------------------------------------------------------
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
    WebView2Thunks.dcf m_pController, 27, "Ctrl2.put_DefaultBackgroundColor", packed
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Controller3
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：DPIスケールを手動指定します(Controller3以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  DPIスケール値(例:1.0=100%、1.5=150%)
'***************************************************************************************************
Property Let RasterizationScale(v As Double)
    If m_controllerLevel < 3 Then Exit Property
    WebView2Thunks.dcf m_pController, 29, "Ctrl3.put_RasterizationScale", v
End Property

'***************************************************************************************************
'* 機能　　：モニターのDPIスケール変更を自動検知するか切り替えます(Controller3以降、既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで自動検知、Falseで`RasterizationScale`を手動管理する
'***************************************************************************************************
Property Let ShouldDetectMonitorScaleChanges(v As Boolean)
    If m_controllerLevel < 3 Then Exit Property
    WebView2Thunks.dcf m_pController, 31, "Ctrl3.put_ShouldDetectMonitorScaleChanges", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：`Bounds`(位置/サイズ)の解釈方式を切り替えます(Controller3以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  `COREWEBVIEW2_BOUNDS_MODE`(Long。0=USE_RAW_PIXELS(既定)、1=USE_RASTERIZATION_SCALE。
'            `CDPCoreViaWebView2`側の`WV2BoundsMode`Enum使用を推奨)
'***************************************************************************************************
Property Let BoundsMode(v As Long)
    If m_controllerLevel < 3 Then Exit Property
    WebView2Thunks.dcf m_pController, 35, "Ctrl3.put_BoundsMode", v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Controller4
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：外部(エクスプローラー等)からのドラッグ&ドロップ受け入れ可否を切り替えます
'            (Controller4以降、既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let AllowExternalDrop(v As Boolean)
    If m_controllerLevel < 4 Then Exit Property
    WebView2Thunks.dcf m_pController, 37, "Ctrl4.put_AllowExternalDrop", CLng(v * -1)
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
    WebView2Thunks.dcf m_pSettings, 12, "Settings.put_AreDevToolsEnabled", CLng(v * -1)   ' put_AreDevToolsEnabled
End Property

'***************************************************************************************************
'* 機能　　：既定の右クリックコンテキストメニューの表示可否を切り替えます
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let ContextMenuEnabled(v As Boolean)
    If m_pSettings = 0 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 14, "Settings.put_AreDefaultContextMenusEnabled", CLng(v * -1)   ' put_AreDefaultContextMenusEnabled
End Property

'***************************************************************************************************
'* 機能　　：ページ内のJavaScript実行可否を切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let ScriptEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 4, "Settings.put_IsScriptEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：`postMessage`/`WebMessageReceived`によるWebページ⇔ホストアプリ間の通信可否を
'            切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let WebMessageEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 6, "Settings.put_IsWebMessageEnabled", CLng(v * -1)
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
    WebView2Thunks.dcf m_pSettings, 8, "Settings.put_AreDefaultScriptDialogsEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：リンクにマウスオーバーした際のURLプレビュー等、ステータスバーの表示可否を
'            切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let StatusBarEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 10, "Settings.put_IsStatusBarEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：`AddHostObjectToScript`で追加したホストオブジェクトへ、スクリプト側からアクセスできるか
'            切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let HostObjectsAllowed(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 16, "Settings.put_AreHostObjectsAllowed", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：Ctrl+ホイール/ピンチ/キーボードショートカット等によるユーザー操作でのズーム可否を
'            切り替えます(既定True)
'***************************************************************************************************
Property Let ZoomControlEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 18, "Settings.put_IsZoomControlEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：ページ読み込み失敗時等に表示される、WebView2内蔵のエラーページ表示可否を
'            切り替えます(既定True)
'***************************************************************************************************
Property Let BuiltInErrorPageEnabled(v As Boolean)
    If m_settingsLevel < 1 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 20, "Settings.put_IsBuiltInErrorPageEnabled", CLng(v * -1)
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Settings2
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：リクエストヘッダーおよび`navigator.userAgent`に使われるUser-Agent文字列を上書きします
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  上書きするUser-Agent文字列
'---------------------------------------------------------------------------------------------------
'* 注意事項：`ICoreWebView2EnvironmentOptions2`とは別物(こちらは実行時にいつでも変更可能)
'***************************************************************************************************
Property Let UserAgentOverride(v As String)
    If m_settingsLevel < 2 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 22, "Settings2.put_UserAgent", StrPtr(v)
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Settings3
'---------------------------------------------------------------------------------------------------
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
    WebView2Thunks.dcf m_pSettings, 24, "Settings3.put_AreBrowserAcceleratorKeysEnabled", CLng(v * -1)
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Settings4
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：パスワードの自動保存機能(保存確認ダイアログ含む)の有効可否を切り替えます(既定True)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let PasswordAutosaveEnabled(v As Boolean)
    If m_settingsLevel < 4 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 26, "Settings4.put_IsPasswordAutosaveEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：住所/フォーム入力内容等の一般的なオートフィル機能(パスワード以外)の有効可否を
'            切り替えます(既定True)
'***************************************************************************************************
Property Let GeneralAutofillEnabled(v As Boolean)
    If m_settingsLevel < 4 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 28, "Settings4.put_IsGeneralAutofillEnabled", CLng(v * -1)
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Settings5
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：タッチ操作によるピンチズームジェスチャーの有効可否を切り替えます(既定True)
'***************************************************************************************************
Property Let PinchZoomEnabled(v As Boolean)
    If m_settingsLevel < 5 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 30, "Settings5.put_IsPinchZoomEnabled", CLng(v * -1)
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Settings6
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：左右スワイプによる進む/戻るナビゲーションジェスチャーの有効可否を切り替えます(既定True)
'***************************************************************************************************
Property Let SwipeNavigationEnabled(v As Boolean)
    If m_settingsLevel < 6 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 32, "Settings6.put_IsSwipeNavigationEnabled", CLng(v * -1)
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Settings7
'---------------------------------------------------------------------------------------------------
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
    WebView2Thunks.dcf m_pSettings, 34, "Settings7.put_HiddenPdfToolbarItems", v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Settings8
'---------------------------------------------------------------------------------------------------
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
    WebView2Thunks.dcf m_pSettings, 36, "Settings8.put_IsReputationCheckingRequired", CLng(v * -1)
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Settings9
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：Webページ側でCSSの`app-region: drag`/`no-drag`を使った独自タイトルバー領域(ドラッグで
'            ウィンドウ自体を移動できる領域)の指定を有効にするか切り替えます(既定False)
'***************************************************************************************************
Property Let NonClientRegionSupportEnabled(v As Boolean)
    If m_settingsLevel < 9 Then Exit Property
    WebView2Thunks.dcf m_pSettings, 38, "Settings9.put_IsNonClientRegionSupportEnabled", CLng(v * -1)
End Property



'***************************************************************************************************
'                               ■■■ ICoreWebView2Profile 一式 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Profile
'---------------------------------------------------------------------------------------------------
'* 注意事項：`ICoreWebView2Profile`は`ICoreWebView2Settings`と紛らわしい同名プロパティ
'            (`IsPasswordAutosaveEnabled`/`IsGeneralAutofillEnabled`)を持つが、こちらは
'            **プロファイル単位**(同じ`user-data-dir`を使う全WebViewに共通)の設定であり、
'            Settings側の**WebViewインスタンス単位**の設定とは別物。名前衝突を避けるため
'            `Profile`を頭に付けている
'***************************************************************************************************
'* 機能　　：プロファイル名を取得します(読み取り専用)
'***************************************************************************************************
Property Get ProfileName() As String
    If m_profileLevel < 1 Then Exit Property
    ProfileName = WebView2Thunks.GetStringProperty(m_pProfile, 3, "Profile.get_ProfileName")
End Property

'***************************************************************************************************
'* 機能　　：InPrivate(プライベート)モードが有効かどうかを取得します(読み取り専用)
'***************************************************************************************************
Property Get IsInPrivateModeEnabled() As Boolean
    If m_profileLevel < 1 Then Exit Property
    IsInPrivateModeEnabled = WebView2Thunks.GetBoolProperty(m_pProfile, 4, "Profile.get_IsInPrivateModeEnabled")
End Property

'***************************************************************************************************
'* 機能　　：プロファイルディレクトリの絶対パスを取得します(読み取り専用)
'***************************************************************************************************
Property Get ProfilePath() As String
    If m_profileLevel < 1 Then Exit Property
    ProfilePath = WebView2Thunks.GetStringProperty(m_pProfile, 5, "Profile.get_ProfilePath")
End Property

'***************************************************************************************************
'* 機能　　：既定のダウンロード先フォルダを設定します
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  フォルダの絶対パス
'***************************************************************************************************
Property Let DefaultDownloadFolderPath(v As String)
    If m_profileLevel < 1 Then Exit Property
    WebView2Thunks.dcf m_pProfile, 7, "Profile.put_DefaultDownloadFolderPath", StrPtr(v)
End Property

'***************************************************************************************************
'* 機能　　：優先カラースキームを設定します
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  `COREWEBVIEW2_PREFERRED_COLOR_SCHEME`(Long。0=Auto(既定), 1=Light, 2=Dark。
'            `CDPCoreViaWebView2`側の`WV2PreferredColorScheme`Enum使用を推奨)
'***************************************************************************************************
Property Let PreferredColorScheme(v As Long)
    If m_profileLevel < 1 Then Exit Property
    WebView2Thunks.dcf m_pProfile, 9, "Profile.put_PreferredColorScheme", v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Profile3
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：トラッキング防止レベルを設定します(Profile3以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  `COREWEBVIEW2_TRACKING_PREVENTION_LEVEL`(Long。0=None, 1=Basic, 2=Balanced(既定),
'            3=Strict。`CDPCoreViaWebView2`側の`WV2TrackingPreventionLevel`Enum使用を推奨)
'***************************************************************************************************
Property Let PreferredTrackingPreventionLevel(v As Long)
    If m_profileLevel < 3 Then Exit Property
    WebView2Thunks.dcf m_pProfile, 14, "Profile3.put_PreferredTrackingPreventionLevel", v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Profile6
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：パスワード自動保存機能の有効可否を、**プロファイル単位**で切り替えます(Profile6以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let ProfilePasswordAutosaveEnabled(v As Boolean)
    If m_profileLevel < 6 Then Exit Property
    WebView2Thunks.dcf m_pProfile, 19, "Profile6.put_IsPasswordAutosaveEnabled", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：一般オートフィル機能の有効可否を、**プロファイル単位**で切り替えます(Profile6以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let ProfileGeneralAutofillEnabled(v As Boolean)
    If m_profileLevel < 6 Then Exit Property
    WebView2Thunks.dcf m_pProfile, 21, "Profile6.put_IsGeneralAutofillEnabled", CLng(v * -1)
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Profile7
'---------------------------------------------------------------------------------------------------
'* 注意事項：ここでの各`IssueXxx`は「バージョンチェック→Vtable呼び出し→(完了通知が要るものは)
'            ハンドラ準備」までを担当し、完了待ち(`WaitFor`)自体はここでは行わない
'            (待機/結果の取り出しは「■■■ 拡張機能 操作一式 ■■■」セクション側の責務)
'***************************************************************************************************
'* 機能　　：指定フォルダの未パック拡張機能を、プロファイルへインストールする要求を発行します
'            (Profile7以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：extensionFolderPath  拡張機能の`manifest.json`を含むトップフォルダの絶対パス
'* 返り値  ：要求の発行に成功したか(完了自体はまだ先。完了通知は`WebView2Host.OnAddBrowserExtensionCompleted`に届く)
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`EnvironmentOptions.AreBrowserExtensionsEnabled=True`でEnvironmentを作成していないと
'              `ERROR_NOT_SUPPORTED`で失敗します(SDK仕様)
'            ・`ICoreWebView2::CallDevToolsProtocolMethod`経由の`Extensions.loadUnpacked`は
'              WebView2では`Method not available`で拒否されるため、本メソッド(SDK公式API)を
'              使う必要がある
'            ・完了通知作成を伴うため、ここのフローのみ詳細化してます
'***************************************************************************************************
Private Function IssueAddBrowserExtension(ByVal extensionFolderPath As String) As Boolean
    Const FromProcedureName As String = "WebView2Host.IssueAddBrowserExtension"


    '-----------------------------------------------------------------------------------------------
    ' 【ステップ1: バージョン確認(身分証チェック)】
    '-----------------------------------------------------------------------------------------------
    ' 拡張機能の追加は「ICoreWebView2Profile7」以降でしか使えない最新機能！
    ' 実行環境のEdgeが古くてProfile7未満なら、門前払いで抜ける(クラッシュ防止)
    If m_profileLevel < 7 Then
        Debug.Print FromProcedureName & ": ICoreWebView2Profile7が未対応です(m_profileLevel=" & m_profileLevel & ")"
        Exit Function
    End If

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ2: 完了通知を受け取る配達員(ハンドラー)を準備する】
    '-----------------------------------------------------------------------------------------------
    ' 2-1. 共有スロットプールから「拡張機能完了用(HK_AddBrowserExtensionCompleted)」の空きブースを借りる
    Set m_ext.addExt.handler = WebView2Thunks.AcquireHandlerFor(HK_AddBrowserExtensionCompleted, Me)
    If m_ext.addExt.handler Is Nothing Then
        Debug.Print FromProcedureName & ": ハンドラのAcquireに失敗しました"
        Exit Function
    End If

    ' 2-2. C++に見せるための「お面(Vtableポインタ)」の先頭アドレスを計算
    Dim pHandlerObj As LongPtr
    pHandlerObj = m_ext.addExt.handler.Slot + WebView2Thunks.VTABLE_OBJ_OFFSET

    ' 2-3. 結果受け取り用の変数をリセットし、「待機中(Whs_Waiting)」の看板を立てる
    m_ext.addExt.resultId = vbNullString
    m_ext.addExt.state = Whs_Waiting

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ3: C++(Profile7)に拡張機能の追加を命令する！】
    '-----------------------------------------------------------------------------------------------
    ' Profile7の Vtable出席番号[22]: AddBrowserExtension を叩く！
    '  - 第1引数: 拡張機能フォルダのパス文字列(StrPtr)
    '  - 第2引数: さっき作った配達員の住所(pHandlerObj)
    Dim hr As Long
    hr = WebView2Thunks.dcf(m_pProfile, 22, "Profile7.AddBrowserExtension", _
                               StrPtr(extensionFolderPath), pHandlerObj)
    If hr <> 0 Then
        Debug.Print FromProcedureName & ": AddBrowserExtension失敗 hr=&H" & Hex(hr)
        m_lastError = hr
        m_ext.addExt.state = Whs_Failed
        Exit Function
    End If

    '4. 依頼完了サインを返す
    IssueAddBrowserExtension = True
End Function

'***************************************************************************************************
'* 機能　　：インストール済み拡張機能一覧の取得要求を発行します(Profile7以降)
'---------------------------------------------------------------------------------------------------
'* 返り値  ：要求の発行に成功したか(完了自体はまだ先。完了通知は`WebView2Host.OnGetBrowserExtensionsCompleted`に届く)
'***************************************************************************************************
Private Function IssueGetBrowserExtensions() As Boolean
    Const FromProcedureName As String = "WebView2Host.IssueGetBrowserExtensions"


    '-----------------------------------------------------------------------------------------------
    ' 【ステップ1: 受け皿の準備 & 身分証チェック】
    '-----------------------------------------------------------------------------------------------
    ' 1-1. 前回取得した古いデータが残っていると困るので、結果格納用Dictionaryを新品にリセット！
    Set m_ext.getExt.Items = New Dictionary

    ' 1-2. 拡張機能の取得機能も「ICoreWebView2Profile7」以降でしか使えない！
    If m_profileLevel < 7 Then
        Debug.Print FromProcedureName & ": ICoreWebView2Profile7が未対応です(m_profileLevel=" & m_profileLevel & ")"
        Exit Function
    End If

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ2: 一覧取得専用の配達員(ハンドラー)を準備する】
    '-----------------------------------------------------------------------------------------------
    ' 2-1. 共有スロットプールから「一覧取得完了用(HK_GetBrowserExtensionsCompleted)」の空き部屋を借りる
    Set m_ext.getExt.handler = WebView2Thunks.AcquireHandlerFor(HK_GetBrowserExtensionsCompleted, Me)
    If m_ext.getExt.handler Is Nothing Then
        Debug.Print FromProcedureName & ": ハンドラのAcquireに失敗しました"
        Exit Function
    End If

    ' 2-2. C++に見せるための「お面(Vtableポインタ)」のアドレスを計算
    Dim pHandlerObj As LongPtr
    pHandlerObj = m_ext.getExt.handler.Slot + WebView2Thunks.VTABLE_OBJ_OFFSET

    ' 2-3. 「待機中(Whs_Waiting)」の看板を立てておく
    m_ext.getExt.state = Whs_Waiting

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ3: C++(Profile7)に「一覧をちょうだい！」と命令する】
    '-----------------------------------------------------------------------------------------------
    ' Profile7::GetBrowserExtensions を叩く！
    ' (引数は結果を届けてもらうための配達員の住所 pHandlerObj だけ！)
    Dim hr As Long
    hr = WebView2Thunks.dcf(m_pProfile, 23, "Profile7.GetBrowserExtensions", pHandlerObj)
    If hr <> 0 Then
        Debug.Print FromProcedureName & ": GetBrowserExtensions失敗 hr=&H" & Hex(hr)
        m_lastError = hr
        m_ext.getExt.state = Whs_Failed
        Exit Function
    End If

    '4. 依頼完了サインを返す
    IssueGetBrowserExtensions = True
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Profile9
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：Service Worker内からWebView固有のスクリプトAPIを使えるようにするか切り替えます
'            (Profile9以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let WebViewScriptApisEnabledForServiceWorkers(v As Boolean)
    If m_profileLevel < 9 Then Exit Property
    WebView2Thunks.dcf m_pProfile, 28, "Profile9.put_AreWebViewScriptApisEnabledForServiceWorkers", CLng(v * -1)
End Property



'***************************************************************************************************
'                          ■■■ ICoreWebView2BrowserExtension 一式 ■■■
'***************************************************************************************************
'* 機能　　：指定の`ICoreWebView2BrowserExtension`に対する、アンインストール要求を発行します
'---------------------------------------------------------------------------------------------------
'* 引数    ：pTarget  対象の`ICoreWebView2BrowserExtension`ポインタ(借用。呼び出し元がRelease管理する)
'* 返り値  ：要求の発行に成功したか(完了自体はまだ先。完了通知は`WebView2Host.OnRemoveBrowserExtensionCompleted`に届く)
'***************************************************************************************************
Private Function IssueRemoveBrowserExtension(ByVal pTarget As LongPtr) As Boolean
    Const FromProcedureName As String = "WebView2Host.IssueRemoveBrowserExtension"


    '-----------------------------------------------------------------------------------------------
    ' 【ステップ1: 完了通知を受け取る配達員(ハンドラー)を準備する】
    '-----------------------------------------------------------------------------------------------
    ' 1-1. 共有スロットプールから「削除完了用(HK_RemoveBrowserExtensionCompleted)」の空き部屋を借りる
    Set m_ext.removeExt.handler = WebView2Thunks.AcquireHandlerFor(HK_RemoveBrowserExtensionCompleted, Me)
    If m_ext.removeExt.handler Is Nothing Then
        Debug.Print FromProcedureName & ": ハンドラのAcquireに失敗しました"
        Exit Function
    End If

    ' 1-2. C++に見せるための「お面(Vtableポインタ)」のアドレスを計算
    Dim pHandlerObj As LongPtr
    pHandlerObj = m_ext.removeExt.handler.Slot + WebView2Thunks.VTABLE_OBJ_OFFSET

    ' 1-3. 「待機中(Whs_Waiting)」の看板を立てておく
    m_ext.removeExt.state = Whs_Waiting

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ2: C++(ICoreWebView2BrowserExtension)に「拡張機能アンインストールして！」と命令する】
    '-----------------------------------------------------------------------------------------------
    ' ICoreWebView2BrowserExtension::Remove を叩く！
    ' (引数は結果を届けてもらうための配達員の住所 pHandlerObj だけ！)
    Dim hr As Long
    hr = WebView2Thunks.dcf(pTarget, 5, "BrowserExtension.Remove", pHandlerObj)
    If hr <> 0 Then
        Debug.Print FromProcedureName & ": Remove失敗 hr=&H" & Hex(hr)
        m_lastError = hr
        m_ext.removeExt.state = Whs_Failed
        Exit Function
    End If

    '3. 依頼完了サインを返す
    IssueRemoveBrowserExtension = True
End Function



'***************************************************************************************************
'                               ■■■ ICoreWebView2 一式 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2_3
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：ローカルフォルダを、仮想ホスト名経由(`https://<hostName>/`のような形)で配信されている
'            かのように見せます(ICoreWebView2_3以降)。`file://`のCORS/CSP制約を回避しつつ、
'            ローカルのHTML/JS/CSSをページに読み込ませたい場合に使う
'---------------------------------------------------------------------------------------------------
'* 引数    ：hostName    仮想ホスト名(例:"my.app")
'            folderPath  配信対象のローカルフォルダの絶対パス
'            accessKind  `COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND`(Long。0=Deny、1=Allow、
'            2=DenyCors。`CDPCoreViaWebView2`側の`WV2HostResourceAccessKind`Enum使用を推奨)
'* 返り値  ：成功可否
'***************************************************************************************************
Public Function SetVirtualHostNameToFolderMapping(ByVal hostname As String, ByVal FolderPath As String, ByVal accessKind As Long) As Boolean
    If m_viewLevel < 3 Then Exit Function

    Dim hr As Long
    hr = WebView2Thunks.dcf(m_pView, 71, "View3.SetVirtualHostNameToFolderMapping", _
        StrPtr(hostname), StrPtr(FolderPath), accessKind)
    SetVirtualHostNameToFolderMapping = (hr = 0)
End Function

'***************************************************************************************************
'* 機能　　：`SetVirtualHostNameToFolderMapping`で設定したマッピングを解除します(ICoreWebView2_3以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：hostName  解除する仮想ホスト名
'***************************************************************************************************
Public Sub ClearVirtualHostNameToFolderMapping(ByVal hostname As String)
    If m_viewLevel < 3 Then Exit Sub
    WebView2Thunks.dcf m_pView, 72, "View3.ClearVirtualHostNameToFolderMapping", StrPtr(hostname)
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2_6
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：タスクマネージャーウィンドウを開きます(ICoreWebView2_6以降)
'---------------------------------------------------------------------------------------------------
'* 注意事項：Chromiumの各プロセス(レンダラー/GPU等)のCPU/メモリ使用状況を確認できる、
'            開発時のデバッグ用ウィンドウです
'***************************************************************************************************
Public Sub OpenTaskManagerWindow()
    If m_viewLevel < 6 Then Exit Sub
    WebView2Thunks.dcf m_pView, 79, "View6.OpenTaskManagerWindow"
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2_8
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：WebView2インスタンス全体の音声出力をミュートするか切り替えます(ICoreWebView2_8以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueでミュート、Falseで解除
'***************************************************************************************************
Property Let IsMuted(v As Boolean)
    If m_viewLevel < 8 Then Exit Property
    WebView2Thunks.dcf m_pView, 84, "View8.put_IsMuted", CLng(v * -1)
End Property

'***************************************************************************************************
'* 機能　　：現在ドキュメントが音声を再生中かどうかを取得します(ICoreWebView2_8以降)
'---------------------------------------------------------------------------------------------------
'* 注意事項：読み取り専用(SDK仕様上、書き込みはできない)
'***************************************************************************************************
Property Get IsDocumentPlayingAudio() As Boolean
    If m_viewLevel < 8 Then Exit Property
    IsDocumentPlayingAudio = WebView2Thunks.GetBoolProperty(m_pView, 87, "View8.get_IsDocumentPlayingAudio")
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2_9
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：既定のダウンロードダイアログ(WebView2標準のダウンロード完了通知UI)を開きます
'            (ICoreWebView2_9以降)
'***************************************************************************************************
Public Sub OpenDefaultDownloadDialog()
    If m_viewLevel < 9 Then Exit Sub
    WebView2Thunks.dcf m_pView, 91, "View9.OpenDefaultDownloadDialog"
End Sub

'***************************************************************************************************
'* 機能　　：既定のダウンロードダイアログを閉じます(ICoreWebView2_9以降)
'***************************************************************************************************
Public Sub CloseDefaultDownloadDialog()
    If m_viewLevel < 9 Then Exit Sub
    WebView2Thunks.dcf m_pView, 92, "View9.CloseDefaultDownloadDialog"
End Sub

'***************************************************************************************************
'* 機能　　：既定のダウンロードダイアログを表示するコーナー位置を設定します(ICoreWebView2_9以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  `COREWEBVIEW2_DEFAULT_DOWNLOAD_DIALOG_CORNER_ALIGNMENT`(Long。0=TopLeft、
'            1=TopRight(既定)、2=BottomLeft、3=BottomRight。`CDPCoreViaWebView2`側の
'            `WV2DownloadDialogCornerAlignment`Enum使用を推奨)
'***************************************************************************************************
Property Let DefaultDownloadDialogCornerAlignment(v As Long)
    If m_viewLevel < 9 Then Exit Property
    WebView2Thunks.dcf m_pView, 94, "View9.put_DefaultDownloadDialogCornerAlignment", v
End Property

'***************************************************************************************************
'* 機能　　：既定のダウンロードダイアログの、コーナーからのオフセットを設定します(ICoreWebView2_9以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：x/y  `DefaultDownloadDialogCornerAlignment`で指定したコーナーからのオフセット(px)
'---------------------------------------------------------------------------------------------------
'* 注意事項：x/yが対等な1組の値であり「代入値」の形にそぐわないため、`Property Let`ではなく
'            `Sub`にしている(`SetDefaultBackgroundColor`と同じ理由)
'***************************************************************************************************
Public Sub SetDefaultDownloadDialogMargin(x As Long, y As Long)
    If m_viewLevel < 9 Then Exit Sub
    Dim pt As WebView2Thunks.POINT
    pt.x = x: pt.y = y
    WebView2Thunks.dcf m_pView, 96, "View9.put_DefaultDownloadDialogMargin", VarPtr(pt)
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2_12
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：リンクホバー時のリンク先URL等、ステータスバーに表示される文字列を取得します
'            (ICoreWebView2_12以降)
'---------------------------------------------------------------------------------------------------
'* 注意事項：読み取り専用。`StatusBarEnabled=False`にしていても値自体は取得できる
'***************************************************************************************************
Property Get StatusBarText() As String
    If m_viewLevel < 12 Then Exit Property
    StatusBarText = WebView2Thunks.GetStringProperty(m_pView, 104, "View12.get_StatusBarText")
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2_15
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：現在のページのfavicon URI(Chromium本体が解決した実際のURI。`<link rel="icon">`未指定時の
'            `/favicon.ico`フォールバックも考慮済み)を取得します(ICoreWebView2_15以降)
'---------------------------------------------------------------------------------------------------
'* 注意事項：読み取り専用。未解決/存在しない場合は空文字
'***************************************************************************************************
Property Get FaviconUri() As String
    If m_viewLevel < 15 Then Exit Property
    FaviconUri = WebView2Thunks.GetStringProperty(m_pView, 111, "View15.get_FaviconUri")
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2_16
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：印刷ダイアログを表示します(ICoreWebView2_16以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：kind  `COREWEBVIEW2_PRINT_DIALOG_KIND`(Long。0=Browser、1=System。
'            `CDPCoreViaWebView2`側の`WV2PrintDialogKind`Enum使用を推奨)
'***************************************************************************************************
Public Sub ShowPrintUI(kind As Long)
    If m_viewLevel < 16 Then Exit Sub
    WebView2Thunks.dcf m_pView, 114, "View16.ShowPrintUI", kind
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2_19
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：WebViewの目標メモリ使用量レベルを設定します(ICoreWebView2_19以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  `COREWEBVIEW2_MEMORY_USAGE_TARGET_LEVEL`(Long。0=Normal、1=Low。
'            `CDPCoreViaWebView2`側の`WV2MemoryUsageTargetLevel`Enum使用を推奨)
'***************************************************************************************************
Property Let MemoryUsageTargetLevel(v As Long)
    If m_viewLevel < 19 Then Exit Property
    WebView2Thunks.dcf m_pView, 120, "View19.put_MemoryUsageTargetLevel", v
End Property

'***************************************************************************************************
'* 機能　　：WebViewの目標メモリ使用量レベルを取得します(ICoreWebView2_19以降)
'***************************************************************************************************
Property Get MemoryUsageTargetLevel() As Long
    If m_viewLevel < 19 Then Exit Property
    MemoryUsageTargetLevel = WebView2Thunks.GetLongProperty(m_pView, 119, "View19.get_MemoryUsageTargetLevel")
End Property



'***************************************************************************************************
'                              ■■■ 初期化/後始末 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    '1. ステータスリセット
    m_core.env.state = Whs_Created
    m_core.ctrl.state = Whs_Created

    '2. 実行マシンの「1秒間の振動数（周波数）」を取得して、bas側で記憶し、タイムアウト判定準備をする
    CDPHelpers.InitQueryPerformanceFrequency
End Sub

Private Sub Class_Terminate()
    If Not m_closed Then CLOSE_
End Sub
