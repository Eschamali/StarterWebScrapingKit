VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrowserCapabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================================
'             指定Classにある`OpenBrowser`系メソッドの引数に渡す、ブラウザ起動時のオプション設定オブジェクトです。
'                       いわゆる、使い切り[ワンショット]的な扱いです。
'       最終到達地点である`CDPCore/CDPCoreViaWebSocket.OpenBrowserAndConnectCDP`にて、`Nothing`によるワンショット仕様にしてます。
'
'  ・各設定の既定値は、`Class_Initialize`にて、設定シートから読み取ります
'  ・既定値を変えたい場合は、`Set_○○`プロパティで上書きしてください
'  ・BiDi用の`session.new`パラメーター(`sessionCapabilitiesRequest`)は、このオブジェクトとは別に引数で渡します
'
'  仕組み上、下記の2通りとなります
'  ・`Set_`始まり   ：ユーザー側
'  ・`ClsSet_`始まり：Class内都合による動的設定
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                 ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 全てのプロトコルコア機能
'---------------------------------------------------------------------------------------------------
Private m_CDPCore               As CDPCore
Private m_WebDriverBiDiCore     As WebDriverBiDiCore

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 初期ブラウザ起動に関する設定
'---------------------------------------------------------------------------------------------------
Private m_CustomChromiumPath    As String       'ブラウザのexeフルパス。空なら`UseBrowserPreset`に従って自動判定
Private m_UserDataDirName       As String       '`--user-data-dir`につけるユーザーディレクトリ名。識別名称としても使う
Private m_AddArguments          As String       '追加の起動引数
Private m_UseBrowserPreset      As Long         '使用するブラウザ種別
Private m_UseWebSocket          As Boolean      '`True`：WebSocket(Port)通信 / `False`：Pipe通信
Private m_StartWindowMode       As WinState     '初期起動ウィンドウ状態
Private m_StartURL              As String       '`--app`につけるURL
Private m_SplashScreenMode      As Boolean      '`True`で、`--app`に渡すURLがスプラッシュ画面としてかっこよくなる

'---------------------------------------------------------------------------------------------------
' [ SECTION ] chromium-bidi関する設定
'---------------------------------------------------------------------------------------------------
Private m_ViewTab_BiDiCDPMapper     As Boolean      '`BiDi-CDP Mapper`のコンソールタブを表示するか
Private m_ViewDebugLogForBiDiMapper As Boolean      'Mapper内部のデバッグログを記録するか
Private m_sessionCapabilityResponse As Dictionary   'BiDi始動時の挙動を決めます

'---------------------------------------------------------------------------------------------------
' [ SECTION ] その他の設定
'---------------------------------------------------------------------------------------------------
Private m_EnableUserAgentOverride   As Boolean  'このツール独自のユーザーエージェントを刻印するか



'***************************************************************************************************
'                                 ■■■ 設定メソッド ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 初期ブラウザ起動に関する設定
'---------------------------------------------------------------------------------------------------
Property Let Set_CustomChromiumPath(FullPath As String)
    m_CustomChromiumPath = FullPath
End Property

Property Let Set_UserDataDirName(UserName As String)
    m_UserDataDirName = UserName
End Property

Property Let Set_AdditionalBrowserArguments(Arguments As String)
    m_AddArguments = Arguments
End Property

Property Let Set_UseBrowserPreset(Preset As Long)
    m_UseBrowserPreset = Preset
End Property

Property Let Set_UseWebSocket(Flag As Boolean)
    m_UseWebSocket = Flag
End Property

Property Let Set_StartWindowMode(Mode As WinState)
    m_StartWindowMode = Mode
End Property

Property Let Set_StartURL(Url As String)
    m_StartURL = Url
End Property

Property Let ClsSet_SplashScreenMode(Flag As Boolean)
    m_SplashScreenMode = Flag
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] chromium-bidi関する設定
'---------------------------------------------------------------------------------------------------
Property Let Set_BiDiViewTab_BiDiCDPMapper(Flag As Boolean)
    m_ViewTab_BiDiCDPMapper = Flag
End Property

Property Let Set_BiDiViewDebugLogForBiDiCDPMapper(Flag As Boolean)
    m_ViewDebugLogForBiDiMapper = Flag
End Property

Property Set Set_BiDiSessionCapabilitiesRequest(cap As Dictionary)
    Set m_sessionCapabilityResponse = cap
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] その他の設定
'---------------------------------------------------------------------------------------------------
Property Let Set_EnableUserAgentOverride(Flag As Boolean)
    m_EnableUserAgentOverride = Flag
End Property



'***************************************************************************************************
'                               ■■■ 処理側が読み取るメソッド ■■■
'***************************************************************************************************
Property Get ClsGet_CDPCore() As CDPCore
    '呼ばれるまで実体化させないようにし、初期化後は同じオブジェクトポインターとして振る舞う
    If m_CDPCore Is Nothing Then Set m_CDPCore = New CDPCore
    Set ClsGet_CDPCore = m_CDPCore
End Property

Property Get ClsGet_WebDriverBiDiCore() As WebDriverBiDiCore
    '呼ばれるまで実体化させないようにし、初期化後は同じオブジェクトポインターとして振る舞う
    If m_WebDriverBiDiCore Is Nothing Then Set m_WebDriverBiDiCore = New WebDriverBiDiCore
    Set ClsGet_WebDriverBiDiCore = m_WebDriverBiDiCore
End Property

Property Get Set_CustomChromiumPath() As String
    Set_CustomChromiumPath = m_CustomChromiumPath
End Property

Property Get Set_UserDataDirName() As String
    Set_UserDataDirName = m_UserDataDirName
End Property

Property Get Set_AdditionalBrowserArguments() As String
    Set_AdditionalBrowserArguments = m_AddArguments
End Property

Property Get Set_UseBrowserPreset() As Long
    Set_UseBrowserPreset = m_UseBrowserPreset
End Property

Property Get Set_UseWebSocket() As Boolean
    Set_UseWebSocket = m_UseWebSocket
End Property

Property Get Set_StartWindowMode() As WinState
    Set_StartWindowMode = m_StartWindowMode
End Property

Property Get Set_StartURL() As String
    Set_StartURL = m_StartURL
End Property

Property Get ClsSet_SplashScreenMode() As Boolean
    ClsSet_SplashScreenMode = m_SplashScreenMode
End Property

Property Get Set_BiDiViewTab_BiDiCDPMapper() As Boolean
    Set_BiDiViewTab_BiDiCDPMapper = m_ViewTab_BiDiCDPMapper
End Property

Property Get Set_BiDiViewDebugLogForBiDiCDPMapper() As Boolean
    Set_BiDiViewDebugLogForBiDiCDPMapper = m_ViewDebugLogForBiDiMapper
End Property

Property Get Set_BiDiSessionCapabilitiesRequest() As Dictionary
    Set Set_BiDiSessionCapabilitiesRequest = m_sessionCapabilityResponse
End Property

Property Get Set_EnableUserAgentOverride() As Boolean
    Set_EnableUserAgentOverride = m_EnableUserAgentOverride
End Property



'***************************************************************************************************
'                               ■■■ 初期化と後始末 ■■■
'***************************************************************************************************
'* 機能　　：設定シートの各セルから設定値を取得し、既定値として適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = "BrowserCapabilities.Class_Initialize"


    With ShSetting01_StartBrowser
        '1. 初期ブラウザ起動に関する設定
        m_CustomChromiumPath = .GetSettingCDP(SettingCDP.CustomChromiumPath, FromProcedureName)
        m_UserDataDirName = .GetSettingCDP(SettingCDP.UserDataDirName, FromProcedureName)
        m_AddArguments = .GetSettingCDP(SettingCDP.AdditionalBrowserArguments, FromProcedureName)
        m_UseBrowserPreset = .GetSettingCDP(SettingCDP.UseBrowserPreset, FromProcedureName)
        m_UseWebSocket = .GetSettingCDP(SettingCDP.UseWebSocket, FromProcedureName)
        m_StartWindowMode = .GetSettingCDP(SettingCDP.StartWindowMode, FromProcedureName)

        '2. chromium-BiDiの設定
        m_ViewDebugLogForBiDiMapper = .GetSettingCDP(SettingCDP.ViewDebugLogForBiDiCDPMapper, FromProcedureName)
        m_ViewTab_BiDiCDPMapper = .GetSettingCDP(SettingCDP.ViewTab_BiDiCDPMapper, FromProcedureName)

        '3. その他
        m_EnableUserAgentOverride = .GetSettingCDP(SettingCDP.EnableUserAgentOverride, FromProcedureName)
    End With
End Sub

Private Sub Class_Terminate()
    '念のため、明示的に破棄
    Set m_CDPCore = Nothing
    Set m_WebDriverBiDiCore = Nothing
End Sub
