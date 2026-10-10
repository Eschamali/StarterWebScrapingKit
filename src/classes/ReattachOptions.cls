VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReattachOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================================
'             指定Classにある`reattach`メソッド系の引数に渡すオプション設定オブジェクトです。
'                       いわゆる、使い切り[ワンショット]的な扱いです。
'        最終到達地点である`CDPBrowser.reattach`にて、`Nothing`によるワンショット仕様にしてます。
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
' [ SECTION ] 起動済みブラウザに接続する際の関連設定
'---------------------------------------------------------------------------------------------------
Private m_destruction               As Boolean  'ハンドル先のメモリにたまった古いデータをどうするか？
Private m_EnableUserAgentOverride   As Boolean  'このツール独自のユーザーエージェントを刻印するか

'---------------------------------------------------------------------------------------------------
' [ SECTION ] chromium-bidi関する再設定
'---------------------------------------------------------------------------------------------------
Private m_ViewTab_BiDiCDPMapper     As Boolean      '`BiDi-CDP Mapper`のコンソールタブを表示するか
Private m_ViewDebugLogForBiDiMapper As Boolean      'Mapper内部のデバッグログを記録するか
Private m_sessionCapabilityResponse As Dictionary   'BiDi再始動時の挙動を決めます



'***************************************************************************************************
'                                 ■■■ 設定メソッド ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 制御経路設定
' 現在は、3つの経路をサポートします
' ・Pipe(remote-debugging-pipe)
' ・Port(remote-debugging-port)
' ・WebView2(CallDevToolsProtocolMethod/ForSession)
'
' New による上書きにより必ずどれか、1つの経路のみとなります
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：Pipeモードでのメッセージ取得/解析としてリアタッチします
'---------------------------------------------------------------------------------------------------
'* Let引数 ：UserName   テーブルから読み取る識別名称
'***************************************************************************************************
Property Let Set_PipeMode(UserName As String)
    Set m_CDPCore = New CDPCore
    m_CDPCore.RemoteDebuggingPipeMode UserName
End Property

'***************************************************************************************************
'* 機能　　：WebSocketモードでのメッセージ取得/解析に切り替えます
'---------------------------------------------------------------------------------------------------
'* Let引数 ：UserName   テーブルから読み取る識別名称
'* 引数    ：接続済みの`WebSocket`オブジェクト
'***************************************************************************************************
Property Let Set_WebSocketMode(ws As CDPCoreViaWebSocket, UserName As String)
    Set m_CDPCore = New CDPCore
    Set m_CDPCore.RemoteDebuggingPortMode(UserName) = ws
End Property

'***************************************************************************************************
'* 機能　　：WebView2モードでのメッセージ取得/解析に切り替えます
'---------------------------------------------------------------------------------------------------
'* Let引数 ：UserName   テーブルから読み取る識別名称
'* 引数    ：接続済みの`WebView2`オブジェクト
'***************************************************************************************************
Property Let Set_WebView2Mode(wv As CDPCoreViaWebView2, UserName As String)
    Set m_CDPCore = New CDPCore
    Set m_CDPCore.CallDevToolsProtocolMethodForSessionMode(UserName) = wv
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
' [ SECTION ] その他
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：リアタッチ後に、内部の古いデータを破棄するか？
'---------------------------------------------------------------------------------------------------
'* Let引数 ：`True` ：メモリにたまった古いデータを破棄します
'            `False`：メモリにたまった古いデータを破棄せずに、イベント情報を受け取ります
'***************************************************************************************************
Property Let Set_Destruction(Flag As Boolean)
    m_destruction = Flag
End Property

Property Let Set_EnableUserAgentOverride(Flag As Boolean)
    m_EnableUserAgentOverride = Flag
End Property



'***************************************************************************************************
'                               ■■■ 処理側が読み取るメソッド ■■■
'***************************************************************************************************
Property Get ClsGet_CDPCore() As CDPCore
    '1. `Nothing`のままの場合は、エラーで止めます
    If m_CDPCore Is Nothing Then Err.Raise vbObjectError, Description:="No control paths have been configured for the rear touch."

    '2. 実際に返す
    Set ClsGet_CDPCore = m_CDPCore
End Property

Property Get ClsGet_WebDriverBiDiCore() As WebDriverBiDiCore
    '呼ばれるまで実体化させないようにし、初期化後は同じオブジェクトポインターとして振る舞う
    If m_WebDriverBiDiCore Is Nothing Then Set m_WebDriverBiDiCore = New WebDriverBiDiCore
    Set ClsGet_WebDriverBiDiCore = m_WebDriverBiDiCore
End Property

Property Get Set_Destruction() As Boolean
    Set_Destruction = m_destruction
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
'                               ■■■ 初期値設定 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = "ReattachOptions.Class_Initialize"


    '1. 既定値を取得
    '※リアタッチに影響のある設定値のみ列挙
    With ShSetting01_StartBrowser
        '1-1. chromium-BiDiの再設定
        m_ViewDebugLogForBiDiMapper = .GetSettingCDP(SettingCDP.ViewDebugLogForBiDiCDPMapper, FromProcedureName)
        m_ViewTab_BiDiCDPMapper = .GetSettingCDP(SettingCDP.ViewTab_BiDiCDPMapper, FromProcedureName)

        '1-2. その他
        m_EnableUserAgentOverride = .GetSettingCDP(SettingCDP.EnableUserAgentOverride, FromProcedureName)
        m_destruction = True    '古いデータは破棄させます
    End With

    '2. 共通設定を再適用
    CDPHelpers.InitQueryPerformanceFrequency    '実行マシンの「1秒間の振動数（周波数）」を取得して、bas側で記憶し、各種タイムアウト判定準備をする
    CDPHelpers.InitLogger                       '初期ログ設定をbas側で記憶する
End Sub



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
Private Sub Class_Terminate()
    '念のため、明示的に破棄
    Set m_CDPCore = Nothing
    Set m_WebDriverBiDiCore = Nothing
End Sub
