VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2EnvOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   `ICoreWebView2EnvironmentOptions`(+`Options2`/`3`/`5`/`6`/`7`/`8`)を実装する
'   VBA側COMオブジェクトです。`BuildNativePointer`が返すポインタを、`WebView2Host.RunWebView2`の
'   引数経由で`CreateCoreWebView2EnvironmentWithOptions`の`environmentOptions`引数として渡します。
'   (`Options4`[CustomSchemeRegistrations]は配列を扱う特殊な形のため未実装)
'
'   ★使い方★
'     `CDPCoreViaWebView2.EnvironmentOptions`(このクラスのインスタンス)のプロパティを、
'     必ず`ConnectCDP`を呼ぶ**前**に設定すること。Environment生成時にしか読まれない値のため、
'     生成後に変更しても反映されない。`m_Host`(`WebView2Host`)の生死とは無関係に
'     `CDPCoreViaWebView2`側で直接保持するため、接続前から自由にアクセスできる。
'
'   ★実装方式★
'     ネイティブ側のvtable組み立て/QueryInterface実装は`WebView2Thunks`の「EnvOpt」
'     セクションが担う。このクラスは値の保持と、WebView2からのget_/put_コールバックの
'     受け皿(`EnvOpt_OnGetXxx`/`EnvOpt_OnPutXxx`、`WebView2CallbackHandler`から動的束縛で
'     呼ばれるため全てPublic Sub必須)に徹する。
'***************************************************************************************************
Option Explicit



#If Win64 Then
    ' OK
#Else
    #Error "このクラスは64ビットVBA(x64)が必要です"
#End If



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'* `ChannelSearchKind`用(`COREWEBVIEW2_CHANNEL_SEARCH_KIND`)
Public Enum WV2ChannelSearchKind
    ChannelSearch_MostStable = 0     ' 既定。安定版から探す
    ChannelSearch_LeastStable = 1    ' 不安定版から探す
End Enum

'* `ReleaseChannels`用(`COREWEBVIEW2_RELEASE_CHANNELS`)。`Or`で組み合わせる
Public Enum WV2ReleaseChannels
    Channel_Stable = &H1
    Channel_Beta = &H2
    Channel_Dev = &H4
    Channel_Canary = &H8
End Enum

'* `ScrollBarStyle`用(`COREWEBVIEW2_SCROLLBAR_STYLE`)
Public Enum WV2ScrollBarStyle
    ScrollBar_Default = 0            ' 既定
    ScrollBar_FluentOverlay = 1
End Enum



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'`ICoreWebView2EnvironmentOptions`管理
Private m_pNative As LongPtr    'Vtableポインター
Private m_closed  As Boolean    'Closeサイン

'各種設定値保持
Private m_additionalBrowserArguments             As String
Private m_language                               As String
Private m_targetCompatibleBrowserVersion         As String
Private m_allowSingleSignOnUsingOSPrimaryAccount As Boolean
Private m_areBrowserExtensionsEnabled            As Boolean
Private m_exclusiveUserDataFolderAccess As Boolean
Private m_isCustomCrashReportingEnabled As Boolean
Private m_enableTrackingPrevention      As Boolean
Private m_channelSearchKind             As Long
Private m_releaseChannels               As Long
Private m_scrollBarStyle                As Long



'***************************************************************************************************
'                               ■■■ ICoreWebView2EnvironmentOptions 一式 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザプロセス起動時に追加するコマンドライン引数を指定します
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2EnvironmentOptions
'---------------------------------------------------------------------------------------------------
Property Let AdditionalBrowserArguments(ByVal v As String)
    m_additionalBrowserArguments = v
End Property

'***************************************************************************************************
'* 機能　　：ブラウザのUI言語を指定します(既定は空文字=OS既定に従う)
'***************************************************************************************************
Property Let Language(ByVal v As String)
    m_language = v
End Property

'***************************************************************************************************
'* 機能　　：互換性チェック対象のブラウザバージョンを指定します(既定は空文字=チェックなし)
'***************************************************************************************************
Property Let TargetCompatibleBrowserVersion(ByVal v As String)
    m_targetCompatibleBrowserVersion = v
End Property

'***************************************************************************************************
'* 機能　　：OSのプライマリアカウントを使ったシングルサインオンの許可可否を切り替えます(既定False)
'***************************************************************************************************
Property Let AllowSingleSignOnUsingOSPrimaryAccount(ByVal v As Boolean)
    m_allowSingleSignOnUsingOSPrimaryAccount = v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2EnvironmentOptions2
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：同じuser-data-folderから、他プロセスが別のWebView2を作って同じブラウザプロセス
'            インスタンスに相乗りすることを禁止するか切り替えます(既定False=相乗り許可)
'***************************************************************************************************
Property Let ExclusiveUserDataFolderAccess(ByVal v As Boolean)
    m_exclusiveUserDataFolderAccess = v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2EnvironmentOptions3
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：Trueにすると、クラッシュ時のデータをMicrosoftへ送らなくなります(既定False=OSの同意設定に従う)
'***************************************************************************************************
Property Let IsCustomCrashReportingEnabled(ByVal v As Boolean)
    m_isCustomCrashReportingEnabled = v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2EnvironmentOptions4
'---------------------------------------------------------------------------------------------------
'新しい自前COMオブジェクトを丸ごと1本(`WebView2EnvOptions.cls`相当)必要なため、保留

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2EnvironmentOptions5
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：トラッキング防止機能(サードパーティCookie制限等)の有効可否を切り替えます(既定True)
'***************************************************************************************************
Property Let EnableTrackingPrevention(ByVal v As Boolean)
    m_enableTrackingPrevention = v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2EnvironmentOptions6
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：ブラウザ拡張機能(Extensions)の使用可否を切り替えます(既定False)
'***************************************************************************************************
Property Let AreBrowserExtensionsEnabled(ByVal v As Boolean)
    m_areBrowserExtensionsEnabled = v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2EnvironmentOptions7
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：インストール済みEdgeチャンネルを探す順序を指定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：v  `COREWEBVIEW2_CHANNEL_SEARCH_KIND`(既定0=MostStable。安定版から探す。
'            1=LeastStableで逆順)
'***************************************************************************************************
Property Let ChannelSearchKind(ByVal v As WV2ChannelSearchKind)
    m_channelSearchKind = v
End Property

'***************************************************************************************************
'* 機能　　：探索対象にするEdgeチャンネルをビットマスクで絞り込みます(既定は全チャンネル許可)
'---------------------------------------------------------------------------------------------------
'* 引数　　：v  `COREWEBVIEW2_RELEASE_CHANNELS`のビットマスク(Stable=&H1/Beta=&H2/Dev=&H4/Canary=&H8)
'***************************************************************************************************
Property Let ReleaseChannels(ByVal v As WV2ReleaseChannels)
    m_releaseChannels = v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2EnvironmentOptions8
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：スクロールバーの見た目を切り替えます
'---------------------------------------------------------------------------------------------------
'* 引数　　：v  `COREWEBVIEW2_SCROLLBAR_STYLE`(既定0=Default。1=FluentOverlayで細い自動非表示型)
'***************************************************************************************************
Property Let ScrollBarStyle(ByVal v As WV2ScrollBarStyle)
    m_scrollBarStyle = v
End Property



'***************************************************************************************************
'                              ■■■ ネイティブオブジェクトの構築/破棄 ■■■
'***************************************************************************************************
'* 機能　　：ネイティブ側のfakeオブジェクトを構築し、そのポインタを返します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`CreateCoreWebView2EnvironmentWithOptions`の`environmentOptions`引数にそのまま渡せる
'***************************************************************************************************
Friend Function BuildNativePointer() As LongPtr
    If m_pNative = 0 Then
        m_pNative = WebView2Thunks.EnvOpt_CreateNative(Me)
    End If
    BuildNativePointer = m_pNative
End Function

'***************************************************************************************************
'* 機能　　：`TargetCompatibleBrowserVersion`の既定値を設定します
'---------------------------------------------------------------------------------------------------
'* 注意事項：★実機検証済み★ 公式のC++ラッパー(`WebView2EnvironmentOptions.h`)はこの値を常に
'            (SDKバンドル版の)具体的なバージョン文字列で初期化している。ここを空文字
'            (nullptr相当)のままにすると`CreateCoreWebView2EnvironmentWithOptions`が
'            hr=&H80070057(E_INVALIDARG)で即失敗することを実機で確認した。SDKバンドル版に
'            合わせた値を既定にすることで解消する
'***************************************************************************************************
Private Sub Class_Initialize()
    m_targetCompatibleBrowserVersion = "152.0.4181.0"

    ' --- Options2/3/5/7/8の既定値。公式サンプル実装(WebView2EnvironmentOptions.h)と同じ値にする ---
    ' `ExclusiveUserDataFolderAccess`/`IsCustomCrashReportingEnabled`/`ChannelSearchKind`
    ' (MOST_STABLE=0)/`ScrollBarStyle`(DEFAULT=0)はVBAのFalse/0既定のままで公式既定と一致するため
    ' 明示設定不要。`EnableTrackingPrevention`(既定True)と`ReleaseChannels`(既定は全チャンネル
    ' 許可=15)だけ明示的に上書きする
    m_enableTrackingPrevention = True
    m_releaseChannels = &H1 Or &H2 Or &H4 Or &H8   ' STABLE|BETA|DEV|CANARY
End Sub

'***************************************************************************************************
'* 機能　　：自分(VBA所有者側)が持っていた分の参照を手放します
'---------------------------------------------------------------------------------------------------
'* 注意事項：ここで即座に実メモリが解放されるとは限らない。`WebView2Thunks`側が参照カウント
'            方式で管理しており、WebView2Loaderがまだ内部でQueryInterface済みの参照を保持している
'            間は、実解放(`VirtualFree`)は行われない。両方(VBA側/WebView2Loader側)が全て
'            手放した時点で初めて解放されるため、呼ぶタイミングを厳密に見計らう必要はない
'***************************************************************************************************
Friend Sub CLOSE_()
    If m_closed Then Exit Sub

    If m_pNative <> 0 Then
        WebView2Thunks.EnvOpt_DestroyNative m_pNative
        m_pNative = 0
    End If

    m_closed = True
End Sub

Private Sub Class_Terminate()
    If Not m_closed Then CLOSE_
End Sub



'***************************************************************************************************
'                  ■■■ WebView2からのコールバック受け皿(動的束縛経由) ■■■
'***************************************************************************************************
'* 注意事項：★重要★ 以下は全て`WebView2CallbackHandler.Handler_Invoke`から`owner As Object`
'            経由(動的束縛)で呼ばれるため、全てPublic Subでなければならない
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] ICoreWebView2EnvironmentOptions
'---------------------------------------------------------------------------------------------------
Public Sub EnvOpt_OnGetAdditionalBrowserArguments(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteStringOut pOut, m_additionalBrowserArguments
End Sub

Public Sub EnvOpt_OnPutAdditionalBrowserArguments(ByVal pValue As LongPtr)
    If pValue <> 0 Then m_additionalBrowserArguments = WebView2Thunks.PtrToString(pValue)
End Sub

Public Sub EnvOpt_OnGetLanguage(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteStringOut pOut, m_language
End Sub

Public Sub EnvOpt_OnPutLanguage(ByVal pValue As LongPtr)
    If pValue <> 0 Then m_language = WebView2Thunks.PtrToString(pValue)
End Sub

Public Sub EnvOpt_OnGetTargetCompatibleBrowserVersion(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteStringOut pOut, m_targetCompatibleBrowserVersion
End Sub

Public Sub EnvOpt_OnPutTargetCompatibleBrowserVersion(ByVal pValue As LongPtr)
    If pValue <> 0 Then m_targetCompatibleBrowserVersion = WebView2Thunks.PtrToString(pValue)
End Sub

Public Sub EnvOpt_OnGetAllowSingleSignOnUsingOSPrimaryAccount(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteBoolOut pOut, m_allowSingleSignOnUsingOSPrimaryAccount
End Sub

Public Sub EnvOpt_OnPutAllowSingleSignOnUsingOSPrimaryAccount(ByVal rawValue As LongPtr)
    m_allowSingleSignOnUsingOSPrimaryAccount = (CLngLng(rawValue) And &HFFFFFFFF^) <> 0
End Sub

Public Sub EnvOpt_OnGetAreBrowserExtensionsEnabled(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteBoolOut pOut, m_areBrowserExtensionsEnabled
End Sub

Public Sub EnvOpt_OnPutAreBrowserExtensionsEnabled(ByVal rawValue As LongPtr)
    m_areBrowserExtensionsEnabled = (CLngLng(rawValue) And &HFFFFFFFF^) <> 0
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ICoreWebView2EnvironmentOptions2
'---------------------------------------------------------------------------------------------------
Public Sub EnvOpt_OnGetExclusiveUserDataFolderAccess(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteBoolOut pOut, m_exclusiveUserDataFolderAccess
End Sub

Public Sub EnvOpt_OnPutExclusiveUserDataFolderAccess(ByVal rawValue As LongPtr)
    m_exclusiveUserDataFolderAccess = (CLngLng(rawValue) And &HFFFFFFFF^) <> 0
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ICoreWebView2EnvironmentOptions3
'---------------------------------------------------------------------------------------------------
Public Sub EnvOpt_OnGetIsCustomCrashReportingEnabled(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteBoolOut pOut, m_isCustomCrashReportingEnabled
End Sub

Public Sub EnvOpt_OnPutIsCustomCrashReportingEnabled(ByVal rawValue As LongPtr)
    m_isCustomCrashReportingEnabled = (CLngLng(rawValue) And &HFFFFFFFF^) <> 0
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ICoreWebView2EnvironmentOptions5
'---------------------------------------------------------------------------------------------------
Public Sub EnvOpt_OnGetEnableTrackingPrevention(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteBoolOut pOut, m_enableTrackingPrevention
End Sub

Public Sub EnvOpt_OnPutEnableTrackingPrevention(ByVal rawValue As LongPtr)
    m_enableTrackingPrevention = (CLngLng(rawValue) And &HFFFFFFFF^) <> 0
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ICoreWebView2EnvironmentOptions7
'---------------------------------------------------------------------------------------------------
Public Sub EnvOpt_OnGetChannelSearchKind(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteLongOut pOut, m_channelSearchKind
End Sub

Public Sub EnvOpt_OnPutChannelSearchKind(ByVal rawValue As LongPtr)
    m_channelSearchKind = CLng(CLngLng(rawValue) And &HFFFFFFFF^)
End Sub

Public Sub EnvOpt_OnGetReleaseChannels(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteLongOut pOut, m_releaseChannels
End Sub

Public Sub EnvOpt_OnPutReleaseChannels(ByVal rawValue As LongPtr)
    m_releaseChannels = CLng(CLngLng(rawValue) And &HFFFFFFFF^)
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ICoreWebView2EnvironmentOptions8
'---------------------------------------------------------------------------------------------------
Public Sub EnvOpt_OnGetScrollBarStyle(ByVal pOut As LongPtr)
    WebView2Thunks.EnvOpt_WriteLongOut pOut, m_scrollBarStyle
End Sub

Public Sub EnvOpt_OnPutScrollBarStyle(ByVal rawValue As LongPtr)
    m_scrollBarStyle = CLng(CLngLng(rawValue) And &HFFFFFFFF^)
End Sub
