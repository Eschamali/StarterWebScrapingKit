VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPWebView2EnvOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   `ICoreWebView2EnvironmentOptions`(+`ICoreWebView2EnvironmentOptions6`)を実装する
'   VBA側COMオブジェクトです。`CDPWebView2Host.RunWebView2`が
'   `CreateCoreWebView2EnvironmentWithOptions`へ渡す`environmentOptions`引数として使います。
'
'   ★使い方★
'     `CDPWebView2Host.EnvironmentOptions`(このクラスのインスタンス)のプロパティを、
'     必ず`RunWebView2`を呼ぶ**前**に設定すること。Environment生成時にしか読まれない値のため、
'     生成後に変更しても反映されない。
'
'   ★実装方式★
'     ネイティブ側のvtable組み立て/QueryInterface実装は`CDPWebView2Thunks`の「EnvOpt」
'     セクションが担う。このクラスは値の保持と、WebView2からのget_/put_コールバックの
'     受け皿(`EnvOpt_OnGetXxx`/`EnvOpt_OnPutXxx`、`CDPWebView2CallbackHandler`から動的束縛で
'     呼ばれるため全てPublic Sub必須)に徹する。
'***************************************************************************************************
Option Explicit



#If Win64 Then
    ' OK
#Else
    #Error "このクラスは64ビットVBA(x64)が必要です"
#End If



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
Private m_pNative As LongPtr
Private m_closed  As Boolean

Private m_additionalBrowserArguments             As String
Private m_language                               As String
Private m_targetCompatibleBrowserVersion         As String
Private m_allowSingleSignOnUsingOSPrimaryAccount As Boolean
Private m_areBrowserExtensionsEnabled            As Boolean

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
End Sub



'***************************************************************************************************
'                               ■■■ ICoreWebView2EnvironmentOptions 一式 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザプロセス起動時に追加するコマンドライン引数を指定します
'***************************************************************************************************
Public Property Get AdditionalBrowserArguments() As String
    AdditionalBrowserArguments = m_additionalBrowserArguments
End Property

Public Property Let AdditionalBrowserArguments(ByVal v As String)
    m_additionalBrowserArguments = v
End Property

'***************************************************************************************************
'* 機能　　：ブラウザのUI言語を指定します(既定は空文字=OS既定に従う)
'***************************************************************************************************
Public Property Get Language() As String
    Language = m_language
End Property

Public Property Let Language(ByVal v As String)
    m_language = v
End Property

'***************************************************************************************************
'* 機能　　：互換性チェック対象のブラウザバージョンを指定します(既定は空文字=チェックなし)
'***************************************************************************************************
Public Property Get TargetCompatibleBrowserVersion() As String
    TargetCompatibleBrowserVersion = m_targetCompatibleBrowserVersion
End Property

Public Property Let TargetCompatibleBrowserVersion(ByVal v As String)
    m_targetCompatibleBrowserVersion = v
End Property

'***************************************************************************************************
'* 機能　　：OSのプライマリアカウントを使ったシングルサインオンの許可可否を切り替えます(既定False)
'***************************************************************************************************
Public Property Get AllowSingleSignOnUsingOSPrimaryAccount() As Boolean
    AllowSingleSignOnUsingOSPrimaryAccount = m_allowSingleSignOnUsingOSPrimaryAccount
End Property

Public Property Let AllowSingleSignOnUsingOSPrimaryAccount(ByVal v As Boolean)
    m_allowSingleSignOnUsingOSPrimaryAccount = v
End Property

'***************************************************************************************************
'                              ■■■ ICoreWebView2EnvironmentOptions6 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザ拡張機能(Extensions)の使用可否を切り替えます(既定False)
'***************************************************************************************************
Public Property Get AreBrowserExtensionsEnabled() As Boolean
    AreBrowserExtensionsEnabled = m_areBrowserExtensionsEnabled
End Property

Public Property Let AreBrowserExtensionsEnabled(ByVal v As Boolean)
    m_areBrowserExtensionsEnabled = v
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
        m_pNative = CDPWebView2Thunks.EnvOpt_CreateNative(Me)
    End If
    BuildNativePointer = m_pNative
End Function

'***************************************************************************************************
'* 機能　　：ネイティブ側のリソース(スロット+ブロック本体)を明示的に解放します
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebView2LoaderがEnvironment作成のために読み終えた後(=完了コールバック確定後)に
'            呼ぶこと
'***************************************************************************************************
Friend Sub CLOSE_()
    If m_closed Then Exit Sub

    If m_pNative <> 0 Then
        CDPWebView2Thunks.EnvOpt_DestroyNative m_pNative
        m_pNative = 0
    End If

    m_closed = True
End Sub

Private Sub Class_Terminate()
    If Not m_closed Then CLOSE_
End Sub



'***************************************************************************************************
'                    ■■■ WebView2からのコールバック受け皿(動的束縛経由) ■■■
'***************************************************************************************************
'* 注意事項：★重要★ 以下は全て`CDPWebView2CallbackHandler.Handler_Invoke`から`owner As Object`
'            経由(動的束縛)で呼ばれるため、全てPublic Subでなければならない
'***************************************************************************************************
' --- ICoreWebView2EnvironmentOptions ---
Public Sub EnvOpt_OnGetAdditionalBrowserArguments(ByVal pOut As LongPtr)
    CDPWebView2Thunks.EnvOpt_WriteStringOut pOut, m_additionalBrowserArguments
End Sub

Public Sub EnvOpt_OnPutAdditionalBrowserArguments(ByVal pValue As LongPtr)
    If pValue <> 0 Then m_additionalBrowserArguments = CDPWebView2Thunks.PtrToString(pValue)
End Sub

Public Sub EnvOpt_OnGetLanguage(ByVal pOut As LongPtr)
    CDPWebView2Thunks.EnvOpt_WriteStringOut pOut, m_language
End Sub

Public Sub EnvOpt_OnPutLanguage(ByVal pValue As LongPtr)
    If pValue <> 0 Then m_language = CDPWebView2Thunks.PtrToString(pValue)
End Sub

Public Sub EnvOpt_OnGetTargetCompatibleBrowserVersion(ByVal pOut As LongPtr)
    CDPWebView2Thunks.EnvOpt_WriteStringOut pOut, m_targetCompatibleBrowserVersion
End Sub

Public Sub EnvOpt_OnPutTargetCompatibleBrowserVersion(ByVal pValue As LongPtr)
    If pValue <> 0 Then m_targetCompatibleBrowserVersion = CDPWebView2Thunks.PtrToString(pValue)
End Sub

Public Sub EnvOpt_OnGetAllowSingleSignOnUsingOSPrimaryAccount(ByVal pOut As LongPtr)
    CDPWebView2Thunks.EnvOpt_WriteBoolOut pOut, m_allowSingleSignOnUsingOSPrimaryAccount
End Sub

Public Sub EnvOpt_OnPutAllowSingleSignOnUsingOSPrimaryAccount(ByVal rawValue As LongPtr)
    m_allowSingleSignOnUsingOSPrimaryAccount = (CLngLng(rawValue) And &HFFFFFFFF^) <> 0
End Sub

Public Sub EnvOpt_OnGetAreBrowserExtensionsEnabled(ByVal pOut As LongPtr)
    CDPWebView2Thunks.EnvOpt_WriteBoolOut pOut, m_areBrowserExtensionsEnabled
End Sub

Public Sub EnvOpt_OnPutAreBrowserExtensionsEnabled(ByVal rawValue As LongPtr)
    m_areBrowserExtensionsEnabled = (CLngLng(rawValue) And &HFFFFFFFF^) <> 0
End Sub
