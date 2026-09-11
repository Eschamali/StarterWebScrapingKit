VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2CallbackHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   WebView2からのCOMコールバック(サンク経由)を受け取るVBA側の受け皿クラスです。
'
'   出典・移植元：WebView2-For-Excel-VBA プロジェクトの`ComCallbackHandler.cls`。
'   `Handler_Invoke`(vtableスロット7固定、`WebView2Thunks.GetClassMethodAddrAtFixedSlot`が
'   この位置を前提に呼び出す)以降の骨格は無変更。`Select Case m_kind`の分岐を、このプロジェクトの
'   スコープ(CDPメソッド呼び出し完了・CDPイベント受信)向けの2種類に絞り、既存の
'   HK_EnvironmentCompleted/HK_ControllerCompletedはそのまま残しています。
'
'   1インスタンス = 1個の登録済みハンドラの関係で生成され、WebView2が握るvtableの
'   AddRef/Release/Invokeを、`m_refCount`と`m_owner`への動的束縛呼び出しで実装します。
'
'   重要：`m_owner.OnXxx`は**Public Sub**でなければ動的束縛経由で呼べません
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
'--- COMコールバック管理全般系 ---
Private m_kind      As HandlerKind  'どんなCOMコールバックか？
Private m_owner     As Object       '転送先VBAClassオブジェクト
Private m_pSlot     As LongPtr      '共有スロットプールのポインター。`WebView2Thunks.AcquireHandlerFor`で作成したCOMコールバックメモリ領域の場所(ポインター)
Private m_refCount  As Long         'このClassオブジェクトの参照カウント値

'--- CDP系に関する管理変数 ---
' HK_CdpMethodCompleted用: CDPコマンドの`id`をそのまま格納し、完了コールバックで
' `OnCdpMethodCompleted`に渡す(1ショット系、同時並行呼び出し可)
Private m_callbackId As Long

' HK_CdpEventReceived用: どのCDPイベント名を購読したものかを記憶する(args側にはイベント名が
' 含まれないため、購読登録時にここへ保存しておく必要がある)
Private m_eventName As String



'***************************************************************************************************
'                              ■■■ 初期化 ■■■
'***************************************************************************************************
'* 機能　　：`WebView2Thunks.AcquireHandlerFor`/`Thunks_AcquireSlot`から呼ばれる初期化処理
'---------------------------------------------------------------------------------------------------
'* 注意事項：refcountは0から開始する。WebView2にハンドラを渡した後はWebView2自身がAddRefして
'            責任を持つため、VBA側では参照を保持しない
'***************************************************************************************************
Friend Sub Init( _
    ByVal kind As HandlerKind, _
    ByVal owner As Object, _
    ByVal pSlot As LongPtr)

    m_kind = kind
    Set m_owner = owner
    m_pSlot = pSlot
    m_refCount = 0
End Sub



'***************************************************************************************************
'                      ■■■ プロパティ群 - 基本情報編(基盤側からの操作用) ■■■
'***************************************************************************************************
Friend Property Get refCount() As Long
    refCount = m_refCount
End Property

Friend Property Let refCount(ByVal v As Long)
    m_refCount = v
End Property

Friend Property Get Slot() As LongPtr
    Slot = m_pSlot
End Property

Friend Property Get kind() As HandlerKind
    kind = m_kind
End Property

'***************************************************************************************************
'* 機能　　：refcount=0になったときに基盤側(HandlerReleaseInternal)から呼ばれ、循環参照を切ります
'***************************************************************************************************
Friend Sub ClearOwner()
    Set m_owner = Nothing
End Sub



'***************************************************************************************************
'               ■■■ Handler_Invoke(vTableスロット7固定 - メインディッシュ) ■■■
'***************************************************************************************************
'* 機能　　：WebView2からのコールバックを受ける唯一のエントリです
'---------------------------------------------------------------------------------------------------
'* 引数仕様：this : サンク後のrdx(= WebView2がrcxに乗せていたハンドラobjアドレス、未使用)
'            arg1 : サンク後のr8 (= WebView2がrdxに乗せていた1st引数)
'            arg2 : サンク後のr9 (= WebView2がr8に乗せていた2nd引数)
'---------------------------------------------------------------------------------------------------
'* 各kindの解釈：
'   HK_EnvironmentCompleted : Invoke(this, errorCode, pEnvironment)
'   HK_ControllerCompleted  : Invoke(this, errorCode, pController)
'   HK_CdpMethodCompleted   : Invoke(this, errorCode, resultAsJsonポインタ)
'                             ※CallDevToolsProtocolMethod/…ForSessionの両方で共用
'   HK_CdpEventReceived     : Invoke(this, sender, args)
'                             ※argsはICoreWebView2DevToolsProtocolEventReceivedEventArgs*
'   HK_EnvOpt_*             : Invoke(this, arg1, 未使用)
'                             ※ICoreWebView2EnvironmentOptions(base/2/3/5/6/7/8)系のget_/put_
'                             1個につき1つのkind。詳細は`WebView2Thunks`の「EnvOpt」セクション参照
'***************************************************************************************************
Public Function Handler_Invoke( _
    ByVal This As LongPtr, _
    ByVal arg1 As LongPtr, _
    ByVal arg2 As LongPtr) As Long

    '1. `Init`で、事前に設定したコールバックタイプに応じて、定義したClassオブジェクトメソッドで、呼び出す
    '   ※各Caseは同時に1つしか実行されないので、エラーコード復元用の変数は使い回す
    Dim ErrorCode As Long

    Select Case m_kind
        '---------------- 必須系 ----------------
        Case HK_EnvironmentCompleted
            '2. 第2引数を基に、エラーコードを復元
            ErrorCode = LongPtrLowDword(arg1)

            '3. `WebView2Host.cls`の完了コールバックプロシージャへ
            If Not (m_owner Is Nothing) Then m_owner.OnEnvironmentCompleted ErrorCode, arg2

        Case HK_ControllerCompleted
            '2. 第2引数を基に、エラーコードを復元
            ErrorCode = LongPtrLowDword(arg1)

            '3. `WebView2Host.cls`のメソッドへ
            If Not (m_owner Is Nothing) Then m_owner.OnControllerCompleted ErrorCode, arg2


        '---------------- CDP系 ----------------
        ' SDKシグネチャ: ICoreWebView2CallDevToolsProtocolMethodCompletedHandler::
        '   Invoke(HRESULT errorCode, LPCWSTR returnObjectAsJson)
        Case HK_CdpMethodCompleted
            '2. 第2引数を基に、エラーコードを復元
            ErrorCode = LongPtrLowDword(arg1)

            '3. `CDPCoreViaWebView2.SendCommandCDP`で設定した`m_callbackId`をコマンドID応答として、`CDPCoreViaWebView2.OnCdpMethodCompleted`へ渡す
            If Not (m_owner Is Nothing) Then m_owner.OnCdpMethodCompleted m_callbackId, ErrorCode, arg2

        ' SDKシグネチャ: ICoreWebView2DevToolsProtocolEventReceivedEventHandler::
        '   Invoke(ICoreWebView2* sender, ICoreWebView2DevToolsProtocolEventReceivedEventArgs* args)
        '   arg1 = sender、arg2 = args
        Case HK_CdpEventReceived
            '2. `CDPCoreViaWebView2.SubscribeCdpEvent`で確保したCOMコールバック領域にて設定したCDPイベント名を、引数に渡して、`CDPCoreViaWebView2.OnCdpEventReceived`へ渡す
            '※これは実行結果ではなく、WebView2自身で起こったイベント情報なので、実行時エラーコードの概念はない
            If Not (m_owner Is Nothing) Then m_owner.OnCdpEventReceived m_eventName, arg2


        '---------------- 拡張機能系 ----------------
        ' SDKシグネチャ: ICoreWebView2ProfileAddBrowserExtensionCompletedHandler::
        '   Invoke(HRESULT errorCode, ICoreWebView2BrowserExtension* result)
        Case HK_AddBrowserExtensionCompleted
            '2. 第2引数を基に、エラーコードを復元
            ErrorCode = LongPtrLowDword(arg1)

            '3. `WebView2Host.cls`の完了コールバックプロシージャへ
            If Not (m_owner Is Nothing) Then m_owner.OnAddBrowserExtensionCompleted ErrorCode, arg2

        ' SDKシグネチャ: ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler::
        '   Invoke(HRESULT errorCode, ICoreWebView2BrowserExtensionList* result)
        Case HK_GetBrowserExtensionsCompleted
            '2. 第2引数を基に、エラーコードを復元
            ErrorCode = LongPtrLowDword(arg1)

            '3. `WebView2Host.cls`の完了コールバックプロシージャへ
            If Not (m_owner Is Nothing) Then m_owner.OnGetBrowserExtensionsCompleted ErrorCode, arg2

        ' SDKシグネチャ: ICoreWebView2BrowserExtensionRemoveCompletedHandler::
        '   Invoke(HRESULT errorCode)  ※引数1個のみ(結果オブジェクトは無い)
        Case HK_RemoveBrowserExtensionCompleted
            '2. 第2引数を基に、エラーコードを復元
            ErrorCode = LongPtrLowDword(arg1)

            '3. `WebView2Host.cls`の完了コールバックプロシージャへ
            If Not (m_owner Is Nothing) Then m_owner.OnRemoveBrowserExtensionCompleted ErrorCode


        '---------------- ICoreWebView2EnvironmentOptions系 ----------------
        ' SDKシグネチャ: get_Xxx(this, [out]value) / put_Xxx(this, value) のいずれも1引数のみ。
        ' arg1が「get_ なら[out]ポインタ」「put_ なら渡された値そのもの」に対応する
        '・`WebView2EnvOptions.BuildNativePointer`で構築した「m_kind」を基に、`WebView2Loader`側からVBEでセットしたプロパティ値を取りに行きます
        Case HK_EnvOpt_GetAdditionalBrowserArguments
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetAdditionalBrowserArguments arg1
        Case HK_EnvOpt_PutAdditionalBrowserArguments
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutAdditionalBrowserArguments arg1

        Case HK_EnvOpt_GetLanguage
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetLanguage arg1
        Case HK_EnvOpt_PutLanguage
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutLanguage arg1

        Case HK_EnvOpt_GetTargetCompatibleBrowserVersion
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetTargetCompatibleBrowserVersion arg1
        Case HK_EnvOpt_PutTargetCompatibleBrowserVersion
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutTargetCompatibleBrowserVersion arg1

        Case HK_EnvOpt_GetAllowSingleSignOnUsingOSPrimaryAccount
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetAllowSingleSignOnUsingOSPrimaryAccount arg1
        Case HK_EnvOpt_PutAllowSingleSignOnUsingOSPrimaryAccount
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutAllowSingleSignOnUsingOSPrimaryAccount arg1

        Case HK_EnvOpt_GetAreBrowserExtensionsEnabled
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetAreBrowserExtensionsEnabled arg1
        Case HK_EnvOpt_PutAreBrowserExtensionsEnabled
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutAreBrowserExtensionsEnabled arg1

        Case HK_EnvOpt_GetExclusiveUserDataFolderAccess
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetExclusiveUserDataFolderAccess arg1
        Case HK_EnvOpt_PutExclusiveUserDataFolderAccess
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutExclusiveUserDataFolderAccess arg1

        Case HK_EnvOpt_GetIsCustomCrashReportingEnabled
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetIsCustomCrashReportingEnabled arg1
        Case HK_EnvOpt_PutIsCustomCrashReportingEnabled
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutIsCustomCrashReportingEnabled arg1

        Case HK_EnvOpt_GetEnableTrackingPrevention
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetEnableTrackingPrevention arg1
        Case HK_EnvOpt_PutEnableTrackingPrevention
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutEnableTrackingPrevention arg1

        Case HK_EnvOpt_GetChannelSearchKind
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetChannelSearchKind arg1
        Case HK_EnvOpt_PutChannelSearchKind
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutChannelSearchKind arg1

        Case HK_EnvOpt_GetReleaseChannels
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetReleaseChannels arg1
        Case HK_EnvOpt_PutReleaseChannels
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutReleaseChannels arg1

        Case HK_EnvOpt_GetScrollBarStyle
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnGetScrollBarStyle arg1
        Case HK_EnvOpt_PutScrollBarStyle
            If Not (m_owner Is Nothing) Then m_owner.EnvOpt_OnPutScrollBarStyle arg1


        Case Else
            Debug.Print "WebView2CallbackHandler.Handler_Invoke: 未知のkind=" & m_kind
    End Select

    Handler_Invoke = 0   ' S_OK
End Function



'***************************************************************************************************
'                      ■■■ プロパティ群 - 拡張用(基盤側からの操作用) ■■■
'***************************************************************************************************
'* 注意事項：★重要★ `WebView2Thunks.GetClassMethodAddrAtFixedSlot`は、このClassの**既定の
'            公開インターフェース**(=`Public`メンバーのみ。`Friend`は早期バインディング専用の
'            プロジェクト内アクセス制御でしかなく、外部/レイトバインディングから見える
'            vtableには乗らない)の宣言順に基づいて、固定スロット番号(7=IUnknown3+IDispatch4の
'            直後、最初の`Public`メンバー)でHandler_Invokeのアドレスを取得している。
'            このClass内で`Public`なのは`Handler_Invoke`だけなので、`Friend`メンバー
'            (`Init`/`RefCount`/`Slot`/`kind`/`ClearOwner`/`callbackId`/`EventName`)は
'            どこに置いても(前でも後でも)スロット番号に影響しない。ただし、もし新たに
'            **`Public`**なメンバーを`Handler_Invoke`より前に追加すると、スロット番号が
'            ズレてWebView2からのコールバックが全く別のメソッドを叩いてしまい、即クラッシュの
'            原因になる。実機で7以外の値(5,6)を試して確認済み(いずれも起動失敗)。
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] CDP系
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：HK_CdpMethodCompleted用のコマンドIDプロパティです
'***************************************************************************************************
Friend Property Let callbackId(ByVal v As Long)
    m_callbackId = v
End Property

'***************************************************************************************************
'* 機能　　：HK_CdpEventReceived用のイベント名プロパティです
'***************************************************************************************************
Friend Property Let EventName(ByVal v As String)
    m_eventName = v
End Property



'***************************************************************************************************
'                               ■■■ ヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：LongPtrの下位32bitを符号付きLongとして取り出します(HRESULT等の抽出用)
'***************************************************************************************************
Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
    Dim u As LongLong
    u = v And &HFFFFFFFF^
    If u >= &H80000000^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
End Function
