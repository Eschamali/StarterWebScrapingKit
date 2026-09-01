VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPWebView2CallbackHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   WebView2からのCOMコールバック(サンク経由)を受け取るVBA側の受け皿クラスです。
'
'   出典・移植元：WebView2-For-Excel-VBA プロジェクトの`ComCallbackHandler.cls`。
'   `Handler_Invoke`(vtableスロット7固定、`CDPWebView2Thunks.GetClassMethodAddrAtFixedSlot`が
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

Private m_kind      As HandlerKind
Private m_owner     As Object
Private m_pSlot     As LongPtr
Private m_refCount  As Long

' HK_CdpMethodCompleted用: CDPコマンドの`id`をそのまま格納し、完了コールバックで
' `OnCdpMethodCompleted`に渡す(1ショット系、同時並行呼び出し可)
Private m_callbackId As Long

' HK_CdpEventReceived用: どのCDPイベント名を購読したものかを記憶する(args側にはイベント名が
' 含まれないため、購読登録時にここへ保存しておく必要がある)
Private m_eventName As String



'***************************************************************************************************
'                              ■■■ 初期化 ■■■
'***************************************************************************************************
'* 機能　　：`CDPWebView2Thunks.AcquireHandlerFor`/`Thunks_AcquireSlot`から呼ばれる初期化処理
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
'                              ■■■ プロパティ群(基盤側からの操作用) ■■■
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

Friend Property Get callbackId() As Long
    callbackId = m_callbackId
End Property

Friend Property Let callbackId(ByVal v As Long)
    m_callbackId = v
End Property

'***************************************************************************************************
'* 機能　　：refcount=0になったときに基盤側(HandlerReleaseInternal)から呼ばれ、循環参照を切ります
'***************************************************************************************************
Friend Sub ClearOwner()
    Set m_owner = Nothing
End Sub



'***************************************************************************************************
'                              ■■■ Handler_Invoke(vTableスロット7固定) ■■■
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
'                             ※ICoreWebView2EnvironmentOptions系のget_/put_1個につき1つのkind。
'                             詳細は`CDPWebView2Thunks`の「EnvOpt」セクション参照
'***************************************************************************************************
Public Function Handler_Invoke( _
    ByVal This As LongPtr, _
    ByVal arg1 As LongPtr, _
    ByVal arg2 As LongPtr) As Long

    Select Case m_kind
        Case HK_EnvironmentCompleted
            Dim ErrorCode As Long
            ErrorCode = LongPtrLowDword(arg1)

            If Not (m_owner Is Nothing) Then
                m_owner.OnEnvironmentCompleted ErrorCode, arg2
            End If

        Case HK_ControllerCompleted
            Dim errorCode2 As Long
            errorCode2 = LongPtrLowDword(arg1)

            If Not (m_owner Is Nothing) Then
                m_owner.OnControllerCompleted errorCode2, arg2
            End If

        Case HK_CdpMethodCompleted
            ' SDKシグネチャ: ICoreWebView2CallDevToolsProtocolMethodCompletedHandler::
            '   Invoke(HRESULT errorCode, LPCWSTR returnObjectAsJson)
            Dim errorCode3 As Long
            errorCode3 = LongPtrLowDword(arg1)
            Dim resultStr As String
            If arg2 <> 0 Then resultStr = PtrToString(arg2)

            If Not (m_owner Is Nothing) Then
                m_owner.OnCdpMethodCompleted m_callbackId, errorCode3, resultStr
            End If

        Case HK_CdpEventReceived
            ' SDKシグネチャ: ICoreWebView2DevToolsProtocolEventReceivedEventHandler::
            '   Invoke(ICoreWebView2* sender, ICoreWebView2DevToolsProtocolEventReceivedEventArgs* args)
            '   arg1 = sender、arg2 = args
            If Not (m_owner Is Nothing) Then
                m_owner.OnCdpEventReceived m_eventName, arg2
            End If

        ' --- ICoreWebView2EnvironmentOptions/Options6の各get_/put_ ---
        ' SDKシグネチャ: get_Xxx(this, [out]value) / put_Xxx(this, value) のいずれも1引数のみ。
        ' arg1が「get_ なら[out]ポインタ」「put_ なら渡された値そのもの」に対応する
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

        Case Else
            Debug.Print "CDPWebView2CallbackHandler.Handler_Invoke: 未知のkind=" & m_kind
    End Select

    Handler_Invoke = 0   ' S_OK
End Function

'***************************************************************************************************
'* 機能　　：HK_CdpEventReceived用のイベント名プロパティです
'---------------------------------------------------------------------------------------------------
'* 注意事項：★重要★ この2メンバーは、`Handler_Invoke`(vtableスロット7固定)より後方に置くこと。
'            `CDPWebView2Thunks.GetClassMethodAddrAtFixedSlot`は、このClassの
'            Public/Friendメンバーの宣言順([Init]→[RefCount Get/Let]→[Slot Get]→[kind Get]→
'            [callbackId Get/Let]→[ClearOwner]→[Handler_Invoke]、移植元と同じ順序)に基づいて
'            固定スロット番号でHandler_Invokeのアドレスを取得している。この並び順より前に
'            新規メンバーを追加すると、スロット番号がズレてWebView2からのコールバックが
'            全く別のメソッドを叩いてしまい、即クラッシュの原因になる。
'***************************************************************************************************
Friend Property Get EventName() As String
    EventName = m_eventName
End Property

Friend Property Let EventName(ByVal v As String)
    m_eventName = v
End Property

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
