VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPCoreViaWebView2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   WebView2(ICoreWebView2)経由での、Chrome DevTools Protocolのやり取りに特化した専用Classです。
'   `CDPCoreViaWebSocket.cls`と同じ「`CDPCore.cls`にRaiseEventで繋ぎ込む」拡張パターンに則った、
'   3つ目のtransportです。
'
'   移植/実装対象は、WebView2 SDKの以下3メソッドのみです(このプロジェクトはCDPフォーマットしか
'   扱わないため、これ以外のWebView2機能[ナビゲーション制御・DOM操作等]は一切実装しません):
'     ・CallDevToolsProtocolMethodAsync           → ICoreWebView2::CallDevToolsProtocolMethod
'     ・CallDevToolsProtocolMethodForSessionAsync → ICoreWebView2_11::CallDevToolsProtocolMethodForSession
'     ・GetDevToolsProtocolEventReceiver           → ICoreWebView2::GetDevToolsProtocolEventReceiver
'
'   ★CDPCoreViaWebSocket.clsとの本質的な違い★
'     WebSocket版は「バイト列のフラグメントを`RaiseEvent WebSocketReceive`で都度流す」設計ですが、
'     WebView2はCOMコールバック経由で「既にUTF-16デコード済み・欠けのない完成JSON文字列」を
'     1件ずつ届けてくるため、このClassは`RaiseEvent CDPMessageReceived(RawJson As String)`という
'     文字列1件そのままの、より単純な形のイベントを発火します。
'
'   ★既知の制約(CDP-over-WebSocketとの違い)★
'     CDP-over-WebSocketは「ドメインをenableすれば以後そのドメインの全イベントが自動で流れる」
'     モデルですが、WebView2の`GetDevToolsProtocolEventReceiver`は「イベント名ごとに個別登録」
'     が必要なモデルです。この違いを隠さず、`SubscribeCdpEvent`/`UnsubscribeCdpEvent`という
'     明示的なAPIとして公開しています。
'     `GetDevToolsProtocolEventReceiver`はトップレベルの`ICoreWebView2`に対して1回購読するだけで、
'     `Target.setAutoAttach`等でアタッチされた子セッション(タブ/iframe/worker等)からの同名イベントも
'     同じコールバックに届きます。どのセッション由来かは`ICoreWebView2DevToolsProtocolEventReceived
'     EventArgs2::get_SessionId`で判別し(`OnCdpEventReceived`/`TryGetEventSessionId`)、有効な
'     `sessionId`があれば再構成後のJsonへ含めるため、`CDPCore.BrowserReceivedDataCheck`は
'     Pipe/WebSocket版と同様に`CDPContextEvent`(タブ向け)/`CDPBrowserEvent`(ブラウザ向け)を
'     正しく振り分けます。
'
'   ★既知の落とし穴(継承不可避、WebView2-For-Excel-VBAプロジェクトの知見より)★
'     ・コールバック待ち中(コマンド送信?完了、イベント購読中)にVBEでブレーク/ステップ実行すると
'       Excelがクラッシュする可能性がある(サンク技法そのものの制約)
'     ・`OnCdpMethodCompleted`/`OnCdpEventReceived`の中から`ExecuteCDP`のような同期待ちの
'       メソッドを呼ばないこと(再入デッドロックの原因になる)
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
Private Const ThisClassName As String = "CDPCoreViaWebView2"                            'ログのトレース用
Private Const IID_ICoreWebView2_11 As String = "0be78e56-c193-4051-b943-23b460c08bdb"   ' ICoreWebView2_11(CallDevToolsProtocolMethodForSessionを持つ派生IF)のIID



'***************************************************************************************************
'                                   ■■■ イベント定義 ■■■
'***************************************************************************************************
'* 完成した1件のCDP-Json文字列(コマンド結果 or 非同期イベント)を、そのまま発信します
Public Event CDPMessageReceived(RawJson As String)



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
Private m_Host          As CDPWebView2Host
Private m_pView11       As LongPtr    ' QueryInterfaceで取得したICoreWebView2_11(未対応環境では0のまま)
Private m_pView11Failed As Boolean  ' 1度QIに失敗したら、以後は再試行しない

' CDPイベント購読の管理(イベント名 → レシーバーポインタ/トークン)
Private m_subReceiver As Dictionary
Private m_subToken    As Dictionary



'***************************************************************************************************
'                                   ■■■ 初期化/後始末 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Set m_subReceiver = New Dictionary
    Set m_subToken = New Dictionary
End Sub

Private Sub Class_Terminate()
    DisconnectCDP
End Sub



'***************************************************************************************************
'                               ■■■ 新規接続関連 ■■■
'***************************************************************************************************
'* 機能　　：`WebView2`方式としてCDP制御できる準備を行います(Environment/Controller/Viewの生成)
'---------------------------------------------------------------------------------------------------
'* 返り値  ：接続成功可否
'* 引数    ：userDataPath   WebView2のユーザーデータフォルダ(省略時は%TEMP%\CDPWebView2)
'---------------------------------------------------------------------------------------------------
'* 注意事項：既存の接続がある場合は、切断されます
'***************************************************************************************************
Public Function ConnectCDP(Optional userDataPath As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ConnectCDP"

    If Not (m_Host Is Nothing) Then DisconnectCDP

    Set m_Host = New CDPWebView2Host
    If Not m_Host.Init(userDataPath) Then
        Debug.Print FromProcedureName & ": WebView2の初期化に失敗しました"
        Set m_Host = Nothing
        Exit Function
    End If

    ConnectCDP = True
End Function

'***************************************************************************************************
'* 機能　　：WebView2(ICoreWebView2)が生きてるか確認します
'***************************************************************************************************
Public Function isAvailability() As Boolean
    isAvailability = (Not m_Host Is Nothing) And m_Host.IsReady
End Function

'***************************************************************************************************
'* 機能　　：接続を切断し、各種ハンドル/購読を解放します
'***************************************************************************************************
Public Sub DisconnectCDP()
    '1. 残っている購読を全解除
    If Not (m_subReceiver Is Nothing) Then
        Dim N As Long: N = m_subReceiver.Count
        If N > 0 Then
            Dim keys() As String
            ReDim keys(0 To N - 1)
            Dim i As Long
            For i = 0 To N - 1
                keys(i) = m_subReceiver.KeyAtIndex(i)
            Next i
            For i = 0 To N - 1
                UnsubscribeCdpEvent keys(i)
            Next i
        End If
    End If

    '2. ICoreWebView2_11(QI取得分)を解放
    If m_pView11 <> 0 Then
        CDPWebView2Thunks.ComRelease m_pView11
        m_pView11 = 0
    End If
    m_pView11Failed = False

    '3. ホスト(Environment/Controller/View)を解放
    If Not (m_Host Is Nothing) Then
        m_Host.CLOSE_
        Set m_Host = Nothing
    End If
End Sub



'***************************************************************************************************
'                               ■■■ CDPコマンド送信 ■■■
'***************************************************************************************************
'* 機能　　：WebView2経由で、ブラウザへCDPコマンドを送信します
'---------------------------------------------------------------------------------------------------
'* 引数　　：strData    完成したCDP-Json文字列コマンド(`id`/`method`/`params`/[`sessionId`]を含む)
'* 返り値  ：実行結果コード(0で成功。`CDPCore.LastErrorWriteFile`相当として扱われます)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`sessionId`の有無で、`CallDevToolsProtocolMethod`(スロット36)と
'            `CallDevToolsProtocolMethodForSession`(スロット99、ICoreWebView2_11経由)を
'            切り分けます。完了は`OnCdpMethodCompleted`で非同期に受け取り、`id`を付け直した
'            CDP形式のJson文字列として`RaiseEvent CDPMessageReceived`します。
'* 注意事項：`params`は`BiDiCDPJson.NodeKey("params").Stringify()`で再構築して取り出します
'            (`RawStringKey`はスカラー値専用でオブジェクト/配列型では空文字を返してしまうため使えません)
'***************************************************************************************************
Public Function SendCommandCDP(strData As String) As Long
    Const FromProcedureName As String = ThisClassName & ".SendCommandCDP"
    Const E_HANDLE As Long = &H80070006   ' ハンドルが無効(=接続が既に切れている)

    '★重要★ ここで`0`(成功扱い)を返すと、呼び出し元(`CDPCore.writeWebView2CDP`経由で
    '`LastPipeError.WriteFile`)が「送信成功」と誤認し、実際には送られていないコマンドの
    '結果を`CDPBrowser.AutoWaitTakeResultCDP`が永遠に待ち続けてしまう
    '(Pipe/WebSocket版は、死んだハンドルへの`WriteFile`/`send`が実際にOSエラーを返すことで
    '`CDPBrowser.ExecuteCDPAsync`の`LastErrorWriteFile`チェックが機能するが、WebView2版は
    '自前でこの「接続が切れている」状態を明示的にエラーとして返す必要がある)
    If m_Host Is Nothing Then
        Debug.Print FromProcedureName & ": 接続済みのHostがありません(切断済み)"
        SendCommandCDP = E_HANDLE
        Exit Function
    End If
    Dim pView As LongPtr: pView = m_Host.ViewPtr
    If pView = 0 Then
        Debug.Print FromProcedureName & ": ViewPtrが0です(切断済みか初期化未完了)"
        SendCommandCDP = E_HANDLE
        Exit Function
    End If

    Dim Node As BiDiCDPJson: Set Node = BiDiCDPJson.Parse(strData)
    Dim commandID As Long: commandID = CLng(Node.NumberKey("id"))
    Dim methodName As String: methodName = Node.StringKey("method")

    'params」オブジェクトの取り出し
    '`RawStringKey`はスカラー値(文字列/数値/真偽値)専用で、オブジェクト/配列型の値では
    '`ValStart`/`ValLen`が設定されないため常に空文字(=NULL相当)を返してしまう(BiDiCDPJson側の仕様)。
    '`params`は常にJSONオブジェクトなので、`NodeKey`で子ノードを取得してから`Stringify`で
    '再構築する(子孫を辿って正しくシリアライズするため、この方法なら安全)
    Dim paramsRaw As String
    If Node.ExistsKey("params") Then
        Dim paramsNode As BiDiCDPJson
        Set paramsNode = Node.NodeKey("params")
        If paramsNode Is Nothing Then paramsRaw = "{}" Else paramsRaw = paramsNode.Stringify()
    Else
        paramsRaw = "{}"
    End If

    Dim sessionID As String
    If Node.ExistsKey("sessionId") Then sessionID = Node.StringKey("sessionId")

    Debug.Print FromProcedureName & ": id=" & commandID & " method=[" & methodName & _
                "] paramsRaw=[" & paramsRaw & "] (len=" & Len(paramsRaw) & ") sessionId=[" & sessionID & "]"

    Dim h As CDPWebView2CallbackHandler
    Set h = CDPWebView2Thunks.AcquireHandlerFor(HK_CdpMethodCompleted, Me)
    If h Is Nothing Then
        Debug.Print FromProcedureName & ": ハンドラのAcquireに失敗しました"
        SendCommandCDP = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If
    h.callbackId = commandID

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + CDPWebView2Thunks.VTABLE_OBJ_OFFSET

    Dim hr As Long
    If LenB(sessionID) = 0 Then
        'ICoreWebView2::CallDevToolsProtocolMethod(vtableスロット36)
        hr = CDPWebView2Thunks.dcf(pView, 36, "View.CallDevToolsProtocolMethod", _
                                    StrPtr(methodName), StrPtr(paramsRaw), pVTableObj)
    Else
        'ICoreWebView2_11::CallDevToolsProtocolMethodForSession(vtableスロット99)
        Dim pView11 As LongPtr: pView11 = EnsureView11(pView)
        If pView11 = 0 Then
            Debug.Print FromProcedureName & ": ICoreWebView2_11が取得できないため、" & _
                        "sessionId指定のコマンドを送信できません"
            Set h = Nothing
            SendCommandCDP = &H80004002   ' E_NOINTERFACE
            Exit Function
        End If
        hr = CDPWebView2Thunks.dcf(pView11, 99, "View11.CallDevToolsProtocolMethodForSession", _
                                    StrPtr(sessionID), StrPtr(methodName), StrPtr(paramsRaw), pVTableObj)
    End If

    'WebView2側がAddRef済みなので、こちらのローカル参照は捨てる
    Set h = Nothing

    SendCommandCDP = hr
End Function

'***************************************************************************************************
'* 機能　　：`ICoreWebView2`から`ICoreWebView2_11`をQueryInterfaceで取得します(1回だけ試行)
'***************************************************************************************************
Private Function EnsureView11(ByVal pView As LongPtr) As LongPtr
    If m_pView11 <> 0 Then EnsureView11 = m_pView11: Exit Function
    If m_pView11Failed Then Exit Function

    Dim iid As CDPWebView2Thunks.GUID
    CDPWebView2Thunks.FillGUID iid, IID_ICoreWebView2_11

    Dim pOut As LongPtr
    Dim hr As Long
    hr = CDPWebView2Thunks.dcf(pView, 0, "View.QueryInterface(ICoreWebView2_11)", VarPtr(iid), VarPtr(pOut))
    If hr = 0 And pOut <> 0 Then
        m_pView11 = pOut
        EnsureView11 = m_pView11
    Else
        m_pView11Failed = True
    End If
End Function

'***************************************************************************************************
'* 機能　　：`CallDevToolsProtocolMethod(ForSession)`の完了通知を受け取り、CDP形式のJsonに
'            組み立てて発信します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`CDPWebView2CallbackHandler.Handler_Invoke`から動的束縛で呼ばれるため、Public必須
'***************************************************************************************************
Public Sub OnCdpMethodCompleted(ByVal callbackId As Long, ByVal ErrorCode As Long, ByVal resultJson As String)
    Dim synthesized As String
    If ErrorCode = 0 Then
        If LenB(resultJson) = 0 Then resultJson = "null"
        synthesized = "{""id"":" & callbackId & ",""result"":" & resultJson & "}"
    Else
        synthesized = "{""id"":" & callbackId & ",""error"":{""code"":" & ErrorCode & _
                      ",""message"":""CallDevToolsProtocolMethod(ForSession) failed""}}"
    End If
    RaiseEvent CDPMessageReceived(synthesized)
End Sub



'***************************************************************************************************
'                               ■■■ CDPイベント購読 ■■■
'***************************************************************************************************
'* 機能　　：指定したCDPイベント名を、WebView2経由で購読します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：購読成功可否(既に購読済みの場合も`True`)
'* 引数　　：eventName   購読するCDPイベント名(例："Page.loadEventFired")
'---------------------------------------------------------------------------------------------------
'* 注意事項：CDP-over-WebSocketと異なり、ドメインをenableしただけでは自動購読されません。
'            必要なイベント名ごとに、このメソッドを個別に呼び出してください。
'***************************************************************************************************
Public Function SubscribeCdpEvent(ByVal EventName As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".SubscribeCdpEvent"

    If m_Host Is Nothing Then Exit Function
    Dim pView As LongPtr: pView = m_Host.ViewPtr
    If pView = 0 Then Exit Function

    If m_subReceiver.Exists(EventName) Then SubscribeCdpEvent = True: Exit Function

    Dim pReceiver As LongPtr
    Dim hr As Long
    hr = CDPWebView2Thunks.dcf(pView, 42, "View.GetDevToolsProtocolEventReceiver", _
                               StrPtr(EventName), VarPtr(pReceiver))
    If hr <> 0 Or pReceiver = 0 Then
        Debug.Print FromProcedureName & ": GetDevToolsProtocolEventReceiver失敗 hr=&H" & Hex(hr)
        Exit Function
    End If

    Dim h As CDPWebView2CallbackHandler
    Set h = CDPWebView2Thunks.AcquireHandlerFor(HK_CdpEventReceived, Me)
    If h Is Nothing Then
        CDPWebView2Thunks.ComRelease pReceiver
        Exit Function
    End If
    h.EventName = EventName

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + CDPWebView2Thunks.VTABLE_OBJ_OFFSET

    Dim token As LongLong
    hr = CDPWebView2Thunks.dcf(pReceiver, 3, "Receiver.add_DevToolsProtocolEventReceived", _
                               pVTableObj, VarPtr(token))
    If hr <> 0 Then
        Debug.Print FromProcedureName & ": add_DevToolsProtocolEventReceived失敗 hr=&H" & Hex(hr)
        Set h = Nothing
        CDPWebView2Thunks.ComRelease pReceiver
        Exit Function
    End If

    'WebView2側がハンドラをAddRef済みなので、ローカル参照hはこれ以上不要(捨ててよい)
    Set h = Nothing

    m_subReceiver.Add EventName, (pReceiver)
    m_subToken.Add EventName, (token)
    SubscribeCdpEvent = True
End Function

'***************************************************************************************************
'* 機能　　：指定したCDPイベント名の購読を解除します
'***************************************************************************************************
Public Function UnsubscribeCdpEvent(ByVal EventName As String) As Boolean
    If m_subReceiver Is Nothing Then Exit Function
    If Not m_subReceiver.Exists(EventName) Then Exit Function

    Dim pReceiver As LongPtr: pReceiver = m_subReceiver(EventName)
    Dim token As LongLong: token = m_subToken(EventName)

    CDPWebView2Thunks.dcf pReceiver, 4, "Receiver.remove_DevToolsProtocolEventReceived", token
    CDPWebView2Thunks.ComRelease pReceiver

    m_subReceiver.Remove EventName
    m_subToken.Remove EventName
    UnsubscribeCdpEvent = True
End Function

'***************************************************************************************************
'* 機能　　：`GetDevToolsProtocolEventReceived`の発火通知を受け取り、CDP形式のJsonに組み立てて
'            発信します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`GetDevToolsProtocolEventReceiver`はトップレベルの`ICoreWebView2`に対して1回購読する
'            だけで、`Target.setAutoAttach`等でアタッチされた子セッション(iframe/worker等)からの
'            同名イベントも同じコールバックに届く(公式サンプルコードで確認済みの仕様)。
'            どのセッション由来かは`args`を`ICoreWebView2DevToolsProtocolEventReceivedEventArgs2`
'            へQueryInterfaceし、`get_SessionId`で判別する(既定のトップレベルターゲット由来なら
'            空文字が返る)。これにより`sessionId`を含めて再構成し、`CDPCore.
'            BrowserReceivedDataCheck`が`CDPContextEvent`(タブ向け)/`CDPBrowserEvent`
'            (ブラウザ向け)を正しく振り分けられるようにする。
'***************************************************************************************************
Public Sub OnCdpEventReceived(ByVal EventName As String, ByVal pArgs As LongPtr)
    Dim paramJson As String
    paramJson = CDPWebView2Thunks.GetStringProperty(pArgs, 3, "Args.get_ParameterObjectAsJson")
    If LenB(paramJson) = 0 Then paramJson = "{}"

    Dim sessionID As String
    sessionID = TryGetEventSessionId(pArgs)


    Dim synthesized As String
    If LenB(sessionID) = 0 Then
        synthesized = "{""method"":""" & EventName & """,""params"":" & paramJson & "}"
    Else
        synthesized = "{""method"":""" & EventName & """,""params"":" & paramJson & _
                      ",""sessionId"":""" & sessionID & """}"
    End If
    RaiseEvent CDPMessageReceived(synthesized)
End Sub

'***************************************************************************************************
'* 機能　　：イベントargsを`…EventArgs2`へQueryInterfaceし、`get_SessionId`(vtableスロット4)で
'            セッションIDを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：セッションID(トップレベルの既定ターゲット由来、または未対応環境の場合は空文字)
'***************************************************************************************************
Private Function TryGetEventSessionId(ByVal pArgs As LongPtr) As String
    Const IID_Args2 As String = "2dc4959d-1494-4393-95ba-bea4cb9ebd1b"   ' ICoreWebView2DevToolsProtocolEventReceivedEventArgs2

    Dim iid As CDPWebView2Thunks.GUID
    CDPWebView2Thunks.FillGUID iid, IID_Args2

    Dim pArgs2 As LongPtr
    Dim hr As Long
    hr = CDPWebView2Thunks.dcf(pArgs, 0, "Args.QueryInterface(Args2)", VarPtr(iid), VarPtr(pArgs2))
    If hr <> 0 Or pArgs2 = 0 Then Exit Function

    TryGetEventSessionId = CDPWebView2Thunks.GetStringProperty(pArgs2, 4, "Args2.get_SessionId")

    'QueryInterfaceで確保した分の参照を解放(GetStringPropertyはQI自体はしない)
    CDPWebView2Thunks.ComRelease pArgs2
End Function
