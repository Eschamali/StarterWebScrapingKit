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
'     明示的なAPIとして公開しています(一括購読はWebView2側に概念自体が無いため未対応。
'     一括解除のみ`UnsubscribeAllCdpEvents`として提供)。
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
Private Const ThisClassName         As String = "CDPCoreViaWebView2"    'ログのトレース用

'* `HiddenPdfToolbarItems`用。`COREWEBVIEW2_PDF_TOOLBAR_ITEMS`のビットマスク値をEnum化したもの。
'  `Or`で組み合わせて指定する(例: `PDF_Save Or PDF_PrintItem`)。VBEはEnum型引数に対して、
'  `Or `と打った直後に同じEnumの候補一覧を再表示してくれるため、複数選択がしやすい
Public Enum WV2PdfToolbarItems
    PDF_Save = &H1
    PDF_PrintItem = &H2
    PDF_SaveAs = &H4
    PDF_ZoomIn = &H8
    PDF_ZoomOut = &H10
    PDF_Rotate = &H20
    PDF_FitPage = &H40
    PDF_PageLayout = &H80
    PDF_Bookmarks = &H100
    PDF_PageSelector = &H200
    PDF_Search = &H400
    PDF_FullScreen = &H800
    PDF_MoreSettings = &H1000
End Enum

'* `BoundsMode`用(`COREWEBVIEW2_BOUNDS_MODE`)
Public Enum WV2BoundsMode
    BoundsMode_UseRawPixels = 0             ' 既定
    BoundsMode_UseRasterizationScale = 1
End Enum

'* `MoveFocus`用(`COREWEBVIEW2_MOVE_FOCUS_REASON`)
Public Enum WV2MoveFocusReason
    MoveFocus_Programmatic = 0
    MoveFocus_Next = 1
    MoveFocus_Previous = 2
End Enum

'* `PreferredColorScheme`用(`COREWEBVIEW2_PREFERRED_COLOR_SCHEME`)
Public Enum WV2PreferredColorScheme
    ColorScheme_Auto = 0             ' 既定
    ColorScheme_Light = 1
    ColorScheme_Dark = 2
End Enum

'* `PreferredTrackingPreventionLevel`用(`COREWEBVIEW2_TRACKING_PREVENTION_LEVEL`)
Public Enum WV2TrackingPreventionLevel
    TrackingPrevention_None = 0
    TrackingPrevention_Basic = 1
    TrackingPrevention_Balanced = 2  ' 既定
    TrackingPrevention_Strict = 3
End Enum

'* `SetVirtualHostNameToFolderMapping`用(`COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND`)
Public Enum WV2HostResourceAccessKind
    HostAccess_Deny = 0
    HostAccess_Allow = 1
    HostAccess_DenyCors = 2
End Enum



'***************************************************************************************************
'                                   ■■■ イベント定義 ■■■
'***************************************************************************************************
'* 完成した1件のCDP-Json文字列(コマンド結果 or 非同期イベント)を、そのまま発信します
Public Event CDPMessageReceived(RawJson As String)



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
Private m_Host          As WebView2Host
Private m_pView11       As LongPtr      ' QueryInterfaceで取得したICoreWebView2_11(未対応環境では0のまま)
Private m_pView11Failed As Boolean      ' 1度QIに失敗したら、以後は再試行しない

' Environment作成オプション(`ICoreWebView2EnvironmentOptions`系)。`m_Host`の生死とは独立に
' こちらで直接保持するため、`ConnectCDP`実行前(=`m_Host`がまだ存在しない間)から自由に設定できる。
' Environment生成時にしか読まれない値のため、必ず`ConnectCDP`を呼ぶ**前**に設定すること
Private m_envOptions As WebView2EnvOptions

' CDPイベント購読の管理(イベント名 → レシーバーポインタ/トークン)
Private m_subReceiver As Dictionary
Private m_subToken    As Dictionary

'--- ログ設定状況 ---
Private currentLog As StateLog



'***************************************************************************************************
'                                ■■■ check用プロパティ ■■■
'***************************************************************************************************
'* 機能　　：WebView2(ICoreWebView2)が生きてるか確認します
'***************************************************************************************************
Property Get isAvailability() As Boolean
    If m_Host Is Nothing Then isAvailability = False Else isAvailability = m_Host.IsReady
End Property

'***************************************************************************************************
'* 機能　　：購読中のイベント数を取得します
'***************************************************************************************************
Property Get SubscribeCdpEventCount() As Long
    SubscribeCdpEventCount = m_subToken.Count
End Property

'***************************************************************************************************
'* 機能　　：購読中のイベント内容を一次元配列で取得します
'***************************************************************************************************
Property Get SubscribeCdpEventDetails()
    SubscribeCdpEventDetails = m_subToken.keys
End Property

'***************************************************************************************************
'* 機能　　：Environment作成時に使うオプション(`ICoreWebView2EnvironmentOptions`系)を取得します
'---------------------------------------------------------------------------------------------------
'* 注意事項：★重要★ 必ず`ConnectCDP`を呼ぶ**前**にこのプロパティ経由で値を設定すること。
'            Environment生成時にしか読まれないため、生成後の変更は反映されない
'***************************************************************************************************
Public Property Get EnvironmentOptions() As WebView2EnvOptions
    If m_envOptions Is Nothing Then Set m_envOptions = New WebView2EnvOptions
    Set EnvironmentOptions = m_envOptions
End Property



'***************************************************************************************************
'                                   ■■■ 接続と切断 ■■■
'***************************************************************************************************
'* 機能　　：`WebView2`方式としてCDP制御できる準備を行います(Environment/Controller/Viewの生成)
'---------------------------------------------------------------------------------------------------
'* 返り値  ：接続成功可否
'* 引数    ：UserName   WebView2のユーザーデータフォルダ名
'            AttachHwnd WebView2を貼り付けるハンドル ※省略時、Excel自身のハンドル(Application.Hwnd)を使用
'---------------------------------------------------------------------------------------------------
'* 注意事項：既存の接続がある場合は、切断されます
'***************************************************************************************************
Public Function ConnectCDP(UserName As String, Optional AttachHwnd As LongPtr) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ConnectCDP"


    '1. 既にオブジェクトを持ってる場合は切断後、再開
    If Not (m_Host Is Nothing) Then DisconnectCDP

    '2. ハンドル未指定の場合は、Excel自身を直接使用する
    If AttachHwnd <= 0 Then AttachHwnd = Application.hWnd

    '3. 専用のホストオブジェクトを新規作成
    Dim RunRes As Boolean
    Set m_Host = New WebView2Host

    '3-1. Environmentオプション(設定されていれば、ネイティブ側のfakeオブジェクトを構築)
    '     ★重要★ 実際に使う場合(=`m_envOptions`が既に生成済み)のみ構築する。未使用なら
    '     ネイティブfakeオブジェクトを一切構築せず、従来通り`environmentOptions=nullptr`の
    '     まま接続する(無駄な構築コストを避けるため)
    Dim pEnvOptions As LongPtr
    If Not (m_envOptions Is Nothing) Then pEnvOptions = m_envOptions.BuildNativePointer()

    With m_Host
        '3-2. 必要な情報を渡す
        .InitInfo UserName, AttachHwnd

        '3-3. セッションを閉じる(保険用)
        .cleanUpSessions

        '3-4. 起動！
        RunRes = .RunWebView2(pEnvOptions)
    End With

    '3-5. Environmentオプションのネイティブリソースは、ここで(成功/失敗問わず)解放してよい
    '     (`RunWebView2`内の完了待ちで、WebView2Loaderは既に読み終えている)
    DisposeEnvOptions

    '4. 正常に起動出来たか？
    If Not RunRes Then
        printMsg ERROR_, "WebView2 initialization failed. ErrorCode: " & m_Host.LastError & " - " & m_Host.LastErrorMessage, FromProcedureName

        Set m_Host = Nothing
        Exit Function
    End If

    '5. 成功サインを返す
    ConnectCDP = True
End Function

'***************************************************************************************************
'* 機能　　：接続を切断し、各種ハンドル/購読を解放します
'***************************************************************************************************
Public Sub DisconnectCDP()
    '1. 残っている購読を全解除
    UnsubscribeAllCdpEvents

    '2. ICoreWebView2_11(QI取得分)を解放
    If m_pView11 <> 0 Then
        WebView2Thunks.ComRelease m_pView11
        m_pView11 = 0
    End If
    m_pView11Failed = False

    '3. ホスト(Environment/Controller/View)を解放
    If Not (m_Host Is Nothing) Then
        m_Host.CLOSE_
        Set m_Host = Nothing
    End If

    '4. Environmentオプションが未使用のまま残っていれば、保険として解放
    '   (通常は`ConnectCDP`内で既に解放済みだが、一度も`ConnectCDP`を呼ばずに
    '   切断/破棄された場合への保険)
    DisposeEnvOptions
End Sub

'***************************************************************************************************
'* 機能　　：Environmentオプションのネイティブリソースを、設定済みなら解放します
'***************************************************************************************************
Private Sub DisposeEnvOptions()
    If Not (m_envOptions Is Nothing) Then
        m_envOptions.CLOSE_
        Set m_envOptions = Nothing
    End If
End Sub



'***************************************************************************************************
'                               ■■■ CDPコマンド送信関連 ■■■
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
'***************************************************************************************************
Public Function SendCommandCDP(strData As String) As Long
    Const FromProcedureName As String = ThisClassName & ".SendCommandCDP"
    Const E_HANDLE As Long = &H80070006   ' ハンドルが無効(=接続が既に切れている)


    '1. WebView2生存チェック
    '★重要★ ここで`0`(成功扱い)を返すと、呼び出し元(`CDPCore.writeWebView2CDP`経由で
    '`LastPipeError.WriteFile`)が「送信成功」と誤認し、実際には送られていないコマンドの
    '結果を`CDPBrowser.AutoWaitTakeResultCDP`が永遠に待ち続けてしまう
    '(Pipe/WebSocket版は、死んだハンドルへの`WriteFile`/`send`が実際にOSエラーを返すことで
    '`CDPBrowser.ExecuteCDPAsync`の`LastErrorWriteFile`チェックが機能するが、WebView2版は
    '自前でこの「接続が切れている」状態を明示的にエラーとして返す必要がある)
    If m_Host Is Nothing Then
        printMsg ERROR_, "No connected WebView2 (disconnected)", FromProcedureName

        SendCommandCDP = E_HANDLE
        Exit Function
    End If

    '2. `ICoreWebView2_1`ポインターの取得を試みる
    Dim pView As LongPtr: pView = m_Host.ViewPtr
    If pView = 0 Then
        printMsg ERROR_, "ViewPtr is 0 (disconnected or initialization incomplete).", FromProcedureName

        SendCommandCDP = E_HANDLE
        Exit Function
    End If

    '3. `CallDevToolsProtocolMethod`/`CallDevToolsProtocolMethodForSession`に渡すため、パラメータを分解する
    Dim Node As BiDiCDPJson: Set Node = BiDiCDPJson.Parse(strData)

    '3-1. スカラー値部分
    Dim commandID   As Long:                                       commandID = Node.NumberKey("id")
    Dim methodName  As String:                                    methodName = Node.StringKey("method")
    Dim sessionID   As String: If Node.ExistsKey("sessionId") Then sessionID = Node.StringKey("sessionId")

    '3-2.`params`オブジェクトのJson文字列化
    Dim paramsRaw As String
    If Node.ExistsKey("params") Then
        Dim paramsNode As BiDiCDPJson
        Set paramsNode = Node.NodeKey("params")
        If paramsNode Is Nothing Then paramsRaw = "{}" Else paramsRaw = paramsNode.Stringify()
    Else
        paramsRaw = "{}"
    End If

    '3-3. `commandID`をCOMコールバックメモリ領域の補助情報として登録して、来たるべき結果の紐づけを行う
    Dim h As WebView2CallbackHandler
    Set h = WebView2Thunks.AcquireHandlerFor(HK_CdpMethodCompleted, Me)
    If h Is Nothing Then
        printMsg ERROR_, "Failed to acquire the handler.", FromProcedureName

        SendCommandCDP = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If
    h.callbackId = commandID
    printMsg Debug_, "Registered with callback ID: " & commandID & ".", FromProcedureName

    '4. `SessionID`の有無に応じて使うコマンド実行CDP関数を決める
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + WebView2Thunks.VTABLE_OBJ_OFFSET
    Dim hr As Long
    If StrPtr(sessionID) Then
        '`SessionID`がある場合は、`CallDevToolsProtocolMethodForSession`で送る

        '4-1. 初代WebView2にはないため1度だけまず、インターフェース存在確認を行う
        'ICoreWebView2_11::CallDevToolsProtocolMethodForSession(vtableスロット99)
        'https://learn.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2.calldevtoolsprotocolmethodforsessionasync
        Dim pView11 As LongPtr: pView11 = EnsureView11(pView)
        If pView11 = 0 Then
            '対象のインターフェースがない場合は、取り扱えないので失敗とみなす
            printMsg ERROR_, "Cannot send a command specifying a `sessionId` because `ICoreWebView2_11` could not be obtained.", FromProcedureName

            Set h = Nothing
            SendCommandCDP = &H80004002   ' E_NOINTERFACE
            Exit Function
        End If

        '4-2. `CallDevToolsProtocolMethodForSession`で、`SessionID`付きで実行
        printMsg Debug_, "ICoreWebView2_11::CallDevToolsProtocolMethodForSession - sessionId:=" & sessionID & " , method:=" & methodName & " , parametersAsJson:=" & paramsRaw, FromProcedureName
        hr = WebView2Thunks.dcf(pView11, 99, "View11.CallDevToolsProtocolMethodForSession", _
                                    StrPtr(sessionID), StrPtr(methodName), StrPtr(paramsRaw), pVTableObj)
    Else
        '`SessionID`がない場合は、`CallDevToolsProtocolMethod`で送る
        printMsg Debug_, "ICoreWebView2::CallDevToolsProtocolMethod - method:=" & methodName & " , parametersAsJson:=" & paramsRaw, FromProcedureName

        '4-1. 初代WebView2で既にあるため、そのまま実行へ
        'ICoreWebView2::CallDevToolsProtocolMethod(vtableスロット36)
        'https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.web.webview2.core.corewebview2.calldevtoolsprotocolmethodasync
        hr = WebView2Thunks.dcf(pView, 36, "View.CallDevToolsProtocolMethod", _
                                    StrPtr(methodName), StrPtr(paramsRaw), pVTableObj)
    End If

    '4. WebView2側がAddRef済みなので、こちらのローカル参照は捨てる
    Set h = Nothing

    '5. 結果コードを返却
    SendCommandCDP = hr
End Function

'***************************************************************************************************
'* 機能　　：`ICoreWebView2`から`ICoreWebView2_11`をQueryInterfaceで取得します(1回だけ試行)
'***************************************************************************************************
Private Function EnsureView11(ByVal pView As LongPtr) As LongPtr
    Const IID_ICoreWebView2_11  As String = "0be78e56-c193-4051-b943-23b460c08bdb"  ' ICoreWebView2_11(CallDevToolsProtocolMethodForSessionを持つ派生IF)のIID


    If m_pView11 <> 0 Then EnsureView11 = m_pView11: Exit Function
    If m_pView11Failed Then Exit Function

    Dim iid As WebView2Thunks.GUID
    WebView2Thunks.FillGUID iid, IID_ICoreWebView2_11

    Dim pOut As LongPtr
    Dim hr As Long
    hr = WebView2Thunks.dcf(pView, 0, "View.QueryInterface(ICoreWebView2_11)", VarPtr(iid), VarPtr(pOut))
    If hr = 0 And pOut <> 0 Then
        m_pView11 = pOut
        EnsureView11 = m_pView11
    Else
        m_pView11Failed = True
    End If
End Function



'***************************************************************************************************
'                               ■■■ CDP受信関連 ■■■
'***************************************************************************************************
'* 機能　　：`CallDevToolsProtocolMethod(ForSession)`の完了通知を受け取り、CDP形式のJsonに
'            組み立てて発信します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：いわゆる、コマンド結果を受け取るところです。`SendCommandCDP`で、`commandID`を`callbackId`に登録したのと一致するのが来ます
'* 注意事項：`WebView2CallbackHandler.Handler_Invoke`から動的束縛で呼ばれるため、Public必須
'***************************************************************************************************
Public Sub OnCdpMethodCompleted(ByVal callbackId As Long, ByVal ErrorCode As Long, ByVal resultJson As String)
    Const FromProcedureName As String = ThisClassName & ".OnCdpMethodCompleted"


    printMsg Debug_, "ICoreWebView2CallDevToolsProtocolMethodCompletedHandler - callbackId:=" & callbackId & " , errorCode:=" & ErrorCode & " , result:=" & resultJson, FromProcedureName
    Dim synthesized As String
    If ErrorCode = 0 Then
        '上手い具合に、`id`付きで復元させる
        If LenB(resultJson) = 0 Then resultJson = "null"
        synthesized = "{""id"":" & callbackId & ",""result"":" & resultJson & "}"
    Else
        'エラーログを出す
        printMsg ERROR_, "ICoreWebView2CallDevToolsProtocolMethodCompletedHandler - ErrorCode: " & ErrorCode & " - " & WinApiError.GetMessage(ErrorCode), FromProcedureName

        '上手い具合に、`id`付きで、返ったエラーもくっつけて、復元させる
        Dim ErrorNode As BiDiCDPJson: Set ErrorNode = BiDiCDPJson.Parse(resultJson).NodeKey("error")
        synthesized = "{""id"":" & callbackId & ",""error"":" & ErrorNode.Stringify & "}"
    End If
    RaiseEvent CDPMessageReceived(synthesized)
End Sub

'***************************************************************************************************
'* 機能　　：`GetDevToolsProtocolEventReceived`の発火通知を受け取り、CDP形式のJsonに組み立てて
'            発信します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・いわゆる、CDPイベントを受け取るところです。
'            ・どのセッション由来かは`args`を`ICoreWebView2DevToolsProtocolEventReceivedEventArgs2`
'            　へQueryInterfaceし、`get_SessionId`で判別する(既定のトップレベルターゲット由来なら
'            　空文字が返る)。これにより`sessionId`を含めて再構成し、`CDPCore.
'            　BrowserReceivedDataCheck`が`CDPContextEvent`(タブ向け)/`CDPBrowserEvent`
'            　(ブラウザ向け)を正しく振り分けられるようにする。
'            ・`GetDevToolsProtocolEventReceiver`はトップレベルの`ICoreWebView2`に対して1回購読する
'            　だけで、`Target.setAutoAttach`等でアタッチされた子セッション(iframe/worker等)からの
'            　同名イベントも同じコールバックに届く(公式サンプルコードで確認済みの仕様)。
'***************************************************************************************************
Public Sub OnCdpEventReceived(ByVal EventName As String, ByVal pArgs As LongPtr)
    '1. `pArgs`ポインターから、`param`部分のJson文字列を取り出す
    Dim paramJson As String
    paramJson = WebView2Thunks.GetStringProperty(pArgs, 3, "Args.get_ParameterObjectAsJson")
    If LenB(paramJson) = 0 Then paramJson = "{}"

    '2. `pArgs`ポインターに紐づく、`SessionID`を取り出す
    Dim sessionID As String
    sessionID = TryGetEventSessionId(pArgs)

    '3. `SessinID`有無に応じた分岐
    Dim synthesized As String
    If LenB(sessionID) = 0 Then
        'ブラウザ単位として組み立て
        synthesized = "{""method"":""" & EventName & """,""params"":" & paramJson & "}"
    Else
        '`SessionID`単位として組み立て
        synthesized = "{""method"":""" & EventName & """,""params"":" & paramJson & _
                      ",""sessionId"":""" & sessionID & """}"
    End If

    '4. `CDPCore.cls`へ発信！
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

    Dim iid As WebView2Thunks.GUID
    WebView2Thunks.FillGUID iid, IID_Args2

    Dim pArgs2 As LongPtr
    Dim hr As Long
    hr = WebView2Thunks.dcf(pArgs, 0, "Args.QueryInterface(Args2)", VarPtr(iid), VarPtr(pArgs2))
    If hr <> 0 Or pArgs2 = 0 Then Exit Function

    TryGetEventSessionId = WebView2Thunks.GetStringProperty(pArgs2, 4, "Args2.get_SessionId")

    'QueryInterfaceで確保した分の参照を解放(GetStringPropertyはQI自体はしない)
    WebView2Thunks.ComRelease pArgs2
End Function



'***************************************************************************************************
'                               ■■■ CDPイベント購読 ■■■
'***************************************************************************************************
'* 機能　　：指定したCDPイベント名を、WebView2経由で購読します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：購読成功可否(既に購読済みの場合も`True`)
'* 引数　　：eventName   購読するCDPイベント名(例："Page.loadEventFired")
'---------------------------------------------------------------------------------------------------
'* 注意事項：・CDP-over-WebSocketと異なり、ドメインをenableしただけでは自動購読されません。
'            　必要なイベント名ごとに、このメソッドを個別に呼び出してください。
'            ・1つでも購読した場合、VBEの`Stop`やブレイクポイントを行うとクラッシュします。
'              どうしてもデバッグが必要な場合は､`UnsubscribeAllCdpEvents`で一旦解除したうえで、行ってください
'            ・このイベント購読は、各SessionIDごとに適用する必要はありません。`.newtab`で複製しても購読情報は継承されます
'***************************************************************************************************
Public Function SubscribeCdpEvent(ByVal EventName As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".SubscribeCdpEvent"


    '-----------------------------------------------------------------------------------------------
    ' 【ステップ1: 事前チェック】
    '-----------------------------------------------------------------------------------------------
    ' 1-1. WebView2のホストが死んでたら何もできないので即終了
    If m_Host Is Nothing Then Exit Function
    Dim pView As LongPtr: pView = m_Host.ViewPtr
    If pView = 0 Then Exit Function

    ' 1-2. すでに同じイベント(例: "Page.loadEventFired")を購読済みなら、二重登録せず成功扱いで抜ける
    If m_subReceiver.Exists(EventName) Then SubscribeCdpEvent = True: Exit Function

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ2: イベント専用の「受話器(Receiver)」をもらいに行く】
    '-----------------------------------------------------------------------------------------------
    ' WebView2に「"Page.loadEventFired" 専用の窓口オブジェクトを作ってくれ！」と注文する
    ' Vtable出席番号[42]: ICoreWebView2::GetDevToolsProtocolEventReceiver
    Dim pReceiver As LongPtr
    Dim hr As Long
    hr = WebView2Thunks.dcf(pView, 42, "View.GetDevToolsProtocolEventReceiver", _
                               StrPtr(EventName), VarPtr(pReceiver))
    If hr <> 0 Or pReceiver = 0 Then
        printMsg ERROR_, "Failed to GetDevToolsProtocolEventReceiver.  hr=&H" & Hex(hr), FromProcedureName
        Exit Function
    End If

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ3: ニセC++窓口(配達員)をプールから借りて準備する】
    '-----------------------------------------------------------------------------------------------
    ' 3-1. 共有スロットプールから空き部屋を1つ借りて、ニセの受付ブース(h)を作る
    Dim h As WebView2CallbackHandler
    Set h = WebView2Thunks.AcquireHandlerFor(HK_CdpEventReceived, Me)
    If h Is Nothing Then
        ' 失敗したらステップ2でもらった受話器(pReceiver)をお片付け(Release)して抜ける
        WebView2Thunks.ComRelease pReceiver
        Exit Function
    End If

    ' 3-2. ★超重要: イベント通知が届いた時に「何のイベントか」識別できるよう、
    '               ハンドラーのポケットにイベント名をメモしておく！
    h.EventName = EventName

    ' 3-3. C++に見せるための「お面(Vtableポインタ)」の先頭アドレスを計算
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + WebView2Thunks.VTABLE_OBJ_OFFSET

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ4: 受話器にニセ窓口をガッチャンコして登録する】
    '-----------------------------------------------------------------------------------------------
    ' 受話器の Vtable出席番号[3]: add_DevToolsProtocolEventReceived を叩く
    ' 登録に成功すると、あとで購読解除(解約)するときに使う「解約用チケット(token: 64bit整数)」がもらえる
    Dim token As LongLong
    hr = WebView2Thunks.dcf(pReceiver, 3, "Receiver.add_DevToolsProtocolEventReceived", _
                               pVTableObj, VarPtr(token))
    If hr <> 0 Then
        printMsg ERROR_, "Failed to add_DevToolsProtocolEventReceived.  hr=&H" & Hex(hr), FromProcedureName

        ' 失敗したら配達員と受話器を両方片付けて抜ける
        Set h = Nothing
        WebView2Thunks.ComRelease pReceiver
        Exit Function
    End If

    '-----------------------------------------------------------------------------------------------
    ' 【ステップ5: 送りっぱなし ＆ 解約用チケットの保管】
    '-----------------------------------------------------------------------------------------------
    ' 5-1. ★出前完了: WebView2(C++)側がハンドラーをAddRefしてキープしてくれたので、
    '                  VBA側の手元の変数は捨てて「送りっぱなし」にしてOK！
    Set h = Nothing

    ' 5-2. 将来「このイベントの購読をやめる(Unsubscribe)」ときのために、
    '      受話器のアドレス(pReceiver)と解約チケット(token)を辞書にメモしておく
    m_subReceiver.Add EventName, (pReceiver)
    m_subToken.Add EventName, (token)

    '6. 成功サインを返す
    SubscribeCdpEvent = True
End Function

'***************************************************************************************************
'* 機能　　：指定したCDPイベント名の購読を解除します
'***************************************************************************************************
Public Function UnsubscribeCdpEvent(ByVal EventName As String) As Boolean
    ' 1. 未購読なら即抜け
    If m_subReceiver Is Nothing Then Exit Function
    If Not m_subReceiver.Exists(EventName) Then Exit Function

    ' 2. 対象購読イベントに紐づくレシーバーポインター/トークンを取り出す
    Dim pReceiver As LongPtr: pReceiver = m_subReceiver(EventName)
    Dim token As LongLong: token = m_subToken(EventName)

    ' 3. トークンを渡して、購読解除する
    WebView2Thunks.dcf pReceiver, 4, "Receiver.remove_DevToolsProtocolEventReceived", token

    ' 4. レシーバーポインターの参照をカット
    WebView2Thunks.ComRelease pReceiver

    ' 5. 必要な解放を終えたので、Dictionaryからも消しておく
    m_subReceiver.Remove EventName
    m_subToken.Remove EventName

    ' 6. 成功サインを返す
    UnsubscribeCdpEvent = True
End Function

'***************************************************************************************************
'* 機能　　：購読中の全CDPイベントを一括解除します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：解除した件数
'---------------------------------------------------------------------------------------------------
'* 詳細説明：これを行うことで、VBEの`Stop`やブレイクポイントをしてもクラッシュしなくなります
'* 注意事項：WebView2の`GetDevToolsProtocolEventReceiver`自体には一括購読/一括解除の概念はなく、
'            イベント名ごとの個別登録が必須です(CDP-over-WebSocketの「ドメインをenableすれば
'            以後そのドメインの全イベントが自動で流れる」モデルとは異なります)。このメソッドは、
'            `SubscribeCdpEvent`で購読済みのイベント名を`m_subReceiver`から一覧取得し、
'            `UnsubscribeCdpEvent`を1件ずつ回すことで「一発解除」に見せているだけです。
'***************************************************************************************************
Public Function UnsubscribeAllCdpEvents() As Long
    If m_subReceiver Is Nothing Then Exit Function

    '1. 購読中のイベント数を確認し、1つもない場合は、即抜け
    Dim N As Long: N = m_subReceiver.Count
    If N = 0 Then Exit Function

    '2. 購読中のイベント名リストを取り出す
    Dim EventNames
    EventNames = m_subReceiver.keys

    '3. Forで、１つずつ購読解除していく
    Dim i As Long
    For i = 0 To N - 1
        If UnsubscribeCdpEvent(EventNames(i)) Then UnsubscribeAllCdpEvents = UnsubscribeAllCdpEvents + 1
    Next i
End Function



'***************************************************************************************************
'                                ■■■ WebView2の挙動変更 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] WebView2の構成系
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：リサイズ処理を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：WV2_Width      横の長さ
'            WV2_Height     縦の長さ
'            WV2_Top        上端からのオフセット値
'            WV2_Left       左端からのオフセット値
'***************************************************************************************************
Public Sub Resize(Optional WV2_Width As Long = 4, Optional WV2_Height As Long = 4, Optional WV2_Top As Long, Optional WV2_Left As Long)
    m_Host.Resize WV2_Width, WV2_Height, WV2_Top, WV2_Left
End Sub

'***************************************************************************************************
'* 機能　　：表示/非表示切り替え
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Web描画の表示/非表示
'***************************************************************************************************
Property Let Visible(v As Boolean)
    m_Host.Visible = v
End Property

'* 機能　　：ブラウザUIレベルのズーム倍率を設定します(CDPの`Emulation.setPageScaleFactor`とは別物)
Property Let ZoomFactor(v As Double)
    m_Host.ZoomFactor = v
End Property

'* 機能　　：表示位置/サイズとズーム倍率を1回のレイアウト計算でまとめて設定します
Public Sub SetBoundsAndZoomFactor(WV2_Right As Long, WV2_Bottom As Long, WV2_Top As Long, WV2_Left As Long, ZoomFactor As Double)
    m_Host.SetBoundsAndZoomFactor WV2_Right, WV2_Bottom, WV2_Top, WV2_Left, ZoomFactor
End Sub

'* 機能　　：フォーカスをWebView2側へプログラム的に移動します
Public Sub MoveFocus(Optional reason As WV2MoveFocusReason = MoveFocus_Programmatic)
    m_Host.MoveFocus reason
End Sub

'* 機能　　：親ウィンドウの位置が動いたことをWebView2へ明示的に通知します
Public Sub NotifyParentWindowPositionChanged()
    m_Host.NotifyParentWindowPositionChanged
End Sub

'* 機能　　：WebView2のページ読み込み前に表示される既定の背景色を設定します(Controller2以降)
Public Sub SetDefaultBackgroundColor(a As Byte, r As Byte, g As Byte, b As Byte)
    m_Host.SetDefaultBackgroundColor a, r, g, b
End Sub

'* 機能　　：DPIスケールを手動指定します(Controller3以降)
Property Let RasterizationScale(v As Double)
    m_Host.RasterizationScale = v
End Property

'* 機能　　：モニターのDPIスケール変更を自動検知するか切り替えます(Controller3以降、既定True)
Property Let ShouldDetectMonitorScaleChanges(v As Boolean)
    m_Host.ShouldDetectMonitorScaleChanges = v
End Property

'* 機能　　：`Bounds`(位置/サイズ)の解釈方式を切り替えます(Controller3以降)
Property Let BoundsMode(v As WV2BoundsMode)
    m_Host.BoundsMode = v
End Property

'* 機能　　：外部からのドラッグ&ドロップ受け入れ可否を切り替えます(Controller4以降、既定True)
Property Let AllowExternalDrop(v As Boolean)
    m_Host.AllowExternalDrop = v
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] interface - ICoreWebView2Profile
'---------------------------------------------------------------------------------------------------
'* 注意事項：`IsPasswordAutosaveEnabled`/`IsGeneralAutofillEnabled`はSettings側と紛らわしいので
'            `Profile`を頭に付けている(詳細は`CDPWebView2Host`側のコメント参照)
'---------------------------------------------------------------------------------------------------
'* 機能　　：既定のダウンロード先フォルダを設定します
Property Let DefaultDownloadFolderPath(v As String)
    m_Host.DefaultDownloadFolderPath = v
End Property

'* 機能　　：優先カラースキームを設定します(既定Auto)
Property Let PreferredColorScheme(v As WV2PreferredColorScheme)
    m_Host.PreferredColorScheme = v
End Property

'* 機能　　：トラッキング防止レベルを設定します(Profile3以降、既定Balanced)
Property Let PreferredTrackingPreventionLevel(v As WV2TrackingPreventionLevel)
    m_Host.PreferredTrackingPreventionLevel = v
End Property

'* 機能　　：パスワード自動保存機能の有効可否を、プロファイル単位で切り替えます(Profile6以降)
Property Let ProfilePasswordAutosaveEnabled(v As Boolean)
    m_Host.ProfilePasswordAutosaveEnabled = v
End Property

'* 機能　　：一般オートフィル機能の有効可否を、プロファイル単位で切り替えます(Profile6以降)
Property Let ProfileGeneralAutofillEnabled(v As Boolean)
    m_Host.ProfileGeneralAutofillEnabled = v
End Property

'* 機能　　：Service Worker内からWebView固有のスクリプトAPIを使えるようにするか切り替えます(Profile9以降)
Property Let WebViewScriptApisEnabledForServiceWorkers(v As Boolean)
    m_Host.WebViewScriptApisEnabledForServiceWorkers = v
End Property

'***************************************************************************************************
'* 機能　　：指定フォルダの未パック拡張機能を、プロファイルへインストールします(Profile7以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：extensionFolderPath  拡張機能の`manifest.json`を含むトップフォルダの絶対パス
'* 返り値  ：インストールされた拡張機能のID(失敗時は空文字)
'---------------------------------------------------------------------------------------------------
'* 注意事項：★重要★ 事前に`AreBrowserExtensionsEnabled = True`を設定して`ConnectCDP`していないと
'            `ERROR_NOT_SUPPORTED`で失敗します。CDPの`Extensions.loadUnpacked`はWebView2では
'            使えない(`Method not available`)ため、この専用APIを使う必要があります
'***************************************************************************************************
Public Function AddBrowserExtension(ByVal extensionFolderPath As String) As String
    If m_Host Is Nothing Then Exit Function
    AddBrowserExtension = m_Host.AddBrowserExtension(extensionFolderPath)
End Function

'***************************************************************************************************
'* 機能　　：インストール済みの拡張機能IDを一覧取得します(Profile7以降)
'---------------------------------------------------------------------------------------------------
'* 返り値  ：拡張機能IDの配列(String)。0件/失敗時は要素数0の配列
'***************************************************************************************************
Public Function GetBrowserExtensionIds() As Variant
    If m_Host Is Nothing Then
        GetBrowserExtensionIds = Array()
        Exit Function
    End If
    GetBrowserExtensionIds = m_Host.GetBrowserExtensionIds()
End Function

'***************************************************************************************************
'* 機能　　：指定IDの拡張機能をアンインストールします(Profile7以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：extensionId  `AddBrowserExtension`/`GetBrowserExtensionIds`で得られる拡張機能ID
'* 返り値  ：成功可否
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebView2にはID指定の直接アンインストールAPIが無いため、内部で一覧取得→対象特定→
'            `Remove`という2段構えになります(CDPの`Extensions.uninstall`は使えないため代替不可)
'***************************************************************************************************
Public Function RemoveBrowserExtension(ByVal extensionId As String) As Boolean
    If m_Host Is Nothing Then Exit Function
    RemoveBrowserExtension = m_Host.RemoveBrowserExtension(extensionId)
End Function

'***************************************************************************************************
'* 機能　　：ローカルフォルダを、仮想ホスト名経由(`https://<hostName>/`のような形)で配信されている
'            かのように見せます(ICoreWebView2_3以降)
'---------------------------------------------------------------------------------------------------
'* 引数    ：hostName    仮想ホスト名(例:"my.app")
'            folderPath  配信対象のローカルフォルダの絶対パス
'            accessKind  既定は`HostAccess_Allow`
'* 返り値  ：成功可否
'***************************************************************************************************
Public Function SetVirtualHostNameToFolderMapping(ByVal hostname As String, ByVal FolderPath As String, _
    Optional ByVal accessKind As WV2HostResourceAccessKind = HostAccess_Allow) As Boolean

    If m_Host Is Nothing Then Exit Function
    SetVirtualHostNameToFolderMapping = m_Host.SetVirtualHostNameToFolderMapping(hostname, FolderPath, accessKind)
End Function

'***************************************************************************************************
'* 機能　　：`SetVirtualHostNameToFolderMapping`で設定したマッピングを解除します(ICoreWebView2_3以降)
'***************************************************************************************************
Public Sub ClearVirtualHostNameToFolderMapping(ByVal hostname As String)
    If m_Host Is Nothing Then Exit Sub
    m_Host.ClearVirtualHostNameToFolderMapping hostname
End Sub

'***************************************************************************************************
'* 機能　　：WebView2インスタンス全体の音声出力をミュートするか切り替えます(ICoreWebView2_8以降)
'***************************************************************************************************
Property Get IsMuted() As Boolean
    If m_Host Is Nothing Then Exit Property
    IsMuted = m_Host.IsMuted
End Property

Property Let IsMuted(v As Boolean)
    If m_Host Is Nothing Then Exit Property
    m_Host.IsMuted = v
End Property

'***************************************************************************************************
'* 機能　　：現在ドキュメントが音声を再生中かどうかを取得します(ICoreWebView2_8以降、読み取り専用)
'***************************************************************************************************
Property Get IsDocumentPlayingAudio() As Boolean
    If m_Host Is Nothing Then Exit Property
    IsDocumentPlayingAudio = m_Host.IsDocumentPlayingAudio
End Property

'***************************************************************************************************
'* 機能　　：リンクホバー時のリンク先URL等、ステータスバーに表示される文字列を取得します
'            (ICoreWebView2_12以降、読み取り専用)
'***************************************************************************************************
Property Get StatusBarText() As String
    If m_Host Is Nothing Then Exit Property
    StatusBarText = m_Host.StatusBarText
End Property

'***************************************************************************************************
'* 機能　　：現在のページのfavicon URIを取得します(ICoreWebView2_15以降、読み取り専用)
'***************************************************************************************************
Property Get FaviconUri() As String
    If m_Host Is Nothing Then Exit Property
    FaviconUri = m_Host.FaviconUri
End Property

'***************************************************************************************************
'* 機能　　：現在アクティブな検索結果の(1始まりの)インデックスを取得します(ICoreWebView2_28以降)
'***************************************************************************************************
Property Get FindActiveMatchIndex() As Long
    If m_Host Is Nothing Then Exit Property
    FindActiveMatchIndex = m_Host.FindActiveMatchIndex
End Property

'***************************************************************************************************
'* 機能　　：現在の検索結果の総件数を取得します(ICoreWebView2_28以降)
'***************************************************************************************************
Property Get FindMatchCount() As Long
    If m_Host Is Nothing Then Exit Property
    FindMatchCount = m_Host.FindMatchCount
End Property

'***************************************************************************************************
'* 機能　　：次の検索結果へ移動します(ICoreWebView2_28以降)
'***************************************************************************************************
Public Sub FindNext()
    If m_Host Is Nothing Then Exit Sub
    m_Host.FindNext
End Sub

'***************************************************************************************************
'* 機能　　：前の検索結果へ移動します(ICoreWebView2_28以降)
'***************************************************************************************************
Public Sub FindPrevious()
    If m_Host Is Nothing Then Exit Sub
    m_Host.FindPrevious
End Sub

'***************************************************************************************************
'* 機能　　：検索を終了し、ハイライト表示を消去します(ICoreWebView2_28以降)
'***************************************************************************************************
Public Sub FindStop()
    If m_Host Is Nothing Then Exit Sub
    m_Host.FindStop
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] WebView2の機能制限系
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：DevTools(F12キー・右クリック「検証」)の起動可否を切り替えます
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'---------------------------------------------------------------------------------------------------
'* 注意事項：CDPリモートデバッグ(`--remote-debugging-pipe`)経由の自動操作には影響しません
'            (詳細は`CDPWebView2Host.DevToolsEnabled`のコメント参照)
'***************************************************************************************************
Property Let DevToolsEnabled(v As Boolean)
    m_Host.DevToolsEnabled = v
End Property

'***************************************************************************************************
'* 機能　　：既定の右クリックコンテキストメニューの表示可否を切り替えます
'---------------------------------------------------------------------------------------------------
'* 引数    ：v  Trueで許可、Falseで禁止
'***************************************************************************************************
Property Let ContextMenuEnabled(v As Boolean)
    m_Host.ContextMenuEnabled = v
End Property

'--- 以下、`ICoreWebView2Settings`(基底)の残りのプロパティ(詳細は`CDPWebView2Host`側のコメント参照) ---
'* 機能　　：JavaScript実行可否を切り替えます(既定True)
Property Let ScriptEnabled(v As Boolean)
    m_Host.ScriptEnabled = v
End Property

'* 機能　　：`postMessage`等によるWebページ⇔ホストアプリ間の通信可否を切り替えます(既定True)
Property Let WebMessageEnabled(v As Boolean)
    m_Host.WebMessageEnabled = v
End Property

'* 機能　　：`alert`/`confirm`/`prompt`等、既定のJavaScriptダイアログ表示可否を切り替えます(既定True)
Property Let DefaultScriptDialogsEnabled(v As Boolean)
    m_Host.DefaultScriptDialogsEnabled = v
End Property

'* 機能　　：ステータスバーの表示可否を切り替えます(既定True)
Property Let StatusBarEnabled(v As Boolean)
    m_Host.StatusBarEnabled = v
End Property

'* 機能　　：`AddHostObjectToScript`で追加したホストオブジェクトへのスクリプト側アクセス可否を
'            切り替えます(既定True)
Property Let HostObjectsAllowed(v As Boolean)
    m_Host.HostObjectsAllowed = v
End Property

'* 機能　　：ユーザー操作(Ctrl+ホイール等)によるズーム可否を切り替えます(既定True)
Property Let ZoomControlEnabled(v As Boolean)
    m_Host.ZoomControlEnabled = v
End Property

'* 機能　　：WebView2内蔵のエラーページ表示可否を切り替えます(既定True)
Property Let BuiltInErrorPageEnabled(v As Boolean)
    m_Host.BuiltInErrorPageEnabled = v
End Property

'--- 以下、`ICoreWebView2Settings2`~`9`(古いWebView2 Runtimeでは無視されるだけで安全) ---
'* 機能　　：User-Agent文字列を上書きします
Property Let UserAgentOverride(v As String)
    m_Host.UserAgentOverride = v
End Property

'* 機能　　：`Ctrl+F`/`F5`/`F12`等、ブラウザ的なアクセラレータキーの有効可否を切り替えます(既定True)
Property Let BrowserAcceleratorKeysEnabled(v As Boolean)
    m_Host.BrowserAcceleratorKeysEnabled = v
End Property

'* 機能　　：パスワード自動保存機能の有効可否を切り替えます(既定True)
Property Let PasswordAutosaveEnabled(v As Boolean)
    m_Host.PasswordAutosaveEnabled = v
End Property

'* 機能　　：住所/フォーム入力等の一般オートフィル機能(パスワード以外)の有効可否を切り替えます(既定True)
Property Let GeneralAutofillEnabled(v As Boolean)
    m_Host.GeneralAutofillEnabled = v
End Property

'* 機能　　：ピンチズームジェスチャーの有効可否を切り替えます(既定True)
Property Let PinchZoomEnabled(v As Boolean)
    m_Host.PinchZoomEnabled = v
End Property

'* 機能　　：スワイプによる進む/戻るナビゲーションジェスチャーの有効可否を切り替えます(既定True)
Property Let SwipeNavigationEnabled(v As Boolean)
    m_Host.SwipeNavigationEnabled = v
End Property

'***************************************************************************************************
'* 機能　　：PDFビューア内蔵ツールバーの非表示ボタンを、`WV2PdfToolbarItems`のEnumを`Or`で
'            組み合わせて指定します(例: `PDF_Save Or PDF_PrintItem`)
'---------------------------------------------------------------------------------------------------
'* 引数　　：v  非表示にしたい項目を`Or`でつないだもの(既定0相当=何も非表示にしない=全表示)
'---------------------------------------------------------------------------------------------------
'* 注意事項：呼び出すたびに全項目を上書きします(前回の指定は引き継がれません)
'***************************************************************************************************
Property Let HiddenPdfToolbarItems(v As WV2PdfToolbarItems)
    m_Host.HiddenPdfToolbarItems = v
End Property

'* 機能　　：SmartScreen(フィッシング/マルウェア検知)の要否を切り替えます(既定True)
Property Let ReputationCheckingRequired(v As Boolean)
    m_Host.ReputationCheckingRequired = v
End Property

'* 機能　　：CSS`app-region`による独自タイトルバー領域(ドラッグでウィンドウ移動可能な領域)の
'            指定を有効にするか切り替えます(既定False)
Property Let NonClientRegionSupportEnabled(v As Boolean)
    m_Host.NonClientRegionSupportEnabled = v
End Property



'***************************************************************************************************
'                                  ■■■ ログヘルパー ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Const LogFileExtension As String = ".WebView2.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath, vbObjectError + 1
End Sub



'***************************************************************************************************
'                                   ■■■ 初期化/後始末 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    Set m_subReceiver = New Dictionary
    Set m_subToken = New Dictionary
    currentLog.LogID = "WVID" & Format(Rnd * 1000, "000")

    'ログ設定
    currentLog.seeRawSendMsgDbg = ShSetting01_StartBrowser.UseRangeID(8, FromProcedureName)
    currentLog.logPath = ShSetting01_StartBrowser.UseRangeID(9, FromProcedureName)
End Sub

Private Sub Class_Terminate()
    DisconnectCDP
End Sub
