VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverBiDiCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'               Automating Chromium-Based Browsers with WebDriverBiDi API and VBA
'--------------------------------------------------------------------------------------------------------------
' WebDriver BiDiの低レベル通信を担うクラス。
' CDPContext経由でBiDiマッパーを実行し、BiDiコマンドの送受信・イベント取得を行います。
' `RaiseEvent`を利用した`CDPCore.cls`→`WebDriverBiDiCore.cls`→`WebDriverBiDiMode.cls`といった連鎖で成り立ってます
' `CDPCore.cls`をベースに処理ロジックを実装しており、軽量化してます
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'「BiDi-CDP Mapper」に関する情報
Private Const ThisClassName As String = "WebDriverBiDiCore"             'ログのトレース用
Private Const MapperURL     As String = "about:blank#" & ThisClassName  '`BiDi-CDP Mapper`の目印URL
Private Const BiDiResDebug  As String = "sendDebugMessage"              '`BiDi-CDP Mapper`内のJSがデバッグ情報を返す際の関数名(chromium-bidi\src\bidiServer\MapperCdpConnection.ts に基づく名称)
Private Const BiDiResMain   As String = "sendBidiResponse"              '`BiDi-CDP Mapper`内のJSが結果を返す際の関数名        (chromium-bidi\src\bidiServer\MapperCdpConnection.ts に基づく名称)
Private Const BiDiSend      As String = "onBidiMessage"                 '`BiDi-CDP Mapper`へコマンド送信する際の関数名        (chromium-bidi\src\bidiTab\bidiTab.ts#L45)

'nmp情報
Private Const MapperJSName  As String = "mapperTab.js"  'WebDriverBiDiのやり取りの核となるJavaScriptファイル名
Private Const npmName       As String = "chromium-bidi" '公式リポジトリ名称`https://github.com/GoogleChromeLabs/chromium-bidi`

'その他
Private Const LimitCommandID    As Long = 2000000000    'BiDiコマンド送信時のID上限値



'***************************************************************************************************
'                                   ■■■ イベント定義 ■■■
'***************************************************************************************************
'BiDiレスポンスの種類に応じた発信機
Public Event detectionBiDiEvent(methodName As String, RawJson As String)    '非同期イベント
Public Event detectionBiDiID(id As Long, RawJson As String)                 'BiDiコマンド結果

'コマンドIDリセットサイン
Public Event ResetCommandID()



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'ブラウザ制御コアクラスオブジェクト
Private WithEvents m_CDPCore    As CDPCore      'イベント受け取りオブジェクト
Attribute m_CDPCore.VB_VarHelpID = -1
Private m_CDPBrowser            As CDPBrowser   'ブラウザ操作用
Private WebDriverBiDi_mapperTab As CDPContext   'BiDi-CDP Mapper 操作用タブ

'「BiDi-CDP Mapper」のデバッグ設定状況
Private Type StateDebug
    ViewDebugTab        As Boolean      '「BiDi-CDP Mapper is controlling this tab」といったデバックタブ画面を表示するか？※非表示でも、「edge://inspect/#other」で存在確認はできます。
    DiagnosticLog       As Boolean      '「BiDi-CDP Mapper」内の診断ログ情報を出力するか
    seeRawSendMsgDbg    As Boolean      'BiDiのresponseをデバッグログに出すかの調整
End Type
Private BiDiDebug As StateDebug

'ステート管理
Private m_MessageID     As Long     '現在のBiDiコマンドID
Private LogID           As String   'このClassオブジェクトとしてのログID



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：「BiDi-CDP Mapper」を付けた状態で、ブラウザに対する所定の起動引数をつけて起動させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：Name                           使用ブラウザ名※現時点では、Edgeとchromeのみです
'            appUrl                         最初の起動時に遷移したいURL
'            userProfile                    UserDataDir名を指定します
'            addArgs                        追加の起動引数を設けます
'            KioskMode                      0(省略)：通常モード(キオスクモードは使いません)
'                                           1      ：キオスクモード デジタル/対話型サイネージ
'                                           2      ：キオスクモード パブリック ブラウジング
'
'            sessionCapabilitiesRequest     `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'---------------------------------------------------------------------------------------------------
'* 機能説明：・BiDiマッパーのスクリプトをExcelテーブルから読み込み、対象のコンテキストで実行可能にします。
'            ・新規立ち上げ用に使い、セッション確立まで担います。
'***************************************************************************************************
Public Sub InitBiDi(Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional userProfile As String, Optional addArgs As String, Optional KioskMode As edgeKioskType, _
    Optional sessionCapabilitiesRequest As Dictionary)

    Const FromProcedureName As String = ThisClassName & ".start"


    '1. 引数から設定を適用し、まずはCDP-browserとして起動
    Set m_CDPBrowser = New CDPBrowser
    m_CDPBrowser.start Name, appUrl, userProfile, addArgs, KioskMode

    '2. WithEventsClassオブジェクトを受け取る
    Set m_CDPCore = m_CDPBrowser.InheritanceCDPCore

    '3. `BiDi-CDP Mapper`用のタブを新規で、用意します
    newTabForBiDi

    '4. `BiDi-CDP Mapper`を始動させます
    StartBiDiCDPMapper
End Sub

'***************************************************************************************************
'* 機能　　：新規タブにて、`BiDi-CDP Mapper`環境を準備します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・各パラメータは、`chromium-bidi\src\bidiServer\MapperCdpConnection.ts`に準拠してます
'            ・普通の新しいタブOpenとは別の手法を用いるため、`CDPBrowser - newtab`とは別のプロシージャで用意します
'***************************************************************************************************
Private Sub newTabForBiDi()
    Const FromProcedureName As String = ThisClassName & ".newTabForBiDi"


    '1. BiDi用に、各種パラメータ運用
    Dim params As New Dictionary
    params.Add "url", MapperURL     '目印用のURLを付与
    params.Add "background", True   'フォアグランドさせない
    If Not (BiDiDebug.ViewDebugTab) Then params.Add "hidden", Not (BiDiDebug.ViewDebugTab)  'タブを見せないか？

    With m_CDPBrowser
        '2. そのパラメータで、新しいタブを開き、`targetId`を取得
        Dim tabId As String
        tabId = .ExecuteCDP("Target.createTarget", params).StringKey("targetId")

        '3. 得た`targetId`に対して、JavaScriptによるCDP制御の許可を与える
        params.RemoveAll
        params.Add "bindingName", "cdp"
        params.Add "targetId", tabId
        params.Add "inheritPermissions", True
        .ExecuteCDP "Target.exposeDevToolsProtocol", params

        '4. 得た`targetId`で、`sessionId`を取得
        Dim sessionID As String
        sessionID = .attachToTab(tabId)

        '5. Excelに、BiDiMapperTab情報を記録してメイン化しておく
        .serializeForMainTab sessionID, tabId
    End With

    '6. このクラス内で用意してる`WebDriverBiDi_mapperTab`にいくつか情報を渡す
    Set WebDriverBiDi_mapperTab = New CDPContext
    With WebDriverBiDi_mapperTab
        '6-1. このブラウザのパイプハンドル情報を渡す
        Set .InheritanceCDPCore = m_CDPBrowser.InheritanceCDPCore

        '6-2. このブラウザ情報を渡す
        Set .InheritanceCDPBrowser = m_CDPBrowser

        '6-3. 先ほど得た`sessionId`を渡す
        .CurrentSessionID = sessionID
    End With

    '7. ログ化
    printMsg Debug_, LogID, "Ready `BiDi-CDP Mapper`... (ViewTab: " & BiDiDebug.ViewDebugTab & ")", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：`BiDi-CDP Mapper`を始動させます
'---------------------------------------------------------------------------------------------------
'* 注意事項：各パラメータは、`chromium-bidi\src\bidiServer\MapperCdpConnection.ts`に準拠してます
'***************************************************************************************************
Private Sub StartBiDiCDPMapper()
    Const FromProcedureName As String = ThisClassName & ".StartBiDiCDPMapper"


    '1. JSからVBAへ返す経路を設定
    Call BiDiResponseBiding

    'タブ単位実行
    With WebDriverBiDi_mapperTab
        '2. パースしたJavaScript - mapperTab.js をブラウザ上で実行
        .jsEval ShLibrary01_JS.TableMapperTabJS

        '3. ブートストラップ（初期化）JSの実行
        .jsEval "window.runMapperInstance('" & .CurrentTargetID & "')"
    End With

    printMsg Debug_, LogID, "`BiDi-CDP Mapper` has been launched", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：BiDiからの結果やデバッグログを受け取るためのマッピング処理を行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`CDPContext`を`New`してからご利用ください
'            ・各パラメータは、`chromium-bidi\src\bidiServer\MapperCdpConnection.ts`に準拠してます
'            ・SessionIDが変わるたびにこれを行わないと、`BiDi-CDP Mapper`は生きてるのに、コマンドのやり取りが出来ない問題が起こります
'***************************************************************************************************
Private Sub BiDiResponseBiding()
    Const FromProcedureName As String = ThisClassName & ".BiDiResponseBiding"


    'タブ単位実行
    Dim paramsCDP As New Dictionary
    With WebDriverBiDi_mapperTab
        '1. JSからVBAへ、BiDiコマンドの結果/イベント payload を返す経路を設定
        paramsCDP.Add "name", BiDiResMain
        .ExecuteCDP "Runtime.addBinding", paramsCDP

        '2. JSからVBAへ、Mapper側で発生したログ（状況説明、警告、エラー寄り情報）を返す経路を設定
        If BiDiDebug.DiagnosticLog Then
            paramsCDP.RemoveAll
            paramsCDP.Add "name", BiDiResDebug
            .ExecuteCDP "Runtime.addBinding", paramsCDP
        End If
    End With

    printMsg Debug_, LogID, "I have configured the route to receive from `BiDi-CDP Mapper`.", FromProcedureName
End Sub



'***************************************************************************************************
'                               ■■■ 再接続処理系 ■■■
'***************************************************************************************************
'* 機能　　：Excelに記録してるタブ情報を基に、既存の`BiDi-CDP Mapper`に再接続します
'---------------------------------------------------------------------------------------------------
'* 引数　　：userProfile                再アタッチしたいユーザー名(user-data-dirに基づく)
'            WebSocketMode              WebSocketでCDP制御する場合は、接続処理済みのClassオブジェクトを指定します
'* 返値    ：再始動したか論理値 `True` ：`BiDi-CDP Mapper`を再始動しました。`WebDriverBiDiMode.cls`側で、`session.new`する必要があります
'                               `False`：既存の`BiDi-CDP Mapper`を流用できたので、`session.new`は不要です
'---------------------------------------------------------------------------------------------------
'* 機能説明：あとから、`BiDi-CDP Mapper`セッションに再接続する際に使い、セッション確立まで担います
'* 注意事項：・以前のタブが失ってもブラウザハンドル情報が生きてる場合は、新しいタブとして復旧し、再開します
'            ・ブラウザ自体が落ちてると、VBAエラーで止まります
'***************************************************************************************************
Public Function ReBootBiDiCDPMapper(userProfile As String, Optional WebSocketMode As CDPCoreViaWebSocket) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ReBootBiDiCDPMapper"


    '1. CDPタブセッションに再接続
    Set WebDriverBiDi_mapperTab = New CDPContext

    '2. Excelに記録されてる`TargetID`の生存確認　※`sessionId`は更新させます
    Dim mapperTabConnection As Boolean
    mapperTabConnection = WebDriverBiDi_mapperTab.reattach(userProfile, WebSocketMode:=WebSocketMode)

    '3. 各種オブジェクトを継承
    Set m_CDPCore = WebDriverBiDi_mapperTab.InheritanceCDPCore
    Set m_CDPBrowser = WebDriverBiDi_mapperTab.InheritanceCDPBrowser

    '`TargetID`有無に応じた分岐
    If mapperTabConnection Then
        'Excelに記録されてる`TargetID`が生きてる = `Target.exposeDevToolsProtocol`が有効
        printMsg info_, LogID, "`Target.exposeDevToolsProtocol` instance is already running. Reattaching to existing BiDi session...", FromProcedureName

        '4. そのタブURLは、このクラスで用意した専用URLか？
        If StrComp(MapperURL, WebDriverBiDi_mapperTab.Url, vbTextCompare) = 0 Then
            '5. すでに当該タブで `BiDi-CDP Mapper`が起動済みかを確認する
            '→ typeof window.onBidiMessage === 'function' かどうかを調べる
            Dim isMapperRunning As Boolean
            isMapperRunning = WebDriverBiDi_mapperTab.jsEval("typeof window." & BiDiSend & " === 'function';")

            '`BiDi-CDP Mapper`の起動状況に応じた分岐
            If isMapperRunning Then
                '6. `BiDi-CDP Mapper`は生きてるので、再バイディングを行う
                printMsg Debug_, LogID, "The `BiDi-CDP Mapper` instance is already running. Update the SessionID and reconnect to it.", FromProcedureName
                Call BiDiResponseBiding

                '7. 既存接続したサインで返す
                ReBootBiDiCDPMapper = False
            Else
                '6. お亡くなりなので、`BiDi-CDP Mapper`を再始動させます
                printMsg Debug_, LogID, "`BiDi-CDP Mapper` instance not found. ReBooting `BiDi-CDP Mapper` into reattached tab...", FromProcedureName
                StartBiDiCDPMapper

                '7. 再始動したサインを返す
                ReBootBiDiCDPMapper = True
            End If

        Else
            '5. 別URL状態なので、空ページに遷移させる
            '※`Target.exposeDevToolsProtocol`が有効の状態でのネットサーフィンは、危険なのでこの措置をとる
            '　超端的に言うと、ネット上のvbsコードがIEで実行されちゃって大変なことになるあれ
            printMsg WARN_, LogID, "The URL is not suitable for `Target.exposeDevToolsProtocol`. You will be redirected to an empty page.", FromProcedureName
            WebDriverBiDi_mapperTab.Url = MapperURL

            '6. `BiDi-CDP Mapper`を始動させる
            StartBiDiCDPMapper

            '7. 再始動したサインを返す
            ReBootBiDiCDPMapper = True
        End If
    Else
        '2. 既知タブ自体が失ってるので、新規タブで準備し直す
        printMsg WARN_, LogID, "Mapper instance not found. Injecting `BiDi-CDP Mapper` into reattached new tab...", FromProcedureName
        newTabForBiDi

        '3. `BiDi-CDP Mapper`を始動させます
        StartBiDiCDPMapper

        '4. 再始動したサインを返す
        ReBootBiDiCDPMapper = True
    End If
End Function



'***************************************************************************************************
'                           ■■■ WebDriverBiDi との読み書き ■■■
'***************************************************************************************************
'* 機能　　：このタブに対する非同期イベントを受け取ります
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName 非同期イベントメソッド名
'            RawJson    非同期イベントの生Json文字列　※As Is. メソッド名付きのままで渡されます
'            sessionID  届いたSessionID
'---------------------------------------------------------------------------------------------------
'* 機能説明：`Runtime.bindingCalled`に含まれる`payload`を上手い具合に抜き出す処理です
'***************************************************************************************************
Private Sub m_CDPCore_CDPContextEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".m_CDPCore_CDPContextEvent"

    'BiDiでのやりとりの核となるイベント名
    Const BiDiForCDPEventName As String = "Runtime.bindingCalled"


    '1. このタブクラスで設定中の`SessionId`以外は即抜け
    If sessionID <> WebDriverBiDi_mapperTab.CurrentSessionID Then Exit Sub

    '2. 対象のイベント以外も即抜け
    If methodName <> BiDiForCDPEventName Then Exit Sub

    '3. PayloadのName部分を取得
    Dim CDPToBiDi As BiDiCDPJson: Set CDPToBiDi = BiDiCDPJson.Parse(RawJson).NodeKey("params")
    Dim bindingCalledName As String: bindingCalledName = CDPToBiDi.StringKey("name")

    '4. `Name`値に応じた処理
    Dim bindingCalledPayload As String
    Select Case bindingCalledName
        'メインとなるWebDriverBiDiの結果
        Case BiDiResMain
            '4-1. payload 部分を抽出 ※ライブラリ側で自動的にエスケープ解除してくれます
            bindingCalledPayload = CDPToBiDi.StringKey("payload")
            If BiDiDebug.seeRawSendMsgDbg Then printMsg Debug_, LogID, "Result - WebDriverBiDi: " & bindingCalledPayload, FromProcedureName

            '4-2. イベント発信用プロシージャへ遷移
            BiDiReceivedDataCheck bindingCalledPayload

        'BiDi-Mapper 内部のデバッグログ用
        Case BiDiResDebug
            '4-1. エスケープ解除済みのpayload 部分を抽出し、ログ表示のみに留める
            bindingCalledPayload = CDPToBiDi.StringKey("payload")
            printMsg Debug_, LogID, "Debug  - WebDriverBiDi: " & bindingCalledPayload, FromProcedureName

        'WebDriverBiDi 以外の`Runtime.bindingCalled`が返った模様
        '※この場合は、追加処理しない
        Case Else
            printMsg Debug_, LogID, "We detected `" & BiDiForCDPEventName & "` instances other than WebDriverBiDi. Name :" & bindingCalledName, FromProcedureName
    End Select
End Sub

'***************************************************************************************************
'* 機能　　：BiDiコマンドが載ったJavaScriptをCDPで非同期実行します
'---------------------------------------------------------------------------------------------------
'* 引数　　：strData        完成したBiDi-Jsonコマンド
'            StopPipeError  CDP側の`TakeEvents`によるPipeエラーを無視するかの設定
'***************************************************************************************************
Private Sub RunCDP(strData As String, StopPipeError As Boolean)
    '1. JSONのダブルクォートやエスケープ処理による構文乱れを防ぐための引数設定
    Dim argObj As New Dictionary: argObj.Add "value", strData

    '2. BiDiパラメータが載ったJavaScript実行で、非同期CDP実行
    '※CDPと同様のロジック、「書き込みだけ」にするため、非同期で実行。同期だと、パイプ内のデータ回収が走って、IDずれが起こるため
    '※非同期実行であれば、`.jsEval`でも十分な軽量コーディングなので問題なし
    WebDriverBiDi_mapperTab.jsEval "function(commandBiDi) { window." & BiDiSend & "(commandBiDi); }", objectArguments:=Array(argObj), RunAsyncCDP:=True, StopPipeError:=StopPipeError
End Sub



'***************************************************************************************************
'                              ■■■ BiDiコマンド実行前の下準備 ■■■
'***************************************************************************************************
'* 機能　　：CDP実行前の組み立てを行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：strMessage     BiDi-Jsonコマンド。最上位要素に、`method`、`params`があるようにしてください。
'            StopPipeError  CDP側の`TakeEvents`によるPipeエラーを無視するかの設定
'---------------------------------------------------------------------------------------------------
'* 機能説明：`id`の組み立て処理を行います。
'***************************************************************************************************
Public Sub ReadyRunBiDi(dicMessage As Scripting.Dictionary, StopPipeError As Boolean)
    Const FromProcedureName As String = ThisClassName & ".ReadyRunBiDi"


    '1. 現在のコマンドIDを発行  ※内部でカウントアップを行います
    Dim currentCommandID As Long: currentCommandID = IssuanceCommandID

    '2. `id`を付与
    dicMessage.Add "id", currentCommandID

    '3. 実際に実行
    Dim strMessage As String
'    strMessage = BiDiCDPJson.StringifyValue(dicMessage)     '理想はこっちにして、1つに完結したいﾋﾟｴﾝ
    strMessage = WebJsonConverter.serialize(dicMessage)     'Dictionary→Json文字列化のみ`VBA-FastJSON`を使用。こっちのほうがパフォーマンスが良いため
    printMsg Debug_, LogID, "sendBiDi - JsonCommand = " & strMessage, FromProcedureName
    RunCDP strMessage, StopPipeError
End Sub

'***************************************************************************************************
'* 機能　　：BiDi実行の際のコマンドIDを発行します
'---------------------------------------------------------------------------------------------------
'* 返り値　：発行されたcommandID
'---------------------------------------------------------------------------------------------------
'* 注意事項：20億を超えると、1にリセットされます
'***************************************************************************************************
Private Function IssuanceCommandID() As Long
    Const FromProcedureName As String = ThisClassName & ".IssuanceCommandID"


    If m_MessageID + 1 > LimitCommandID Then
        '1. リクエスト数の上限に達したら、1にリセットしておく
        m_MessageID = 1
        printMsg info_, LogID, "Reset to `1` to prevent internal counter overflow.", FromProcedureName

        '2. リセットした旨を伝える
        RaiseEvent ResetCommandID
    Else
        '1. カウントアップ
        m_MessageID = m_MessageID + 1   'We increase the global ID counter
    End If

    '2. 返却
    IssuanceCommandID = m_MessageID
End Function



'***************************************************************************************************
'                               ■■■ レスポンス分類 ■■■
'***************************************************************************************************
'* 機能　　：BiDiコマンドの受信データを高速にチェックし、仕分けます
'---------------------------------------------------------------------------------------------------
'* 引数　　：BiDiJsonString         1件分のBiDiの生JSON文字列
'---------------------------------------------------------------------------------------------------
'* 詳細説明：イベントデータ or CDP-Jsonコマンド結果 or 構文エラー の確認をし、各種`RaiseEvent`で発信します
'* 注意事項；構文エラーの場合、不整合により不安定につながるため、エラーで停止します
'***************************************************************************************************
Private Function BiDiReceivedDataCheck(BiDiJsonString As String)
    Const FromProcedureName As String = ThisClassName & ".BiDiReceivedDataCheck"


    Dim ResultCommandID As Long, AsynchronousEventName As String
    Dim BiDiNode As BiDiCDPJson: Set BiDiNode = BiDiCDPJson.Parse(BiDiJsonString)

    '--------- 1. BiDiからのイベントか？ ---------
    If BiDiNode.StringKey("type") = "event" Then
        '非同期イベントとして発信
        AsynchronousEventName = BiDiNode.StringKey("method")
        RaiseEvent detectionBiDiEvent(AsynchronousEventName, BiDiJsonString)
    Else
        '--------- 2. `id`項目があるか？ ---------
        If BiDiNode.ExistsKey("id") Then
            'コマンド結果として発信
            ResultCommandID = BiDiNode.NumberKey("id")
            RaiseEvent detectionBiDiID(ResultCommandID, BiDiJsonString)
        Else
            '`method`も`id`項目もない場合、BiDi-Jsonコマンドの構文エラーとして返った可能性があるとし、受信ループ終了とする
            printMsg ERROR_, LogID, "A syntax error was detected when executing with ExecuteBiDi/ExecuteBiDiAsync. Processing was interrupted to prevent inconsistencies." & vbCrLf & "> " & BiDiJsonString, FromProcedureName, doRaiseError:=True
        End If
    End If
End Function



'***************************************************************************************************
'                              ■■■ 複製用プロパティ ■■■
'***************************************************************************************************
'* 機能　　：このClass内で定義してる`CDPBrowser.cls`を受け取り、WebDriverBiDiモードでありつつ、直接CDP制御できる仕組みを提供します
'---------------------------------------------------------------------------------------------------
'* 注意事項：利用する際は、「BiDi-CDP Mapper」タブ自体を操作しないように気を付ける必要がある上級者向けです
'***************************************************************************************************
Property Get DirectControlCDP() As CDPBrowser
    Set DirectControlCDP = m_CDPBrowser
End Property



'***************************************************************************************************
'                               ■■■ プロパティ機能 ■■■
'***************************************************************************************************
'* 機能　　：最後の○○系取得用プロパティ
'***************************************************************************************************
Property Get LastErrorPeekNamedPipe() As Long           '`PeekNamedPipe`使用時のエラーコード
    LastErrorPeekNamedPipe = m_CDPCore.LastErrorPeekNamedPipe
End Property

Property Get LastErrorReadFile() As Long                '`ReadFile`使用時のエラーコード
    LastErrorReadFile = m_CDPCore.LastErrorReadFile
End Property

Property Get LastErrorWriteFile() As Long               '`WriteFile`使用時のエラーコード
    LastErrorWriteFile = m_CDPCore.LastErrorWriteFile
End Property

Property Get LastCommandID() As Long                    '最後に実行したcommandID値
    LastCommandID = m_MessageID
End Property

'***************************************************************************************************
'* 機能　　：ログ設定値、取得用プロパティ
'***************************************************************************************************
Property Get ShowLogBiDiresponse() As Boolean
    ShowLogBiDiresponse = m_CDPCore.ShowLogCDPresponse
End Property



'***************************************************************************************************
'                            ■■■ CDPCore.cls 直通メソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, LogID As String, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    m_CDPCore.printMsg LogLevel_, LogID, strMsg, From, isHeader, doRaiseError
End Sub

Public Sub sleep(Optional seconds As Double = 0.5)
    m_CDPCore.sleep seconds
End Sub

Public Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    m_CDPCore.TakeEvents StopPipeError, destruction
End Sub

Public Function TimerCounter() As Double
    TimerCounter = m_CDPCore.GetTicksMs
End Function

Public Sub CloseHandleBrowser()
    m_CDPCore.CloseHandleBrowser
End Sub



'***************************************************************************************************
'                               ■■■ `mapperTab.js`更新関連 ■■■
'***************************************************************************************************
'* 機能　　：ローカルファイルから、JSコードを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功論理値
'* 引数　　：FolderPath     フォルダパス
'            FileName       JSファイル名
'---------------------------------------------------------------------------------------------------
'* 機能説明：読み取ったファイルと、Excelテーブルに挿入した文字列との一致確認も行っていきます
'***************************************************************************************************
Public Function UpdateFromLocalFile(FolderPath As String, Optional FileName As String = MapperJSName) As Boolean
    Const FromProcedureName As String = ThisClassName & ".UpdateFromLocalFile"


    '1. ファイルから、Jsコードを取得
    '※Windows方式の改行コード(CRLF)だと、テーブル挿入時の分割にて、`CR`と`LF`での分割が起こって不整合になる可能性があるため、Linux方式(LF)に統一させます
    Dim m_CharConv As New CharacterCodeConversion
    Dim JsCode As String
    With m_CharConv
        JsCode = Replace(.BytesToString(.BytesFromSavedFile(FolderPath, FileName)), vbCrLf, vbLf)
    End With

    With ShLibrary01_JS
        '2. 専用のテーブルに配置
        .TableMapperTabJS = JsCode

        '3. 正しく置けてるか確認
        UpdateFromLocalFile = (JsCode = .TableMapperTabJS)

        '4. ミスの場合はテーブルの中身を削除
        If Not (UpdateFromLocalFile) Then .DeleteMapperTabJS

        'ローカルファイルで更新した旨を記録
        '※ミスの場合は、空欄にする
        .VersionMapperTabJS = IIf(UpdateFromLocalFile, " Local", "")
    End With
End Function

'***************************************************************************************************
'* 機能　　：インターネット経由で、JSコードを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功論理値
'* 引数　　：Version        バージョン指定する場合は、指定してください。デフォルトは最新です
'---------------------------------------------------------------------------------------------------
'* 機能説明：`https://www.jsdelivr.com/`というオープンソースファイル向けの無料のCDN(コンテンツ配信ネットワーク)を利用して、
'            `chromium-bidi`の`mapperTab.js`を更新させます
'
'* 注意事項：chromium-bidi v17 以降は、ビルドシステムが GN (Generate Ninja) に移行しました。
'
'            その影響で､npm に公開されるファイル構成が大きく変更され､
'            従来の「lib/iife/mapperTab.js」ではなく、
'            GN が生成するビルド成果物のディレクトリ
'                out/Default/gen/src/
'            配下に mapperTab.js が配置されるようになりました。
'
'            ＜背景＞
'            ・v16 までは rollup ベースの IIFE ビルドが npm にそのまま含まれていた
'            ・v17 で GN ベースのビルドに統合され、出力ディレクトリが再編成された
'            ・そのため jsDelivr 経由で取得するパスも変更が必要になった
'
'            ＜注意＞
'            ・v16 以前と v17 以降で mapperTab.js の場所が異なるため、
'            　必要に応じて URL を切り替えてください。
'            ・将来のバージョンでさらに構成が変わる可能性があります。
'            　（GN 移行期のため、ディレクトリ構造が安定していない可能性あり）
'
'            ＜旧構成（v16 以前）＞
'               npm/chromium-bidi@{Version}/lib/iife/mapperTab.js
'
'            ＜新構成（v17 以降）＞
'               npm/chromium-bidi@{Version}/out/Default/gen/src/mapperTab.js
'
'            この変更はライブラリの不具合ではなく､公式ビルドシステムの移行による仕様変更です｡
'***************************************************************************************************
Public Function UpdateFromNPMFile(Optional Version As String = "latest") As Boolean
    'リクエストパラメーターを設定する
    Dim リクエスト情報一式 As WebRequest: Set リクエスト情報一式 = New WebRequest
    With リクエスト情報一式
        'GET
        .Method = vbWebHttpGet

        'レスポンスのデータ形式を指定
        .ResponseFormat = vbWebFormatPlainText

        '`chromium-bidi`のnpmパスパラメーターを指定(ホスト名に続くURL)
'        .Resource = "npm/" & npmName & "@{Version}/lib/iife/" & MapperJSName   '※v16.0.1以下を使いたい場合はこちらをご利用ください。文字列同士では数値の大小の比較ができないため、ユーザー側の明示的な切り替えが必要です。
        .Resource = "npm/" & npmName & "@{Version}/out/Default/gen/src/" & MapperJSName
        .AddUrlSegment "Version", Version
    End With

    'リクエストして、結果を返す
    With jsDelivrへリクエスト(リクエスト情報一式, True)
        '200確認
        If .StatusCode = vbWebStatusOk Then
            'テーブルに置く作業
            ShLibrary01_JS.TableMapperTabJS = .Content

            '正しく置けてるか確認
            UpdateFromNPMFile = (.Content = ShLibrary01_JS.TableMapperTabJS)

            'ミスの場合はテーブルの中身を削除
            If Not (UpdateFromNPMFile) Then ShLibrary01_JS.DeleteMapperTabJS

        Else
            MsgBox "npm経由による`" & MapperJSName & "`の取得に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & .Content, vbCritical, "ErrorCode:" & .StatusCode
            UpdateFromNPMFile = False
        End If
    End With
End Function

'***************************************************************************************************
'* 機能　　：インターネット経由で、現在のJSコードバージョンを確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：現在の最新バージョン
'---------------------------------------------------------------------------------------------------
'* 機能説明：`https://www.jsdelivr.com/`というオープンソースファイル向けの無料のCDN(コンテンツ配信ネットワーク)を利用して、
'            `chromium-bidi`の`mapperTab.js`を更新させます
'***************************************************************************************************
Public Function UpdateCheckNPMVersion() As String
    Dim ResultResponse As WebResponse

    'リクエストパラメーターを設定する
    Dim リクエスト情報一式 As WebRequest: Set リクエスト情報一式 = New WebRequest
    With リクエスト情報一式
        'GET
        .Method = vbWebHttpGet

        'レスポンスのデータ形式を指定
        .ResponseFormat = vbWebFormatJson

        '`chromium-bidi`のnpmパスパラメーターを指定(ホスト名に続くURL)
        .Resource = "v1/package/npm/" & npmName
    End With

    'リクエストして、結果を返す
    With jsDelivrへリクエスト(リクエスト情報一式, False)
        '200確認
        If .StatusCode = vbWebStatusOk Then
            UpdateCheckNPMVersion = .Data("tags")("latest")
        Else
            MsgBox "npm経由による`" & MapperJSName & "`のバージョン確認に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & .Content, vbCritical, "ErrorCode:" & .StatusCode
        End If
    End With
End Function

'***************************************************************************************************
'* 機能　　：jsDelivr関連のホスト設定です
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスオブジェクト - WebResponse   ここに結果が格納されます
'* 引数　　：クラスオブジェクト - WebRequest    事前に設定したリクエストパラメーター一式
'            GetContents                        `True` :ファイルそのもの用のホストととして
'                                               `False`:バージョン確認用のホストととして
'***************************************************************************************************
Private Function jsDelivrへリクエスト(リクエスト情報一式 As WebRequest, GetContents As Boolean) As WebResponse
    'エンドポイント通信の設定
    Dim EndPointAPI As New WebClient
    With EndPointAPI
        'ホスト名URL
        '実際のファイルか、バージョン確認かの切り替えをします
        .BaseUrl = IIf(GetContents, "https://cdn.jsdelivr.net/", "https://data.jsdelivr.com/")
    End With

    'リクエスト設定とホスト設定を使ってリクエストし、結果を返します
    Set jsDelivrへリクエスト = EndPointAPI.Execute(リクエスト情報一式)
End Function



'***************************************************************************************************
'                               ■■■ 初期化処理 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'Assign an objID for easy debug identification
    LogID = "BCID" & Format(Rnd * 1000, "000")

    'WebDriverBiDiのデバッグ設定を、ワークシートから適用します
    With ShSetting01_StartBrowser
        BiDiDebug.ViewDebugTab = .UseRangeID(11, FromProcedureName)    '`BiDi-CDP Mapper`のコンソールタブを表示するか？
        BiDiDebug.seeRawSendMsgDbg = .UseRangeID(8, FromProcedureName) 'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
        BiDiDebug.DiagnosticLog = .UseRangeID(13, FromProcedureName)   'Mapper内部のデバッグログを記録するか？
    End With
End Sub



'***************************************************************************************************
'                              ■■■ HISTORICAL ARCHIVE ■■■
'***************************************************************************************************
' かつて、VBAのJSONパース制限と戦うため、極限の位置計算によって
' 巨大データから必要項目を1ミリ秒で撃ち抜いていた伝説のロマンコード。
' ゼロコピーパーサー(vbacollective/json)の導入に伴い、ここに栄誉を称え封印する。
' 2026/06/28 退役
'
' 消すのがなんかもったいないので、残しておきます。※`Github`に残るだろ、というクレームは受け付けません
'***************************************************************************************************
'* 機能　　：`Runtime.bindingCalled` のCDPメッセージから name 部分を抜き出します
'---------------------------------------------------------------------------------------------------
'* 返値　　：`name`部分
'* 引数　　：resCDP     CDPからの非同期イベント:`Runtime.bindingCalled`の内容そのまんま
'***************************************************************************************************
'Private Function ExtractName(resCDP As String) As String
'    Const FromProcedureName As String = ThisClassName & ".ExtractName"
'
'
'    '検索設定
'    Const SearchKeyName     As String = "{""method"":""Runtime.bindingCalled"",""params"":{""name"":""" '検索文字列
'    Const StartSearchPos    As Long = 1                                                                 'あるべき位置
'
'
'    '1. 文字列を探索
'    Dim StartPos As Long
'    StartPos = InStr(StartSearchPos, resCDP, SearchKeyName)
'
'    '2. 所定の位置にない場合は、存在しないとし、抜ける
'    '例：{"method":"Runtime.bindingCalled","params":{"name":"sendBidiResponse","...
'    If StartPos = StartSearchPos Then
'
'        '3. 開始位置をうまい具合にずらす
'        StartPos = StartSearchPos + Len(SearchKeyName)
'
'        '4. name値の終点を求める
'        Dim EndPos As Long
'        EndPos = InStr(StartPos, resCDP, """,""")
'
'        '5. name値を取り出す
'        ExtractName = Mid(resCDP, StartPos, EndPos - StartPos)
'        printMsg Trace_, LogID, "`Name` of `Runtime.bindingCalled` > " & ExtractName, FromProcedureName
'    Else
'        'name値が特定できない場合、`Runtime.bindingCalled`の構造が変わってる可能性があり、続行不可とする
'        printMsg ERROR_, LogID, "Failed to identify the `name` value. The structure of `Runtime.bindingCalled` may have changed." & vbCrLf & "> " & resCDP, FromProcedureName, doRaiseError:=True
'    End If
'End Function
'
''***************************************************************************************************
''* 機能　　：`Runtime.bindingCalled` のCDPメッセージから payload 部分を高速に抜き出します
''---------------------------------------------------------------------------------------------------
''* 返値　　：エスケープ解除済み`payload`部分
''* 引数　　：resCDP             CDPからの非同期イベント:`Runtime.bindingCalled`の内容そのまんま
''            bindingCalledName  `Runtime.bindingCalled`の`Name`値を指定。抽出精度を高めるために使用します
''***************************************************************************************************
'Private Function ExtractPayload(resCDP As String, bindingCalledName As String) As String
'    Const FromProcedureName As String = ThisClassName & ".ExtractPayload"
'
'
'    '検索設定
'    Const SearchKeyName As String = """,""payload"":"""                                             '実際の検索文字列
'    Const FirstLetter   As String = "{""method"":""Runtime.bindingCalled"",""params"":{""name"":""" '先頭の固定文字列
'    Dim StartSearchPos As Long: StartSearchPos = Len(FirstLetter) + Len(bindingCalledName) + 1      '`payload`があるべき位置を動的に特定
'
'
'    '1. 文字列を探索
'    Dim StartPos As Long
'    StartPos = InStr(StartSearchPos, resCDP, SearchKeyName)
'
'    '2. 所定の位置あるか確認する
'    If StartPos = StartSearchPos Then
'        '3. 開始位置をうまい具合にずらす
'        StartPos = StartSearchPos + Len(SearchKeyName)
'
'        '4. payloadの終了位置を探す。通常はexecutionContextIdの前
'        'https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#event-bindingCalled
'        Dim EndPos As Long: EndPos = InStrRev(resCDP, """,""executionContextId""")
'
'        '5. 取り出せる計算かチェック
'        If EndPos > StartPos Then
'            '6.payload値を取り出す
'            ExtractPayload = Mid(resCDP, StartPos, EndPos - StartPos)
'
'            '7.「JsonOfJson」の外殻だけ剥がすための、エスケープ（\" -> "）を復元
'            'これにより、`WebJsonConverter`クラスのパース処理に渡せれる形になります
'            ExtractPayload = Replace(ExtractPayload, "\""", """")
'            ExtractPayload = Replace(ExtractPayload, "\\", "\")
'        Else
'            'payload値が取れない場合、`Runtime.bindingCalled`の構造が変わってる可能性があり、続行不可とする
'            printMsg ERROR_, LogID, "Failed to retrieve the `payload` value. The structure of `Runtime.bindingCalled` may have changed." & vbCrLf & "> " & resCDP, FromProcedureName, doRaiseError:=True
'        End If
'    Else
'        'payload値が特定できない場合、`Runtime.bindingCalled`の構造が変わってる可能性があり、続行不可とする
'        printMsg ERROR_, LogID, "Failed to identify the `payload` value. The structure of `Runtime.bindingCalled` may have changed." & vbCrLf & "> " & resCDP, FromProcedureName, doRaiseError:=True
'    End If
'End Function
''***************************************************************************************************
''* 機能　　：BiDiからの受信データの構造をチェックする際の1つ、`id`チェックを行います
''---------------------------------------------------------------------------------------------------
''* 返り値  ：存在有無論理値
''* 引数　　：jsonString         BiDiのレスポンス(JSON文字列そのまま)
''*            idValueContents   コマンド結果のID値を格納する変数(参照渡し)
''---------------------------------------------------------------------------------------------------
''* 注意事項：コマンド結果のエラー、成功の区別はしません。
''***************************************************************************************************
'Private Function GetContentsID(jsonString As String, ByRef idValueContents As Long) As Boolean
'    Const FromProcedureName As String = ThisClassName & ".GetContentsID"
'
'
'    '検索設定
'    Const SearchKeyNameOK   As String = "success"       '検索文字列(コマンド成功時)
'    Const SearchKeyNameNG   As String = "error"         '検索文字列(コマンド失敗時)
'    Const SearchKeyid       As String = """,""id"":"    '検索文字列の末尾のID部分
'    Const StartSearchPos    As Long = 10                'あるべき位置
'
'
'    '1. 文字列を探索
'    Dim existenceSuccess As Boolean, existenceError As Boolean
'    existenceSuccess = (InStr(StartSearchPos, jsonString, SearchKeyNameOK & SearchKeyid) = StartSearchPos)
'    existenceError = (InStr(StartSearchPos, jsonString, SearchKeyNameNG & SearchKeyid) = StartSearchPos)
'
'    '2. どっちも所定の位置にない場合は、存在しないとし、抜ける
'    '下記の例に基づき、10文字目にない場合は、コマンド結果ではないとして、扱います。
'    '・{"type":"success", "id":... （コマンドの成功）
'    '・{"type":"error", "id":... （コマンドの失敗）
'    '・{"type":"error","error":...(構文エラー)　※これは通さない
'    If existenceSuccess Or existenceError Then
'        GetContentsID = True
'
'        '3. 検索ワードを適用
'        Dim SearchKeyName As String
'        SearchKeyName = IIf(existenceSuccess, SearchKeyNameOK, SearchKeyNameNG)
'
'        '4. 開始位置をうまい具合にずらす
'        Dim StartPos As Long
'        StartPos = StartSearchPos + Len(SearchKeyName) + Len(SearchKeyid) '「","id":」分、ずらすイメージ
'
'        '5. ID値の終点を求める
'        Dim EndPos As Long
'        EndPos = InStr(StartPos, jsonString, ",")
'
'        '6. ID値を取り出す
'        idValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
'        printMsg Trace_, LogID, "This BiDi response is the BiDicommand result. ID:" & idValueContents, FromProcedureName
'    Else
'        GetContentsID = False
'        printMsg Trace_, LogID, "This BiDi response does not have an `id`.", FromProcedureName
'    End If
'End Function
'
''***************************************************************************************************
''* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`method`チェックを行います
''---------------------------------------------------------------------------------------------------
''* 返り値  ：存在有無論理値
''* 引数　　：jsonString             BiDiのレスポンス(JSON文字列そのまま)
''*            methodValueContents  非同期イベントのメソッド名(参照渡し)
''---------------------------------------------------------------------------------------------------
''* 詳細説明：前方一致、後方一致を確認します
''***************************************************************************************************
'Private Function GetContentsMethod(jsonString As String, ByRef methodValueContents As String) As Boolean
'    Const FromProcedureName As String = ThisClassName & ".GetContentsMethod"
'
'
'    '検索設定
'    Const SearchKeyName     As String = "event" '検索文字列
'    Const StartSearchPos    As Long = 10        'あるべき位置
'
'
'    '1. 前方文字列を探索
'    Dim StartPos As Long, EndPos As Long
'    StartPos = InStr(StartSearchPos, jsonString, SearchKeyName)
'
'    '2. 所定の位置あるか確認する
'    '下記の例に基づき、10文字目にない場合は、次に後方一致を確認
'    '・{"type":"event","method":...
'    If StartPos = StartSearchPos Then
'        GetContentsMethod = True
'
'        '3. 開始位置をうまい具合にずらす
'        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""",""method"":""") '「","method":"」分、ずらすイメージ
'
'        '4. method名の終点を求める
'        EndPos = InStr(StartPos, jsonString, """,")
'
'        '5. method名を取り出す
'        methodValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
'        printMsg Trace_, LogID, "This BiDi response is an asynchronous event: " & methodValueContents, FromProcedureName
'
'    '下記の例に基づき、後方にない場合は、非同期イベントではないとして、扱います。
'    '・{"method":"network.beforeRequestSent", ... (中略) ..., "type":"event"}
'    ElseIf Right(jsonString, Len(SearchKeyName) + Len("""}")) = SearchKeyName & """}" Then
'        GetContentsMethod = True
'
'        '3. 開始位置をうまい具合にずらす
'        StartPos = 1 + Len("{""method"":""") '「{"method":"」分、ずらすイメージ
'
'        '4. method名の終点を求める
'        EndPos = InStr(StartPos, jsonString, """,")
'
'        '5. method名を取り出す
'        methodValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
'        printMsg Trace_, LogID, "This BiDi response is an asynchronous event: " & methodValueContents, FromProcedureName
'    Else
'        GetContentsMethod = False
'        printMsg Trace_, LogID, "This BiDi response does not have an `method`.", FromProcedureName
'    End If
'End Function
