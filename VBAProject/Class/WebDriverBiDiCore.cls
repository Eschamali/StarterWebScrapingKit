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
Attribute m_CDPBrowser.VB_VarHelpID = -1
Private WebDriverBiDi_mapperTab As CDPContext   'BiDi-CDP Mapper 操作用タブ

'定義側からも呼び出せるチェーンクラスオブジェクト
Public ErrorDetail      As New WinApiError      'エラーコードから、詳細を取得するやつ

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
Public StopPipeError    As Boolean  'Pipe関連のエラーがあったら、`Err.Raise`で止めるか？



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：「BiDi-CDP Mapper」を付けた状態で、ブラウザに対する所定の起動引数をつけて起動させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：Name                           使用ブラウザ名※現時点では、Edgeとchromeのみです
'            appUrl                         最初の起動時に遷移したいURL
'            cleanActive                    `True` :デバッグブラウザを(あれば)Kill後、ブラウザを起動する
'                                           `False`:起動チェックをせずに、ブラウザを起動します
'
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
Public Sub start(Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional cleanActive As Boolean = True, _
    Optional userProfile As String, Optional addArgs As String, Optional KioskMode As edgeKioskType, _
    Optional sessionCapabilitiesRequest As Dictionary)

    Const FromProcedureName As String = ThisClassName & ".start"


    '1. 引数から設定を適用し、まずはCDPとして起動
    Set m_CDPBrowser = New CDPBrowser
    m_CDPBrowser.start Name, appUrl, cleanActive, userProfile, addArgs, KioskMode

    '2. WithEventsClassオブジェクトを受け取る
    Set m_CDPCore = m_CDPBrowser.InheritanceCDPCore

    '3. BiDi-CDP Mapper を始動させます
    StartBiDiCDPMapper
End Sub

'***************************************************************************************************
'* 機能　　：新規タブにて、`BiDi-CDP Mapper`を始動させます
'* 引数　　：ReConnect  Excelテーブルにある`targetID`に対して再度、`BiDi-CDP Mapper`を再始動する際は、`True`にします
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`CDPBrowser`を`New`してからご利用ください
'            ・各パラメータは、`chromium-bidi\src\bidiServer\MapperCdpConnection.ts`に準拠してます
'***************************************************************************************************
Private Sub StartBiDiCDPMapper(Optional ReConnect As Boolean)
    Const FromProcedureName As String = ThisClassName & ".StartBiDiCDPMapper"


    'ブラウザ単位実行
    With m_CDPBrowser
        '1. 新規タブに対して、BiDiコンソールタブを準備し、セッションターゲットをメイン化
        '※タブ表示モードの際は、そのタブがフォアグラウンドにならないように仕向けます
        If Not (ReConnect) Then Set WebDriverBiDi_mapperTab = .newTab(MapperURL, setMain:=True, isHidden:=Not (BiDiDebug.ViewDebugTab), isBackground:=True)
        printMsg Debug_, LogID, "Booting `BiDi-CDP Mapper`... (ViewTab: " & CStr(BiDiDebug.ViewDebugTab) & ")", FromProcedureName

        '2. BiDi用バインディングの追加
        Dim paramsCDP As New Dictionary
        paramsCDP.Add "bindingName", "cdp"
        paramsCDP.Add "targetId", WebDriverBiDi_mapperTab.CurrentTargetID
        paramsCDP.Add "inheritPermissions", True
        .ExecuteCDP "Target.exposeDevToolsProtocol", paramsCDP, Not ReConnect   '再接続モードの場合は、エラー無視
    End With

    '3. JSからVBAへ返す経路を設定
    Call BiDiResponseBiding

    'タブ単位実行
    With WebDriverBiDi_mapperTab
        '4. パースしたJavaScriptをブラウザ上で実行
        .jsEval Take_mapperTabJSstring

        '5. ブートストラップ（初期化）JSの実行
        .jsEval "window.runMapperInstance('" & WebDriverBiDi_mapperTab.CurrentTargetID & "')"
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
'* 返値    ：再始動したか論理値 `True` ：`BiDi-CDP Mapper`を再始動しました。`WebDriverBiDiMode.cls`側で、`session.new`する必要があります
'                               `False`：既存の`BiDi-CDP Mapper`を流用できたので、`session.new`は不要です
'---------------------------------------------------------------------------------------------------
'* 機能説明：あとから、`BiDi-CDP Mapper`セッションに再接続する際に使い、セッション確立まで担います
'* 注意事項：・以前のタブが失ってもブラウザハンドル情報が生きてる場合は、新しいタブとして復旧し、再開します
'            ・ブラウザ自体が落ちてると、VBAエラーで止まります
'***************************************************************************************************
Public Function ReBootBiDiCDPMapper(userProfile As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".ReBootBiDiCDPMapper"


    '1. CDPタブセッションに再接続
    Set WebDriverBiDi_mapperTab = New CDPContext

    '2. Excelに記録されてる`TargetID`の生存確認　※`sessionId`は更新させます
    Dim mapperTabConnection As Boolean
    mapperTabConnection = WebDriverBiDi_mapperTab.reattach(userProfile)

    '3. 各種オブジェクトを継承
    Set m_CDPCore = WebDriverBiDi_mapperTab.InheritanceCDPCore
    Set m_CDPBrowser = WebDriverBiDi_mapperTab.InheritanceCDPBrowser

    '`TargetID`有無に応じた分岐
    If mapperTabConnection Then
        'Excelに記録されてる`TargetID`が生きてる = `Target.exposeDevToolsProtocol`が有効
        printMsg info_, LogID, "`Target.exposeDevToolsProtocol` instance is already running. Reattaching to existing BiDi session...", FromProcedureName

        '4. そのタブURLは、このクラスで用意した専用URLか？
        If LCase(MapperURL) = LCase(WebDriverBiDi_mapperTab.Url) Then
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
                StartBiDiCDPMapper True

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
            StartBiDiCDPMapper True

            '7. 再始動したサインを返す
            ReBootBiDiCDPMapper = True
        End If
    Else
        '2. 既知タブ自体が失ってるので、新規タブで始動しなおす
        printMsg WARN_, LogID, "Mapper instance not found. Injecting `BiDi-CDP Mapper` into reattached new tab...", FromProcedureName
        StartBiDiCDPMapper

        '3. 再始動したサインを返す
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
    If sessionID <> WebDriverBiDi_mapperTab.id Then Exit Sub

    '2. 対象のイベント以外も即抜け
    If methodName <> BiDiForCDPEventName Then Exit Sub

    '3. PayloadのName部分を取得
    Dim NamePayload As String: NamePayload = ExtractName(RawJson)

    '4. `Name`値に応じた処理
    Dim tmp As String
    Select Case NamePayload
        'メインとなるWebDriverBiDiの結果
        Case BiDiResMain
            '4-1. エスケープ解除済みのpayload 部分を抽出
            tmp = ExtractPayload(RawJson)
            If BiDiDebug.seeRawSendMsgDbg Then printMsg Debug_, LogID, "Result - WebDriverBiDi: " & tmp, FromProcedureName

            '4-2. イベント発信用プロシージャへ遷移
            BiDiReceivedDataCheck tmp

        'BiDi-Mapper 内部のデバッグログ用
        Case BiDiResDebug
            '4-1. エスケープ解除済みのpayload 部分を抽出し、ログ表示のみに留める
            tmp = ExtractPayload(RawJson)
            printMsg Debug_, LogID, "Debug  - WebDriverBiDi: " & tmp, FromProcedureName

        'WebDriverBiDi 以外の`Runtime.bindingCalled`が返った模様
        '※この場合は、追加処理しない
        Case Else
            printMsg Debug_, LogID, "We detected `" & BiDiForCDPEventName & "` instances other than WebDriverBiDi. Name :" & NamePayload, FromProcedureName
    End Select
End Sub

'***************************************************************************************************
'* 機能　　：`Runtime.bindingCalled` のCDPメッセージから name 部分を抜き出します
'---------------------------------------------------------------------------------------------------
'* 返値　　：`name`部分
'* 引数　　：resCDP     CDPからの非同期イベント:`Runtime.bindingCalled`の内容そのまんま
'***************************************************************************************************
Private Function ExtractName(resCDP As String) As String
    Const FromProcedureName As String = ThisClassName & ".ExtractName"

    '--------- params-name項目から、`payload`の種類を特定 ---------
    Const SearchKey_ParamsName As String = "{""method"":""Runtime.bindingCalled"",""params"":{""name"":"""
    Dim StartPos As Long, EndPos As Long
    Dim CalledName As String

    'name項目を抽出
    StartPos = Len(SearchKey_ParamsName) + 1    '「{"method":"Runtime.bindingCalled","params":{"name":"」分、検索位置をずらす想定
    EndPos = InStr(StartPos, resCDP, """,""")
    ExtractName = Mid(resCDP, StartPos, EndPos - StartPos)
End Function

'***************************************************************************************************
'* 機能　　：`Runtime.bindingCalled` のCDPメッセージから payload 部分を高速に抜き出します
'---------------------------------------------------------------------------------------------------
'* 返値　　：エスケープ解除済み`payload`部分
'* 引数　　：resCDP     CDPからの非同期イベント:`Runtime.bindingCalled`の内容そのまんま
'***************************************************************************************************
Private Function ExtractPayload(resCDP As String) As String
    Const SearchKey As String = """payload"":"""
    Dim StartPos As Long: StartPos = InStr(1, resCDP, SearchKey)
    If StartPos = 0 Then Exit Function

    StartPos = StartPos + Len(SearchKey)

    ' payloadの終了位置を探す。通常はexecutionContextIdの前
    Dim EndPos As Long: EndPos = InStr(StartPos, resCDP, """,""executionContextId""")
    If EndPos = 0 Then EndPos = InStrRev(resCDP, """}") ' 万が一の末尾判定

    If EndPos > StartPos Then
        ' JSONエンコードされた文字列をデコード（本来はもっと複雑だが、ここではWebJsonConverter.ParseJsonに任せるために、
        ' いったん「JsonOfJson」の外殻だけ剥がす。ただしVBAのParserを完全に避けるならここが正念場）
        ' 実はParseJsonにこの部分文字列を渡すのが一番安全。
        ExtractPayload = Mid(resCDP, StartPos, EndPos - StartPos)

        ' 文字列内のエスケープ（\" -> "）を復元
        ' ※WebJsonConverter.ParseJson は「JSON全体」を期待するので、
        ' ここで完全に復元してしまうと、今度はその中の文字列が壊れる可能性がある。
        ' なので、ここでは「JSONとしての文字列値」を取得するにとどめ、
        ' クラス内のパース処理にそのまま渡す。
        ' ただし、CDPのJSONエンコードにより \" になっているものを " に戻す必要がある。
        ExtractPayload = Replace(ExtractPayload, "\""", """")
        ExtractPayload = Replace(ExtractPayload, "\\", "\")
    End If
End Function

'***************************************************************************************************
'* 機能　　：実際にCDPを実行します
'---------------------------------------------------------------------------------------------------
'* 引数　　：strData    完成したBiDi-Jsonコマンド
'---------------------------------------------------------------------------------------------------
'* 機能説明：実行前に、エスケープ処理を行い、`Runtime.evaluate`として実行します
'***************************************************************************************************
Private Sub RunCDP(strData As String)
    '1. `Runtime.evaluate`実行前の、JSONのダブルクォートやエスケープ処理による構文乱れを防ぐためエスケープ
    Dim safeJson As String
    safeJson = Replace(strData, "\", "\\")
    safeJson = Replace(safeJson, "'", "\'")

    '2. `expression`を組み立て
    Dim paramsCDP As New Dictionary
    paramsCDP.Add "expression", "window." & BiDiSend & "('" & safeJson & "');"

    '3. CDP実行
    '※JavaScript実行自体は同期処理。エラー時は、内部でStopがかかります。(chromium-bidi-main\src\bidiServer\MapperCdpConnection.ts　に基づく)
    '※また内部にて、`RaiseEvent`よる発信も行われます
    '※`.jsEval`は、内部のお節介によりオーバーヘッドが積み重なることから、直通メソッドで使用します
    WebDriverBiDi_mapperTab.ExecuteCDP "Runtime.evaluate", paramsCDP, StopPipeError
End Sub



'***************************************************************************************************
'                              ■■■ BiDiコマンド実行前の下準備 ■■■
'***************************************************************************************************
'* 機能　　：CDP実行前の組み立てを行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：strMessage BiDi-Jsonコマンド
'---------------------------------------------------------------------------------------------------
'* 機能説明：`id`の組み立て処理を行います。
'***************************************************************************************************
Public Sub ReadyRunBiDi(strMessage As String)
    Const FromProcedureName As String = ThisClassName & ".ReadyRunBiDi"


    '1. 現在のコマンドIDを発行  ※内部でカウントアップを行います
    Dim currentCommandID As Long: currentCommandID = IssuanceCommandID

    '2. 末尾付与しやすいように、末尾1文字削減
    strMessage = Left(strMessage, Len(strMessage) - 1)

    '3. 末尾に`id`を付与
    strMessage = strMessage & ", ""id"":" & currentCommandID & "}"

    '4. 実際に実行
    printMsg Debug_, LogID, "sendBiDi - JsonCommand = " & strMessage, FromProcedureName
    RunCDP strMessage
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

    '--------- 1. BiDiからのイベントか？ ---------
    If GetContentsMethod(BiDiJsonString, AsynchronousEventName) Then
        '非同期イベントとして発信
        RaiseEvent detectionBiDiEvent(AsynchronousEventName, BiDiJsonString)
    Else
        '--------- 2. `id`項目があるか？ ---------
        If GetContentsID(BiDiJsonString, ResultCommandID) Then
            'コマンド結果として発信
            RaiseEvent detectionBiDiID(ResultCommandID, BiDiJsonString)
        Else
            '`method`も`id`項目もない場合、BiDi-Jsonコマンドの構文エラーとして返った可能性があるとし、受信ループ終了とする
            printMsg ERROR_, LogID, "A syntax error was detected when executing with ExecuteBiDi/ExecuteBiDiAsync. Processing was interrupted to prevent inconsistencies." & vbCrLf & "> " & BiDiJsonString, FromProcedureName, doRaiseError:=True
        End If
    End If
End Function

'***************************************************************************************************
'* 機能　　：BiDiからの受信データの構造をチェックする際の1つ、`id`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：jsonString         BiDiのレスポンス(JSON文字列そのまま)
'*            idValueContents   コマンド結果のID値を格納する変数(参照渡し)
'---------------------------------------------------------------------------------------------------
'* 注意事項：コマンド結果のエラー、成功の区別はしません。
'***************************************************************************************************
Private Function GetContentsID(jsonString As String, ByRef idValueContents As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".GetContentsID"


    '検索設定
    Const SearchKeyNameOK   As String = "success"       '検索文字列(コマンド成功時)
    Const SearchKeyNameNG   As String = "error"         '検索文字列(コマンド失敗時)
    Const SearchKeyid       As String = """,""id"":"    '検索文字列の末尾のID部分
    Const StartSearchPos    As Long = 10                'あるべき位置


    '1. 文字列を探索
    Dim existenceSuccess As Boolean, existenceError As Boolean
    existenceSuccess = (InStr(StartSearchPos, jsonString, SearchKeyNameOK & SearchKeyid) = StartSearchPos)
    existenceError = (InStr(StartSearchPos, jsonString, SearchKeyNameNG & SearchKeyid) = StartSearchPos)

    '2. どっちも所定の位置にない場合は、存在しないとし、抜ける
    '下記の例に基づき、10文字目にない場合は、コマンド結果ではないとして、扱います。
    '・{"type":"success", "id":... （コマンドの成功）
    '・{"type":"error", "id":... （コマンドの失敗）
    '・{"type":"error","error":...(構文エラー)　※これは通さない
    If existenceSuccess Or existenceError Then
        GetContentsID = True

        '3. 検索ワードを適用
        Dim SearchKeyName As String
        SearchKeyName = IIf(existenceSuccess, SearchKeyNameOK, SearchKeyNameNG)

        '4. 開始位置をうまい具合にずらす
        Dim StartPos As Long
        StartPos = StartSearchPos + Len(SearchKeyName) + Len(SearchKeyid) '「","id":」分、ずらすイメージ

        '5. ID値の終点を求める
        Dim EndPos As Long
        EndPos = InStr(StartPos, jsonString, ",")

        '6. ID値を取り出す
        idValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, LogID, "This BiDi response is the BiDicommand result. ID:" & idValueContents, FromProcedureName
    Else
        GetContentsID = False
        printMsg Trace_, LogID, "This BiDi response does not have an `id`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`method`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：jsonString             BiDiのレスポンス(JSON文字列そのまま)
'*            methodValueContents  非同期イベントのメソッド名(参照渡し)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：前方一致、後方一致を確認します
'***************************************************************************************************
Private Function GetContentsMethod(jsonString As String, ByRef methodValueContents As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".GetContentsMethod"


    '検索設定
    Const SearchKeyName     As String = "event" '検索文字列
    Const StartSearchPos    As Long = 10        'あるべき位置


    '1. 前方文字列を探索
    Dim StartPos As Long, EndPos As Long
    StartPos = InStr(StartSearchPos, jsonString, SearchKeyName)

    '2. 所定の位置あるか確認する
    '下記の例に基づき、10文字目にない場合は、次に後方一致を確認
    '・{"type":"event","method":...
    If StartPos = StartSearchPos Then
        GetContentsMethod = True

        '3. 開始位置をうまい具合にずらす
        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""",""method"":""") '「","method":"」分、ずらすイメージ

        '4. method名の終点を求める
        EndPos = InStr(StartPos, jsonString, """,")

        '5. method名を取り出す
        methodValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, LogID, "This BiDi response is an asynchronous event: " & methodValueContents, FromProcedureName
    
    '下記の例に基づき、後方にない場合は、非同期イベントではないとして、扱います。
    '・{"method":"network.beforeRequestSent", ... (中略) ..., "type":"event"}
    ElseIf Right(jsonString, Len(SearchKeyName) + Len("""}")) = SearchKeyName & """}" Then
        GetContentsMethod = True

        '3. 開始位置をうまい具合にずらす
        StartPos = 1 + Len("{""method"":""") '「{"method":"」分、ずらすイメージ

        '4. method名の終点を求める
        EndPos = InStr(StartPos, jsonString, """,")

        '5. method名を取り出す
        methodValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, LogID, "This BiDi response is an asynchronous event: " & methodValueContents, FromProcedureName
    Else
        GetContentsMethod = False
        printMsg Trace_, LogID, "This BiDi response does not have an `method`.", FromProcedureName
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

Property Get LastCommandID() As Long    '最後に実行したcommandID値
    LastCommandID = m_MessageID
End Property

'***************************************************************************************************
'* 機能　　：ログ設定値、取得用プロパティ
'***************************************************************************************************
Property Get ShowLogBiDiresponse() As Boolean
    ShowLogBiDiresponse = m_CDPCore.ShowLogCDPresponse
End Property



'***************************************************************************************************
'                           ■■■ ワークシートからJS取得 ■■■
'***************************************************************************************************
'* 機能　　：指定のテーブルから、mapperTab.jsコードを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：mapperTab.jsコード
'---------------------------------------------------------------------------------------------------
'* 機能説明：テーブルに格納した分割文字列を結合して返却します
'***************************************************************************************************
Private Function Take_mapperTabJSstring() As String
    Dim tmp2
    Dim chunkNum As Long

    'JS置き場のワークシート制御へ
    With ShLibrary01_JS
        With .ListObjects(.UseTableName(1, "WebDriverBiDiCore.Take_mapperTabJSstring"))
            'データがそもそもない場合は、エラーで止める
            If .ListRows.Count = 0 Then
                Err.Raise vbObjectError, , "mapperTab.js がセットされていません"
                Exit Function
            End If

            '一旦、2次元配列で受け取る
            tmp2 = .DataBodyRange
            chunkNum = UBound(tmp2)
        End With
    End With

    '1次元配列に変換
    ReDim tmp1(1 To chunkNum)
    Dim i As Long
    For i = 1 To chunkNum
        tmp1(i) = tmp2(i, 1)
    Next

    '返却
    Take_mapperTabJSstring = Join(tmp1, "")
End Function



'***************************************************************************************************
'                               ■■■ 更新関連 ■■■
'***************************************************************************************************
'* 機能　　：受け取った`mapperTab.js`のコードを、32,767文字ごとに分割し、テーブルに格納します
'***************************************************************************************************
Private Sub InsertMapperTab(Code As String)
    '1セルごとの格納文字数上限
    Const LimitCellCharacter As Long = 32767

    'テーブル追加時のアニメーション動作を止める
    Application.ScreenUpdating = False


    '1. 分割数を確認
    Dim NumberOfSplit As Long: NumberOfSplit = Len(Code) / LimitCellCharacter

    '2. 配列確保
    ReDim JsCodeChunk(NumberOfSplit, 0)

    '3. 分割して格納する
    Dim i As Long
    For i = 0 To NumberOfSplit
        JsCodeChunk(i, 0) = Mid(Code, LimitCellCharacter * i + 1, LimitCellCharacter)
    Next

    'JS置き場のワークシート制御へ
    With ShLibrary01_JS
        With .ListObjects(.UseTableName(1, "WebDriverBiDiCore.InsertMapperTab"))

            '4. テーブルの中身を全消去
            If .ListRows.Count Then .DataBodyRange.Delete
            
            '5. 挿入範囲の確定
            Dim TargetRange As Range
            Set TargetRange = .ListColumns(1).Range(2).Resize(NumberOfSplit + 1, 1)

            '6. 挿入前に書式を「文字列(@)」に設定
            TargetRange.NumberFormatLocal = "@"

            '7. 挿入
            TargetRange.value = JsCodeChunk
        End With
    End With


    '元に戻す
    Application.ScreenUpdating = True
End Sub

'***************************************************************************************************
'* 機能　　：`mapperTab.js`置き場のExcelテーブルの中身を削除します
'***************************************************************************************************
Private Sub DeleteMapperTab()
    'テーブル削除時のアニメーション動作を止める
    Application.ScreenUpdating = False


    'JS置き場のワークシート制御へ
    With ShLibrary01_JS
        With .ListObjects(.UseTableName(1, "WebDriverBiDiCore.DeleteMapperTab"))
            'テーブルの中身を全消去
            If .ListRows.Count Then .DataBodyRange.Delete
        End With
    End With


    '元に戻す
    Application.ScreenUpdating = True
End Sub

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
    'ファイルから、Jsコードを取得
    '※Windows方式の改行コード(CRLF)だと、`CR`と`LF`での分割が起こって不整合になる可能性があるため、Linux方式(LF)に統一させます
    Dim m_CharConv As New CharacterCodeConversion
    Dim JsCode As String
    With m_CharConv
        JsCode = Replace(.BytesToString(.BytesFromSavedFile(FolderPath, FileName)), vbCrLf, vbLf)
    End With

    'テーブルに置く作業
    InsertMapperTab JsCode

    '正しく置けてるか確認
    UpdateFromLocalFile = (Take_mapperTabJSstring = JsCode)

    'ミスの場合はテーブルの中身を削除
    If Not (UpdateFromLocalFile) Then DeleteMapperTab

    'ローカルファイルで更新した旨を記録
    '※ミスの場合は、空欄にする
    With ShLibrary01_JS
        .Range(.UseRangeName(1, "WebDriverBiDiCore.UpdateFromLocalFile")).value = IIf(UpdateFromLocalFile, " Local", "")
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
'***************************************************************************************************
Public Function UpdateFromNPMFile(Optional Version As String = "latest") As Boolean
    Dim ResultResponse As WebResponse

    'リクエストパラメーターを設定する
    Dim リクエスト情報一式 As WebRequest: Set リクエスト情報一式 = New WebRequest
    With リクエスト情報一式
        'GET
        .Method = vbWebHttpGet

        'レスポンスのデータ形式を指定
        .ResponseFormat = vbWebFormatPlainText

        '`chromium-bidi`のnpmパスパラメーターを指定(ホスト名に続くURL)
        .Resource = "npm/" & npmName & "@{Version}/lib/iife/" & MapperJSName
        .AddUrlSegment "Version", Version
    End With

    'リクエストして、結果を返す
    With jsDelivrへリクエスト(リクエスト情報一式, True)
        '200確認
        If .StatusCode = vbWebStatusOk Then
            'テーブルに置く作業
            InsertMapperTab .Content

            '正しく置けてるか確認
            UpdateFromNPMFile = (Take_mapperTabJSstring = .Content)

            'ミスの場合はテーブルの中身を削除
            If Not (UpdateFromNPMFile) Then DeleteMapperTab

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
'                            ■■■ CDPCore.cls 直通メソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, LogID As String, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    m_CDPCore.printMsg LogLevel_, LogID, strMsg, From, isHeader, doRaiseError
End Sub

Public Sub sleep(Optional seconds As Double = 0.5)
    m_CDPCore.sleep seconds
End Sub

Public Sub TakeEvents(Optional destruction As Boolean)
    m_CDPCore.TakeEvents destruction
End Sub

Public Function TimerCounter() As Double
    TimerCounter = m_CDPCore.GetTicksMs
End Function

Public Sub ClosePipeCDP()
    m_CDPCore.ClosePipeCDP
End Sub



'***************************************************************************************************
'                               ■■■ 初期化処理 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'Assign an objID for easy debug identification
    LogID = "BCID" & Format(Rnd * 1000, "000")

    'WebDriverBiDiのデバッグ設定を、ワークシートから適用します
    With ShSetting01_StartBrowser
        BiDiDebug.ViewDebugTab = .Range(.UseRangeName(11, FromProcedureName)).value    '`BiDi-CDP Mapper`のコンソールタブを表示するか？
        BiDiDebug.seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value 'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
        BiDiDebug.DiagnosticLog = .Range(.UseRangeName(13, FromProcedureName)).value   'Mapper内部のデバッグログを記録するか？
    End With
End Sub
