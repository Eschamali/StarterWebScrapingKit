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
' CDPBrowser経由でBiDiマッパーを実行し、BiDiコマンドの送受信・イベント取得を行います。
' `RaiseEvent`を利用した`CDPCore.cls`←→`BiDiCore`の行き来で、実現していきます
' 処理ロジックはほとんど、`CDPBrowser`を踏襲してます
'==============================================================================================================
Option Explicit



'===================================
' Global Variables
'===================================
'ブラウザ制御クラスオブジェクト
Private WithEvents m_CDPCore    As CDPCore      'イベント受け取りオブジェクト
Attribute m_CDPCore.VB_VarHelpID = -1
Private m_CDPBrowser            As CDPBrowser   'ブラウザ操作用
Attribute m_CDPBrowser.VB_VarHelpID = -1
Private WebDriverBiDi_mapperTab As CDPTab       'BiDi-CDP Mapper 操作用タブ

'非同期イベント発火
Public Event detectionBiDiEvent(methodName As String, EventRawJson As String)

'クラスオブジェクト
Private m_CharConv              As New CharacterCodeConversion
Private m_JsonConverter         As New WebJsonConverter

'CDPからWebDriverBiDi用のPayload結果を受け取るためのキュー
Private QueueResultBiDi As Collection

'BiDiコマンドの結果を一時蓄積する際の管理(`WithEvents`によるズレ防止の仕組み作り)
Private Type cacheResultBiDi
    response    As Dictionary   '「ID - それに対するBiDiコマンドの結果」というDictionary
    Limit       As Long         'その`Dictionary`に詰め込む上限数
End Type
Private currentResultBiDi As cacheResultBiDi

'デバッグ設定状況
Private Type StateDebug
    ViewDebugTab        As Boolean      '「BiDi-CDP Mapper is controlling this tab」といったデバックタブ画面を表示するか？※非表示でも、「edge://inspect/#other」で存在確認はできます。
    DiagnosticLog       As Boolean      '「BiDi-CDP Mapper」内の診断ログ情報を出力するか
    seeRawSendMsgDbg    As Boolean      'BiDiのresponseをデバッグログに出すかの調整
End Type
Private currentDebug As StateDebug

'操作中のタブ情報一式
Private Type StateTab
    sessionID   As String   'The ID to the current session tab.
    targetID    As String   'The targetID of the current tab for ExecuteCDP
End Type
Private BiDiTab As StateTab

'ステート管理
Private m_MessageID                 As Long         'BiDiコマンド結果を確認するためのID管理
Private m_Events                    As Dictionary   'BiDi向けの非同期イベント蓄積
Private m_ErrorDetails              As Dictionary   '同期BiDiコマンドした際の最後のエラー情報
Private timerOut                    As Long         '同期BiDiコマンドした際のタイムアウト設定



'===================================
' Constants
'===================================
'内部使用用
Private Const ThisClassName As String = "WebDriverBiDiCore"             'ログのトレース用
Private Const MapperURL     As String = "about:blank#" & ThisClassName  '`BiDi-CDP Mapper`の目印URL
Private Const BiDiResDebug  As String = "sendDebugMessage"              '`BiDi-CDP Mapper`内のJSがデバッグ情報を返す際の関数名(chromium-bidi\src\bidiServer\MapperCdpConnection.ts に基づく名称)
Private Const BiDiResMain   As String = "sendBidiResponse"              '`BiDi-CDP Mapper`内のJSが結果を返す際の関数名        (chromium-bidi\src\bidiServer\MapperCdpConnection.ts に基づく名称)
Private Const LimitID           As Long = 2000000000                    'BiDiコマンド送信時のID上限値
Private Const TIMEOUT_DEFAULT   As Long = 30                            '初期タイムアウト秒数

'nmp情報
Private Const MapperJSName  As String = "mapperTab.js"  'WebDriverBiDiのやり取りの核となるJavaScriptファイル名
Private Const npmName       As String = "chromium-bidi" '公式リポジトリ名称`https://github.com/GoogleChromeLabs/chromium-bidi`

'BiDiからのイベント情報を格納する際のDictionaryKey名設定
Private Const EventsDictionaryKeyName01 As String = "TotalEvents"
Private Const EventsDictionaryKeyName02 As String = "EventMethods"



'***************************************************************************************************
'               ■■■ CDP非同期イベントからBiDi用の結果を受け取る処理 ■■■
'***************************************************************************************************
Private Sub m_CDPCore_CDPTabEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".m_CDPCore_CDPTabEvent"

    'BiDiでのやりとりの核となるイベント名
    Const BiDiForCDPEventName As String = "Runtime.bindingCalled"


    '1. このタブクラスで設定中の`SessionId`以外は即抜け
    If sessionID <> BiDiTab.sessionID Then Exit Sub

    '2. 対象のイベント以外も即抜け
    If methodName <> BiDiForCDPEventName Then Exit Sub
    If currentDebug.seeRawSendMsgDbg Then printMsg Trace_, "`Runtime.bindingCalled` was detected. Raw-Json : " & RawJson, FromProcedureName

    '3. PayloadのName部分を取得
    Dim NamePayload As String: NamePayload = ExtractName(RawJson)

    '4. `Name`値に応じた処理
    Dim tmp As String
    Select Case NamePayload
        'メインとなるWebDriverBiDiの結果
        Case BiDiResMain
            'エスケープ解除済みのpayload 部分を抽出
            tmp = ExtractPayload(RawJson)
            If currentDebug.seeRawSendMsgDbg Then printMsg Debug_, "Result - WebDriverBiDi: " & tmp, FromProcedureName

            '専用のCollectionに蓄積
            QueueResultBiDi.Add tmp

        'BiDi-Mapper 内部のデバッグログ用
        Case BiDiResDebug
            'エスケープ解除済みのpayload 部分を抽出し、ログ表示のみに留める
            tmp = ExtractPayload(RawJson)
            printMsg Debug_, "Debug - WebDriverBiDi: " & tmp, FromProcedureName

        'WebDriverBiDi 以外の`Runtime.bindingCalled`が返った模様
        '※この場合は、負荷の高いJsonパース処理はしない
        Case Else
            printMsg Debug_, "We detected `Runtime.bindingCalled` instances other than WebDriverBiDi. Name :" & NamePayload, FromProcedureName
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
'                              ■■■ BiDiセッション確保関連 ■■■
'***************************************************************************************************
'* 機能　　：初期化処理
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


    Dim paramsCDP As New Dictionary
    With m_CDPBrowser
        '1. 新規タブに対して、BiDiコンソールタブを準備し、セッションターゲットをメイン化
        '※タブ表示モードの際は、そのタブがフォアグラウンドにならないように仕向けます
        If Not (ReConnect) Then Set WebDriverBiDi_mapperTab = .newTab(MapperURL, setMain:=True, isHidden:=Not (currentDebug.ViewDebugTab), isBackground:=True)
        printMsg Debug_, "Booting `BiDi-CDP Mapper`... (ViewTab: " & CStr(currentDebug.ViewDebugTab) & ")", FromProcedureName

        '2. 「BiDi-CDP Mapper」のタブ情報を取り出す
        BiDiTab.targetID = WebDriverBiDi_mapperTab.CurrentTargetID
        BiDiTab.sessionID = WebDriverBiDi_mapperTab.id

        '3. BiDi用バインディングの追加
        paramsCDP.Add "bindingName", "cdp"
        paramsCDP.Add "targetId", BiDiTab.targetID
        paramsCDP.Add "inheritPermissions", True
        .ExecuteCDP "Target.exposeDevToolsProtocol", paramsCDP, Not ReConnect   '再接続モードの場合は、エラー無視
    End With

    '4. JSからVBAへ返す経路を設定
    Call BiDiResponseBiding

    With WebDriverBiDi_mapperTab
        '5. パースしたJavaScriptをブラウザ上で実行
        .jsEval Take_mapperTabJSstring

        '6. ブートストラップ（初期化）JSの実行
        .jsEval "window.runMapperInstance('" & BiDiTab.targetID & "')"
    End With

    printMsg Debug_, "`BiDi-CDP Mapper` has been launched", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：BiDiからの結果やデバッグログを受け取るためのマッピング処理を行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`CDPBrowser`を`New`してからご利用ください
'            ・各パラメータは、`chromium-bidi\src\bidiServer\MapperCdpConnection.ts`に準拠してます
'            ・SessionIDが変わるたびにこれを行わないと、`BiDi-CDP Mapper`は生きてるのに、コマンドのやり取りが出来ない問題が起こります
'***************************************************************************************************
Private Sub BiDiResponseBiding()
    Const FromProcedureName As String = ThisClassName & ".BiDiResponseBiding"


    Dim paramsCDP As New Dictionary
    With WebDriverBiDi_mapperTab
        '1. JSからVBAへ、BiDiコマンドの結果/イベント payload を返す経路を設定
        paramsCDP.Add "name", BiDiResMain
        .ExecuteCDP "Runtime.addBinding", paramsCDP

        '2. JSからVBAへ、Mapper側で発生したログ（状況説明、警告、エラー寄り情報）を返す経路を設定
        If currentDebug.DiagnosticLog Then
            paramsCDP.RemoveAll
            paramsCDP.Add "name", BiDiResDebug
            .ExecuteCDP "Runtime.addBinding", paramsCDP
        End If
    End With

    printMsg Debug_, "I have configured the route to receive from `BiDi-CDP Mapper`.", FromProcedureName
End Sub



'***************************************************************************************************
'* 機能　　：BiDiコマンドの受信データを高速にチェックし、仕分けます
'---------------------------------------------------------------------------------------------------
'* 返値    ：TRUE :受信ループを止める合図。
'               ・`BiDiJsonString`に、ExcelからのBiDi-Jsonコマンド結果がある
'               ・コマンド構文エラーによる、ID項目なし
'
'* 引数　　：BiDiJsonString         BiDiの生JSON文字列
'            InternalUseMessages    クラス内部プロシージャで、イベント内容を扱う用として蓄積させる参照変数
'*           checkCommandID         待ち受け中のコマンド結果ID
'***************************************************************************************************
Private Function BrowserReceivedDataCheck(BiDiJsonString As String, InternalUseMessages As Dictionary, checkCommandID As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".BrowserReceivedDataCheck"



    '======================================= チェック開始 =======================================
    Dim ResultCommandID As Long, AsynchronousEventName As String
    Dim JsonDictionaryObj As Dictionary
    Dim FlagValue As Long

    Const resultCommand As Long = 2 ^ 0 'IDのコマンド結果を検知出来た
    Const mismatchID    As Long = 2 ^ 1 'クラス内変数に記録してるIDとは違うコマンド結果が検知された。恐らく以前、非同期で実行したときの結果と思われる
    Const browserEvent  As Long = 2 ^ 2 '非同期イベントである
    Const EventCapture  As Long = 2 ^ 3 '非同期イベントキャプチャモードである
    Dim mask As Long: mask = resultCommand + mismatchID + browserEvent + EventCapture


    '--------- 1. BiDiからのイベントか？ ---------
    If GetContentsMethod(BiDiJsonString, AsynchronousEventName) Then
        FlagValue = FlagValue Or browserEvent   '非同期イベントとして、フラグ付け

        'VBAでのイベント発火を行う
        '※拡張機能の柔軟性を高めるため、ブラウザから得た非同期イベントは全て、イベント発火させます
        RaiseEvent detectionBiDiEvent(AsynchronousEventName, BiDiJsonString)


        '--------- 2. いずれかのDictionary変数が、非同期イベントキャプチャモードか？ ---------
        If (m_Events Is Nothing) And (InternalUseMessages Is Nothing) Then
            '両方とも`Nothing`なら、わざわざ中身を解析する意味がないので、パーサーは使わない
            printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName
        Else
            'イベントキャプチャモードなので、JsonPurserを使用する
            FlagValue = FlagValue Or EventCapture   'イベントキャプチャモードとして、フラグ付け
            printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
            Set JsonDictionaryObj = m_JsonConverter.ParseJson(BiDiJsonString)
        End If

    Else
        FlagValue = FlagValue Or resultCommand  'IDのコマンド結果を検知出来たとして、フラグ付け

        '--------- 2. `id`項目があるか？ ---------
        If GetContentsID(BiDiJsonString, ResultCommandID) Then
            '--------- 3. そのIDは、クラス内変数に記録してるIDと一致してるか？ ---------
            If ResultCommandID = checkCommandID Then
                '正しいコマンド結果が返ったので次項に、受信ループ終了フラグを返すようにする
                printMsg Trace_, "It was confirmed that the execution BiDicommand ID matches the one recorded in Excel. ID:" & checkCommandID, FromProcedureName
            Else
                '`WithEvents`によるズレなので、結果をキャッシュしておく
                FlagValue = FlagValue Or mismatchID '現在のコマンド結果IDと返ったコマンド結果IDが異なる場合のフラグ付け
                printMsg info_, "The BiDicommand result (ID:" & ResultCommandID & ") was returned, but it does not match the BiDicommand ID (" & checkCommandID & ") recorded in Excel. This is probably an asynchronous result. The results will be accumulated.", FromProcedureName

                'コマンド結果ID用受信Boxへ格納しておく
                '※パーサー処理負荷を抑えるため、Json文字列まんまとします
                currentResultBiDi.response.Add ResultCommandID, BiDiJsonString

                '格納件数チェックし、上限を超えたら、最古データを削除します
                If currentResultBiDi.response.Count > currentResultBiDi.Limit Then
                    ' 一番古いキーを取得して削除
                    printMsg Trace_, "Delete asynchronous processing result ID:" & currentResultBiDi.response.keys(0), FromProcedureName
                    currentResultBiDi.response.Remove currentResultBiDi.response.keys(0)
                    printMsg info_, "The accumulation limit has exceeded " & currentResultBiDi.Limit & ", so the oldest has been deleted.", FromProcedureName
                End If
            End If
        Else
            '`method`も`id`項目もない場合、BiDi-Jsonコマンドの構文エラーとして返った可能性があるとし、受信ループ終了とする
            printMsg WARN_, "Both the `id` and `method` fields are missing. This may be a syntax error.", FromProcedureName
        End If
    End If


    '======================================= このプロシージャから抜けるか判定 =======================================
    Select Case (FlagValue And mask)
        '送ったコマンドの結果は来たので、受信ループを止めて、抜ける
        Case resultCommand
            BrowserReceivedDataCheck = True
            Exit Function

        'コマンド結果だがID不一致。以前、非同期で実行したときの結果と思われるので、まだ受信ループは止めずに、抜ける
        Case resultCommand + mismatchID
            BrowserReceivedDataCheck = False
            Exit Function

        '受信ループは止めないが、イベントキャプチャしないモード(格納用のDictionaryがどちらもNothing or 古いSessionID)なので、次項の格納処理はせず、抜ける
        Case browserEvent
            BrowserReceivedDataCheck = False
            Exit Function

        'イベントキャプチャモードにより、Dictionaryオブジェクト格納フェーズへ行けるようにしておく。なお、受信ループは止めません
        Case browserEvent + EventCapture
            BrowserReceivedDataCheck = False

        '未知の組み合わせ
        Case Else
            printMsg ERROR_, "An unknown flag was passed. (" & (FlagValue And mask) & ")", FromProcedureName, doRaiseError:=True

    End Select


    '======================================= イベント情報を格納するフェーズへ =======================================
    'パブリック(呼び出し側使用)用
    If m_Events Is Nothing Then
        printMsg Trace_, "No event information is required on the module side.", FromProcedureName
    Else
        printMsg Trace_, "The module is requesting the accumulation of event information. Saving...", FromProcedureName
        OnNewEventReceived m_Events, JsonDictionaryObj, AsynchronousEventName
    End If

    'プライベート(クラス内変数)用
    If InternalUseMessages Is Nothing Then
        printMsg Trace_, "No event information needed inside the class.", FromProcedureName
    Else
        printMsg Trace_, "The class is requesting the accumulation of event information. Saving...", FromProcedureName
        OnNewEventReceived InternalUseMessages, JsonDictionaryObj, AsynchronousEventName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックし、`id`を高速抽出します
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
    Const SearchKeyNameOK   As String = "success"   '検索文字列(コマンド成功時)
    Const SearchKeyNameNG   As String = "error"     '検索文字列(コマンド失敗時)
    Const StartSearchPos    As Long = 10            'あるべき位置


    '1. 文字列を探索
    Dim StartPos_OK As Long, StartPos_NG As Long
    StartPos_OK = InStr(StartSearchPos, jsonString, SearchKeyNameOK)
    StartPos_NG = InStr(StartSearchPos, jsonString, SearchKeyNameNG)

    '2. 2つの結果を見て1つにまとめる
    Dim StartPos As Long
    Dim SearchKeyName As String
    StartPos = IIf(StartPos_OK, StartPos_OK, StartPos_NG)
    SearchKeyName = IIf(StartPos_OK, SearchKeyNameOK, SearchKeyNameNG)

    '3. 所定の位置にない場合は、存在しないとし、抜ける
    '下記の例に基づき、10文字目にない場合は、コマンド結果ではないとして、扱います。
    '・{"type":"success", "id":... （コマンドの成功）
    '・{"type":"error", "id":... （コマンドの失敗）
    If StartPos = StartSearchPos Then
        GetContentsID = True

        '4. 開始位置をうまい具合にずらす
        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""",""id"":") '「","id":」分、ずらすイメージ

        '5. ID値の終点を求める
        Dim EndPos As Long
        EndPos = InStr(StartPos, jsonString, ",")

        '6. ID値を取り出す
        idValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, "This BiDi response is the BiDicommand result. ID:" & idValueContents, FromProcedureName
    Else
        GetContentsID = False
        printMsg Trace_, "This BiDi response does not have an `id`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックし、`method`を高速抽出します
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
        printMsg Trace_, "This BiDi response is an asynchronous event: " & methodValueContents, FromProcedureName
    
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
        printMsg Trace_, "This BiDi response is an asynchronous event: " & methodValueContents, FromProcedureName

    Else
        GetContentsMethod = False
        printMsg Trace_, "This BiDi response does not have an `method`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：コマンドの結果にエラーが含まれているかチェックします
'---------------------------------------------------------------------------------------------------
'* 返り値　：エラー有無
'* 引数　　：ResultsObj    `invokeMethod`の生結果のDictionary
'***************************************************************************************************
Private Function invokeError(ResultsObj As Dictionary) As Boolean
    Const FromProcedureName As String = ThisClassName & ".invokeError"


    Dim ErrorType As String, ErrorMessage As String, ErrorStackTrace As String
    '結果が空っぽになってる場合は、タイムアウトエラーか、コードのミスの可能性
    If ResultsObj Is Nothing Then
        'ログを出しておく
        ErrorType = "error" 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        printMsg WARN_, "InvokeMethod Error: [" & ErrorType & "] " & ErrorMessage, FromProcedureName

        'エラーフラグを立てる
        invokeError = True
        Set m_ErrorDetails = m_JsonConverter.ParseJson("{""type"":" & ErrorType & ",""message"":""" & ErrorMessage & """}")     'カスタムエラーメッセージJsonで返却

    'BiDiコマンドエラーは "type":"error" で返る
    ElseIf ResultsObj("type") = "error" Then
        'エラーを残す
        Set m_ErrorDetails = ResultsObj

        'ログを出しておく
        ErrorType = ResultsObj("error")
        ErrorMessage = ResultsObj("message")
        printMsg WARN_, "BiDi-InvokeMethod Error: [" & ErrorType & "] " & ErrorMessage, FromProcedureName

        '`StackTrace`があったら別枠で表示
        If ResultsObj.Exists("stacktrace") Then
            ErrorStackTrace = ResultsObj("stacktrace")
            printMsg WARN_, "       " & ErrorStackTrace, FromProcedureName
        End If

        'エラーフラグを立てる
        invokeError = True
    End If
End Function



'***************************************************************************************************
'                           ■■■ `CDPCore.cls`直通メソッド ■■■
'***************************************************************************************************
Public Function TimerCounter() As Double
    TimerCounter = m_CDPCore.LogControl.GetTicksMs
End Function

Public Sub sleep(Optional seconds As Double = 0.5)
    m_CDPCore.sleep seconds
End Sub



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
        With .ListObjects(.UseTableName(1, "BiDiCore.InsertMapperTab"))
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



Private Sub Class_Initialize()
    '------------------------- クラス内変数の一部を初期化 -------------------------
    'タイムアウト設定(デフォルトで、30s)
    timerOut = TIMEOUT_DEFAULT * 1000   'ミリ秒に補正

    '非同期でのコマンド結果を蓄積させる件数(デフォルトで、256件)
    Set currentResultBiDi.response = New Dictionary
    currentResultBiDi.Limit = 2 ^ 8

    '初期化
    Set QueueResultBiDi = New Collection

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        currentDebug.ViewDebugTab = .Range(.UseRangeName(11, "WebDriverBiDiCore.Class_Initialize")).value    '`BiDi-CDP Mapper`のコンソールタブを表示するか？
        currentDebug.seeRawSendMsgDbg = .Range(.UseRangeName(8, "WebDriverBiDiCore.Class_Initialize")).value 'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
        currentDebug.DiagnosticLog = .Range(.UseRangeName(13, "WebDriverBiDiCore.Class_Initialize")).value   'Mapper内部のデバッグログを記録するか？
    End With
End Sub
