VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverBiDiPlusContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'                         BiDi+として扱いやすくするためのラッパークラスです
'                           使用イメージとしてはアップグレードみたいな感じです
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const ThisClassName     As String = "WebDriverBiDiPlusContext"  'ログのトレース用
Private Const BiDiPlusSign      As String = "goog:cdp."                 'BiDi+の目印



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_BiDiCore  As WebDriverBiDiCore
Attribute ex_BiDiCore.VB_VarHelpID = -1
Private ex_BiDiContext          As WebDriverBiDiContext

'操作中のタブ情報一式
Private Type StateTab
    'コアとなる情報
    sessionID           As String   'The ID to the current session tab.
    targetID            As String   'The targetID of the current tab for ExecuteCDP
    browserContextId    As String   'どの`Profile`に属してるか

    '補助情報
    WID         As String   'Useful for windows repositioning
    Title       As String   'The focused tab name / window name of the active CDPBrowser
    Url         As String   'The focused tab url of the active CDPBrowser
End Type
Private BiDiPlusTab As StateTab

'CDPイベント情報
Private Type CdpEventReceivedParameters
    event   As String
    params  As BiDiCDPJson
    session As String
End Type

'その他
Private currentLog              As StateLog     'ログ設定



'***************************************************************************************************
'                                   ■■■ イベント定義 ■■■
'***************************************************************************************************
'BiDi+によるCDPレスポンスの種類に応じた発信機
Public Event CDPBrowserEvent(methodName As String, params As BiDiCDPJson)                       'ブラウザに対する非同期イベント
Public Event CDPContextEvent(methodName As String, params As BiDiCDPJson, sessionID As String)  'タブに対する非同期イベント



'***************************************************************************************************
'                              ■■■ BiDi+として扱う準備 ■■■
'***************************************************************************************************
Public Sub UpgradeBiDiPlus(TargetContext As WebDriverBiDiContext)
    Const FromProcedureName As String = ThisClassName & ".UpgradeBiDiPlus"


    '1. オブジェクトを継承
    Set ex_BiDiContext = TargetContext
    Set ex_BiDiCore = TargetContext.ThisWebDriverBiDiCore

    '2. ログ情報を継承
    CDPHelpers.InheritLogSettings currentLog, ex_BiDiCore.DirectControlCDP.ThisCDPCore.LogSettings

    '3. BiDi+用の`SessionID`を取得
    BiDiPlusTab.sessionID = ExecuteBiDiPlusGetSession()
    If LenB(BiDiPlusTab.sessionID) = 0 Then printMsg ERROR_, "The upgrade to BiDi+ failed.", FromProcedureName, doRaiseError:=True

    '4. 情報を更新
    getSessionInfo
    printParams
End Sub

'***************************************************************************************************
'* 機能　　：`goog:cdp.getSession`を実行し、BiDiの`Context`に紐づく、BiDi+用SessionIDを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`goog:cdp.sendCommand`で扱える`SessionID`
'***************************************************************************************************
Private Function ExecuteBiDiPlusGetSession() As String
    '内部SessionIDを取得するBiDi+コマンドを叩く
    ExecuteBiDiPlusGetSession = ex_BiDiContext.ExecuteBiDi(BiDiPlusSign & "getSession").StringKey("session")
End Function

'***************************************************************************************************
'* 機能　　：CDP系のイベント購読切り替えをします
'---------------------------------------------------------------------------------------------------
'* LET引数 ：`events`に渡す1次元配列 or 文字列
'     引数 ：subscribe  `True` :`session.subscribe`として実行   ※既定
'                       `False`:`session.unsubscribe`として実行
'---------------------------------------------------------------------------------------------------
'* 詳細説明：仕様としては、WebView2のイベント購読と一緒です
'* 注意事項：・`goog:cdp.○○`と付けず、CDPイベント名そのまんまで付与してください
'            ・`Collection`には非対応です
'***************************************************************************************************
Property Let SubscribeCdpEvent(Optional subscribe As Boolean = True, events)
    '配列判定
    If IsArray(events) Then
        'ループさせて、接頭辞付けていく
        Dim i As Long
        For i = LBound(events) To UBound(events)
            events(i) = BiDiPlusSign & events(i)
        Next
    Else
        '接頭辞付けて、配列化
        events = Array(BiDiPlusSign & events)
    End If

    '実際の購読へ
    ex_BiDiContext.sessionSubscribe(subscribe) = events
End Property

Public Sub getSessionInfo()
'----------------------------------------------------------------------------------
' Obtain session info on the fly for printParams.
' Updated: 22/03/2023
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".getSessionInfo"


    'Get infos
    Dim TargetInfo As BiDiCDPJson: Set TargetInfo = refreshTargetInfo
    BiDiPlusTab.Url = TargetInfo.StringKey("url")
    BiDiPlusTab.Title = TargetInfo.StringKey("title")

    '対象タブが非表示の場合、`Browser.getWindowForTarget`が効かなくなるので、判定をかます
    Dim tmp As BiDiCDPJson: Set tmp = ExecuteCDP("Browser.getWindowForTarget", StopBiDiError:=False)
    If tmp Is Nothing Then printMsg info_, "The tab you are connecting to may be hidden.", FromProcedureName Else BiDiPlusTab.WID = tmp.StringKey("windowId")

    'タブIDを記録
    BiDiPlusTab.targetID = TargetInfo.StringKey("targetId")

    'ProfileIDを記録
    BiDiPlusTab.browserContextId = TargetInfo.StringKey("browserContextId")

    'JavaScript関連の変化も捉えれるようにする
    '※`reattach`を想定して一旦OFFにして、再度ONにすることで、`Runtime.executionContextCreated`等が飛んでくれます
    runtimeEnable False
    runtimeEnable True
End Sub

'***************************************************************************************************
'* 機能　　：`SessionID`に紐づく`Target.TargetInfo`オブジェクトを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`Target.TargetInfo`オブジェクト。　※取得失敗時は、`Nothing`で返ります。大抵は、context(タブ等)が消えてることが多いです
'***************************************************************************************************
Private Function refreshTargetInfo() As BiDiCDPJson
    '1. このClassで設定中の`SessionID`を基に情報取得を試みる
    Dim TargetInfo As BiDiCDPJson
    Set TargetInfo = ExecuteCDP("Target.getTargetInfo")

    '2. エラーがなければ、`targetInfo`の中身を返却
    If TargetInfo.ExistsKey("targetInfo") Then Set refreshTargetInfo = TargetInfo.NodeKey("targetInfo")

End Function



'***************************************************************************************************
'                              ■■■ CDPcommand操作関連 ■■■
'***************************************************************************************************
'* 機能　　：`goog:cdp.sendCommand`を実行し、`goog:cdp.getSession`で取った`SessionID`に紐づくCDPコマンドを非同期で、実行します
'---------------------------------------------------------------------------------------------------
'* 返り値　：実行コマンドID         実行時のID値(整理券番号のイメージでOK) ※API失敗時は0で、成功時は1以上です。
'* 引数　　：methodName             method名。"Network.getCookies"や"DOM.getDocument"というCDPコマンド名として、記入します
'            params                 Json-Dictionary形式で、KeyとItemを事前に定義したオブジェクト変数を指定します。
'            BrowserContext         `True`で`SessionID`なしによる、ブラウザ単位への命令とします
'            StopApiError           `True`で、実行失敗時に`Err.Raise`で停止し、以降の処理を止めます
'***************************************************************************************************
Public Function ExecuteCDPAsync(methodName As String, Optional params As Dictionary, Optional BrowserContext As Boolean, Optional StopApiError As Boolean = True) As Long
    'CDP実行し、実行IDだけ回収
    ExecuteCDPAsync = ExecuteBiDiAsync(BiDiPlusSign & "sendCommand", MakeCDPparams(methodName, params, BrowserContext), StopApiError)
End Function

'***************************************************************************************************
'* 機能　　：`goog:cdp.sendCommand`を実行し、CDPコマンドを同期実行します
'---------------------------------------------------------------------------------------------------
'* 返値    ：ブラウザから返却されたCDP結果 (パースされたNode情報)
'* 引数　　：※`ExecuteBiDiPlusAsyncSendCommand`と一緒なので省略
'---------------------------------------------------------------------------------------------------
'*注意事項 ：CDPエラーについては、BiDi内部が解決してくれるため、ここでのエラーチェックは飛ばします
'***************************************************************************************************
Public Function ExecuteCDP(methodName As String, Optional params As Dictionary, Optional BrowserContext As Boolean, Optional StopBiDiError As Boolean = True) As BiDiCDPJson
    Const FromProcedureName As String = ThisClassName & ".ExecuteCDP"


    '1. CDP実行し、CDP結果部分を回収
    Set ExecuteCDP = ExecuteBiDi(BiDiPlusSign & "sendCommand", MakeCDPparams(methodName, params, BrowserContext), StopBiDiError).NodeKey("result")

    '2. デバッグ用に、As Is 出力
    If currentLog.seeRawSendMsgDbg Then printMsg Debug_, "ExecuteBiDi+ `" & methodName & "` with result: " & ExecuteCDP.Stringify, FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンドのParamsを作成し、CDPコマンドを実行できるように準備します
'***************************************************************************************************
Private Function MakeCDPparams(methodName As String, params As Dictionary, BrowserContext As Boolean) As Dictionary
    'パラメーターを組み立て
    Set MakeCDPparams = New Dictionary
    MakeCDPparams.Add "method", methodName
    If Not BrowserContext Then MakeCDPparams.Add "session", BiDiPlusTab.sessionID  'Browser単位の場合は、`session`項目を省略
    If Not params Is Nothing Then MakeCDPparams.Add "params", params               '省略時は、`params`項目を省略
End Function



'***************************************************************************************************
'                           ■■■ 各種CDPヘルパーコマンド ■■■
'***************************************************************************************************
'* 機能　　：`Runtime`系のイベント機能を有効/無効化を行います
'***************************************************************************************************
Public Sub runtimeEnable(Optional enable As Boolean = True)
    ExecuteCDP IIf(enable, "Runtime.enable", "Runtime.disable"), StopBiDiError:=False
End Sub



'***************************************************************************************************
'                           ■■■ CDPイベント捌き用 ■■■
'***************************************************************************************************
Private Sub ex_BiDiCore_detectionBiDiEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_BiDiCore_detectionBiDiEvent"


    '1. CDPイベント以外は即抜け
    Dim tmpLen As Long: tmpLen = Len(BiDiPlusSign)
    If Left(methodName, tmpLen) <> BiDiPlusSign Then Exit Sub

    '2. CDPイベント部分を取り出す
    Dim CDPparams As BiDiCDPJson: Set CDPparams = BiDiCDPJson.Parse(RawJson).NodeKey("params")
    Dim CurrentCDP As CdpEventReceivedParameters
    CurrentCDP.event = CDPparams.StringKey("event")
    Set CurrentCDP.params = CDPparams.NodeKey("params")
    CurrentCDP.session = CDPparams.StringKey("session")
    If currentLog.seeRawSendMsgDbg Then printMsg Debug_, "BiDi+ Event: " & CDPparams.Stringify, FromProcedureName

    '3. SessionIDの状況を確認
    With CurrentCDP
        If LenB(.session) Then
            'タブ単位として発信
            RaiseEvent CDPContextEvent(.event, .params, .session)
        Else
            'ブラウザ単位として発信
            RaiseEvent CDPBrowserEvent(.event, .params)
        End If
    End With
End Sub



'***************************************************************************************************
'                                   ■■■ ステータス系 ■■■
'***************************************************************************************************
Public Sub printParams()
'------------------------------------------------------------------
' Print all parameters of the current session for debuggings.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".printParams"


    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, currentLog.LogID & " Tab - Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "Session ID        | sessionID    | " & BiDiPlusTab.sessionID, FromProcedureName
    printMsg Debug_, "Window ID         | brWId        | " & BiDiPlusTab.WID, FromProcedureName
    printMsg Debug_, "Tab URL           | brUrl        | " & BiDiPlusTab.Url, FromProcedureName
    printMsg Debug_, "Tab Title         | brName       | " & BiDiPlusTab.Title, FromProcedureName
    printMsg Debug_, "Tab ID            | targetID     | " & BiDiPlusTab.targetID, FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

End Sub



'***************************************************************************************************
'                           ■■■ CDPHelpers ラッパーメソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".bidi+.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub



'***************************************************************************************************
'                           ■■■ 各種BiDiラッパーメソッド ■■■
'***************************************************************************************************
Private Function ExecuteBiDiAsync(methodName As String, Optional params As Dictionary, Optional StopError As Boolean = True) As Long
    ExecuteBiDiAsync = ex_BiDiContext.ThisWebDriverBiDiMode.ExecuteBiDiAsync(methodName, params, StopError)
End Function

Private Function ExecuteBiDi(methodName As String, Optional params As Dictionary, Optional StopError As Boolean = True) As BiDiCDPJson
    Set ExecuteBiDi = ex_BiDiContext.ThisWebDriverBiDiMode.ExecuteBiDi(methodName, params, StopError)
End Function



'***************************************************************************************************
'                               ■■■ 純BiDiメソッド提供 ■■■
'***************************************************************************************************
Property Get ThisWebDriverBiDiContext() As WebDriverBiDiContext
    Set ThisWebDriverBiDiContext = ex_BiDiContext
End Property



'***************************************************************************************************
'                       ■■■ このClassオブジェクト生成/破棄時のアクション ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'Assign an objID for easy debug identification
    currentLog.LogID = "BPID" & Format(Rnd * 1000, "000")
End Sub
