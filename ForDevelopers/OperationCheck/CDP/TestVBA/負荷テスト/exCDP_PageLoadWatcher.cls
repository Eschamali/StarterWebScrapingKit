VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_PageLoadWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           タブ(CDPContext)のページ読み込み完了を、CDP非同期イベントで検知する拡張クラスです
'   `ForDevelopers/TemplateExtensions/CDP/Normal/exCDP_Template.cls`をベースに作成しています
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定義変数/定数 ■■■
'***************************************************************************************************
'ブラウザ制御オブジェクト
Private WithEvents ex_CDPCore       As CDPCore      'イベント受け取りオブジェクト
Attribute ex_CDPCore.VB_VarHelpID = -1
Private ex_CDPBrowser               As CDPBrowser   '継承元のブラウザクラス
Private ex_CDPContext               As CDPContext   '継承元のタブ

'内部使用用
Private Const ThisClassName         As String = "exCDP_PageLoadWatcher"

'ステート管理
Private currentLog                 As StateLog     'ログ設定
Private EnableBrowserEvents         As Boolean      '非同期イベント処理を有効化するかのスイッチング
Private CurrentSessionID            As String       '`CDPContext.cls`継承時のSesssionID

'監視結果
Private HasLoadedFlag               As Boolean      '直近の`NavigateAsync`以降、`Page.loadEventFired`を受信したか
Private NetworkEventCountValue      As Long         '直近の`NavigateAsync`以降に受信した`Network.requestWillBeSent`の件数



'***************************************************************************************************
'                                   ■■■ 公開プロパティ ■■■
'***************************************************************************************************
'* 機能　　：非同期イベント処理を行うかのスイッチングです
'***************************************************************************************************
Public Property Get EnableEvents() As Boolean
    EnableEvents = EnableBrowserEvents
End Property
Public Property Let EnableEvents(ByVal v As Boolean)
    EnableBrowserEvents = v
End Property

'***************************************************************************************************
'* 機能　　：直近の`NavigateAsync`以降、`Page.loadEventFired`を受信したかを返します
'***************************************************************************************************
Public Property Get HasLoaded() As Boolean
    HasLoaded = HasLoadedFlag
End Property

'***************************************************************************************************
'* 機能　　：直近の`NavigateAsync`以降に受信した`Network.requestWillBeSent`の件数を返します
'***************************************************************************************************
Public Property Get NetworkEventCount() As Long
    NetworkEventCount = NetworkEventCountValue
End Property



'***************************************************************************************************
'                     ■■■ メインとなる機能拡張部分 - 非同期イベント処理 ■■■
'***************************************************************************************************
'* 機能　　：タブ単位の非同期イベントを受信した時にここが呼ばれます
'***************************************************************************************************
Private Sub ex_CDPCore_CDPContextEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPCore_CDPContextEvent"


    '1. 非同期イベント処理そのものが無効状態なら、即抜け
    If Not EnableBrowserEvents Then Exit Sub

    '2. このタブクラスで設定中の`SessionId`以外は無視
    If CurrentSessionID <> sessionID Then Exit Sub

    '3. 監視対象の非同期イベントを判定
    Select Case methodName
        Case "Page.loadEventFired"
            HasLoadedFlag = True
            If currentLog.seeRawSendMsgDbg Then printMsg Debug_, "`" & methodName & "` was detected.", FromProcedureName

        Case "Network.requestWillBeSent"
            NetworkEventCountValue = NetworkEventCountValue + 1

        Case Else
            'このClassでは処理しない非同期イベント
    End Select
End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************
'* 機能　　：このタブを、指定URLへ非同期で遷移させます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：発行前に、`HasLoaded`/`NetworkEventCount`をリセットします（前回遷移分の値を次回に持ち込まないため）
'* 注意事項：遷移の完了確認は行いません。別途`HasLoaded`や`AutoWaitLoaded`をご利用ください
'***************************************************************************************************
Public Sub NavigateAsync(Url As String)
    HasLoadedFlag = False
    NetworkEventCountValue = 0

    Dim CDPparams As New Dictionary
    CDPparams.Add "url", url
    ex_CDPContext.ExecuteCDPAsync "Page.navigate", CDPparams
End Sub

'***************************************************************************************************
'* 機能　　：`HasLoaded`が`True`になるまで待機し続けます
'---------------------------------------------------------------------------------------------------
'* 返り値　：`True`で読み込み完了、`False`でタイムアウト
'* 引数　　：TimeOutSecond  タイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 注意事項：内部で`TakeEvents`を呼ぶため、複数タブを並行監視したい場合は、
'            この関数ではなく、`.bas`側で`TakeEvents`と`HasLoaded`を使ってバリアループを組んでください
'***************************************************************************************************
Public Function AutoWaitLoaded(Optional TimeOutSecond As Double = 30) As Boolean
    Const FromProcedureName As String = ThisClassName & ".AutoWaitLoaded"


    Dim timerStart As Double: timerStart = CDPHelpers.TimerCounter
    Do
        If HasLoadedFlag Then
            AutoWaitLoaded = True
            Exit Function
        End If

        If CDPHelpers.TimerCounter - timerStart > TimeOutSecond * 1000 Then
            printMsg WARN_, "AutoWaitLoaded timed out.", FromProcedureName
            Exit Function
        End If

        TakeEvents  '内部でパイプをポンプします
        CDPHelpers.Sleep 0.05
    Loop
End Function



'***************************************************************************************************
'                           ■■■ `CDPContext/CDPBrowser/CDPCore`直通メソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub

Private Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    ex_CDPCore.TakeEvents StopPipeError, destruction
End Sub



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：このClassを動かすのに必要なオブジェクトを継承し、監視に必要なCDPドメインを有効化します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`Page.loadEventFired`/`Network.requestWillBeSent`はタブ(session)単位のイベントのため、
'            `CDPContext`（`CDPBrowser`ではない）から継承する構成にしています
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    '必要なCDPオブジェクトを継承
    Set ex_CDPContext = Inheritance
    Set ex_CDPBrowser = Inheritance.ThisCDPBrowser
    Set ex_CDPCore = Inheritance.ThisCDPCore
    CDPHelpers.InheritLogSettings currentLog, ex_CDPCore.LogSettings

    'このタブのSessionIdを記録（他タブの非同期イベントを無視するため）
    CurrentSessionID = ex_CDPContext.CurrentSessionID

    '監視に必要なCDPドメインを有効化
    ex_CDPContext.ExecuteCDP "Page.enable"
    ex_CDPContext.ExecuteCDP "Network.enable"
End Sub

'***************************************************************************************************
'* 機能　　：初期設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'タブ非同期イベント処理を有効化
    EnableBrowserEvents = True

    currentLog.LogID = "PLID" & Format(Rnd * 1000, "000")
End Sub
