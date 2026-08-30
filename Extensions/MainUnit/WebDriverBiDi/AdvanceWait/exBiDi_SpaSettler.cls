VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exBiDi_SpaSettler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'                           SPAの待機に特化したアシストを提供します
'ロジックは、「https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA/tree/main/src」準拠です
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_BiDiCore  As WebDriverBiDiCore
Attribute ex_BiDiCore.VB_VarHelpID = -1
Private ex_BiDiContext          As WebDriverBiDiContext
Attribute ex_BiDiContext.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName     As String = "exBiDi_SpaSettler.cls" 'ログのトレース用

'ステート管理
Private currentLog             As StateLog  'ログ設定
Private EnableBrowserEvents     As Boolean  'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private m_context               As String   '`WebDriverBiDiContext`継承時の`Context`情報

'待機系ログ状況管理用
Private Type DiscoveryLog
    isRecording       As Boolean
    recorderLogs        As Collection
    recSuppByKey        As Dictionary
    recorderStartTime   As Double
    excludeResources    As Boolean
    droppedHeadCount    As Long
    Limit               As Long
End Type
Private CurrentDiscoveryLog As DiscoveryLog



'***************************************************************************************************
'                               ■■■ 公開プロパティ ■■■
'***************************************************************************************************
'* 機能　　：非同期イベント処理を行うかのスイッチングです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：スコープ範囲の調整にどうぞ。名称は、「Application.EnableEvents」から取りました
'***************************************************************************************************
Public Property Get EnableEvents() As Boolean
    EnableEvents = EnableBrowserEvents
End Property

Public Property Let EnableEvents(ByVal v As Boolean)
    EnableBrowserEvents = v
End Property



'***************************************************************************************************
'                     ■■■ 非同期イベントの処理 ■■■
'***************************************************************************************************
'* 機能　　：WebDirverBiDiからの非同期イベントを処理します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`ex_BiDiContext.cls`から受け取った`Context`だけ処理を通す場合はその判定処理を別途、case内に記述して下さい。
'            CDP版と違って、イベント内容によって`Context`の記載情報位置が異なるため、イベントmethod名ごとのcontext判定というフローになります
'***************************************************************************************************
Private Sub ex_BiDiCore_detectionBiDiEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_BiDiCore_detectionBiDiEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. 待機系に必要なイベントを`FormatAndStoreLog`プロシージャへ委託
    Select Case methodName
        Case "network.beforeRequestSent", "network.responseCompleted", "network.fetchError", "browsingContext.load", "browsingContext.domContentLoaded"
            '3. 記録モードなら委託
            If CurrentDiscoveryLog.isRecording Then FormatAndStoreLog methodName, EventRawJson

    End Select
End Sub



'***************************************************************************************************
'                               ■■■ 調査機能関連 ■■■
'***************************************************************************************************
'* 機能　　：ページ読み込み待機系に関するイベントを購読し、記録モードに入ります
'---------------------------------------------------------------------------------------------------
'* 引数    ：excludeImagesAndCss    `True` :ImagesとCss を含めない
'                                   `False`:ImagesとCss を含める
'***************************************************************************************************
Public Sub StartDiscoveryLog(Optional excludeImagesAndCss As Boolean = True)
    Const FromProcedureName As String = ThisClassName & ".StartDiscoveryLog"


    '1. エラー情報をクリア
    ex_BiDiContext.ThisWebDriverBiDiMode.LastBiDiJsonError.RemoveAll

    '2. 必要なイベントを購読
    Dim TargetEvents As New Collection
    TargetEvents.Add "network.responseCompleted"
    TargetEvents.Add "network.fetchError"
    TargetEvents.Add "network.beforeRequestSent"
    TargetEvents.Add "browsingContext.load"
    TargetEvents.Add "browsingContext.domContentLoaded"
    Set ex_BiDiContext.ThisWebDriverBiDiMode.sessionSubscribe = TargetEvents

    '3. エラーcheck
    If ex_BiDiContext.ThisWebDriverBiDiMode.LastBiDiJsonError.Count Then printMsg ERROR_, "The execution of `session.subscribe` failed." & vbCrLf & "> " & ex_BiDiContext.ThisWebDriverBiDiMode.LastBiDiJsonError("message"), FromProcedureName, doRaiseError:=True

    '4. 各種初期化
    With CurrentDiscoveryLog
        Set .recorderLogs = New Collection
        Set .recSuppByKey = New Dictionary
        .recorderStartTime = CDPHelpers.TimerCounter
        .excludeResources = excludeImagesAndCss
        .droppedHeadCount = 0
        .isRecording = True
    End With
    printMsg info_, "BiDi Recorder: Started. (Filter ImagesAndCss: " & excludeImagesAndCss & ")", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：ログ生成を行います
'***************************************************************************************************
Private Sub FormatAndStoreLog(methodName As String, RawJson As String)
    '1. 経過時間を取得
    Dim relTime As String
    relTime = Format(CDPHelpers.TimerCounter - CurrentDiscoveryLog.recorderStartTime, "0.00") & "s"

    '2. `params`の取り出し
    Dim paramsNode As BiDiCDPJson
    Set paramsNode = BiDiCDPJson.Parse(RawJson).NodeKey("params")

    '3. 各イベントごとの蓄積処理
    Dim reqNode As BiDiCDPJson
    Dim LogLine As String
    Select Case methodName
        Case "network.beforeRequestSent"
            '3-1. contextが一致しないなら即抜け
            If paramsNode.StringKey("context") <> m_context Then Exit Sub

            '3-2. `request`ノード部分を取り出す
            Set reqNode = paramsNode.NodeKey("request")

            '3-3. 必要なパラメーターを取り出す
            Dim reqId       As String:      reqId = reqNode.StringKey("request")
            Dim reqMethod   As String:  reqMethod = reqNode.StringKey("method")
            Dim reqUrl      As String:     reqUrl = reqNode.StringKey("url")

            '3-4. フィルタリングモードの場合はURLのチェックに入る
            If CurrentDiscoveryLog.excludeResources Then
                If IsNoiseUrl(reqUrl) Then BumpSuppressed NormalizeEndpointKey(reqUrl): Exit Sub
            End If

            '3-4. 書き込むログ文字列を生成
            LogLine = "[" & relTime & "] [REQ] " & reqMethod & " | id=" & reqId & _
                      " | key=" & NormalizeEndpointKey(reqUrl) & _
                      " | idle=" & IIf(IsIdleIgnored(reqUrl), "skip", "track")




            ' url / method → "[rel] [REQ] ..."
        Case "network.responseCompleted"
            ' status / mime / url
            ' excludeResources なら画像・CSS は BumpSuppressed して Exit
            ' → "[rel] [RES] ..."
        Case "network.fetchError"
            ' → "[rel] [ERR] ..."
        Case "browsingContext.load", "browsingContext.domContentLoaded"
            ' → "[rel] [NAV] ..."
    End Select
    PushTimelineLine line
End Sub

Private Sub PushTimelineLine(s As String)
    If LenB(s) = 0 Then Exit Sub
    With CurrentDiscoveryLog
        If .recorderLogs.Count >= 5000 Then
            .recorderLogs.Remove 1
            .droppedHeadCount = .droppedHeadCount + 1
        End If
        .recorderLogs.Add s
    End With
End Sub

'***************************************************************************************************
'* 機能　　：待機系と直接関係ないURLかをcheckします
'***************************************************************************************************
Private Function IsNoiseUrl(Url As String) As Boolean
    '1. 小文字に統一しておく
    Dim lUrl As String: lUrl = LCase(Url)

    '2. 【先頭一致の特殊ガード】data: URI は一瞬で判定して即抜け（超巨大テキストへの無駄なInStrを防止）
    If Left$(Url, 5) = "data:" Then IsStaticAsset = True: Exit Function

    '3.【固定NGリストを用意】配列の生成は最初の1回のみ！2回目以降は0秒で使い回します
    Static patterns
    If IsEmpty(patterns) Then
        patterns = Array("/_next/static/", "webworker.js", ".png", ".jpg", ".jpeg", _
                         ".gif", ".webp", ".svg", ".css", ".woff", ".woff2", ".ico")
    End If

    '4. 【部分一致チェック】1つでも合致したらその瞬間に即抜け
    Dim p
    For Each p In patterns
        If InStr(1, lUrl, p) Then
            IsStaticAsset = True
            Exit Function   ' ヒットした瞬間にループを即終了して高速化！
        End If
    Next

    '5. ユーザー独自のフィルター設定側でのチェック
    If Not m_userIgnoreNetworkPatterns Is Nothing Then
        For Each p In m_userIgnoreNetworkPatterns
            If UrlMatchesPattern(1, lUrl, CStr(pattern)) Then
                IsNoiseUrl = True   ' ヒットした瞬間にループを即終了して高速化！
                Exit Function
            End If
        Next
    End If
End Function



'***************************************************************************************************
'                           ■■■ `WebDriverBiDi`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.ExecuteBiDi`といった記述を防ぐ措置となります。
'---------------------------------------------------------------------------------------------------
'* 注意事項：・前後の処理等も出来ますが基本は、1行だけに留めておくことをお勧めします
'            ・各詳細は、`WebDriverBiDi`の方をご覧ください
'***************************************************************************************************
'負荷を抑えるときにどうぞ
Private Sub sleep(Optional seconds As Double = 0.5)
    CDPHelpers.Sleep seconds
End Sub

'ログメッセージを出すときにどうぞ
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Const LogFileExtension As String = ".bidi.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub

'BiDi実行に必要です
Private Function ExecuteBiDi(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As BiDiCDPJson
    Set ExecuteBiDi = ex_BiDiContext.ExecuteBiDi(methodName, params, StopError)
End Function

'非同期イベント発火(RaiseEvent)に必要です
Private Sub TakeEvents(Optional destruction As Boolean)
    ex_BiDiCore.TakeEvents , destruction
End Sub



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：情報継承を行います
'***************************************************************************************************
Friend Sub Init(Inheritance As WebDriverBiDiContext)
    '継承
    Set ex_BiDiContext = Inheritance
    Set ex_BiDiCore = Inheritance.ThisWebDriverBiDiCore
    CDPHelpers.InheritLogSettings currentLog, Inheritance.ThisWebDriverBiDiCore.DirectControlCDP.ThisCDPCore.LogSettings
    
    'コンテキスト情報
    m_context = Inheritance.context
End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    '初期上限をセット
    CurrentDiscoveryLog.Limit = 5000

    currentLog.LogID = "SPID" & Format(Rnd * 1000, "000")
End Sub
