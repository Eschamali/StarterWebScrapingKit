VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_WebSocketEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'                             WebSocketに関する非同期イベントです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1
Private WithEvents st               As SafeTimer
Attribute st.VB_VarHelpID = -1
Private JsonConv                    As New WebJsonConverter

'内部使用用
Private Const ThisClassName         As String = "exCDP_NetworkEvents"             'ログのトレース用

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'responseをデバッグログに出すかの調整
Private EnableBrowserEvents         As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング



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
'                     ■■■ メインとなる機能拡張部分 非同期イベント編 ■■■
'***************************************************************************************************
'* 機能　　：非同期イベントを処理したいときにご利用ください。あなたのイメージをコードに落とし込む場所の1つです
'***************************************************************************************************
Private Sub ex_CDPBrowser_detectionCDPEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPBrowser_detectionCDPEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. WebSocket関連の非同期イベント名を、各Case に沿ったDemoプロシージャ実行
    Dim tmp As Dictionary
    Select Case methodName
        Case "Network.webSocketClosed"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName
            Set tmp = JsonConv.ParseJson(EventRawJson)("params")
            Call ShowClosed(tmp("requestId"), tmp("timestamp"))

        Case "Network.webSocketCreated"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName
            Set tmp = JsonConv.ParseJson(EventRawJson)("params")
            Call ShowCreated(tmp("requestId"), tmp("url"))

        Case "Network.webSocketFrameError"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName
            Set tmp = JsonConv.ParseJson(EventRawJson)("params")
            Call ShowFrameError(tmp("requestId"), tmp("timestamp"), tmp("errorMessage"))

        Case "Network.webSocketFrameReceived"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName
            Set tmp = JsonConv.ParseJson(EventRawJson)("params")
            Call ShowFrameReceived(tmp("requestId"), tmp("timestamp"), CStr(tmp("response")("payloadData")))

        Case "Network.webSocketFrameSent"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName
            Set tmp = JsonConv.ParseJson(EventRawJson)("params")
            Call ShowFrameSent(tmp("requestId"), tmp("timestamp"), CStr(tmp("response")("payloadData")))

        Case "Network.webSocketHandshakeResponseReceived"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName
            Set tmp = JsonConv.ParseJson(EventRawJson)("params")
            Call ShowHandshakeResponseReceived(tmp("requestId"), tmp("timestamp"), tmp("response"))

        Case "Network.webSocketWillSendHandshakeRequest"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName
            Set tmp = JsonConv.ParseJson(EventRawJson)("params")
            Call ShowWillSendHandshakeRequest(tmp("requestId"), tmp("timestamp"), tmp("request"))

        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName
            Exit Sub

    End Select

End Sub

'***************************************************************************************************
'* 機能　　：継承したClassオブジェト内にあるパイプを覗き、非同期イベントを発火(RaiseEvent detectionCDPEvent)させます
'***************************************************************************************************
Private Sub st_TimerCall(ByVal wTime As Long, varData As Variant)
    TakeEvents
End Sub



'***************************************************************************************************
'                               ■■■ イベント監視設定 ■■■
'***************************************************************************************************
'* 機能　　：WindowsAPI - SetTimer による定期的なブラウザ非同期イベントの監視を行います
'---------------------------------------------------------------------------------------------------
'* 引数    ：msec       イベント監視間隔(ms) ※デフォルトは50ms
'            varData    `st_TimerCall`に渡したい追加データ  ※このサンプルでは、未使用です
'***************************************************************************************************
Public Sub StartCheckAsyncEvents(Optional msec As Long = 50, Optional addData As Variant)
    Set st = New SafeTimer
    st.StartTimer msec, addData
End Sub



'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.invokeMethod`といった記述を防ぐ措置となります。
'---------------------------------------------------------------------------------------------------
'* 注意事項：・メインの`CDPBrowser`メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておくことをお勧めします
'            ・各詳細は、`CDPBrowser`の方をご覧ください
'***************************************************************************************************
'負荷を抑えるときにどうぞ
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_CDPBrowser.sleep seconds
End Sub

'ログメッセージを出すときにどうぞ
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

'CDP実行に必要です
Private Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set invokeMethod = ex_CDPBrowser.invokeMethod(methodName, params, alwaysBrowserContext, StopError)
End Function

'非同期イベント発火(RaiseEvent)に必要です
Private Sub TakeEvents(Optional destruction As Boolean)
    ex_CDPBrowser.TakeEvents destruction
End Sub



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`CDPBrowser.cls`を継承させます
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPBrowser)
    '継承
    Set ex_CDPBrowser = Inheritance
End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value    'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
    End With
End Sub
