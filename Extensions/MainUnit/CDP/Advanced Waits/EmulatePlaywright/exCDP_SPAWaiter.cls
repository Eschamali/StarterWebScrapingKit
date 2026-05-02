VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_SPAWaiter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           PlaywrightのSPA待機ロジックをエミュレートする拡張機能クラスファイルです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "exCDP_SPAWaiter"             'ログのトレース用

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'responseをデバッグログに出すかの調整
Private EnableBrowserEvents         As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング

'--- Network Idle 判定用 ---
Private m_InflightCount             As Long         ' 現在通信中のリクエスト数
Private m_LastIdleTime              As Double       ' 最後に InflightCount が 0 になった時刻 (Timer値)
Private m_NetworkIdleDetected       As Boolean      ' networkidle 判定フラグ

'--- Lifecycle 判定用 ---
Private m_LoadFired                 As Boolean      ' Page.lifecycleEvent "load" 検知済み
Private m_DOMContentLoaded          As Boolean      ' Page.lifecycleEvent "DOMContentLoaded" 検知済み

'--- DOM 安定判定用 ---
Private m_LastDOMSnapshot           As String       ' 直前の DOM スナップショット (innerHTML.length)
Private m_DOMStableCount            As Long         ' DOM が変化しなかった連続回数



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
    If Not EnableBrowserEvents Then Exit Sub

    '2. 処理したい対象の非同期イベント名を、各Case に記述してください。
    Select Case methodName
        ' === Network Idle 検出 ===
        Case "Network.requestWillBeSent"
            m_InflightCount = m_InflightCount + 1
            m_NetworkIdleDetected = False
            printMsg Trace_, "Network request started. Inflight: " & m_InflightCount, FromProcedureName

        Case "Network.loadingFinished", "Network.loadingFailed"
            If m_InflightCount > 0 Then m_InflightCount = m_InflightCount - 1
            If m_InflightCount = 0 Then m_LastIdleTime = Timer
            printMsg Trace_, "Network request finished. Inflight: " & m_InflightCount, FromProcedureName

        ' === Lifecycle イベント検出 ===
        Case "Page.lifecycleEvent"
            Dim EventName As String
            EventName = ExtractLifecycleEventName(EventRawJson)
            Select Case LCase(EventName)
                Case "load":              m_LoadFired = True
                Case "domcontentloaded":  m_DOMContentLoaded = True
            End Select
            printMsg Trace_, "Lifecycle event detected: " & EventName, FromProcedureName

        ' === ナビゲーション検出 ===
        Case "Page.frameNavigated"
            ' 新しいドキュメントへの遷移（リダイレクト等含む）時は全リセット
            ' ※本来はメインフレームかどうかを判定すべきだが、簡易的に全リセット
            Call ResetState
            printMsg info_, "New document navigation detected. State reset.", FromProcedureName

        ' === SPA ルート変更検出 ===
        Case "Page.navigatedWithinDocument"
            ' SPA の pushState/replaceState による遷移。
            ' ネットワークとDOMの安定待機のみリセットし、ライフサイクルフラグは維持する。
            m_NetworkIdleDetected = False
            m_LastDOMSnapshot = ""
            m_DOMStableCount = 0
            printMsg info_, "Internal SPA navigation detected. Network/DOM flags reset.", FromProcedureName
            
    End Select
End Sub

'***************************************************************************************************
'* 機能　　：EventRawJsonから簡易的にライフサイクルのイベント名を取り出します
'***************************************************************************************************
Private Function ExtractLifecycleEventName(ByVal rawJson As String) As String
    ' "name":"load" などの箇所を探索
    Dim namePos As Long
    namePos = InStr(rawJson, """name"":""")
    If namePos > 0 Then
        Dim StartPos As Long: StartPos = namePos + 8
        Dim EndPos As Long: EndPos = InStr(StartPos, rawJson, """")
        If EndPos > StartPos Then
            ExtractLifecycleEventName = Mid(rawJson, StartPos, EndPos - StartPos)
        End If
    End If
End Function



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分編 ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：SPA待機に必要な設定 (Network.enable, Page.setLifecycleEventsEnabled) を発行します
'***************************************************************************************************
Public Sub Setup()
    Const FromProcedureName As String = ThisClassName & ".Setup"
    
    ' Networkの監視を有効化
    Call invokeMethod("Network.enable", StopError:=False)
    
    ' ライフサイクルイベントの監視を有効化
    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "enabled", True
    Call invokeMethod("Page.setLifecycleEventsEnabled", params, StopError:=False)
    
    ' イベントフックと状態の初期化
    EnableBrowserEvents = True
    ResetState
    
    printMsg info_, "SPA Waiter is successfully setup and events are enabled.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：SPA待機のリソースを開放します
'***************************************************************************************************
Public Sub Teardown()
    EnableBrowserEvents = False
    ResetState
End Sub

'***************************************************************************************************
'* 機能　　：SPA待機のフラグをリセットします。ページ遷移の直前などに呼び出します
'***************************************************************************************************
Public Sub ResetState()
    m_InflightCount = 0
    m_LastIdleTime = 0
    m_NetworkIdleDetected = False
    m_LoadFired = False
    m_DOMContentLoaded = False
    m_LastDOMSnapshot = ""
    m_DOMStableCount = 0
End Sub

'***************************************************************************************************
'* 機能　　：ネットワークが静止状態 (NetworkIdle) になるまで待機します
'---------------------------------------------------------------------------------------------------
'* 引数　　：timeoutSec       タイムアウト（秒）
'*           idleThresholdMs  Inflightな通信が0になってから、NetworkIdleと判定するまでの待機時間（ミリ秒）
'***************************************************************************************************
Public Function WaitForNetworkIdle(Optional TimeoutSec As Double = 30, Optional idleThresholdMs As Long = 500) As Boolean
    Const FromProcedureName As String = ThisClassName & ".WaitForNetworkIdle"
    
    Dim startTime As Double: startTime = Timer
    Dim idleThresholdSec As Double: idleThresholdSec = idleThresholdMs / 1000

    ' 状態を一旦リセット
    m_NetworkIdleDetected = False
    If m_InflightCount = 0 Then m_LastIdleTime = Timer

    Do
        ' パイプから非同期イベントを取り出す
        TakeEvents

        ' networkidle 判定
        If m_InflightCount = 0 And m_LastIdleTime > 0 Then
            If (Timer - m_LastIdleTime) >= idleThresholdSec Then
                m_NetworkIdleDetected = True
                WaitForNetworkIdle = True
                printMsg Trace_, "NetworkIdle is reached.", FromProcedureName
                Exit Function
            End If
        End If

        ' タイムアウトチェック
        If (Timer - startTime) > TimeoutSec Then
            printMsg WARN_, "Timeout: Failed to reach NetworkIdle state within " & TimeoutSec & " seconds.", FromProcedureName
            WaitForNetworkIdle = False
            Exit Function
        End If

        sleep 0.05
    Loop
End Function

'***************************************************************************************************
'* 機能　　：指定したライフサイクルイベント (load, domcontentloaded) が発火するまで待機します
'***************************************************************************************************
Public Function WaitForLoadState(Optional state As String = "load", Optional TimeoutSec As Double = 30) As Boolean
    Const FromProcedureName As String = ThisClassName & ".WaitForLoadState"
    Dim startTime As Double: startTime = Timer
    Dim isLoaded As Boolean
    
    Do
        TakeEvents
        
        If LCase(state) = "load" Then
            isLoaded = m_LoadFired
        ElseIf LCase(state) = "domcontentloaded" Then
            isLoaded = m_DOMContentLoaded
        End If
        
        If isLoaded Then
            WaitForLoadState = True
            printMsg Trace_, "Lifecycle event '" & state & "' is reached.", FromProcedureName
            Exit Function
        End If

        If (Timer - startTime) > TimeoutSec Then
            printMsg WARN_, "Timeout: Failed to reach lifecycle state '" & state & "' within " & TimeoutSec & " seconds.", FromProcedureName
            WaitForLoadState = False
            Exit Function
        End If

        sleep 0.05
    Loop
End Function

'***************************************************************************************************
'* 機能　　：DOMスナップショットのサイズが指定回数連続して変化しなくなるまで待機します
'***************************************************************************************************
Public Function WaitForDOMStable(Optional TimeoutSec As Double = 10, Optional intervalMs As Long = 200, Optional requiredStableCount As Long = 3) As Boolean
    Const FromProcedureName As String = ThisClassName & ".WaitForDOMStable"
    Dim startTime As Double: startTime = Timer
    m_DOMStableCount = 0

    Do
        ' 現在の DOM サイズ (文字数) を取得。JSエラー時は "0" とする。
        Dim currentSnapshot As String
        currentSnapshot = CStr(ex_CDPBrowser.jsEval("document.body ? document.body.innerHTML.length : 0", IFEXCEPTION:="0", dbgMsg:=False))

        ' 変化チェック
        If currentSnapshot = m_LastDOMSnapshot And currentSnapshot <> "0" Then
            m_DOMStableCount = m_DOMStableCount + 1
        Else
            m_DOMStableCount = 0
        End If
        m_LastDOMSnapshot = currentSnapshot

        ' 安定判定
        If m_DOMStableCount >= requiredStableCount Then
            WaitForDOMStable = True
            printMsg Trace_, "DOM is stable.", FromProcedureName
            Exit Function
        End If

        ' タイムアウトチェック
        If (Timer - startTime) > TimeoutSec Then
            printMsg WARN_, "Timeout: DOM did not stabilize within " & TimeoutSec & " seconds.", FromProcedureName
            WaitForDOMStable = False
            Exit Function
        End If

        sleep intervalMs / 1000
    Loop
End Function

'***************************************************************************************************
'* 機能　　：SPAの準備が完了するまで (DOMContentLoaded + NetworkIdle) 待機するおまかせ機能です
'***************************************************************************************************
Public Function WaitForSPAReady(Optional TimeoutSec As Double = 30) As Boolean
    Const FromProcedureName As String = ThisClassName & ".WaitForSPAReady"
    Dim startTime As Double: startTime = Timer

    printMsg Trace_, "Waiting for SPA ready (1/2: DOMContentLoaded)...", FromProcedureName
    If Not WaitForLoadState("domcontentloaded", TimeoutSec) Then
        WaitForSPAReady = False
        Exit Function
    End If

    ' 残り時間の計算
    Dim remaining As Double: remaining = TimeoutSec - (Timer - startTime)
    If remaining <= 0 Then
        WaitForSPAReady = False
        Exit Function
    End If

    printMsg Trace_, "Waiting for SPA ready (2/2: NetworkIdle)...", FromProcedureName
    If Not WaitForNetworkIdle(remaining) Then
        WaitForSPAReady = False
        Exit Function
    End If

    printMsg Trace_, "SPA is ready.", FromProcedureName
    WaitForSPAReady = True
End Function



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
Private Function invokeMethod(methodName As String, Optional params As Object, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Object
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
    On Error Resume Next
    With ShSetting01_StartBrowser
        seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value    'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
    End With
    On Error GoTo 0
End Sub
