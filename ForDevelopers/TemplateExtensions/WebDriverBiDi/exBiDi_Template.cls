VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exBiDi_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           `WebDriverBiDiMode.cls`に対する機能拡張を提供するテンプレートクラスファイルです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_BiDiMode  As WebDriverBiDiMode
Attribute ex_BiDiMode.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName     As String = "exBiDi_Template"   'ログのトレース用

'ステート管理
Private m_LogID                 As String   'ログ識別ID
Private EnableBrowserEvents     As Boolean  'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private m_context               As String   '特定タブ向けに処理したい場合などに



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
'---------------------------------------------------------------------------------------------------
'* 注意事項：特定の`Context`だけ処理を通す場合はその判定処理を別途、case内に記述して下さい。
'            なお、CDP版と違って、イベント内容によって`Context`の記載情報位置が異なるため、イベントmethod名ごとにcontext判定というフローにして下さい
'***************************************************************************************************
Private Sub ex_BiDiMode_ExEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_BiDiMode_ExEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. 処理したい対象の非同期イベント名を、各Case に記述してください。
    Select Case methodName
        Case "○○.△△"
            '3. ↓ここから、あなたの非同期Aの処理イメージをコードに落とし込む↓
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName


        Case "□□.◇◇"
            '3. ↓ここから、あなたの非同期Bの処理イメージをコードに落とし込む↓
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName


        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName
            Exit Sub

    End Select

    '基本的には、`Case`内完結が望ましいですが、必要に応じて、下準備や後始末要因で、非同期イベントの処理前後であっても、コードを記述していただいて構いません。

End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分編 ■■■
'***************************************************************************************************
'* 機能　　：メインとなる機能拡張部分の1つです。あなたのイメージをコードに落とし込む場所の1つです
'***************************************************************************************************
Public Sub Main01(hoge1 As String, hoge2 As String)
    Const FromProcedureName As String = ThisClassName & ".Main01"   'Logのトレース用に、プロシージャ名を定数で用意しておくと便利です


    '↓ここから、独自の機能拡張処理イメージをコードに落とし込む↓



End Sub



'***************************************************************************************************
'                           ■■■ `WebDriverBiDiMode`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：親クラスへの直通ラッパーです。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：下記は一例です。待機は`CDPHelpers.Sleep` / `CDPHelpers.TimerCounter`を直接使ってください。必要に応じて削除や追加も可能です
'* 注意事項：各種メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておき、Private化すること
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Const LogFileExtension  As String = ".bidi.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, CDPHelpers.currentLog.logPath
End Sub

Private Function ExecuteBiDi(methodName As String, Optional params As Dictionary, Optional StopError As Boolean = True) As BiDiCDPJson
    Set ExecuteBiDi = ex_BiDiMode.ExecuteBiDi(methodName, params, StopError)
End Function

'非同期イベント発火(RaiseEvent)に必要です
Private Sub TakeEvents(Optional destruction As Boolean)
    ex_BiDiMode.TakeEvents , destruction
End Sub



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：BiDi制御のベースとなる`WithEvents`を継承させます
'---------------------------------------------------------------------------------------------------
'* 注意事項：・BiDiの場合、CDPと違いイベント情報にタブ宛て/ブラウザ宛てという概念がありません。
'              そのため、`WithEvents`継承は常に`As WebDriverBiDiMode`となります。
'            ・状況に応じて、どっちかを選んでください
'***************************************************************************************************
Friend Sub Init(Inheritance As WebDriverBiDiMode, Optional contextID As String)
    '必要なBiDiオブジェクトを継承
    Set ex_BiDiMode = Inheritance

    'コンテキスト情報
    m_context = contextID
End Sub

Friend Sub Init(Inheritance As WebDriverBiDiContext)
    '必要なBiDiオブジェクトを継承
    Set ex_BiDiMode = Inheritance.ThisWebDriverBiDiMode

    'コンテキスト情報
    m_context = Inheritance.context
End Sub

'***************************************************************************************************
'* 機能　　：`LogID`を採番します
'***************************************************************************************************
Private Sub Class_Initialize()
    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    m_LogID = "XBID" & Format(Rnd * 1000, "000")
End Sub
