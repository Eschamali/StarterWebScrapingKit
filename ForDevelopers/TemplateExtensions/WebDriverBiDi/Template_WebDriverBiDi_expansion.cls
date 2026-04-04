VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Template_WebDriverBiDi_expansion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           `WebDriverBiDiCore.cls`に対する機能拡張を提供するテンプレートクラスファイルです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_WebDriverBiDi As WebDriverBiDiCore

'内部使用用
Private Const ThisClassName         As String = "Template_WebDriverBiDi_expansion"             'ログのトレース用

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
Private Sub ex_WebDriverBiDi_detectionBiDiEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_WebDriverBiDi_detectionBiDiEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. 対象のイベント以外も即抜け
    '処理したい対象の非同期イベント名を、各Case に記述してください。
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

End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分編 ■■■
'***************************************************************************************************
'* 機能　　：メインとなる機能拡張部分の1つです。あなたのイメージをコードに落とし込む場所の1つです
'***************************************************************************************************





'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.invokeMethod`といった記述を防ぐ措置となります。
'---------------------------------------------------------------------------------------------------
'* 注意事項：メインの`WebDriverBiDiCore`メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておくことをお勧めします
'***************************************************************************************************
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_WebDriverBiDi.sleep seconds
End Sub

Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_WebDriverBiDi.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set invokeMethod = ex_WebDriverBiDi.invokeMethod(methodName, params, StopError)
End Function



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：WebDriverBiDi制御のベースとなる`WebDriverBiDiCore.cls`を継承させます
'***************************************************************************************************
Friend Sub Init(Inheritance As WebDriverBiDiCore)
    '継承
    Set ex_WebDriverBiDi = Inheritance
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
