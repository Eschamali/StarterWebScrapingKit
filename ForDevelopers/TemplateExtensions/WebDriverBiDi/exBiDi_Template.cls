VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exBiDi_Template.cls"
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
Private WithEvents ex_BiDiCore  As WebDriverBiDiCore
Attribute ex_BiDiCore.VB_VarHelpID = -1
Private ex_BiDiContext          As WebDriverBiDiContext
Attribute ex_BiDiContext.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "exBiDi_Template.cls"             'ログのトレース用

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'responseをデバッグログに出すかの調整
Private EnableBrowserEvents         As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private logFolderPath              As String       'ログファイルの保存フォルダパス



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
'* 注意事項：`ex_BiDiContext.cls`から受け取った`Context`だけ処理を通す場合はその判定処理を別途、case内に記述して下さい。
'            なお、CDP版と違って、イベント内容によって`Context`の記載情報位置が異なるため、イベントmethod名ごとにcontext判定というフローにして下さい
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
'                           ■■■ `WebDriverBiDi`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.ExecuteBiDi`といった記述を防ぐ措置となります。
'---------------------------------------------------------------------------------------------------
'* 注意事項：・前後の処理等も出来ますが基本は、1行だけに留めておくことをお勧めします
'            ・各詳細は、`WebDriverBiDi`の方をご覧ください
'***************************************************************************************************
'負荷を抑えるときにどうぞ
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_BiDiCore.sleep seconds
End Sub

'ログメッセージを出すときにどうぞ
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Const LogFileExtension         As String = ".bidi.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, , isHeader, doRaiseError, logFolderPath
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
    Set ex_BiDiCore = Inheritance.InheritanceWebDriverBiDiCore
    
    'コンテキスト情報
    m_context = Inheritance.context
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
        seeRawSendMsgDbg = .UseRangeID(8, FromProcedureName)    'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
        logFolderPath = .UseRangeID(9, FromProcedureName)          'ログファイルの保存フォルダパス
    End With
End Sub
