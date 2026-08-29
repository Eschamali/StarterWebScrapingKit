VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_TemplateWithSafeTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'        `CDPBrowser/CDPContext.cls`に対する機能拡張を提供するテンプレートクラスファイルです
'   このクラスは、`VBA-SafeTimer`機能を足したものです。より高度な連携機能を実装する際にこのテンプレートは役立つでしょう
'                      https://github.com/cristianbuse/VBA-SafeTimer
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'ブラウザ制御オブジェクト
Private WithEvents ex_CDPCore       As CDPCore      'イベント受け取りオブジェクト
Attribute ex_CDPCore.VB_VarHelpID = -1
Private ex_CDPBrowser               As CDPBrowser   '継承元のコンテキスト(タブ)
Private ex_CDPContext               As CDPContext   '継承元のブラウザ

'定期実行オブジェクト
Private WithEvents st               As SafeTimer
Attribute st.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "exCDP_TemplateWithSafeTimer"

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'ログ表示管理
Private EnableBrowserEvents         As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private CurrentSessionId            As String       '`CDPContext.cls`継承時のSesssionID
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
'* 機能　　：タブ向けの非同期イベントを処理したいときにご利用ください。あなたのイメージをコードに落とし込む場所の1つです
'***************************************************************************************************
Private Sub ex_CDPCore_CDPContextEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPCore_CDPContextEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. このコンテキスト(タブ)クラスで設定中の`SessionId`以外も即抜け
    If CurrentSessionId <> sessionID Then Exit Sub

    '3. 処理したい対象の非同期イベント名を、各Case に記述してください。
    Select Case methodName
        Case "○○.△△"
            '4. ↓ここから、あなたの非同期Aの処理イメージをコードに落とし込む↓
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & RawJson, FromProcedureName


        Case "□□.◇◇"
            '4. ↓ここから、あなたの非同期Bの処理イメージをコードに落とし込む↓
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & RawJson, FromProcedureName


        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName
            Exit Sub

    End Select

    '基本的には、`Case`内完結が望ましいですが、必要に応じて、下準備や後始末要因で、非同期イベントの処理前後であっても、コードを記述していただいて構いません。

End Sub

'***************************************************************************************************
'* 機能　　：ブラウザ単位での非同期イベントを処理したいときにご利用ください。あなたのイメージをコードに落とし込む場所の1つです
'***************************************************************************************************
Private Sub ex_CDPCore_CDPBrowserEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPCore_CDPBrowserEvent"


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
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & RawJson, FromProcedureName


        Case "□□.◇◇"
            '3. ↓ここから、あなたの非同期Bの処理イメージをコードに落とし込む↓
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & RawJson, FromProcedureName


        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName
            Exit Sub

    End Select

    '基本的には、`Case`内完結が望ましいですが、必要に応じて、下準備や後始末要因で、非同期イベントの処理前後であっても、コードを記述していただいて構いません。

End Sub



'***************************************************************************************************
'                               ■■■ VBA-SafeTimer 機能 ■■■
'***************************************************************************************************
'* 機能　　：継承したClassオブジェト内にあるパイプを覗き、非同期イベントを発火(RaiseEvent detectionCDPEvent)させます
'---------------------------------------------------------------------------------------------------
'* 注意事項：`TakeEvents`プロシージャ内にて`DoEvents`が呼ばれると、`SetTimer`のイベント処理の仕様で、このプロシージャが再度呼ばれスタックが積み重なります。
'           スタック領域不足を防ぐため常に1スタック分にすること
'***************************************************************************************************
Private Sub st_TimerCall(ByVal wTime As Long, varData As Variant)
    '1. 既に`TakeEvents`プロシージャに入ってる場合は、多重介入によるスタック領域不足を防ぐため即抜け
    Static isInnerTakeEvents As Boolean
    if isInnerTakeEvents then exit Sub

    '2. フラグを`True`にして、多重介入を阻止し、イベント取り出し処理へ
    isInnerTakeEvents = True
    TakeEvents

    '3. 終わったら、元に戻す
    isInnerTakeEvents = False
End Sub

'***************************************************************************************************
'* 機能　　：WindowsAPI - SetTimer による定期的なブラウザ非同期イベントの監視を開始します
'---------------------------------------------------------------------------------------------------
'* 引数    ：msec       イベント監視間隔(ms) ※デフォルトは50ms
'            varData    `st_TimerCall`に渡したい追加データ  ※テンプレートでは未使用です
'***************************************************************************************************
Public Sub StartAsyncTakeEvents(Optional msec As Long = 50, Optional addData As Variant)
    Set st = New SafeTimer
    st.StartTimer msec, addData
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
'                           ■■■ `CDPContext,CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：親クラスへの直通ラッパーです。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：下記は一例です。待機は`CDPHelpers.Sleep` / `CDPHelpers.TimerCounter`を直接使ってください。必要に応じて削除や追加も可能です
'* 注意事項：各種メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておき、Private化すること
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, , isHeader, doRaiseError, logFolderPath
End Sub

Private Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    ex_CDPCore.TakeEvents StopPipeError, destruction
End Sub



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`CDPCore.cls`等を継承させます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：CDP制御に関する全てのオブジェクトをここで継承させ、タブのSesssionIDも適用させてください
'* 注意事項：・タブの場合は、「CDPContext,CDPBrowser,CDPCore」の3つ、ブラウザの場合は、「CDPBrowser,CDPCore」の2つとなります
'           ・SessionId は、タブで継承する際に使用します
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    '必要なCDPオブジェクトを継承
    Set ex_CDPContext = Inheritance
    Set ex_CDPBrowser = Inheritance.InheritanceCDPBrowser
    Set ex_CDPCore = Inheritance.InheritanceCDPCore

    'SessionIdを取得
    CurrentSessionId = ex_CDPContext.CurrentSessionId
End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    seeRawSendMsgDbg = ShSetting01_StartBrowser.UseRangeID(8, FromProcedureName)    'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
    logFolderPath = ShSetting01_StartBrowser.UseRangeID(9, FromProcedureName)      'ログファイルの保存フォルダパス
End Sub
