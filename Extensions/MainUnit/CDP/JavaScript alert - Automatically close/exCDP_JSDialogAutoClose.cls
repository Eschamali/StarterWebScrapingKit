VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_JSDialogAutoClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
' CDPContext 拡張 - JavaScript ダイアログ（alert / confirm / prompt 等）の自動応答
'
' Page.javascriptDialogOpening を detectionCDPEvent で受け、Page.handleJavaScriptDialog を実行します。
' ・DefaultAccept … confirm / prompt / beforeunload で「はい相当」か「いいえ相当」か（alert は常に OK）
' ・DefaultPromptText … prompt で accept するときに送る文字列。空なら CDP の defaultPrompt を使用
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


'内部使用用
Private Const ThisClassName         As String = "exCDP_JSDialogAutoClose"

'ステート管理
Private seeRawSendMsgDbg            As Boolean
Private EnableBrowserEvents         As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private CurrentSessionId            As String       '継承時のSesssionID

' 既定: はい=OK 系 / prompt 用の入力文字列
Private Type JavaScriptDialog
    Accept      As Boolean
    promptText  As String
End Type
Private JavaScriptDialogClose As JavaScriptDialog



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
'* 機能　　：JavaScrpt-alert関数のデフォルトボタン操作の設定です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：confirm / prompt / beforeunload で True=はい・OK、False=いいえ・キャンセル（alert は常に True で閉じる）
'***************************************************************************************************
Public Property Get DefaultAccept() As Boolean
    DefaultAccept = JavaScriptDialogClose.Accept
End Property
Public Property Let DefaultAccept(ByVal v As Boolean)
    JavaScriptDialogClose.Accept = v
End Property

'***************************************************************************************************
'* 機能　　：JavaScrpt-alert関数のデフォルト入力値の設定です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：prompt かつ accept するときに Page.handleJavaScriptDialog の promptText に使う
'***************************************************************************************************
Public Property Get DefaultPromptText() As String
    DefaultPromptText = JavaScriptDialogClose.promptText
End Property
Public Property Let DefaultPromptText(ByVal v As String)
    JavaScriptDialogClose.promptText = v
End Property



'***************************************************************************************************
'                               ■■■ メインとなる機能拡張部分 ■■■
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

    '3. 対象のイベント以外も即抜け
    '処理したい対象の非同期イベント名を、各Case に記述してください。
    Select Case methodName
        Case "Page.javascriptDialogOpening"
            '4. デバッグログ用に、Jsonパース処理
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & RawJson, FromProcedureName
            If seeRawSendMsgDbg Then
                Dim DialogInfos As BiDiCDPJson: Set DialogInfos = BiDiCDPJson.Parse(RawJson).NodeKey("params")
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "javascriptDialogOpening information", FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "url           : " & DialogInfos.StringKey("url"), FromProcedureName
                printMsg Debug_, "message       : " & DialogInfos.StringKey("message"), FromProcedureName
                printMsg Debug_, "type          : " & DialogInfos.StringKey("type"), FromProcedureName
                printMsg Debug_, "defaultPrompt : " & DialogInfos.StringKey("defaultPrompt"), FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
            End If

            '5. ダイアログを閉じる処理を発動
            With JavaScriptDialogClose
                ex_CDPContext.handleDialog .Accept, .promptText
            End With

        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName

    End Select

End Sub



'***************************************************************************************************
'                           ■■■ `CDPContext,CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.ExecuteCDP`といった記述を防ぐ措置となります。
'            メインの`CDPBrowser`メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておくこと
'***************************************************************************************************
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_CDPBrowser.sleep seconds
End Sub

Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPContext.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function ExecuteCDP(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set ExecuteCDP = ex_CDPContext.ExecuteCDP(methodName, params, StopError)
End Function



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`CDPCore.cls`を継承させます
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    '必要なCDPオブジェクトを継承
    Set ex_CDPContext = Inheritance
    Set ex_CDPBrowser = Inheritance.InheritanceCDPBrowser
    Set ex_CDPCore = Inheritance.InheritanceCDPCore

    'SessionIdを取得
    CurrentSessionId = ex_CDPContext.CurrentSessionId

    'Pageドメインの有効化
    ex_CDPContext.pageEnable
End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    'ダイアログへのデフォルト値を設定
    JavaScriptDialogClose.Accept = True

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    seeRawSendMsgDbg = ShSetting01_StartBrowser.UseRangeID(8, FromProcedureName)   'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
End Sub



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
'* 機能　　：拡張機能との連携を切る際に記述します
'***************************************************************************************************
Private Sub Class_Terminate()
    'Pageドメインの無効化
    ex_CDPContext.pageEnable False
End Sub
