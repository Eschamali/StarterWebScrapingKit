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
Private ex_CDPContext               As CDPContext   '継承元のコンテキスト(タブ)


'内部使用用
Private Const ThisClassName         As String = "exCDP_JSDialogAutoClose"

'基本ステート管理
Private currentLog                 As StateLog   'ログ設定
Private EnableBrowserEvents         As Boolean      'VBA側のイベント処理を有効/無効化のスイッチング
Private CurrentSessionID            As String       '継承時のSesssionID

'最後の記録用
Private LastJavascriptDialogOpening As String
Private LastJavascriptDialogClosed  As String

'既定: はい=OK 系 / prompt 用の入力文字列設定
Private Type JavaScriptDialog
    Accept      As Boolean
    promptText  As String
End Type
Private JavaScriptDialogClose As JavaScriptDialog



'***************************************************************************************************
'                            ■■■ スコープ調整プロパティ ■■■
'***************************************************************************************************
'* 機能　　：非同期イベント処理を行うかのスイッチングです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：スコープ範囲の調整にどうぞ。名称は、「Application.EnableEvents」から取りました
'***************************************************************************************************
Public Property Let EnableEvents(v As Boolean)
    '1. VBA側のフィルター設定
    EnableBrowserEvents = v

    '2. ブラウザ側からも調整
    ex_CDPContext.pageEnable v
End Property



'***************************************************************************************************
'                               ■■■ 設定プロパティ ■■■
'***************************************************************************************************
'* 機能　　：JavaScrpt-alert関数のデフォルトボタン操作の設定です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：confirm / prompt / beforeunload で True=はい・OK、False=いいえ・キャンセル（alert は常に True で閉じる）
'***************************************************************************************************
Public Property Let DefaultAccept(v As Boolean)
    JavaScriptDialogClose.Accept = v
End Property

'***************************************************************************************************
'* 機能　　：JavaScrpt-alert関数のデフォルト入力値の設定です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：prompt かつ accept するときに Page.handleJavaScriptDialog の promptText に使う
'***************************************************************************************************
Public Property Let DefaultPromptText(v As String)
    JavaScriptDialogClose.promptText = v
End Property



'***************************************************************************************************
'                               ■■■ 確認プロパティ ■■■
'***************************************************************************************************
'* 機能　　：各JavascriptDialog系の最後の情報を拝見します。
'---------------------------------------------------------------------------------------------------
'* 注意事項：中身は、文字列Jsonのままにしてます
'***************************************************************************************************
Public Property Get ViewLastJavascriptDialogOpening() As String
    ViewLastJavascriptDialogOpening = LastJavascriptDialogOpening
End Property
Public Property Get ViewLastJavascriptDialogClosed() As String
    ViewLastJavascriptDialogClosed = LastJavascriptDialogClosed
End Property



'***************************************************************************************************
'                      ■■■ メインとなる機能拡張部分 - イベント処理 ■■■
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
    If CurrentSessionID <> sessionID Then Exit Sub

    '3. 対象のイベント以外も即抜け
    Select Case methodName
        Dim DialogInfos As BiDiCDPJson

        Case "Page.javascriptDialogOpening"
            '4. デバッグログ用に、Jsonパース処理
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & RawJson, FromProcedureName
            If currentLog.seeRawSendMsgDbg Then
                Set DialogInfos = BiDiCDPJson.Parse(RawJson).NodeKey("params")

                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "javascriptDialogOpening information", FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "url               : " & DialogInfos.StringKey("url"), FromProcedureName
                printMsg Debug_, "message           : " & DialogInfos.StringKey("message"), FromProcedureName
                printMsg Debug_, "type              : " & DialogInfos.StringKey("type"), FromProcedureName
                printMsg Debug_, "hasBrowserHandler : " & DialogInfos.BoolKey("hasBrowserHandler"), FromProcedureName
                printMsg Debug_, "defaultPrompt     : " & DialogInfos.StringKey("defaultPrompt"), FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
            End If

            '5. ダイアログを閉じる処理を発動
            With JavaScriptDialogClose
                ex_CDPContext.handleDialog .Accept, .promptText
            End With

            '6. 最後の履歴として残しておく
            LastJavascriptDialogOpening = RawJson

        Case "Page.javascriptDialogClosed"
            '4. デバッグログ用に、Jsonパース処理
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & RawJson, FromProcedureName
            If currentLog.seeRawSendMsgDbg Then
                Set DialogInfos = BiDiCDPJson.Parse(RawJson).NodeKey("params")

                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "javascriptDialogClosed information", FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "result    : " & DialogInfos.BoolKey("result"), FromProcedureName
                printMsg Debug_, "userInput : " & DialogInfos.StringKey("userInput"), FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
            End If

            '5. 最後の履歴として残しておく
            LastJavascriptDialogClosed = RawJson

        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName

    End Select

End Sub



'***************************************************************************************************
'                           ■■■ `CDPContext`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.ExecuteCDP`といった記述を防ぐ措置となります。
'            メインの`CDPBrowser`メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておくこと
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Const LogFileExtension As String = ".cdp.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`CDPCore.cls`等を継承させます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：この拡張機能Classを動かすのに必要なオブジェクトや情報を渡し、イベント処理機能を有効化させます
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    '必要なCDPオブジェクトを継承
    Set ex_CDPContext = Inheritance
    Set ex_CDPCore = Inheritance.ThisCDPCore
    CDPHelpers.InheritLogSettings currentLog, Inheritance.ThisCDPCore.LogSettings

    'SessionIdを取得
    CurrentSessionID = ex_CDPContext.CurrentSessionID

    'Pageドメインの有効化
    EnableEvents = True
End Sub

'***************************************************************************************************
'* 機能　　：初期設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    'ダイアログへのデフォルト値を設定
    JavaScriptDialogClose.Accept = True

    currentLog.LogID = "JDID" & Format(Rnd * 1000, "000")
End Sub



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
'* 機能　　：拡張機能との連携を切る際に記述します
'***************************************************************************************************
Private Sub Class_Terminate()
    'Pageドメインの無効化
    EnableEvents = False
End Sub
