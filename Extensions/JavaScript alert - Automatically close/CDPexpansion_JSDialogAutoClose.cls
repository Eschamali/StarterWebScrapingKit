VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPexpansion_JSDialogAutoClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
' CDPBrowser 拡張 - JavaScript ダイアログ（alert / confirm / prompt 等）の自動応答
'
' Page.javascriptDialogOpening を detectionCDPEvent で受け、Page.handleJavaScriptDialog を実行します。
' ・DefaultAccept … confirm / prompt / beforeunload で「はい相当」か「いいえ相当」か（alert は常に OK）
' ・DefaultPromptText … prompt で accept するときに送る文字列。空なら CDP の defaultPrompt を使用
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "CDPexpansion_JSDialogAutoClose"
Private mJson                       As New WebJsonConverter

'ステート管理
Private seeRawSendMsgDbg            As Boolean

Public EnableBrowserEvents          As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング

' 既定: はい=OK 系 / prompt 用の入力文字列
Private mDefaultAccept              As Boolean
Private mDefaultPromptText          As String



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
    DefaultAccept = mDefaultAccept
End Property
Public Property Let DefaultAccept(ByVal v As Boolean)
    mDefaultAccept = v
End Property

'***************************************************************************************************
'* 機能　　：JavaScrpt-alert関数のデフォルト入力値の設定です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：prompt かつ accept するときに Page.handleJavaScriptDialog の promptText に使う
'***************************************************************************************************
Public Property Get DefaultPromptText() As String
    DefaultPromptText = mDefaultPromptText
End Property
Public Property Let DefaultPromptText(ByVal v As String)
    mDefaultPromptText = v
End Property



'***************************************************************************************************
'                               ■■■ メインとなる機能拡張部分 ■■■
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

    '2. 対象のイベント以外も即抜け
    '処理したい対象の非同期イベント名を、各Case に記述してください。
    Select Case methodName
        Case "Page.javascriptDialogOpening"
            '3. デバッグログ用に、Jsonパース処理
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName
            If seeRawSendMsgDbg Then
                Dim DialogInfos As Dictionary: Set DialogInfos = mJson.ParseJson(EventRawJson)("params")
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "javascriptDialogOpening information", FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "url           : " & DialogInfos("url"), FromProcedureName
                printMsg Debug_, "message       : " & DialogInfos("message"), FromProcedureName
                printMsg Debug_, "type          : " & DialogInfos("type"), FromProcedureName
                printMsg Debug_, "defaultPrompt : " & DialogInfos("defaultPrompt"), FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
            End If

            '4. ダイアログを閉じる処理を発動
            ex_CDPBrowser.handleDialog mDefaultAccept, mDefaultPromptText

        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName
            Exit Sub
    End Select

End Sub



'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.invokeMethod`といった記述を防ぐ措置となります。
'            メインの`CDPBrowser`メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておくこと
'***************************************************************************************************
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_CDPBrowser.sleep seconds
End Sub

Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set invokeMethod = ex_CDPBrowser.invokeMethod(methodName, params, alwaysBrowserContext, StopError)
End Function



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`CDPBrowser.cls`を継承させます
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPBrowser)
    '継承
    Set ex_CDPBrowser = Inheritance

    'Pageドメインの有効化
    ex_CDPBrowser.pageEnable
End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    'ダイアログへのデフォルト値を設定
    mDefaultAccept = True
    mDefaultPromptText = vbNullString

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value    'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
    End With
End Sub



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
'* 機能　　：拡張機能との連携を切る際に記述します
'***************************************************************************************************
Private Sub Class_Terminate()
    'Pageドメインの無効化
    ex_CDPBrowser.pageEnable False
End Sub
