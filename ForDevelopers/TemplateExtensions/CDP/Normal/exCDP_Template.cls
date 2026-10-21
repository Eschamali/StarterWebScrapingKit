VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'        `CDPBrowser/CDPContext.cls`に対する機能拡張を提供するテンプレートクラスファイルです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'ブラウザ制御オブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser   '継承元のコンテキスト(タブ)
Private WithEvents ex_CDPContext    As CDPContext   '継承元のブラウザ

'内部使用用
Private Const ThisClassName         As String = "exCDP_Template"

'ステート管理
Private m_LogID                     As String   'ログ識別ID
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
'* 機能　　：タブ向けの非同期イベントを処理したいときにご利用ください。あなたのイメージをコードに落とし込む場所の1つです
'***************************************************************************************************
Private Sub ex_CDPContext_ExEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPContext_ExEvent"


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
'* 機能　　：ブラウザ単位での非同期イベントを処理したいときにご利用ください。あなたのイメージをコードに落とし込む場所の1つです
'***************************************************************************************************
Private Sub ex_CDPBrowser_ExEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPBrowser_ExEvent"


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
    Const LogFileExtension  As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, CDPHelpers.currentLog.logPath
End Sub

Private Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    '※ブラウザ向けコマンドが必要な場合にのみ、このようなヘルパーをもうけること
    ex_CDPContext.ThisCDPBrowser.TakeEvents StopPipeError, destruction
End Sub



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`WithEvents`を継承させます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：CDP制御に関するオブジェクトをここで継承さます
'* 注意事項：・タブの場合は、「CDPContext」、ブラウザの場合は、「CDPBrowser」となります
'            ・`CDPContext`だけであっても一応`.ThisCDPBrowser`で同時に取れますが、なるべくスコープ範囲を守るようにしましょう
'              タブで完結するなら､「CDPContext」だけといった感じです。
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    '必要なCDPオブジェクトを継承
    Set ex_CDPContext = Inheritance
    Set ex_CDPBrowser = Inheritance.ThisCDPBrowser  '※ブラウザ向けイベントが必要な場合のみ使用
End Sub

'***************************************************************************************************
'* 機能　　：`LogID`を採番します
'***************************************************************************************************
Private Sub Class_Initialize()
    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    m_LogID = "XTID" & Format(Rnd * 1000, "000")
End Sub
