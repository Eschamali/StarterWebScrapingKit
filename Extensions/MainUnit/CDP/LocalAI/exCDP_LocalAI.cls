VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_LocalAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           ブラウザのローカルAI(Prompt API)に対する機能拡張を提供するクラスファイルです
'   実装内容は全て、2026/4/15時点の「https://learn.microsoft.com/ja-jp/microsoft-edge/web-platform/prompt-api」を参考に記述してます
'           まだまだ、試験的扱いのためある日突然、動作しなくなることをご了承ください
'   現時点では、全てのサイトで動作することができません。「edge://」等と言った限られた場所のみ動作します
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部クラスオブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "exCDP_LocalAI"             'ログのトレース用

'`Runtime.bindingCalled`の`Name`定義用
Private Const BindingName01         As String = "sendDLprogressNotification"    'AIモデルデータのDL進捗用

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'responseをデバッグログに出すかの調整
Private EnableBrowserEvents         As Boolean      'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private AI_ObjectID                 As String       'AIへのセッションに対するObjectID
Private ProgressValue               As Double       'AIモデルデータの進捗値



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
'* 機能　　：AIモデルデータのDL進捗値を取得します
'***************************************************************************************************
Public Property Get DLProgressValue() As Double
    DLProgressValue = ProgressValue
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
'                               ■■■ Prompt API - 準備編 ■■■
'***************************************************************************************************
'* 機能　　：API が有効かどうかを確認する
'---------------------------------------------------------------------------------------------------
'* 返り値  ：利用可否論理値
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Prompt API を使用する前に、LanguageModel オブジェクトの存在をテストして API が有効になっていることを確認します
'***************************************************************************************************
Public Function IsPromptApiAvailable() As Boolean
    '1. 所定のJavaScriptコードを記述
    Dim JsCode As String
    JsCode = "!LanguageModel"

    '2. CDPのJavaScript実行メソッドを使用し、ブラウザ内で、JavaScriptコードを実行
    Dim tmp
    tmp = ex_CDPBrowser.jsEval(JsCode)

    '3. 結果を整理
    'そもそも、用意されてないか？
    If IsError(tmp) Then
        '「Uncaught ReferenceError: LanguageModel is not defined」と出るので、`False`とする
        IsPromptApiAvailable = False
    Else
        '機能があれば、普通の論理値なので、それを返却
        '※元のJSが、否定演算子を使用してるため、結果も反転しておく
        IsPromptApiAvailable = Not tmp
    End If

End Function

'***************************************************************************************************
'* 機能　　：モデルを使用できるかどうかを確認する
'---------------------------------------------------------------------------------------------------
'* 返り値  ：現在は、下記の4種類が返ります
'            ・"unavailable"
'            ・"downloadable"
'            ・"downloading"
'            ・"available"
'---------------------------------------------------------------------------------------------------
'* 詳細説明：言語モデルとモデル ランタイムが Microsoft Edge によってダウンロードされたかを確認します
'* 注意事項：これを行う前に`IsPromptApiAvailable`の確認をしないとエラー落ちします
'***************************************************************************************************
Public Function CheckAvailability() As String
    '1. 所定のJavaScriptコードを記述
    Dim JsCode As String
    JsCode = "(async () => { return await LanguageModel.availability(); })()"

    '2. CDPのJavaScript実行メソッドを使用し、ブラウザ内で、JavaScriptコードを実行し、結果を返却
    CheckAvailability = ex_CDPBrowser.jsEval(JsCode, awaitPromise:=True, StopException:=True)
End Function

'***************************************************************************************************
'* 機能　　：モデルをDLし、進捗表示機能を有効化します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`monitor`オプションを使用して、モデルのダウンロードの進行状況を追跡できます
'***************************************************************************************************
Public Function ModelDownloadProgress() As String
    With ex_CDPBrowser
        '1. ブラウザ→VBA への受け口を用意
        Dim params As New Dictionary
        params.Add "name", BindingName01
        .invokeMethod "Runtime.addBinding", params

        '2. JavaScript側でその窓口を叩くコードを準備
        Dim jsProgressCode As String
        jsProgressCode = "(function() { " & _
            "  LanguageModel.create({ " & _
            "    monitor: m => { " & _
            "      m.addEventListener('downloadprogress', event => { " & _
            "        const progress = (event.loaded / event.total) * 100; " & _
            "        if (typeof " & BindingName01 & " === 'function') { " & _
            "          " & BindingName01 & "(progress.toFixed(2)) " & _
            "        } " & _
            "      }); " & _
            "    } " & _
            "  }); " & _
            "  return 'DOWNLOAD_STARTED'; " & _
            "})()"

        '3. Run!
        ModelDownloadProgress = .jsEval(jsProgressCode, userGesture:=True, StopException:=True)
    End With
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
