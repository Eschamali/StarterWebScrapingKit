VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'                   Automating WebView2 Browsers with Chrome Dev Protocol API and VBA
'--------------------------------------------------------------------------------------------------------------
'                   WebView2 を Chrome Dev Protocol を駆使したブラウザ自動操作を担います
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                              ■■■ Windows API定義関連 ■■■
'***************************************************************************************************
Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)



'***************************************************************************************************
'                                   ■■■ 定数関連 ■■■
'***************************************************************************************************
Private Const brWebView2            As String = "msedgewebview2.exe"
Private Const WebView2WindowClass   As String = "Chrome_WidgetWin_1"    'same window class for WebView2
Private Const ThisClassName         As String = "WebView2Browser"       'ログのトレース用

Public Enum StartWebView2Mode
    UserFormModeless    'モードレスとして、UserFormを起動します
    UserFormModal       'モーダルとして、UserFormを起動します
    Background          'UserFormを出さず、内部としてWebView2を起動します(Excelの`Hwnd`を使用)
End Enum

Private Const ERR_TIMEOUT As Long = 900
Private Const ERR_CoreWebView2 As Long = 901
Private Const ERR_PROTOCOL As Long = 902
Private Const ERR_DictionaryStructure As Long = 903
Private Const ERR_UnregisteredProtocol As Long = 904
Private Const ERR_PolicyViolation As Long = 905
Private Const TIMEOUT_DEFAULT As Single = 30
Private Const TIMEOUT_NAVIGATE As Single = 45


'***************************************************************************************************
'                              ■■■ クラス内変数定義関連 ■■■
'***************************************************************************************************
Private jsConverter As New WebJsonConverter 'Excelで、JsonデータをDictionary(Collection)に変換できるすごいやつ
Private LogControl  As New Logger           'ログレベルの制御
Private core        As WebView2Core         'Must be initialized for CDP engine
Private ErrorDetail As New WinApiError      'エラーコードから、詳細を取得するやつ

Private brID                    As String
Private sessionID               As String
Private doLog                   As Boolean  'ログファイルにも、デバッグメッセージを残すか
Private logPath                 As String   'ログファイルの保存フォルダパス
Private seeRawSendMsgDbg        As Boolean  'set to True if needing to see the raw json msg received from the browser for debugging
Private errorDetails            As Dictionary      'CDP-Jsonコマンド実行時の最後のエラー情報


Private InternalResolutionWidth As Long
Private InternalResolutionHight As Long



'***************************************************************************************************
'                                   ■■■ 始動関連 ■■■
'***************************************************************************************************
'* 機能　　：初期化処理
'---------------------------------------------------------------------------------------------------
'* 返り値　：初期化成功可否論理値
'* 引数　　：userProfile                    UserDataDir名を指定します
'            StartMode                      UserFormModeless：モードレスとして、UserForm を起動します(規定値)
'                                           UserFormModal   ：モーダルとして、UserForm を起動します。この場合、この`Start`メソッド以降は処理が走りません。
'                                           Background      ：UserForm を出さず、内部でWebView2を起動した状態で、処理を進めます。Excel自体の`hwnd`を使用することで、実現してます。
'
'            FirstURL                       最初の起動時のURLを渡します。省略時は、「"about:blank"」とします。
'            addArgs                        WebView2に渡す追加引数を渡します
'---------------------------------------------------------------------------------------------------
'* 機能説明：WebView2 の初期化が完了するまで、担います
'* 注意事項：WebView2 の初期化などに失敗すると内部でErr.Raise`が発動され、続行されなくなります
'***************************************************************************************************
Public Sub start(userProfile As String, Optional StartMode As StartWebView2Mode, Optional FirstURL As String, Optional addArgs As String)
    Const FromProcedureName As String = ThisClassName & ".start"


    'WebView2準備の待機時間(s)
    Const ReadyWebView2Time As Long = 15

    '新規オブジェクトをセット
    Set core = New WebView2Core

    '起動モードを設定する
    Dim t As Single
    Select Case StartMode
        Case StartWebView2Mode.UserFormModeless
            '1. モードレスの場合は、処理は続行されます
            ' WebView2Core を UserForm 側に渡して初期化・表示してもらう
            printMsg info_, "モードレスウィンドウ上で、WebView2を起動します...", FromProcedureName
            With WebView2InExcelForm
                .AttachCore core, FirstURL, addArgs, userProfile
                .show vbModeless
            End With

            '2. WebView2 の内部処理が終わるまで待機
            '※UserForm側のイベント発火のタイミングを合わせるため、待機処理
            t = Timer
            Do While Not core.IsReady And Timer - t < ReadyWebView2Time
                core.ProcessMessages
            Loop

            '3. 時間内に準備できなかったら、停止
            If Not core.IsReady Then
                quit
                printMsg ERROR_, "WebView2 の初期化には成功しましたが時間内に、準備ができませんでした。", FromProcedureName, doRaiseError:=True
            Else
                printMsg info_, "WebView2 の使用準備が完了しました", FromProcedureName
            End If

        Case StartWebView2Mode.UserFormModal
            'モーダル表示の場合は仕様上、ウィンドウを閉じるまで、ここで止まります
            printMsg info_, "モーダルウィンドウ上で、WebView2を起動します...", FromProcedureName
            With WebView2InExcelForm
                .AttachCore core, FirstURL, addArgs, userProfile
                .show vbModal
            End With

        Case StartWebView2Mode.Background
            '1. WebView2 初期化
            printMsg info_, "バックグラウンドで、WebView2を起動します...", FromProcedureName
            If core.Initialize(Application.hWnd, 0, 0, InternalResolutionWidth, InternalResolutionHight, FirstURL, userProfile, addArgs) Then
                printMsg info_, "WebView2 の初期化、成功", FromProcedureName
            Else
                printMsg ERROR_, "WebView2 の初期化に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & core.LastErrorDescription, FromProcedureName, doRaiseError:=True
            End If

            '2. WebView2 の内部処理が終わるまで待機
            t = Timer
            Do While Not core.IsReady And Timer - t < ReadyWebView2Time
                core.ProcessMessages
            Loop

            '3. 時間内に準備できなかったら、停止
            If Not core.IsReady Then
                quit
                printMsg ERROR_, "WebView2 の初期化には成功しましたが時間内に、準備ができませんでした。", FromProcedureName, doRaiseError:=True
            Else
                printMsg info_, "WebView2 の使用準備が完了しました", FromProcedureName
            End If

        Case Else: printMsg ERROR_, "この起動モードは定義されてません", FromProcedureName, doRaiseError:=True
    End Select
End Sub



'***************************************************************************************************
'                              ■■■ CDP送信準備関連 ■■■
'***************************************************************************************************
'* 機能　　：CDPのMethod名と、parametersを渡して、低レベルにてブラウザ制御をします
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド の実行結果(Json-Dictionary 形式)。失敗時は、`Nothing`となります。
'* 引数　　：methodName             method名。"Network.getCookies"や"DOM.getDocument"というのを記入します
'            params                 Json-Dictionary形式で、KeyとItemを事前に定義したオブジェクト変数を指定します。
'            alwaysBrowserContext   `True` ：`sessionID`の有無問わず、常にブラウザ自身を対象したコマンド実行モードとなります
'                                   `False`：`sessionID`があれば、その`SessionID`に対するコマンド実行。なければ、ブラウザ自身を対象したコマンド実行となります。
'
'            StopError              `True`で、実行失敗時に`Err.Raise`で停止し、以降の処理を止めます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：This function enables advanced method invocation under CDP.
'            For a list of all methods available, see:
'               https://chromedevtools.github.io/devtools-protocol/tot/Browser/
'            Error message will be prompted automatically by invokeError in the Immediate Window.
'            The return object is Json String that has been converted to a dictionary object.
'            Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'
'* 更新履歴：・18/02/26 - SessionIDがあっても、BrowserContext優先にコマンド実行できるように改良。エラーで止めるかのスイッチングを導入
'            ・28/04/23 - Set this to Public as it is useful
'            ・20/03/23 - Added printMsg msgs & set to Private
'
'* 注意事項：・CDPの運営思想の都合上、多数の`Deprecated`,`Experimental`が存在します。
'              これらのタグのmethodを使用する際は、`StopError`を`True`にすることを推奨します。失敗時の以降の処理の予期せぬデータ破損などを防ぐためです。
'
'            ・`StopError`の規定値は、`True`なので、特別な理由がない限りは、省略推奨です
'            ・このメソッドは、コマンド結果が来るまで待機するメソッドです。送るコマンド内容や、サイトの挙動によっては、タイムアウトエラーになる場合があります。
'              その場合は、後述の結果を待たないタイプのメソッド`invokeMethodAsync`をご利用ください。
'***************************************************************************************************
Public Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Const FromProcedureName As String = ThisClassName & ".invokeMethod"


    Dim paramStr As String
    Dim results As Scripting.Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = jsConverter.ConvertToJson(params)

    '関数内部で、結果が返るまで待つようにします
    Set results = SendMessage(methodName, paramStr, alwaysBrowserContext, Async:=False)

   'Print error if there is
    If invokeError(results) Then
        'エラーによる`Err.Raise`で止めるか？
        If StopError Then
            'エラーになったmethod名と、結果生responseを表示して停止
            printMsg ERROR_, "Failed to invoke method """ & methodName & """" & vbCrLf & vbCrLf & "=== RawErrorResponse ===" & vbCrLf & jsConverter.ConvertToJson(results), FromProcedureName, doRaiseError:=True
        Else
            'エラーになったmethod名だけログ出して続行
            printMsg WARN_, "Failed to invoke method """ & methodName & """", FromProcedureName
        End If
    End If

   'Return result object
    If results.Exists("result") Then Set invokeMethod = results.Item("result")
    printMsg Debug_, "Invoked """ & methodName & """ with result: " & jsConverter.ConvertToJson(results), FromProcedureName

End Function



'***************************************************************************************************
'                           ■■■ 核となるプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDPコマンド送信のコア処理
'***************************************************************************************************
'* 機能　　：CDP にメッセージを送信し、Chromium ブラウザでメッセージが受信されているかどうかを確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：Async が、`False` → CDP-Jsonコマンド の実行結果(Json-Dictionary 形式)
'                      `True`  → 実行したコマンドIDと、WriteFile結果コードのペアーDictionary
'
'* 引数　　：methodName             CDP-Jsonコマンドのうち、メソッド名のみを指定します
'            paramsJsonString       WebView2に送るCDP-Jsonコマンドのうち、パラメーターJson部分(プレーンテキスト)
'            alwaysBrowserContext   `True` ：`sessionID`の有無問わず、常にブラウザ自身を対象したコマンド実行モードとなります
'                                   `False`：`sessionID`があれば、その`SessionID`に対するコマンド実行。なければ、ブラウザ自身を対象したコマンド実行となります。
'
'            InternalUseMessages    非同期イベントをクラス内部で後程、使用したいときに、`New Dictionry`として、渡して有効化します。
'            Async                  `True` ：コマンドの結果を待ちません。Do Loopフェーズには入りません。
'                                   `False`：コマンドの結果を待ちます。Do Loopフェーズに入り、そのIDが来るまで待ちます。この場合、非同期に実行した結果のID回収もついでに行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：引数-Async の値によって、Dictionaryの構造が変わる特殊なメソッドです。別途公開メソッドで、役割分担しております
'***************************************************************************************************
Private Function SendMessage(methodName As String, paramsJsonString As String, Optional alwaysBrowserContext As Boolean, Optional InternalUseMessages As Scripting.Dictionary, Optional Async As Boolean) As Scripting.Dictionary
    Const FromProcedureName As String = ThisClassName & ".SendMessage"


    Dim strRes As String
    printMsg Debug_, "sendCDP-Command：method = `" & methodName & "` , paramsJson = `" & paramsJsonString & "`", FromProcedureName

    'コマンド結果が来るまで待機するか？
    If Async Then
        '投げっぱなしにして、即次へ

    Else
        '内部で、コマンド結果のコールバックが来るまで待ち続けます
        strRes = core.CallDevToolsProtocolMethod(methodName, paramsJsonString)
        
        'ついでに、非同期イベント情報も回収しておく
        

        '結果を返却
        Set SendMessage = jsConverter.ParseJson(strRes)
    End If
End Function

'***************************************************************************************************
'* 機能　　：コマンドの結果にエラーが含まれているかチェックします
'---------------------------------------------------------------------------------------------------
'* 返り値　：エラー有無
'* 引数　　：ResultsObj    `invokeMethod`の生結果のDictionary
'***************************************************************************************************
Private Function invokeError(result As Scripting.Dictionary) As Boolean
    Const FromProcedureName As String = ThisClassName & ".invokeError"


    Dim ErrorCode As Long, ErrorMessage As String, ErrorData As Variant
    If result Is Nothing Then
        ErrorCode = -1 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        ErrorData = vbEmpty
        printMsg WARN_, "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage, FromProcedureName
        invokeError = True
        Set errorDetails = jsConverter.ParseJson("{""code"":" & ErrorCode & ",""message"":""" & ErrorMessage & """}")     'カスタムエラーメッセージJsonで返却
    ElseIf result.Exists("error") Then
        Set errorDetails = result("error")
        ErrorCode = errorDetails("code")
        ErrorMessage = errorDetails("message")
        printMsg WARN_, "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage, FromProcedureName
        If errorDetails.Exists("data") Then
            ErrorData = errorDetails("data")
            printMsg WARN_, "       " & ErrorData, FromProcedureName
        Else
            ErrorData = vbEmpty ' vbNullString? is errorDetails("data") a string or object?
        End If
        invokeError = True
    Else
        ErrorCode = 0
        ErrorMessage = vbNullString
        ErrorData = vbEmpty
    End If

End Function

'***************************************************************************************************
'* 機能　　：同期モードでの CDP-Json コマンド結果のうち、最後のエラー情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後のエラー情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」のエラー情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます。
'            ・適切な`invokeMethod`を実行したつもりなのに返り値が、`Nothing`となるのが、エラーの目印となるので、それでこのエラー情報を見に行ってください。
'            ・`invokeMethod`の`StopError`が`False`だと実質機能しません。
'            ・非同期によるあとで取り出したパターンとは、連動しません。
'***************************************************************************************************
Property Get LastCDPJsonError() As Dictionary
    Set LastCDPJsonError = errorDetails
End Property

'***************************************************************************************************
'* 機能　　：ログレベル付きのスマートなログをイミディエイトへ出力させるヘルパーメソッドです
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Dim logName As String, strFormattedMsg As String

    If isHeader Then strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-") _
    Else strFormattedMsg = " | " & brID & " | " & strMsg

   'Perform msg logging if needed
    If doLog Then
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".cdp.log"
        If Dir(logPath, vbDirectory) = "" Then MkDir logPath
    End If

   'Perform msg output to Immediate Window if needed
    With LogControl
        Select Case LogLevel_
            Case LogLevelName.Trace_: .LogTrace strFormattedMsg, From, logPath, logName
            Case LogLevelName.Debug_: .LogDebug strFormattedMsg, From, logPath, logName
            Case LogLevelName.info_: .LogInfo strFormattedMsg, From, logPath, logName
            Case LogLevelName.WARN_: .LogWarn strFormattedMsg, From, logPath, logName
            Case LogLevelName.ERROR_: .LogError strFormattedMsg, From, logPath, logName, Err.Number
            Case Else: 'ログ出力無効化
        End Select
    End With


    'WebView2自身や、非同期処理自体に問題や、タイムアウトになったと判断された場合は、このツールでの処理(CDP-Json)ができないので、エラーで停止させておく
    If doRaiseError Then Err.Raise ERR_CoreWebView2, From, Description:=strMsg

End Sub

'***************************************************************************************************
'* 機能　　：WebView2を終了させます
'***************************************************************************************************
Public Sub quit()
    Const FromProcedureName As String = ThisClassName & ".invokeError"


    printMsg info_, "WebView2を終了します...", FromProcedureName
    core.quit
    printMsg info_, "WebView2を終了しました", FromProcedureName
End Sub



Private Sub Class_Initialize()
    'Assign an objID for easy debug identification
    brID = "BRID" & Format(Rnd * 1000, "000")


    '------------------------- 設定シートの各セルから設定値を取得し、クラス内変数に適用します -------------------------
    With ShSetting01_StartBrowser
        '1. イミディエイトに関する設定を行う
        seeRawSendMsgDbg = .Range(.UseRangeName(8, "WebView2Browser.Class_Initialize")).value           'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？

        '2. パスが設定されてたら、ログファイル保存を有効化
        Dim LogFolderPath As String: LogFolderPath = .Range(.UseRangeName(9, "WebView2Browser.Class_Initialize")).value
        If LogFolderPath <> "" Then
            doLog = True                'ログファイルとして保存するか？
            logPath = LogFolderPath     '保存先のフォルダパスを設定
        End If

        '3. WebView2が持つ、内部解像度を設定します。
        '※この設定は、`StartWebView2Mode`が、`Background`のみ適用されます
        InternalResolutionWidth = .Range(.UseRangeName(13, "WebView2Browser.Class_Initialize")).value
        InternalResolutionHight = .Range(.UseRangeName(14, "WebView2Browser.Class_Initialize")).value
    End With
End Sub
