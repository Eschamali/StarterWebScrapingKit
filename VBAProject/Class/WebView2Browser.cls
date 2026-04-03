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

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" ( _
    ) As Long

Private Declare PtrSafe Function ProcessIdToSessionId Lib "kernel32" ( _
    ByVal dwProcessId As Long, _
    ByRef pSessionId As Long) As Long

Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nCmdShow As Long) As Boolean



'***************************************************************************************************
'                                   ■■■ 定数関連 ■■■
'***************************************************************************************************
Private Const brWebView2            As String = "msedgewebview2.exe"
Private Const WebView2WindowClass   As String = "Chrome_WidgetWin_1"    'same window class for WebView2
Private Const ThisClassName         As String = "WebView2Browser"       'ログのトレース用
Private Const CDPVer                As String = "3.1.0"                 '本家流用なので、そのままにしておく

Public Enum StartWebView2Mode
    UserFormModeless    'モードレスとして、UserFormを起動します
    UserFormModal       'モーダルとして、UserFormを起動します
    Background          'UserFormを出さず、内部としてWebView2を起動します(Excelの`Hwnd`を使用)
End Enum

Public Enum ReadyState      'Used for .wait method
    isLoading = 0           'equivalence of the browser's "loading" state
    isInteractive = 1       'equivalence of the browser's "interactive" state
    isComplete = 2          'equivalence of the browser's "complete" state
End Enum

Public Enum TargetgetTargetsType
    Page
    iFrame
End Enum

Public Enum edgeKioskType
    none            'キオスクモードなしの通常起動
    fullscreen      'キオスクモード デジタル/対話型サイネージ
    publicBrowsing  'キオスクモードのパブリック ブラウジング
End Enum

Public Enum WinState
    asHidden = 0            'Hides the window and activates another window.
    asNormal = 1            'Activates and displays a window. If the window is minimized or maximized, the system restores it to its original
    asMinimized = 2         'Activates the window and displays it as a minimized window.
    asMaximized = 3         'Activates the window and displays it as a maximized window.
    doShowNoActivate = 4    'Displays a window in its most recent size and position. This value is similar to asNormal, except that the window is not activated.
    doShow = 5              'Activates the window and displays it in its current size and position.
    doMinimize = 6          'Minimizes the specified window and activates the next top-level window in the Z order.
    doShowMinNoActivate = 7 'Displays the window as a minimized window. This value is similar to asMinimized, except the window is not activated.
    doShowNA = 8            'Displays the window in its current size and position. This value is similar to doShow, except that the window is not activated.
    doRestore = 9           'Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    doShowDefault = 10      'Sets the show state based on the WS_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
    doForceMin = 11         'Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread.
End Enum

Private Const ERR_TIMEOUT As Long = 900
Private Const ERR_CoreWebView2 As Long = 901
Private Const ERR_PROTOCOL As Long = 902
Private Const ERR_DictionaryStructure As Long = 903
Private Const ERR_UnregisteredProtocol As Long = 904
Private Const ERR_PolicyViolation As Long = 905
Private Const TIMEOUT_DEFAULT As Single = 30
Private Const TIMEOUT_NAVIGATE As Single = 45

' セッション情報
Private brURL   As String         ' タブの URL
Private brName  As String         ' タブのタイトル
Private brHWnd  As LongPtr        ' 親ウィンドウ（WebView2 の親 HWND）
Private brType  As String         ' ブラウザタイプ（例: "Edg/xxx"）
Private brVer   As String         ' ブラウザバージョン
Private brPath  As String         ' WebView2 実行ファイルパス
Private targetID As String        ' CDP 的な targetId
Private brWID As String           ' Useful for windows repositioning
Private brUser As String          ' Browser user profile
Private sessionID As String       ' 一応残しておく
Private lastExceptionDetailsJS  As Dictionary  'JavaScript実行時の最後のエラー情報



'***************************************************************************************************
'                              ■■■ クラス内変数定義関連 ■■■
'***************************************************************************************************
Private jsConverter As New WebJsonConverter 'Excelで、JsonデータをDictionary(Collection)に変換できるすごいやつ
Private LogControl  As New Logger           'ログレベルの制御
Private WithEvents core As WebView2Core     'Must be initialized for CDP engine
Attribute core.VB_VarHelpID = -1

Private ErrorDetail As New WinApiError      'エラーコードから、詳細を取得するやつ

Private brID                    As String
Private doLog                   As Boolean  'ログファイルにも、デバッグメッセージを残すか
Private logPath                 As String   'ログファイルの保存フォルダパス
Private seeRawSendMsgDbg        As Boolean  'set to True if needing to see the raw json msg received from the browser for debugging
Private errorDetails            As Dictionary      'CDP-Jsonコマンド実行時の最後のエラー情報


Private InternalResolutionWidth As Long
Private InternalResolutionHight As Long
Private lngLastID As Long              'Async呼び出し時の整理券ID（ローカル採番）



'非同期イベント通知（WebView2Core から CDP イベントを受け取り、そのまま伝播）
' EventRawJson … CDP の params に相当する JSON 文字列（WebView2 からの As-Is。フル Envelope ではない）
Public Event detectionCDPEvent(ByVal methodName As String, ByVal EventRawJson As String)



'***************************************************************************************************
'                                   ■■■ 初回処理関連 ■■■
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


    '作業フォルダの記録
    brUser = userProfile 'store the user profile name for cpid serialization
    Call cleanUpSessions

    'WebView2準備の待機時間(s)
    Const ReadyWebView2Time As Long = 15

    '新規オブジェクトをセット
    Set core = New WebView2Core

    '起動モードを設定する
    Dim timeStart As Single
    Select Case StartMode
        Case StartWebView2Mode.UserFormModeless
            '1. モードレスの場合は、処理は続行されます
            ' WebView2Core を UserForm 側に渡して初期化・表示してもらう
            printMsg info_, "Launch WebView2 in a modeless window...", FromProcedureName
            With WebView2InExcelForm
                .AttachCore core, checkURI(FirstURL), addArgs, userProfile
                .show vbModeless
            End With

            '2. WebView2 の内部処理が終わるまで待機
            '※UserForm側のイベント発火のタイミングを合わせるため、待機処理
            timeStart = Timer
            Do While Not core.IsReady And Timer - timeStart < ReadyWebView2Time
                core.ProcessMessages
            Loop

            '3. 時間内に準備できなかったら、停止
            If Not core.IsReady Then
                quit
                printMsg ERROR_, "Initialization of WebView2 was successful, but the setup could not be completed within the time limit.", FromProcedureName, doRaiseError:=True
            End If

        Case StartWebView2Mode.UserFormModal
            'モーダル表示の場合は仕様上、ウィンドウを閉じるまで、ここで止まります
            printMsg info_, "Launch WebView2 in a modal window...", FromProcedureName
            With WebView2InExcelForm
                .AttachCore core, checkURI(FirstURL), addArgs, userProfile
                .show vbModal
            End With

            '以降の処理は無意味なので、抜ける
            Exit Sub

        Case StartWebView2Mode.Background
            '1. WebView2 初期化
            printMsg info_, "WebView2 will launch in the background...", FromProcedureName
            If core.Initialize(Application.hWnd, 0, 0, InternalResolutionWidth, InternalResolutionHight, checkURI(FirstURL), userProfile, addArgs) Then
                printMsg info_, "WebView2 initialization successful.", FromProcedureName
            Else
                printMsg ERROR_, "WebView2 initialization failed." & vbCrLf & vbCrLf & "<cause>" & core.LastErrorDescription, FromProcedureName, doRaiseError:=True
            End If

            '2. WebView2 の内部処理が終わるまで待機
            timeStart = Timer
            Do While Not core.IsReady And Timer - timeStart < ReadyWebView2Time
                core.ProcessMessages
            Loop

            '3. 時間内に準備できなかったら、停止
            If Not core.IsReady Then
                quit
                printMsg ERROR_, "Initialization of WebView2 was successful, but the setup could not be completed within the time limit.", FromProcedureName, doRaiseError:=True
            End If

        Case Else: printMsg ERROR_, "This startup mode is not defined.", FromProcedureName, doRaiseError:=True
    End Select

    'Complete initialization ------------------------------------------------------
    If lngLastID <= 0 Then lngLastID = 1
    invokeMethod ("Runtime.enable")    '最もよく使うであろう、JavaScript機能を有効化しておく
    getSessionInfo                      'Store initial session info for printParams debugging
    serialize
    '------------------------------------------------------------------------------

    'FirstURL が指定されてる場合は、読み込みが終るまで待機しておく
    Dim timerStart As Single, strLoc As String
    If FirstURL <> "" Then
        timerStart = Timer
        Do
            If Me.Name <> "" Then Exit Do
            If Timer - timerStart > ReadyWebView2Time Then Err.Raise ERR_TIMEOUT, , "Timeout waiting for the browser to load the homepage."
            DoEvents
        Loop
       wait
    End If

    'Finally print the session details for debugging ------------------------------
    printParams
    printMsg info_, "WebView2 initialization completed after " & Timer - timeStart & "s", FromProcedureName
    '------------------------------------------------------------------------------

End Sub



'***************************************************************************************************
'                              ■■■ プロパティ関連 ■■■
'***************************************************************************************************
'* 機能　　：同期モードでの JavaScript結果のうち、最後の例外情報を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：最後の「例外（Exception）」情報
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、「最後」の例外情報なので、成功しても消滅しません。`Err.LastDllError`に準拠してます。
'            ・これはあくまでも、`例外`のみ扱います。「try { ... } catch(e) { return e }」系統の取得は対象外です。
'            ・非同期によるあとで取り出したパターンとは、連動しません。
'***************************************************************************************************
Property Get LastJavaScriptException() As Dictionary
    Set LastJavaScriptException = lastExceptionDetailsJS
End Property



'***************************************************************************************************
'                              ■■■ CDP送信準備関連 ■■■
'***************************************************************************************************
'* 機能　　：CDPのMethod名と、parametersを渡して、低レベルにてブラウザ制御をします
'---------------------------------------------------------------------------------------------------
'* 返り値　：CDP-Jsonコマンド の実行結果(result)オブジェクト(Json-Dictionary 形式)。失敗時は、`Nothing`となります。
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
    Dim HasError As Boolean: HasError = invokeError(results)
    If HasError Then
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
    '※WebView2の場合は、Resultオブジェクトが直で返る仕様なので、エラー有無で結果を返す返さない判定をします
    If Not (HasError) Then Set invokeMethod = results
    printMsg Debug_, "Invoked """ & methodName & """ with result: " & jsConverter.ConvertToJson(results), FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：`invokeMethod`の非同期実行(コマンド結果を待たない)バージョンです
'---------------------------------------------------------------------------------------------------
'* 返り値　：実行時の整理券ID（ResultCDPForAsync 相当の将来拡張用）
'* 引数　　：※`invokeMethod`と同じ
'***************************************************************************************************
Public Function invokeMethodAsync(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Long
    Const FromProcedureName As String = ThisClassName & ".invokeMethodAsync"

    Dim paramStr As String
    Dim results As Scripting.Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = jsConverter.ConvertToJson(params)
    Set results = SendMessage(methodName, paramStr, alwaysBrowserContext, Async:=True)

    '非同期送信の失敗（COM 呼び出し失敗）だけここで検知
    Dim hr As Long
    hr = CLng(results.Items(0))
    If hr <> 0 Then
        If StopError Then
            printMsg ERROR_, "Failed to invokeAsync method """ & methodName & """" & vbCrLf & vbCrLf & "=== AsyncInvokeError ===" & vbCrLf & "HRESULT: 0x" & Hex$(hr), FromProcedureName, doRaiseError:=True
        Else
            printMsg WARN_, "Failed to invokeAsync method """ & methodName & """ HRESULT: 0x" & Hex$(hr), FromProcedureName
        End If
    End If

    invokeMethodAsync = CLng(results.keys(0))
    printMsg Debug_, "AsyncInvoked """ & methodName & """ with HRESULT: 0x" & Hex$(hr) & " (CommandID=" & invokeMethodAsync & ")", FromProcedureName
End Function



'***************************************************************************************************
'* 機能　　：Async CDPコマンド結果の後回収
'---------------------------------------------------------------------------------------------------
'* 引数　　：CommandID（invokeMethodAsync の戻り値）
'* 返り値　：結果 Dictionary（取得できない場合は Nothing）
'***************************************************************************************************
Public Function ResultCDPForAsync(ByVal CommandID As Long, Optional ByRef ErrorExist As Boolean) As Scripting.Dictionary
    Dim raw As String
    raw = core.ResultCDPForAsync(CommandID, ErrorExist)
    If Len(raw) = 0 Then Exit Function
    Set ResultCDPForAsync = jsConverter.ParseJson(raw)
End Function


'***************************************************************************************************
'                              ■■■ 便利メソッド集 ■■■
'***************************************************************************************************
'* 機能　　：記述したJavaScriptをブラウザ内で実行します。JavaScript版低レベル操作としての位置付けです。
'            Webスクレイピングに関わる引数は全て設けています。
'---------------------------------------------------------------------------------------------------
'* 返り値　：Variant型。実行したJavaScriptコード、引数:returnByValue や、IFEXCEPTION に依存します。
'* 引数　　：strString                      実行したいJavaScriptコードそのもの。`As Is`として実行されます。※ほとんどのエスケープ処理は、`WebJsonConverter`が担うため、ここでわざわざ手動エスケープの記述は必要ありません。
'
'            objectId                       実行の基準要素オブジェクトIDを指定します。これを指定すると`Runtime.callFunctionOn`として優先実行されます。
'
'            objectArguments                これを指定すると`Runtime.callFunctionOn`として優先実行されます。
'                                           予め、呼び出し側にてCollectionを組み立てて渡してください。
'
'            contextId                      `iframe`内で、JavaScript実行をしたいときに指定します。`objectId`が指定されてる場合は、無視されます
'            IFEXCEPTION                    JavaScript実行中に「例外（Exception）」が発生した場合の代替値を指定します。名前の通り、IFERROR関数チックに使えます。
'
'            returnByValue                  True ：type:"object" 等の場合でも、`value`で受け取れるようにCDPが試みます。実行コードによっては、Dictionary or Collection形式で返されるため、`SET`で受け取る必要があります。なお、空のオブジェクトで返る場合は、対応してないデータ型の可能性があるためその際は、`false`に切り替えてください。
'                                           False：type:"object" 等の場合は、`{"objectId": "-4985739485.1.3"}`にある`-4985739485.1.3`というブラウザ内変数名的なものとして返えります。再度`jsEval`を使う際の引数`objectId`に流用できます。「このJSコードは`SET`で...」という意識は不要です。
'
'            awaitPromise                   True ：JavaScriptそのものの実行結果が非同期処理であっても、処理が終わるまで待ってから結果を返すことができます
'                                           False：JavaScriptそのものの実行結果が非同期処理の場合、結果を待たずに続行されます
'
'            userGesture                    True ：「人間がクリックやキーボード操作をした結果」としてブラウザに偽装します。スクレピング対策が施されてる場合は、`True`にしておくといいかもしれません
'                                           False：特に施しません
'
'            allowUnsafeEvalBlockedByCSP    True ：サイトの外部JavaScript実行規制を無視して、全てのJavaScript実行を許可します
'                                           False：サイトの外部JavaScript実行規制に従います
'
'            serializationOptions           意地でも、type:"object" 等となっても、`value`で受け取れるように、細かい調整をします。予め、呼び出し側にてDictionaryを組み立てて渡してください。
'                                           このパラメーターがあると、`returnByValue`の引数は無視されます。
'
'            generatePreview                `objectId`で返った中身をプレーンテキスト化して、プレビューします。予め、ログレベルを`DEBUG`以下にしてください。基本は、開発時のみ`True`にしてください。
'
'            RunAsyncCDP                    True ：結果を待たないCDPで実行します。実行時のIDが返却されます。主に、「この操作、`alert`関数が発動するな...」という場面で使用します。
'                                           False：結果を待つCDPで実行します。基本はこれでOK
'
'            StopException                  `True`で、JavaScriptコード内で例外が起こった場合、Err.Raiseで止まります。基本は、開発時のみ使用する想定です
'
'            dbgMsg                         Allow the option to bypass error if intended
'---------------------------------------------------------------------------------------------------
'* 機能説明：This function commands the attached browser instance to execute a Javascript expression to perform required tasks.
'            Note: the original function has a bool retry variable which was removed for simplicity.
'            Added dbgMsg to read debug line on the go when needed.
'
'* 注意事項：・利用時は、「名前付き引数`xxx:=yyy`」として運用することをお勧めします
'            ・JavaScript実行中の例外確認は、「IsError関数」で判断してください。詳細の確認は、`LastJavaScriptException`プロパティで確認できます。「IsError関数」での確認が面倒なら、引数`IFEXCEPTION`を活用してください。
'            ・`returnByValue`=true の場合は、SET or LET で受け取る意識が必要になります
'
'            ・「このJavaScriptコードの場合は、こういう返り値が来るな」といった知識がないと、返り値に困惑します。
'               開発時は、デバッグログを活用した試行錯誤をすることをお勧めします。
'
'            ・`objectId`や`objectArguments`を指定する`Runtime.callFunctionOn`方式の場合は、`strString`を 「function() { ... }」 形式で記述する必要があります
'
'            ・実行したJavaScriptコードの結果が、プリミティブ型の場合、`returnByValue`の設定問わずそのプリミティブ型の中身が返ります。
'              常に`objectId`で返したい場合は、JavaScriptコードそのものを工夫する必要があります。
'
'            ・JavaScriptの場合、例外を利用した処理が多々あるため、運用時は`StopException`を省略しておくことを推奨します
'            ・CDP実行自体のエラーキャッチは常に`True`となります
'            ・`IFEXCEPTION`はあくまでも、「例外」時のみ機能します。JavaScript自体は成功してるけど、結果がエラーの場合はISERROR関数で対処する必要があります。
'---------------------------------------------------------------------------------------------------
'* 元URL   ：https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#method-evaluate
'***************************************************************************************************
Public Function jsEval(strString As String, Optional objectId As String, Optional contextId As Long, Optional IFEXCEPTION As Variant, _
                       Optional objectArguments As Collection, _
                       Optional returnByValue As Boolean, _
                       Optional awaitPromise As Boolean, _
                       Optional userGesture As Boolean, _
                       Optional allowUnsafeEvalBlockedByCSP As Boolean, _
                       Optional serializationOptions As Dictionary, _
                       Optional generatePreview As Boolean, _
                       Optional RunAsyncCDP As Boolean, _
                       Optional StopException As Boolean, _
                       Optional dbgMsg As Boolean = True) As Variant
    Const FromProcedureName As String = ThisClassName & ".jsEval"


    Dim CDPresults  As Scripting.Dictionary
    Dim strMessage  As String
    Dim objMessage  As New Scripting.Dictionary
    Dim CDPparams   As New Scripting.Dictionary

step1:

   'Construct message for sending ----------------------------------------------------------
    CDPparams.RemoveAll

    '`callFunctionOn`か`evaluate`を決める判定に入ります
    Dim UseJsMethod As String

    ' 【パターン1】特定の要素を操作する場合
    If StrPtr(objectId) Then
        UseJsMethod = "Runtime.callFunctionOn"

        '必要なパラメーターをセット
        CDPparams.Add "functionDeclaration", strString
        If Not objectArguments Is Nothing Then CDPparams.Add "arguments", objectArguments   '引数があったら適用しておく

        'objectId があるので、場所(ContextID)を指定する必要はない
        CDPparams.Add "objectId", objectId

    ' 【パターン2】要素指定はないが、安全に引数(arguments)を渡したい場合
    ElseIf Not objectArguments Is Nothing Then
        UseJsMethod = "Runtime.callFunctionOn"

        '必要なパラメーターをセット
        CDPparams.Add "functionDeclaration", strString
        CDPparams.Add "arguments", objectArguments

        ' objectId がないので、実行する「場所(executionContextId)」を教える必要がある
        If contextId Then CDPparams.Add "executionContextId", contextId

    ' 【パターン3】引数も要素指定もない、ただのコード実行
    Else
        ' 最もシンプルな evaluate を使う
        UseJsMethod = "Runtime.evaluate"

        '必要なパラメーターをセット
        CDPparams.Add "expression", strString
        If contextId Then CDPparams.Add "contextId", contextId
    End If

    '以降は、共通設定を施す
    '※文字数節約により、IF文をかます
    If returnByValue Then CDPparams.Add "returnByValue", returnByValue
    If awaitPromise Then CDPparams.Add "awaitPromise", awaitPromise
    If userGesture Then CDPparams.Add "userGesture", userGesture
    If allowUnsafeEvalBlockedByCSP Then CDPparams.Add "allowUnsafeEvalBlockedByCSP", allowUnsafeEvalBlockedByCSP
    If generatePreview Then CDPparams.Add "generatePreview", generatePreview
    If Not serializationOptions Is Nothing Then CDPparams.Add "serializationOptions", serializationOptions

    'parameter部分を組み立て
    strMessage = jsConverter.ConvertToJson(CDPparams)
   '----------------------------------------------------------------------------------------

    'Send and check CDPresult -------------------------------------------------------------------------------------------
    '引数に応じて、CDP自体を同期/非同期で実行し、分岐します
    Set CDPresults = SendMessage(UseJsMethod, strMessage, Async:=RunAsyncCDP)
    If RunAsyncCDP Then
        '`WriteFile`のエラーをチェック
        If CDPresults.Items(0) Then
            printMsg ERROR_, "Failed to AsyncCDP-WriteFile JavaScript: " & strString & vbCrLf & vbCrLf & "=== WriteFileErrorDetail ===" & vbCrLf & "ErrorCode:" & CDPresults.Items(0) & " - " & ErrorDetail.GetMessage(CDPresults.Items(0)), FromProcedureName, doRaiseError:=True
        Else
            'ID返却して抜ける
            jsEval = CLng(CDPresults.keys(0))
        End If

        printMsg Debug_, "Async CDP JavaScript with WriteFile Result: " & CDPresults.Items(0) & " - " & ErrorDetail.GetMessage(CDPresults.Items(0), "kernel32") & " ID: " & CDPresults.keys(0), FromProcedureName
        Exit Function
    Else
        If seeRawSendMsgDbg Then printMsg Trace_, "jsEval - ResultCDP - JsonString = " & jsConverter.ConvertToJson(CDPresults), FromProcedureName
        If CDPresults Is Nothing Then printMsg ERROR_, "Unable to contact the Browser instance!", FromProcedureName, doRaiseError:=True
        If Not CDPresults.Exists("error") Then GoTo step2   'CDP自体のエラーがなければ、`step2`へ
    End If
    '-----------------------------------------------------------------------------------------------------------------

   'Handle errors ---------------------------------------------------------------------------------------------------
    Dim errMsg As String, retried As Boolean
    errMsg = CDPresults("error")("message")
    printMsg WARN_, "jsEval Error: " & errMsg, FromProcedureName
    If errMsg = "'" & UseJsMethod & "' wasn't found" Then printMsg WARN_, "Possibly the main session has been abruptly closed.", FromProcedureName
    If errMsg = "Session with given id not found." And Not retried Then
        printMsg info_, "Attempting to update Session ID " & sessionID, FromProcedureName
        sessionID = "": sessionID = attachToTab(targetID)
        printMsg info_, "New Session ID obtained: " & sessionID, FromProcedureName
        retried = True: GoTo step1
    End If
    printMsg ERROR_, "The Browser was unable to evaluate the JS expression: " & strString & vbCrLf & vbCrLf & "<ErrorDetail>" & vbCrLf & errMsg, FromProcedureName, doRaiseError:=True
   '-----------------------------------------------------------------------------------------------------------------

step2:

    'JavaScript実行結果部分の解析フェーズに移ります ----------------------------------------------------------------
    Dim RemoteObject As Dictionary
    Set RemoteObject = CDPresults.Item("result")
    If seeRawSendMsgDbg Then printMsg Debug_, "jsEval - RemoteObject - JsonString = " & jsConverter.ConvertToJson(RemoteObject), FromProcedureName

    '例外（JS側のエラー）が発生していないかチェックし、あったら、ログを表示し、抜けます
    If CDPresults.Exists("exceptionDetails") Then
        '最後の詳細例外情報を記録
        Set lastExceptionDetailsJS = CDPresults("exceptionDetails")

        '例外による`Err.Raise`で止めるか？
        If StopException Then
            'ログ表示用に、詳細の方を文字列化
            Dim exceptionDetails As String: exceptionDetails = jsConverter.ConvertToJson(lastExceptionDetailsJS)
            printMsg ERROR_, "An exception occurred while executing JavaScript: " & strString & vbCrLf & vbCrLf & "<exceptionDetails>" & vbCrLf & exceptionDetails, FromProcedureName, doRaiseError:=True
        Else
            If IsMissing(IFEXCEPTION) Then
                '省略時は、エラー型として返します。https://learn.microsoft.com/ja-jp/office/vba/excel/Concepts/Cells-and-Ranges/cell-error-values
                '※例外の種類は一律「#VALUE!」扱いとします。あくまでも「JavaScript実行中に例外が起きた」という目印として判断してください
                jsEval = CVErr(xlErrValue)
                printMsg WARN_, "An exception occurred while executing JavaScript: " & strString & "  It was returned as an error type.", FromProcedureName
            Else
                '引数:IFEXCEPTION が指定されてる場合はそっちを優先して返します
                If IsObject(IFEXCEPTION) Then Set jsEval = IFEXCEPTION Else jsEval = IFEXCEPTION
                printMsg WARN_, "An exception occurred while executing JavaScript: " & strString & " The argument `IFERROR` was specified, so it was returned.", FromProcedureName
            End If

            '簡易部分の例外内容を出す
            printMsg WARN_, "ExceptionType: " & RemoteObject("className") & " , " & "ExceptionDescription: " & RemoteObject("description"), FromProcedureName
        End If

        Exit Function
    End If

   'If the return type has a specific type we can return the result -------------------------------------------
    Dim tmp
    If JSresultAnalysis(RemoteObject, serializationOptions, Not returnByValue, generatePreview, tmp) Then
        Set jsEval = tmp
    Else
        jsEval = tmp
    End If
   '-----------------------------------------------------------------------------------------------------------

step3:

'Output to Immediate Window if enabled -------------------------------------------------------------------
    Dim strLogValue As String

    ' VarTypeを使って型ごとにログ用の文字列を作る
    Select Case VarType(jsEval)
        Case vbDouble
            strLogValue = "[Double] " & jsEval

        Case vbString
            strLogValue = "[String] " & jsEval

        Case vbBoolean
            strLogValue = "[Boolean] " & jsEval

        Case Else
            strLogValue = "[" & TypeName(jsEval) & "]"
    End Select

    ' 改行を消して、長すぎる場合はカットする（整形）
    strLogValue = Replace(strLogValue, vbNewLine, " ")
    If Len(strLogValue) > 200 Then
        strLogValue = Left(strLogValue, 200) & "... (Total: " & Len(strLogValue) & " chars)"
    End If

    ' 最終的なデバッグ出力
    printMsg Debug_, "Script executed: " & strString & " | Returned: " & strLogValue, FromProcedureName
   '---------------------------------------------------------------------------------------------------------

End Function

'***************************************************************************************************
'* 機能　　：JavaScript実行で返ったJsonDictionaryから、結果値`value`だけを取り出すように試みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：RemoteObject                       `Runtime.RemoteObject`でのJson-Dictionaryを指定します。　https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#type-RemoteObject
'            PriorityReturn_deepSerializedValue `deepSerializedValue`があったら、そっちを最優先 に返すようにします。`serializationOptions`に依存します。
'            PriorityReturn_objectId            `objectId`           があったら、そっちを第2優先に返すようにします。`returnByValue`に依存します。
'            ReturnValue                        結果値。参照渡しによる実質返り値です。`primitive values` OR `Object` 両方が想定されます。
'
'* 返り値　：ReturnValueがObjectであるかの論理値。`ReturnValue`がObject型であるかを呼び出し側で正しく代入するために必要です。
'---------------------------------------------------------------------------------------------------
'* 注意事項：返す値は、`value` OR `objectId` OR `deepSerializedValue` に限ります。`description`等といった補助情報やエラー情報は返しません。
'***************************************************************************************************
Private Function JSresultAnalysis(RemoteObject As Dictionary, PriorityReturn_deepSerializedValue As Dictionary, PriorityReturn_objectId As Boolean, ObjectPreview As Boolean, ByRef ReturnValue As Variant) As Boolean
    Const FromProcedureName As String = ThisClassName & ".JSresultAnalysis"
    Const ErrorLandmark     As String = "error"


    '------------------------ 1. 特定項目有無チェック ------------------------
    '`serializationOptions`が指定されてたら、`deepSerializedValue`を最優先に返します
    If Not PriorityReturn_deepSerializedValue Is Nothing Then
        printMsg info_, "The parameter `serializationOptions` is specified. Check if `deepSerializedValue` exists.", FromProcedureName

        '`deepSerializedValue`があるか？
        If RemoteObject.Exists("deepSerializedValue") Then
            '見つかったので、`deepSerializedValue`を返して、ここで抜ける
            Set ReturnValue = RemoteObject("deepSerializedValue")
            printMsg info_, "JS-Return type: object-deepSerializedValue", FromProcedureName

            'Objectで返したのでその目印もつけておく
            JSresultAnalysis = True
            Exit Function
        Else
            printMsg WARN_, "`The parameter `serializationOptions` was specified, but `deepSerializedValue` was missing.", FromProcedureName
        End If
    End If

    '`objectId`を返す優先になってたら、こっちを返します
    If PriorityReturn_objectId Then
        printMsg Trace_, "Since `returnByValue` is `False`, we check for the existence of `objectId`.", FromProcedureName

        '`objectId`があるか？
        If RemoteObject.Exists("objectId") Then
            '見つかったので、`objectId`を返す
            ReturnValue = CStr(RemoteObject("objectId"))
            printMsg info_, "JS-Return type: objectId", FromProcedureName

            '返ったObjectID情報を列挙する
            Dim objectSubtype As String: objectSubtype = RemoteObject("subtype")
            printMsg Debug_, String(100, "="), FromProcedureName
            printMsg Debug_, "objectId information", FromProcedureName
            printMsg Debug_, String(100, "="), FromProcedureName
            printMsg Debug_, "objectId          : " & ReturnValue, FromProcedureName
            printMsg Debug_, "objectType        : " & RemoteObject("type"), FromProcedureName
            printMsg Debug_, "objectSubtype     : " & objectSubtype, FromProcedureName
            printMsg Debug_, "objectClassName   : " & RemoteObject("className"), FromProcedureName
            printMsg Debug_, "objectDescription : " & RemoteObject("description"), FromProcedureName
            If ObjectPreview Then printMsg Debug_, "objectPreview     : " & jsConverter.ConvertToJson(RemoteObject("preview")), FromProcedureName
            printMsg Debug_, String(100, "="), FromProcedureName

            'エラー情報があればそのログを出しておく
            '※「try { ... } catch(e) { return e }」といった書き方をしないとこのログは表示されない想定
            If objectSubtype = ErrorLandmark Then printMsg WARN_, "This objectId(" & ReturnValue & ") contains an error. ErrorType: " & RemoteObject("className") & " , " & "ErrorDescription: " & RemoteObject("description"), FromProcedureName

            Exit Function
        Else
            'primitive values OR Objectがnullの可能性あり
            printMsg Debug_, "`objectId` was not found. It may be a primitive value or `null`, so we will try to retrieve it.", FromProcedureName
        End If
    Else
        printMsg Trace_, "`returnByValue` is `True`. Attempting to retrieve the primitive value.", FromProcedureName
    End If


    '------------------------ 2. primitive values OR JsonDictionaryの取得を試みる ------------------------
    'primitive values の可能性があるので、それの取得を試みます
    '以降は基本的に、`returnByValue:=True`でないと通りません
    'Objectとして返そうとしたが`null`の場合もここを通って、null返却します

    ' 格納型情報を取得
    Dim ValueType       As String: ValueType = RemoteObject.Item("type")
    Dim ExistsSubtype   As Boolean: ExistsSubtype = RemoteObject.Exists("subtype")
    Dim ValueSubtype    As String: ValueSubtype = RemoteObject.Item("subtype")
    printMsg Debug_, "Type info. type: " & ValueType & " , subtype: " & ValueSubtype, FromProcedureName

    '型に応じた分岐処理
    Select Case ValueType
        Case "object"
            ' object型の場合はsubtypeで更に細かく分岐
            Select Case ValueSubtype
                Case "null"
                    ReturnValue = Null
                    printMsg info_, "JS-Return type: object-null", FromProcedureName

                Case "array"
                    ' 配列やDOM要素(returnByValue時)は、valueの中に中身が入っている
                    ' 配列の場合は Scripting.Dictionary (連番) が入ります
                    '※誇らしげに説明してますが、`returnByValue:=True`で、`array`だと、subtype そのものがなくなります...が一応書いておきます。
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-array", FromProcedureName

                Case "node"
                    '多分、来ない...が一応書いておきます。
                    '`returnByValue:=True`で、`node`も、subtype そのものがなくなります...
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-node", FromProcedureName

                Case "regexp"
                    '以下同文
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-regexp", FromProcedureName

                Case "map"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-map", FromProcedureName

                Case "date"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-date", FromProcedureName

                Case "set"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-set", FromProcedureName

                Case "weakmap"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-weakmap", FromProcedureName

                Case "weakset"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-weakset", FromProcedureName

                Case "iterator"     '※ObjectIDモードでも空欄です?!
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-iterator", FromProcedureName
                
                Case "generator"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-generator", FromProcedureName

                Case "promise"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-promise", FromProcedureName

                Case "proxy"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-proxy", FromProcedureName

                Case "typedarray"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-typedarray", FromProcedureName

                Case "arraybuffer"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-arraybuffer", FromProcedureName

                Case "dataview"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-dataview", FromProcedureName

                Case "webassemblymemory"
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-webassemblymemory", FromProcedureName

                Case "trustedtype"  '※ObjectIDモードでも空欄です?!
                    Set ReturnValue = RemoteObject.Item("value")
                    printMsg info_, "JS-Return type: object-trustedtype", FromProcedureName

                Case ErrorLandmark
                    '`returnByValue:=True`では意味をなさないようです...が念のため
                    printMsg WARN_, "This value contains an error. ErrorType: " & RemoteObject("className") & " , " & "ErrorDescription: " & RemoteObject("description"), FromProcedureName
                    ReturnValue = CVErr(xlErrValue)
                    printMsg info_, "JS-Return type: object-error", FromProcedureName

                Case Else
                    '"subtype"自体がない場合は、`array`として扱います
                    '`array`,`typedarray`が機能してます。他は空({})です
                    If Not ExistsSubtype Then
                        ' 配列やDOM要素(returnByValue時)は、valueの中に中身が入っている
                        ' 配列の場合は Scripting.Dictionary (連番) が入ります
                        ' returnByValue=True なら value に Dictionary が入っているはず
                        printMsg WARN_, "It's of type Object, but no `subtype` was found. It will be treated as an `array`.", FromProcedureName
                        Set ReturnValue = RemoteObject.Item("value")
                        printMsg info_, "JS-Return type: object-array", FromProcedureName

                    Else
                        'ドキュメント化されてない型にも備えておく
                        printMsg WARN_, "This data type cannot be handled in VBA. type: " & ValueType & " , subtype: " & ValueSubtype, FromProcedureName
    
                        '返せる物があれば、返しておく
                        If RemoteObject.Exists("value") Then
                            ' オブジェクト(Dictionary)を返す場合は Set が必要
                            If IsObject(RemoteObject.Item("value")) Then
                                Set ReturnValue = RemoteObject.Item("value")
                            Else
                                ReturnValue = RemoteObject.Item("value")
                            End If
                        ElseIf RemoteObject.Exists("unserializableValue") Then
                            ReturnValue = CStr(RemoteObject.Item("unserializableValue"))
                        Else
                            ReturnValue = RemoteObject.Item("description")
                        End If
                    End If
            End Select

        Case "function"
            ' 関数が返ってきた場合は、「{"type":"function","value":{}}」なので、`SET`
            Set ReturnValue = RemoteObject.Item("value")
            printMsg info_, "JS-Return type: function", FromProcedureName

        Case "undefined"
            ' JSの undefined は VBAでは Empty が一番近い
            ReturnValue = Empty
            printMsg info_, "JS-Return type: undefined", FromProcedureName

        Case "string"
            ' 文字列として認識するように返す
            ReturnValue = CStr(RemoteObject.Item("value"))
            printMsg info_, "JS-Return type: string", FromProcedureName

        Case "number"
            'VBAで扱える型か？
            If RemoteObject.Exists("unserializableValue") Then
                'VBAで扱えない特殊な"number"の場合は、警告を出し、「#NUM!」で返します。
                printMsg WARN_, "This `number` can't be processed in VBA. It returns `#NUM!` instead. Actual JS result: " & RemoteObject("unserializableValue") & " , description: " & RemoteObject("description"), FromProcedureName
                ReturnValue = CVErr(xlErrNum)
            Else
                ' Double型として認識するように返す
                ' ※JavaScriptは、整数専用等の型がないため
                ReturnValue = CDbl(RemoteObject.Item("value"))
            End If

            printMsg info_, "JS-Return type: number", FromProcedureName

        Case "boolean"
            ' boolean型として認識するように返す
            ReturnValue = CBool(RemoteObject.Item("value"))
            printMsg info_, "JS-Return type: boolean", FromProcedureName

        Case "symbol"
            '※`symbol`は、`returnByValue:=True`にすると、「Object couldn't be returned by value.」というCDPエラーで返るので通らないと思うが念のため。
            ReturnValue = RemoteObject.Item("value")
            printMsg info_, "JS-Return type: symbol", FromProcedureName

        Case "bigint"
            ' BigInt は JSON に変換できない巨大整数のため、unserializableValue の中にある値をそのまんま返却します
            ReturnValue = CStr(RemoteObject.Item("unserializableValue"))
            printMsg WARN_, "Unserializable BigInt detected. Returning as String: " & ReturnValue & " , description: " & RemoteObject("description"), FromProcedureName
            printMsg info_, "JS-Return type: bigint", FromProcedureName

        Case Else
            'ドキュメント化されてない型にも備えておく
            printMsg WARN_, "This data type cannot be handled in VBA. type: " & ValueType, FromProcedureName

            '返せる物があれば、返しておく
            If RemoteObject.Exists("value") Then
                ' オブジェクト(Dictionary)を返す場合は Set が必要
                If IsObject(RemoteObject.Item("value")) Then
                    Set ReturnValue = RemoteObject.Item("value")
                Else
                    ReturnValue = RemoteObject.Item("value")
                End If
            ElseIf RemoteObject.Exists("unserializableValue") Then
                ReturnValue = CStr(RemoteObject.Item("unserializableValue"))
            Else
                ReturnValue = RemoteObject.Item("description")
            End If
    End Select

    'SETステートメントするかの判定を用意する
    JSresultAnalysis = IsObject(ReturnValue)
End Function

Public Sub snapPage(FolderPath As String, FileName As String, Optional getFullPage As Boolean = False)
'----------------------------------------------------------------------------------
' Capture a screenshot of the entire HTML page from the top to bottom. The method
' employs html2canvas external library is which loaded via addJsLibrary. To snap
' only a certain element within the page, indicate the same with the targetElement
' argument (default is "document.body" for the entire page).
' Original source: https://stackoverflow.com/a/51478809
' Updated: 14/04/23
'          23/10/25 Daniel Polak - fixed bug where the all files in the Downloads folder were deleted instead of just the temporary screenshot
'          24/10/25 Daniel Polak - added CDP native screenshot capability which avoids use of external JavaScript library
'          07/01/26 Long Vh      - removed the use of html2canvas js lib in favor of using the native CDP method to capture the page. Overhauled the function design.
'----------------------------------------------------------------------------------
' Note : if getFullPage = false, snapPage will only get a screenshot of the current
'        view. If true, it will get all the way from the start to the bottom. See
'        demo.getSnapShot sub procedure for example.
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".snapPage"


    Dim params As New Scripting.Dictionary
    Dim r As Scripting.Dictionary
    Dim b64 As String, tmpPng As String

   'Verify if the page can be captured with screenshot
    Set r = invokeMethod("Page.getLayoutMetrics")
    If invokeError(r) Then
        printMsg WARN_, "Unable to perform snapPage on this screen as it has no layout metrics.", FromProcedureName
        Exit Sub
    End If

   'Perform the snip
    If getFullPage Then params.Add "captureBeyondViewport", True
    Set r = invokeMethod("Page.captureScreenshot", params)

   'Save to the same folder if not indicated
    If Not invokeError(r) Then
        b64 = r("data")
        SaveBase64ToFile b64, FolderPath, FileName
        printMsg info_, "Screenshot saved to: " & FolderPath & "\" & FileName, FromProcedureName
    End If

End Sub

Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
    Const FromProcedureName As String = ThisClassName & ".navigate"


   'Check if already navigated
    Dim curURL As String
    curURL = jsEval("window.location.href")
    If curURL = checkURI(strURL) Then
        printMsg WARN_, "Already navigated to """ & strURL & """", FromProcedureName
        Exit Sub
    End If

   'If not yet then do navigate
    Dim results As Scripting.Dictionary
    Set results = SendMessage("Page.navigate", "{""url"":""" & strURL & """}")

    'should also add a maxTimeLimit to Wait
    wait till:=till
    getSessionInfo 'update new session info for printParams
    printMsg info_, "Navigated to: " & strURL, FromProcedureName

End Sub

Public Sub wait(Optional till As ReadyState = isComplete, Optional dbgState As Boolean = False)
'--------------------------------------------------------------------------------------------------
' Wait until browser reaches a specific state. Default is until it fully loads ("complete")
' Possible states:  loading / interactive / complete
' Complete state is preferable when certain information is only showed when fully loaded.
' Otherwise, an interactive state is good enough and can potentially save a lot of wait time.
' However, do note that some web apps only fully displays elements at complete readystate.
' To read ReadyState on the go, set the dbgState to True to print the state to the Immediate Window
' Original Name: waitCompletion
' Updated: 26/04/23 - Enhanced 'till' to Enum type for ease of use
'                   - Added exit for targetState = interactive in case CDP misses the state change
'--------------------------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".wait"


    Dim strState As String, targetState As String
    Select Case till
        Case isLoading: targetState = "loading"
        Case isInteractive: targetState = "interactive"
        Case isComplete: targetState = "complete"
    End Select

    '初期値を設定しておく
    '最初はほぼ必ずloadingだから、`targetState`が`loading`になってたら、わざわざチェックする必要性はない
    strState = "loading"

    Do Until strState = targetState
        strState = jsEval("document.readyState")                 'dbgMsg = False to not spam Immediate Window
        printMsg Debug_, "Current ReadyState: " & strState, FromProcedureName   'option to print ReadyState info during iteration
        If targetState = "interactive" And strState = "complete" Then Exit Do   'also exit if better than interactive is satisfied
        sleep 0.1                                                               'reduce sleep will speed up but will cost cpu power
    Loop

    printMsg Debug_, "Wait '" & targetState & "' satisfied", FromProcedureName

End Sub

Public Sub sleep(Optional seconds As Double = 0.5)
'------------------------------------------------------------------
' Custom sleep function. Sleep by 0.5s by default.
' Useful for a quick necessary pause when needed.
'------------------------------------------------------------------

    Const baseUnit As Long = 1000    'ie. millisecs

    sleep2 seconds * baseUnit
    DoEvents

End Sub

Public Sub notify(msg)
'----------------------------------------------------------------------------------
' Send a custom notification message on the bottom right corner of the screen. This
' popup message is nicer than alert() and it is not modal: ie. does not the user to
' interact with it while pausing the automation.
' Credit: Claude AI
' Updated: 01/12/2025
'----------------------------------------------------------------------------------

   'Make sure backslash if any is escaped
    msg = Replace(msg, "\", "\\")

    jsEval "function showStatus(msg, duration = 10000) {" & _
                "let c = document.getElementById('statusContainer');" & _
                "if (!c) {" & _
                    "c = document.createElement('div');" & _
                    "c.id = 'statusContainer';" & _
                    "c.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:9999;';" & _
                    "document.body.appendChild(c);" & _
                "}" & _
                "const m = document.createElement('div');" & _
                "m.style.cssText = 'background:#4CAF50;color:white;padding:16px 24px;margin-top:10px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3);white-space:nowrap;transition:all 0.3s;opacity:0;transform:translateX(400px);font-size:16px;font-family:system-ui,-apple-system,BlinkMacSystemFont,Roboto,sans-serif;font-weight:300;line-height:1.5;';" & _
                "m.textContent = msg;" & _
                "c.appendChild(m);" & _
                "setTimeout(() => { m.style.opacity = '1'; m.style.transform = 'translateX(0)'; }, 10);" & _
                "setTimeout(() => {" & _
                    "m.style.opacity = '0';" & _
                    "m.style.transform = 'translateX(400px)';" & _
                    "setTimeout(() => c.removeChild(m), 300);" & _
                "}, duration);" & _
            "}"

    jsEval "showStatus('" & msg & "')"

End Sub

Public Sub activate(Optional focus As Boolean = True)
'----------------------------------------------------------------------------------
' Switch focus to the current tab being attached to under the current BRID.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    If focus Then invokeMethod "Target.activateTarget", params

End Sub

Public Sub handleDialog(Accept As Boolean, Optional promptText As String, Optional AsyncMode As Boolean)
'----------------------------------------------------------------------------------
' Accepts or dismisses a JavaScript initiated dialog (alert, confirm, prompt, or
' onbeforeunload). accept = True to accept the dialog, else dimiss it. promptText
' is the text to enter into the dialog prompt before accepting, used only if this is a
' prompt dialog.
' Note: this method requires the execution of "pageEnable" before the dialog appears
' so that the dialog event is probably exposed and captured as an event.
' Updated: 23/07/2025 - added as requested by Jane Zhang for her project
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary

    params("accept") = CBool(Accept)
    If StrPtr(promptText) Then params("promptText") = promptText
    If AsyncMode Then invokeMethodAsync "Page.handleJavaScriptDialog", params Else invokeMethod "Page.handleJavaScriptDialog", params

End Sub

Public Sub pageEnable(Optional enable As Boolean = True)
'----------------------------------------------------------------------------------
' Enable the page for events listening. Required for handleDialog for instance.
'----------------------------------------------------------------------------------

    invokeMethod IIf(enable, "Page.enable", "Page.disable")

End Sub

Private Sub SaveBase64ToFile(strB64 As String, FolderPath As String, FileName As String)
'----------------------------------------------------------------------------------
' Decode a base64 string and save to file without external dependencies (late bound).
'----------------------------------------------------------------------------------

    Dim dom As Object, el As Object, Bytes() As Byte, CharConv As CharacterCodeConversion

    Set dom = CreateObject("MSXML2.DOMDocument")
    Set el = dom.createElement("tmp")
    el.DataType = "bin.base64"
    el.Text = strB64
    Bytes = el.nodeTypedValue

    Set CharConv = New CharacterCodeConversion
    CharConv.BytesToSaveFile Bytes, FolderPath, FileName

End Sub



'***************************************************************************************************
'                              ■■■ 機能拡張系 ■■■
'***************************************************************************************************
Public Sub jsAddLib(strURL As String)
'----------------------------------------------------------------------------------
' Load the external JS Library to the document for use of specialized JS methods.
' This is very useful when built-in JS commands are not sufficient to execute a
' task easily, such as when taking a screenshot of an entire webpage while this can
' be done relatively easy using external JS libraries. This function shall enables
' massive possibility of extended features.
' https://stackoverflow.com/a/57572107
' Updated: 23/03/2023
'          23/10/2025 Daniel Polak - fixed bug in check whether JS library is loaded
'----------------------------------------------------------------------------------
' Note: the script library should only be added after the page is fully loaded. If
' it is loaded when the HTML document is refereshing, it risks adding the script to
' the old HTML document which will be replaced by the new one causing the script to
' be inexistent and resulting in "ReferenceError" due to its functions not defined.
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".jsAddLib"


   'Check if already loaded, if so then no need to add again
    If getElementByQuery("script[src='" & strURL & "']").varResult <> "null" Then
        printMsg WARN_, "The HTML already contains the JS Library: """ & strURL & """", FromProcedureName
        Exit Sub
    End If

   'Add script loading function
   '(Javascript Original)
    '    function loadScript(name) {
    '        var script = document.createElement('script');
    '        script.type = 'text/javascript'; script.src = name;
    '        document.head.appendChild(script);
    '        script.onload = () => {};                              //added to make sure script has been loaded
    '        return 'Script Library fully loaded';
    '    }
    jsEval "function loadScript(name) { var script = document.createElement('script'); script.type = 'text/javascript'; script.src = name; document.head.appendChild(script); script.onload = () => {}; return 'Script library loaded'; }"

   'Load the script URL
    jsEval "loadScript('" & strURL & "')"
    printMsg info_, "Successfully added the JS Library: """ & strURL & """", FromProcedureName

End Sub

Public Sub jsAddScript(strLocalScriptPath As String)
'----------------------------------------------------------------------------------
' Load the contents of the local script file directly to the HTML page. This is
' useful when you have a custom local script to add in a certain feature to a web
' application (eg adding a custom popup message).
' Updated: 01/12/2025
'----------------------------------------------------------------------------------
' Note: the library should only be added after the page is fully loaded. If
' it is loaded when the HTML document is refereshing, it risks adding the script to
' the old HTML document which will be replaced by the new one causing the script to
' be inexistent and resulting in "ReferenceError" due to its functions not defined.
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".jsAddScript"


    If Dir(strLocalScriptPath) = "" Then Exit Sub

    Dim fso As New FileSystemObject
    jsEval fso.OpenTextFile(strLocalScriptPath).ReadAll 'Load the script contents
    printMsg info_, "Successfully added the local script: """ & strLocalScriptPath & """", FromProcedureName

    Set fso = Nothing

End Sub



'***************************************************************************************************
'                              ■■■ タブ操作系 ■■■
'***************************************************************************************************
Public Sub closeTab()
'------------------------------------------------------------------
' Close the current tab of the session. Requires targetId which can
' be retrieved by using getTargetId. This is the targetId of the
' active tab under focus of the automation pipe.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".closeTab"


   'Get the original serialized tab ID
    Dim mainID As String, serial As Dictionary
    Set serial = ShSetting01_StartBrowser.deserializeFromTable(brUser, 2, FromProcedureName)
    mainID = serial.Item("targetID")

    '操作中のタブ自体を閉じないように仕向ける
    If targetID = mainID Then
        printMsg WARN_, "Unable to close this tab " & targetID & " as it is the main session tab", FromProcedureName
        Exit Sub
    End If

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    invokeMethod "Target.closeTarget", params
    printMsg info_, brID & " was closed", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：既に開いているタブのセッションIDを、タブの名前またはURLに基づいて検索します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - CDPBrowser
'* 引数　　：tabName    タブ名
'            Url        URL
'            setMain    たった今使ってる クラスモジュール - CDPBrowserに、この引数の結果に基づいて、セッションIDを置き換えます
'            SearchType "page","iframe" といった、セッションの種類を指定します
'            doRetry    `True`で、出現まで何回かリトライします。
'---------------------------------------------------------------------------------------------------
'* 更新履歴：06/06/23               - Added doRetry argument
'            03/05/23               - Combined the method with the getTabNew
'            24/10/25 Daniel Polak  - Predictable, precise, and script-friendly tab selection that scales well when there are many similar tabs open
'                                     Exact title and URL-prefix matching avoids accidentally selecting tabs with similar titles/URLs
'            23/10/25 Daniel Polak  - Fixed typo timeStart instead of timerStart and fixed bug 60000s instead of 60s
'            01/08/26 Eschamali     - iframeへの切り替えができるように改良
'            22/01/26 Long Vh       - Added .reattach where appropriate
'---------------------------------------------------------------------------------------------------
'* 機能説明：tabName引数とurl引数は部分的な指定(部分一致)が可能です。SearchType はデフォルトは、"page"にしてます。
'* 注意事項：これらの引数を両方とも省略すると、getTabは最も近い未接続のタブに接続しようとします。
'***************************************************************************************************
Public Function getTab(Optional tabName As String, Optional Url As String, Optional setMain As Boolean, Optional SearchTypeID As TargetgetTargetsType, Optional doRetry As Boolean = False) As WebView2Browser
    Const FromProcedureName As String = ThisClassName & ".getTab"


    Dim results As Scripting.Dictionary, Obj As Scripting.Dictionary
    Dim tabId As String, candidateTabId As String, errMsg As String
    Dim targets As Collection
    Dim timerStart As Single
    Dim tabfound As Boolean
    Dim bestScore As Long, score As Long
    Dim info As String

    If doRetry Then timerStart = Timer

    'Typeを決める
    Dim SearchType As String
    Select Case SearchTypeID
        Case TargetgetTargetsType.Page: SearchType = "page"
        Case TargetgetTargetsType.iFrame: SearchType = "iframe"
        Case Else: Err.Raise vbObjectError + 1, FromProcedureName, "このパラメーターは、未定義です"
    End Select

searchit:

    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

    'スコア方式を利用して、理想の対象タブ情報を取得します
    Const MaxScore As Long = 100
    bestScore = -1
    candidateTabId = ""
    For Each Obj In targets
        If Obj("type") = SearchType Then
            score = -1
            Select Case True
                Case tabName <> ""
                    If Obj.Exists("title") Then
                        If Obj("title") = tabName Then
                            score = 100
                        ElseIf InStr(Obj("title"), tabName) Then
                            score = 50
                        Else
                            errMsg = "with name containing """ & tabName & """"
                        End If
                    End If
                Case Url <> ""
                    If Obj.Exists("url") Then
                        If Left$(Obj("url"), Len(Url)) = Url Then
                            score = 90
                        ElseIf InStr(Obj("url"), Url) Then
                            score = 40
                        Else
                            errMsg = "with url containing """ & Url & """"
                        End If
                    End If
                Case Else
                    If Obj("attached") = False Then
                        score = 80
                    Else
                        score = 30
                        errMsg = "as no new unattached tab could be found"
                    End If
            End Select
            If score > bestScore Then
                bestScore = score
                candidateTabId = Obj("targetId")
            End If

            '最大スコアが出てるなら、捜査打ち切り
            If bestScore = MaxScore Then Exit For
        End If
    Next

    If bestScore >= 0 Then
        tabfound = True
        tabId = candidateTabId
    End If

   'Do retry as long as not timeout (default 60s)
    If Not tabfound And doRetry And Timer - timerStart < 60 Then GoTo searchit

   'Returns nothing if not found
   'Nothing is preferred over throwing error
    If Not tabfound Then
        info = IIf(tabName <> "", tabName, IIf(Url <> "", Url, "(no new tab found)"))
        printMsg WARN_, "Unable to find the tab: " & info, FromProcedureName
        Exit Function
    Else
        printMsg info_, "Found the tab with ID: " & tabId & ". (Certainty：" & bestScore & "%)", FromProcedureName
    End If

searchfound:

    Dim sessionID_ As String
   'Return the session object for reference
    sessionID_ = attachToTab(tabId)

   'Return the new CDP object
   '22/01/26 .reattach brUser is needed from v3.1 onwards to enable multi-instances automation
    Set getTab = New WebView2Browser
    getTab.reattach brUser, False   '親子Session関係により、旧SessionIDは破棄しないようにする
    getTab.id = sessionID_  'will also update the targetID
    sleep 'give some time so that the debugging session can switch to the new tab

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = sessionID_
    getSessionInfo
    serialize
    printMsg info_, "Tab [" & targetID & "] is now the main session tab for " & brID, FromProcedureName

End Function

Public Function newTab(Optional Url As String, Optional newWindow As Boolean, Optional setMain As Boolean, Optional isHidden As Boolean, Optional isBackground As Boolean) As WebView2Browser
'------------------------------------------------------------------
' Create a new blank tab within the same instance.
' url: The initial URL the page will be navigated to. An empty
' string indicates about:blank.
' newWindow: If true will open a separate window instead of a tab.
' setMain: If true will set the new tab as the main session tab.
' isBackground: If true, creates the new target in background.
' In the original CDP document, each new tab is called "Target" and
' the equivalence of this method is Target.createTarget
' This returns the Session ID of the newly created tab.
' Updated: 20/03/23 Return the result as a new Browser object
'          26/04/23 Format url argument to accept simplified input
'          22/01/26 Added .reattach where appropriate
'          07/03/26 Added hidden argument
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".newTab"


    Dim params As New Scripting.Dictionary
    Dim sessionID_ As String

    params("url") = CStr(checkURI(Url))
    If newWindow Then params("newWindow") = CBool(newWindow)
    If isHidden Then params("hidden") = CBool(isHidden)
    If isBackground Then params("background") = CBool(isBackground)

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.createTarget", params)

   'Return the new tab's Session ID
    Dim tabId As String
    tabId = results("targetId")
    sessionID_ = attachToTab(tabId)

   'Switch attachment to the new tab
   '22/01/26: From v3.1, reattachment is no longer automatic so .reattach is needed
    printMsg info_, "newTab will now open a new CDP instance", FromProcedureName
    Set newTab = New WebView2Browser
    newTab.reattach brUser, False   '親子Session関係により、旧SessionIDは破棄しないようにする
    newTab.id = sessionID_          '新しくタブ生成した際の`SessionID`として、改めて割り当て
    If Url <> "" Then newTab.wait

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = sessionID_
    getSessionInfo
    serialize
    printMsg info_, "Tab [" & brName & " | " & targetID & "] is now the main session tab for " & brID, FromProcedureName

End Function



'***************************************************************************************************
'                              ■■■ デバッグメソッド集 ■■■
'***************************************************************************************************
'* 機能　　：このクラスオブジェト内部で持っている、現在のセッションのすべてのパラメータを表示します
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前にログレベルを`DEBUG`にする必要があります
'***************************************************************************************************
Public Sub printParams()
    Const FromProcedureName As String = ThisClassName & ".printParams"


    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, brID & " Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    printMsg Debug_, "Session ID        | sessionID    | " & sessionID, FromProcedureName
    printMsg Debug_, "Last ID           | lngLastID    | " & lngLastID, FromProcedureName
    printMsg Debug_, "Window Handle     | brHWnd       | " & brHWnd, FromProcedureName
    printMsg Debug_, "Window ID         | brWId        | " & brWID, FromProcedureName
    printMsg Debug_, "Browser Type      | brType       | " & brType, FromProcedureName
    printMsg Debug_, "Browser Ver       | brVer        | " & brVer, FromProcedureName
    printMsg Debug_, "Browser Path      | brPath       | " & brPath, FromProcedureName
    printMsg Debug_, "Tab URL           | brUrl        | " & brURL, FromProcedureName
    printMsg Debug_, "Tab Title         | brName       | " & brName, FromProcedureName
    printMsg Debug_, "Tab ID            | targetID     | " & targetID, FromProcedureName
    printMsg Debug_, "Core dwProcessId  | dwProcessId  | " & core.dwProcessId, FromProcedureName
    printMsg Debug_, "CDP Version       | CDPVer       | " & CDPVer, FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：内部で持っているタブ情報を表示します
'---------------------------------------------------------------------------------------------------
'* 引数　　：AllShow    `True` ：全ての`type`を表示させます
'                       `False`：`type`が`page`のみ表示させます(規定値)
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前にログレベルを`DEBUG`にする必要があります
'***************************************************************************************************
Public Sub printTabs(Optional AllShow As Boolean)
    Const FromProcedureName As String = ThisClassName & ".printTabs"


    Dim results As Scripting.Dictionary, target_ As Scripting.Dictionary
    Dim i As Long
    Dim targets As Collection
    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

   'Read info
    printMsg Debug_, "---------------------------------------------------", FromProcedureName
    printMsg Debug_, "Tabs Information", FromProcedureName
    printMsg Debug_, "---------------------------------------------------", FromProcedureName

    i = 1
    For Each target_ In targets
        If target_("type") = "page" Or AllShow Then
            printMsg Debug_, "#" & i & vbTab & _
                        "type: " & target_("type") & " | " & _
                        "attached: " & target_("attached") & " | " & _
                        "title: " & target_("title") & " | " & _
                        "url: " & target_("url") & " | " & _
                        "target ID: " & target_("targetId"), FromProcedureName
            i = i + 1
        End If
    Next

End Sub

'***************************************************************************************************
'* 機能　　：ログレベル付きのスマートなログをイミディエイトへ出力させるヘルパーメソッドです
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    Dim logName As String, strFormattedMsg As String

    If isHeader Then strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-") _
    Else strFormattedMsg = " | " & brID & " | " & strMsg

   'Perform msg logging if needed
    If doLog Then
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".WebView2.log"
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

'=============================================================================
' PUBLIC PROPERTIES
'=============================================================================

Property Get Url() As String
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    brURL = jsEval("document.URL")
    Url = brURL

End Property

Property Let Url(newURL As String)
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    navigate newURL
    brURL = newURL

End Property

Property Get Name() As String
'------------------------------------------------------------
' Get the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    brName = jsEval("document.title")
    Name = brName

End Property

Property Let Name(newName As String)
'------------------------------------------------------------
' Set the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".Name"


    jsEval "document.title = '" & newName & "'"
    printMsg info_, "Tab name changed to " & newName, FromProcedureName
    brName = newName

End Property

Property Get html() As String
'------------------------------------------------------------
' Get the html text of the entire document
' Added 02/03/2023
'------------------------------------------------------------

    html = jsEval("document.documentElement.innerHTML")

End Property

Property Get id() As String
'------------------------------------------------------------
' Get the session ID of the current browser
' Added 02/03/2023
'------------------------------------------------------------

    id = sessionID

End Property

Property Let id(newID As String)
'------------------------------------------------------------
' Set the session ID of the current browser to a new value
' Added 02/03/2023
'------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".id"


    sessionID = newID
    getSessionInfo        'update all session infos
    printMsg info_, brID & " updated Session ID to " & sessionID, FromProcedureName

End Property

'***************************************************************************************************
'* 機能　　：このクラスで定義中のtargetIDを取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：「どのタブか」を特定するための専用ID
'***************************************************************************************************
Property Get CurrentTargetID() As String
    CurrentTargetID = targetID
End Property

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


Public Sub ProcessMessages(Optional ByVal iterations As Long = 100)
    If core Is Nothing Then Exit Sub
    core.ProcessMessages iterations
End Sub

'----------------------------------------------------------------------
' DevToolsProtocol event receiver 購読（例: Network.responseReceived）
'----------------------------------------------------------------------
Public Function SubscribeDevToolsProtocolEvent(ByVal eventName As String) As Boolean
    If core Is Nothing Then
        SubscribeDevToolsProtocolEvent = False
    Else
        SubscribeDevToolsProtocolEvent = core.SubscribeDevToolsProtocolEvent(eventName)
    End If
End Function

'----------------------------------------------------------------------
' DevToolsProtocol event receiver 購読解除
'----------------------------------------------------------------------
Public Function UnsubscribeDevToolsProtocolEvent(Optional ByVal eventName As String) As Boolean
    If core Is Nothing Then
        UnsubscribeDevToolsProtocolEvent = True
    Else
        UnsubscribeDevToolsProtocolEvent = core.UnsubscribeDevToolsProtocolEvent(eventName)
    End If
End Function

Private Sub core_CDPEventReceived(methodName As String, eventParamsJson As String)
    RaiseEvent detectionCDPEvent(methodName, eventParamsJson)
End Sub

'***************************************************************************************************
'* 機能　　：このClassで起動したブラウザのウィンドウハンドルを取得します
'---------------------------------------------------------------------------------------------------
'* 引数　　：requestUpdate      `True` :必ずウィンドウハンドルを調査したうえで、ウィンドウハンドルを取得します
'                               `False`:既存の値があればそれを流用し、なければ再調査します
'* 返値    ：ブラウザのウィンドウハンドル
'---------------------------------------------------------------------------------------------------
'* 機能説明：このハンドルを使った高度なウィンドウの調整などができます
'***************************************************************************************************
Property Get BrowserWindowHandle(Optional alwaysRequest As Boolean) As LongPtr
    If alwaysRequest Then
        BrowserWindowHandle = getHandle
    Else
        '既存の値があれば、それを流用、ない場合は、特定処理に入ります
        If brHWnd Then BrowserWindowHandle = brHWnd Else BrowserWindowHandle = getHandle
    End If
End Property

'***************************************************************************************************
'* 機能　　：このClassで起動したブラウザのウィンドウIDを取得します
'---------------------------------------------------------------------------------------------------
'* 引数　　：requestUpdate      `True` :必ずウィンドウIDを調査したうえで、ウィンドウIDを取得します
'                               `False`:既存の値があればそれを流用し、なければ再調査します
'* 返値    ：ブラウザのウィンドウID
'---------------------------------------------------------------------------------------------------
'* 機能説明：このIDを使ったCDP-Jsonとしてのウィンドウ調整("Browser.setWindowBounds"等)ができます
'***************************************************************************************************
Property Get BrowserWindowID(Optional alwaysRequest As Boolean) As Long
    If alwaysRequest Then
        BrowserWindowID = invokeMethod("Browser.getWindowForTarget").Item("windowId")
    Else
        '既存の値があれば、それを流用、ない場合は、特定処理に入ります
        If brWID Then BrowserWindowID = brWID Else BrowserWindowID = invokeMethod("Browser.getWindowForTarget").Item("windowId")
    End If
End Property



'***************************************************************************************************
'                           ■■■ 要素特定系プロシージャ ■■■
'***************************************************************************************************
Public Function getElementByID(strID As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class from the element ID, similar to the getElementByID of JS.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Get the object (no global variable assignment - objectId only)
    Dim varJS As String: Dim result As Variant
    varJS = "document.getElementById(""" & strID & """)"
    result = jsEval(varJS, returnByValue:=False)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByID = New CDPElement
    getElementByID.Init Me, varJS, VarType(result), oid

End Function

Public Function getElementByXPath(strXPath As String) As CDPElement
'---------------------------------------------------------------------------------------
' Get an element class from XPath. The element class is init in CDPElement module.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' XPath Syntax help: https://www.w3schools.com/xml/xpath_syntax.asp
' Updated: 02/03/2023
' Examples -----------------------------------------------------------------------------
' eg.1 "//b[text()='T-BILL CURVE']": tag "b" with innertext to be 'T-BILL CURVE'
' eg.2 "//input[@name='username']": tag "input" & attribute "name" to be 'username'
' eg.3 "//a[@a1='abc' and @a2='xyz']": tag "a" with attribute a1 = abc and a2 = xyz
' eg.4 "//div[@id='abc']/div": get first div child of the div with id = abc
' eg.5 "//div[@id='abc'].//div": get first div descendant of the div with id = abc
' eg.6 "div": get first div child from the prescribed contextNode (default = 'document')
'---------------------------------------------------------------------------------------

   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")

   'Get the object (no global variable assignment - objectId only)
    Dim varJS As String: Dim result As Variant
    varJS = "document.evaluate(""" & strXPath & """, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
    result = jsEval(varJS, returnByValue:=False)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByXPath = New CDPElement
    getElementByXPath.Init Me, varJS, VarType(result), oid

End Function

Public Function getElementsByXPath(strXPath As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by XPath. For further info on the particular XPath
' syntax used for retrieving the elements, refer to Mozilla doc below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")

   'Get the XPathResult objectId (ORDERED_NODE_SNAPSHOT_TYPE)
    Dim collResult As Variant
    collResult = jsEval("document.evaluate(""" & strXPath & """, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)", returnByValue:=False)
    If VarType(collResult) <> vbString Then Exit Function
    Dim collOid As String: collOid = CStr(collResult)

   'Get snapshotLength via callFunctionOn
    Dim elements As Collection: Set elements = New Collection
    Dim length_ As Long, N As Long
    length_ = jsEval("function(){ return this.snapshotLength }", objectId:=collOid)

   'Get each item's objectId via callFunctionOn
    Dim tElement As CDPElement
    Dim itemArgs As Collection, itemArg As Scripting.Dictionary
    Dim itemResult As Variant, itemOid As String
    For N = 0 To length_ - 1
        Set itemArgs = New Collection
        Set itemArg = New Scripting.Dictionary
        itemArg.Add "value", N
        itemArgs.Add itemArg
        itemResult = jsEval("function(n){ return this.snapshotItem(n) }", objectId:=collOid, objectArguments:=itemArgs, returnByValue:=False)
        itemOid = "": If VarType(itemResult) = vbString Then itemOid = CStr(itemResult)
        Set tElement = New CDPElement
        tElement.Init Me, , VarType(itemResult), itemOid
        elements.Add tElement
    Next

    Set getElementsByXPath = elements

End Function

Public Function getElementByQuery(strQuery As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class by QuerySelector - similar to that of the same in JS.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation. Refer to the below for some examples.
' https://www.w3schools.com/jsref/met_document_queryselector.asp
' https://stackoverflow.com/a/29937844
' https://www.w3.org/TR/selectors-3/#context
' Updated: 08/11/2023
' Examples ------------------------------------------------------------------------
' eg.1 "#abc": get the first element with id attribute = abc
' eg.2 "tb.xyz": get the first element with tag tb and with class xyz
' eg.3 "a[name='abc']": get the first element with tag a and attribute name = abc
' eg.4 "div > p": get the first element with tag 'p' and has parent with tag 'div'
' eg.5 "[a1='abc'][a2='xyz']": get element with attribute a1 = abc and a2 = 123
' eg.6 "[a1='abc'] [a2='xyz']": get element with attribute a2 = xyz and is a child
' of an element with attribute a1 = abc
' eg.7 "#processType [id*='select_option']": get child element with id text that
' contains "select_option" and is a child of element with id "processType"
' eg.8 "#upload:not([disabled])": get the first element with id=upload and does not
' have attribute 'disabled'.
' eg.9 "#upload:disabled": get the first element with id=upload and has attribute
' 'disabled'.
' eg.10 "#id1 tr:last-child [id*='invoiceno']": first get element with id text
' 'id1', then look for the last child with tag 'tr', then finally look within the
' child's children for the first element with id containing 'invoiceno'.
'----------------------------------------------------------------------------------

   'Get the object (no global variable assignment - objectId only)
    Dim varJS As String: Dim result As Variant
    varJS = "document.querySelector(""" & strQuery & """)"
    result = jsEval(varJS, returnByValue:=False)

   'Return result
    Dim oid As String: If VarType(result) = vbString Then oid = CStr(result)
    Set getElementByQuery = New CDPElement
    getElementByQuery.Init Me, varJS, VarType(result), oid

End Function

Public Function getElementsByQuery(strQuery As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by QuerySelectorAll. The syntax is as below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Get the NodeList objectId
    Dim collResult As Variant
    collResult = jsEval("document.querySelectorAll(""" & strQuery & """)", returnByValue:=False)
    If VarType(collResult) <> vbString Then Exit Function
    Dim collOid As String: collOid = CStr(collResult)

   'Get length via callFunctionOn
    Dim result As Variant, elements As Collection: Set elements = New Collection
    Dim length_ As Long, N As Long
    length_ = jsEval("function(){ return this.length }", objectId:=collOid)

   'Get each item's objectId via callFunctionOn
    Dim tElement As CDPElement
    Dim itemArgs As Collection, itemArg As Scripting.Dictionary
    Dim itemResult As Variant, itemOid As String
    For N = 0 To length_ - 1
        Set itemArgs = New Collection
        Set itemArg = New Scripting.Dictionary
        itemArg.Add "value", N
        itemArgs.Add itemArg
        itemResult = jsEval("function(n){ return this.item(n) }", objectId:=collOid, objectArguments:=itemArgs, returnByValue:=False)
        itemOid = "": If VarType(itemResult) = vbString Then itemOid = CStr(itemResult)
        Set tElement = New CDPElement
        tElement.Init Me, , VarType(itemResult), itemOid
        elements.Add tElement
    Next

    Set getElementsByQuery = elements

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
'            alwaysBrowserContext   `True` ：`sessionID`の有無問わず、常にブラウザ自身を対象したコマンド実行モードとなります→CallDevToolsProtocolMethodAsync or CallDevToolsProtocolMethod
'                                   `False`：`sessionID`があれば、その`SessionID`に対するコマンド実行。→CallDevToolsProtocolMethodForSessionAsync or CallDevToolsProtocolMethodForSession
'                                            なければ、ブラウザ自身を対象したコマンド実行となります。→CallDevToolsProtocolMethodAsync or CallDevToolsProtocolMethod
'
'            Async                  `True` ：コマンドの結果を待ちません。Do Loopフェーズには入りません。
'                                   `False`：コマンドの結果を待ちます。Do Loopフェーズに入り、そのIDが来るまで待ちます。この場合、非同期に実行した結果のID回収もついでに行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：・引数-Async の値によって、Dictionaryの構造が変わる特殊なメソッドです。別途公開メソッドで、役割分担しております
'            ・非同期イベントの取得は、分離しました。ここではその取得はしません。
'***************************************************************************************************
Private Function SendMessage(methodName As String, paramsJsonString As String, Optional alwaysBrowserContext As Boolean, Optional Async As Boolean) As Scripting.Dictionary
    Const FromProcedureName As String = ThisClassName & ".SendMessage"


    Dim strRes As String
    printMsg Debug_, "sendCDP-Command：method = `" & methodName & "` , paramsJson = `" & paramsJsonString & "` , sessionID = " & sessionID, FromProcedureName
    printMsg Debug_, "alwaysBrowserContext：" & alwaysBrowserContext, FromProcedureName
    
    'コマンド結果が来るまで待機するか？
    If Async Then
        '投げっぱなしにして、即次へ
        Dim CommandID As Long
        Dim hrAsync As Long
        CommandID = lngLastID
        lngLastID = lngLastID + 1

        If alwaysBrowserContext Then
            printMsg Trace_, "Use CallDevToolsProtocolMethodAsync.", FromProcedureName
            hrAsync = core.CallDevToolsProtocolMethodAsync(methodName, paramsJsonString, CommandID)
        Else
            If Len(sessionID) > 0 Then
                printMsg Trace_, "Use CallDevToolsProtocolMethodForSessionAsync.", FromProcedureName
                hrAsync = core.CallDevToolsProtocolMethodForSessionAsync(sessionID, methodName, paramsJsonString, CommandID)
            Else
                printMsg Trace_, "Use CallDevToolsProtocolMethodAsync.", FromProcedureName
                hrAsync = core.CallDevToolsProtocolMethodAsync(methodName, paramsJsonString, CommandID)
            End If
        End If
        
        Set SendMessage = New Dictionary
        SendMessage.Add CommandID, hrAsync

    Else
        '内部で、コマンド結果のコールバックが来るまで待ち続けます
        If alwaysBrowserContext Then
            printMsg Trace_, "Use CallDevToolsProtocolMethod.", FromProcedureName
            strRes = core.CallDevToolsProtocolMethod(methodName, paramsJsonString)
        Else
            If Len(sessionID) > 0 Then
                printMsg Trace_, "Use CallDevToolsProtocolMethodForSession.", FromProcedureName
                strRes = core.CallDevToolsProtocolMethodForSession(sessionID, methodName, paramsJsonString)
            Else
                printMsg Trace_, "Use CallDevToolsProtocolMethod.", FromProcedureName
                strRes = core.CallDevToolsProtocolMethod(methodName, paramsJsonString)
            End If
        End If

        If seeRawSendMsgDbg Then printMsg Debug_, "strRes = " & strRes, FromProcedureName

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
'* 機能　　：URIのプロトコルチェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数の内容そのまんま
'* 引数　　：URI   ブラウザのURLに渡す予定の文字列
'---------------------------------------------------------------------------------------------------
'* 注意事項：・セキュリティ重視により、ホワイトリスト外のプロトコルが来たら、エラーで止まります
'            ・既に開いてるURIに関して(`getTab`メソッド)はチェック対象外とします
'***************************************************************************************************
Private Function checkURI(URI As String) As String
    Const FromProcedureName As String = ThisClassName & ".checkURI"


    '0. 空文字の場合は、チェックしません
    printMsg Debug_, "Checking the protocol of the URI...", FromProcedureName
    If URI = "" Then
        printMsg Debug_, "Checking is skipped because the URI is empty.", FromProcedureName

        checkURI = URI
        Exit Function
    End If

    '1. ホワイトリストのプロトコルを取得
    With ShSetting01_StartBrowser
        Dim tmp, WhiteList
        tmp = .ListObjects(.UseTableName(3, FromProcedureName)).DataBodyRange.Columns(Table03ColumnID.接頭辞)

        '配列か？
        If IsArray(tmp) Then
            'そのまんま
            WhiteList = tmp
        Else
            '配列じゃない場合は配列にしておく
            ReDim WhiteList(1 To 1, 1 To 1)
            WhiteList(1, 1) = tmp
        End If

        '解放
        tmp = vbNullString
    End With

    '2. 先頭始まりかチェック
    Dim i As Long
    For i = 1 To UBound(WhiteList)
        '先頭始まりなら、引数の内容そのまんまを返して抜ける
        If InStr(1, URI, WhiteList(i, 1)) = 1 Then
            printMsg Debug_, "Protocol check passed：" & WhiteList(i, 1), FromProcedureName

            checkURI = URI
            Exit Function
        End If
    Next

    '3. プロトコル名をチェック
    Dim ProtocolName, ErrorMes As String
    ProtocolName = Split(URI, ":")
    If UBound(ProtocolName) = 0 Then
        ErrorMes = "URLの形式が正しくありません。" & vbCrLf & "安全のため省略せず、`https://`や`file://`のような、プロトコル名を、URLの先頭に付けてください。" & vbCrLf & vbCrLf & "[問題のURL]" & vbCrLf & URI
    Else
        ErrorMes = "プロトコル`" & ProtocolName(0) & "`は、ホワイトリストに登録されていません。" & vbCrLf & "意図しないアプリケーションの起動などを防ぐため、処理を中断しました。" & vbCrLf & vbCrLf & "【解決策】" & vbCrLf & "このプロトコルへのアクセスが、本当に必要である場合は、シート「ブラウザ起動設定」のテーブル「ProtocolWhiteList」に、`" & ProtocolName(0) & "`を追加してから、もう一度、お試しください。"
    End If

    '4. エラーで中断
    printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True

End Function

Private Function attachToTab(tabId As String) As String
'----------------------------------------------------------------------------------
' Attach to a specific tab with its tabId (a.k.a targetId)
' flatten: Enables "flat" access to the session via specifying sessionId attribute
' in the commands. Set to True by default to make attachToTab work.
' Returns: a session ID string that is needed for CDP automation.
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(tabId)
    params("flatten") = True

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.attachToTarget", params)
    attachToTab = results("sessionId")

End Function



'***************************************************************************************************
'                           ■■■ 情報収集系プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：セッション情報の更新を行います
'***************************************************************************************************
Private Sub getSessionInfo()
    Const FromProcedureName As String = ThisClassName & ".getSessionInfo"


   'Activate to getHandle correctly
    activate False

   'Get infos
    brURL = Url
    brName = Name
    brHWnd = getHandle

    '対象タブが非表示の場合、`Browser.getWindowForTarget`が効かなくなるので、判定をかます
    Dim tmp As Dictionary: Set tmp = invokeMethod("Browser.getWindowForTarget", StopError:=False)
    If tmp Is Nothing Then printMsg info_, "The tab you are connecting to may be hidden.", FromProcedureName Else brWID = tmp.Item("windowId")

    brType = invokeMethod("Browser.getVersion").Item("product")
    brVer = Split(brType, "/")(1)
    brPath = getBrowserPath(brType)
    targetID = getTargetID

End Sub

'***************************************************************************************************
'* 機能　　：現在のTargetIDを取得します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Retrieve the targetId of the current tab (a.k.k "target" in CDP Documentation).
'            This targetId is useful for tab automation, such as the .closeTab method.
'***************************************************************************************************
Private Function getTargetID() As String

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.getTargetInfo")
    getTargetID = results("targetInfo").Item("targetId")

End Function

'***************************************************************************************************
'* 機能　　：WebView2実行パスを取得します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：レジストリから、インストールパスを特定します
'***************************************************************************************************
Private Function getBrowserPath(strBrowserType As String) As String
    Const FromProcedureName As String = ThisClassName & ".getBrowserPath"


    '一択しかないと思うが一応
    Select Case True
        Case InStr(LCase(strBrowserType), "edg"): brType = brWebView2         'On edge it is just "Edg"
        Case Else: printMsg ERROR_, "Unable to getBrowserPath for browser type: " & brType, FromProcedureName, doRaiseError:=True
    End Select

    'レジストリ操作に必要な定数/オブジェクトの定義
    Dim RegistryCheck As Object: Set RegistryCheck = GetObject("winmgmts:\\.\root\default:StdRegProv")
    Const HKEY_LOCAL_MACHINE    As Long = &H80000002           'このPCの、"すべて"のユーザーに影響する、強力な設定
    Const HKEY_CURRENT_USER     As Long = &H80000001           '"今、ログインしているユーザー"だけに、影響する設定
    Const wbemNoErr             As Long = 0                    '存在するときは、`0`を返します。　https://learn.microsoft.com/ja-jp/windows/win32/api/wbemdisp/ne-wbemdisp-wbemerrorenum#constants

    'WebView2に関する情報のあるパスを設定します
    Const BaseKeyPath As String = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft EdgeWebView"

    '必要な変数を用意
    Dim checkValueName  As String   '値の名前
    Dim ValueExists     As Long     '存在有無結果
    Dim ValueContents   As Variant  '値の中身


    ' ------------ 1. WebView2のインストールパス ------------
    Dim InstallLocation As String
    checkValueName = "InstallLocation"

    '1-1. 文字列として、チェックする
    ValueExists = RegistryCheck.GetStringValue(HKEY_LOCAL_MACHINE, BaseKeyPath, checkValueName, ValueContents)

    '1-2. そのデータ名があったら、取得
    If ValueExists = wbemNoErr Then InstallLocation = ValueContents


    ' ------------ 2. WebView2のバージョン ------------
    Dim Version As String
    checkValueName = "Version"

    '2-1. 文字列として、チェックする
    ValueExists = RegistryCheck.GetStringValue(HKEY_LOCAL_MACHINE, BaseKeyPath, checkValueName, ValueContents)

    '2-2. そのデータ名があったら取得する
    If ValueExists = wbemNoErr Then Version = ValueContents


    '返却
    getBrowserPath = InstallLocation & "\" & Version & "\" & brWebView2

End Function

'***************************************************************************************************
'* 機能　　：起動時の WebView2 ハンドルを取得します
'***************************************************************************************************
Private Function getHandle() As LongPtr
    getHandle = core.ParentHwnd
End Function



'***************************************************************************************************
'                           ■■■ ウィンドウ処理関連プロシージャ ■■■
'***************************************************************************************************
Public Sub show(Optional state As WinState = asNormal, Optional xywh As String = "0 0 0 0")
'-----------------------------------------------------------------------------------
' Display the window with various optional state. The default state is "normal".
' Can also use after .attach to bring the window topmost as well. By default at
' start, the browser window's state will be that of the previous session run.
' xywh is used for inputing left top width height parameters:
' eg.1: xywh = "100 200" left = 100 and top = 200, width and height unchanged
' eg.2: xywh = "100 100 300 400" left = top = 100, width = 300 and height = 400
' eg.3: xywh = "0 0 0 0" the method will ignore and not make any resizing
'-----------------------------------------------------------------------------------
' Note: This method targets the window based on its name so it is best applied when
' the browser window is not refreshing a URL which may potentially change the name.
' Updated: 02/03/23 - Changed CInt to Int to prevent Overflow error.
'-----------------------------------------------------------------------------------

    If brHWnd = 0 Then brHWnd = getHandle
    ShowWindow brHWnd, doShow 'Make visible
    ShowWindow brHWnd, state

   'Resize to specific position and size
   'Will not work properly if window is currently maximized
    If xywh <> "0 0 0 0" Then

        sleep 0.3      'Sleep a bit for previous Win Api executions to finish, else the resize effect will be accidentally cancelled out

        Dim params As New Scripting.Dictionary
        If brWID = "" Then brWID = invokeMethod("Browser.getWindowForTarget").Item("windowId")
        params("windowId") = Int(brWID)     '02/03/23 CInt() will throw error, Int() is better

        Dim bounds As New Scripting.Dictionary
        Dim x As Long, y As Long, w As Long, h As Long
        On Error Resume Next
        x = CLng(Split(xywh, " ")(0))
        y = CLng(Split(xywh, " ")(1))
        w = CLng(Split(xywh, " ")(2))
        h = CLng(Split(xywh, " ")(3))
        On Error GoTo 0

        If x <> 0 Then bounds("left") = x
        If y <> 0 Then bounds("top") = y
        If w <> 0 Then bounds("width") = w
        If h <> 0 Then bounds("height") = h
        Set params("bounds") = bounds
        invokeMethod "Browser.setWindowBounds", params

    End If

End Sub

Public Sub hide()
'------------------------------------------------------------------
' This is the equivalent of show "hidden" but this function is
' added to make it easy to access and use the same feature.
'------------------------------------------------------------------

    If brHWnd = 0 Then brHWnd = getHandle
    ShowWindow brHWnd, asHidden

End Sub



'***************************************************************************************************
'                           ■■■ 再接続関連プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：既存(Public の Class WebView2Core)のオブジェクトを参照して、処理を再開させます
'---------------------------------------------------------------------------------------------------
'* 注意事項：VBE編集が入ると機能しなくなります
'***************************************************************************************************
Public Function reattach(userProfile As String, Optional detachCurrent As Boolean = True) As Boolean
    Const FromProcedureName As String = ThisClassName & ".reattach"


    brUser = userProfile
    deserialize brUser, detachCurrent

    If isLive Then
        reattach = True
        printParams
    Else
        reattach = False
        printMsg WARN_, "Failed to reattach to previous CPID.", FromProcedureName
    End If

End Function

Private Sub serialize()
    Const FromProcedureName As String = ThisClassName & ".serialize"


    'データ記録準備
    Dim excelSerialize As String
    With WorksheetFunction
        excelSerialize = "{""" & .TextJoin(""";""", False, _
                                    .TextJoin(""",""", False, brUser, "lngLastID", lngLastID), _
                                    .TextJoin(""",""", False, brUser, "sessionID", sessionID), _
                                    .TextJoin(""",""", False, brUser, "targetID", targetID) _
        ) & """}"
    End With
 
    'テーブルに、ブラウザセッション情報を記録する
    'Evaluate関数を利用して、上記で作った記述方式で簡易的な2次元配列を作成します
    ShSetting01_StartBrowser.serializeToTable(2, FromProcedureName) = Evaluate(excelSerialize)

End Sub

Private Sub deserialize(UserName As String, detachCurrent As Boolean)
    Const FromProcedureName As String = ThisClassName & ".deserialize"


    On Error GoTo iferr

    Dim objSerialize As Scripting.Dictionary

    '`UserName`に紐づく(予定)オブジェクト参照を行う
1   Set core = g_WebView2Core

    'そのSetしたオブジェクトは、生きてるか？
    If core Is Nothing Then
        'VBEをいじったことによる参照切れ
        '※そもそも、オブジェクトが切れたら何もできないので、Err.Raiseで止めちゃう
        printMsg ERROR_, "Failed to reattach to previous WebView2.", FromProcedureName, doRaiseError:=True
    Else
        '生きてるので改めて、情報を取り直す
        printMsg info_, "An existing WebView2 reference object was found.", FromProcedureName

        getSessionInfo
        printParams
    End If

    'Excelテーブルから、`UserName`に紐づくブラウザセッション情報を読み込む
2   Set objSerialize = ShSetting01_StartBrowser.deserializeFromTable(UserName, 2, FromProcedureName)

    'クラス内変数に適用
3   'lngLastID = objSerialize.Item("lngLastID")
4   targetID = objSerialize.Item("targetID")

    '旧SessionIDの破棄を試みる
    '※`getTab`,`newTab`からの場合は破棄すると、親子関係Sessionの影響で道ずれ破棄されるので、破棄しないように引数で制御します
    If detachCurrent Then
        Dim ResultDetach As Dictionary, CDPparameters As Dictionary
5       Dim oldSessionID As String: oldSessionID = objSerialize.Item("sessionID")
        Set CDPparameters = New Dictionary

        CDPparameters.Add "sessionId", oldSessionID
        printMsg Debug_, "Discard the old session ID... - " & oldSessionID, FromProcedureName
6       Set ResultDetach = invokeMethod("Target.detachFromTarget", CDPparameters, StopError:=False)

        '結果を判定
        '※エラーによるエラーメッセージは、`invokeMethod`が担います
7       If Not (ResultDetach Is Nothing) Then printMsg info_, "The old session ID has been detached. - " & oldSessionID, FromProcedureName

        'SessionIDを再生成し、適用
8       sessionID = attachToTab(targetID)
        printMsg info_, "A new session ID has been assigned. - " & sessionID, FromProcedureName

        '記録しておく
        '※旧SessionIDの破棄をしない場合は、上書きせずそのままにします
9       ShSetting01_StartBrowser.serializeSingleToHandleInfoTable(UserName, Table02RowID.sessionID, FromProcedureName) = sessionID
    Else
        '旧SessionIDを再利用
        sessionID = objSerialize.Item("sessionID")
        printMsg info_, "Reattached by reusing the Session ID. - " & sessionID, FromProcedureName
    End If
    Exit Sub

iferr: 'if deserialize fails, simply move on to starting anew

    printMsg ERROR_, "deserialization failed at line " & Erl & " for serial string " & WorksheetFunction.TextJoin(",", True, "lngLastID:" & lngLastID, "targetID:" & targetID, "sessionID:" & sessionID), FromProcedureName

End Sub

Public Function isLive() As Boolean
'-----------------------------------------------------------------------------------
' Check if there is ongoing connection with an attached instance
' Original Name: connectionAlive
'-----------------------------------------------------------------------------------

    On Error GoTo iferr
    Dim strLoc As String
    strLoc = jsEval("window.location.href", dbgMsg:=False)
    isLive = True
    Exit Function

iferr:
    isLive = False

End Function




'***************************************************************************************************
'                                    ■■■ 後始末系 ■■■
'***************************************************************************************************
Private Sub cleanUpSessions()
'-----------------------------------------------------------------------------------------------------------------------------------------
' Provides cleaning before firing up a new session to prevent
' pipe error
' Update: 05/11/24: Daniel Polak - modified to solve problems with closing sessions (that do not belong to this user) on terminal server
'         23/10/25: Daniel Polak - p.name and p.terminate were not correct, replaced by proc.name and proc.terminate (discovered this bug because we are using Option Explicit now). On Error Resume Next was hiding the bug too.
'                                - only terminate automation-processes (where command line contains --remote-debugging-pipe)
'-----------------------------------------------------------------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".cleanUpSessions"


    '1. 今、Excelが動いてるプロセスIDとそのログインユーザーのセッションIDを取得します
    Dim lProcessId As Long, lSessionId As Long
    lProcessId = GetCurrentProcessId()          'Get the current(Excel) process ID
    ProcessIdToSessionId lProcessId, lSessionId 'Get the Session ID for the current(Excel) process

    'Clean up all running sessions
    Dim proc As Object
    On Error Resume Next

    '2. 前述の2つの情報を基に、起動中のアプリたちを取得する
    For Each proc In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process Where SessionId = " & lSessionId)
        '3. 起動中のブラウザ名であるか確認
        If LCase$(proc.Name) = LCase$(brType) Then
            '4. 対象ユーザー名なら、Killします
            If InStr(1, proc.CommandLine, brUser) > 0 Then
                 proc.Terminate
            End If
        End If
        DoEvents
    Next
    On Error GoTo 0
    Err.Number = 0

    'Also clean up all prev CDP
    Call deleteCDPID

   'Clean core params for new session
    lngLastID = 0   'To make a clean slate for new session
    sessionID = ""  'Likewise

    printMsg info_, "cleanUpSessions completed and previous CdpSessionID removed", FromProcedureName

End Sub

Private Sub deleteCDPID()
'----------------------------------------------------------------------------------
' By      : Eschamali
' 機能    ：デバッグブラウザのパイプハンドル/セッション情報を消去します
'----------------------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".deleteCDPID"


    With ShSetting01_StartBrowser
        .DeleteCDPToTable(2, FromProcedureName) = brUser  'ブラウザセッション情報
    End With

End Sub

'***************************************************************************************************
'* 機能　　：WebView2を終了させます
'***************************************************************************************************
Public Sub quit()
    Const FromProcedureName As String = ThisClassName & ".quit"


    printMsg info_, "Exiting WebView2...", FromProcedureName
    core.quit
    printMsg info_, "WebView2 has been closed.", FromProcedureName
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
