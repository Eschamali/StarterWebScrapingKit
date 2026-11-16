Attribute VB_Name = "Demo_OfficeScript"
'==============================================================================================================
'         OfficeScriptをCDP経由で動かすためのDemoです。コードエディタが使えない環境ではご利用できません
'
'       ニーズ/安全上、以下の経路のみ載せています
'       ・普段使いのブラウザで、「https://excel.cloud.microsoft/」から任意のブックを開いて、コードエディタも開けてる状態
'       ・DesktopExcelのコードエディタからのWebView2操作
'==============================================================================================================
Option Explicit
Option Private Module



'***************************************************************************************************
'                                   ■■■ 共通設定用 ■■■
'***************************************************************************************************
Private Const UseWeb            As Boolean = False   'True:Web版Excelで　False:Desktop版Excelで



'***************************************************************************************************
'                 ■■■ 長いJavaScriptコードを見やすくするためのヘルパー変数 ■■■
'***************************************************************************************************
Private JSList()        As String
Private CurrentIndex    As Long
Private SetSplit        As String



'***************************************************************************************************
'          ■■■ Demo_ViewBookInfo(デスクトップ版)の結果を、マクロ終了後に OnTime で回収する用 ■■■
'***************************************************************************************************
Private PendingEditor       As CDPContext   '結果を回収し終わるまで、コードエディターへの接続を保つ
Private PendingStartSec     As Double       '実行を始めた時刻(Timer)
Private PendingTimeoutSec   As Double       '待つ上限(秒)
Private PendingPollSec      As Long         '覗く間隔(秒)
Private PendingSheetName    As String       '書き出し先のシート名



'***************************************************************************************************
'                                   ■■■ 接続儀式 ■■■
'***************************************************************************************************
'* 機能　　：環境変数を適用させた状態のコードエディターを起動させます
'---------------------------------------------------------------------------------------------------
'* 注意事項：現状、初回起動時はこれを2回、叩かないとうまくいかないようです
'***************************************************************************************************
Private Sub DebugModeCodeEditorInit()
    '1. Portを開ける環境変数を適用
    Dim WebView2Debug As CDPCoreViaWebSocket: Set WebView2Debug = New CDPCoreViaWebSocket
    WebView2Debug.EnsureWebView2DebugPort = 0

    '2. WebView2をクリーン処理させるパスを指定する
    Dim DropWebView2 As WebView2Host: Set DropWebView2 = New WebView2Host
    BrowserCapabilities.Set_UserDataDirName = Environ("LOCALAPPDATA") & "\Microsoft\Office\16.0\Wef\webview2\2\1"
    DropWebView2.Init BrowserCapabilities, 0
    DropWebView2.cleanUpSessions
    Set BrowserCapabilities = Nothing

    '3. コードエディタを、遅延起動
    CommandBars.ExecuteMso "CreateInEditor"
End Sub

'***************************************************************************************************
'* 機能　　：起動済み「コードエディタ」に接続します
'***************************************************************************************************
Private Function ConnectForWebExcel(WebMode As Boolean) As CDPContext
    '1. 探す目印となるUserNameを指定
    '※ローカルブラウザ依存の「DevToolsActivePort」探索の都合により「BrowserCapabilities」を使用します
    If WebMode Then
        'Web版の場合、普段使いとして、「User Data」とする
        '※`UserDataDir`ポリシーがある場合は、今はここでは自動対応できません
        BrowserCapabilities.Set_UserDataDirName = "User Data"
    Else
        'DesktopExcelの場合は、内部で使ってる「UserDataDir」の絶対パスを指定
        '※お使いの環境によっては違う場合があるためその場合は、調整が必要です
        BrowserCapabilities.Set_UserDataDirName = Environ("LOCALAPPDATA") & "\Microsoft\Office\16.0\Wef\webview2\2\1\EBWebView"
    End If

    '2. 指定のユーザ名のWebSocketForDevToolsActivePortへ接続
    Dim UserName As String: UserName = BrowserCapabilities.Set_UserDataDirName
    Dim WebSocketCDP As CDPCoreViaWebSocket: Set WebSocketCDP = New CDPCoreViaWebSocket
    WebSocketCDP.AutoConnectDevToolsActivePort BrowserCapabilities

    '3. WebSocketモードとして設定
    ReattachOptions.Set_WebSocketMode(WebSocketCDP) = UserName

    '4. 再接続情報を`reattach`メソッドに渡す
    Dim WebSocketChromium As CDPBrowser: Set WebSocketChromium = New CDPBrowser
    WebSocketChromium.reattach ReattachOptions

    '5. モードに応じて、`Office.js`がある場所を特定させる
    If WebMode Then
        'Web版は、`index.html#/edit`が、コードエディタです
        Set ConnectForWebExcel = WebSocketChromium.getTab(Url:="index.html#/edit", SearchTypeID:=kFrame)
    Else
        'Desktop版は、`$$$$16#/edit`が、コードエディタです
        Set ConnectForWebExcel = WebSocketChromium.getTab(Url:="$$$$16#/edit")
    End If
End Function



'***************************************************************************************************
'                                   ■■■ 各種Demo ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] ExcelらしいDemo
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：アクティブシートに、いくつか書き込みを行います。Officeスクリプト(コードエディターで実行)経由で
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Officeスクリプトをコードエディターに書き込み、「実行」ボタンを押して、完了を待ちます。
'            `main(workbook, text, when)`の text / when は、VBA から渡した値の const 宣言に置き換えてから実行します
'* 注意事項：・既にある場所の文字列は、上書きされるので、新しいシートで試すように
'            ・コードエディターで開いているスクリプトの中身は、このデモのスクリプトで上書きされます(自動保存されます)
'***************************************************************************************************
Sub Demo_Simple()
    Const InputText As String = "Hello World OfficeScript"     'ここに、任意のテキストを入力してみよう！
    Const TimeoutSec As Double = 120                            '実行の完了を待つ上限(秒)


    '1. 接続儀式を済ませる
    Dim CodeEditor As CDPContext
    Set CodeEditor = ConnectForWebExcel(UseWeb)
    If CodeEditor Is Nothing Then MsgBox "コードエディタが開いてる場所の確認が取れませんでした。事前に開く必要があります。", vbCritical: Exit Sub


    '2. 可変値部分を`Runtime.callFunctionOn`の`arguments`に渡せるように準備
    '   1つ目は実行するOfficeスクリプト本体。2つ目以降が、`main`の workbook より後ろの引数に順番に入る
    Dim args(2) As Dictionary
    Dim arg0 As New Dictionary: arg0.Add "value", OS_SimpleDemo(): Set args(0) = arg0   '実行するOfficeスクリプト
    Dim arg1 As New Dictionary: arg1.Add "value", InputText:       Set args(1) = arg1   'text：A1に書く文字列
    Dim arg2 As New Dictionary: arg2.Add "value", Now:             Set args(2) = arg2   'when：A2に書く実行日時

    '3. 実行！(エディターに書き込み、「実行」ボタンを押すところまで。スクリプトの完了は待たない)
    Dim result
    CDPHelpers.SetVar(result) = CodeEditor.jsEval(JS_StartOfficeScript(), objectArguments:=args, awaitPromise:=True)

    '4. 完了を待って、結果を表示
    If IsError(result) Then
        Debug.Print "[Demo_Simple] JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
    ElseIf result <> "STARTED" Then
        Debug.Print "[Demo_Simple] " & result
    Else
        PrintOfficeScriptResult WaitOfficeScriptResult(CodeEditor, TimeoutSec), "Demo_Simple"
    End If

    '5. 後始末
    CodeEditor.ThisCDPBrowser.DisconnectCDP
End Sub

'***************************************************************************************************
'* 機能　　：コードエディターを開いているブックの全シートを Officeスクリプトで調べ、結果をこのマクロのブックにレポートとグラフで書き出します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：調べる項目は、使用範囲・行数・列数・数式の数・エラーセルの数と場所(先頭5件)・テーブル数・グラフ数です。
'            Officeスクリプトは結果を JSON で console.log に出し、VBA は出力ログから受け取ります
'            ・Web 版        ：完了まで待ってから書き出します(同期)
'            ・デスクトップ版：実行を始めたらマクロをいったん終え、`Application.OnTime`で結果を覗きます(Demo_ViewBookInfo_Poll)。
'                              デスクトップ版の Officeスクリプトは Excel と同じプロセスで動くため、
'                              VBA が待っている間は、ブックの読み取りが進まないためです
'* 注意事項：・調べる側のブックは書き換えません。Office 側が「ループ内の読み取りは遅い」という注意を問題ログに出しますが、
'              実行には影響しません
'            ・結果は、このマクロのブック(ThisWorkbook)の「OfficeScriptReport」シートに書き出します(既にあれば中身を消して書き直します)
'            ・デスクトップ版で結果を待っている間は、もう一度実行できません
'***************************************************************************************************
Sub Demo_ViewBookInfo()
    Const ReportSheetName As String = "OfficeScriptReport"     '書き出し先のシート名(このマクロのブック)
    Const ScanLimit As Long = 200000                            '使用範囲のセル数がこれより多いシートは、数式とエラーを数えない
    Const TimeoutSec As Double = 300                            '実行の完了を待つ上限(秒)
    Const PollSec As Long = 2                                   'デスクトップ版で、結果を覗く間隔(秒)


    '0. デスクトップ版で前回の結果を待っている間は、実行しない
    If Not PendingEditor Is Nothing Then MsgBox "前回の棚卸しの結果を待っています。終わるまでお待ちください。", vbExclamation: Exit Sub

    '1. 接続儀式を済ませる
    Dim CodeEditor As CDPContext
    Set CodeEditor = ConnectForWebExcel(UseWeb)
    If CodeEditor Is Nothing Then MsgBox "コードエディタが開いてる場所の確認が取れませんでした。事前に開く必要があります。", vbCritical: Exit Sub

    '2. 可変値部分を`Runtime.callFunctionOn`の`arguments`に渡せるように準備
    Dim args(1) As Dictionary
    Dim arg0 As New Dictionary: arg0.Add "value", OS_ViewBookInfo(): Set args(0) = arg0   '実行するOfficeスクリプト
    Dim arg1 As New Dictionary: arg1.Add "value", ScanLimit:         Set args(1) = arg1   'limit：数式とエラーを数えるシートの上限(セル数)

    '3. 実行！(エディターに書き込み、「実行」ボタンを押すところまで。スクリプトの完了は待たない)
    Dim result
    CDPHelpers.SetVar(result) = CodeEditor.jsEval(JS_StartOfficeScript(), objectArguments:=args, awaitPromise:=True)
    If IsError(result) Then
        Debug.Print "[Demo_ViewBookInfo] JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
        CodeEditor.ThisCDPBrowser.DisconnectCDP: Exit Sub
    ElseIf result <> "STARTED" Then
        Debug.Print "[Demo_ViewBookInfo] " & result
        CodeEditor.ThisCDPBrowser.DisconnectCDP: Exit Sub
    End If

    '4. Web 版：完了まで待ってから書き出す(同期)
    If UseWeb Then
        WriteViewBookInfoResult WaitOfficeScriptResult(CodeEditor, TimeoutSec), ReportSheetName
        CodeEditor.ThisCDPBrowser.DisconnectCDP
        Exit Sub
    End If

    '5. デスクトップ版：接続と設定を保ったままマクロを終え、OnTime で結果を覗く(非同期)
    Set PendingEditor = CodeEditor
    PendingStartSec = Timer
    PendingTimeoutSec = TimeoutSec
    PendingPollSec = PollSec
    PendingSheetName = ReportSheetName
    Debug.Print "[Demo_ViewBookInfo] 実行を始めました(デスクトップ版)。マクロを終えて、結果を待ちます"
    Application.OnTime Now + TimeSerial(0, 0, PollSec), "Demo_OfficeScript.Demo_ViewBookInfo_Poll"
End Sub

'***************************************************************************************************
'* 機能　　：Demo_ViewBookInfo(デスクトップ版)の結果を覗きます。`Application.OnTime`から呼ばれます
'---------------------------------------------------------------------------------------------------
'* 注意事項：まだ終わっていなければ、PendingPollSec 秒後に自分自身を予約し直します
'***************************************************************************************************
Sub Demo_ViewBookInfo_Poll()
    If PendingEditor Is Nothing Then Exit Sub

    '1. 1回だけ覗く。まだ終わっていなければ、上限を過ぎていない限り予約し直す
    Dim result As String, Elapsed As Double
    result = CheckOfficeScriptResult(PendingEditor)
    If Len(result) = 0 Then
        Elapsed = Timer - PendingStartSec
        If Elapsed < 0 Then Elapsed = Elapsed + 86400#     '日付またぎ
        If Elapsed < PendingTimeoutSec Then
            Application.OnTime Now + TimeSerial(0, 0, PendingPollSec), "Demo_OfficeScript.Demo_ViewBookInfo_Poll"
            Exit Sub
        End If
        result = "NG: Timed out after " & PendingTimeoutSec & " seconds"
    End If

    '2. 待ち状態を先に解いてから(書き出しでエラーになっても、次が実行できるように)、書き出す
    Dim CodeEditor As CDPContext: Set CodeEditor = PendingEditor
    Set PendingEditor = Nothing
    WriteViewBookInfoResult result, PendingSheetName

    '3. 後始末
    CodeEditor.ThisCDPBrowser.DisconnectCDP
End Sub

Sub Demo_TakeVakue()
    Const TimeoutSec As Double = 120                            '実行の完了を待つ上限(秒)


    '1. 接続儀式を済ませる
    Dim CodeEditor As CDPContext
    Set CodeEditor = ConnectForWebExcel(UseWeb)
    If CodeEditor Is Nothing Then MsgBox "コードエディタが開いてる場所の確認が取れませんでした。事前に開く必要があります。", vbCritical: Exit Sub


    '2. 可変値部分を`Runtime.callFunctionOn`の`arguments`に渡せるように準備
    '   1つ目は実行するOfficeスクリプト本体。2つ目以降が、`main`の workbook より後ろの引数に順番に入る
    Dim args(0) As Dictionary
    Dim arg0 As New Dictionary: arg0.Add "value", Sheet7.Range("B1").value: Set args(0) = arg0  '実行するOfficeスクリプト

    '3. 実行！(エディターに書き込み、「実行」ボタンを押すところまで。スクリプトの完了は待たない)
    Dim result
    CDPHelpers.SetVar(result) = CodeEditor.jsEval(JS_StartOfficeScript(), objectArguments:=args, awaitPromise:=True)

    '4. 完了を待って、結果を表示
    If IsError(result) Then
        Debug.Print "[Demo_TakeVakue] JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
    ElseIf result <> "STARTED" Then
        Debug.Print "[Demo_TakeVakue>] " & result
    Else
        Dim Json As String: Json = WaitOfficeScriptResult(CodeEditor, TimeoutSec)
        PrintOfficeScriptResult Json, "Demo_TakeVakue"

        '5. 最後に console.log した値を受け取る(配列やオブジェクトは BiDiCDPJConv で届く)
        Dim Taken As Variant, i As Long
        CDPHelpers.SetVar(Taken) = OfficeScriptLastLog(Json)
        If IsObject(Taken) Then
            For i = 0 To Taken.Count - 1
                '※返り値のタイプによってはエラーになります。どういう返り値が来るかによって書き方を変えてください
                Debug.Print "[Demo_TakeVakue] Taken(" & i & ") = " & Taken.Item(i)
            Next
        ElseIf Not IsNull(Taken) Then
            Debug.Print "[Demo_TakeVakue] Taken = " & Taken
        End If
    End If

    '6. 後始末
    CodeEditor.ThisCDPBrowser.DisconnectCDP
End Sub



'***************************************************************************************************
'                         ■■■ 各種Demoヘルパー ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 共通。実行結果の受け取り
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：`JS_StartOfficeScript`で始めたOfficeスクリプトの結果を、1回だけ確認します(待ちません)
'---------------------------------------------------------------------------------------------------
'* 返り値　：まだ終わっていなければ空文字。終わっていれば JSON(`WaitOfficeScriptResult`と同じ形)。
'            問い合わせに失敗したら "NG: <理由>"
'***************************************************************************************************
Private Function CheckOfficeScriptResult(CodeEditor As CDPContext) As String
    Dim v
    CDPHelpers.SetVar(v) = CodeEditor.jsEval("(" & JS_OfficeScriptResult() & ")()")
    If IsError(v) Then
        CheckOfficeScriptResult = "NG: JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
    Else
        CheckOfficeScriptResult = v
    End If
End Function

'***************************************************************************************************
'* 機能　　：`JS_StartOfficeScript`で始めたOfficeスクリプトの完了を待ち、結果を JSON で返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：{status, ms, history, outputs:[…], logs:[…], errors:[…], messages:[…]}  または "NG: <理由>"
'            status は Succeeded / Failed / Unsupported のいずれか
'* 引数　　：CodeEditor     接続済みのコードエディター
'            TimeoutSec     待つ上限(秒)。時間切れでも、スクリプトの実行自体は続きます
'---------------------------------------------------------------------------------------------------
'* 注意事項：完了まで1回の`awaitPromise`で待たず、短い問い合わせを繰り返します。
'            デスクトップ版で VBA が長く止まったままだと、スクリプトの実行が進まないことがあるためです
'***************************************************************************************************
Private Function WaitOfficeScriptResult(CodeEditor As CDPContext, TimeoutSec As Double) As String
    Const PollSec As Double = 0.5


    Dim StartSec As Double: StartSec = Timer
    Dim Elapsed As Double, result As String
    Do
        CDPHelpers.Sleep PollSec
        result = CheckOfficeScriptResult(CodeEditor)
        If Len(result) > 0 Then
            WaitOfficeScriptResult = result
            Exit Function
        End If

        Elapsed = Timer - StartSec
        If Elapsed < 0 Then Elapsed = Elapsed + 86400#     '日付またぎ
    Loop While Elapsed < TimeoutSec

    WaitOfficeScriptResult = "NG: Timed out after " & TimeoutSec & " seconds"
End Function

'***************************************************************************************************
'* 機能　　：`WaitOfficeScriptResult`の結果を、イミディエイトウィンドウに表示します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Json       `WaitOfficeScriptResult`の返り値
'            Caption    表示の先頭に付ける名前
'***************************************************************************************************
Private Sub PrintOfficeScriptResult(Json As String, Caption As String)
    If Left$(Json, 1) <> "{" Then
        Debug.Print "[" & Caption & "] " & Json
        Exit Sub
    End If

    Dim Root As BiDiCDPJConv: Set Root = BiDiCDPJConv.Parse(Json)
    Debug.Print "[" & Caption & "] status = " & Root.StringKey("status") & " (" & Root.NumberKey("ms") & "ms / " & Root.StringKey("history") & ")"

    Dim Key As Variant, List As BiDiCDPJConv, i As Long
    For Each Key In Array("outputs", "errors", "messages")
        Set List = Root.NodeKey(CStr(Key))
        If Not List Is Nothing Then
            For i = 0 To List.Count - 1
                Debug.Print "[" & Caption & "] " & Key & ": " & List.StringAt(i)
            Next
        End If
    Next
End Sub

'***************************************************************************************************
'* 機能　　：`WaitOfficeScriptResult` / `CheckOfficeScriptResult`の結果から、最後に console.log した値を取り出します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：「VBA に渡したい最終結果は、最後に console.log する」という約束で、Officeスクリプトから値を受け取る用です。
'            Power Automate 向けの`return`と並べて書けば、同じスクリプトのまま両方に渡せます
'              console.log(values);     // VBA へ(最後の console.log)
'              return values;           // Power Automate へ
'* 返り値　：文字列・数値・真偽値は、そのままの値。配列やオブジェクトは BiDiCDPJConv(`Count`/`Item(i)`/`Stringify`等で読む)。
'            成功していない・console.log が1件も無い・結果が "NG: …" の場合は Null
'* 使い方　：Dim v: CDPHelpers.SetVar(v) = OfficeScriptLastLog(WaitOfficeScriptResult(CodeEditor, 120))
'***************************************************************************************************
Private Function OfficeScriptLastLog(Json As String) As Variant
    OfficeScriptLastLog = Null
    If Left$(Json, 1) <> "{" Then Exit Function

    Dim Root As BiDiCDPJConv: Set Root = BiDiCDPJConv.Parse(Json)
    If Root.StringKey("status") <> "Succeeded" Then Exit Function

    Dim Logs As BiDiCDPJConv: Set Logs = Root.NodeKey("logs")
    If Logs Is Nothing Then Exit Function
    If Logs.Count = 0 Then Exit Function

    Dim v As Variant
    CDPHelpers.SetVar(v) = Logs.Item(Logs.Count - 1)
    If IsObject(v) Then
        Set OfficeScriptLastLog = v
    Else
        OfficeScriptLastLog = v
    End If
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] Demo_ViewBookInfo 用ヘルパー。結果の書き出し
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：Demo_ViewBookInfo の結果を受け取り、出力ログの JSON(console.log)を、レポートとして書き出します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Json           `WaitOfficeScriptResult` / `CheckOfficeScriptResult`の返り値
'            SheetName      書き出し先のシート名
'***************************************************************************************************
Private Sub WriteViewBookInfoResult(Json As String, SheetName As String)
    Const Caption As String = "Demo_ViewBookInfo"


    '1. 成功していなければ、理由を表示して終わる
    If Left$(Json, 1) <> "{" Then Debug.Print "[" & Caption & "] " & Json: Exit Sub
    Dim Root As BiDiCDPJConv: Set Root = BiDiCDPJConv.Parse(Json)
    If Root.StringKey("status") <> "Succeeded" Then PrintOfficeScriptResult Json, Caption: Exit Sub

    '2. 出力ログのうち、最後の JSON(「{」で始まるもの)を探す
    Dim Outputs As BiDiCDPJConv, i As Long, Report As String
    Set Outputs = Root.NodeKey("outputs")
    If Not Outputs Is Nothing Then
        For i = Outputs.Count - 1 To 0 Step -1
            If Left$(Outputs.StringAt(i), 1) = "{" Then Report = Outputs.StringAt(i): Exit For
        Next
    End If
    If Len(Report) = 0 Then Debug.Print "[" & Caption & "] 棚卸しの結果(JSON)を受け取れませんでした": Exit Sub

    '3. 書き出す
    Debug.Print "[" & Caption & "] status = Succeeded (" & Root.NumberKey("ms") & "ms)"
    WriteInspectReport Report, SheetName
End Sub

'***************************************************************************************************
'* 機能　　：このマクロのブックに、書き出し用のシートを用意します(既にあれば、テーブル・グラフ・中身を消す)
'***************************************************************************************************
Private Function PrepareLocalSheet(SheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SheetName
    Else
        Do While ws.ListObjects.Count > 0
            ws.ListObjects(1).Delete
        Loop
        If ws.ChartObjects.Count > 0 Then ws.ChartObjects.Delete
        ws.Cells.Clear
    End If
    Set PrepareLocalSheet = ws
End Function

'***************************************************************************************************
'* 機能　　：棚卸しの JSON を、このマクロのブックの指定シートにテーブルとグラフで書き出します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Json           Officeスクリプトが console.log に出した棚卸しの JSON
'            SheetName      書き出し先のシート名(無ければ作り、あれば中身を消して書き直す)
'***************************************************************************************************
Private Sub WriteInspectReport(Json As String, SheetName As String)
    Const FirstRow As Long = 4


    '1. JSON を読む
    Dim Root As BiDiCDPJConv, SheetList As BiDiCDPJConv
    Set Root = BiDiCDPJConv.Parse(Json)
    Set SheetList = Root.NodeKey("sheets")
    If SheetList Is Nothing Then
        Debug.Print "[Demo_ViewBookInfo] 棚卸しの JSON に sheets がありません: " & Left$(Json, 200)
        Exit Sub
    End If

    '2. 見出し
    Dim ws As Worksheet
    Set ws = PrepareLocalSheet(SheetName)
    ws.Range("A1").value = "ブックの棚卸し(Officeスクリプト)：" & Root.StringKey("workbook")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").value = "取得日時：" & Format$(Now, "yyyy/mm/dd hh:nn:ss") & "　名前の定義：" & Root.NumberKey("names") & " 件"
    ws.Range("A" & FirstRow).Resize(1, 10).value = Array("シート名", "使用範囲", "行数", "列数", "セル数", "数式", "エラー", "エラーの場所(先頭5件)", "テーブル", "グラフ")

    '3. シートごとの行
    Dim r As Long, Tok As Long, Node As BiDiCDPJConv, ErrCells As BiDiCDPJConv, i As Long, List As String
    r = FirstRow
    Tok = SheetList.FirstChildToken
    Do While Tok > 0
        r = r + 1
        Set Node = SheetList.NodeFromToken(Tok)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).NumberFormat = "@"
        ws.Cells(r, 8).NumberFormat = "@"
        ws.Cells(r, 1).value = Node.StringKey("name")
        ws.Cells(r, 2).value = Node.StringKey("used")
        ws.Cells(r, 3).value = Node.NumberKey("rows")
        ws.Cells(r, 4).value = Node.NumberKey("cols")
        ws.Cells(r, 5).value = Node.NumberKey("rows") * Node.NumberKey("cols")
        If Node.BoolKey("scanned") Then
            ws.Cells(r, 6).value = Node.NumberKey("formulas")
            ws.Cells(r, 7).value = Node.NumberKey("errors")
        Else
            ws.Cells(r, 6).value = "-"
            ws.Cells(r, 7).value = "-"
        End If

        List = ""
        Set ErrCells = Node.NodeKey("errorCells")
        If Not ErrCells Is Nothing Then
            For i = 0 To ErrCells.Count - 1
                If Len(List) Then List = List & ", "
                List = List & ErrCells.StringAt(i)
            Next
        End If
        ws.Cells(r, 8).value = List
        ws.Cells(r, 9).value = Node.NumberKey("tables")
        ws.Cells(r, 10).value = Node.NumberKey("charts")

        Debug.Print "[Demo_ViewBookInfo] " & Node.StringKey("name") & " : " & Node.StringKey("used") & " / 数式 " & ws.Cells(r, 6).value & " / エラー " & ws.Cells(r, 7).value & IIf(Len(List), " (" & List & ")", "")
        Tok = SheetList.NextToken(Tok)
    Loop
    If r = FirstRow Then Exit Sub

    '4. テーブルと条件付き書式(セル数はデータバー、エラーがあれば赤)
    Dim Table As ListObject
    Set Table = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(FirstRow, 1), ws.Cells(r, 10)), , xlYes)
    Table.TableStyle = "TableStyleMedium2"
    Table.ListColumns(5).DataBodyRange.FormatConditions.AddDatabar
    With Table.ListColumns(7).DataBodyRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With
    ws.Columns("A:J").AutoFit

    '5. シートごとのセル数のグラフ
    Dim ChartObj As ChartObject
    Set ChartObj = ws.ChartObjects.Add(Left:=ws.Cells(FirstRow, 12).Left, Top:=ws.Cells(FirstRow, 12).Top, Width:=420, height:=260)
    With ChartObj.Chart
        .ChartType = xlBarClustered
        With .SeriesCollection.NewSeries
            .Name = "セル数"
            .Values = Table.ListColumns(5).DataBodyRange
            .XValues = Table.ListColumns(1).DataBodyRange
        End With
        .HasTitle = True
        .ChartTitle.Text = "シートごとの使用セル数"
        .HasLegend = False
    End With

    ws.activate
    ws.Range("A1").Select
End Sub



'***************************************************************************************************
'                         ■■■ コードエディター周りのJavaScript ■■■
'***************************************************************************************************
'* 機能　　：Officeスクリプトをコードエディターに書き込み、「実行」ボタンを押します(完了は待ちません)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・エディター(Monaco)と画面の状態(Redux の store)は、React の内部から探します
'            ・`main(workbook, 引数…)`の「引数…」は、実行時の入力ダイアログを避けるため、
'              渡された値の const 宣言に置き換えてから書き込みます
'            ・実行の様子は store を見張って window.__exCDP_OfficeScriptRun に記録し、`JS_OfficeScriptResult`で取りに行きます
'* 引数　　：script         実行するOfficeスクリプト(arguments の1つ目)
'            ...params      `main`の workbook より後ろの引数に入れる値(arguments の2つ目以降)
'* 返り値　："STARTED"  または "NG: <理由>"
'* 注意事項：・エディターで開いているスクリプトの中身は上書きされます(Ctrl+Z で戻せます。自動保存もされます)
'            ・`main`の引数の型に「,」や「)」を含むもの(Map<string, number> 等)は、置き換えに対応していません
'            ・React / Redux の中身の構造は、Office 側の更新で変わる可能性があります
'***************************************************************************************************
Private Function JS_StartOfficeScript() As String
    InitJSList(vbLf) = 79    '行数 - 1(JSList の上限)
    SetJS = "async function (script, ...params) {"
    SetJS = "    // Office スクリプトをコードエディターに書き込み、「実行」ボタンを押す(結果は待たない。JS_OfficeScriptResult で取りに行く)"
    SetJS = "    // main(workbook, 引数…) の「引数…」は、そのままだと実行時に入力ダイアログが出るため、VBA から渡された値の const 宣言に置き換える"
    SetJS = "    const norm = (t) => String(t).replace(/\r\n?/g, ""\n"");"
    SetJS = "    const fmt = (v) => (typeof v === ""string"" ? v : JSON.stringify(v));"
    SetJS = "    const fiberOf = (el) => { const k = Object.keys(el).find((k) => k.startsWith(""__reactFiber$"")); return k ? el[k] : null; };"
    SetJS = "    const isEd = (v) => !!v && typeof v === ""object"" && typeof v.getModel === ""function"" && typeof v.executeEdits === ""function"";"
    SetJS = "    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));"
    SetJS = ""
    SetJS = "    // 1. main の引数を const 宣言に置き換えた、実行用のスクリプトを作る"
    SetJS = "    const m = /function\s+main\s*\(([^)]*)\)/.exec(script);"
    SetJS = "    if (!m) { return ""NG: function main(...) not found""; }"
    SetJS = "    const decls = m[1].split("","").map((s) => s.trim()).filter((s) => s.length > 0);"
    SetJS = "    const first = decls.shift();"
    SetJS = "    if (decls.length !== params.length) { return ""NG: main has "" + decls.length + "" parameter(s), but "" + params.length + "" value(s) were passed""; }"
    SetJS = "    const consts = decls.map((d, i) => ""const "" + d + "" = "" + JSON.stringify(params[i]) + "";"");"
    SetJS = "    const code = norm((consts.length > 0 ? ""// ===== VBA から渡された引数 =====\n"" + consts.join(""\n"") + ""\n\n"" : """") + script.replace(m[0], ""function main("" + first + "")""));"
    SetJS = ""
    SetJS = "    // 2. エディター(Monaco)と、画面の状態(Redux の store)を、React の内部から探す"
    SetJS = "    let ed = null;"
    SetJS = "    for (let el = document.querySelector("".monaco-editor""); el && !ed; el = el.parentElement) {"
    SetJS = "        for (let f = fiberOf(el), d = 0; f && d < 40 && !ed; f = f.return, d++) {"
    SetJS = "            const sn = f.stateNode;"
    SetJS = "            if (sn && !(sn instanceof Node)) { if (isEd(sn.editor)) { ed = sn.editor; } else if (isEd(sn)) { ed = sn; } }"
    SetJS = "        }"
    SetJS = "    }"
    SetJS = "    if (!ed) { return ""NG: monaco editor not found (open a script in the code editor)""; }"
    SetJS = "    let store = null;"
    SetJS = "    for (const n of document.body.querySelectorAll(""*"")) {"
    SetJS = "        for (let f = fiberOf(n); f && !store; f = f.return) { const p = f.memoizedProps; if (p && p.store && typeof p.store.getState === ""function"") { store = p.store; } }"
    SetJS = "        if (store) { break; }"
    SetJS = "    }"
    SetJS = "    if (!store) { return ""NG: redux store not found""; }"
    SetJS = "    if (store.getState().run.runningStatus === ""InProgress"") { return ""NG: another script is running""; }"
    SetJS = ""
    SetJS = "    // 3. エディターの全文を置き換える(Ctrl+Z で戻せる編集として)"
    SetJS = "    const model = ed.getModel();"
    SetJS = "    ed.pushUndoStop();"
    SetJS = "    ed.executeEdits(""exCDP_OfficeScript"", [{ range: model.getFullModelRange(), text: code, forceMoveMarkers: true }]);"
    SetJS = "    ed.pushUndoStop();"
    SetJS = ""
    SetJS = "    // 4. 実行ボタンは store 側の本文を実行するので、書き込みが store に反映されるまで待つ(最大10秒)"
    SetJS = "    const synced = () => {"
    SetJS = "        const s = store.getState(), id = s.editor.active, c = [];"
    SetJS = "        if (s.scripts.scripts && s.scripts.scripts[id]) { c.push(s.scripts.scripts[id]); }"
    SetJS = "        for (const x of (s.scripts.mruScripts || [])) { if (x && x.storageScript && x.storageScript.id === id) { c.push(x); } }"
    SetJS = "        return c.some((x) => norm((x.localUpdates && typeof x.localUpdates.body === ""string"") ? x.localUpdates.body : x.storageScript.body) === code);"
    SetJS = "    };"
    SetJS = "    for (let t = 0; !synced(); t += 100) {"
    SetJS = "        if (t >= 10000) { return ""NG: the editor content was not reflected in the store""; }"
    SetJS = "        await sleep(100);"
    SetJS = "    }"
    SetJS = ""
    SetJS = "    // 5. 実行ボタン(aria-label は日本語 UI「実行」/ 英語 UI「Run」)"
    SetJS = "    const btn = [...document.querySelectorAll(""button"")].find((b) => [""実行"", ""Run""].includes(b.getAttribute(""aria-label"")));"
    SetJS = "    if (!btn) { return ""NG: run button not found""; }"
    SetJS = "    if (btn.disabled || btn.getAttribute(""aria-disabled"") === ""true"") { return ""NG: run button disabled""; }"
    SetJS = ""
    SetJS = "    // 6. store を見張って、状態の移り変わりとメッセージバーを記録してから、実行ボタンを押す"
    SetJS = "    const old = window.__exCDP_OfficeScriptRun;"
    SetJS = "    if (old && old.unsub) { old.unsub(); }"
    SetJS = "    const s0 = store.getState();"
    SetJS = "    const r = { store: store, status: """", startedAt: Date.now(), doneAt: 0, lastMsgAt: 0, history: [], messages: [], logStart: s0.bottomPanel.outputLogs.length, unsub: null };"
    SetJS = "    let prev = s0.run.runningStatus, prevMsg = (s0.messageBar && s0.messageBar.isVisible) ? fmt(s0.messageBar.text || """") : """";"
    SetJS = "    r.unsub = store.subscribe(() => {"
    SetJS = "        const s = store.getState();"
    SetJS = "        if (s.bottomPanel.outputLogs.length < r.logStart) { r.logStart = 0; }"
    SetJS = "        const mb = (s.messageBar && s.messageBar.isVisible) ? fmt(s.messageBar.text || """") : """";"
    SetJS = "        if (mb && mb !== prevMsg) { r.messages.push(mb); r.lastMsgAt = Date.now(); }"
    SetJS = "        prevMsg = mb;"
    SetJS = "        const cur = s.run.runningStatus;"
    SetJS = "        if (cur === prev) { return; }"
    SetJS = "        prev = cur;"
    SetJS = "        r.history.push(cur);"
    SetJS = "        if (!r.status && cur !== ""None"" && cur !== ""InProgress"") { r.status = cur; r.doneAt = Date.now(); }"
    SetJS = "    });"
    SetJS = "    window.__exCDP_OfficeScriptRun = r;"
    SetJS = "    btn.click();"
    SetJS = "    return ""STARTED"";"
    SetJS = "}"
    JS_StartOfficeScript = SetJS
End Function

'***************************************************************************************************
'* 機能　　：`JS_StartOfficeScript`で始めた実行の結果を返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：まだ終わっていなければ空文字。終わっていれば JSON
'            {status, ms, history, outputs:[…], logs:[…], errors:[…], messages:[…]}
'              status   : Succeeded / Failed / Unsupported
'              outputs  : 出力ログ(console.log やエラー)を、表示用に文字列にしたもの
'              logs     : 出力ログのうちエラー以外を、console.log に渡した値のまま(数値・配列・オブジェクトの型を保つ)。
'                         console.log("x", 2) のように複数渡すと、"x" と 2 の2件になる
'              errors   : 出力ログのうちエラーのもの(Web版は例外のメッセージがそのまま、デスクトップ版は汎用のメッセージ)
'              messages : メッセージバー(「スクリプトを実行中...」→「スクリプトが正常に実行されました。」等)
'***************************************************************************************************
Private Function JS_OfficeScriptResult() As String
    InitJSList(vbLf) = 24    '行数 - 1(JSList の上限)
    SetJS = "function () {"
    SetJS = "    // JS_StartOfficeScript で始めた実行の結果を JSON で返す。まだ終わっていなければ空文字"
    SetJS = "    // 終わった後も、結果のメッセージバー(「実行中」→「結果」の2件)が出揃うまで、最大 SETTLE_MS 待ってから確定する"
    SetJS = "    const SETTLE_MS = 3000;"
    SetJS = "    const fmt = (v) => (typeof v === ""string"" ? v : JSON.stringify(v));"
    SetJS = "    const r = window.__exCDP_OfficeScriptRun;"
    SetJS = "    if (!r) { return ""NG: no run (call JS_StartOfficeScript first)""; }"
    SetJS = "    if (!r.status) { return """"; }"
    SetJS = ""
    SetJS = "    const now = Date.now();"
    SetJS = "    const settled = (r.messages.length >= 2 && r.lastMsgAt >= r.doneAt && now - r.lastMsgAt >= 200) || now - r.doneAt >= SETTLE_MS;"
    SetJS = "    if (!settled) { return """"; }"
    SetJS = "    if (r.unsub) { r.unsub(); r.unsub = null; }"
    SetJS = ""
    SetJS = "    const logs = r.store.getState().bottomPanel.outputLogs.slice(r.logStart).filter((l) => l && l.message !== undefined);"
    SetJS = "    return JSON.stringify({"
    SetJS = "        status: r.status,"
    SetJS = "        ms: r.doneAt - r.startedAt,"
    SetJS = "        history: r.history.join("" -> ""),"
    SetJS = "        outputs: logs.map((l) => fmt(l.message)),"
    SetJS = "        logs: logs.filter((l) => l.severity !== ""error"").map((l) => l.message),"
    SetJS = "        errors: logs.filter((l) => l.severity === ""error"").map((l) => fmt(l.message)),"
    SetJS = "        messages: r.messages"
    SetJS = "    });"
    SetJS = "}"
    JS_OfficeScriptResult = SetJS
End Function



'***************************************************************************************************
'                   ■■■ OfficeScriptに渡すテンプレートコード一式 ■■■
'***************************************************************************************************
'* 機能　　：Demo_Simple 用。アクティブシートの A1 に text を太字で、A2 に when を日時として書きます
'---------------------------------------------------------------------------------------------------
'* 引数　　：text / when    A1 に書く文字列 / A2 に書く日時(VBA の Date は "yyyy-mm-dd hh:nn:ss" の文字列で届く)
'* 注意事項：Power Automate の「スクリプトの実行」に置く場合は、このままの形で、text / when をフローから渡せます
'***************************************************************************************************
Private Function OS_SimpleDemo() As String
    InitJSList(vbLf) = 21    '行数 - 1(JSList の上限)
    SetJS = "function main(workbook: ExcelScript.Workbook, text: string, when: string) {"
    SetJS = "    // アクティブなシートの A1 に text を太字で、A2 に when(VBA の Now)を日時として書く"
    SetJS = "    // ※VBA からは text / when を const 宣言に置き換えて実行する。Power Automate の「スクリプトの実行」なら、このまま引数として渡せる"
    SetJS = "    const sheet = workbook.getActiveWorksheet();"
    SetJS = "    const a1 = sheet.getRange(""A1"");"
    SetJS = "    a1.setValue(text);"
    SetJS = "    a1.getFormat().getFont().setBold(true);"
    SetJS = ""
    SetJS = "    const a2 = sheet.getRange(""A2"");"
    SetJS = "    a2.setValue(toSerial(when));"
    SetJS = "    a2.setNumberFormat(""yyyy/mm/dd hh:mm:ss"");"
    SetJS = "    sheet.getRange(""A1:A2"").getFormat().autofitColumns();"
    SetJS = "    console.log(""OK: "" + sheet.getName() + ""!A1:A2"");"
    SetJS = "}"
    SetJS = ""
    SetJS = "// ""yyyy-mm-dd hh:nn:ss""(VBA の Date を JSON にした形。区切りが「T」でも可)を、Excel のシリアル値にする"
    SetJS = "function toSerial(iso: string): number {"
    SetJS = "    const m = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})/.exec(iso);"
    SetJS = "    if (!m) { return 0; }"
    SetJS = "    const utc = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6]));"
    SetJS = "    return utc / 86400000 + 25569;"
    SetJS = "}"
    OS_SimpleDemo = SetJS
End Function

'***************************************************************************************************
'* 機能　　：Demo_ViewBookInfo 用。全シートを調べて、結果の JSON を console.log に出します(ブックは書き換えません)
'---------------------------------------------------------------------------------------------------
'* 引数　　：limit    使用範囲のセル数がこれより多いシートは、数式とエラーを数えない
'* 出力　　：{workbook, names, sheets:[{name, used, rows, cols, formulas, errors, errorCells, tables, charts, scanned}]}
'***************************************************************************************************
Private Function OS_ViewBookInfo() As String
    InitJSList(vbLf) = 33    '行数 - 1(JSList の上限)
    SetJS = "function main(workbook: ExcelScript.Workbook, limit: number) {"
    SetJS = "    // 全シートの使用範囲・数式・エラー・テーブル・グラフを調べて、JSON を console.log に出す(ブックは書き換えない)"
    SetJS = "    // 使用範囲のセル数が limit を超えるシートは、数式とエラーを数えない(時間がかかりすぎるため)"
    SetJS = "    const MAX_ERR = 5;"
    SetJS = "    const col = (n: number): string => { let s = """"; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };"
    SetJS = "    const sheets: { name: string; used: string; rows: number; cols: number; formulas: number; errors: number; errorCells: string[]; tables: number; charts: number; scanned: boolean }[] = [];"
    SetJS = "    for (const ws of workbook.getWorksheets()) {"
    SetJS = "        const info = { name: ws.getName(), used: """", rows: 0, cols: 0, formulas: 0, errors: 0, errorCells: [] as string[], tables: ws.getTables().length, charts: ws.getCharts().length, scanned: false };"
    SetJS = "        const used = ws.getUsedRange(true);"
    SetJS = "        if (used) {"
    SetJS = "            info.used = used.getAddress().split(""!"").pop() || """";"
    SetJS = "            info.rows = used.getRowCount();"
    SetJS = "            info.cols = used.getColumnCount();"
    SetJS = "            if (info.rows * info.cols <= limit) {"
    SetJS = "                info.scanned = true;"
    SetJS = "                const top = used.getRowIndex(), left = used.getColumnIndex();"
    SetJS = "                const types = used.getValueTypes();"
    SetJS = "                const formulas = used.getFormulas();"
    SetJS = "                for (let r = 0; r < info.rows; r++) {"
    SetJS = "                    for (let c = 0; c < info.cols; c++) {"
    SetJS = "                        const f = formulas[r][c];"
    SetJS = "                        if (typeof f === ""string"" && f.startsWith(""="")) { info.formulas++; }"
    SetJS = "                        if (types[r][c] === ExcelScript.RangeValueType.error) {"
    SetJS = "                            info.errors++;"
    SetJS = "                            if (info.errorCells.length < MAX_ERR) { info.errorCells.push(col(left + c) + (top + r + 1)); }"
    SetJS = "                        }"
    SetJS = "                    }"
    SetJS = "                }"
    SetJS = "            }"
    SetJS = "        }"
    SetJS = "        sheets.push(info);"
    SetJS = "    }"
    SetJS = "    console.log(JSON.stringify({ workbook: workbook.getName(), names: workbook.getNames().length, sheets: sheets }));"
    SetJS = "}"
    OS_ViewBookInfo = SetJS
End Function



'***************************************************************************************************
'                 ■■■ なが～いJavaScriptコードの設定をスマートにする用 ■■■
'***************************************************************************************************
Private Property Let InitJSList(Optional JsJoin As String, Num As Long)
    ReDim JSList(Num)
    CurrentIndex = 0
    SetSplit = JsJoin
End Property
Private Property Let SetJS(JSStr As String)
    JSList(CurrentIndex) = JSStr
    CurrentIndex = CurrentIndex + 1
End Property
Private Property Get SetJS() As String
    SetJS = Join(JSList, SetSplit)
End Property
