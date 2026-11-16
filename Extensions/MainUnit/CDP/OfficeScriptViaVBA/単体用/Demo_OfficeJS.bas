Attribute VB_Name = "Demo_OfficeJS"
'==============================================================================================================
'           OfficeJSをCDP経由で動かすためのDemoです。いわゆる、即席Officeアドインみたいなポジションです
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
Private Const ShowCodeInEditor  As Boolean = True   '演出：実行する関数と引数を、コードエディターにも表示する(実行は Office.js)



'***************************************************************************************************
'                 ■■■ 長いJavaScriptコードを見やすくするためのヘルパー変数 ■■■
'***************************************************************************************************
Private JSList()        As String
Private CurrentIndex    As Long
Private SetSplit        As String



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
        '※`taskpane.html`があれば実はどこでも機能(Excel Labs内のOffice.jsを拝借など)しますが、コードが見える見栄え的に、ここにする
        Set ConnectForWebExcel = WebSocketChromium.getTab(Url:="index.html#/edit", SearchTypeID:=kFrame)
    Else
        'Desktop版は、`$$$$16#/edit`が、コードエディタです
        '※実はどこでも機能(ヘルプ内のOffice.jsを拝借など)しますが、コードが見える見栄え的に、ここにする
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
'* 機能　　：アクティブシートに、いくつか書き込みを行います。Office.js経由で
'---------------------------------------------------------------------------------------------------
'* 注意事項：既にある場所の文字列は、上書きされるので、新しいシートで試すように
'***************************************************************************************************
Sub Demo_Simple()
    Const InputText As String = "Hello World Office.js"     'ここに、任意のテキストを入力してみよう！


    '1. 接続儀式を済ませる
    Dim CodeEditor As CDPContext
    Set CodeEditor = ConnectForWebExcel(UseWeb)
    If CodeEditor Is Nothing Then MsgBox "コードエディタが開いてる場所の確認が取れませんでした。事前に開く必要があります。", vbCritical: Exit Sub


    '2. 可変値部分を`Runtime.callFunctionOn`の`arguments`に渡せるように準備
    Dim args(1) As Dictionary
    Dim arg1 As New Dictionary: arg1.Add "value", InputText: Set args(0) = arg1  'A1に書く文字列
    Dim arg2 As New Dictionary: arg2.Add "value", Now:       Set args(1) = arg2  'A2に書く実行日時

    '3. 演出：実行する関数と引数を、呼び出しの形でコードエディターに表示(結果は待たない)
    If ShowCodeInEditor Then ShowInEditor CodeEditor, JS_SimpleDemo(), args, "SimpleDemo"

    '4. 実行！(Excel.run は Promise を返すので、awaitPromise で完了まで待つ)
    Dim result
    CDPHelpers.SetVar(result) = CodeEditor.jsEval(JS_SimpleDemo(), objectArguments:=args, awaitPromise:=True)
    Debug.Print "[SimpleDemo] " & result

    '5. 後始末
    CodeEditor.ThisCDPBrowser.DisconnectCDP
End Sub

'***************************************************************************************************
'* 機能　　：コードエディターを開いているブックの全シートを調べ、結果をこのマクロのブックにレポートとグラフで書き出します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：調べる項目は、使用範囲・行数・列数・数式の数・エラーセルの数と場所(先頭5件)・テーブル数・グラフ数です。
'            Office.js の`sync`を3回にまとめて読むため、Officeスクリプトとして実行するより速く終わります
'* 注意事項：・調べる側のブックは書き換えません
'            ・結果は、このマクロのブック(ThisWorkbook)の「OfficeJsReport」シートに書き出します(既にあれば中身を消して書き直します)
'***************************************************************************************************
Sub Demo_ViewBookInfo()
    Const ReportSheetName As String = "OfficeJsReport"     '書き出し先のシート名(このマクロのブック)


    '1. 接続儀式を済ませる
    Dim CodeEditor As CDPContext
    Set CodeEditor = ConnectForWebExcel(UseWeb)
    If CodeEditor Is Nothing Then MsgBox "コードエディタが開いてる場所の確認が取れませんでした。事前に開く必要があります。", vbCritical: Exit Sub

    '2. 演出：実行する関数を、呼び出しの形でコードエディターに表示(結果は待たない)
    If ShowCodeInEditor Then ShowInEditor CodeEditor, JS_ViewBookInfo(), , "ViewBookInfoDemo"

    '3. 実行！(引数がないので`Runtime.evaluate`になる。関数宣言を「( … )()」で呼ぶ式にして渡す)
    Dim result
    CDPHelpers.SetVar(result) = CodeEditor.jsEval("(" & JS_ViewBookInfo() & ")()", awaitPromise:=True)

    '4. 結果(JSON)を、手元のブックにレポートとして書き出す
    If IsError(result) Then
        Debug.Print "[ViewBookInfoDemo] JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
    ElseIf Left$(result, 1) = "{" Then
        WriteInspectReport CStr(result), ReportSheetName
    Else
        Debug.Print "[ViewBookInfoDemo] " & result
    End If

    '5. 後始末
    CodeEditor.ThisCDPBrowser.DisconnectCDP
End Sub

'***************************************************************************************************
'* 機能　　：Web 版 Excel でアクティブなシートを繰り返し読み取り、このマクロのブックのシートに写します(ライブミラー)
'            前回から変わったセルは黄色に光ります
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Web 版 Excel と手元の Excel を並べて、Web 版でセルに入力すると、手元に写る様子を見るデモです。
'            読み取るのは、使用範囲の左上から最大 MaxRows 行 x MaxCols 列の「表示されている文字」です
'* 注意事項：・常に Web 版で動かします(デスクトップ版だと、このマクロのブック自身を読むことになりやすいため)
'            ・読み取る側のブックは書き換えません
'            ・結果は、このマクロのブック(ThisWorkbook)の「OfficeJsMirror」シートに写します(既にあれば中身を消して書き直します)
'            ・ループ中に止めたいときは、Ctrl+Break で VBA を中断してください
'***************************************************************************************************
Sub Demo_LiveMirror()
    Const MirrorSheetName As String = "OfficeJsMirror"     '写す先のシート名(このマクロのブック)
    Const Polls As Long = 120                               '読み取る回数
    Const WaitSec As Double = 0.5                           '読み取りの間隔(秒)
    Const MaxRows As Long = 40                              '読み取る最大の行数
    Const MaxCols As Long = 12                              '読み取る最大の列数
    Const TopRow As Long = 4                                '写す先の開始行


    '1. 接続儀式を済ませる(このデモは常に Web 版)
    Dim CodeEditor As CDPContext
    Set CodeEditor = ConnectForWebExcel(True)
    If CodeEditor Is Nothing Then MsgBox "Web 版 Excel のコードエディタが開いてる場所の確認が取れませんでした。事前に開く必要があります。", vbCritical: Exit Sub

    '2. 可変値部分を`Runtime.callFunctionOn`の`arguments`に渡せるように準備
    Dim args(1) As Dictionary
    Dim arg1 As New Dictionary: arg1.Add "value", MaxRows: Set args(0) = arg1   '読み取る最大の行数
    Dim arg2 As New Dictionary: arg2.Add "value", MaxCols: Set args(1) = arg2   '読み取る最大の列数

    '3. 写す先のシートを用意
    Dim ws As Worksheet
    Set ws = PrepareLocalSheet(MirrorSheetName)
    ws.Range("A1").value = "Web 版 Excel → 手元の Excel (ライブミラー / Office.js)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.activate

    '4. 演出：実行する関数と引数を、呼び出しの形でコードエディターに表示(ループ前に1回だけ。結果は待たない)
    Dim js As String: js = JS_LiveMirror()
    If ShowCodeInEditor Then ShowInEditor CodeEditor, js, args, "LiveMirrorDemo (" & Polls & " polls)"

    '5. 読み取り → 書き出し を繰り返す
    Dim result, p As Long, Prev As Variant, LastAddress As String
    For p = 1 To Polls
        CDPHelpers.SetVar(result) = CodeEditor.jsEval(js, objectArguments:=args, awaitPromise:=True)
        If IsError(result) Then
            Debug.Print "[LiveMirrorDemo] JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
            Exit For
        ElseIf Left$(result, 1) <> "{" Then
            Debug.Print "[LiveMirrorDemo] " & result
            Exit For
        End If
        WriteMirror ws, CStr(result), p, Polls, Prev, LastAddress, TopRow, MaxRows, MaxCols
        CDPHelpers.Sleep WaitSec
    Next

    '6. 後始末
    CodeEditor.ThisCDPBrowser.DisconnectCDP
End Sub

Sub Demo_TakeValue()
    '1. 接続儀式を済ませる
    Dim CodeEditor As CDPContext
    Set CodeEditor = ConnectForWebExcel(UseWeb)
    If CodeEditor Is Nothing Then MsgBox "コードエディタが開いてる場所の確認が取れませんでした。事前に開く必要があります。", vbCritical: Exit Sub

    '2. 演出：実行する関数と引数を、呼び出しの形でコードエディターに表示(結果は待たない)
    If ShowCodeInEditor Then ShowInEditor CodeEditor, Sheet7.Range("A1").value, , "Demo_TakeValue"

    '3. 実行！(Excel.run は Promise を返すので、awaitPromise で完了まで待つ)
    Dim result
    CDPHelpers.SetVar(result) = CodeEditor.jsEval(Sheet7.Range("A1").value, awaitPromise:=True, returnByValue:=True, StopException:=True)
    Debug.Print result(0)
    Debug.Print result(1)
 
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] JavaScriptの計算処理を活かしたDemo
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：マンデルブロ集合を、少しずつズームしながらセルに描きます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：1フレームごとに、GridW x GridH 個のセルについて「何回で発散したか」を JavaScript で計算し、
'            Office.js でまとめて書き込みます。色は条件付き書式(カラースケール)に任せています。
'            計算は JavaScript 側なので、1フレームの計算は数十ミリ秒で終わります
'* 注意事項：・コードエディターを開いているブックに「MandelbrotJs」シートを作って描きます(既にあれば中身を消して描き直します)
'            ・描いたシートは自動では消しません。不要になったら手で削除してください
'            ・ループ中に止めたいときは、Ctrl+Break で VBA を中断してください
'***************************************************************************************************
Sub Demo_Mandelbrot()
    Const SheetName As String = "MandelbrotJs"              '描く先のシート名(コードエディターを開いているブック)
    Const Frames As Long = 120                              'フレーム数
    Const GridW As Long = 150                               '横のセル数
    Const GridH As Long = 100                               '縦のセル数
    Const CellSize As Long = 4                              'セル1つの大きさ(ポイント)
    Const CenterX As Double = -0.743643887037151            'ズームしていく中心(実部)
    Const CenterY As Double = 0.13182590420533              'ズームしていく中心(虚部)
    Const StartScale As Double = 3                          '最初に見せる範囲(縦の幅)
    Const ZoomRate As Double = 0.94                         '1フレームごとの倍率(1より小さいほど速く寄る)
    Const StartIter As Long = 120                           '計算回数の上限(最初のフレーム)
    Const IterStep As Long = 4                              '1フレームごとに増やす計算回数の上限(寄るほど細かく計算する)
    Const WaitSec As Double = 0.05                          '1フレームごとに待つ時間(秒)。Excel に描画させるための間


    '1. 接続儀式を済ませる
    Dim CodeEditor As CDPContext
    Set CodeEditor = ConnectForWebExcel(UseWeb)
    If CodeEditor Is Nothing Then MsgBox "コードエディタが開いてる場所の確認が取れませんでした。事前に開く必要があります。", vbCritical: Exit Sub

    '2. 描く先のシートを用意(シート名・横のセル数・縦のセル数・セルの大きさ を arguments で渡す)
    Dim SetupArgs(3) As Dictionary
    Set SetupArgs(0) = JsArg(SheetName)
    Set SetupArgs(1) = JsArg(GridW)
    Set SetupArgs(2) = JsArg(GridH)
    Set SetupArgs(3) = JsArg(CellSize)

    Dim result
    CDPHelpers.SetVar(result) = CodeEditor.jsEval(JS_MandelbrotSetup(), objectArguments:=SetupArgs, awaitPromise:=True)
    If IsError(result) Then
        Debug.Print "[Demo_Mandelbrot] JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
        CodeEditor.ThisCDPBrowser.DisconnectCDP: Exit Sub
    ElseIf Left$(result, 3) <> "OK:" Then
        Debug.Print "[Demo_Mandelbrot] " & result
        CodeEditor.ThisCDPBrowser.DisconnectCDP: Exit Sub
    End If

    '3. 1フレームの関数に渡す arguments のうち、フレームごとに変わらないもの
    Dim FrameArgs(7) As Dictionary
    Set FrameArgs(0) = JsArg(SheetName)
    Set FrameArgs(1) = JsArg(GridW)
    Set FrameArgs(2) = JsArg(GridH)
    Set FrameArgs(4) = JsArg(CenterX)
    Set FrameArgs(5) = JsArg(CenterY)

    '4. 計算 → 書き込み を Frames 回繰り返す(フレームごとに、計算回数の上限・範囲・見出しが変わる)
    Dim js As String: js = JS_MandelbrotFrame()
    Dim f As Long, Scale_ As Double, T0 As Single
    Scale_ = StartScale
    T0 = Timer
    For f = 1 To Frames
        Set FrameArgs(3) = JsArg(StartIter + (f - 1) * IterStep)
        Set FrameArgs(6) = JsArg(Scale_)
        Set FrameArgs(7) = JsArg("VBA -> CDP -> Office.js : Mandelbrot  x" & Format$(StartScale / Scale_, "0") & "  (" & f & "/" & Frames & ")")

        '演出：実行する関数と引数をコードエディターに表示
        If ShowCodeInEditor Then ShowInEditor CodeEditor, js, FrameArgs, "Demo_Mandelbrot (" & Frames & " frames)"

        CDPHelpers.SetVar(result) = CodeEditor.jsEval(js, objectArguments:=FrameArgs, awaitPromise:=True)
        If IsError(result) Then
            Debug.Print "[Demo_Mandelbrot] JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
            Exit For
        ElseIf Left$(result, 3) <> "OK:" Then
            Debug.Print "[Demo_Mandelbrot] " & result
            Exit For
        End If
        Scale_ = Scale_ * ZoomRate
        CDPHelpers.Sleep WaitSec
    Next
    If f > Frames Then Debug.Print "[Demo_Mandelbrot] " & Frames & " frames / " & Format$(Timer - T0, "0.0") & " 秒 (最後のフレーム：" & result & ")"

    '5. 後始末
    CodeEditor.ThisCDPBrowser.DisconnectCDP
End Sub

'***************************************************************************************************
'* 機能　　：ライフゲーム(グライダー銃とパルサー)を、セルの上で世代を進めながら描きます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：1フレームごとに、シートの今の状態を読み、次の世代を JavaScript で計算して書き戻します。
'            セルの値は「生き続けている世代数」で、色は条件付き書式(カラースケール)に任せています(若い:緑 → 長生き:紫)
'* 注意事項：・コードエディターを開いているブックに「LifeGameJs」シートを作って描きます(既にあれば中身を消して描き直します)
'            ・状態はシートが持っているので、動いている途中に盤面のセルへ 1 を入力すると、そこにも命が生まれます
'            ・描いたシートは自動では消しません。不要になったら手で削除してください
'            ・ループ中に止めたいときは、Ctrl+Break で VBA を中断してください
'***************************************************************************************************
Sub Demo_LifeGame()
    Const SheetName As String = "LifeGameJs"                '描く先のシート名(コードエディターを開いているブック)
    Const Frames As Long = 240                              'フレーム数
    Const StepsPerFrame As Long = 1                         '1フレームで進める世代数
    Const GridW As Long = 80                                '横のセル数
    Const GridH As Long = 50                                '縦のセル数
    Const CellSize As Long = 8                              'セル1つの大きさ(ポイント)
    Const WaitSec As Double = 0.05                          '1フレームごとに待つ時間(秒)。Excel に描画させるための間


    '1. 接続儀式を済ませる
    Dim CodeEditor As CDPContext
    Set CodeEditor = ConnectForWebExcel(UseWeb)
    If CodeEditor Is Nothing Then MsgBox "コードエディタが開いてる場所の確認が取れませんでした。事前に開く必要があります。", vbCritical: Exit Sub

    '2. 描く先のシートを用意して、最初の模様を置く(シート名・横のセル数・縦のセル数・セルの大きさ を arguments で渡す)
    Dim SetupArgs(3) As Dictionary
    Set SetupArgs(0) = JsArg(SheetName)
    Set SetupArgs(1) = JsArg(GridW)
    Set SetupArgs(2) = JsArg(GridH)
    Set SetupArgs(3) = JsArg(CellSize)

    Dim result
    CDPHelpers.SetVar(result) = CodeEditor.jsEval(JS_LifeGameSetup(), objectArguments:=SetupArgs, awaitPromise:=True)
    If IsError(result) Then
        Debug.Print "[Demo_LifeGame] JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
        CodeEditor.ThisCDPBrowser.DisconnectCDP: Exit Sub
    ElseIf Left$(result, 3) <> "OK:" Then
        Debug.Print "[Demo_LifeGame] " & result
        CodeEditor.ThisCDPBrowser.DisconnectCDP: Exit Sub
    End If

    '3. 1フレームの関数に渡す arguments のうち、フレームごとに変わらないもの
    Dim StepArgs(4) As Dictionary
    Set StepArgs(0) = JsArg(SheetName)
    Set StepArgs(1) = JsArg(GridW)
    Set StepArgs(2) = JsArg(GridH)
    Set StepArgs(3) = JsArg(StepsPerFrame)

    '4. 読み取り → 次の世代の計算 → 書き戻し を Frames 回繰り返す(フレームごとに変わるのは見出しだけ)
    Dim js As String: js = JS_LifeGameStep()
    Dim f As Long, T0 As Single
    T0 = Timer
    For f = 1 To Frames
        Set StepArgs(4) = JsArg("VBA -> CDP -> Office.js : Game of Life  gen " & f * StepsPerFrame)

        '演出：実行する関数と引数をコードエディターに表示
        If ShowCodeInEditor Then ShowInEditor CodeEditor, js, StepArgs, "Demo_LifeGame (" & Frames & " frames)"

        CDPHelpers.SetVar(result) = CodeEditor.jsEval(js, objectArguments:=StepArgs, awaitPromise:=True)
        If IsError(result) Then
            Debug.Print "[Demo_LifeGame] JavaScript exception: " & CodeEditor.LastJavaScriptException.Stringify
            Exit For
        ElseIf Left$(result, 3) <> "OK:" Then
            Debug.Print "[Demo_LifeGame] " & result
            Exit For
        End If
        CDPHelpers.Sleep WaitSec
    Next
    If f > Frames Then Debug.Print "[Demo_LifeGame] " & Frames & " frames / " & Format$(Timer - T0, "0.0") & " 秒 (最後のフレーム：" & result & ")"

    '5. 後始末
    CodeEditor.ThisCDPBrowser.DisconnectCDP
End Sub




'***************************************************************************************************
'                         ■■■ 各種Demoヘルパー ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 共通。`Runtime.callFunctionOn`の`arguments`作り
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：`arguments`の1要素({"value": 値})を作ります。引数が多いデモで、Dictionary を1つずつ書く手間を省く用
'***************************************************************************************************
Private Function JsArg(value As Variant) As Dictionary
    Set JsArg = New Dictionary
    JsArg.Add "value", value
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] Demo_ViewBookInfo 用ヘルパー。結果の書き出し
'---------------------------------------------------------------------------------------------------
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
'* 引数　　：Json           ViewBookInfoDemo の結果(JSON)
'            SheetName      書き出し先のシート名(無ければ作り、あれば中身を消して書き直す)
'***************************************************************************************************
Private Sub WriteInspectReport(Json As String, SheetName As String)
    Const FirstRow As Long = 4


    '1. JSON を読む
    Dim Root As BiDiCDPJConv, SheetList As BiDiCDPJConv
    Set Root = BiDiCDPJConv.Parse(Json)
    Set SheetList = Root.NodeKey("sheets")
    If SheetList Is Nothing Then
        Debug.Print "[Showcase] 棚卸しの JSON に sheets がありません: " & Left$(Json, 200)
        Exit Sub
    End If

    '2. 見出し
    Dim ws As Worksheet
    Set ws = PrepareLocalSheet(SheetName)
    ws.Range("A1").value = "ブックの棚卸し(Office.js)：" & Root.StringKey("workbook")
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

        Debug.Print "[Showcase] " & Node.StringKey("name") & " : " & Node.StringKey("used") & " / 数式 " & ws.Cells(r, 6).value & " / エラー " & ws.Cells(r, 7).value & IIf(Len(List), " (" & List & ")", "")
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

'---------------------------------------------------------------------------------------------------
' [ SECTION ] Demo_LiveMirror 用ヘルパー
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：ライブミラーの JSON を、このマクロのブックのシートに写します。前回から変わったセルは黄色にします
'---------------------------------------------------------------------------------------------------
'* 引数　　：ws                 写す先のシート
'            Json               LiveMirrorDemo の結果(JSON)
'            Poll / Polls       何回目 / 全部で何回(状況表示用)
'            Prev / LastAddress (ByRef) 前回の内容と範囲。呼び出し側で保持して、次の回に渡す
'            TopRow             写す先の開始行
'            MaxRows / MaxCols  読み取る最大の行数 / 列数(前回の光と内容を消す範囲)
'***************************************************************************************************
Private Sub WriteMirror(ws As Worksheet, Json As String, Poll As Long, Polls As Long, Prev As Variant, LastAddress As String, TopRow As Long, MaxRows As Long, MaxCols As Long)
    If Len(Json) = 0 Then Exit Sub

    '1. JSON を読む
    Dim Root As BiDiCDPJConv, Texts As BiDiCDPJConv, Address As String
    Set Root = BiDiCDPJConv.Parse(Json)
    Address = Root.StringKey("address")
    ws.Range("A2").value = "読み取り " & Poll & "/" & Polls & "　" & Format$(Now, "hh:nn:ss") & "　" & Root.StringKey("sheet") & "!" & Address

    '2. 前回の光を消す
    Dim Area As Range
    Set Area = ws.Cells(TopRow, 1).Resize(MaxRows, MaxCols)
    Area.Interior.ColorIndex = xlNone

    Set Texts = Root.NodeKey("texts")
    If Texts Is Nothing Then Exit Sub
    If Texts.Count = 0 Then
        Area.ClearContents
        Prev = Empty
        LastAddress = Address
        Exit Sub
    End If

    '3. JSON の2次元配列 → VBA の配列
    Dim RowsN As Long, ColsN As Long, Cur() As Variant
    RowsN = Texts.Count
    ColsN = Texts.NodeAt(0).Count
    ReDim Cur(1 To RowsN, 1 To ColsN)

    Dim r As Long, c As Long, RowTok As Long, CellTok As Long, RowNode As BiDiCDPJConv
    RowTok = Texts.FirstChildToken
    Do While RowTok > 0
        r = r + 1
        Set RowNode = Texts.NodeFromToken(RowTok)
        c = 0
        CellTok = RowNode.FirstChildToken
        Do While CellTok > 0
            c = c + 1
            If c <= ColsN Then Cur(r, c) = RowNode.TokenStringValue(CellTok)
            CellTok = RowNode.NextToken(CellTok)
        Loop
        RowTok = Texts.NextToken(RowTok)
    Loop

    '4. 書き出し(範囲が変わったら、先に前回の分を消す。文字のまま写すため、表示形式は文字列にする)
    If Address <> LastAddress Then
        Area.ClearContents
        Prev = Empty
    End If
    Dim Target As Range
    Set Target = ws.Cells(TopRow, 1).Resize(RowsN, ColsN)
    Target.NumberFormat = "@"
    Target.value = Cur

    '5. 前回から変わったセルを光らせる(同じ範囲のときだけ)
    If IsArray(Prev) Then
        For r = 1 To RowsN
            For c = 1 To ColsN
                If Cur(r, c) <> Prev(r, c) Then Target.Cells(r, c).Interior.Color = RGB(255, 235, 59)
            Next
        Next
    End If
    Prev = Cur
    LastAddress = Address
End Sub



'***************************************************************************************************
'                         ■■■ 演出(コードエディターへの表示) ■■■
'***************************************************************************************************
'* 機能　　：`Runtime.callFunctionOn`で実行する関数と引数を、「関数をその引数で呼んでいる形」にして、
'            コードエディターに表示します。見た目のための演出で、実際の実行はエディターではなく Office.js で行います
'            引数は、長い関数でもスクロールせずに見えるよう、先頭に「const 引数名 = 値;」として並べます
'                // 見出し…
'                const text = "Hello";
'                const when = "2026-11-09T12:34:56";
'                (async function (text, when) { … })(text, when);
'---------------------------------------------------------------------------------------------------
'* 引数　　：CodeEditor     コードエディターのページ
'            FunctionJs     `jsEval`の`JavaScriptStr`に渡す関数宣言(例：JS_SimpleDemo())
'            Args           `jsEval`の`objectArguments`に渡す配列(各要素は "value" を持つ Dictionary)
'            Caption        表示の3行目に入れる見出し
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・引数名は、関数宣言の最初の「( … )」から取り出します。取り出せない分は arg0, arg1 … になります
'            ・値は`BiDiCDPJConv.StringifyValue`で JSON にします。CDP に送る値と同じ中身で、日本語は読める形のまま表示されます
'              (日時は、CDP に送るときと同じく "yyyy-mm-ddTHH:nn:ss" の文字列になります)
'* 注意事項：・表示は結果を待たない CDP(`RunAsyncCDP:=True`)で送るため、実行の速さはほぼ変わりません
'            ・エディターの中身は表示用のコードで上書きされ、Officeスクリプトとして自動保存されます
'            ・表示するのは Office.js のコードなので、エディターの「問題」に警告が出ます。実行には影響しません
'***************************************************************************************************
Private Sub ShowInEditor(CodeEditor As CDPContext, ByVal FunctionJs As String, Optional args As Variant, Optional Caption As String)
    '1. 関数宣言の最初の「( … )」から、引数名を取り出す(既定値「x = 1」は名前だけにする)
    Dim ParamNames() As String, ParamCount As Long, p1 As Long, p2 As Long
    p1 = InStr(1, FunctionJs, "(")
    If p1 > 0 Then p2 = InStr(p1 + 1, FunctionJs, ")")
    If p2 > p1 + 1 Then
        ParamNames = Split(Mid$(FunctionJs, p1 + 1, p2 - p1 - 1), ",")
        ParamCount = UBound(ParamNames) + 1
    End If

    '2. 引数を「const 引数名 = 値;」の行にする(日本語が \uXXXX にならないよう、表示用の値は BiDiCDPJConv で JSON にする)
    Dim ArgLines As String, CallList As String, ArgName As String, i As Long, N As Long
    If Not IsMissing(args) Then
        For i = LBound(args) To UBound(args)
            N = i - LBound(args)
            ArgName = ""
            If N < ParamCount Then
                If Len(Trim$(ParamNames(N))) Then ArgName = Trim$(Split(ParamNames(N), "=")(0))
            End If
            If Len(ArgName) = 0 Then ArgName = "arg" & N
            ArgLines = ArgLines & "const " & ArgName & " = " & BiDiCDPJConv.StringifyValue(args(i)("value")) & ";" & vbLf
            If N > 0 Then CallList = CallList & ", "
            CallList = CallList & ArgName
        Next
    End If

    '3. 表示するテキスト(関数宣言の末尾の改行は外して、閉じかっこの直後に引数名を続ける)
    Do While Right$(FunctionJs, 1) = vbLf
        FunctionJs = Left$(FunctionJs, Len(FunctionJs) - 1)
    Loop
    Dim Text As String
    Text = "// ===== VBA -> CDP -> Office.js (Runtime.callFunctionOn) =====" & vbLf & _
           "// 表示用です。実行はこのエディターではなく、Office.js で直接行っています" & vbLf & _
           "// " & Caption & IIf(Len(Caption), "  ", "") & Format$(Now, "yyyy/mm/dd hh:nn:ss") & vbLf & _
           ArgLines & _
           "(" & FunctionJs & ")(" & CallList & ");" & vbLf

    '4. エディターに送る(結果は待たない)
    Dim v As Variant: v = Text
    CodeEditor.jsEval Replace(JS_ShowInEditor(), "__TEXT__", WebJsonConverter.serialize(v)), RunAsyncCDP:=True
End Sub

'***************************************************************************************************
'* 機能　　：演出用。Monaco エディター(コードエディター)の全文を置き換えます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：__TEXT__ に JSON 文字列リテラル(WebJsonConverter.serialize の結果)を入れて使います。
'            エディターのインスタンスは React の内部情報からさかのぼって探し、2回目からはページ内に覚えておいたものを使います
'* 返り値　："OK" / "same"(同じ内容) / "NG: editor not found"
'***************************************************************************************************
Private Function JS_ShowInEditor() As String
    InitJSList(vbLf) = 21    '行数 - 1(JSList の上限)
    SetJS = "(() => {"
    SetJS = "    const W = window, K = ""__exCDP_OfficeJsShow"";"
    SetJS = "    const isEd = (v) => !!v && typeof v === ""object"" && typeof v.getModel === ""function"" && typeof v.executeEdits === ""function"";"
    SetJS = "    let ed = W[K];"
    SetJS = "    if (!isEd(ed) || !ed.getModel()) {"
    SetJS = "        ed = null;"
    SetJS = "        const fiberOf = (el) => { const k = Object.keys(el).find((x) => x.startsWith(""__reactFiber$"")); return k ? el[k] : null; };"
    SetJS = "        for (let el = document.querySelector("".monaco-editor""); el && !ed; el = el.parentElement) {"
    SetJS = "            for (let f = fiberOf(el), d = 0; f && d < 40 && !ed; f = f.return, d++) {"
    SetJS = "                const sn = f.stateNode;"
    SetJS = "                if (sn && !(sn instanceof Node)) { if (isEd(sn.editor)) { ed = sn.editor; } else if (isEd(sn)) { ed = sn; } }"
    SetJS = "            }"
    SetJS = "        }"
    SetJS = "        if (!ed || !ed.getModel()) { return ""NG: editor not found""; }"
    SetJS = "        W[K] = ed;"
    SetJS = "    }"
    SetJS = "    const m = ed.getModel(), t = __TEXT__;"
    SetJS = "    if (m.getValue() === t) { return ""same""; }"
    SetJS = "    ed.executeEdits(""exCDP_OfficeJs"", [{ range: m.getFullModelRange(), text: t, forceMoveMarkers: true }]);"
    SetJS = "    ed.setScrollTop(0);"
    SetJS = "    return ""OK"";"
    SetJS = "})()"
    JS_ShowInEditor = SetJS
End Function



'***************************************************************************************************
'               ■■■ Office.jsに渡すテンプレートJavaScriptコード一式 ■■■
'***************************************************************************************************
'* 機能　　：Demo_Simple 用。アクティブシートの A1 に文字列、A2 に日時を書きます
'---------------------------------------------------------------------------------------------------
'* 引数　　：text   A1 に書く文字列(arguments の1つ目)
'            when   A2 に書く日時(arguments の2つ目)。VBA の Date は "yyyy-mm-ddTHH:nn:ss" の文字列で届くため、
'                   JavaScript 側で Excel のシリアル値に直してから書きます
'* 返り値　："OK: シート名!A1:A2" または "NG: <エラーコード>"
'***************************************************************************************************
Private Function JS_SimpleDemo() As String
    InitJSList(vbLf) = 24    '行数 - 1(JSList の上限)
    SetJS = "async function (text, when) {"
    SetJS = "    // VBA の Date は ""yyyy-mm-ddTHH:nn:ss""(タイムゾーンなし)の文字列で届くので、Excel のシリアル値に直す"
    SetJS = "    const toSerial = (v) => {"
    SetJS = "        if (typeof v === ""number"") { return v; }"
    SetJS = "        const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/.exec(String(v));"
    SetJS = "        if (!m) { return v; }"
    SetJS = "        return (Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]) - Date.UTC(1899, 11, 30)) / 86400000;"
    SetJS = "    };"
    SetJS = "    try {"
    SetJS = "        return await Excel.run(async (context) => {"
    SetJS = "            const ws = context.workbook.worksheets.getActiveWorksheet();"
    SetJS = "            const a1 = ws.getRange(""A1""), a2 = ws.getRange(""A2"");"
    SetJS = "            a1.values = [[text]];"
    SetJS = "            a1.format.font.bold = true;"
    SetJS = "            a2.values = [[toSerial(when)]];"
    SetJS = "            a2.numberFormat = [[""yyyy/mm/dd hh:mm:ss""]];"
    SetJS = "            ws.getRange(""A1:A2"").format.autofitColumns();"
    SetJS = "            ws.load(""name"");"
    SetJS = "            await context.sync();"
    SetJS = "            return ""OK: "" + ws.name + ""!A1:A2"";"
    SetJS = "        });"
    SetJS = "    } catch (e) {"
    SetJS = "        return ""NG: "" + String((e && (e.code || e.message)) || e);"
    SetJS = "    }"
    SetJS = "}"
    JS_SimpleDemo = SetJS
End Function

'***************************************************************************************************
'* 機能　　：Demo_ViewBookInfo 用。全シートの使用範囲・数式・エラー・テーブル・グラフを調べ、JSON で返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：{workbook, names, sheets:[{name, used, rows, cols, formulas, errors, errorCells, tables, charts, scanned}], ms}
'            または "NG: <エラーコード>"
'***************************************************************************************************
Private Function JS_ViewBookInfo() As String
    InitJSList(vbLf) = 51    '行数 - 1(JSList の上限)
    SetJS = "async function () {"
    SetJS = "    // 全シートの使用範囲・数式・エラー・テーブル・グラフを調べて JSON で返す(ブックは書き換えない)"
    SetJS = "    const LIMIT = 200000, MAX_ERR = 5, T0 = Date.now();"
    SetJS = "    const col = (n) => { let s = """"; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };"
    SetJS = "    try {"
    SetJS = "        return await Excel.run(async (context) => {"
    SetJS = "            // 1回目の sync：シート一覧・名前の定義"
    SetJS = "            const wb = context.workbook; wb.load(""name"");"
    SetJS = "            const sheets = wb.worksheets; sheets.load(""items/name"");"
    SetJS = "            const names = wb.names; names.load(""items/name"");"
    SetJS = "            await context.sync();"
    SetJS = ""
    SetJS = "            // 2回目の sync：シートごとの使用範囲・テーブル・グラフ"
    SetJS = "            const infos = sheets.items.map((ws) => {"
    SetJS = "                const used = ws.getUsedRangeOrNullObject(true); used.load(""address,rowCount,columnCount,rowIndex,columnIndex"");"
    SetJS = "                const tables = ws.tables; tables.load(""items/name"");"
    SetJS = "                const charts = ws.charts; charts.load(""items/name"");"
    SetJS = "                return { ws: ws, used: used, tables: tables, charts: charts, scan: false };"
    SetJS = "            });"
    SetJS = "            await context.sync();"
    SetJS = ""
    SetJS = "            // 3回目の sync：使用範囲が LIMIT 以下のシートだけ、数式と値の種類を読む"
    SetJS = "            for (const i of infos) { if (!i.used.isNullObject && i.used.rowCount * i.used.columnCount <= LIMIT) { i.used.load(""valueTypes,formulas""); i.scan = true; } }"
    SetJS = "            await context.sync();"
    SetJS = ""
    SetJS = "            const out = infos.map((i) => {"
    SetJS = "                const info = { name: i.ws.name, used: """", rows: 0, cols: 0, formulas: 0, errors: 0, errorCells: [], tables: i.tables.items.length, charts: i.charts.items.length, scanned: i.scan };"
    SetJS = "                if (!i.used.isNullObject) {"
    SetJS = "                    info.used = String(i.used.address).split(""!"").pop();"
    SetJS = "                    info.rows = i.used.rowCount;"
    SetJS = "                    info.cols = i.used.columnCount;"
    SetJS = "                    if (i.scan) {"
    SetJS = "                        for (let r = 0; r < info.rows; r++) {"
    SetJS = "                            for (let c = 0; c < info.cols; c++) {"
    SetJS = "                                const f = i.used.formulas[r][c];"
    SetJS = "                                if (typeof f === ""string"" && f.charAt(0) === ""="") { info.formulas++; }"
    SetJS = "                                if (i.used.valueTypes[r][c] === ""Error"") {"
    SetJS = "                                    info.errors++;"
    SetJS = "                                    if (info.errorCells.length < MAX_ERR) { info.errorCells.push(col(i.used.columnIndex + c) + (i.used.rowIndex + r + 1)); }"
    SetJS = "                                }"
    SetJS = "                            }"
    SetJS = "                        }"
    SetJS = "                    }"
    SetJS = "                }"
    SetJS = "                return info;"
    SetJS = "            });"
    SetJS = "            return JSON.stringify({ workbook: wb.name, names: names.items.length, sheets: out, ms: Date.now() - T0 });"
    SetJS = "        });"
    SetJS = "    } catch (e) {"
    SetJS = "        return ""NG: "" + String((e && (e.code || e.message)) || e);"
    SetJS = "    }"
    SetJS = "}"
    JS_ViewBookInfo = SetJS
End Function

'***************************************************************************************************
'* 機能　　：Demo_LiveMirror 用。アクティブなシートの使用範囲の左上から、表示されている文字を JSON で返します
'---------------------------------------------------------------------------------------------------
'* 引数　　：maxRows / maxCols   読み取る最大の行数 / 列数(arguments の1つ目 / 2つ目)
'* 返り値　：{sheet, address, texts:[[…]]}  または "NG: <エラーコード>"
'***************************************************************************************************
Private Function JS_LiveMirror() As String
    InitJSList(vbLf) = 17    '行数 - 1(JSList の上限)
    SetJS = "async function (maxRows, maxCols) {"
    SetJS = "    // アクティブなシートの使用範囲の左上から、最大 maxRows 行 x maxCols 列の「表示されている文字」を JSON で返す(ブックは書き換えない)"
    SetJS = "    try {"
    SetJS = "        return await Excel.run(async (context) => {"
    SetJS = "            const ws = context.workbook.worksheets.getActiveWorksheet(); ws.load(""name"");"
    SetJS = "            const used = ws.getUsedRangeOrNullObject(true); used.load(""rowCount,columnCount,rowIndex,columnIndex"");"
    SetJS = "            await context.sync();"
    SetJS = "            if (used.isNullObject) { return JSON.stringify({ sheet: ws.name, address: """", texts: [] }); }"
    SetJS = ""
    SetJS = "            const area = ws.getRangeByIndexes(used.rowIndex, used.columnIndex, Math.min(used.rowCount, maxRows), Math.min(used.columnCount, maxCols));"
    SetJS = "            area.load(""address,text"");"
    SetJS = "            await context.sync();"
    SetJS = "            return JSON.stringify({ sheet: ws.name, address: String(area.address).split(""!"").pop(), texts: area.text });"
    SetJS = "        });"
    SetJS = "    } catch (e) {"
    SetJS = "        return ""NG: "" + String((e && (e.code || e.message)) || e);"
    SetJS = "    }"
    SetJS = "}"
    JS_LiveMirror = SetJS
End Function

'***************************************************************************************************
'* 機能　　：Demo_Mandelbrot 用。描く先のシートを用意します(無ければ作り、あれば中身を消す)
'---------------------------------------------------------------------------------------------------
'* 引数　　：sheetName / w / h / cellSize   シート名 / 横のセル数 / 縦のセル数 / セル1つの大きさ(ポイント)
'* 返り値　："OK: <シート名>"  または "NG: <エラーコード>"
'***************************************************************************************************
Private Function JS_MandelbrotSetup() As String
    InitJSList(vbLf) = 36    '行数 - 1(JSList の上限)
    SetJS = "async function (sheetName, w, h, cellSize) {"
    SetJS = "    // 描く先のシートを用意する(無ければ作る)。セルを小さな正方形にし、値の大小を条件付き書式(カラースケール)の色で見せる"
    SetJS = "    try {"
    SetJS = "        return await Excel.run(async (context) => {"
    SetJS = "            const wb = context.workbook;"
    SetJS = "            let ws = wb.worksheets.getItemOrNullObject(sheetName);"
    SetJS = "            await context.sync();"
    SetJS = "            if (ws.isNullObject) { ws = wb.worksheets.add(sheetName); }"
    SetJS = "            ws.activate();"
    SetJS = "            ws.getRange().clear(Excel.ClearApplyTo.all);"
    SetJS = "            ws.showGridlines = false;"
    SetJS = ""
    SetJS = "            const title = ws.getRangeByIndexes(0, 0, 1, w);"
    SetJS = "            title.merge(false);"
    SetJS = "            title.getCell(0, 0).values = [[""VBA -> CDP -> Office.js : Mandelbrot""]];"
    SetJS = "            title.format.font.bold = true;"
    SetJS = "            title.format.font.size = 14;"
    SetJS = ""
    SetJS = "            const grid = ws.getRangeByIndexes(1, 0, h, w);"
    SetJS = "            grid.format.columnWidth = cellSize;"
    SetJS = "            grid.format.rowHeight = cellSize;"
    SetJS = "            grid.numberFormat = "";;;"";"
    SetJS = "            grid.format.fill.color = ""#000000"";"
    SetJS = "            const cf = grid.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);"
    SetJS = "            cf.colorScale.criteria = {"
    SetJS = "                minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: ""#020B3A"" },"
    SetJS = "                midpoint: { formula: ""85"", type: Excel.ConditionalFormatColorCriterionType.percentile, color: ""#FF9A1F"" },"
    SetJS = "                maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: ""#FFFFF0"" }"
    SetJS = "            };"
    SetJS = "            ws.getRange(""A1"").select();"
    SetJS = "            await context.sync();"
    SetJS = "            return ""OK: "" + sheetName;"
    SetJS = "        });"
    SetJS = "    } catch (e) {"
    SetJS = "        return ""NG: "" + String((e && (e.code || e.message)) || e);"
    SetJS = "    }"
    SetJS = "}"
    JS_MandelbrotSetup = SetJS
End Function

'***************************************************************************************************
'* 機能　　：Demo_Mandelbrot 用。1フレーム分を JavaScript で計算し、セルにまとめて書き込みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：sheetName / w / h    シート名 / 横のセル数 / 縦のセル数
'            maxIter              計算回数の上限
'            cx / cy / scale      描く範囲の中心(実部 / 虚部)と、縦の幅
'            caption              A1 に書く見出し
'* 返り値　："OK: calc <計算の ms> / total <書き込みまでの ms>"  または "NG: <エラーコード>"
'***************************************************************************************************
Private Function JS_MandelbrotFrame() As String
    InitJSList(vbLf) = 34    '行数 - 1(JSList の上限)
    SetJS = "async function (sheetName, w, h, maxIter, cx, cy, scale, caption) {"
    SetJS = "    // 中心(cx, cy)・縦の幅 scale の範囲を w x h のセルに割り当て、セルごとに「何回で発散したか」を計算して書き込む(色は条件付き書式に任せる)"
    SetJS = "    const T0 = Date.now();"
    SetJS = "    const rows = [];"
    SetJS = "    for (let y = 0; y < h; y++) {"
    SetJS = "        const row = [];"
    SetJS = "        const ci = cy + (y / h - 0.5) * scale;"
    SetJS = "        for (let x = 0; x < w; x++) {"
    SetJS = "            const cr = cx + (x / w - 0.5) * scale * (w / h);"
    SetJS = "            let zr = 0, zi = 0, i = 0;"
    SetJS = "            while (zr * zr + zi * zi < 16 && i < maxIter) {"
    SetJS = "                const t = zr * zr - zi * zi + cr;"
    SetJS = "                zi = 2 * zr * zi + ci;"
    SetJS = "                zr = t;"
    SetJS = "                i++;"
    SetJS = "            }"
    SetJS = "            // 集合の内側は空白(黒)。外側は、色の段差が目立たないように回数を滑らかにした値"
    SetJS = "            row.push(i >= maxIter ? """" : Math.round(Math.sqrt(i + 1 - Math.log(Math.log(Math.sqrt(zr * zr + zi * zi))) / Math.LN2) * 1000) / 1000);"
    SetJS = "        }"
    SetJS = "        rows.push(row);"
    SetJS = "    }"
    SetJS = "    const calcMs = Date.now() - T0;"
    SetJS = ""
    SetJS = "    try {"
    SetJS = "        return await Excel.run(async (context) => {"
    SetJS = "            const ws = context.workbook.worksheets.getItem(sheetName);"
    SetJS = "            ws.getRangeByIndexes(1, 0, h, w).values = rows;"
    SetJS = "            ws.getRange(""A1"").values = [[caption]];"
    SetJS = "            await context.sync();"
    SetJS = "            return ""OK: calc "" + calcMs + ""ms / total "" + (Date.now() - T0) + ""ms"";"
    SetJS = "        });"
    SetJS = "    } catch (e) {"
    SetJS = "        return ""NG: "" + String((e && (e.code || e.message)) || e);"
    SetJS = "    }"
    SetJS = "}"
    JS_MandelbrotFrame = SetJS
End Function

'***************************************************************************************************
'* 機能　　：Demo_LifeGame 用。描く先のシートを用意し、グライダー銃とパルサーを置きます
'---------------------------------------------------------------------------------------------------
'* 引数　　：sheetName / w / h / cellSize   シート名 / 横のセル数 / 縦のセル数 / セル1つの大きさ(ポイント)
'* 返り値　："OK: <シート名>"  または "NG: <エラーコード>"
'***************************************************************************************************
Private Function JS_LifeGameSetup() As String
    InitJSList(vbLf) = 74    '行数 - 1(JSList の上限)
    SetJS = "async function (sheetName, w, h, cellSize) {"
    SetJS = "    // 描く先のシートを用意し、グライダー銃(左上)とパルサー(右上)を置く。セルの値は「生きている世代数」で、色は条件付き書式(カラースケール)に任せる"
    SetJS = "    try {"
    SetJS = "        return await Excel.run(async (context) => {"
    SetJS = "            const wb = context.workbook;"
    SetJS = "            let ws = wb.worksheets.getItemOrNullObject(sheetName);"
    SetJS = "            await context.sync();"
    SetJS = "            if (ws.isNullObject) { ws = wb.worksheets.add(sheetName); }"
    SetJS = "            ws.activate();"
    SetJS = "            ws.getRange().clear(Excel.ClearApplyTo.all);"
    SetJS = "            ws.showGridlines = false;"
    SetJS = ""
    SetJS = "            const title = ws.getRangeByIndexes(0, 0, 1, w);"
    SetJS = "            title.merge(false);"
    SetJS = "            title.getCell(0, 0).values = [[""VBA -> CDP -> Office.js : Game of Life  gen 0""]];"
    SetJS = "            title.format.font.bold = true;"
    SetJS = "            title.format.font.size = 14;"
    SetJS = ""
    SetJS = "            const grid = ws.getRangeByIndexes(1, 0, h, w);"
    SetJS = "            grid.format.columnWidth = cellSize;"
    SetJS = "            grid.format.rowHeight = cellSize;"
    SetJS = "            grid.numberFormat = "";;;"";"
    SetJS = "            grid.format.fill.color = ""#05050F"";"
    SetJS = "            const cf = grid.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);"
    SetJS = "            cf.colorScale.criteria = {"
    SetJS = "                minimum: { formula: ""1"", type: Excel.ConditionalFormatColorCriterionType.number, color: ""#39FF14"" },"
    SetJS = "                midpoint: { formula: ""6"", type: Excel.ConditionalFormatColorCriterionType.number, color: ""#00B4FF"" },"
    SetJS = "                maximum: { formula: ""30"", type: Excel.ConditionalFormatColorCriterionType.number, color: ""#B000FF"" }"
    SetJS = "            };"
    SetJS = ""
    SetJS = "            // 「O」の場所を 1(生まれたばかり)にする"
    SetJS = "            const cells = [];"
    SetJS = "            for (let y = 0; y < h; y++) { const row = []; for (let x = 0; x < w; x++) { row.push(""""); } cells.push(row); }"
    SetJS = "            const put = (pattern, top, left) => {"
    SetJS = "                for (let y = 0; y < pattern.length; y++) {"
    SetJS = "                    for (let x = 0; x < pattern[y].length; x++) {"
    SetJS = "                        if (pattern[y][x] === ""O"" && top + y < h && left + x < w) { cells[top + y][left + x] = 1; }"
    SetJS = "                    }"
    SetJS = "                }"
    SetJS = "            };"
    SetJS = "            put(["
    SetJS = "                ""........................O..........."","
    SetJS = "                ""......................O.O..........."","
    SetJS = "                ""............OO......OO............OO"","
    SetJS = "                ""...........O...O....OO............OO"","
    SetJS = "                ""OO........O.....O...OO.............."","
    SetJS = "                ""OO........O...O.OO....O.O..........."","
    SetJS = "                ""..........O.....O.......O..........."","
    SetJS = "                ""...........O...O...................."","
    SetJS = "                ""............OO......................"""
    SetJS = "            ], 2, 2);"
    SetJS = "            put(["
    SetJS = "                ""..OOO...OOO.."","
    SetJS = "                ""............."","
    SetJS = "                ""O....O.O....O"","
    SetJS = "                ""O....O.O....O"","
    SetJS = "                ""O....O.O....O"","
    SetJS = "                ""..OOO...OOO.."","
    SetJS = "                ""............."","
    SetJS = "                ""..OOO...OOO.."","
    SetJS = "                ""O....O.O....O"","
    SetJS = "                ""O....O.O....O"","
    SetJS = "                ""O....O.O....O"","
    SetJS = "                ""............."","
    SetJS = "                ""..OOO...OOO.."""
    SetJS = "            ], 3, w - 16);"
    SetJS = "            grid.values = cells;"
    SetJS = "            ws.getRange(""A1"").select();"
    SetJS = "            await context.sync();"
    SetJS = "            return ""OK: "" + sheetName;"
    SetJS = "        });"
    SetJS = "    } catch (e) {"
    SetJS = "        return ""NG: "" + String((e && (e.code || e.message)) || e);"
    SetJS = "    }"
    SetJS = "}"
    JS_LifeGameSetup = SetJS
End Function

'***************************************************************************************************
'* 機能　　：Demo_LifeGame 用。シートの今の状態を読み、steps 世代進めて書き戻します
'---------------------------------------------------------------------------------------------------
'* 引数　　：sheetName / w / h    シート名 / 横のセル数 / 縦のセル数
'            steps                進める世代数
'            caption              A1 に書く見出し
'* 返り値　："OK: alive <生きているセル数> / calc <計算の ms> / total <読み書き込みの ms>"  または "NG: <エラーコード>"
'***************************************************************************************************
Private Function JS_LifeGameStep() As String
    InitJSList(vbLf) = 44    '行数 - 1(JSList の上限)
    SetJS = "async function (sheetName, w, h, steps, caption) {"
    SetJS = "    // シートの今の状態を読み、steps 世代ぶん JavaScript で進めて書き戻す(状態はシートが持つので、途中でセルに 1 を入れると、そこにも命が生まれる)"
    SetJS = "    try {"
    SetJS = "        return await Excel.run(async (context) => {"
    SetJS = "            const T0 = Date.now();"
    SetJS = "            const ws = context.workbook.worksheets.getItem(sheetName);"
    SetJS = "            const grid = ws.getRangeByIndexes(1, 0, h, w);"
    SetJS = "            grid.load(""values"");"
    SetJS = "            await context.sync();"
    SetJS = ""
    SetJS = "            const T1 = Date.now();"
    SetJS = "            let age = grid.values.map((r) => r.map((v) => (typeof v === ""number"" && v > 0) ? v : 0));"
    SetJS = "            for (let s = 0; s < steps; s++) {"
    SetJS = "                const next = [];"
    SetJS = "                for (let y = 0; y < h; y++) {"
    SetJS = "                    const row = [];"
    SetJS = "                    for (let x = 0; x < w; x++) {"
    SetJS = "                        let n = 0;"
    SetJS = "                        for (let dy = -1; dy <= 1; dy++) {"
    SetJS = "                            for (let dx = -1; dx <= 1; dx++) {"
    SetJS = "                                if (dx === 0 && dy === 0) { continue; }"
    SetJS = "                                const yy = y + dy, xx = x + dx;"
    SetJS = "                                if (yy >= 0 && yy < h && xx >= 0 && xx < w && age[yy][xx] > 0) { n++; }"
    SetJS = "                            }"
    SetJS = "                        }"
    SetJS = "                        // 生きているセルは、周りが2か3なら生き残って世代数+1。死んでいるセルは、周りがちょうど3なら誕生"
    SetJS = "                        const a = age[y][x];"
    SetJS = "                        row.push(a > 0 ? ((n === 2 || n === 3) ? Math.min(a + 1, 99) : 0) : (n === 3 ? 1 : 0));"
    SetJS = "                    }"
    SetJS = "                    next.push(row);"
    SetJS = "                }"
    SetJS = "                age = next;"
    SetJS = "            }"
    SetJS = "            const alive = age.reduce((sum, r) => sum + r.filter((v) => v > 0).length, 0);"
    SetJS = "            const calcMs = Date.now() - T1;"
    SetJS = ""
    SetJS = "            grid.values = age.map((r) => r.map((v) => (v > 0 ? v : """")));"
    SetJS = "            ws.getRange(""A1"").values = [[caption]];"
    SetJS = "            await context.sync();"
    SetJS = "            return ""OK: alive "" + alive + "" / calc "" + calcMs + ""ms / total "" + (Date.now() - T0) + ""ms"";"
    SetJS = "        });"
    SetJS = "    } catch (e) {"
    SetJS = "        return ""NG: "" + String((e && (e.code || e.message)) || e);"
    SetJS = "    }"
    SetJS = "}"
    JS_LifeGameStep = SetJS
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
