Attribute VB_Name = "Demo_OfficeJs_Showcase"
'***************************************************************************************************
'       exCDP_OfficeJs 拡張 - ショーケース(Office.js ダイレクト版) モジュール
'***************************************************************************************************
'* 機能　　：Demo_OfficeScript_Showcase を、Officeスクリプト(エディターに書き込んで実行)ではなく、
'            コードエディターのページの Office.js(Excel.run)を直接呼ぶ形に作り直したリメイク版です
'            ・Showcase_00 全部まとめて(録画用の通し)
'            ・Showcase_01 マンデルブロ集合のズーム
'            ・Showcase_02 ライフゲーム
'            ・Showcase_03 ドット絵バナー
'            ・Showcase_04 ブックの棚卸し(読み取り。デスクトップ版でも同期で待てる)
'            ・Showcase_05 ライブミラー(読み取り)
'            ・Showcase_99 後片付け
'---------------------------------------------------------------------------------------------------
'* 元の版との違い：
'            ・1回あたり、元の版は Web 版で約1.5秒・デスクトップ版で約4～5秒でしたが、この版は数百ミリ秒以下です。
'              そのぶんフレーム数を増やし、1フレームごとに FRAME_WAIT_SEC だけ待って Excel に描画させます
'            ・棚卸しは、デスクトップ版でも OnTime を使わず、同期で待てます
'            ・作るシート名は、元の版と重ならないように末尾に「Js」を付けています
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\OfficeScriptViaVBA\exCDP_OfficeJs.cls
'---------------------------------------------------------------------------------------------------
'* 注意事項：・「exCDP_OfficeJs.cls」と「Demo_OfficeScript.bas」(ConnectWebExcel / ConnectDesktopExcel を使用)も
'              取り込んでください
'            ・事前に、Excel でスクリプトを開き「コード エディター」を表示してください(エディターの中身は書き換えません)
'            ・既定は Web 版 Excel(Edge)で実行します。デスクトップ版で動かす場合は SHOWCASE_WEB を False にしてください
'            ・描画系はブックに「MandelbrotJs」「LifeGameJs」「BannerJs」シートを作って描きます。Showcase_99 で削除できます
'            ・読み取り系(04 / 05)は対象のブックを書き換えません。結果は、このマクロのブック(ThisWorkbook)の
'              「OfficeJsReport」「OfficeJsMirror」シートに書き出します
'            ・ループ中に止めたいときは、Ctrl+Break で VBA を中断してください
'***************************************************************************************************
Option Explicit

'接続先(True:Web 版 Excel / False:デスクトップ版 Excel)
Private Const SHOWCASE_WEB      As Boolean = True

'演出：実行するコードを、コードエディターにも表示する(実際の実行は Office.js。表示は結果を待たない)
Private Const SHOW_IN_EDITOR    As Boolean = True

'1フレームごとに待つ時間(秒)。Excel に描画させるための間
Private Const FRAME_WAIT_SEC    As Double = 0.05

'マンデルブロ集合
Private Const MANDEL_SHEET      As String = "MandelbrotJs"
Private Const MANDEL_W          As Long = 150
Private Const MANDEL_H          As Long = 100
Private Const MANDEL_SIZE       As Long = 4
Private Const MANDEL_CX         As Double = -0.743643887037151
Private Const MANDEL_CY         As Double = 0.13182590420533
Private Const MANDEL_ZOOM       As Double = 0.94        '1フレームごとの倍率(元の版より細かく寄る)

'ライフゲーム
Private Const LIFE_SHEET        As String = "LifeGameJs"
Private Const LIFE_W            As Long = 80
Private Const LIFE_H            As Long = 50
Private Const LIFE_SIZE         As Long = 8

'ドット絵バナー
Private Const BANNER_SHEET      As String = "BannerJs"
Private Const BANNER_SIZE       As Long = 12
Private Const BANNER_CHARS      As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 -><!.:?*"

'読み取り系(結果の書き出し先は、このマクロのブック)
Private Const REPORT_SHEET      As String = "OfficeJsReport"
Private Const INSPECT_LIMIT     As Long = 200000
Private Const MIRROR_SHEET      As String = "OfficeJsMirror"
Private Const MIRROR_MAX_ROWS   As Long = 40
Private Const MIRROR_MAX_COLS   As Long = 12
Private Const MIRROR_TOP        As Long = 4
Private Const MIRROR_WAIT_SEC   As Double = 0.5         '読み取りの間隔(秒)



'***************************************************************************************************
'                            ■■■ 初めに ■■■
'***************************************************************************************************
'* 機能　　：デバッグ経路を確保します
'***************************************************************************************************
Sub FirstFunction()
    Dim tmp As New exCDP_OfficeScript
    tmp.OpenDebugPortWebView2
End Sub



'***************************************************************************************************
'                         ■■■ CDPBrowser の用意(bas側の責務) ■■■
'***************************************************************************************************
'* 機能　　：デスクトップ版Excel：OfficeアドインのWebView2へ接続した`CDPBrowser`を返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：接続済みの`CDPBrowser`。接続できなければ Nothing
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebView2のデバッグポートは、環境変数`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`で開けるため、
'            Office側のWebView2が起動する「前」に`EnsureWebView2DebugPort`が呼ばれている必要があります。
'            接続できない場合は、一度実行した後にコードエディターを開き直して再実行してください
'***************************************************************************************************
Public Function ConnectDesktopExcel() As CDPBrowser
    Const USERDATA_SUBPATH As String = "\Microsoft\Office\16.0\Wef\webview2\2\1\EBWebView"   '%LOCALAPPDATA%配下
'    Const USERDATA_SUBPATH As String = "\Microsoft\Office\16.0\Wef\webview2\561d27ca26ad86cf_LiveId\2\1\EBWebView"   '%LOCALAPPDATA%配下


    On Error GoTo Failed

    '1. OfficeアドインのWebView2のユーザーデータフォルダ
    Dim UserDataDir As String
    UserDataDir = Environ("LOCALAPPDATA") & USERDATA_SUBPATH

    '2. WebView2のデバッグポートを開け、`DevToolsActivePort`から接続
    BrowserCapabilities.Set_UserDataDirName = UserDataDir
    Dim WebSocketCDP As CDPCoreViaWebSocket: Set WebSocketCDP = New CDPCoreViaWebSocket
    If Not WebSocketCDP.AutoConnectDevToolsActivePort(BrowserCapabilities) Then
        Debug.Print "[ConnectDesktopExcel] `DevToolsActivePort`から接続できませんでした。コードエディターを開き直して再実行してください"
        Exit Function
    End If

    '3. WebSocketモードとして設定
    ReattachOptions.Set_WebSocketMode(WebSocketCDP) = UserDataDir

    '4. 再接続情報を`reattach`メソッドに渡す
    Set ConnectDesktopExcel = New CDPBrowser
    ConnectDesktopExcel.reattach ReattachOptions
    Exit Function

Failed:
    Debug.Print "[ConnectDesktopExcel] 接続に失敗しました。Error " & Err.Number & ": " & Err.Description
End Function

'***************************************************************************************************
'* 機能　　：Web版Excel：Excelを開いているEdgeへ接続した`CDPBrowser`を返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：接続済みの`CDPBrowser`。接続できなければ Nothing
'---------------------------------------------------------------------------------------------------
'* 注意事項：・事前にEdgeのリモートデバッグを有効にしてください(`edge://inspect/#remote-debugging`で許可、等)
'            ・`DevToolsActivePort`に書かれたブラウザ用WebSocketで繋ぐため、`/json`系が無効でも繋がります
'            ・別のユーザーデータフォルダで起動したEdgeの場合は、UserDataDirを書き換えてください
'***************************************************************************************************
Public Function ConnectWebExcel() As CDPBrowser
    Const MainUserData As String = "User Data"


    On Error GoTo Failed

    '1. ワークシートの設定に委ねた、今目の前のブラウザへ接続させる
    Dim WebSocketCDP As CDPCoreViaWebSocket: Set WebSocketCDP = New CDPCoreViaWebSocket
    BrowserCapabilities.Set_UserDataDirName = MainUserData
    If Not WebSocketCDP.AutoConnectDevToolsActivePort(BrowserCapabilities) Then
        Debug.Print "[ConnectWebExcel] `DevToolsActivePort`から接続できませんでした。Edgeのリモートデバッグを有効にしてください"
        Exit Function
    End If

    '2. WebSocketモードとして設定
    ReattachOptions.Set_WebSocketMode(WebSocketCDP) = MainUserData

    '3. 再接続情報を`reattach`メソッドに渡す
    Set ConnectWebExcel = New CDPBrowser
    ConnectWebExcel.reattach ReattachOptions
    Exit Function

Failed:
    Debug.Print "[ConnectWebExcel] 接続に失敗しました。Error " & Err.Number & ": " & Err.Description
End Function



'***************************************************************************************************
'                         ■■■ Showcase 00：全部まとめて(録画用) ■■■
'***************************************************************************************************
Sub Showcase_00_全部まとめて()
    Dim Browser As CDPBrowser, ojs As exCDP_OfficeJs
    If Not OpenShowcase(Browser, ojs) Then Exit Sub

    If PlayBanner(ojs, "VBA -> OFFICE.JS", 40) Then
        If PlayMandelbrot(ojs, 120) Then
            If PlayLife(ojs, 240) Then
                PlayBanner ojs, "DIRECT!", 40
            End If
        End If
    End If

    CloseShowcase Browser, ojs
End Sub



'***************************************************************************************************
'                         ■■■ Showcase 01：マンデルブロ集合のズーム ■■■
'***************************************************************************************************
Sub Showcase_01_マンデルブロズーム()
    Dim Browser As CDPBrowser, ojs As exCDP_OfficeJs
    If Not OpenShowcase(Browser, ojs) Then Exit Sub

    PlayMandelbrot ojs, 120

    CloseShowcase Browser, ojs
End Sub



'***************************************************************************************************
'                             ■■■ Showcase 02：ライフゲーム ■■■
'***************************************************************************************************
Sub Showcase_02_ライフゲーム()
    Dim Browser As CDPBrowser, ojs As exCDP_OfficeJs
    If Not OpenShowcase(Browser, ojs) Then Exit Sub

    PlayLife ojs, 240

    CloseShowcase Browser, ojs
End Sub



'***************************************************************************************************
'                            ■■■ Showcase 03：ドット絵バナー ■■■
'***************************************************************************************************
'* 注意事項：使える文字は英大文字・数字・空白と「-><!.:?*」です。英小文字は大文字に、それ以外は「?」になります
'***************************************************************************************************
Sub Showcase_03_ドット絵バナー()
    Dim Browser As CDPBrowser, ojs As exCDP_OfficeJs
    If Not OpenShowcase(Browser, ojs) Then Exit Sub

    PlayBanner ojs, "VBA -> OFFICE.JS", 60

    CloseShowcase Browser, ojs
End Sub



'***************************************************************************************************
'                         ■■■ Showcase 04：ブックの棚卸し(読み取り) ■■■
'***************************************************************************************************
'* 機能　　：対象のブックの全シートを Office.js で調べ、結果を手元のブックにレポートとグラフで書き出します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：元の版(Officeスクリプト)は、デスクトップ版で VBA が同期で待つと結果が返らなかったため OnTime で
'            回収していました。この版は、デスクトップ版でもそのまま同期で待てます
'***************************************************************************************************
Sub Showcase_04_ブック棚卸し()
    Dim Browser As CDPBrowser, ojs As exCDP_OfficeJs
    If Not OpenShowcase(Browser, ojs) Then Exit Sub

    If ojs.Run(Replace(JS_Inspect(), "__LIMIT__", CStr(INSPECT_LIMIT))) Then
        Debug.Print "[Showcase] Inspect : " & Format$(ojs.LastElapsedSec, "0.00") & " 秒"
        WriteInspectReport ojs.LastResult
    Else
        Debug.Print "[Showcase] Inspect : " & ojs.LastError
    End If

    CloseShowcase Browser, ojs
End Sub



'***************************************************************************************************
'                         ■■■ Showcase 05：ライブミラー(読み取り) ■■■
'***************************************************************************************************
'* 機能　　：対象のブックでアクティブなシートを繰り返し読み取り、手元のブックの「OfficeJsMirror」シートに写します。
'            前回から変わったセルは黄色に光ります
'---------------------------------------------------------------------------------------------------
'* 注意事項：デスクトップ版(SHOWCASE_WEB = False)では、コードエディターを開いているブックを読みます。
'            このマクロのブックでエディターを開いていると、OfficeJsMirror シート自身を読むことになるので注意してください
'***************************************************************************************************
Sub Showcase_05_ライブミラー()
    Const Polls As Long = 120


    Dim Browser As CDPBrowser, ojs As exCDP_OfficeJs
    If Not OpenShowcase(Browser, ojs) Then Exit Sub

    '1. 写す先のシートを用意
    Dim ws As Worksheet
    Set ws = PrepareLocalSheet(MIRROR_SHEET)
    ws.Range("A1").Value = "Excel → 手元の Excel (ライブミラー / Office.js ダイレクト)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Activate

    '2. 読み取り → 書き出し を繰り返す
    Dim s As String, p As Long, Prev As Variant, LastAddress As String
    s = JS_MirrorRead()
    s = Replace(s, "__MAX_ROWS__", CStr(MIRROR_MAX_ROWS))
    s = Replace(s, "__MAX_COLS__", CStr(MIRROR_MAX_COLS))
    For p = 1 To Polls
        If Not ojs.Run(s) Then Debug.Print "[Showcase] Mirror : " & ojs.LastError: Exit For
        WriteMirror ws, ojs.LastResult, p, Polls, Prev, LastAddress
        Pause MIRROR_WAIT_SEC
    Next

    CloseShowcase Browser, ojs
End Sub



'***************************************************************************************************
'                              ■■■ Showcase 99：後片付け ■■■
'***************************************************************************************************
Sub Showcase_99_後片付け()
    Dim Browser As CDPBrowser, ojs As exCDP_OfficeJs
    If Not OpenShowcase(Browser, ojs) Then Exit Sub

    Dim s As String
    s = Replace(JS_Cleanup(), "__SHEETS__", """" & MANDEL_SHEET & """, """ & LIFE_SHEET & """, """ & BANNER_SHEET & """")
    RunStep ojs, s, "Cleanup"
    Debug.Print "[Showcase] " & ojs.LastResult

    CloseShowcase Browser, ojs
End Sub



'***************************************************************************************************
'                                  ■■■ 再生処理 ■■■
'***************************************************************************************************
Private Function PlayMandelbrot(ojs As exCDP_OfficeJs, Frames As Long) As Boolean
    Dim s As String
    s = JS_MandelSetup()
    s = Replace(s, "__W__", CStr(MANDEL_W))
    s = Replace(s, "__H__", CStr(MANDEL_H))
    s = Replace(s, "__SIZE__", CStr(MANDEL_SIZE))
    s = Replace(s, "__MID__", "85")
    s = Replace(s, "__SHEET__", MANDEL_SHEET)
    If Not RunStep(ojs, s, "Mandelbrot setup") Then Exit Function

    Dim f As Long, Scale_ As Double
    Scale_ = 3#
    For f = 0 To Frames - 1
        s = JS_MandelFrame()
        s = Replace(s, "__W__", CStr(MANDEL_W))
        s = Replace(s, "__H__", CStr(MANDEL_H))
        s = Replace(s, "__MAX__", CStr(120 + f * 4))
        s = Replace(s, "__CX__", Num(MANDEL_CX))
        s = Replace(s, "__CY__", Num(MANDEL_CY))
        s = Replace(s, "__SCALE__", Num(Scale_))
        s = Replace(s, "__SHEET__", MANDEL_SHEET)
        If Not RunStep(ojs, s, "Mandelbrot frame " & f + 1 & "/" & Frames) Then Exit Function
        Scale_ = Scale_ * MANDEL_ZOOM
        Pause FRAME_WAIT_SEC
    Next

    PlayMandelbrot = True
End Function

Private Function PlayLife(ojs As exCDP_OfficeJs, Generations As Long) As Boolean
    Dim s As String
    s = JS_LifeSetup()
    s = Replace(s, "__W__", CStr(LIFE_W))
    s = Replace(s, "__H__", CStr(LIFE_H))
    s = Replace(s, "__SIZE__", CStr(LIFE_SIZE))
    s = Replace(s, "__SHEET__", LIFE_SHEET)
    If Not RunStep(ojs, s, "Life setup") Then Exit Function

    Dim g As Long
    For g = 0 To Generations - 1
        s = JS_LifeStep()
        s = Replace(s, "__W__", CStr(LIFE_W))
        s = Replace(s, "__H__", CStr(LIFE_H))
        s = Replace(s, "__STEPS__", "1")
        s = Replace(s, "__GEN__", CStr(g))
        s = Replace(s, "__SHEET__", LIFE_SHEET)
        If Not RunStep(ojs, s, "Life gen " & g + 1 & "/" & Generations) Then Exit Function
        Pause FRAME_WAIT_SEC
    Next

    PlayLife = True
End Function

Private Function PlayBanner(ojs As exCDP_OfficeJs, ByVal Text As String, Frames As Long) As Boolean
    Text = SanitizeBannerText(Text)

    Dim Size As Long
    Size = BANNER_SIZE
    If Len(Text) > 10 Then Size = Int(BANNER_SIZE * 10 / Len(Text))
    If Size < 4 Then Size = 4

    Dim s As String
    s = JS_BannerSetup()
    s = Replace(s, "__TEXT__", Text)
    s = Replace(s, "__SIZE__", CStr(Size))
    s = Replace(s, "__SHEET__", BANNER_SHEET)
    If Not RunStep(ojs, s, "Banner setup") Then Exit Function

    Dim f As Long
    For f = 1 To Frames
        s = JS_BannerFrame()
        s = Replace(s, "__FRAME__", CStr(f))
        s = Replace(s, "__LEN__", CStr(Len(Text)))
        s = Replace(s, "__SHEET__", BANNER_SHEET)
        If Not RunStep(ojs, s, "Banner frame " & f & "/" & Frames) Then Exit Function
        Pause FRAME_WAIT_SEC
    Next

    PlayBanner = True
End Function



'***************************************************************************************************
'                                  ■■■ 共通処理 ■■■
'***************************************************************************************************
Private Function OpenShowcase(Browser As CDPBrowser, ojs As exCDP_OfficeJs) As Boolean
    If SHOWCASE_WEB Then
        Set Browser = ConnectWebExcel()
    Else
        Set Browser = ConnectDesktopExcel()
    End If
    If Browser Is Nothing Then
        MsgBox "接続できませんでした。イミディエイトウィンドウを確認してください。", vbCritical
        Exit Function
    End If

    Set ojs = New exCDP_OfficeJs
    If Not ojs.Init(Browser, WebMode:=SHOWCASE_WEB) Then
        CloseShowcase Browser, ojs
        MsgBox "Office.js が使えるページ(コード エディター)が見つかりませんでした。イミディエイトウィンドウを確認してください。", vbCritical
        Exit Function
    End If
    ojs.ShowInEditor = SHOW_IN_EDITOR

    OpenShowcase = True
End Function

Private Sub CloseShowcase(Browser As CDPBrowser, ojs As exCDP_OfficeJs)
    If Not ojs Is Nothing Then ojs.Dispose
    Set ojs = Nothing
    If Not Browser Is Nothing Then Browser.DisconnectCDP
    Set Browser = Nothing
End Sub

'* Excel.run を1回実行します。失敗したら理由を出して False を返します
Private Function RunStep(ojs As exCDP_OfficeJs, Body As String, Label As String) As Boolean
    ojs.EditorCaption = Label      '演出：エディターに表示する見出し
    If ojs.Run(Body) Then
        Debug.Print "[Showcase] " & Label & " : " & Format$(ojs.LastElapsedSec * 1000, "0") & " ms"
        RunStep = True
    Else
        Debug.Print "[Showcase] " & Label & " : " & ojs.LastError
    End If
End Function

'* 指定秒だけ待ちます(待つ間も DoEvents で Excel に描画させる)
Private Sub Pause(Seconds As Double)
    If Seconds > 0 Then CDPHelpers.Sleep Seconds Else DoEvents
End Sub

Private Function Num(Value As Double) As String
    Num = Trim$(Str$(Value))
End Function

Private Function SanitizeBannerText(ByVal Text As String) As String
    Dim i As Long, c As String
    Text = UCase$(Text)
    For i = 1 To Len(Text)
        c = Mid$(Text, i, 1)
        If InStr(1, BANNER_CHARS, c, vbBinaryCompare) = 0 Then c = "?"
        SanitizeBannerText = SanitizeBannerText & c
    Next
    If Len(SanitizeBannerText) = 0 Then SanitizeBannerText = "VBA"
End Function

'***************************************************************************************************
'                          ■■■ 読み取り結果の書き出し(手元のブック) ■■■
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
'* 機能　　：棚卸しの JSON を、手元のブックの「OfficeJsReport」シートにテーブルとグラフで書き出します
'***************************************************************************************************
Private Sub WriteInspectReport(Json As String)
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
    Set ws = PrepareLocalSheet(REPORT_SHEET)
    ws.Range("A1").Value = "ブックの棚卸し(Office.js)：" & Root.StringKey("workbook")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "取得日時：" & Format$(Now, "yyyy/mm/dd hh:nn:ss") & "　名前の定義：" & Root.NumberKey("names") & " 件"
    ws.Range("A" & FirstRow).Resize(1, 10).Value = Array("シート名", "使用範囲", "行数", "列数", "セル数", "数式", "エラー", "エラーの場所(先頭5件)", "テーブル", "グラフ")

    '3. シートごとの行
    Dim r As Long, Tok As Long, Node As BiDiCDPJConv, ErrCells As BiDiCDPJConv, i As Long, List As String
    r = FirstRow
    Tok = SheetList.FirstChildToken
    Do While Tok > 0
        r = r + 1
        Set Node = SheetList.NodeFromToken(Tok)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).NumberFormat = "@"
        ws.Cells(r, 8).NumberFormat = "@"
        ws.Cells(r, 1).Value = Node.StringKey("name")
        ws.Cells(r, 2).Value = Node.StringKey("used")
        ws.Cells(r, 3).Value = Node.NumberKey("rows")
        ws.Cells(r, 4).Value = Node.NumberKey("cols")
        ws.Cells(r, 5).Value = Node.NumberKey("rows") * Node.NumberKey("cols")
        If Node.BoolKey("scanned") Then
            ws.Cells(r, 6).Value = Node.NumberKey("formulas")
            ws.Cells(r, 7).Value = Node.NumberKey("errors")
        Else
            ws.Cells(r, 6).Value = "-"
            ws.Cells(r, 7).Value = "-"
        End If

        List = ""
        Set ErrCells = Node.NodeKey("errorCells")
        If Not ErrCells Is Nothing Then
            For i = 0 To ErrCells.Count - 1
                If Len(List) Then List = List & ", "
                List = List & ErrCells.StringAt(i)
            Next
        End If
        ws.Cells(r, 8).Value = List
        ws.Cells(r, 9).Value = Node.NumberKey("tables")
        ws.Cells(r, 10).Value = Node.NumberKey("charts")

        Debug.Print "[Showcase] " & Node.StringKey("name") & " : " & Node.StringKey("used") & " / 数式 " & ws.Cells(r, 6).Value & " / エラー " & ws.Cells(r, 7).Value & IIf(Len(List), " (" & List & ")", "")
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
    Set ChartObj = ws.ChartObjects.Add(Left:=ws.Cells(FirstRow, 12).Left, Top:=ws.Cells(FirstRow, 12).Top, Width:=420, Height:=260)
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

    ws.Activate
    ws.Range("A1").Select
End Sub

'***************************************************************************************************
'* 機能　　：ライブミラーの JSON を、手元のブックの「OfficeJsMirror」シートに写します。前回から変わったセルは黄色にします
'---------------------------------------------------------------------------------------------------
'* 引数　　：Prev / LastAddress   (ByRef) 前回の内容と範囲。呼び出し側で保持して、次の回に渡す
'***************************************************************************************************
Private Sub WriteMirror(ws As Worksheet, Json As String, Poll As Long, Polls As Long, Prev As Variant, LastAddress As String)
    If Len(Json) = 0 Then Exit Sub

    '1. JSON を読む
    Dim Root As BiDiCDPJConv, Texts As BiDiCDPJConv, Address As String
    Set Root = BiDiCDPJConv.Parse(Json)
    Address = Root.StringKey("address")
    ws.Range("A2").Value = "読み取り " & Poll & "/" & Polls & "　" & Format$(Now, "hh:nn:ss") & "　" & Root.StringKey("sheet") & "!" & Address

    '2. 前回の光を消す
    Dim Area As Range
    Set Area = ws.Cells(MIRROR_TOP, 1).Resize(MIRROR_MAX_ROWS, MIRROR_MAX_COLS)
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
    Set Target = ws.Cells(MIRROR_TOP, 1).Resize(RowsN, ColsN)
    Target.NumberFormat = "@"
    Target.Value = Cur

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
'                     ■■■ Office.js(Excel.run の中身)のひな形 ■■■
'***************************************************************************************************
'* 機能　　：「__名前__」の部分を、VBA 側で Replace して使います。`context` は Excel.RequestContext です
'***************************************************************************************************
'* マンデルブロ集合：シートの準備
Private Function JS_MandelSetup() As String
    Dim s As String
    s = s & "        const W = __W__, H = __H__, SIZE = __SIZE__, NAME = ""__SHEET__"";" & vbLf
    s = s & "        const wb = context.workbook;" & vbLf
    s = s & "        let ws = wb.worksheets.getItemOrNullObject(NAME);" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        if (ws.isNullObject) { ws = wb.worksheets.add(NAME); }" & vbLf
    s = s & "        ws.activate();" & vbLf
    s = s & "        ws.getRange().clear(Excel.ClearApplyTo.all);" & vbLf
    s = s & "        ws.showGridlines = false;" & vbLf
    s = s & "        const title = ws.getRangeByIndexes(0, 0, 1, W);" & vbLf
    s = s & "        title.merge(false);" & vbLf
    s = s & "        title.getCell(0, 0).values = [[""VBA -> CDP -> Office.js : Mandelbrot""]];" & vbLf
    s = s & "        title.format.font.bold = true;" & vbLf
    s = s & "        title.format.font.size = 14;" & vbLf
    s = s & "        const grid = ws.getRangeByIndexes(1, 0, H, W);" & vbLf
    s = s & "        grid.format.columnWidth = SIZE;" & vbLf
    s = s & "        grid.format.rowHeight = SIZE;" & vbLf
    s = s & "        grid.numberFormat = "";;;"";" & vbLf
    s = s & "        grid.format.fill.color = ""#000000"";" & vbLf
    s = s & "        const cf = grid.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);" & vbLf
    s = s & "        cf.colorScale.criteria = {" & vbLf
    s = s & "            minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: ""#020B3A"" }," & vbLf
    s = s & "            midpoint: { formula: ""__MID__"", type: Excel.ConditionalFormatColorCriterionType.percentile, color: ""#FF9A1F"" }," & vbLf
    s = s & "            maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: ""#FFFFF0"" }" & vbLf
    s = s & "        };" & vbLf
    s = s & "        ws.getRange(""A1"").select();" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        return ""OK"";" & vbLf
    JS_MandelSetup = s
End Function

'* マンデルブロ集合：1フレーム
Private Function JS_MandelFrame() As String
    Dim s As String
    s = s & "        const W = __W__, H = __H__, MAX = __MAX__;" & vbLf
    s = s & "        const CX = __CX__, CY = __CY__, SCALE = __SCALE__;" & vbLf
    s = s & "        const ws = context.workbook.worksheets.getItem(""__SHEET__"");" & vbLf
    s = s & "        const rows = [];" & vbLf
    s = s & "        for (let y = 0; y < H; y++) {" & vbLf
    s = s & "            const row = [];" & vbLf
    s = s & "            const ci = CY + (y / H - 0.5) * SCALE;" & vbLf
    s = s & "            for (let x = 0; x < W; x++) {" & vbLf
    s = s & "                const cr = CX + (x / W - 0.5) * SCALE * (W / H);" & vbLf
    s = s & "                let zr = 0, zi = 0, i = 0;" & vbLf
    s = s & "                while (zr * zr + zi * zi < 16 && i < MAX) {" & vbLf
    s = s & "                    const t = zr * zr - zi * zi + cr;" & vbLf
    s = s & "                    zi = 2 * zr * zi + ci;" & vbLf
    s = s & "                    zr = t;" & vbLf
    s = s & "                    i++;" & vbLf
    s = s & "                }" & vbLf
    s = s & "                row.push(i >= MAX ? """" : Math.round(Math.sqrt(i + 1 - Math.log(Math.log(Math.sqrt(zr * zr + zi * zi))) / Math.LN2) * 1000) / 1000);" & vbLf
    s = s & "            }" & vbLf
    s = s & "            rows.push(row);" & vbLf
    s = s & "        }" & vbLf
    s = s & "        ws.getRangeByIndexes(1, 0, H, W).values = rows;" & vbLf
    s = s & "        ws.getRange(""A1"").values = [[""VBA -> CDP -> Office.js : Mandelbrot  x"" + Math.round(3 / SCALE)]];" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        return ""OK"";" & vbLf
    JS_MandelFrame = s
End Function

'* ライフゲーム：シートの準備(グライダー銃とパルサー)
Private Function JS_LifeSetup() As String
    Dim s As String
    s = s & "        const W = __W__, H = __H__, SIZE = __SIZE__, NAME = ""__SHEET__"";" & vbLf
    s = s & "        const wb = context.workbook;" & vbLf
    s = s & "        let ws = wb.worksheets.getItemOrNullObject(NAME);" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        if (ws.isNullObject) { ws = wb.worksheets.add(NAME); }" & vbLf
    s = s & "        ws.activate();" & vbLf
    s = s & "        ws.getRange().clear(Excel.ClearApplyTo.all);" & vbLf
    s = s & "        ws.showGridlines = false;" & vbLf
    s = s & "        const title = ws.getRangeByIndexes(0, 0, 1, W);" & vbLf
    s = s & "        title.merge(false);" & vbLf
    s = s & "        title.getCell(0, 0).values = [[""VBA -> CDP -> Office.js : Game of Life  gen 0""]];" & vbLf
    s = s & "        title.format.font.bold = true;" & vbLf
    s = s & "        title.format.font.size = 14;" & vbLf
    s = s & "        const grid = ws.getRangeByIndexes(1, 0, H, W);" & vbLf
    s = s & "        grid.format.columnWidth = SIZE;" & vbLf
    s = s & "        grid.format.rowHeight = SIZE;" & vbLf
    s = s & "        grid.numberFormat = "";;;"";" & vbLf
    s = s & "        grid.format.fill.color = ""#05050F"";" & vbLf
    s = s & "        const cf = grid.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);" & vbLf
    s = s & "        cf.colorScale.criteria = {" & vbLf
    s = s & "            minimum: { formula: ""1"", type: Excel.ConditionalFormatColorCriterionType.number, color: ""#39FF14"" }," & vbLf
    s = s & "            midpoint: { formula: ""6"", type: Excel.ConditionalFormatColorCriterionType.number, color: ""#00B4FF"" }," & vbLf
    s = s & "            maximum: { formula: ""30"", type: Excel.ConditionalFormatColorCriterionType.number, color: ""#B000FF"" }" & vbLf
    s = s & "        };" & vbLf
    s = s & "        const cells = [];" & vbLf
    s = s & "        for (let y = 0; y < H; y++) { const row = []; for (let x = 0; x < W; x++) { row.push(""""); } cells.push(row); }" & vbLf
    s = s & "        const put = (pattern, top, left) => {" & vbLf
    s = s & "            for (let y = 0; y < pattern.length; y++) {" & vbLf
    s = s & "                for (let x = 0; x < pattern[y].length; x++) {" & vbLf
    s = s & "                    if (pattern[y][x] === ""O"" && top + y < H && left + x < W) { cells[top + y][left + x] = 1; }" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        };" & vbLf
    s = s & "        put([" & vbLf
    s = s & "            ""........................O...........""," & vbLf
    s = s & "            ""......................O.O...........""," & vbLf
    s = s & "            ""............OO......OO............OO""," & vbLf
    s = s & "            ""...........O...O....OO............OO""," & vbLf
    s = s & "            ""OO........O.....O...OO..............""," & vbLf
    s = s & "            ""OO........O...O.OO....O.O...........""," & vbLf
    s = s & "            ""..........O.....O.......O...........""," & vbLf
    s = s & "            ""...........O...O....................""," & vbLf
    s = s & "            ""............OO......................""" & vbLf
    s = s & "        ], 2, 2);" & vbLf
    s = s & "        put([" & vbLf
    s = s & "            ""..OOO...OOO..""," & vbLf
    s = s & "            "".............""," & vbLf
    s = s & "            ""O....O.O....O""," & vbLf
    s = s & "            ""O....O.O....O""," & vbLf
    s = s & "            ""O....O.O....O""," & vbLf
    s = s & "            ""..OOO...OOO..""," & vbLf
    s = s & "            "".............""," & vbLf
    s = s & "            ""..OOO...OOO..""," & vbLf
    s = s & "            ""O....O.O....O""," & vbLf
    s = s & "            ""O....O.O....O""," & vbLf
    s = s & "            ""O....O.O....O""," & vbLf
    s = s & "            "".............""," & vbLf
    s = s & "            ""..OOO...OOO..""" & vbLf
    s = s & "        ], 3, W - 16);" & vbLf
    s = s & "        grid.values = cells;" & vbLf
    s = s & "        ws.getRange(""A1"").select();" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        return ""OK"";" & vbLf
    JS_LifeSetup = s
End Function

'* ライフゲーム：STEPS 世代進める
Private Function JS_LifeStep() As String
    Dim s As String
    s = s & "        const W = __W__, H = __H__, STEPS = __STEPS__, GEN = __GEN__;" & vbLf
    s = s & "        const ws = context.workbook.worksheets.getItem(""__SHEET__"");" & vbLf
    s = s & "        const grid = ws.getRangeByIndexes(1, 0, H, W);" & vbLf
    s = s & "        grid.load(""values"");" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        let age = grid.values.map((r) => r.map((v) => (typeof v === ""number"" && v > 0) ? v : 0));" & vbLf
    s = s & "        for (let s = 0; s < STEPS; s++) {" & vbLf
    s = s & "            const next = [];" & vbLf
    s = s & "            for (let y = 0; y < H; y++) {" & vbLf
    s = s & "                const row = [];" & vbLf
    s = s & "                for (let x = 0; x < W; x++) {" & vbLf
    s = s & "                    let n = 0;" & vbLf
    s = s & "                    for (let dy = -1; dy <= 1; dy++) {" & vbLf
    s = s & "                        for (let dx = -1; dx <= 1; dx++) {" & vbLf
    s = s & "                            if (dx === 0 && dy === 0) { continue; }" & vbLf
    s = s & "                            const yy = y + dy, xx = x + dx;" & vbLf
    s = s & "                            if (yy >= 0 && yy < H && xx >= 0 && xx < W && age[yy][xx] > 0) { n++; }" & vbLf
    s = s & "                        }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                    const a = age[y][x];" & vbLf
    s = s & "                    row.push(a > 0 ? ((n === 2 || n === 3) ? Math.min(a + 1, 99) : 0) : (n === 3 ? 1 : 0));" & vbLf
    s = s & "                }" & vbLf
    s = s & "                next.push(row);" & vbLf
    s = s & "            }" & vbLf
    s = s & "            age = next;" & vbLf
    s = s & "        }" & vbLf
    s = s & "        grid.values = age.map((r) => r.map((v) => (v > 0 ? v : """")));" & vbLf
    s = s & "        ws.getRange(""A1"").values = [[""VBA -> CDP -> Office.js : Game of Life  gen "" + (GEN + STEPS)]];" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        return ""OK"";" & vbLf
    JS_LifeStep = s
End Function

'* ドット絵バナー：シートの準備(5x7ドットのフォント)
Private Function JS_BannerSetup() As String
    Dim s As String
    s = s & "        const TEXT = ""__TEXT__"", SIZE = __SIZE__, PAD = 2, NAME = ""__SHEET__"";" & vbLf
    s = s & "        const FONT = {" & vbLf
    s = s & "            ""0"": ""01110100011001110101110011000101110"", ""1"": ""00100011000010000100001000010001110"", ""2"": ""01110100010000100010001000100011111""," & vbLf
    s = s & "            ""3"": ""11111000100010000010000011000101110"", ""4"": ""00010001100101010010111110001000010"", ""5"": ""11111100001111000001000011000101110""," & vbLf
    s = s & "            ""6"": ""00110010001000011110100011000101110"", ""7"": ""11111000010001000100010000100001000"", ""8"": ""01110100011000101110100011000101110""," & vbLf
    s = s & "            ""9"": ""01110100011000101111000010001001100"", ""A"": ""01110100011000111111100011000110001"", ""B"": ""11110100011000111110100011000111110""," & vbLf
    s = s & "            ""C"": ""01110100011000010000100001000101110"", ""D"": ""11110100011000110001100011000111110"", ""E"": ""11111100001000011110100001000011111""," & vbLf
    s = s & "            ""F"": ""11111100001000011110100001000010000"", ""G"": ""01110100011000010111100011000101111"", ""H"": ""10001100011000111111100011000110001""," & vbLf
    s = s & "            ""I"": ""01110001000010000100001000010001110"", ""J"": ""00111000100001000010000101001001100"", ""K"": ""10001100101010011000101001001010001""," & vbLf
    s = s & "            ""L"": ""10000100001000010000100001000011111"", ""M"": ""10001110111010110101100011000110001"", ""N"": ""10001100011100110101100111000110001""," & vbLf
    s = s & "            ""O"": ""01110100011000110001100011000101110"", ""P"": ""11110100011000111110100001000010000"", ""Q"": ""01110100011000110001101011001001101""," & vbLf
    s = s & "            ""R"": ""11110100011000111110101001001010001"", ""S"": ""01111100001000001110000010000111110"", ""T"": ""11111001000010000100001000010000100""," & vbLf
    s = s & "            ""U"": ""10001100011000110001100011000101110"", ""V"": ""10001100011000110001100010101000100"", ""W"": ""10001100011000110101101011010101010""," & vbLf
    s = s & "            ""X"": ""10001100010101000100010101000110001"", ""Y"": ""10001100011000101010001000010000100"", ""Z"": ""11111000010001000100010001000011111""," & vbLf
    s = s & "            "" "": ""00000000000000000000000000000000000"", ""-"": ""00000000000000011111000000000000000"", "">"": ""01000001000001000001000100010001000""," & vbLf
    s = s & "            ""<"": ""00010001000100010000010000010000010"", ""!"": ""00100001000010000100001000000000100"", ""."": ""00000000000000000000000000110001100""," & vbLf
    s = s & "            "":"": ""00000011000110000000011000110000000"", ""?"": ""01110100010000100010001000000000100"", ""*"": ""00000100010101000100010101000100000""" & vbLf
    s = s & "        };" & vbLf
    s = s & "        const W = TEXT.length * 6 - 1 + PAD * 2, H = 7 + PAD * 2;" & vbLf
    s = s & "        const wb = context.workbook;" & vbLf
    s = s & "        let ws = wb.worksheets.getItemOrNullObject(NAME);" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        if (ws.isNullObject) { ws = wb.worksheets.add(NAME); }" & vbLf
    s = s & "        ws.activate();" & vbLf
    s = s & "        ws.getRange().clear(Excel.ClearApplyTo.all);" & vbLf
    s = s & "        ws.showGridlines = false;" & vbLf
    s = s & "        const area = ws.getRangeByIndexes(0, 0, H, W);" & vbLf
    s = s & "        area.format.columnWidth = SIZE;" & vbLf
    s = s & "        area.format.rowHeight = SIZE;" & vbLf
    s = s & "        area.numberFormat = "";;;"";" & vbLf
    s = s & "        area.format.fill.color = ""#000000"";" & vbLf
    s = s & "        const cf = area.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);" & vbLf
    s = s & "        cf.colorScale.criteria = {" & vbLf
    s = s & "            minimum: { formula: ""0"", type: Excel.ConditionalFormatColorCriterionType.number, color: ""#FF2E88"" }," & vbLf
    s = s & "            midpoint: { formula: ""50"", type: Excel.ConditionalFormatColorCriterionType.number, color: ""#00E5FF"" }," & vbLf
    s = s & "            maximum: { formula: ""100"", type: Excel.ConditionalFormatColorCriterionType.number, color: ""#FFE600"" }" & vbLf
    s = s & "        };" & vbLf
    s = s & "        const cells = [];" & vbLf
    s = s & "        for (let y = 0; y < H; y++) { const row = []; for (let x = 0; x < W; x++) { row.push(""""); } cells.push(row); }" & vbLf
    s = s & "        for (let c = 0; c < TEXT.length; c++) {" & vbLf
    s = s & "            const glyph = FONT[TEXT[c]] || FONT[""?""];" & vbLf
    s = s & "            for (let y = 0; y < 7; y++) {" & vbLf
    s = s & "                for (let x = 0; x < 5; x++) {" & vbLf
    s = s & "                    if (glyph[y * 5 + x] === ""1"") {" & vbLf
    s = s & "                        const gx = PAD + c * 6 + x, gy = PAD + y;" & vbLf
    s = s & "                        const p = (gx * 6 + gy * 3) % 200;" & vbLf
    s = s & "                        cells[gy][gx] = p <= 100 ? p : 200 - p;" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        area.values = cells;" & vbLf
    s = s & "        const caption = ws.getRangeByIndexes(H, 0, 1, W);" & vbLf
    s = s & "        caption.merge(false);" & vbLf
    s = s & "        caption.getCell(0, 0).values = [[""Drawn by Office.js, driven from VBA via CDP""]];" & vbLf
    s = s & "        caption.format.rowHeight = 20;" & vbLf
    s = s & "        caption.format.horizontalAlignment = Excel.HorizontalAlignment.center;" & vbLf
    s = s & "        caption.format.font.color = ""#888888"";" & vbLf
    s = s & "        ws.getRange(""A1"").select();" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        return ""OK"";" & vbLf
    JS_BannerSetup = s
End Function

'* ドット絵バナー：グラデーションを1フレーム進める
Private Function JS_BannerFrame() As String
    Dim s As String
    s = s & "        const FRAME = __FRAME__, LEN = __LEN__, PAD = 2;" & vbLf
    s = s & "        const W = LEN * 6 - 1 + PAD * 2, H = 7 + PAD * 2;" & vbLf
    s = s & "        const area = context.workbook.worksheets.getItem(""__SHEET__"").getRangeByIndexes(0, 0, H, W);" & vbLf
    s = s & "        area.load(""values"");" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        area.values = area.values.map((r, y) => r.map((v, x) => {" & vbLf
    s = s & "            if (typeof v !== ""number"") { return """"; }" & vbLf
    s = s & "            const p = (x * 6 + y * 3 + FRAME * 12) % 200;" & vbLf
    s = s & "            return p <= 100 ? p : 200 - p;" & vbLf
    s = s & "        }));" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        return ""OK"";" & vbLf
    JS_BannerFrame = s
End Function

'* ブックの棚卸し：全シートを調べて JSON で返す
Private Function JS_Inspect() As String
    Dim s As String
    s = s & "        const LIMIT = __LIMIT__, MAX_ERR = 5, T0 = Date.now();" & vbLf
    s = s & "        const col = (n) => { let s = """"; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };" & vbLf
    s = s & "        const wb = context.workbook; wb.load(""name"");" & vbLf
    s = s & "        const sheets = wb.worksheets; sheets.load(""items/name"");" & vbLf
    s = s & "        const names = wb.names; names.load(""items/name"");" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        const infos = sheets.items.map((ws) => {" & vbLf
    s = s & "            const used = ws.getUsedRangeOrNullObject(true); used.load(""address,rowCount,columnCount,rowIndex,columnIndex"");" & vbLf
    s = s & "            const tables = ws.tables; tables.load(""items/name"");" & vbLf
    s = s & "            const charts = ws.charts; charts.load(""items/name"");" & vbLf
    s = s & "            return { ws: ws, used: used, tables: tables, charts: charts, scan: false };" & vbLf
    s = s & "        });" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        for (const i of infos) { if (!i.used.isNullObject && i.used.rowCount * i.used.columnCount <= LIMIT) { i.used.load(""valueTypes,formulas""); i.scan = true; } }" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        const out = infos.map((i) => {" & vbLf
    s = s & "            const info = { name: i.ws.name, used: """", rows: 0, cols: 0, formulas: 0, errors: 0, errorCells: [], tables: i.tables.items.length, charts: i.charts.items.length, scanned: i.scan };" & vbLf
    s = s & "            if (!i.used.isNullObject) {" & vbLf
    s = s & "                info.used = String(i.used.address).split(""!"").pop();" & vbLf
    s = s & "                info.rows = i.used.rowCount;" & vbLf
    s = s & "                info.cols = i.used.columnCount;" & vbLf
    s = s & "                if (i.scan) {" & vbLf
    s = s & "                    for (let r = 0; r < info.rows; r++) {" & vbLf
    s = s & "                        for (let c = 0; c < info.cols; c++) {" & vbLf
    s = s & "                            const f = i.used.formulas[r][c];" & vbLf
    s = s & "                            if (typeof f === ""string"" && f.charAt(0) === ""="") { info.formulas++; }" & vbLf
    s = s & "                            if (i.used.valueTypes[r][c] === ""Error"") {" & vbLf
    s = s & "                                info.errors++;" & vbLf
    s = s & "                                if (info.errorCells.length < MAX_ERR) { info.errorCells.push(col(i.used.columnIndex + c) + (i.used.rowIndex + r + 1)); }" & vbLf
    s = s & "                            }" & vbLf
    s = s & "                        }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            return info;" & vbLf
    s = s & "        });" & vbLf
    s = s & "        return JSON.stringify({ workbook: wb.name, names: names.items.length, sheets: out, ms: Date.now() - T0 });" & vbLf
    JS_Inspect = s
End Function

'* ライブミラー：アクティブなシートの表示テキストを JSON で返す
Private Function JS_MirrorRead() As String
    Dim s As String
    s = s & "        const MAX_ROWS = __MAX_ROWS__, MAX_COLS = __MAX_COLS__;" & vbLf
    s = s & "        const ws = context.workbook.worksheets.getActiveWorksheet(); ws.load(""name"");" & vbLf
    s = s & "        const used = ws.getUsedRangeOrNullObject(true); used.load(""rowCount,columnCount,rowIndex,columnIndex"");" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        if (used.isNullObject) { return JSON.stringify({ sheet: ws.name, address: """", texts: [] }); }" & vbLf
    s = s & "        const area = ws.getRangeByIndexes(used.rowIndex, used.columnIndex, Math.min(used.rowCount, MAX_ROWS), Math.min(used.columnCount, MAX_COLS));" & vbLf
    s = s & "        area.load(""address,text"");" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        return JSON.stringify({ sheet: ws.name, address: String(area.address).split(""!"").pop(), texts: area.text });" & vbLf
    JS_MirrorRead = s
End Function

'* 後片付け：ショーケースのシートを削除
Private Function JS_Cleanup() As String
    Dim s As String
    s = s & "        const NAMES = [__SHEETS__];" & vbLf
    s = s & "        const sheets = context.workbook.worksheets;" & vbLf
    s = s & "        const found = NAMES.map((n) => sheets.getItemOrNullObject(n));" & vbLf
    s = s & "        sheets.load(""items/name"");" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        let deleted = 0, left = sheets.items.length;" & vbLf
    s = s & "        for (const ws of found) {" & vbLf
    s = s & "            if (ws.isNullObject) { continue; }" & vbLf
    s = s & "            if (left <= 1) { sheets.add(); left++; }" & vbLf
    s = s & "            ws.delete();" & vbLf
    s = s & "            deleted++;" & vbLf
    s = s & "            left--;" & vbLf
    s = s & "        }" & vbLf
    s = s & "        await context.sync();" & vbLf
    s = s & "        return ""deleted "" + deleted;" & vbLf
    JS_Cleanup = s
End Function
