Attribute VB_Name = "Demo_OfficeScript_ShowcaseLoop"
'***************************************************************************************************
'       exCDP_OfficeScript 拡張 - ショーケース(ループ一括送信版) モジュール
'***************************************************************************************************
'* 機能　　：Demo_OfficeScript_Showcase の描画系を、「CDP で投げるのは1回、ループは Officeスクリプトの中で回す」形にした版です
'            ・ShowcaseLoop_00 全部まとめて(バナー → マンデルブロ → ライフゲーム → バナー を、1本のスクリプトで通しで流す)
'            ・ShowcaseLoop_01 マンデルブロ集合のズーム
'            ・ShowcaseLoop_02 ライフゲーム
'            ・ShowcaseLoop_03 ドット絵バナー
'---------------------------------------------------------------------------------------------------
'* 元の版との違い：
'            ・元の版は1フレームごとにスクリプトを書き込んで実行していました(Web 版で1フレーム約1.5秒)。
'              この版は、1つのデモにつき書き込みと実行は1回で、フレームのループは `async function main` の中で回します
'            ・フレームの区切りでは、セルを1つ読み取って、溜まった書き込みを Excel へ送らせます
'              (Officeスクリプトは、読み取りが来るまで書き込みをまとめて送るため。読み取らないと最後のフレームしか見えません)
'            ・ライフゲームとバナーは、前のフレームを Excel から読み直さず、スクリプト側のメモリに持ったまま進めます
'            ・結果として、フレーム数・処理時間(ms)を console.log の JSON で受け取ります
'            ・フレーム数と倍率は、比較しやすいように元の版と同じにしています
'            ・読み取り系(棚卸し・ライブミラー)は、1回ごとに VBA が手元のブックへ書き出す作りのため、元の版を使ってください
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\OfficeScriptViaVBA\OfficialRoute\exCDP_OfficeScript.cls
'---------------------------------------------------------------------------------------------------
'* 注意事項：・「exCDP_OfficeScript.cls」と「Demo_OfficeScript.bas」(ConnectWebExcel / ConnectDesktopExcel を使用)も
'              取り込んでください
'            ・既定は Web 版 Excel(Edge)で実行します。デスクトップ版で動かす場合は SHOWCASE_WEB を False にしてください
'            ・デスクトップ版では、実行を始めたらマクロをいったん終え、`Application.OnTime` で結果を覗きます
'              (ShowcaseLoop_Poll)。デスクトップ版の Officeスクリプトは Excel と同じプロセスで動くため、VBA が待機ループを
'              回している間は、フレームごとの読み取りが進まないためです
'            ・描画先のシート名は元の版と同じ(Mandelbrot / LifeGame / Banner)です。
'              不要になったら Demo_OfficeScript_Showcase の Showcase_99_後片付け で削除できます
'            ・エディターで開いているスクリプトの中身は、デモのスクリプトで上書きされます
'            ・ループは Officeスクリプトの中で回るため、Ctrl+Break で VBA を止めても、スクリプトは最後まで続きます。
'              LOOP_TIMEOUT_SEC の10秒前を過ぎると、スクリプト側で次のフレームに入る前に例外で打ち切ります
'            ・フレーム間の待ちには setTimeout を使います。Officeスクリプトの実行環境に無い場合は、時刻を見ながら
'              空回りして待ちます(FRAME_WAIT_MS = 0 で待ち自体を無くせます)
'***************************************************************************************************
Option Explicit

'接続先(True:Web 版 Excel / False:デスクトップ版 Excel)
Private Const SHOWCASE_WEB      As Boolean = True

'1フレームごとに、スクリプト側で待つ時間(ミリ秒)。Excel に描画させるための間
Private Const FRAME_WAIT_MS     As Long = 50

'1本のスクリプトを待つ上限(秒)。スクリプト側のループは、この10秒前で打ち切る
Private Const LOOP_TIMEOUT_SEC  As Double = 600

'デスクトップ版で、結果を覗く間隔(秒)
Private Const POLL_SEC          As Long = 2

'マンデルブロ集合
Private Const MANDEL_SHEET      As String = "Mandelbrot"
Private Const MANDEL_W          As Long = 150
Private Const MANDEL_H          As Long = 100
Private Const MANDEL_SIZE       As Long = 4
Private Const MANDEL_CX         As Double = -0.743643887037151
Private Const MANDEL_CY         As Double = 0.13182590420533
Private Const MANDEL_ZOOM       As Double = 0.82        '1フレームごとの倍率
Private Const MANDEL_MAX_STEP   As Long = 12            '1フレームごとに増やす反復回数

'ライフゲーム
Private Const LIFE_SHEET        As String = "LifeGame"
Private Const LIFE_W            As Long = 80
Private Const LIFE_H            As Long = 50
Private Const LIFE_SIZE         As Long = 8

'ドット絵バナー
Private Const BANNER_SHEET      As String = "Banner"
Private Const BANNER_SIZE       As Long = 12
Private Const BANNER_CHARS      As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 -><!.:?*"

'デスクトップ版で、OnTime で結果を覗くあいだ保持しておくもの
Private m_Browser               As CDPBrowser
Private m_OS                    As exCDP_OfficeScript
Private m_Ticket                As Long
Private m_Label                 As String
Private m_StartSec              As Double



'***************************************************************************************************
'                            ■■■ 初めに ■■■
'***************************************************************************************************
'* 機能　　：デバッグ経路を確保します
'***************************************************************************************************
Sub FirstFunction()
    Dim tmp As New exCDP_OfficeJs
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
'                         ■■■ ShowcaseLoop 00：全部まとめて(録画用) ■■■
'***************************************************************************************************
'* 機能　　：バナー → マンデルブロ集合のズーム → ライフゲーム → 締めのバナー を、1本のスクリプトで通しで流します
'***************************************************************************************************
Sub ShowcaseLoop_00_全部まとめて()
    Dim Body As String
    Body = PartBanner("VBA -> CDP", 20) & _
           PartMandelbrot(40) & _
           PartLife(90) & _
           PartBanner("ONE SCRIPT!", 20)
    PlayLoop BuildLoop(Body), "All in one script"
End Sub



'***************************************************************************************************
'                         ■■■ ShowcaseLoop 01：マンデルブロ集合のズーム ■■■
'***************************************************************************************************
Sub ShowcaseLoop_01_マンデルブロズーム()
    PlayLoop BuildLoop(PartMandelbrot(40)), "Mandelbrot 40 frames"
End Sub



'***************************************************************************************************
'                             ■■■ ShowcaseLoop 02：ライフゲーム ■■■
'***************************************************************************************************
Sub ShowcaseLoop_02_ライフゲーム()
    PlayLoop BuildLoop(PartLife(90)), "Life 90 generations"
End Sub



'***************************************************************************************************
'                            ■■■ ShowcaseLoop 03：ドット絵バナー ■■■
'***************************************************************************************************
'* 注意事項：使える文字は英大文字・数字・空白と「-><!.:?*」です。英小文字は大文字に、それ以外は「?」になります
'***************************************************************************************************
Sub ShowcaseLoop_03_ドット絵バナー()
    PlayLoop BuildLoop(PartBanner("VBA -> CDP", 30)), "Banner 30 frames"
End Sub



'***************************************************************************************************
'                                  ■■■ 再生処理 ■■■
'***************************************************************************************************
'* 機能　　：接続 → スクリプトを1回だけ実行 → 結果の表示 → 後片付け、を行います
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Web 版は同期で待ちます。デスクトップ版は実行を始めたらマクロを終え、OnTime(ShowcaseLoop_Poll)で回収します
'***************************************************************************************************
Private Sub PlayLoop(Script As String, Label As String)
    If Not m_OS Is Nothing Then
        MsgBox "前回のスクリプトの結果を待っています。終わるまでお待ちください。", vbExclamation
        Exit Sub
    End If

    Dim Browser As CDPBrowser, os As exCDP_OfficeScript
    If Not OpenShowcase(Browser, os) Then Exit Sub
    Debug.Print "[ShowcaseLoop] " & Label & " : start"

    '1. Web 版：同期実行で待つ
    If SHOWCASE_WEB Then
        Dim StartSec As Double, Status As String
        StartSec = Timer
        Status = os.RunScript(Script, LOOP_TIMEOUT_SEC)
        ReportResult os, Label, Status, ElapsedSec(StartSec)
        CloseShowcase Browser, os
        Exit Sub
    End If

    '2. デスクトップ版：実行を始めたらマクロを終え、OnTime で結果を覗く
    m_Ticket = os.RunScriptAsync(Script)
    If m_Ticket = 0 Then
        Debug.Print "[ShowcaseLoop] " & Label & " : 開始できませんでした" & vbLf & os.LastLogs
        CloseShowcase Browser, os
        Exit Sub
    End If

    Set m_Browser = Browser
    Set m_OS = os
    m_Label = Label
    m_StartSec = Timer
    Debug.Print "[ShowcaseLoop] " & Label & " : started (desktop). ticket: " & m_Ticket & " / マクロを終えて、結果を待ちます"
    Application.OnTime Now + TimeSerial(0, 0, POLL_SEC), "Demo_OfficeScript_ShowcaseLoop.ShowcaseLoop_Poll"
End Sub

'***************************************************************************************************
'* 機能　　：デスクトップ版の結果を覗きます(`Application.OnTime`から呼ばれる)
'***************************************************************************************************
Sub ShowcaseLoop_Poll()
    If m_OS Is Nothing Then Exit Sub

    '1. 整理券番号で、1回だけ確認
    Dim Status As String
    If Not m_OS.TryGetResult(m_Ticket, Status) Then
        '未完了：上限を過ぎていなければ再予約
        If ElapsedSec(m_StartSec) < LOOP_TIMEOUT_SEC Then
            Application.OnTime Now + TimeSerial(0, 0, POLL_SEC), "Demo_OfficeScript_ShowcaseLoop.ShowcaseLoop_Poll"
            Exit Sub
        End If
        Status = "Timeout"
    End If

    '2. 結果の表示
    ReportResult m_OS, m_Label, Status, ElapsedSec(m_StartSec)

    '3. 後片付け(先に Class を手放してから切断)
    CloseShowcase m_Browser, m_OS
End Sub

'***************************************************************************************************
'* 機能　　：結果(状態・かかった時間・スクリプトが出したフレーム数と処理時間)を表示します
'***************************************************************************************************
Private Sub ReportResult(os As exCDP_OfficeScript, Label As String, Status As String, Elapsed As Double)
    Debug.Print "[ShowcaseLoop] " & Label & " : " & Status & " / " & Format$(Elapsed, "0.00") & " 秒 " & LastJsonOutput(os)
    If Status <> "Succeeded" Then
        Debug.Print os.LastLogs
        If Len(os.LastExceptionMessage) Then Debug.Print "[ShowcaseLoop] exception: " & os.LastExceptionMessage
    End If
End Sub



'***************************************************************************************************
'                                  ■■■ 共通処理 ■■■
'***************************************************************************************************
Private Function OpenShowcase(Browser As CDPBrowser, os As exCDP_OfficeScript) As Boolean
    If SHOWCASE_WEB Then
        Set Browser = ConnectWebExcel()
    Else
        Set Browser = ConnectDesktopExcel()
    End If
    If Browser Is Nothing Then
        MsgBox "接続できませんでした。イミディエイトウィンドウを確認してください。", vbCritical
        Exit Function
    End If

    Set os = New exCDP_OfficeScript
    If Not os.Init(Browser, WebMode:=SHOWCASE_WEB) Then
        CloseShowcase Browser, os
        MsgBox "コード エディターが見つかりませんでした。イミディエイトウィンドウを確認してください。", vbCritical
        Exit Function
    End If

    OpenShowcase = True
End Function

Private Sub CloseShowcase(Browser As CDPBrowser, os As exCDP_OfficeScript)
    If Not os Is Nothing Then os.Dispose
    Set os = Nothing
    If Not Browser Is Nothing Then Browser.DisconnectCDP
    Set Browser = Nothing
End Sub

'* 直近の実行の console.log のうち、最後の JSON(「{」で始まる行)を返します。無ければ空文字
Private Function LastJsonOutput(os As exCDP_OfficeScript) As String
    Dim i As Long, Line As String
    For i = os.LastOutputs.Count To 1 Step -1
        Line = CStr(os.LastOutputs(i))
        If Left$(Line, 1) = "{" Then
            LastJsonOutput = Line
            Exit Function
        End If
    Next
End Function

Private Function ElapsedSec(StartSec As Double) As Double
    ElapsedSec = Timer - StartSec
    If ElapsedSec < 0 Then ElapsedSec = ElapsedSec + 86400#   '日付またぎ
End Function

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
'                     ■■■ Officeスクリプト(TypeScript)の組み立て ■■■
'***************************************************************************************************
'* 機能　　：`async function main` と共通の前後処理(フレームの区切り・待ち・打ち切り)で、各パートを包みます
'***************************************************************************************************
Private Function BuildLoop(Parts As String) As String
    Dim s As String
    s = TS_LoopHeader()
    s = Replace(s, "__WAIT_MS__", CStr(FRAME_WAIT_MS))
    s = Replace(s, "__LIMIT_MS__", CStr(CLng((LOOP_TIMEOUT_SEC - 10) * 1000)))
    BuildLoop = s & Parts & TS_LoopFooter()
End Function

Private Function PartMandelbrot(Frames As Long) As String
    Dim s As String
    s = TS_MandelPart()
    s = Replace(s, "__W__", CStr(MANDEL_W))
    s = Replace(s, "__H__", CStr(MANDEL_H))
    s = Replace(s, "__SIZE__", CStr(MANDEL_SIZE))
    s = Replace(s, "__MID__", "85")
    s = Replace(s, "__FRAMES__", CStr(Frames))
    s = Replace(s, "__CX__", Num(MANDEL_CX))
    s = Replace(s, "__CY__", Num(MANDEL_CY))
    s = Replace(s, "__ZOOM__", Num(MANDEL_ZOOM))
    s = Replace(s, "__MAX_STEP__", CStr(MANDEL_MAX_STEP))
    s = Replace(s, "__SHEET__", MANDEL_SHEET)
    PartMandelbrot = s
End Function

Private Function PartLife(Generations As Long) As String
    Dim s As String
    s = TS_LifePart()
    s = Replace(s, "__W__", CStr(LIFE_W))
    s = Replace(s, "__H__", CStr(LIFE_H))
    s = Replace(s, "__SIZE__", CStr(LIFE_SIZE))
    s = Replace(s, "__GENS__", CStr(Generations))
    s = Replace(s, "__SHEET__", LIFE_SHEET)
    PartLife = s
End Function

Private Function PartBanner(ByVal Text As String, Frames As Long) As String
    '1. フォントにない文字を置き換える(スクリプトの文字列にそのまま埋め込むため、ここで必ず絞る)
    Text = SanitizeBannerText(Text)

    '2. 長い文字は、横幅が収まるようにセルを小さくする
    Dim Size As Long
    Size = BANNER_SIZE
    If Len(Text) > 10 Then Size = Int(BANNER_SIZE * 10 / Len(Text))
    If Size < 4 Then Size = 4

    Dim s As String
    s = TS_BannerPart()
    s = Replace(s, "__TEXT__", Text)
    s = Replace(s, "__SIZE__", CStr(Size))
    s = Replace(s, "__FRAMES__", CStr(Frames))
    s = Replace(s, "__SHEET__", BANNER_SHEET)
    PartBanner = s
End Function



'***************************************************************************************************
'                     ■■■ Officeスクリプト(TypeScript)のひな形 ■■■
'***************************************************************************************************
'* 機能　　：「__名前__」の部分を、VBA 側で Replace して使います
'***************************************************************************************************
'* 共通の前処理：frame() で1フレームを確定(読み取りで同期)し、打ち切りの確認と、フレーム間の待ちを行う
Private Function TS_LoopHeader() As String
    Dim s As String
    s = s & "async function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const WAIT_MS = __WAIT_MS__, LIMIT_MS = __LIMIT_MS__, T0 = Date.now();" & vbLf
    s = s & "    let frames = 0;" & vbLf
    s = s & "    const g = globalThis as unknown as { setTimeout?: (handler: () => void, ms: number) => number };" & vbLf
    s = s & "    const frame = async (ws: ExcelScript.Worksheet) => {" & vbLf
    s = s & "        ws.getRange(""A1"").getValue();" & vbLf
    s = s & "        frames++;" & vbLf
    s = s & "        if (Date.now() - T0 > LIMIT_MS) { throw new Error(""loop limit exceeded ("" + LIMIT_MS + "" ms, "" + frames + "" frames)""); }" & vbLf
    s = s & "        if (WAIT_MS <= 0) { return; }" & vbLf
    s = s & "        if (typeof g.setTimeout === ""function"") {" & vbLf
    s = s & "            await new Promise<void>((resolve) => { g.setTimeout!(() => resolve(), WAIT_MS); });" & vbLf
    s = s & "        } else {" & vbLf
    s = s & "            const until = Date.now() + WAIT_MS;" & vbLf
    s = s & "            while (Date.now() < until) { }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    };" & vbLf
    s = s & "    const prepareSheet = (name: string): ExcelScript.Worksheet => {" & vbLf
    s = s & "        let ws = workbook.getWorksheet(name);" & vbLf
    s = s & "        if (!ws) { ws = workbook.addWorksheet(name); }" & vbLf
    s = s & "        ws.activate();" & vbLf
    s = s & "        ws.getRange().clear(ExcelScript.ClearApplyTo.all);" & vbLf
    s = s & "        ws.setShowGridlines(false);" & vbLf
    s = s & "        return ws;" & vbLf
    s = s & "    };" & vbLf
    TS_LoopHeader = s
End Function

'* 共通の後処理：フレーム数と処理時間を console.log に出す
Private Function TS_LoopFooter() As String
    Dim s As String
    s = s & "    console.log(JSON.stringify({ frames: frames, ms: Date.now() - T0 }));" & vbLf
    s = s & "}" & vbLf
    TS_LoopFooter = s
End Function

'* マンデルブロ集合：シートの準備 + FRAMES フレームのズーム
Private Function TS_MandelPart() As String
    Dim s As String
    s = s & "    {" & vbLf
    s = s & "        const W = __W__, H = __H__, SIZE = __SIZE__, FRAMES = __FRAMES__;" & vbLf
    s = s & "        const CX = __CX__, CY = __CY__, ZOOM = __ZOOM__, MAX_STEP = __MAX_STEP__;" & vbLf
    s = s & "        const ws = prepareSheet(""__SHEET__"");" & vbLf
    s = s & "        const title = ws.getRangeByIndexes(0, 0, 1, W);" & vbLf
    s = s & "        title.merge(false);" & vbLf
    s = s & "        title.getFormat().getFont().setBold(true);" & vbLf
    s = s & "        title.getFormat().getFont().setSize(14);" & vbLf
    s = s & "        const grid = ws.getRangeByIndexes(1, 0, H, W);" & vbLf
    s = s & "        grid.getFormat().setColumnWidth(SIZE);" & vbLf
    s = s & "        grid.getFormat().setRowHeight(SIZE);" & vbLf
    s = s & "        grid.setNumberFormat("";;;"");" & vbLf
    s = s & "        grid.getFormat().getFill().setColor(""#000000"");" & vbLf
    s = s & "        const cf = grid.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);" & vbLf
    s = s & "        cf.getColorScale().setCriteria({" & vbLf
    s = s & "            minimum: { type: ExcelScript.ConditionalFormatColorCriterionType.lowestValue, color: ""#020B3A"" }," & vbLf
    s = s & "            midpoint: { type: ExcelScript.ConditionalFormatColorCriterionType.percentile, formula: ""__MID__"", color: ""#FF9A1F"" }," & vbLf
    s = s & "            maximum: { type: ExcelScript.ConditionalFormatColorCriterionType.highestValue, color: ""#FFFFF0"" }" & vbLf
    s = s & "        });" & vbLf
    s = s & "        ws.getRange(""A1"").select();" & vbLf
    s = s & "        let SCALE = 3;" & vbLf
    s = s & "        for (let f = 0; f < FRAMES; f++) {" & vbLf
    s = s & "            const MAX = 120 + f * MAX_STEP;" & vbLf
    s = s & "            const rows: (number | string)[][] = [];" & vbLf
    s = s & "            for (let y = 0; y < H; y++) {" & vbLf
    s = s & "                const row: (number | string)[] = [];" & vbLf
    s = s & "                const ci = CY + (y / H - 0.5) * SCALE;" & vbLf
    s = s & "                for (let x = 0; x < W; x++) {" & vbLf
    s = s & "                    const cr = CX + (x / W - 0.5) * SCALE * (W / H);" & vbLf
    s = s & "                    let zr = 0, zi = 0, i = 0;" & vbLf
    s = s & "                    while (zr * zr + zi * zi < 16 && i < MAX) {" & vbLf
    s = s & "                        const t = zr * zr - zi * zi + cr;" & vbLf
    s = s & "                        zi = 2 * zr * zi + ci;" & vbLf
    s = s & "                        zr = t;" & vbLf
    s = s & "                        i++;" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                    if (i >= MAX) {" & vbLf
    s = s & "                        row.push("""");" & vbLf
    s = s & "                    } else {" & vbLf
    s = s & "                        const smooth = i + 1 - Math.log(Math.log(Math.sqrt(zr * zr + zi * zi))) / Math.LN2;" & vbLf
    s = s & "                        row.push(Math.round(Math.sqrt(smooth) * 1000) / 1000);" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "                rows.push(row);" & vbLf
    s = s & "            }" & vbLf
    s = s & "            grid.setValues(rows);" & vbLf
    s = s & "            title.getCell(0, 0).setValue(""VBA -> CDP(1 call) -> Office Scripts loop : Mandelbrot  x"" + Math.round(3 / SCALE) + ""  "" + (f + 1) + ""/"" + FRAMES);" & vbLf
    s = s & "            await frame(ws);" & vbLf
    s = s & "            SCALE *= ZOOM;" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    TS_MandelPart = s
End Function

'* ライフゲーム：シートの準備(グライダー銃とパルサー) + GENS 世代
Private Function TS_LifePart() As String
    Dim s As String
    s = s & "    {" & vbLf
    s = s & "        const W = __W__, H = __H__, SIZE = __SIZE__, GENS = __GENS__;" & vbLf
    s = s & "        const ws = prepareSheet(""__SHEET__"");" & vbLf
    s = s & "        const title = ws.getRangeByIndexes(0, 0, 1, W);" & vbLf
    s = s & "        title.merge(false);" & vbLf
    s = s & "        title.getFormat().getFont().setBold(true);" & vbLf
    s = s & "        title.getFormat().getFont().setSize(14);" & vbLf
    s = s & "        const grid = ws.getRangeByIndexes(1, 0, H, W);" & vbLf
    s = s & "        grid.getFormat().setColumnWidth(SIZE);" & vbLf
    s = s & "        grid.getFormat().setRowHeight(SIZE);" & vbLf
    s = s & "        grid.setNumberFormat("";;;"");" & vbLf
    s = s & "        grid.getFormat().getFill().setColor(""#05050F"");" & vbLf
    s = s & "        const cf = grid.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);" & vbLf
    s = s & "        cf.getColorScale().setCriteria({" & vbLf
    s = s & "            minimum: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""1"", color: ""#39FF14"" }," & vbLf
    s = s & "            midpoint: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""6"", color: ""#00B4FF"" }," & vbLf
    s = s & "            maximum: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""30"", color: ""#B000FF"" }" & vbLf
    s = s & "        });" & vbLf
    s = s & "        ws.getRange(""A1"").select();" & vbLf
    s = s & "        let age: number[][] = [];" & vbLf
    s = s & "        for (let y = 0; y < H; y++) {" & vbLf
    s = s & "            const row: number[] = [];" & vbLf
    s = s & "            for (let x = 0; x < W; x++) { row.push(0); }" & vbLf
    s = s & "            age.push(row);" & vbLf
    s = s & "        }" & vbLf
    s = s & "        const put = (pattern: string[], top: number, left: number) => {" & vbLf
    s = s & "            for (let y = 0; y < pattern.length; y++) {" & vbLf
    s = s & "                for (let x = 0; x < pattern[y].length; x++) {" & vbLf
    s = s & "                    if (pattern[y][x] === ""O"" && top + y < H && left + x < W) { age[top + y][left + x] = 1; }" & vbLf
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
    s = s & "        for (let gen = 0; gen <= GENS; gen++) {" & vbLf
    s = s & "            if (gen > 0) {" & vbLf
    s = s & "                const next: number[][] = [];" & vbLf
    s = s & "                for (let y = 0; y < H; y++) {" & vbLf
    s = s & "                    const row: number[] = [];" & vbLf
    s = s & "                    for (let x = 0; x < W; x++) {" & vbLf
    s = s & "                        let n = 0;" & vbLf
    s = s & "                        for (let dy = -1; dy <= 1; dy++) {" & vbLf
    s = s & "                            for (let dx = -1; dx <= 1; dx++) {" & vbLf
    s = s & "                                if (dx === 0 && dy === 0) { continue; }" & vbLf
    s = s & "                                const yy = y + dy, xx = x + dx;" & vbLf
    s = s & "                                if (yy >= 0 && yy < H && xx >= 0 && xx < W && age[yy][xx] > 0) { n++; }" & vbLf
    s = s & "                            }" & vbLf
    s = s & "                        }" & vbLf
    s = s & "                        const a = age[y][x];" & vbLf
    s = s & "                        row.push(a > 0 ? ((n === 2 || n === 3) ? Math.min(a + 1, 99) : 0) : (n === 3 ? 1 : 0));" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                    next.push(row);" & vbLf
    s = s & "                }" & vbLf
    s = s & "                age = next;" & vbLf
    s = s & "            }" & vbLf
    s = s & "            grid.setValues(age.map(r => r.map(v => (v > 0 ? v : """"))));" & vbLf
    s = s & "            title.getCell(0, 0).setValue(""VBA -> CDP(1 call) -> Office Scripts loop : Game of Life  gen "" + gen + ""/"" + GENS);" & vbLf
    s = s & "            await frame(ws);" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    TS_LifePart = s
End Function

'* ドット絵バナー：シートの準備(5x7ドットのフォント) + FRAMES フレームのグラデーション
Private Function TS_BannerPart() As String
    Dim s As String
    s = s & "    {" & vbLf
    s = s & "        const TEXT = ""__TEXT__"", SIZE = __SIZE__, PAD = 2, FRAMES = __FRAMES__;" & vbLf
    s = s & "        const FONT: { [key: string]: string } = {" & vbLf
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
    s = s & "        const ws = prepareSheet(""__SHEET__"");" & vbLf
    s = s & "        const area = ws.getRangeByIndexes(0, 0, H, W);" & vbLf
    s = s & "        area.getFormat().setColumnWidth(SIZE);" & vbLf
    s = s & "        area.getFormat().setRowHeight(SIZE);" & vbLf
    s = s & "        area.setNumberFormat("";;;"");" & vbLf
    s = s & "        area.getFormat().getFill().setColor(""#000000"");" & vbLf
    s = s & "        const cf = area.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);" & vbLf
    s = s & "        cf.getColorScale().setCriteria({" & vbLf
    s = s & "            minimum: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""0"", color: ""#FF2E88"" }," & vbLf
    s = s & "            midpoint: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""50"", color: ""#00E5FF"" }," & vbLf
    s = s & "            maximum: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""100"", color: ""#FFE600"" }" & vbLf
    s = s & "        });" & vbLf
    s = s & "        const caption = ws.getRangeByIndexes(H, 0, 1, W);" & vbLf
    s = s & "        caption.merge(false);" & vbLf
    s = s & "        caption.getCell(0, 0).setValue(""Drawn by Office Scripts loop, driven from VBA via 1 CDP call"");" & vbLf
    s = s & "        caption.getFormat().setRowHeight(20);" & vbLf
    s = s & "        caption.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);" & vbLf
    s = s & "        caption.getFormat().getFont().setColor(""#888888"");" & vbLf
    s = s & "        ws.getRange(""A1"").select();" & vbLf
    s = s & "        const dots: boolean[][] = [];" & vbLf
    s = s & "        for (let y = 0; y < H; y++) {" & vbLf
    s = s & "            const row: boolean[] = [];" & vbLf
    s = s & "            for (let x = 0; x < W; x++) { row.push(false); }" & vbLf
    s = s & "            dots.push(row);" & vbLf
    s = s & "        }" & vbLf
    s = s & "        for (let c = 0; c < TEXT.length; c++) {" & vbLf
    s = s & "            const glyph = FONT[TEXT[c]] || FONT[""?""];" & vbLf
    s = s & "            for (let y = 0; y < 7; y++) {" & vbLf
    s = s & "                for (let x = 0; x < 5; x++) {" & vbLf
    s = s & "                    if (glyph[y * 5 + x] === ""1"") { dots[PAD + y][PAD + c * 6 + x] = true; }" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        for (let f = 0; f <= FRAMES; f++) {" & vbLf
    s = s & "            const out: (number | string)[][] = dots.map((r, y) => r.map((on, x) => {" & vbLf
    s = s & "                if (!on) { return """"; }" & vbLf
    s = s & "                const p = (x * 6 + y * 3 + f * 12) % 200;" & vbLf
    s = s & "                return p <= 100 ? p : 200 - p;" & vbLf
    s = s & "            }));" & vbLf
    s = s & "            area.setValues(out);" & vbLf
    s = s & "            await frame(ws);" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    TS_BannerPart = s
End Function
