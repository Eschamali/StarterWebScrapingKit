Attribute VB_Name = "Demo_OfficeScript_Showcase"
'***************************************************************************************************
'       exCDP_OfficeScript 拡張 - ショーケース(録画向けデモ) モジュール
'***************************************************************************************************
'* 機能　　：VBA が Officeスクリプト(TypeScript)の計算力を借りて、セルをピクセル代わりに絵を描くデモです
'            ・Showcase_01 マンデルブロ集合のズーム(1フレームごとに中心と倍率を差し替えて実行)
'            ・Showcase_02 ライフゲーム(グライダー銃とパルサー。1世代ごとに実行)
'            ・Showcase_03 ドット絵バナー(5x7ドットの文字に、虹色のグラデーションを流す)
'            ・Showcase_00 全部まとめて(録画用の通し)
'            読み取り系(Web 版ブックの中身を、VBA が受け取って手元のブックで使う)
'            ・Showcase_04 ブックの棚卸し(全シートの使用範囲・数式・エラー・テーブルを調べ、手元にレポートとグラフを作る)
'            ・Showcase_05 ライブミラー(Web 版でアクティブなシートを繰り返し読み取り、手元のブックに写す。変わったセルは光る)
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\OfficeScriptViaVBA\exCDP_OfficeScript.cls
'---------------------------------------------------------------------------------------------------
'* 注意事項：・「exCDP_OfficeScript.cls」と「Demo_OfficeScript.bas」(ConnectWebExcel / ConnectDesktopExcel を使用)
'              も取り込んでください
'            ・既定は Web 版 Excel(Edge)で実行します。デスクトップ版で動かす場合は SHOWCASE_WEB を False に
'              してください。1フレームあたり、Web 版は約1.5秒、デスクトップ版は約4～5秒かかります
'            ・ブックに「Mandelbrot」「LifeGame」「Banner」シートを作って描きます(既にあれば中身を消して描き直します)。
'              不要になったら Showcase_99_後片付け で削除できます
'            ・読み取り系(04 / 05)は、Web 版のブックを書き換えません。結果は、このマクロのブック(ThisWorkbook)の
'              「WebBookReport」「WebMirror」シートに書き出します(既にあれば中身を消して書き直します)
'            ・エディターで開いているスクリプトの中身は、デモのスクリプトで上書きされます
'            ・ループ中に止めたいときは、Ctrl+Break で VBA を中断してください
'            ・録画するときは、Excel の表示倍率を下げて全体が入るようにし、録画後に2～4倍速にすると滑らかに見えます
'***************************************************************************************************
Option Explicit

'接続先(True:Web 版 Excel / False:デスクトップ版 Excel)
Private Const SHOWCASE_WEB      As Boolean = True

'マンデルブロ集合
Private Const MANDEL_SHEET      As String = "Mandelbrot"
Private Const MANDEL_W          As Long = 150           '横のピクセル(セル)数
Private Const MANDEL_H          As Long = 100           '縦のピクセル(セル)数
Private Const MANDEL_SIZE       As Long = 4             'セル1つの大きさ(ポイント)
Private Const MANDEL_CX         As Double = -0.743643887037151  'ズームの中心(タツノオトシゴの谷)
Private Const MANDEL_CY         As Double = 0.13182590420533
Private Const MANDEL_ZOOM       As Double = 0.82        '1フレームごとの倍率(小さいほど速く寄る)

'ライフゲーム
Private Const LIFE_SHEET        As String = "LifeGame"
Private Const LIFE_W            As Long = 80
Private Const LIFE_H            As Long = 50
Private Const LIFE_SIZE         As Long = 8

'ドット絵バナー
Private Const BANNER_SHEET      As String = "Banner"
Private Const BANNER_SIZE       As Long = 12
Private Const BANNER_CHARS      As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 -><!.:?*"  'フォントにある文字

'読み取り系(結果の書き出し先は、このマクロのブック)
Private Const REPORT_SHEET      As String = "WebBookReport"
Private Const INSPECT_LIMIT     As Long = 200000        'これより大きい使用範囲のシートは、数式とエラーを数えない
Private Const MIRROR_SHEET      As String = "WebMirror"
Private Const MIRROR_MAX_ROWS   As Long = 40
Private Const MIRROR_MAX_COLS   As Long = 12
Private Const MIRROR_TOP        As Long = 4             '写す先の開始行

'棚卸し(デスクトップ版)：マクロを終えてから OnTime で結果を覗く
Private Const INSPECT_POLL_SEC      As Long = 2         '覗く間隔(秒)
Private Const INSPECT_TIMEOUT_SEC   As Double = 300     '待つ上限(秒)

'棚卸し(デスクトップ版)：マクロ終了後も、接続と整理券番号を保持する
Private m_InspectBrowser    As CDPBrowser
Private m_InspectOS         As exCDP_OfficeScript
Private m_InspectTicket     As Long
Private m_InspectStart      As Double



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
'* 機能　　：バナー → マンデルブロ集合のズーム → ライフゲーム → 締めのバナー、の順に通しで流します
'***************************************************************************************************
Sub Showcase_00_全部まとめて()
    Dim Browser As CDPBrowser, os As exCDP_OfficeScript
    If Not OpenShowcase(Browser, os) Then Exit Sub

    If PlayBanner(os, "VBA -> CDP", 20) Then
        If PlayMandelbrot(os, 40) Then
            If PlayLife(os, 90) Then
                PlayBanner os, "OFFICE SCRIPTS!", 20
            End If
        End If
    End If

    CloseShowcase Browser, os
End Sub



'***************************************************************************************************
'                         ■■■ Showcase 01：マンデルブロ集合のズーム ■■■
'***************************************************************************************************
'* 機能　　：150x100 のセルをピクセルにして、マンデルブロ集合へズームインしていきます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：計算は Officeスクリプト(JavaScript)側で行い、セルには値を一括で書き込むだけです。
'            色は条件付き書式のカラースケールに任せ、数値は表示形式「;;;」で隠しています
'***************************************************************************************************
Sub Showcase_01_マンデルブロズーム()
    Dim Browser As CDPBrowser, os As exCDP_OfficeScript
    If Not OpenShowcase(Browser, os) Then Exit Sub

    PlayMandelbrot os, 40

    CloseShowcase Browser, os
End Sub



'***************************************************************************************************
'                             ■■■ Showcase 02：ライフゲーム ■■■
'***************************************************************************************************
'* 機能　　：グライダー銃とパルサーを置いて、1世代ずつ進めます。生まれたての細胞は緑、古い細胞ほど紫になります
'***************************************************************************************************
Sub Showcase_02_ライフゲーム()
    Dim Browser As CDPBrowser, os As exCDP_OfficeScript
    If Not OpenShowcase(Browser, os) Then Exit Sub

    PlayLife os, 90

    CloseShowcase Browser, os
End Sub



'***************************************************************************************************
'                            ■■■ Showcase 03：ドット絵バナー ■■■
'***************************************************************************************************
'* 機能　　：VBA から渡した文字を5x7ドットで描き、虹色のグラデーションを流します
'---------------------------------------------------------------------------------------------------
'* 注意事項：使える文字は英大文字・数字・空白と「-><!.:?*」です。英小文字は大文字に、それ以外は「?」になります
'***************************************************************************************************
Sub Showcase_03_ドット絵バナー()
    Dim Browser As CDPBrowser, os As exCDP_OfficeScript
    If Not OpenShowcase(Browser, os) Then Exit Sub

    PlayBanner os, "VBA -> CDP", 30

    CloseShowcase Browser, os
End Sub



'***************************************************************************************************
'                         ■■■ Showcase 04：ブックの棚卸し(読み取り) ■■■
'***************************************************************************************************
'* 機能　　：Web 版ブックの全シートを Officeスクリプトで調べ、結果を手元のブックにレポートとグラフで書き出します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：調べる項目は、使用範囲・行数・列数・数式の数・エラーセルの数と場所(先頭5件)・テーブル数・グラフ数です。
'            Officeスクリプトは結果を JSON で console.log に出し、VBA は `LastOutputs` から受け取って読みます
'* 注意事項：・Web 版のブックは書き換えません。Office 側が「ループ内の読み取りは遅い」という注意を問題ログに出しますが、
'              実行には影響しません
'            ・デスクトップ版(SHOWCASE_WEB = False)では、実行を始めたらマクロをいったん終え、`Application.OnTime` で
'              結果を覗きます(Showcase_04_Poll)。デスクトップ版の Officeスクリプトは Excel と同じプロセスで動くため、
'              VBA が待機ループを回している間は、ブックの読み取りが進まないためです
'***************************************************************************************************
Sub Showcase_04_ブック棚卸し()
    If Not m_InspectOS Is Nothing Then
        MsgBox "前回の棚卸しの結果を待っています。終わるまでお待ちください。", vbExclamation
        Exit Sub
    End If

    Dim Browser As CDPBrowser, os As exCDP_OfficeScript
    If Not OpenShowcase(Browser, os) Then Exit Sub

    Dim s As String
    s = Replace(TS_Inspect(), "__LIMIT__", CStr(INSPECT_LIMIT))

    '1. Web 版：同期実行で待つ
    If SHOWCASE_WEB Then
        If RunStep(os, s, "Inspect") Then ShowInspectResult os
        CloseShowcase Browser, os
        Exit Sub
    End If

    '2. デスクトップ版：実行を始めたらマクロを終え、OnTime で結果を覗く
    m_InspectTicket = os.RunScriptAsync(s)
    If m_InspectTicket = 0 Then
        Debug.Print "[Showcase] 棚卸しを開始できませんでした: " & os.LastLogs
        CloseShowcase Browser, os
        Exit Sub
    End If

    Set m_InspectBrowser = Browser
    Set m_InspectOS = os
    m_InspectStart = Timer
    Debug.Print "[Showcase] Inspect started (desktop). ticket: " & m_InspectTicket & " / マクロを終えて、結果を待ちます"
    Application.OnTime Now + TimeSerial(0, 0, INSPECT_POLL_SEC), "Demo_OfficeScript_Showcase.Showcase_04_Poll"
End Sub

'***************************************************************************************************
'* 機能　　：Showcase 04(デスクトップ版)の結果を覗きます(`Application.OnTime`から呼ばれる)
'***************************************************************************************************
Sub Showcase_04_Poll()
    If m_InspectOS Is Nothing Then Exit Sub

    '1. 整理券番号で、1回だけ確認
    Dim Status As String, Elapsed As Double
    If Not m_InspectOS.TryGetResult(m_InspectTicket, Status) Then
        '未完了：上限を過ぎていなければ再予約
        Elapsed = Timer - m_InspectStart
        If Elapsed < 0 Then Elapsed = Elapsed + 86400#   '日付またぎ
        If Elapsed < INSPECT_TIMEOUT_SEC Then
            Application.OnTime Now + TimeSerial(0, 0, INSPECT_POLL_SEC), "Demo_OfficeScript_Showcase.Showcase_04_Poll"
            Exit Sub
        End If
        Status = "Timeout"
    End If

    '2. 結果の書き出し
    Debug.Print "[Showcase] Inspect : " & Status
    If Status = "Succeeded" Then
        ShowInspectResult m_InspectOS
    Else
        Debug.Print m_InspectOS.LastLogs
    End If

    '3. 後片付け(先に Class を手放してから切断)
    CloseShowcase m_InspectBrowser, m_InspectOS
End Sub



'***************************************************************************************************
'                         ■■■ Showcase 05：ライブミラー(読み取り) ■■■
'***************************************************************************************************
'* 機能　　：Web 版でアクティブなシートを繰り返し読み取り、手元のブックの「WebMirror」シートに写します。
'            前回から変わったセルは黄色に光ります
'---------------------------------------------------------------------------------------------------
'* 詳細説明：録画するときは、Web 版 Excel と手元の Excel を並べ、Web 版でセルに入力すると、手元に写る様子を撮ります。
'            読み取るのは、使用範囲の左上から最大 40行 x 12列の「表示されている文字」(getTexts)です
'* 注意事項：・Web 版のブックは書き換えません
'            ・デスクトップ版(SHOWCASE_WEB = False)では、Officeスクリプトの作業ウィンドウを開いているブックを読みます。
'              このマクロのブックで作業ウィンドウを開いていると、WebMirror シート自身を読むことになるので注意してください
'***************************************************************************************************
Sub Showcase_05_ライブミラー()
    Const Polls As Long = 60


    Dim Browser As CDPBrowser, os As exCDP_OfficeScript
    If Not OpenShowcase(Browser, os) Then Exit Sub

    '1. 写す先のシートを用意
    Dim ws As Worksheet
    Set ws = PrepareLocalSheet(MIRROR_SHEET)
    ws.Range("A1").Value = "Web 版 Excel → 手元の Excel (ライブミラー)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Activate

    '2. 読み取り → 書き出し を繰り返す
    Dim s As String, p As Long, Prev As Variant, LastAddress As String
    s = TS_MirrorRead()
    s = Replace(s, "__MAX_ROWS__", CStr(MIRROR_MAX_ROWS))
    s = Replace(s, "__MAX_COLS__", CStr(MIRROR_MAX_COLS))
    For p = 1 To Polls
        If Not RunStep(os, s, "Mirror " & p & "/" & Polls) Then Exit For
        WriteMirror ws, LastJsonOutput(os), p, Polls, Prev, LastAddress
        DoEvents
    Next

    CloseShowcase Browser, os
End Sub



'***************************************************************************************************
'                              ■■■ Showcase 99：後片付け ■■■
'***************************************************************************************************
'* 機能　　：ショーケースで作ったシート(Mandelbrot / LifeGame / Banner)を削除します
'***************************************************************************************************
Sub Showcase_99_後片付け()
    Dim Browser As CDPBrowser, os As exCDP_OfficeScript
    If Not OpenShowcase(Browser, os) Then Exit Sub

    Dim s As String
    s = TS_Cleanup()
    s = Replace(s, "__SHEETS__", """" & MANDEL_SHEET & """, """ & LIFE_SHEET & """, """ & BANNER_SHEET & """")
    RunStep os, s, "Cleanup"

    CloseShowcase Browser, os
End Sub



'***************************************************************************************************
'                                  ■■■ 再生処理 ■■■
'***************************************************************************************************
'* 機能　　：マンデルブロ集合のズームを再生します
'* 返り値　：最後まで再生できたか
'***************************************************************************************************
Private Function PlayMandelbrot(os As exCDP_OfficeScript, Frames As Long) As Boolean
    '1. シートの準備(セルを正方形にし、カラースケールを付ける)
    Dim s As String
    s = TS_MandelSetup()
    s = Replace(s, "__W__", CStr(MANDEL_W))
    s = Replace(s, "__H__", CStr(MANDEL_H))
    s = Replace(s, "__SIZE__", CStr(MANDEL_SIZE))
    s = Replace(s, "__MID__", "85")
    s = Replace(s, "__SHEET__", MANDEL_SHEET)
    If Not RunStep(os, s, "Mandelbrot setup") Then Exit Function

    '2. 1フレームごとに、中心と倍率を差し替えて描く(寄るほど反復回数を増やす)
    Dim f As Long, Scale_ As Double
    Scale_ = 3#
    For f = 0 To Frames - 1
        s = TS_MandelFrame()
        s = Replace(s, "__W__", CStr(MANDEL_W))
        s = Replace(s, "__H__", CStr(MANDEL_H))
        s = Replace(s, "__MAX__", CStr(120 + f * 12))
        s = Replace(s, "__CX__", Num(MANDEL_CX))
        s = Replace(s, "__CY__", Num(MANDEL_CY))
        s = Replace(s, "__SCALE__", Num(Scale_))
        s = Replace(s, "__SHEET__", MANDEL_SHEET)
        If Not RunStep(os, s, "Mandelbrot frame " & f + 1 & "/" & Frames) Then Exit Function
        Scale_ = Scale_ * MANDEL_ZOOM
    Next

    PlayMandelbrot = True
End Function

'***************************************************************************************************
'* 機能　　：ライフゲームを再生します
'* 返り値　：最後まで再生できたか
'***************************************************************************************************
Private Function PlayLife(os As exCDP_OfficeScript, Generations As Long) As Boolean
    '1. シートの準備(グライダー銃とパルサーを置く)
    Dim s As String
    s = TS_LifeSetup()
    s = Replace(s, "__W__", CStr(LIFE_W))
    s = Replace(s, "__H__", CStr(LIFE_H))
    s = Replace(s, "__SIZE__", CStr(LIFE_SIZE))
    s = Replace(s, "__SHEET__", LIFE_SHEET)
    If Not RunStep(os, s, "Life setup") Then Exit Function

    '2. 1世代ずつ進める
    Dim g As Long
    For g = 0 To Generations - 1
        s = TS_LifeStep()
        s = Replace(s, "__W__", CStr(LIFE_W))
        s = Replace(s, "__H__", CStr(LIFE_H))
        s = Replace(s, "__STEPS__", "1")
        s = Replace(s, "__GEN__", CStr(g))
        s = Replace(s, "__SHEET__", LIFE_SHEET)
        If Not RunStep(os, s, "Life gen " & g + 1 & "/" & Generations) Then Exit Function
    Next

    PlayLife = True
End Function

'***************************************************************************************************
'* 機能　　：ドット絵バナーを再生します
'* 返り値　：最後まで再生できたか
'***************************************************************************************************
Private Function PlayBanner(os As exCDP_OfficeScript, ByVal Text As String, Frames As Long) As Boolean
    '1. フォントにない文字を置き換える(スクリプトの文字列にそのまま埋め込むため、ここで必ず絞る)
    Text = SanitizeBannerText(Text)

    '2. シートの準備(文字をドットで描く。長い文字は、横幅が収まるようにセルを小さくする)
    Dim Size As Long
    Size = BANNER_SIZE
    If Len(Text) > 10 Then Size = Int(BANNER_SIZE * 10 / Len(Text))
    If Size < 4 Then Size = 4

    Dim s As String
    s = TS_BannerSetup()
    s = Replace(s, "__TEXT__", Text)
    s = Replace(s, "__SIZE__", CStr(Size))
    s = Replace(s, "__SHEET__", BANNER_SHEET)
    If Not RunStep(os, s, "Banner setup") Then Exit Function

    '3. 1フレームごとに、グラデーションをずらす
    Dim f As Long
    For f = 1 To Frames
        s = TS_BannerFrame()
        s = Replace(s, "__FRAME__", CStr(f))
        s = Replace(s, "__LEN__", CStr(Len(Text)))
        s = Replace(s, "__SHEET__", BANNER_SHEET)
        If Not RunStep(os, s, "Banner frame " & f & "/" & Frames) Then Exit Function
    Next

    PlayBanner = True
End Function



'***************************************************************************************************
'                                  ■■■ 共通処理 ■■■
'***************************************************************************************************
'* 機能　　：CDPBrowser を用意し、コードエディターを確保します(SHOWCASE_WEB で Web 版 / デスクトップ版を切り替え)
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

'***************************************************************************************************
'* 機能　　：後片付け(先に Class を手放してから切断)
'***************************************************************************************************
Private Sub CloseShowcase(Browser As CDPBrowser, os As exCDP_OfficeScript)
    If Not os Is Nothing Then os.Dispose
    Set os = Nothing
    If Not Browser Is Nothing Then Browser.DisconnectCDP
    Set Browser = Nothing
End Sub

'***************************************************************************************************
'* 機能　　：スクリプトを1本実行します。失敗したらログを出して False を返します
'***************************************************************************************************
Private Function RunStep(os As exCDP_OfficeScript, Script As String, Label As String) As Boolean
    Dim Status As String
    Status = os.RunScript(Script, 60)
    Debug.Print "[Showcase] " & Label & " : " & Status
    If Status <> "Succeeded" Then
        Debug.Print os.LastLogs
        If Len(os.LastExceptionMessage) Then Debug.Print "[Showcase] exception: " & os.LastExceptionMessage
        Exit Function
    End If
    RunStep = True
End Function

'***************************************************************************************************
'* 機能　　：Double を、JavaScript の数値リテラルとして使える文字列にします(地域設定によらず小数点は「.」)
'***************************************************************************************************
Private Function Num(Value As Double) As String
    Num = Trim$(Str$(Value))
End Function

'***************************************************************************************************
'* 機能　　：バナーの文字を、フォントにある文字だけに絞ります(英小文字は大文字に、それ以外は「?」に)
'***************************************************************************************************
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
'* 機能　　：直近の実行の console.log のうち、最後の JSON(「{」で始まる行)を返します。無ければ空文字
'***************************************************************************************************
Private Function LastJsonOutput(os As exCDP_OfficeScript) As String
    Dim i As Long, Line As String
    For i = os.LastOutputs.Count To 1 Step -1
        Line = os.LastOutputs(i)
        If Left$(Line, 1) = "{" Then
            LastJsonOutput = Line
            Exit Function
        End If
    Next
End Function

'***************************************************************************************************
'* 機能　　：棚卸しの結果(JSON)を受け取り、手元のブックに書き出します
'***************************************************************************************************
Private Sub ShowInspectResult(os As exCDP_OfficeScript)
    Dim Json As String
    Json = LastJsonOutput(os)
    If Len(Json) Then
        WriteInspectReport Json
    Else
        Debug.Print "[Showcase] 棚卸しの結果(JSON)を受け取れませんでした"
    End If
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
'* 機能　　：棚卸しの JSON を、手元のブックの「WebBookReport」シートにテーブルとグラフで書き出します
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
    ws.Range("A1").Value = "Web 版ブックの棚卸し：" & Root.StringKey("workbook")
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
'* 機能　　：ライブミラーの JSON を、手元のブックの「WebMirror」シートに写します。前回から変わったセルは黄色にします
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
'                     ■■■ Officeスクリプト(TypeScript)のひな形 ■■■
'***************************************************************************************************
'* 機能　　：「__名前__」の部分を、VBA 側で Replace して使います
'***************************************************************************************************
'* マンデルブロ集合：シートの準備
Private Function TS_MandelSetup() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const W = __W__, H = __H__, SIZE = __SIZE__;" & vbLf
    s = s & "    const NAME = ""__SHEET__"";" & vbLf
    s = s & "    let ws = workbook.getWorksheet(NAME);" & vbLf
    s = s & "    if (!ws) { ws = workbook.addWorksheet(NAME); }" & vbLf
    s = s & "    ws.activate();" & vbLf
    s = s & "    ws.getRange().clear(ExcelScript.ClearApplyTo.all);" & vbLf
    s = s & "    ws.setShowGridlines(false);" & vbLf
    s = s & "    const title = ws.getRangeByIndexes(0, 0, 1, W);" & vbLf
    s = s & "    title.merge(false);" & vbLf
    s = s & "    title.getCell(0, 0).setValue(""VBA -> CDP -> Office Scripts : Mandelbrot"");" & vbLf
    s = s & "    title.getFormat().getFont().setBold(true);" & vbLf
    s = s & "    title.getFormat().getFont().setSize(14);" & vbLf
    s = s & "    const grid = ws.getRangeByIndexes(1, 0, H, W);" & vbLf
    s = s & "    grid.getFormat().setColumnWidth(SIZE);" & vbLf
    s = s & "    grid.getFormat().setRowHeight(SIZE);" & vbLf
    s = s & "    grid.setNumberFormat("";;;"");" & vbLf
    s = s & "    grid.getFormat().getFill().setColor(""#000000"");" & vbLf
    s = s & "    const cf = grid.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);" & vbLf
    s = s & "    cf.getColorScale().setCriteria({" & vbLf
    s = s & "        minimum: { type: ExcelScript.ConditionalFormatColorCriterionType.lowestValue, color: ""#020B3A"" }," & vbLf
    s = s & "        midpoint: { type: ExcelScript.ConditionalFormatColorCriterionType.percentile, formula: ""__MID__"", color: ""#FF9A1F"" }," & vbLf
    s = s & "        maximum: { type: ExcelScript.ConditionalFormatColorCriterionType.highestValue, color: ""#FFFFF0"" }" & vbLf
    s = s & "    });" & vbLf
    s = s & "    ws.getRange(""A1"").select();" & vbLf
    s = s & "}" & vbLf
    TS_MandelSetup = s
End Function

'* マンデルブロ集合：1フレーム
Private Function TS_MandelFrame() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const W = __W__, H = __H__, MAX = __MAX__;" & vbLf
    s = s & "    const CX = __CX__, CY = __CY__, SCALE = __SCALE__;" & vbLf
    s = s & "    const ws = workbook.getWorksheet(""__SHEET__"");" & vbLf
    s = s & "    const rows: (number | string)[][] = [];" & vbLf
    s = s & "    for (let y = 0; y < H; y++) {" & vbLf
    s = s & "        const row: (number | string)[] = [];" & vbLf
    s = s & "        const ci = CY + (y / H - 0.5) * SCALE;" & vbLf
    s = s & "        for (let x = 0; x < W; x++) {" & vbLf
    s = s & "            const cr = CX + (x / W - 0.5) * SCALE * (W / H);" & vbLf
    s = s & "            let zr = 0, zi = 0, i = 0;" & vbLf
    s = s & "            while (zr * zr + zi * zi < 16 && i < MAX) {" & vbLf
    s = s & "                const t = zr * zr - zi * zi + cr;" & vbLf
    s = s & "                zi = 2 * zr * zi + ci;" & vbLf
    s = s & "                zr = t;" & vbLf
    s = s & "                i++;" & vbLf
    s = s & "            }" & vbLf
    s = s & "            if (i >= MAX) {" & vbLf
    s = s & "                row.push("""");" & vbLf
    s = s & "            } else {" & vbLf
    s = s & "                const smooth = i + 1 - Math.log(Math.log(Math.sqrt(zr * zr + zi * zi))) / Math.LN2;" & vbLf
    s = s & "                row.push(Math.round(Math.sqrt(smooth) * 1000) / 1000);" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        rows.push(row);" & vbLf
    s = s & "    }" & vbLf
    s = s & "    ws.getRangeByIndexes(1, 0, H, W).setValues(rows);" & vbLf
    s = s & "    ws.getRange(""A1"").setValue(""VBA -> CDP -> Office Scripts : Mandelbrot  x"" + Math.round(3 / SCALE));" & vbLf
    s = s & "}" & vbLf
    TS_MandelFrame = s
End Function

'* ライフゲーム：シートの準備(グライダー銃とパルサー)
Private Function TS_LifeSetup() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const W = __W__, H = __H__, SIZE = __SIZE__;" & vbLf
    s = s & "    const NAME = ""__SHEET__"";" & vbLf
    s = s & "    let ws = workbook.getWorksheet(NAME);" & vbLf
    s = s & "    if (!ws) { ws = workbook.addWorksheet(NAME); }" & vbLf
    s = s & "    ws.activate();" & vbLf
    s = s & "    ws.getRange().clear(ExcelScript.ClearApplyTo.all);" & vbLf
    s = s & "    ws.setShowGridlines(false);" & vbLf
    s = s & "    const title = ws.getRangeByIndexes(0, 0, 1, W);" & vbLf
    s = s & "    title.merge(false);" & vbLf
    s = s & "    title.getCell(0, 0).setValue(""VBA -> CDP -> Office Scripts : Game of Life  gen 0"");" & vbLf
    s = s & "    title.getFormat().getFont().setBold(true);" & vbLf
    s = s & "    title.getFormat().getFont().setSize(14);" & vbLf
    s = s & "    const grid = ws.getRangeByIndexes(1, 0, H, W);" & vbLf
    s = s & "    grid.getFormat().setColumnWidth(SIZE);" & vbLf
    s = s & "    grid.getFormat().setRowHeight(SIZE);" & vbLf
    s = s & "    grid.setNumberFormat("";;;"");" & vbLf
    s = s & "    grid.getFormat().getFill().setColor(""#05050F"");" & vbLf
    s = s & "    const cf = grid.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);" & vbLf
    s = s & "    cf.getColorScale().setCriteria({" & vbLf
    s = s & "        minimum: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""1"", color: ""#39FF14"" }," & vbLf
    s = s & "        midpoint: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""6"", color: ""#00B4FF"" }," & vbLf
    s = s & "        maximum: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""30"", color: ""#B000FF"" }" & vbLf
    s = s & "    });" & vbLf
    s = s & "    const cells: (number | string)[][] = [];" & vbLf
    s = s & "    for (let y = 0; y < H; y++) {" & vbLf
    s = s & "        const row: (number | string)[] = [];" & vbLf
    s = s & "        for (let x = 0; x < W; x++) { row.push(""""); }" & vbLf
    s = s & "        cells.push(row);" & vbLf
    s = s & "    }" & vbLf
    s = s & "    const put = (pattern: string[], top: number, left: number) => {" & vbLf
    s = s & "        for (let y = 0; y < pattern.length; y++) {" & vbLf
    s = s & "            for (let x = 0; x < pattern[y].length; x++) {" & vbLf
    s = s & "                if (pattern[y][x] === ""O"" && top + y < H && left + x < W) { cells[top + y][left + x] = 1; }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    };" & vbLf
    s = s & "    put([" & vbLf
    s = s & "        ""........................O...........""," & vbLf
    s = s & "        ""......................O.O...........""," & vbLf
    s = s & "        ""............OO......OO............OO""," & vbLf
    s = s & "        ""...........O...O....OO............OO""," & vbLf
    s = s & "        ""OO........O.....O...OO..............""," & vbLf
    s = s & "        ""OO........O...O.OO....O.O...........""," & vbLf
    s = s & "        ""..........O.....O.......O...........""," & vbLf
    s = s & "        ""...........O...O....................""," & vbLf
    s = s & "        ""............OO......................""" & vbLf
    s = s & "    ], 2, 2);" & vbLf
    s = s & "    put([" & vbLf
    s = s & "        ""..OOO...OOO..""," & vbLf
    s = s & "        "".............""," & vbLf
    s = s & "        ""O....O.O....O""," & vbLf
    s = s & "        ""O....O.O....O""," & vbLf
    s = s & "        ""O....O.O....O""," & vbLf
    s = s & "        ""..OOO...OOO..""," & vbLf
    s = s & "        "".............""," & vbLf
    s = s & "        ""..OOO...OOO..""," & vbLf
    s = s & "        ""O....O.O....O""," & vbLf
    s = s & "        ""O....O.O....O""," & vbLf
    s = s & "        ""O....O.O....O""," & vbLf
    s = s & "        "".............""," & vbLf
    s = s & "        ""..OOO...OOO..""" & vbLf
    s = s & "    ], 3, W - 16);" & vbLf
    s = s & "    grid.setValues(cells);" & vbLf
    s = s & "    ws.getRange(""A1"").select();" & vbLf
    s = s & "}" & vbLf
    TS_LifeSetup = s
End Function

'* ライフゲーム：STEPS 世代進める
Private Function TS_LifeStep() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const W = __W__, H = __H__, STEPS = __STEPS__, GEN = __GEN__;" & vbLf
    s = s & "    const ws = workbook.getWorksheet(""__SHEET__"");" & vbLf
    s = s & "    const grid = ws.getRangeByIndexes(1, 0, H, W);" & vbLf
    s = s & "    const src = grid.getValues();" & vbLf
    s = s & "    let age: number[][] = src.map(r => r.map(v => (typeof v === ""number"" && v > 0) ? v : 0));" & vbLf
    s = s & "    for (let s = 0; s < STEPS; s++) {" & vbLf
    s = s & "        const next: number[][] = [];" & vbLf
    s = s & "        for (let y = 0; y < H; y++) {" & vbLf
    s = s & "            const row: number[] = [];" & vbLf
    s = s & "            for (let x = 0; x < W; x++) {" & vbLf
    s = s & "                let n = 0;" & vbLf
    s = s & "                for (let dy = -1; dy <= 1; dy++) {" & vbLf
    s = s & "                    for (let dx = -1; dx <= 1; dx++) {" & vbLf
    s = s & "                        if (dx === 0 && dy === 0) { continue; }" & vbLf
    s = s & "                        const yy = y + dy, xx = x + dx;" & vbLf
    s = s & "                        if (yy >= 0 && yy < H && xx >= 0 && xx < W && age[yy][xx] > 0) { n++; }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "                const a = age[y][x];" & vbLf
    s = s & "                row.push(a > 0 ? ((n === 2 || n === 3) ? Math.min(a + 1, 99) : 0) : (n === 3 ? 1 : 0));" & vbLf
    s = s & "            }" & vbLf
    s = s & "            next.push(row);" & vbLf
    s = s & "        }" & vbLf
    s = s & "        age = next;" & vbLf
    s = s & "    }" & vbLf
    s = s & "    grid.setValues(age.map(r => r.map(v => (v > 0 ? v : """"))));" & vbLf
    s = s & "    ws.getRange(""A1"").setValue(""VBA -> CDP -> Office Scripts : Game of Life  gen "" + (GEN + STEPS));" & vbLf
    s = s & "}" & vbLf
    TS_LifeStep = s
End Function

'* ドット絵バナー：シートの準備(5x7ドットのフォント)
Private Function TS_BannerSetup() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const TEXT = ""__TEXT__"", SIZE = __SIZE__, PAD = 2;" & vbLf
    s = s & "    const NAME = ""__SHEET__"";" & vbLf
    s = s & "    const FONT: { [key: string]: string } = {" & vbLf
    s = s & "        ""0"": ""01110100011001110101110011000101110"", ""1"": ""00100011000010000100001000010001110"", ""2"": ""01110100010000100010001000100011111""," & vbLf
    s = s & "        ""3"": ""11111000100010000010000011000101110"", ""4"": ""00010001100101010010111110001000010"", ""5"": ""11111100001111000001000011000101110""," & vbLf
    s = s & "        ""6"": ""00110010001000011110100011000101110"", ""7"": ""11111000010001000100010000100001000"", ""8"": ""01110100011000101110100011000101110""," & vbLf
    s = s & "        ""9"": ""01110100011000101111000010001001100"", ""A"": ""01110100011000111111100011000110001"", ""B"": ""11110100011000111110100011000111110""," & vbLf
    s = s & "        ""C"": ""01110100011000010000100001000101110"", ""D"": ""11110100011000110001100011000111110"", ""E"": ""11111100001000011110100001000011111""," & vbLf
    s = s & "        ""F"": ""11111100001000011110100001000010000"", ""G"": ""01110100011000010111100011000101111"", ""H"": ""10001100011000111111100011000110001""," & vbLf
    s = s & "        ""I"": ""01110001000010000100001000010001110"", ""J"": ""00111000100001000010000101001001100"", ""K"": ""10001100101010011000101001001010001""," & vbLf
    s = s & "        ""L"": ""10000100001000010000100001000011111"", ""M"": ""10001110111010110101100011000110001"", ""N"": ""10001100011100110101100111000110001""," & vbLf
    s = s & "        ""O"": ""01110100011000110001100011000101110"", ""P"": ""11110100011000111110100001000010000"", ""Q"": ""01110100011000110001101011001001101""," & vbLf
    s = s & "        ""R"": ""11110100011000111110101001001010001"", ""S"": ""01111100001000001110000010000111110"", ""T"": ""11111001000010000100001000010000100""," & vbLf
    s = s & "        ""U"": ""10001100011000110001100011000101110"", ""V"": ""10001100011000110001100010101000100"", ""W"": ""10001100011000110101101011010101010""," & vbLf
    s = s & "        ""X"": ""10001100010101000100010101000110001"", ""Y"": ""10001100011000101010001000010000100"", ""Z"": ""11111000010001000100010001000011111""," & vbLf
    s = s & "        "" "": ""00000000000000000000000000000000000"", ""-"": ""00000000000000011111000000000000000"", "">"": ""01000001000001000001000100010001000""," & vbLf
    s = s & "        ""<"": ""00010001000100010000010000010000010"", ""!"": ""00100001000010000100001000000000100"", ""."": ""00000000000000000000000000110001100""," & vbLf
    s = s & "        "":"": ""00000011000110000000011000110000000"", ""?"": ""01110100010000100010001000000000100"", ""*"": ""00000100010101000100010101000100000""" & vbLf
    s = s & "    };" & vbLf
    s = s & "    const W = TEXT.length * 6 - 1 + PAD * 2, H = 7 + PAD * 2;" & vbLf
    s = s & "    let ws = workbook.getWorksheet(NAME);" & vbLf
    s = s & "    if (!ws) { ws = workbook.addWorksheet(NAME); }" & vbLf
    s = s & "    ws.activate();" & vbLf
    s = s & "    ws.getRange().clear(ExcelScript.ClearApplyTo.all);" & vbLf
    s = s & "    ws.setShowGridlines(false);" & vbLf
    s = s & "    const area = ws.getRangeByIndexes(0, 0, H, W);" & vbLf
    s = s & "    area.getFormat().setColumnWidth(SIZE);" & vbLf
    s = s & "    area.getFormat().setRowHeight(SIZE);" & vbLf
    s = s & "    area.setNumberFormat("";;;"");" & vbLf
    s = s & "    area.getFormat().getFill().setColor(""#000000"");" & vbLf
    s = s & "    const cf = area.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);" & vbLf
    s = s & "    cf.getColorScale().setCriteria({" & vbLf
    s = s & "        minimum: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""0"", color: ""#FF2E88"" }," & vbLf
    s = s & "        midpoint: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""50"", color: ""#00E5FF"" }," & vbLf
    s = s & "        maximum: { type: ExcelScript.ConditionalFormatColorCriterionType.number, formula: ""100"", color: ""#FFE600"" }" & vbLf
    s = s & "    });" & vbLf
    s = s & "    const cells: (number | string)[][] = [];" & vbLf
    s = s & "    for (let y = 0; y < H; y++) {" & vbLf
    s = s & "        const row: (number | string)[] = [];" & vbLf
    s = s & "        for (let x = 0; x < W; x++) { row.push(""""); }" & vbLf
    s = s & "        cells.push(row);" & vbLf
    s = s & "    }" & vbLf
    s = s & "    for (let c = 0; c < TEXT.length; c++) {" & vbLf
    s = s & "        const glyph = FONT[TEXT[c]] || FONT[""?""];" & vbLf
    s = s & "        for (let y = 0; y < 7; y++) {" & vbLf
    s = s & "            for (let x = 0; x < 5; x++) {" & vbLf
    s = s & "                if (glyph[y * 5 + x] === ""1"") {" & vbLf
    s = s & "                    const gx = PAD + c * 6 + x, gy = PAD + y;" & vbLf
    s = s & "                    const p = (gx * 6 + gy * 3) % 200;" & vbLf
    s = s & "                    cells[gy][gx] = p <= 100 ? p : 200 - p;" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    area.setValues(cells);" & vbLf
    s = s & "    const caption = ws.getRangeByIndexes(H, 0, 1, W);" & vbLf
    s = s & "    caption.merge(false);" & vbLf
    s = s & "    caption.getCell(0, 0).setValue(""Drawn by Office Scripts, driven from VBA via CDP"");" & vbLf
    s = s & "    caption.getFormat().setRowHeight(20);" & vbLf
    s = s & "    caption.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);" & vbLf
    s = s & "    caption.getFormat().getFont().setColor(""#888888"");" & vbLf
    s = s & "    ws.getRange(""A1"").select();" & vbLf
    s = s & "}" & vbLf
    TS_BannerSetup = s
End Function

'* ドット絵バナー：グラデーションを1フレーム進める
Private Function TS_BannerFrame() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const FRAME = __FRAME__, LEN = __LEN__, PAD = 2;" & vbLf
    s = s & "    const W = LEN * 6 - 1 + PAD * 2, H = 7 + PAD * 2;" & vbLf
    s = s & "    const area = workbook.getWorksheet(""__SHEET__"").getRangeByIndexes(0, 0, H, W);" & vbLf
    s = s & "    const src = area.getValues();" & vbLf
    s = s & "    const out: (number | string)[][] = src.map((r, y) => r.map((v, x) => {" & vbLf
    s = s & "        if (typeof v !== ""number"") { return """"; }" & vbLf
    s = s & "        const p = (x * 6 + y * 3 + FRAME * 12) % 200;" & vbLf
    s = s & "        return p <= 100 ? p : 200 - p;" & vbLf
    s = s & "    }));" & vbLf
    s = s & "    area.setValues(out);" & vbLf
    s = s & "}" & vbLf
    TS_BannerFrame = s
End Function

'* ブックの棚卸し：全シートを調べて JSON で返す
Private Function TS_Inspect() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const LIMIT = __LIMIT__, MAX_ERR = 5;" & vbLf
    s = s & "    const col = (n: number): string => { let s = """"; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };" & vbLf
    s = s & "    const sheets: { name: string; used: string; rows: number; cols: number; formulas: number; errors: number; errorCells: string[]; tables: number; charts: number; scanned: boolean }[] = [];" & vbLf
    s = s & "    for (const ws of workbook.getWorksheets()) {" & vbLf
    s = s & "        const info = { name: ws.getName(), used: """", rows: 0, cols: 0, formulas: 0, errors: 0, errorCells: [] as string[], tables: ws.getTables().length, charts: ws.getCharts().length, scanned: false };" & vbLf
    s = s & "        const used = ws.getUsedRange(true);" & vbLf
    s = s & "        if (used) {" & vbLf
    s = s & "            info.used = used.getAddress().split(""!"").pop() || """";" & vbLf
    s = s & "            info.rows = used.getRowCount();" & vbLf
    s = s & "            info.cols = used.getColumnCount();" & vbLf
    s = s & "            if (info.rows * info.cols <= LIMIT) {" & vbLf
    s = s & "                info.scanned = true;" & vbLf
    s = s & "                const top = used.getRowIndex(), left = used.getColumnIndex();" & vbLf
    s = s & "                const types = used.getValueTypes();" & vbLf
    s = s & "                const formulas = used.getFormulas();" & vbLf
    s = s & "                for (let r = 0; r < info.rows; r++) {" & vbLf
    s = s & "                    for (let c = 0; c < info.cols; c++) {" & vbLf
    s = s & "                        const f = formulas[r][c];" & vbLf
    s = s & "                        if (typeof f === ""string"" && f.startsWith(""="")) { info.formulas++; }" & vbLf
    s = s & "                        if (types[r][c] === ExcelScript.RangeValueType.error) {" & vbLf
    s = s & "                            info.errors++;" & vbLf
    s = s & "                            if (info.errorCells.length < MAX_ERR) { info.errorCells.push(col(left + c) + (top + r + 1)); }" & vbLf
    s = s & "                        }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        sheets.push(info);" & vbLf
    s = s & "    }" & vbLf
    s = s & "    console.log(JSON.stringify({ workbook: workbook.getName(), names: workbook.getNames().length, sheets: sheets }));" & vbLf
    s = s & "}" & vbLf
    TS_Inspect = s
End Function

'* ライブミラー：アクティブなシートの表示テキストを JSON で返す
Private Function TS_MirrorRead() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const MAX_ROWS = __MAX_ROWS__, MAX_COLS = __MAX_COLS__;" & vbLf
    s = s & "    const ws = workbook.getActiveWorksheet();" & vbLf
    s = s & "    const used = ws.getUsedRange(true);" & vbLf
    s = s & "    let texts: string[][] = [];" & vbLf
    s = s & "    let address = """";" & vbLf
    s = s & "    if (used) {" & vbLf
    s = s & "        const rows = Math.min(used.getRowCount(), MAX_ROWS), cols = Math.min(used.getColumnCount(), MAX_COLS);" & vbLf
    s = s & "        const area = ws.getRangeByIndexes(used.getRowIndex(), used.getColumnIndex(), rows, cols);" & vbLf
    s = s & "        address = area.getAddress().split(""!"").pop() || """";" & vbLf
    s = s & "        texts = area.getTexts();" & vbLf
    s = s & "    }" & vbLf
    s = s & "    console.log(JSON.stringify({ sheet: ws.getName(), address: address, texts: texts }));" & vbLf
    s = s & "}" & vbLf
    TS_MirrorRead = s
End Function

'* 後片付け：ショーケースのシートを削除
Private Function TS_Cleanup() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const NAMES: string[] = [__SHEETS__];" & vbLf
    s = s & "    for (const n of NAMES) {" & vbLf
    s = s & "        const ws = workbook.getWorksheet(n);" & vbLf
    s = s & "        if (ws) {" & vbLf
    s = s & "            if (workbook.getWorksheets().length <= 1) { workbook.addWorksheet(); }" & vbLf
    s = s & "            ws.delete();" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "}" & vbLf
    TS_Cleanup = s
End Function
