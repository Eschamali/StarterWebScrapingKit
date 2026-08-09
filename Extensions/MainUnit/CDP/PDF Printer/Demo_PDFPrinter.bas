Attribute VB_Name = "Demo_PDFPrinter"
'***************************************************************************************************
'       exCDP_PDFPrinter 拡張 - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：`exCDP_PDFPrinter.cls` を使ったPDF保存のサンプルコードです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\CDP\PDF Printer\exCDP_PDFPrinter.cls
'* 参考元  ：ForAI\vba-cdp-webdriver\Module\SampleModule.bas - Sample_11_Screenshot_And_Pdf
'---------------------------------------------------------------------------------------------------
'* 注意事項：・実行前に「exCDP_PDFPrinter.cls」をVBAプロジェクトに取り込んでください
'            ・`Page.printToPDF` は about:blank では動作しません
'            ・ページ読み込みが完了してから呼び出してください
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                            ■■■ Demo 01：基本的なPDF保存 ■■■
'***************************************************************************************************
'* 機能　　：最もシンプルなPDF保存のデモです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：example.com をブラウザで開き、Downloadsフォルダへ PDF として保存します
'* 確認ポイント：
'   - `PrintToPDF` が保存先フルパスを返すこと
'   - Downloadsフォルダに "demo_basic.pdf" が生成されること
'***************************************************************************************************
Sub Demo_PDFPrinter_01_基本保存()

    '1. ブラウザ起動
    Dim browserTab As CDPContext: Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("https://example.com")

    '2. PDF拡張の初期化
    Dim pdf As New exCDP_PDFPrinter
    pdf.Init browserTab

    '3. PDF保存（デフォルト設定 = A4縦、背景あり）
    Dim outDir As String: outDir = Environ("UserProfile") & "\Downloads"
    Dim savedPath As String
    savedPath = pdf.PrintToPDF(outDir, "demo_basic")

    '4. 結果確認
    If savedPath <> "" Then
        browserTab.notify "PDF保存完了！: " & savedPath
    Else
        MsgBox "PDF保存に失敗しました。イミディエイトウィンドウを確認してください。", vbCritical, "Error"
    End If

    '5. ブラウザを閉じる
    browserTab.InheritanceCDPBrowser.quit

End Sub



'***************************************************************************************************
'                         ■■■ Demo 02：A4横・背景なし・スケール変更 ■■■
'***************************************************************************************************
'* 機能　　：パラメーターをフル指定したPDF保存のデモです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：各パラメーターの効果を確認するためのサンプルです
'* 確認ポイント：
'   - Landscape:=True で横向きPDFが生成されること
'   - PrintBackground:=False で背景色が消えること
'   - Scale:=0.8 でコンテンツが縮小されること
'***************************************************************************************************
Sub Demo_PDFPrinter_02_パラメーター指定()

    '1. ブラウザ起動
    Dim browserTab As CDPContext: Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("https://www.wikipedia.org")

    '2. PDF拡張の初期化
    Dim pdf As New exCDP_PDFPrinter
    pdf.Init browserTab

    '3. パラメーター指定PDF保存
    Dim outDir As String: outDir = Environ("UserProfile") & "\Downloads"
    Dim savedPath As String
    
    'PaperWidth:=11.69　→　A4横
    savedPath = pdf.PrintToPDF( _
        FolderPath:=outDir, _
        FileBaseName:="demo_landscape", _
        PrintBackground:=False, _
        Landscape:=True, _
        PaperWidth:=11.69, _
        PaperHeight:=8.27, _
        Scale_:=0.8, _
        MarginTop:=0.5, _
        MarginBottom:=0.5, _
        MarginLeft:=0.5, _
        MarginRight:=0.5 _
    )

    '4. 結果確認
    Debug.Print "保存パス: " & savedPath
    If savedPath <> "" Then browserTab.notify "パラメーター指定PDF保存完了！: " & savedPath

    '5. ブラウザを閉じる
    browserTab.InheritanceCDPBrowser.quit

End Sub



'***************************************************************************************************
'                      ■■■ Demo 03：プリセット指定（A4/A3/Letter） ■■■
'***************************************************************************************************
'* 機能　　：用紙サイズをプリセット名で指定するデモです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`PrintToPDFWithPreset` を使うと、用紙サイズを文字列で直感的に指定できます
'* 確認ポイント：
'   - "A4" / "A3" / "Letter" / "Legal" のプリセットが正しく機能すること
'   - 未対応プリセットはA4にフォールバックされること（イミディエイトに警告ログが出る）
'***************************************************************************************************
Sub Demo_PDFPrinter_03_プリセット指定()

    '1. ブラウザ起動
    Dim browserTab As CDPContext: Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("https://www.wikipedia.org")

    '2. PDF拡張の初期化
    Dim pdf As New exCDP_PDFPrinter
    pdf.Init browserTab

    Dim outDir As String: outDir = Environ("UserProfile") & "\Downloads"

    '3a. A4縦（デフォルト）
    Debug.Print "A4縦: " & pdf.PrintToPDFWithPreset(outDir, "demo_preset_A4", PaperPreset:="A4")

    '3b. A3縦
    Debug.Print "A3縦: " & pdf.PrintToPDFWithPreset(outDir, "demo_preset_A3", PaperPreset:="A3")

    '3c. Letter横
    Debug.Print "Letter横: " & pdf.PrintToPDFWithPreset(outDir, "demo_preset_Letter", PaperPreset:="Letter", Landscape:=True)

    '3d. 未対応プリセット → A4フォールバック（WARN_ ログが出る）
    Debug.Print "未対応: " & pdf.PrintToPDFWithPreset(outDir, "demo_preset_unknown", PaperPreset:="B5")

    browserTab.notify "プリセット指定PDF保存を3種類完了しました！"

    '4. ブラウザを閉じる
    browserTab.InheritanceCDPBrowser.quit

End Sub



'***************************************************************************************************
'                   ■■■ Demo 04：スクリーンショット + PDF の同時保存 ■■■
'***************************************************************************************************
'* 機能　　：同一ページをPNGスクリーンショットとPDFの両方で保存するデモです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`CDPbrowserTab.snapPage`（スクショ）と `exCDP_PDFPrinter.PrintToPDF`（PDF）を
'            組み合わせて使用します。ForAI\vba-cdp-webdriver の Sample_11 に相当する使い方です
'* 確認ポイント：
'   - ScreenShotとPDFが同じディレクトリに保存されること
'   - どちらも同じページ内容が反映されていること
'***************************************************************************************************
Sub Demo_PDFPrinter_04_スクショとPDF同時保存()

    '保存先フォルダ（なければ自動作成）
    Dim outDir As String: outDir = Environ("UserProfile") & "\Downloads\CDPcapture"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    '1. ブラウザ起動（Googleの検索結果ページ）
    Dim browserTab As CDPContext: Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("https://www.google.com/search?q=1USD+to+JPY")

    '2. スクリーンショット（CDPbrowserTab.snapPage）
    browserTab.snapPage outDir, "capture_shot.png"
    Debug.Print "スクショ保存: " & outDir & "\capture_shot.png"

    '3. PDF保存（exCDP_PDFPrinter）
    Dim pdf As New exCDP_PDFPrinter
    pdf.Init browserTab
    Dim pdfPath As String: pdfPath = pdf.PrintToPDF(outDir, "capture_pdf")
    Debug.Print "PDF保存: " & pdfPath

    '4. 結果通知
    If pdfPath <> "" Then
        browserTab.notify "スクショ & PDF を同時保存しました！" & vbCrLf & outDir
    Else
        MsgBox "PDF保存に失敗しました。", vbCritical
    End If

    '5. ブラウザを閉じる
    browserTab.InheritanceCDPBrowser.quit

End Sub
