VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_PDFPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           `CDPBrowser.cls` に対する機能拡張 - PDF保存（Page.printToPDF）
'***************************************************************************************************
'* 機能　　：現在ブラウザが表示しているページをPDFファイルとして保存します
'---------------------------------------------------------------------------------------------------
'* 使い方　：
'    Dim pdf As New exCDP_PDFPrinter
'    pdf.Init browser                              ' CDPBrowserを継承
'    pdf.PrintToPDF "C:\output\", "report"         ' → report.pdf が生成される
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - `Page.printToPDF` は about:blank など印刷できないページでは失敗します
'    - CDPコマンドが非同期で動くため、ページ読み込み完了後に呼び出してください
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "exCDP_PDFPrinter"   'ログのトレース用



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分編 ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：現在のページをPDFとして指定フォルダに保存します
'---------------------------------------------------------------------------------------------------
'* 返り値　：保存完了したファイルのフルパス。失敗時は空文字を返します
'* 引数　　：FolderPath       保存先フォルダのパス（末尾の「\」は有無どちらでも可）
'            FileBaseName     拡張子なしのファイル名（例: "report" → "report.pdf"）
'            PrintBackground  背景色・背景画像をPDFに含めるか（デフォルト: True）
'            Landscape        横向き印刷にするか（デフォルト: False = 縦向き）
'            PaperWidth       用紙幅（インチ単位）。0指定でA4縦（8.27インチ）
'            PaperHeight      用紙高さ（インチ単位）。0指定でA4縦（11.69インチ）
'            Scale_            スケール倍率（デフォルト: 1.0）
'            MarginTop        上余白（インチ、デフォルト: 0.4）
'            MarginBottom     下余白（インチ、デフォルト: 0.4）
'            MarginLeft       左余白（インチ、デフォルト: 0.4）
'            MarginRight      右余白（インチ、デフォルト: 0.4）
'***************************************************************************************************
Public Function PrintToPDF( _
        FolderPath As String, _
        FileBaseName As String, _
        Optional PrintBackground As Boolean = True, _
        Optional Landscape As Boolean = False, _
        Optional PaperWidth As Double = 0, _
        Optional PaperHeight As Double = 0, _
        Optional Scale_ As Double = 1#, _
        Optional MarginTop As Double = 0.4, _
        Optional MarginBottom As Double = 0.4, _
        Optional MarginLeft As Double = 0.4, _
        Optional MarginRight As Double = 0.4 _
    ) As String
    Const FromProcedureName As String = ThisClassName & ".PrintToPDF"


    printMsg info_, "PDF printing started. File: " & FileBaseName & ".pdf", FromProcedureName

    '--- 1. パラメーター組み立て ---
    Dim params As New Scripting.Dictionary

    params("printBackground") = PrintBackground
    params("landscape") = Landscape
    params("scale") = Scale_
    params("marginTop") = MarginTop
    params("marginBottom") = MarginBottom
    params("marginLeft") = MarginLeft
    params("marginRight") = MarginRight

    '用紙サイズ（0指定はA4デフォルト）
    Dim w As Double: w = IIf(PaperWidth > 0, PaperWidth, 8.27)     'A4横幅
    Dim h As Double: h = IIf(PaperHeight > 0, PaperHeight, 11.69)  'A4縦幅
    params("paperWidth") = w
    params("paperHeight") = h

    printMsg Debug_, "PDF params - landscape:" & Landscape & " printBackground:" & PrintBackground _
                   & " paperWidth:" & w & " paperHeight:" & h & " Scale:" & Scale_, FromProcedureName

    '--- 2. CDP コマンド発行 ---
    printMsg Trace_, "Sending `Page.printToPDF` command...", FromProcedureName
    Dim result As Scripting.Dictionary
    Set result = invokeMethod("Page.printToPDF", params, , False)  'StopError:=False でエラー時もNothingで返す

    '--- 3. エラー確認 ---
    If result Is Nothing Then
        printMsg WARN_, "Page.printToPDF returned Nothing. The page may not be printable (e.g. about:blank).", FromProcedureName
        PrintToPDF = ""
        Exit Function
    End If

    If Not result.Exists("data") Then
        printMsg WARN_, "Response does not contain `data` key. Keys returned: " & Join(result.keys, ", "), FromProcedureName
        PrintToPDF = ""
        Exit Function
    End If

    '--- 4. Base64デコード → ファイル保存 ---
    Dim b64 As String: b64 = CStr(result("data"))
    If Len(b64) = 0 Then
        printMsg WARN_, "PDF data (Base64) is empty. Nothing to save.", FromProcedureName
        PrintToPDF = ""
        Exit Function
    End If

    printMsg Debug_, "PDF Base64 data received. Length: " & Len(b64) & " chars", FromProcedureName

    '保存先フォルダの末尾「\」を統一して除去
    Dim folderNorm As String: folderNorm = FolderPath
    If Right(folderNorm, 1) = "\" Then folderNorm = Left(folderNorm, Len(folderNorm) - 1)

    'フォルダが存在しない場合は作成
    If Dir(folderNorm, vbDirectory) = "" Then
        printMsg info_, "Creating output folder: " & folderNorm, FromProcedureName
        MkDir folderNorm
    End If

    'ファイル名組み立て（拡張子 .pdf を保証）
    Dim FileName As String: FileName = FileBaseName
    If LCase(Right(FileName, 4)) <> ".pdf" Then FileName = FileName & ".pdf"

    'Base64 → バイナリ変換（MSXML2 DOM経由 ／ CDPBrowser.SaveBase64ToFileと同方式）
    Dim dom As Object, el As Object, Bytes() As Byte
    Set dom = CreateObject("MSXML2.DOMDocument")
    Set el = dom.createElement("pdfnode")
    el.DataType = "bin.base64"
    el.Text = b64
    Bytes = el.nodeTypedValue

    'CharacterCodeConversion 経由でファイル書き出し
    Dim CharConv As New CharacterCodeConversion
    CharConv.BytesToSaveFile Bytes, folderNorm, FileName

    Dim fullPath As String: fullPath = folderNorm & "\" & FileName
    printMsg info_, "PDF saved successfully: " & fullPath, FromProcedureName

    PrintToPDF = fullPath

End Function


'***************************************************************************************************
'* 機能　　：用紙サイズのプリセットを文字列で指定できる上位互換メソッドです
'---------------------------------------------------------------------------------------------------
'* 引数　　：PaperPreset    "A4" / "A3" / "Letter" / "Legal" のいずれか（大文字小文字不問）
'            その他の引数は PrintToPDF と同じです
'***************************************************************************************************
Public Function PrintToPDFWithPreset( _
        FolderPath As String, _
        FileBaseName As String, _
        Optional PaperPreset As String = "A4", _
        Optional PrintBackground As Boolean = True, _
        Optional Landscape As Boolean = False, _
        Optional Scale_ As Double = 1# _
    ) As String
    Const FromProcedureName As String = ThisClassName & ".PrintToPDFWithPreset"


    Dim w As Double, h As Double

    '用紙サイズ（単位：インチ）
    Select Case UCase(Trim(PaperPreset))
        Case "A4":      w = 8.27:  h = 11.69
        Case "A3":      w = 11.69: h = 16.54
        Case "LETTER":  w = 8.5:   h = 11#
        Case "LEGAL":   w = 8.5:   h = 14#
        Case Else
            printMsg WARN_, "Unknown paper preset: """ & PaperPreset & """. Falling back to A4.", FromProcedureName
            w = 8.27: h = 11.69
    End Select

    printMsg Debug_, "Paper preset: " & UCase(Trim(PaperPreset)) & " (" & w & " x " & h & " inches)", FromProcedureName

    'Landscape の場合は縦横を入れ替え
    If Landscape Then
        Dim tmp As Double: tmp = w: w = h: h = tmp
        printMsg Debug_, "Landscape mode: swapped to " & w & " x " & h & " inches", FromProcedureName
    End If

    PrintToPDFWithPreset = PrintToPDF(FolderPath, FileBaseName, PrintBackground, Landscape, w, h, Scale_)

End Function



'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_CDPBrowser.sleep seconds
End Sub

Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set invokeMethod = ex_CDPBrowser.invokeMethod(methodName, params, alwaysBrowserContext, StopError)
End Function



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
    '現バージョンでは設定シートからの読み込みは不要
    '将来的に保存先デフォルトフォルダ等を設定シートから取得する場合はここに追加
End Sub
