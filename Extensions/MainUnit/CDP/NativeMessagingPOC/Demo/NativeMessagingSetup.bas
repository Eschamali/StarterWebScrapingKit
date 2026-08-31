Attribute VB_Name = "NativeMessagingSetup"
'***************************************************************************************************
'   Chrome拡張機能の`chrome.debugger`×`Native Messaging`経由でCDP制御する機能のための、
'   Native Messagingホストマニフェスト(.json)の生成ヘルパーです。
'---------------------------------------------------------------------------------------------------
'   Chromeの`Native Messaging`は、拡張機能が`chrome.runtime.connectNative`した瞬間に、
'   OS側(レジストリ)に登録されたホスト実行ファイルをChrome自身が起動する仕様のため、
'   「ホストとして何を起動させるか("path")」を2択で選べるようにしてあります。
'
'     A. NM_DirectExcel  : EXCEL.EXE自身を直接ホストとして登録
'                          (Chromeが渡す引数をEXCEL.EXEがファイルパスとして誤解釈し、
'                           エラーダイアログが出る可能性・DDEシングルインスタンス動作で
'                           新規プロセスが即終了する可能性があります。事前に検証してください)
'
'     B. NM_WrapperBat   : 動的生成した.batファイルをホストとして登録し、そこから
'                          `EXCEL.EXE /x "対象ワークブック"`を起動する
'                          (`/x`で強制的に新規インスタンス化するため、DDE問題を回避しやすく、
'                           Chromeが渡す不正な引数もbat側で握りつぶすため、ダイアログも出ません)
'
'   レジストリへの登録作業自体は、このマクロでは行いません(絶対禁止事項:管理者権限操作の回避、
'   および意図しないレジストリ変更を防ぐため)。生成したファイルを、手順書に従って手動で登録してください。
'***************************************************************************************************
Option Explicit
Option Private Module

Public Enum NMHostMode
    NM_DirectExcel = 0  'A. EXCEL.EXE自身を直接ホストにする
    NM_WrapperBat = 1   'B. ラッパー.bat経由でEXCEL.EXEを`/x`起動する
End Enum

'レジストリキーの案内文言切り替え用。マニフェストJSONの中身自体は、どちらのブラウザでも同一形式(`allowed_origins`の
'書式も"chrome-extension://<id>/"のまま)。Chromium系ブラウザなら共通で、レジストリの登録先だけが変わる
Public Enum NMBrowserTarget
    NM_Chrome = 0
    NM_Edge = 1
End Enum

'Native MessagingホストのID。拡張機能IDと組み合わせて、レジストリ/allowed_originsに使う
Public Const HOST_NAME As String = "com.starterwebscrapingkit.cdpbridge"



'***************************************************************************************************
'* 機能　　：Native Messagingホストマニフェスト(.json)、および必要ならラッパー.batを生成します
'---------------------------------------------------------------------------------------------------
'* 引数    ：Mode           `NM_DirectExcel` or `NM_WrapperBat`
'            TargetBrowser  `NM_Chrome` or `NM_Edge` ※案内するレジストリキーの表示切り替えのみに使用
'            OutputFolder   生成先フォルダ(存在しない場合は作成します)
'            ExtensionID    `chrome://extensions`または`edge://extensions`で確認した、
'                           このNative Messagingを呼び出す拡張機能のID
'---------------------------------------------------------------------------------------------------
'* 機能説明：生成されるのは、あくまでこのローカルフォルダへの「ファイル生成」のみです。
'            Windowsレジストリへの登録は、手順書に従って手動で行ってください
'* 注意事項：マニフェストJSONの中身自体(`allowed_origins`の書式含む)は、Chrome/Edgeどちらでも共通です。
'            変わるのは、手動登録するレジストリキーの場所(`Google\Chrome` or `Microsoft\Edge`)のみです
'***************************************************************************************************
Public Sub GenerateNativeMessagingHostManifest(Mode As NMHostMode, TargetBrowser As NMBrowserTarget, OutputFolder As String, ExtensionID As String)
    Const FromProcedureName As String = "NativeMessagingSetup.GenerateNativeMessagingHostManifest"


    '1. 引数チェック
    If Len(ExtensionID) = 0 Then Err.Raise vbObjectError + 1, FromProcedureName, "ExtensionIDが指定されていません。`chrome://extensions`で拡張機能IDを確認してください。"

    '2. 出力フォルダを準備
    If Right$(OutputFolder, 1) = "\" Then OutputFolder = Left$(OutputFolder, Len(OutputFolder) - 1)
    EnsureFolderExists OutputFolder

    '3. ホストとして起動する実行ファイルパスを決定
    Dim ExcelPath As String: ExcelPath = Application.Path & "\EXCEL.EXE"
    Dim HostPath As String

    Select Case Mode
        Case NM_WrapperBat
            '3-1. ラッパー.batを生成 ※`/x`で強制的に新規インスタンス化させ、Chromeが渡す引数は握りつぶす
            Dim BatPath As String: BatPath = OutputFolder & "\launch_excel_host.bat"
            Dim BatContent As String
            BatContent = "@echo off" & vbCrLf & _
                         Chr(34) & ExcelPath & Chr(34) & " /x " & Chr(34) & ThisWorkbook.FullName & Chr(34)
            WriteTextFile BatPath, BatContent
            HostPath = BatPath
            Debug.Print "ラッパー.batを生成しました: " & BatPath

        Case Else 'NM_DirectExcel
            HostPath = ExcelPath
    End Select

    '4. ホストマニフェストJSONを生成
    Dim ManifestJson As String
    ManifestJson = "{" & vbCrLf & _
        "  " & Chr(34) & "name" & Chr(34) & ": " & Chr(34) & HOST_NAME & Chr(34) & "," & vbCrLf & _
        "  " & Chr(34) & "description" & Chr(34) & ": " & Chr(34) & "StarterWebScrapingKit CDP bridge host" & Chr(34) & "," & vbCrLf & _
        "  " & Chr(34) & "path" & Chr(34) & ": " & Chr(34) & JsonEscapePath(HostPath) & Chr(34) & "," & vbCrLf & _
        "  " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "stdio" & Chr(34) & "," & vbCrLf & _
        "  " & Chr(34) & "allowed_origins" & Chr(34) & ": [" & Chr(34) & "chrome-extension://" & ExtensionID & "/" & Chr(34) & "]" & vbCrLf & _
        "}"

    Dim ManifestPath As String: ManifestPath = OutputFolder & "\" & HOST_NAME & ".json"
    WriteTextFile ManifestPath, ManifestJson

    '5. 案内 ※ブラウザによってレジストリの登録先パスが異なる(マニフェストJSONの中身自体は共通)
    Dim RegistryKeyPath As String
    Select Case TargetBrowser
        Case NM_Edge
            RegistryKeyPath = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Edge\NativeMessagingHosts\" & HOST_NAME
        Case Else 'NM_Chrome
            RegistryKeyPath = "HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\NativeMessagingHosts\" & HOST_NAME
    End Select

    Debug.Print "-----------------------------------------------------------"
    Debug.Print "Native Messagingホストマニフェストを生成しました:"
    Debug.Print "  " & ManifestPath
    Debug.Print "ホスト実行ファイル(path):"
    Debug.Print "  " & HostPath
    Debug.Print "この後、手順書(NativeMessagingBridge\手順書.md)に従って、"
    Debug.Print "レジストリキー " & RegistryKeyPath
    Debug.Print "の(既定)値に、上記マニフェストJSONのフルパスを手動で設定してください。"
    Debug.Print "-----------------------------------------------------------"
End Sub

'***************************************************************************************************
'* 機能　　：JSON文字列内で、Windowsパスのバックスラッシュをエスケープします
'***************************************************************************************************
Private Function JsonEscapePath(FilePath As String) As String
    JsonEscapePath = Replace(FilePath, "\", "\\")
End Function

'***************************************************************************************************
'* 機能　　：フォルダを作成します(`MkDir`は途中の階層が無いと失敗するため、1階層ずつ辿って作成します)
'***************************************************************************************************
Private Sub EnsureFolderExists(FolderPath As String)
    Dim Parts() As String: Parts = Split(FolderPath, "\")

    Dim BuildPath As String: BuildPath = Parts(0)  'ドライブレター部分(例："C:")
    Dim i As Long
    For i = 1 To UBound(Parts)
        BuildPath = BuildPath & "\" & Parts(i)
        If Len(Dir(BuildPath, vbDirectory)) = 0 Then MkDir BuildPath
    Next i
End Sub

'***************************************************************************************************
'* 機能　　：テキストファイルをUTF-8(BOMなし)で新規書き出しします(既存があれば上書き)
'---------------------------------------------------------------------------------------------------
'* 注意事項：・ChromeのNative Messagingホストマニフェストは、UTF-8であることが求められます
'            ・`.bat`ファイルは、先頭にBOMが付くと`cmd.exe`が1行目を正しく解釈できない場合があるため、
'            　通常の`Print #`(ANSI/システムのコードページ)ではなく、UTF-8バイト配列を直接書き出します
'***************************************************************************************************
Private Sub WriteTextFile(FilePath As String, Content As String)
    Dim Utf8Converter As New CharacterCodeConversion
    Dim BinData() As Byte: BinData = Utf8Converter.BytesFromString(Content)

    'Binaryモードは既存ファイルを自動で切り詰めないため、古い内容が末尾に残らないよう先に削除しておく
    If Len(Dir(FilePath)) > 0 Then Kill FilePath

    Dim f As Integer: f = FreeFile
    Open FilePath For Binary Access Write As #f
    Put #f, , BinData
    Close #f
End Sub
