Attribute VB_Name = "Demo_NativeMessaging"
'***************************************************************************************************
'          Chrome拡張機能の`chrome.debugger`×`Native Messaging`経由でCDP制御するデモです
'---------------------------------------------------------------------------------------------------
'   ＜セットアップの流れ＞
'   1. 対象ブラウザで「デベロッパーモード」を有効化し、
'      「NativeMessagingBridge\Extension」フォルダをパッケージ化されていない拡張機能として読み込む
'      (Chromeなら`chrome://extensions`、Edgeなら`edge://extensions`)
'      → 表示された拡張機能IDをメモしておく
'
'   2. 対象ブラウザ・方式に合わせて、下記の生成マクロのいずれかを実行し、ホストマニフェスト(.json)を
'      生成する ※`NativeMessagingBridge\手順書.md`も参照
'        - `GenerateNativeHostFiles_Chrome_直接Excel` / `GenerateNativeHostFiles_Chrome_ラッパーBAT経由`
'        - `GenerateNativeHostFiles_Edge_直接Excel`   / `GenerateNativeHostFiles_Edge_ラッパーBAT経由`
'
'   3. 手順書に従って、Windowsレジストリへ手動登録する
'      - Chrome: HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\NativeMessagingHosts\com.starterwebscrapingkit.cdpbridge
'      - Edge  : HKEY_CURRENT_USER\SOFTWARE\Microsoft\Edge\NativeMessagingHosts\com.starterwebscrapingkit.cdpbridge
'
'   4. このワークブックの「コピー」を、Excelの`XLStart`フォルダ
'      (`%APPDATA%\Microsoft\Excel\XLStart\`)に配置する
'      → そのコピーの`ThisWorkbook`モジュールの`Workbook_Open`から、
'        `Demo_NativeMessaging.StartNativeMessagingHostLoop`を呼び出すよう1行追加しておく
'
'   5. 対象ブラウザで対象タブを開き、1.で読み込んだ拡張機能のツールバーアイコンをクリック
'      → ブラウザがXLStartのExcelコピーを新規プロセスとして起動し、
'        `StartNativeMessagingHostLoop`がCDP待受を開始する
'      → 同プロシージャ内で`CDPBrowser.reattachNativeMessaging`によりブラウザ接続確認(`Browser.getVersion`)を
'        行った上で、いつも通りの`CDPBrowser`/`CDPContext`のAPIで後続の自動化コードを書き足せます
'
'   ※Chrome/Edgeはどちらも同じChromiumベースのため、Native Messagingのワイヤーフォーマット・
'     起動時に渡される引数・`allowed_origins`の書式(`chrome-extension://<id>/`のまま)は共通です。
'     変わるのはレジストリの登録先パスと、拡張機能ID(ブラウザごとに別々に発行される)だけです
'***************************************************************************************************
Option Explicit
Option Private Module



'***************************************************************************************************
'                             ■■■ ホストマニフェスト生成(Chrome向け) ■■■
'***************************************************************************************************
'* 機能　　：A. EXCEL.EXE自身を直接ホストにする構成で、Chrome向けホストマニフェストを生成します
'***************************************************************************************************
Sub GenerateNativeHostFiles_Chrome_直接Excel()
    Dim ExtensionID As String
    ExtensionID = InputBox("`chrome://extensions`で確認した、拡張機能ID(32文字の英字)を入力してください。", "Native Messagingホスト生成(Chrome)")
    If Len(ExtensionID) = 0 Then Exit Sub

    NativeMessagingSetup.GenerateNativeMessagingHostManifest NM_DirectExcel, NM_Chrome, DefaultNativeMessagingOutputFolder(NM_Chrome), ExtensionID
End Sub

'***************************************************************************************************
'* 機能　　：B. ラッパー.bat経由で`EXCEL.EXE /x`起動する構成で、Chrome向けホストマニフェストを生成します
'---------------------------------------------------------------------------------------------------
'* 注意事項：このワークブック(`ThisWorkbook.FullName`)を、`XLStart`へ配置する「コピー」と同じ場所で
'            実行してください。バッチには、実行時点の`ThisWorkbook.FullName`が焼き込まれます
'***************************************************************************************************
Sub GenerateNativeHostFiles_Chrome_ラッパーBAT経由()
    Dim ExtensionID As String
    ExtensionID = InputBox("`chrome://extensions`で確認した、拡張機能ID(32文字の英字)を入力してください。", "Native Messagingホスト生成(Chrome)")
    If Len(ExtensionID) = 0 Then Exit Sub

    NativeMessagingSetup.GenerateNativeMessagingHostManifest NM_WrapperBat, NM_Chrome, DefaultNativeMessagingOutputFolder(NM_Chrome), ExtensionID
End Sub



'***************************************************************************************************
'                             ■■■ ホストマニフェスト生成(Edge向け) ■■■
'***************************************************************************************************
'* 機能　　：A. EXCEL.EXE自身を直接ホストにする構成で、Edge向けホストマニフェストを生成します
'***************************************************************************************************
Sub GenerateNativeHostFiles_Edge_直接Excel()
    Dim ExtensionID As String
    ExtensionID = InputBox("`edge://extensions`で確認した、拡張機能ID(32文字の英字)を入力してください。", "Native Messagingホスト生成(Edge)")
    If Len(ExtensionID) = 0 Then Exit Sub

    NativeMessagingSetup.GenerateNativeMessagingHostManifest NM_DirectExcel, NM_Edge, DefaultNativeMessagingOutputFolder(NM_Edge), ExtensionID
End Sub

'***************************************************************************************************
'* 機能　　：B. ラッパー.bat経由で`EXCEL.EXE /x`起動する構成で、Edge向けホストマニフェストを生成します
'---------------------------------------------------------------------------------------------------
'* 注意事項：このワークブック(`ThisWorkbook.FullName`)を、`XLStart`へ配置する「コピー」と同じ場所で
'            実行してください。バッチには、実行時点の`ThisWorkbook.FullName`が焼き込まれます
'***************************************************************************************************
Sub GenerateNativeHostFiles_Edge_ラッパーBAT経由()
    Dim ExtensionID As String
    ExtensionID = InputBox("`edge://extensions`で確認した、拡張機能ID(32文字の英字)を入力してください。", "Native Messagingホスト生成(Edge)")
    If Len(ExtensionID) = 0 Then Exit Sub

    NativeMessagingSetup.GenerateNativeMessagingHostManifest NM_WrapperBat, NM_Edge, DefaultNativeMessagingOutputFolder(NM_Edge), ExtensionID
End Sub

'***************************************************************************************************
'* 機能　　：既定の生成先フォルダを返します(ブラウザごとにサブフォルダを分け、共存できるようにする)
'***************************************************************************************************
Private Function DefaultNativeMessagingOutputFolder(TargetBrowser As NMBrowserTarget) As String
    Dim SubFolderName As String: SubFolderName = IIf(TargetBrowser = NM_Edge, "Edge", "Chrome")
    DefaultNativeMessagingOutputFolder = Environ$("LOCALAPPDATA") & "\StarterWebScrapingKit\NativeMessaging\" & SubFolderName
End Function



'***************************************************************************************************
'                          ■■■ XLStartホスト側の待受ループ ■■■
'***************************************************************************************************
'* 機能　　：ChromeのNative Messagingホストとして起動された場合のみ、CDP待受ループを開始します
'---------------------------------------------------------------------------------------------------
'* 機能説明：`CDPCoreViaNativeMessaging.ConnectCDP`が`GetFileType`で標準入力がパイプかどうかを判定するため、
'            通常のダブルクリック起動時は何もせず即終了します(=このマクロをXLStartに置きっぱなしでも安全)
'* 注意事項：`XLStart`に配置する「コピー」の`ThisWorkbook.Workbook_Open`から、
'            `Demo_NativeMessaging.StartNativeMessagingHostLoop`を呼び出すようにしてください
'***************************************************************************************************
Public Sub StartNativeMessagingHostLoop()
    Const FromProcedureName As String = "Demo_NativeMessaging.StartNativeMessagingHostLoop"


    '1. 自プロセスが本当にChromeのNative Messagingホストとして起動されたか確認
    '※`ExpectedExtensionOrigin`は参考ログのみに使うため、ここでは省略(判定自体は`GetFileType`で行う)
    Dim nm As New CDPCoreViaNativeMessaging
    If Not nm.ConnectCDP() Then
        '通常のダブルクリック起動など。何もせず抜ける
        Exit Sub
    End If

    '2.
    Dim b As New CDPBrowser
    b.reattachNativeMessaging "NativeMessaging", nm
    
    Dim c As CDPContext
    Set c = b.getTab(setMain:=True)
    
    c.navigate "https://developer.chrome.com/docs/extensions/develop/concepts/native-messaging?hl=ja"

End Sub
