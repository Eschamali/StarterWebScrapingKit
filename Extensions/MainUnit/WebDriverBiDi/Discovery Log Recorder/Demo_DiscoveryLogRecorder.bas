Attribute VB_Name = "Demo_DiscoveryLogRecorder"
'***************************************************************************************************
'       exBiDi_DiscoveryLogRecorder 拡張 - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：一定時間、手動でブラウザを操作してもらい、その間に起きた出来事を
'            discovery_log.txt として保存するデモです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\WebDriverBiDi\Discovery Log Recorder\exBiDi_DiscoveryLogRecorder.cls
'---------------------------------------------------------------------------------------------------
'* 流れ　　：
'   1. ブラウザを起動し、記録対象ページへ遷移
'   2. 「OKを押したら記録開始」の案内を表示
'   3. OK後、指定秒数の間、ネットワーク/ナビゲーション/コンソール/DOM変化を記録
'      （この間に、ユーザーが手動でクリックやページ遷移などを行うことを想定）
'   4. discovery_log.txt を保存して完了報告
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・WORKSPACE_PATH をご自身の環境に合わせて設定してください
'***************************************************************************************************
Option Explicit



'ワークスペースパス
'※ StarterWebScrapingKit のルートフォルダを入力してください
Private Const WORKSPACE_PATH As String = ""

Private Const SAMPLE_DIR As String = "\Extensions\MainUnit\WebDriverBiDi\Discovery Log Recorder"

Private Const RECORDING_SECONDS As Double = 20



'***************************************************************************************************
'      ■■■ Demo：手動操作の一定時間記録 ■■■
'***************************************************************************************************
'* 機能　　：指定ページを開いた状態で、20秒間の手動操作をDiscovery Logとして記録します
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - 保存された discovery_log.txt に、[REQ]/[RES]/[NAV]/[LOG]/[DOM-ADD]等の行が時系列で並ぶこと
'   - excludeImagesAndCss=True により、画像/CSS等の通信が本文から除外され、末尾のサマリーに集計されること
'   - 記録中にページ遷移しても、DOM監視が自動的に再設置されること（[NAV] load の直後にDOM-ADD等が続く）
'***************************************************************************************************
Sub Demo_DiscoveryLogRecorder_手動操作の記録()

    '--- 1. ブラウザを起動し、記録対象ページへ遷移 ---
    Dim browserTab As WebDriverBiDiContext
    Set browserTab = ShSetting01_StartBrowser.StartBiDiModeContext("https://note.com/")

    '--- 2. Discovery Log Recorderの準備 ---
    Dim recorder As New exBiDi_DiscoveryLogRecorder
    recorder.Init browserTab

    '--- 3. 記録準備の案内 ---
    MsgBox "ブラウザの準備ができたら、[OK]を押してください。" & vbCrLf & vbCrLf & _
           "[OK]を押した直後から " & RECORDING_SECONDS & " 秒間、手動でページを操作してください。" & vbCrLf & _
           "（リンクのクリックやページ遷移、フォーム入力など、何でも構いません）", vbInformation, "記録準備"

    '--- 4. 記録開始 → 手動操作の記録 → 記録停止・保存 ---
    Debug.Print "[Demo] Discovery Log の記録を開始します..."
    recorder.StartDiscoveryLog excludeImagesAndCss:=True, captureDom:=True

    recorder.RecordEventsForSeconds RECORDING_SECONDS

    Dim logText As String
    logText = recorder.StopAndSaveDiscoveryLog(WORKSPACE_PATH & SAMPLE_DIR, "discovery_log.txt")

    Debug.Print logText

    '--- 5. 完了報告 ---
    MsgBox "Discovery Log を保存しました。" & vbCrLf & vbCrLf & _
           WORKSPACE_PATH & SAMPLE_DIR & "\discovery_log.txt" & vbCrLf & vbCrLf & _
           "内容はイミディエイトウィンドウにも出力しています。", vbInformation, "記録完了"

    browserTab.ThisWebDriverBiDiMode.quit

End Sub
