Attribute VB_Name = "Demo_DownloadWatcher"
'***************************************************************************************************
'       CDPexpansion_DownloadWatcher 拡張 - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：`CDPexpansion_DownloadWatcher.cls` を使ったダウンロード監視のサンプルコードです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\Download Watcher\CDPexpansion_DownloadWatcher.cls
'---------------------------------------------------------------------------------------------------
'* テストサイト：https://custom-img.lb-product.com/
'*   サイズ・単位・拡張子を指定して任意サイズの画像を生成＆ダウンロードできる無料サービスです
'*   フォーム仕様:
'*     POST https://custom-img.lb-product.com/download
'*     size      : 1 ~ 999  （GBの場合は 1 まで）
'*     unit      : KB / MB / GB
'*     extension : png / jpeg / gif / webp
'---------------------------------------------------------------------------------------------------
'* 内部データ構造（新・Dictionary完結設計）：
'*   downloadInfos
'*    └ Key: Guid
'*    └ Value: Dictionary {
'*        "WillBegin"       → Dictionary  ← downloadWillBegin のraw params 丸ごと
'*                              { "guid", "url", "suggestedFilename", "frameId", ... }
'*        "Progress"        → Dictionary  ← downloadProgress のraw params 丸ごと（最新更新）
'*                              { "guid", "state", "receivedBytes", "totalBytes", "filePath", ... }
'*        "isBrowserMethod" → Boolean
'*      }
'---------------------------------------------------------------------------------------------------
'* ネットワークスロットリングについて：
'*   dw.ThrottleNetwork(KB/s) でネット速度を意図的に制限します。
'*   custom-img は「JS fetch → Blob → <a click>」方式なので、
'*   ThrottleNetwork は fetch フェーズ（サーバーからBlobを取得する段階）に効きます。
'*   Blob → ディスク書き込み段階はローカル処理のため速度制限が効きにくい場合があります。
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'*   ・WORKSPACE_PATH を各自の環境に合わせて設定してください
'*   ・ダウンロード先フォルダは WatchStart が自動作成します
'***************************************************************************************************
Option Explicit



'ワークスペースパス
'※ StarterWebScrapingKit のルートフォルダを入力してください
Private Const WORKSPACE_PATH As String = ""

'テストサイト
Private Const TESTSITE_URL As String = "https://custom-img.lb-product.com/"



'***************************************************************************************************
'              ■■■ Demo 01：シンプルなダウンロード（5MB PNG） ■■■
'***************************************************************************************************
'* 機能　　：custom-img.lb-product.com で 5MB の PNG をダウンロードする基本デモです
'---------------------------------------------------------------------------------------------------
'* スロットリング: 300 KB/s → 5MB ≒ 約17秒
'* 確認ポイント：
'   - WatchStart → フォーム送信 → WaitCompleted の流れで完了待機できること
'   - SuggestedFilename / DownloadedFilePath が正しく取れること
'   - GetWillBeginParams / GetProgressParams で生 params が参照できること
'***************************************************************************************************
Sub Demo_DownloadWatcher_01_5MBのPNGをダウンロード()

    Dim outDir As String
    outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- 1. テストサイトを開く ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動(TESTSITE_URL)
    Debug.Print "[Demo01] ページ読み込み完了"

    '--- 2. DownloadWatcher 初期化 & 監視開始 ---
    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser
    dw.WatchStart outDir   '← ダウンロードトリガーの前に必ず呼ぶ

    '--- 3. ネット速度を制限（進捗を実感できるように）---
    dw.ThrottleNetwork DownloadKBps:=300   '300KB/s → 5MB ≒ 17秒

    '--- 4. フォームを操作：5MB / PNG ---
    Call SetDownloadForm(browser, Size:=5, Unit:="MB", Extension:="png")

    '--- 5. ダウンロードボタンをクリック ---
    Debug.Print "[Demo01] ダウンロードボタンをクリック..."
    browser.getElementByID("submit-button").click

    '--- 6. 完了まで待機（最大 120 秒） ---
    Debug.Print "[Demo01] ダウンロード完了待ち..."
    If dw.WaitCompleted(TimeoutSec:=120) Then
        Debug.Print "[Demo01] ○ 完了！"
        Debug.Print "[Demo01]   SuggestedFilename : " & dw.SuggestedFilename
        Debug.Print "[Demo01]   DownloadedFilePath: " & dw.DownloadedFilePath

        '--- GetWillBeginParams / GetProgressParams で生データを参照 ---
        Dim allGuids As Collection: Set allGuids = dw.GetAllGuids()
        Dim latestGuid As String: latestGuid = CStr(allGuids(allGuids.Count))

        Dim wbP As Scripting.Dictionary: Set wbP = dw.GetWillBeginParams(latestGuid)
        Dim pgP As Scripting.Dictionary: Set pgP = dw.GetProgressParams(latestGuid)
        Debug.Print "[Demo01]   [生params] url      : " & wbP("url")
        Debug.Print "[Demo01]   [生params] totalBytes: " & pgP("totalBytes")
        Debug.Print "[Demo01]   [生params] filePath  : " & pgP("filePath")

        MsgBox "ダウンロード完了！" & vbCrLf & _
               "ファイル名: " & dw.SuggestedFilename & vbCrLf & _
               "保存先: " & dw.DownloadedFilePath, vbInformation
    Else
        Debug.Print "[Demo01] × タイムアウトまたはキャンセル。State=" & dw.state
        MsgBox "ダウンロード失敗。State=" & dw.state, vbCritical
    End If

    dw.UnthrottleNetwork
    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 02：複数サイズ・形式を連続ダウンロード ■■■
'***************************************************************************************************
'* 機能　　：サイズや形式を変えながら複数ファイルを連続ダウンロードするデモです
'---------------------------------------------------------------------------------------------------
'* ダウンロード構成：
'   Round 1:  1MB JPEG  → 300KB/s ≒  3秒
'   Round 2:  5MB PNG   → 300KB/s ≒ 17秒
'   Round 3: 10MB WEBP  → 300KB/s ≒ 34秒
'* 確認ポイント：
'   - 毎ラウンド WatchStart を呼んで downloadInfos がリセットされること
'   - 各ラウンド完了後に PrintSummary でGUID別情報が確認できること
'***************************************************************************************************
Sub Demo_DownloadWatcher_02_複数ファイルを連続ダウンロード()

    Dim outDir As String
    outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- ダウンロード構成 ---
    Dim sizes(1 To 3) As Integer
    Dim units(1 To 3) As String
    Dim exts(1 To 3) As String
    sizes(1) = 1:  units(1) = "MB": exts(1) = "jpeg"
    sizes(2) = 5:  units(2) = "MB": exts(2) = "png"
    sizes(3) = 10: units(3) = "MB": exts(3) = "webp"

    '--- 1. テストサイトを開く ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動(TESTSITE_URL)

    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser

    '--- 3ラウンド連続ダウンロード ---
    Dim i As Integer
    For i = 1 To 3
        Debug.Print "[Demo02] ─── ラウンド " & i & " / 3 (" & sizes(i) & units(i) & " " & exts(i) & ") ───"

        '毎回 WatchStart で downloadInfos をリセット＆スナップショット更新
        dw.WatchStart outDir
        dw.ThrottleNetwork DownloadKBps:=300

        'フォーム入力 → クリック → 完了待ち
        Call SetDownloadForm(browser, sizes(i), units(i), exts(i))
        browser.getElementByID("submit-button").click
        Debug.Print "[Demo02]   完了待ち..."

        If dw.WaitCompleted(TimeoutSec:=180) Then
            Debug.Print "[Demo02]   ○ 保存完了: " & dw.DownloadedFilePath _
                      & " (" & Format(dw.TotalBytes / 1024 / 1024, "0.00") & " MB)"
            dw.PrintSummary   '← GUIDベースのサマリーを確認
        Else
            Debug.Print "[Demo02]   × 失敗 (State=" & dw.state & ")"
            MsgBox "Round " & i & " でダウンロード失敗。", vbCritical
            dw.UnthrottleNetwork
            Exit For
        End If

        dw.UnthrottleNetwork
        browser.sleep 1.5   '次のダウンロードの前にボタンが再有効化されるまで待つ
    Next i

    Debug.Print "[Demo02] 全ラウンド完了！"
    MsgBox "3ファイルのダウンロードが完了しました！" & vbCrLf & outDir, vbInformation

    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 03：進捗を表示しながら大容量ダウンロード（1GB PNG） ■■■
'***************************************************************************************************
'* 機能　　：カスタムポーリングループで進捗をリアルタイム表示するデモです
'---------------------------------------------------------------------------------------------------
'* スロットリング: 1000 KB/s（1Mbps） → 1GB ≒ 約17分
'*   速度目安：5000KB/s→3.4分 / 2000KB/s→8.5分 / 1000KB/s→17分
'*   ※ 途中でブレークして動作確認するだけでも OK です
'* 確認ポイント：
'   - Progress / ReceivedBytes / TotalBytes が段階的に更新されること
'   - GetProgressParams で生params の receivedBytes を直接確認できること
'***************************************************************************************************
Sub Demo_DownloadWatcher_03_進捗表示しながら大容量ダウンロード()

    Dim outDir As String
    outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- 1. テストサイトを開く ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動(TESTSITE_URL)

    '--- 2. DownloadWatcher 初期化 & 監視開始 ---
    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser
    dw.WatchStart outDir
'    dw.ThrottleNetwork DownloadKBps:=1000   '1000KB/s → 1GB ≒ 17分

    '--- 3. フォームを操作：1GB / png ---
    Call SetDownloadForm(browser, Size:=1, Unit:="GB", Extension:="png")

    '--- 4. クリック ---
    browser.getElementByID("submit-button").click
    Debug.Print "[Demo03] ダウンロードをトリガーしました（fetchフェーズ開始）"
    Debug.Print "[Demo03] ※ fetch完了後にダウンロードイベントが発生します..."

    '--- 5. カスタムポーリングループで進捗表示 ---
    Dim t As Double: t = Timer
    Dim lastPct As Long: lastPct = -1

    Do
        browser.TakeEvents
        DoEvents

        If dw.HasStarted Then
            '後方互換プロパティで進捗表示（内部では mLastGuid → Progress params を参照）
            If dw.Progress >= 0 Then
                Dim pct As Long: pct = CLng(dw.Progress * 100)
                If pct <> lastPct Then
                    Debug.Print "[Demo03]   " & Format(pct, "00") & "%" _
                              & Format(dw.ReceivedBytes / 1024 / 1024, "0.00") & " MB / " _
                              & Format(dw.TotalBytes / 1024 / 1024, "0.00") & " MB"
                    lastPct = pct
                End If
            Else
                Debug.Print "[Demo03]   受信: " & Format(dw.ReceivedBytes / 1024 / 1024, "0.00") & " MB (合計不明)"
            End If
        End If

        If dw.IsCompleted Then
            Debug.Print "[Demo03] ○ 完了！ " & Format(Timer - t, "0.0") & "秒"
            Debug.Print "[Demo03]   保存先: " & dw.DownloadedFilePath
            dw.PrintSummary
            MsgBox "完了！ " & Format(Timer - t, "0.0") & "秒" & vbCrLf & dw.DownloadedFilePath, vbInformation
            Exit Do
        End If

        If dw.state = "canceled" Then
            Debug.Print "[Demo03] × キャンセルされました"
            Exit Do
        End If

        If (Timer - t) > 1800 Then   '30分タイムアウト
            Debug.Print "[Demo03] × タイムアウト"
            Exit Do
        End If

        browser.sleep 0.5
    Loop

    dw.UnthrottleNetwork
    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 04：ダウンロード後に別名・別フォルダへ保存（SaveAs） ■■■
'***************************************************************************************************
'* 機能　　：3MB WEBP をダウンロードして、タイムスタンプ付きのファイル名で別フォルダへ保存するデモです
'---------------------------------------------------------------------------------------------------
'* スロットリング: 200 KB/s → 3MB ≒ 約15秒
'* 確認ポイント：
'   - WaitCompleted 後に dw.SaveAs でリネーム & 移動できること
'***************************************************************************************************
Sub Demo_DownloadWatcher_04_別名で保存()

    Dim outDir As String:  outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"
    Dim saveDir As String: saveDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Saved"

    '--- 1. テストサイトを開く ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動(TESTSITE_URL)

    '--- 2. DownloadWatcher 初期化 & 監視開始 ---
    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser
    dw.WatchStart outDir
    dw.ThrottleNetwork DownloadKBps:=200   '200KB/s → 3MB ≒ 15秒

    '--- 3. フォーム操作：3MB / WEBP ---
    Call SetDownloadForm(browser, Size:=3, Unit:="MB", Extension:="webp")

    '--- 4. クリック → 完了待ち ---
    browser.getElementByID("submit-button").click
    Debug.Print "[Demo04] ダウンロード完了待ち..."

    If Not dw.WaitCompleted(TimeoutSec:=120) Then
        MsgBox "ダウンロード失敗。State=" & dw.state, vbCritical
        dw.UnthrottleNetwork
        browser.quit
        Exit Sub
    End If
    Debug.Print "[Demo04] ダウンロード完了: " & dw.DownloadedFilePath
    dw.UnthrottleNetwork

    '--- 5. 日付付きファイル名で別フォルダに保存 ---
    Dim newName As String: newName = "testimg_" & Format(Now, "yyyymmdd_HHmmss") & ".webp"
    Dim saved As String:   saved = dw.SaveAs(saveDir, newName)   'Guid省略 → mLastGuid 自動使用

    If Len(saved) > 0 Then
        Debug.Print "[Demo04] ○ 別名保存完了！ → " & saved
        MsgBox "別名保存完了！" & vbCrLf & saved, vbInformation
    Else
        Debug.Print "[Demo04] × 別名保存失敗"
        MsgBox "SaveAs に失敗しました。", vbCritical
    End If

    'browser.quit

End Sub



'***************************************************************************************************
'   ■■■ Demo 05：複数ダウンロードを同時トリガー → WaitAllCompleted で一括待機 ■■■
'***************************************************************************************************
'* 機能　　：3件のダウンロードを jsEval で同時トリガーし、
'            downloadInfos の Dictionary 完結構造（WillBegin/Progress サブ辞書）を実演します
'---------------------------------------------------------------------------------------------------
'* ダウンロード構成（同時）：
'   DL1:  3MB PNG
'   DL2:  5MB JPEG
'   DL3:  2MB GIF
'---------------------------------------------------------------------------------------------------
'* スロットリング: 500 KB/s → 合計10MB ≒ 20秒（3件並行）
'---------------------------------------------------------------------------------------------------
'* ポイント：
'   - GetInfo(guid) が返す Dictionary の構造を確認
'     info("WillBegin")("suggestedFilename")  ← WillBegin 生params を直接参照
'     info("Progress")("state")               ← Progress 生params を直接参照
'   - GetWillBeginParams / GetProgressParams で個別サブ辞書も取得可能
'   - WaitAllCompleted で全件完了まで一括待機
'***************************************************************************************************
Sub Demo_DownloadWatcher_05_複数同時DLをまとめて待つ()

    Dim outDir As String
    outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- 1. テストサイトを開く ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動(TESTSITE_URL)
    Debug.Print "[Demo05] ページ読み込み完了"

    '--- 2. DownloadWatcher 初期化 & 監視開始 ---
    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser
    dw.WatchStart outDir
    dw.ThrottleNetwork DownloadKBps:=500   '500KB/s → 合計10MB ≒ 20秒

    '--- 3. JS で 3件のダウンロードを「同時」トリガー ---
    '    ① ページから CSRF トークンを取得
    '    ② 3本の fetch を同時発行（各 fetch は独立して非同期処理）
    '    ③ 各 fetch 完了後に <a click> → downloadWillBegin が 3回発火
    '    ④ downloadInfos に 3つの GUID エントリが蓄積される
    '行継続文字の24個上限を回避するため、js = js & "..." 形式で分割して構築
    Dim js As String
    js = "(function(){"
    js = js & "  var token = document.querySelector('input[name=""_token""]').value;"
    js = js & "  var jobs = ["
    js = js & "    {size:'3',unit:'MB',ext:'png'},"
    js = js & "    {size:'5',unit:'MB',ext:'jpeg'},"
    js = js & "    {size:'2',unit:'MB',ext:'gif'}"
    js = js & "  ];"
    js = js & "  jobs.forEach(function(cfg) {"
    js = js & "    var fd = new FormData();"
    js = js & "    fd.append('_token', token);"
    js = js & "    fd.append('size',      cfg.size);"
    js = js & "    fd.append('unit',      cfg.unit);"
    js = js & "    fd.append('extension', cfg.ext);"
    js = js & "    fetch('https://custom-img.lb-product.com/download', {method:'POST', body:fd})"
    js = js & "    .then(function(r)    { return r.blob(); })"
    js = js & "    .then(function(blob) {"
    js = js & "      var url = URL.createObjectURL(blob);"
    js = js & "      var a   = document.createElement('a');"
    js = js & "      a.href  = url;"
    js = js & "      a.download = cfg.size + cfg.unit + '.' + cfg.ext;"
    js = js & "      document.body.appendChild(a);"
    js = js & "      a.click();"
    js = js & "      URL.revokeObjectURL(url);"
    js = js & "    });"
    js = js & "  });"
    js = js & "})();"

    browser.jsEval js
    Debug.Print "[Demo05] 3件の fetch を同時発行しました..."
    Debug.Print "[Demo05] ※ fetch完了後に downloadWillBegin が 3回発火します"

    '--- 4. 3件の WillBegin を受信するまで待機 ---
    Dim tWait As Double: tWait = Timer
    Do
        browser.TakeEvents
        DoEvents
        Debug.Print "[Demo05]   待機中... DownloadCount=" & dw.DownloadCount & " / 3"
        If dw.DownloadCount >= 3 Then Exit Do
        If (Timer - tWait) > 120 Then
            Debug.Print "[Demo05] × 検出タイムアウト。検出数=" & dw.DownloadCount
            dw.UnthrottleNetwork
            Exit Sub
        End If
        browser.sleep 0.5
    Loop
    Debug.Print "[Demo05] " & dw.DownloadCount & "件のダウンロードを検出！"

    '--- 4. WaitAllCompleted で全件（3件）完了まで一括待機 ---
    '    ExpectedCount:=3 を指定 → 3件全部受信＆完了するまで待つ
    '    Phase1（WillBegin待ち）と Phase2（完了待ち）が内部で自動切替
    If dw.WaitAllCompleted(ExpectedCount:=3, TimeoutSec:=180) Then
        Debug.Print "[Demo05] ○ 全件完了！"
        dw.PrintSummary   '← GUID別サマリーをイミディエイトウィンドウに出力

        '--- 6. GetAllGuids で全GUIDを列挙して、Dictionary 完結構造を実演 ---
        Debug.Print "[Demo05] ─── Dictionary 完結構造の実演 ───"
        Debug.Print "[Demo05]   GetInfo(guid) が返す内部構造："
        Debug.Print "[Demo05]   info(""WillBegin"")  → WillBegin のraw params Dictionary"
        Debug.Print "[Demo05]   info(""Progress"")   → Progress  のraw params Dictionary"
        Debug.Print ""

        Dim guids As Collection: Set guids = dw.GetAllGuids()
        Dim guid As Variant
        For Each guid In guids
            '-- GetInfo() で エントリ辞書 を取得 --
            Dim info As Scripting.Dictionary: Set info = dw.GetInfo(CStr(guid))

            '-- WillBegin サブ辞書を直接参照 --
            Dim wbParams As Scripting.Dictionary: Set wbParams = info("WillBegin")
            Dim pgParams As Scripting.Dictionary: Set pgParams = info("Progress")

            Debug.Print "[Demo05]   ┌─ guid: " & Left(CStr(guid), 8) & "..."
            Debug.Print "[Demo05]   │  WillBegin.suggestedFilename = " & wbParams("suggestedFilename")
            Debug.Print "[Demo05]   │  WillBegin.url               = " & wbParams("url")
            Debug.Print "[Demo05]   │  Progress.state              = " & pgParams("state")
            Debug.Print "[Demo05]   │  Progress.totalBytes         = " & pgParams("totalBytes")
            Debug.Print "[Demo05]   └  Progress.filePath           = " & pgParams("filePath")
            Debug.Print ""

            '-- GetWillBeginParams / GetProgressParams での個別取得 --
            Dim wb2 As Scripting.Dictionary: Set wb2 = dw.GetWillBeginParams(CStr(guid))
            Dim pg2 As Scripting.Dictionary: Set pg2 = dw.GetProgressParams(CStr(guid))
            Debug.Print "[Demo05]   ※ GetWillBeginParams / GetProgressParams でも同一データ取得可"
            Debug.Print "[Demo05]      wb2(""suggestedFilename"") = " & wb2("suggestedFilename")
            Debug.Print "[Demo05]      pg2(""filePath"")          = " & pg2("filePath")
            Debug.Print ""
        Next guid

        MsgBox "全3件のダウンロード完了！" & vbCrLf & _
               "完了数: " & dw.CompletedCount() & " / " & dw.DownloadCount & vbCrLf & _
               "保存先: " & outDir, vbInformation
    Else
        Debug.Print "[Demo05] × WaitAllCompleted タイムアウト"
        Debug.Print "[Demo05]   完了: " & dw.CompletedCount() & "/" & dw.DownloadCount
        dw.PrintSummary
        MsgBox "タイムアウト。完了数: " & dw.CompletedCount() & "/" & dw.DownloadCount, vbCritical
    End If

    dw.UnthrottleNetwork
    'browser.quit

End Sub



'***************************************************************************************************
'                           ■■■ ユーティリティ（内部用） ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：custom-img.lb-product.com のダウンロードフォームに値を入力します
'---------------------------------------------------------------------------------------------------
'* 引数　　：browser    CDPBrowser インスタンス
'            Size       ファイルサイズ（1~999、GBの場合は1まで）
'            Unit       "KB" / "MB" / "GB"
'            Extension  "png" / "jpeg" / "gif" / "webp"
'***************************************************************************************************
Private Sub SetDownloadForm(browser As CDPBrowser, Size As Integer, Unit As String, Extension As String)

    '--- サイズ入力 ---
    Dim sizeEl As CDPElement: Set sizeEl = browser.getElementByID("size")
    sizeEl.clearValue
    sizeEl.sendString CStr(Size)
    Debug.Print "[SetDownloadForm] size=" & Size

    '--- 単位セレクト ---
    browser.getElementByID("unit").setSelection Unit
    Debug.Print "[SetDownloadForm] unit=" & Unit

    '--- 拡張子セレクト ---
    browser.getElementByID("extension").setSelection Extension
    Debug.Print "[SetDownloadForm] extension=" & Extension

End Sub
