Attribute VB_Name = "Demo_DownloadWatcher"
'***************************************************************************************************
'       exCDP_DownloadWatcher 拡張 - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：`exCDP_DownloadWatcher.cls` を使ったダウンロード監視のサンプルコードです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\Download Watcher\exCDP_DownloadWatcher.cls
'---------------------------------------------------------------------------------------------------
'* テストサイト：https://custom-img.lb-product.com/
'*   サイズ・単位・拡張子を指定して任意サイズの画像を生成しダウンロードできる無料サービスです
'*   フォーム仕様:
'*     POST https://custom-img.lb-product.com/download
'*     size      : 1 ~ 999  （GBの場合は 1 まで）
'*     unit      : KB / MB / GB
'*     extension : png / jpeg / gif / webp
'---------------------------------------------------------------------------------------------------
'* `exCDP_DownloadWatcher.cls` の実際の公開APIに合わせて実装しています（GUID単位のDictionary管理）：
'*   ・setDownloadBehavior(behavior, downloadPath) = True/False  … 監視の開始/停止
'*   ・DownloadWillBeginInfo(guid, downloadWillBegin_parameters) … url / suggestedFilename / frameID
'*   ・DownloadProgressInfo(guid, downloadProgress_parameters)   … state / totalBytes / receivedBytes / filePath
'*   ・DownloadGuidList / DownloadGuidCount / DownloadStateCount(state)
'*   ・AutoWaitSingleCompleted(guid, TimeOutSecond)   … 単一GUID完了待ち（Boolean）
'*   ・AutoWaitMultiCompleted(CompleteCount, TimeOutSecond, ClearDLHistory) … 複数完了待ち（完了数Long）
'*   ・PrintDownloadSummary(Optional guid)
'*   ・ClearDownloadCompleted / ClearDownloadCanceled
'---------------------------------------------------------------------------------------------------
'* ネットワークスロットリングについて：
'*   ThrottleNetwork(browserTab, KB/s) で意図的に速度を制限します（このモジュールのローカルヘルパー。
'*   `exCDP_DownloadWatcher.cls` 自体にはスロットリング機能は無いため、CDP の
'*   `Network.emulateNetworkConditions` を直接呼び出しています）。
'*   custom-img は「JS fetch → Blob → <a click>」構成なので、
'*   ThrottleNetwork は fetch フェーズ（サーバーからBlobを取得する段階）に効きます。
'*   Blob → ディスク書き込み段階はローカル処理のため速度制限が効きにくい場合があります。
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'*   ・WORKSPACE_PATH をご自身の環境に合わせて設定してください
'*   ・ダウンロード先フォルダは Downloads サブフォルダを自動作成します
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
'* スロットリング: 300 KB/s → 5MB で約17秒
'* 確認ポイント：
'   - setDownloadBehavior → フォーム送信 → AutoWaitSingleCompleted の流れで完全同期待機できること
'   - DownloadWillBeginInfo / DownloadProgressInfo で suggestedFilename / filePath が取得できること
'   - 同メソッドで url / totalBytes など、その他の生paramsも個別に参照できること
'***************************************************************************************************
Sub Demo_DownloadWatcher_01_5MBのPNGをダウンロード()

    Dim outDir As String
    outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- 1. テストサイトを開く ---
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext(TESTSITE_URL)
    Debug.Print "[Demo01] ページ読み込み完了"

    '--- 2. DownloadWatcher 準備 & 監視開始 ---
    Dim dw As New exCDP_DownloadWatcher
    dw.Init browserTab
    dw.setDownloadBehavior(behavior:=behavior_allow, downloadPath:=outDir) = True   '※ ダウンロードトリガーの前に必ず呼ぶ

    '--- 3. ネット速度を制限（進捗が観測しやすいように） ---
    ThrottleNetwork browserTab, DownloadKBps:=300   '300KB/s → 5MB は約17秒

    '--- 4. フォームを操作：5MB / PNG ---
    Dim priorCount As Long: priorCount = dw.DownloadGuidCount
    Call SetDownloadForm(browserTab, Size:=5, Unit:="MB", Extension:="png")

    '--- 5. ダウンロードボタンをクリック ---
    Debug.Print "[Demo01] ダウンロードボタンをクリック..."
    browserTab.getElementByID("submit-button").click

    '--- 6. downloadWillBegin を検知してguidを確保 ---
    Dim guid As String
    guid = WaitForNewGuid(browserTab, dw, priorCount, TimeOutSecond:=30)

    If Len(guid) = 0 Then
        Debug.Print "[Demo01] × downloadWillBegin を検知できませんでした"
        MsgBox "ダウンロードが開始されませんでした。", vbCritical
        UnthrottleNetwork browserTab
        browserTab.ThisCDPBrowser.quit
        Exit Sub
    End If

    '--- 7. 完了まで待機（最大 120 秒） ---
    Debug.Print "[Demo01] ダウンロード完了待ち..."
    If dw.AutoWaitSingleCompleted(guid, TimeOutSecond:=120) Then
        Debug.Print "[Demo01] ○ 完了！"
        Debug.Print "[Demo01]   suggestedFilename : " & dw.DownloadWillBeginInfo(guid, downloadWillBegin_suggestedFilename)
        Debug.Print "[Demo01]   filePath          : " & dw.DownloadProgressInfo(guid, downloadProgress_filePath)

        '--- 生paramsを個別に参照できることの確認 ---
        Debug.Print "[Demo01]   [生params] url        : " & dw.DownloadWillBeginInfo(guid, downloadWillBegin_Url)
        Debug.Print "[Demo01]   [生params] totalBytes  : " & dw.DownloadProgressInfo(guid, downloadProgress_totalBytes)

        MsgBox "ダウンロード完了！" & vbCrLf & _
               "ファイル名: " & dw.DownloadWillBeginInfo(guid, downloadWillBegin_suggestedFilename) & vbCrLf & _
               "保存先: " & dw.DownloadProgressInfo(guid, downloadProgress_filePath), vbInformation
    Else
        Debug.Print "[Demo01] × タイムアウトまたはキャンセル。State=" & dw.DownloadProgressInfo(guid, downloadProgress_state)
        MsgBox "ダウンロード失敗。State=" & dw.DownloadProgressInfo(guid, downloadProgress_state), vbCritical
    End If

    UnthrottleNetwork browserTab
    browserTab.ThisCDPBrowser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 02：複数サイズ・拡張子を連続ダウンロード ■■■
'***************************************************************************************************
'* 機能　　：サイズ・拡張子を変えながら複数ファイルを連続ダウンロードするデモです
'---------------------------------------------------------------------------------------------------
'* ダウンロード構成：
'   Round 1:  1MB JPEG  → 300KB/s で約  3秒
'   Round 2:  5MB PNG   → 300KB/s で約 17秒
'   Round 3: 10MB WEBP  → 300KB/s で約 34秒
'* 確認ポイント：
'   - 同じ dw インスタンスで複数ラウンドの guid が積み上がっていくこと（履歴は自動リセットされない）
'   - 各ラウンド完了後、PrintDownloadSummary(guid) でGUID別情報が確認できること
'   - 最後に ClearDownloadCompleted で完了済み履歴を一括削除できること
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
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext(TESTSITE_URL)

    Dim dw As New exCDP_DownloadWatcher
    dw.Init browserTab
    dw.setDownloadBehavior(behavior:=behavior_allow, downloadPath:=outDir) = True

    '--- 3ラウンド連続ダウンロード ---
    Dim i As Integer
    For i = 1 To 3
        Debug.Print "[Demo02] ────── ラウンド " & i & " / 3 (" & sizes(i) & units(i) & " " & exts(i) & ") ──────"

        Dim priorCount As Long: priorCount = dw.DownloadGuidCount
        ThrottleNetwork browserTab, DownloadKBps:=300

        'フォーム操作 → クリック → 完了待ち
        Call SetDownloadForm(browserTab, sizes(i), units(i), exts(i))
        browserTab.getElementByID("submit-button").click
        Debug.Print "[Demo02]   完了待ち..."

        Dim guid As String
        guid = WaitForNewGuid(browserTab, dw, priorCount, TimeOutSecond:=180)

        If Len(guid) = 0 Then
            Debug.Print "[Demo02]   × downloadWillBegin を検知できませんでした"
            MsgBox "Round " & i & " でダウンロードが開始されませんでした。", vbCritical
            UnthrottleNetwork browserTab
            Exit For
        End If

        If dw.AutoWaitSingleCompleted(guid, TimeOutSecond:=180) Then
            Debug.Print "[Demo02]   ○ 保存完了: " & dw.DownloadProgressInfo(guid, downloadProgress_filePath) _
                      & " (" & Format(CDbl(dw.DownloadProgressInfo(guid, downloadProgress_totalBytes)) / 1024 / 1024, "0.00") & " MB)"
            dw.PrintDownloadSummary guid   '※ このGUIDのサマリーを確認
        Else
            Debug.Print "[Demo02]   × 失敗 (State=" & dw.DownloadProgressInfo(guid, downloadProgress_state) & ")"
            MsgBox "Round " & i & " でダウンロード失敗。", vbCritical
            UnthrottleNetwork browserTab
            Exit For
        End If

        UnthrottleNetwork browserTab
        CDPHelpers.Sleep 1.5    '次のダウンロードの前にボタンが再有効化されるまで待つ
    Next i

    Debug.Print "[Demo02] 全ラウンド完了！ 現在の履歴: " & dw.DownloadGuidCount & " 件 (completed=" & dw.DownloadStateCount(state_completed) & ")"

    '--- 完了済み履歴を一括削除できることの確認 ---
    dw.ClearDownloadCompleted
    Debug.Print "[Demo02] ClearDownloadCompleted 後の履歴: " & dw.DownloadGuidCount & " 件"

    MsgBox "3ファイルのダウンロードが完了しました！" & vbCrLf & outDir, vbInformation

    browserTab.ThisCDPBrowser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 03：進捗表示しながら大容量ダウンロード（1GB PNG） ■■■
'***************************************************************************************************
'* 機能　　：カスタムポーリングループで進捗をリアルタイム表示するデモです
'---------------------------------------------------------------------------------------------------
'* スロットリング: 1000 KB/s（1Mbps） → 1GB で約17分
'*   速度目安：5000KB/s≒3.4分 / 2000KB/s≒8.5分 / 1000KB/s≒17分
'*   ※ 途中でブレークして動作確認するだけでも OK です
'* 確認ポイント：
'   - downloadWillBegin 検知（DownloadGuidCount が増える）タイミングが分かること
'   - DownloadProgressInfo で receivedBytes / totalBytes / state が段階的に更新されること
'***************************************************************************************************
Sub Demo_DownloadWatcher_03_進捗表示しながら大容量ダウンロード()

    Dim outDir As String
    outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- 1. テストサイトを開く ---
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext(TESTSITE_URL)

    '--- 2. DownloadWatcher 準備 & 監視開始 ---
    Dim dw As New exCDP_DownloadWatcher
    dw.Init browserTab
    dw.setDownloadBehavior(behavior:=behavior_allow, downloadPath:=outDir) = True
'    ThrottleNetwork browserTab, DownloadKBps:=1000   '1000KB/s → 1GB は約17分

    '--- 3. フォームを操作：1GB / png ---
    Dim priorCount As Long: priorCount = dw.DownloadGuidCount
    Call SetDownloadForm(browserTab, Size:=1, Unit:="GB", Extension:="png")

    '--- 4. クリック ---
    browserTab.getElementByID("submit-button").click
    Debug.Print "[Demo03] ダウンロードをトリガーしました（fetchフェーズ開始）"
    Debug.Print "[Demo03] → fetch完了後にダウンロードイベントが発火します..."

    '--- 5. downloadWillBegin を検知 ---
    Dim guid As String
    guid = WaitForNewGuid(browserTab, dw, priorCount, TimeOutSecond:=600)
    If Len(guid) = 0 Then
        Debug.Print "[Demo03] × downloadWillBegin タイムアウト"
        browserTab.ThisCDPBrowser.quit
        Exit Sub
    End If
    Debug.Print "[Demo03] ○ downloadWillBegin 検知！ guid=" & Left(guid, 8) & "..."

    '--- 6. カスタムポーリングループで進捗表示 ---
    Dim t As Double: t = Timer
    Dim lastPct As Long: lastPct = -1

    Do
        browserTab.ThisCDPBrowser.TakeEvents

        Dim state As String: state = dw.DownloadProgressInfo(guid, downloadProgress_state)
        Dim totalBytes As Double: totalBytes = dw.DownloadProgressInfo(guid, downloadProgress_totalBytes)
        Dim receivedBytes As Double: receivedBytes = dw.DownloadProgressInfo(guid, downloadProgress_receivedBytes)

        If totalBytes > 0 Then
            Dim pct As Long: pct = CLng(receivedBytes / totalBytes * 100)
            If pct <> lastPct Then
                Debug.Print "[Demo03]   " & Format(pct, "00") & "% " _
                          & Format(receivedBytes / 1024 / 1024, "0.00") & " MB / " _
                          & Format(totalBytes / 1024 / 1024, "0.00") & " MB"
                lastPct = pct
            End If
        Else
            Debug.Print "[Demo03]   受信: " & Format(receivedBytes / 1024 / 1024, "0.00") & " MB (合計不明)"
        End If

        If state = "completed" Then
            Debug.Print "[Demo03] ○ 完了！ " & Format(Timer - t, "0.0") & "秒"
            Debug.Print "[Demo03]   保存先: " & dw.DownloadProgressInfo(guid, downloadProgress_filePath)
            dw.PrintDownloadSummary guid
            MsgBox "完了！ " & Format(Timer - t, "0.0") & "秒" & vbCrLf & dw.DownloadProgressInfo(guid, downloadProgress_filePath), vbInformation
            Exit Do
        End If

        If state = "canceled" Then
            Debug.Print "[Demo03] × キャンセルされました"
            Exit Do
        End If

        If (Timer - t) > 1800 Then   '30分タイムアウト
            Debug.Print "[Demo03] × タイムアウト"
            Exit Do
        End If

        CDPHelpers.Sleep 0.5
    Loop

    UnthrottleNetwork browserTab
    browserTab.ThisCDPBrowser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 04：ダウンロード後に別名・別フォルダへ保存 ■■■
'***************************************************************************************************
'* 機能　　：3MB WEBP をダウンロードして、タイムスタンプ付きのファイル名で別フォルダへ保存するデモです
'---------------------------------------------------------------------------------------------------
'* スロットリング: 200 KB/s → 3MB で約15秒
'* 確認ポイント：
'   - AutoWaitSingleCompleted 完了後、SaveDownloadedFileAs でリネーム＆移動できること
'     （`exCDP_DownloadWatcher.cls` 自体には保存先変更機能は無いため、このモジュールのローカル
'      ヘルパーで FileCopy + Kill による移動を行っています）
'***************************************************************************************************
Sub Demo_DownloadWatcher_04_別名で保存()

    Dim outDir As String:  outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"
    Dim saveDir As String: saveDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Saved"

    '--- 1. テストサイトを開く ---
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext(TESTSITE_URL)

    '--- 2. DownloadWatcher 準備 & 監視開始 ---
    Dim dw As New exCDP_DownloadWatcher
    dw.Init browserTab
    dw.setDownloadBehavior(behavior:=behavior_allow, downloadPath:=outDir) = True
    ThrottleNetwork browserTab, DownloadKBps:=200   '200KB/s → 3MB は約15秒

    '--- 3. フォーム操作：3MB / WEBP ---
    Dim priorCount As Long: priorCount = dw.DownloadGuidCount
    Call SetDownloadForm(browserTab, Size:=3, Unit:="MB", Extension:="webp")

    '--- 4. クリック → 完了待ち ---
    browserTab.getElementByID("submit-button").click
    Debug.Print "[Demo04] ダウンロード完了待ち..."

    Dim guid As String
    guid = WaitForNewGuid(browserTab, dw, priorCount, TimeOutSecond:=30)
    If Len(guid) = 0 Or Not dw.AutoWaitSingleCompleted(guid, TimeOutSecond:=120) Then
        MsgBox "ダウンロード失敗。State=" & dw.DownloadProgressInfo(guid, downloadProgress_state), vbCritical
        UnthrottleNetwork browserTab
        browserTab.ThisCDPBrowser.quit
        Exit Sub
    End If

    Dim downloadedPath As String: downloadedPath = dw.DownloadProgressInfo(guid, downloadProgress_filePath)
    Debug.Print "[Demo04] ダウンロード完了: " & downloadedPath
    UnthrottleNetwork browserTab

    '--- 5. 日付付きファイル名で別フォルダに保存 ---
    Dim newName As String: newName = "testimg_" & Format(Now, "yyyymmdd_HHmmss") & ".webp"
    Dim saved As String:   saved = SaveDownloadedFileAs(downloadedPath, saveDir, newName)

    If Len(saved) > 0 Then
        Debug.Print "[Demo04] ○ 別名保存成功！ → " & saved
        MsgBox "別名保存成功！" & vbCrLf & saved, vbInformation
    Else
        Debug.Print "[Demo04] × 別名保存失敗"
        MsgBox "別名保存に失敗しました。", vbCritical
    End If

    browserTab.ThisCDPBrowser.quit

End Sub



'***************************************************************************************************
'   ■■■ Demo 05：複数ダウンロードを同時トリガー → AutoWaitMultiCompleted で一括待機 ■■■
'***************************************************************************************************
'* 機能　　：3件のダウンロードを jsEval で同時トリガーするデモです
'---------------------------------------------------------------------------------------------------
'* ダウンロード構成（同時）：
'   DL1:  3MB PNG
'   DL2:  5MB JPEG
'   DL3:  2MB GIF
'---------------------------------------------------------------------------------------------------
'* スロットリング: 500 KB/s → 合計10MB で約20秒（3件並行）
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - AutoWaitMultiCompleted(CompleteCount:=3) で3件全ての完了を一括待機できること（戻り値は完了数Long）
'   - PrintDownloadSummary（引数なし）で全GUIDのサマリーが出ること
'   - DownloadGuidList で全GUIDを列挙し、DownloadWillBeginInfo / DownloadProgressInfo で
'     GUIDごとの個別フィールド（url / suggestedFilename / state / totalBytes / filePath）を取得できること
'***************************************************************************************************
Sub Demo_DownloadWatcher_05_複数同時DLをまとめて待つ()

    Dim outDir As String
    outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- 1. テストサイトを開く ---
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext(TESTSITE_URL)
    Debug.Print "[Demo05] ページ読み込み完了"

    '--- 2. DownloadWatcher 準備 & 監視開始 ---
    Dim dw As New exCDP_DownloadWatcher
    dw.Init browserTab
    dw.setDownloadBehavior(behavior:=behavior_allow, downloadPath:=outDir) = True
    ThrottleNetwork browserTab, DownloadKBps:=500   '500KB/s → 合計10MB は約20秒

    '--- 3. JS で 3件のダウンロードを「同時」トリガー ---
    '    ① ページ内の CSRF トークンを取得
    '    ② 3本の fetch を同時実行（各 fetch は独立して非同期完了）
    '    ③ 各 fetch 完了後に <a click> → downloadWillBegin が 3回発火
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

    browserTab.jsEval js
    Debug.Print "[Demo05] 3件の fetch を同時実行しました..."
    Debug.Print "[Demo05] → fetch完了後に downloadWillBegin が 3回発火します"

    '--- 4. AutoWaitMultiCompleted で3件全ての完了を一括待機 ---
    If dw.AutoWaitMultiCompleted(CompleteCount:=3, TimeOutSecond:=180) = 3 Then
        Debug.Print "[Demo05] ○ 全て完了！"
        dw.PrintDownloadSummary   '※ 全GUID分のサマリーをイミディエイトウィンドウに出力

        '--- 5. DownloadGuidList で全GUIDを列挙し、個別フィールドを確認 ---
        Debug.Print "[Demo05] ────── GUID別の個別フィールド確認 ──────"

        Dim guids: guids = dw.DownloadGuidList
        Dim guid As Variant
        For Each guid In guids
            Debug.Print "[Demo05]   ── guid: " & Left(CStr(guid), 8) & "..."
            Debug.Print "[Demo05]     suggestedFilename = " & dw.DownloadWillBeginInfo(CStr(guid), downloadWillBegin_suggestedFilename)
            Debug.Print "[Demo05]     url               = " & dw.DownloadWillBeginInfo(CStr(guid), downloadWillBegin_Url)
            Debug.Print "[Demo05]     state             = " & dw.DownloadProgressInfo(CStr(guid), downloadProgress_state)
            Debug.Print "[Demo05]     totalBytes        = " & dw.DownloadProgressInfo(CStr(guid), downloadProgress_totalBytes)
            Debug.Print "[Demo05]     filePath          = " & dw.DownloadProgressInfo(CStr(guid), downloadProgress_filePath)
            Debug.Print ""
        Next guid

        MsgBox "全3件のダウンロードが完了！" & vbCrLf & _
               "完了数: " & dw.DownloadStateCount(state_completed) & " / " & dw.DownloadGuidCount & vbCrLf & _
               "保存先: " & outDir, vbInformation
    Else
        Debug.Print "[Demo05] × AutoWaitMultiCompleted タイムアウト"
        Debug.Print "[Demo05]   完了数: " & dw.DownloadStateCount(state_completed) & "/" & dw.DownloadGuidCount
        dw.PrintDownloadSummary
        MsgBox "タイムアウト。完了数: " & dw.DownloadStateCount(state_completed) & "/" & dw.DownloadGuidCount, vbCritical
    End If

    UnthrottleNetwork browserTab
    browserTab.ThisCDPBrowser.quit

End Sub



'***************************************************************************************************
'  ■■■ Demo 06：【直接リンク型】10MB PNG をリアルタイム進捗表示しながらダウンロード ■■■
'***************************************************************************************************
'* テストサイト：https://sample-img.lb-product.com/
'*   直接リンク型 → <a href="…URL"> をクリックするだけでブラウザ側からDLを開始します
'---------------------------------------------------------------------------------------------------
'* blob型（custom-img）との違い：
'   [blob型]    クリック → サーバー生成(数秒) → Blob受信 → <a click> → WillBegin
'               ・WillBeginが遅れて来る（Phase1に時間かかる）
'               ・転送はBlob→PCなので一瞬（Progressが0%→100%が一気に来る）
'               ・ThrottleNetwork は fetch フェーズに効く
'
'   [直接リンク型] クリック → 即 WillBegin → 転送(実際に時間かかる) → 完了
'               ・WillBeginがclick直後に来る（Phase1がほぼ0秒）
'               ・Progressが0%→…→100%と段階的に来る（リアルタイム観察できる）
'               ・ThrottleNetwork が転送速度に直接効く → こちらのほうが観察しやすい！
'---------------------------------------------------------------------------------------------------
'* スロットリング: 200 KB/s → 10MB で約50秒（体感できる速度）
'* 確認ポイント：
'   - downloadWillBegin がクリック直後に発火すること
'   - Progress が段階的に更新されること（blob型とは対照的）
'   - ThrottleNetwork が転送速度に直接反映される
'***************************************************************************************************
Sub Demo_DownloadWatcher_06_直接リンク型_10MBをリアルタイム進捗表示()

    Const SITE_URL  As String = "https://sample-img.lb-product.com/1gb/"
    Const DL_URL    As String = "https://sample-img.lb-product.com/wp-content/themes/hitchcock/images/1GB.png"
    Dim OUT_DIR   As String: OUT_DIR = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- 1. テストサイトを開く ---
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext(SITE_URL)
    Debug.Print "[Demo06] ページ読み込み完了"
    Debug.Print "[Demo06] → 直接リンク型：click直後に WillBegin が来る（blob型とは逆）"

    '--- 2. DownloadWatcher 準備 & 監視開始 ---
    Dim dw As New exCDP_DownloadWatcher
    dw.Init browserTab
    dw.setDownloadBehavior(behavior:=behavior_allow, downloadPath:=OUT_DIR) = True

    '--- 3. throttle（転送フェーズに直接効く） ---
    ThrottleNetwork browserTab, DownloadKBps:=200   '200KB/s → 10MB は約50秒

    '--- 4. リンクの href で直接クリック（jsEval で <a> 要素を取得）---
    '    sample-img はフォームではなく直接リンクなので、href一致 <a> を特定する
    Dim priorCount As Long: priorCount = dw.DownloadGuidCount
    Dim clickJs As String
    clickJs = "document.querySelector('a[href=""" & DL_URL & """]').click();"
    browserTab.jsEval clickJs
    Debug.Print "[Demo06] [" & Format(Time, "hh:mm:ss") & "] DLリンクをクリックしました"
    Debug.Print "[Demo06] → 直接リンク型なので downloadWillBegin はすぐに来るはずです..."

    '--- 5. downloadWillBegin を検知 ---
    Dim guid As String
    guid = WaitForNewGuid(browserTab, dw, priorCount, TimeOutSecond:=30)
    If Len(guid) = 0 Then
        Debug.Print "[Demo06] × downloadWillBegin タイムアウト"
        UnthrottleNetwork browserTab
        browserTab.ThisCDPBrowser.quit
        Exit Sub
    End If
    Debug.Print "[Demo06] [" & Format(Time, "hh:mm:ss") & "] ○ downloadWillBegin 受信！"
    Debug.Print "[Demo06]   suggestedFilename: " & dw.DownloadWillBeginInfo(guid, downloadWillBegin_suggestedFilename)

    '--- 6. カスタムポーリングで進捗をリアルタイム表示 ---
    Dim t As Double: t = Timer
    Dim lastPct As Long: lastPct = -1

    Do
        browserTab.ThisCDPBrowser.TakeEvents

        Dim state As String: state = dw.DownloadProgressInfo(guid, downloadProgress_state)
        Dim totalBytes As Double: totalBytes = dw.DownloadProgressInfo(guid, downloadProgress_totalBytes)
        Dim receivedBytes As Double: receivedBytes = dw.DownloadProgressInfo(guid, downloadProgress_receivedBytes)

        If totalBytes > 0 Then
            Dim pct As Long: pct = CLng(receivedBytes / totalBytes * 100)
            If pct <> lastPct Then
                Debug.Print "[Demo06]   " & Format(pct, "00") & "% | " _
                          & Format(receivedBytes / 1024 / 1024, "0.00") & " / " _
                          & Format(totalBytes / 1024 / 1024, "0.00") & " MB" _
                          & "  (" & Format(Timer - t, "0.0") & "s)"
                lastPct = pct
            End If
        End If

        If state = "completed" Then
            Debug.Print "[Demo06] ○ 完了！ 合計 " & Format(Timer - t, "0.0") & "秒"
            Debug.Print "[Demo06]   保存先: " & dw.DownloadProgressInfo(guid, downloadProgress_filePath)
            dw.PrintDownloadSummary guid
            MsgBox "完了！" & vbCrLf & _
                   "合計時間: " & Format(Timer - t, "0.0") & "秒" & vbCrLf & _
                   "保存先: " & dw.DownloadProgressInfo(guid, downloadProgress_filePath), vbInformation
            Exit Do
        End If

        If state = "canceled" Then
            Debug.Print "[Demo06] × キャンセル"
            Exit Do
        End If

        If (Timer - t) > 300 Then   '5分タイムアウト
            Debug.Print "[Demo06] × タイムアウト"
            Exit Do
        End If

        CDPHelpers.Sleep 0.3
    Loop

    UnthrottleNetwork browserTab
    browserTab.ThisCDPBrowser.quit

End Sub



'***************************************************************************************************
'  ■■■ Demo 07：【直接リンク型】複数サイズを同時クリック → AutoWaitMultiCompleted で一括待機 ■■■
'***************************************************************************************************
'* テストサイト：https://sample-img.lb-product.com/
'*   1MB / 10MB / 100MB の 3ファイルを「ほぼ同時」にクリックして並行DLする
'---------------------------------------------------------------------------------------------------
'* 直接リンク型 × 複数同時DL の確認ポイント：
'   - downloadWillBegin が3回すばやく、続けて来ること
'   - downloadProgress が3件のGUIDについて同時進行に来ること（並行転送の証拠）
'   - AutoWaitMultiCompleted(CompleteCount:=3) で3件全ての完了を待てること
'   - PrintDownloadSummary（引数なし）で全GUIDのサマリーが出ること
'---------------------------------------------------------------------------------------------------
'* スロットリング: 500 KB/s → 1MB≒2秒 / 10MB≒20秒 / 100MB≒200秒
'*   ※ 100MB は完了まで時間がかかります。確認目的なら途中でブレークしても OK。
'***************************************************************************************************
Sub Demo_DownloadWatcher_07_直接リンク型_複数ファイル同時DL()

    '--- 各DLの直URL ---
    Dim dlUrls(1 To 3) As String
    dlUrls(1) = "https://sample-img.lb-product.com/wp-content/themes/hitchcock/images/1MB.png"
    dlUrls(2) = "https://sample-img.lb-product.com/wp-content/themes/hitchcock/images/10MB.png"
    dlUrls(3) = "https://sample-img.lb-product.com/wp-content/themes/hitchcock/images/100MB.png"

    Const SITE_URL  As String = "https://sample-img.lb-product.com/10mb/"   '10MBページから起動
    Dim OUT_DIR   As String: OUT_DIR = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- 1. テストサイトを開く ---
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext(SITE_URL)
    Debug.Print "[Demo07] ページ読み込み完了"
    Debug.Print "[Demo07] → 直接リンク型 × 複数同時DL：3件のリンクをほぼ同時にクリック"

    '--- 2. DownloadWatcher 準備 & 監視開始 ---
    Dim dw As New exCDP_DownloadWatcher
    dw.Init browserTab
    dw.setDownloadBehavior(behavior:=behavior_allow, downloadPath:=OUT_DIR) = True
    ThrottleNetwork browserTab, DownloadKBps:=500   '500KB/s → 1MB≒2秒 / 10MB≒20秒 / 100MB≒200秒

    '--- 3. JS で 3件のダウンロードをほぼ同時にトリガー ---
    '    直接リンク型なので fetch は不要。<a> タグを動的生成して click() するだけ。
    Dim js As String
    js = "(function(){"
    js = js & "  var urls = ["
    js = js & "    '" & dlUrls(1) & "',"
    js = js & "    '" & dlUrls(2) & "',"
    js = js & "    '" & dlUrls(3) & "'"
    js = js & "  ];"
    js = js & "  urls.forEach(function(url) {"
    js = js & "    var a = document.createElement('a');"
    js = js & "    a.href = url;"
    js = js & "    a.download = '';"
    js = js & "    document.body.appendChild(a);"
    js = js & "    a.click();"
    js = js & "    document.body.removeChild(a);"
    js = js & "  });"
    js = js & "})();"

    Dim priorCount As Long: priorCount = dw.DownloadGuidCount
    browserTab.jsEval js
    Debug.Print "[Demo07] [" & Format(Time, "hh:mm:ss") & "] 3件のリンクをクリックしました"
    Debug.Print "[Demo07] → 直接リンク型なので downloadWillBegin が瞬時に 3回来るはずです..."

    '--- 4. 全3件の検知を待つ ---
    Dim t As Double: t = Timer
    Do
        browserTab.ThisCDPBrowser.TakeEvents
        Debug.Print "[Demo07]   検知待ち... DownloadGuidCount=" & dw.DownloadGuidCount & " / 3"
        If dw.DownloadGuidCount - priorCount >= 3 Then Exit Do
        If (Timer - t) > 120 Then
            Debug.Print "[Demo07] × 検出タイムアウト。検出数=" & (dw.DownloadGuidCount - priorCount)
            UnthrottleNetwork browserTab
            browserTab.ThisCDPBrowser.quit
            Exit Sub
        End If
        CDPHelpers.Sleep 0.5
    Loop
    Debug.Print "[Demo07] " & (dw.DownloadGuidCount - priorCount) & "件のダウンロードを検出！"

    '--- 5. 全3件の完了をまとめて待つ ---
    If dw.AutoWaitMultiCompleted(CompleteCount:=3, TimeOutSecond:=600) = 3 Then
        Debug.Print "[Demo07] ○ 全3件完了！ 合計 " & Format(Timer - t, "0.0") & "秒"
        dw.PrintDownloadSummary

        '--- 6. 全GUIDを列挙してサマリー表示 ---
        Dim guids: guids = dw.DownloadGuidList
        Dim guid As Variant
        Dim summary As String: summary = "全3件完了！" & vbCrLf & vbCrLf
        For Each guid In guids
            summary = summary & dw.DownloadWillBeginInfo(CStr(guid), downloadWillBegin_suggestedFilename) _
                    & " (" & Format(CDbl(dw.DownloadProgressInfo(CStr(guid), downloadProgress_totalBytes)) / 1024 / 1024, "0.0") & "MB)" _
                    & vbCrLf
        Next guid
        summary = summary & vbCrLf & "保存先: " & OUT_DIR

        MsgBox summary, vbInformation
    Else
        Debug.Print "[Demo07] × AutoWaitMultiCompleted タイムアウト"
        Debug.Print "[Demo07]   完了数: " & dw.DownloadStateCount(state_completed) & "/" & dw.DownloadGuidCount
        dw.PrintDownloadSummary
        MsgBox "タイムアウト。完了数: " & dw.DownloadStateCount(state_completed) & "/" & dw.DownloadGuidCount, vbCritical
    End If

    UnthrottleNetwork browserTab
    browserTab.ThisCDPBrowser.quit

End Sub



'***************************************************************************************************
'                           ■■■ ユーティリティ（共通利用） ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：custom-img.lb-product.com のダウンロードフォームに値を入力します
'---------------------------------------------------------------------------------------------------
'* 引数　　：browserTab CDPContext インスタンス
'            Size       ファイルサイズ（1~999、GBの場合は1まで）
'            Unit       "KB" / "MB" / "GB"
'            Extension  "png" / "jpeg" / "gif" / "webp"
'***************************************************************************************************
Private Sub SetDownloadForm(browserTab As CDPContext, Size As Integer, Unit As String, Extension As String)

    '--- サイズ入力 ---
    Dim sizeEl As CDPElement: Set sizeEl = browserTab.getElementByID("size")
    sizeEl.clearValue
    sizeEl.sendString CStr(Size)
    Debug.Print "[SetDownloadForm] size=" & Size

    '--- 単位セレクト ---
    browserTab.getElementByID("unit").setSelection Unit
    Debug.Print "[SetDownloadForm] unit=" & Unit

    '--- 拡張子セレクト ---
    browserTab.getElementByID("extension").setSelection Extension
    Debug.Print "[SetDownloadForm] extension=" & Extension

End Sub

'***************************************************************************************************
'* 機能　　：`Page.navigate`直後などで、`dw`が新規に検知した最新のguidを待ちます
'---------------------------------------------------------------------------------------------------
'* 返り値　：検知できたguid。タイムアウト時は空文字
'* 引数　　：browserTab   タブ
'            dw           exCDP_DownloadWatcher インスタンス
'            priorCount   トリガー前に確認しておいた DownloadGuidCount
'            TimeOutSecond タイムアウト秒数
'***************************************************************************************************
Private Function WaitForNewGuid(browserTab As CDPContext, dw As exCDP_DownloadWatcher, priorCount As Long, Optional TimeOutSecond As Double = 30) As String
    Dim t As Double: t = Timer

    Do
        browserTab.ThisCDPBrowser.TakeEvents

        If dw.DownloadGuidCount > priorCount Then
            Dim guids: guids = dw.DownloadGuidList
            WaitForNewGuid = CStr(guids(UBound(guids)))
            Exit Function
        End If

        If (Timer - t) > TimeOutSecond Then Exit Function   '空文字のまま返す

        CDPHelpers.Sleep 0.2
    Loop
End Function

'***************************************************************************************************
'* 機能　　：ダウンロード済みファイルを、指定フォルダ・ファイル名で保存し直します（移動＋リネーム）
'---------------------------------------------------------------------------------------------------
'* 返り値　：保存後のフルパス。失敗時は空文字
'* 引数　　：sourcePath   元のダウンロード済みファイルパス
'            destFolder   保存先フォルダ（無ければ作成）
'            newFileName  新しいファイル名
'***************************************************************************************************
Private Function SaveDownloadedFileAs(sourcePath As String, destFolder As String, newFileName As String) As String
    If Len(Dir(sourcePath)) = 0 Then Exit Function   '元ファイルが無ければ空文字で抜ける

    If Len(Dir(destFolder, vbDirectory)) = 0 Then MkDir destFolder

    Dim destPath As String
    destPath = destFolder & "\" & newFileName

    FileCopy sourcePath, destPath
    Kill sourcePath

    SaveDownloadedFileAs = destPath
End Function

'***************************************************************************************************
'* 機能　　：ダウンロード速度を意図的に制限します（`Network.emulateNetworkConditions`）
'---------------------------------------------------------------------------------------------------
'* 引数　　：browserTab      タブ
'            DownloadKBps    制限速度（KB/秒）
'***************************************************************************************************
Private Sub ThrottleNetwork(browserTab As CDPContext, DownloadKBps As Long)
    browserTab.ExecuteCDP "Network.enable"

    Dim params As New Dictionary
    params.Add "offline", False
    params.Add "latency", 0
    params.Add "downloadThroughput", DownloadKBps * 1024
    params.Add "uploadThroughput", DownloadKBps * 1024
    browserTab.ExecuteCDP "Network.emulateNetworkConditions", params
End Sub

'***************************************************************************************************
'* 機能　　：`ThrottleNetwork`で掛けた速度制限を解除します
'***************************************************************************************************
Private Sub UnthrottleNetwork(browserTab As CDPContext)
    Dim params As New Dictionary
    params.Add "offline", False
    params.Add "latency", 0
    params.Add "downloadThroughput", -1
    params.Add "uploadThroughput", -1
    browserTab.ExecuteCDP "Network.emulateNetworkConditions", params
End Sub
