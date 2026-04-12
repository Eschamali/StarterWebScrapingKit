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
'*     size      : 1 ? 999  （GBの場合は 1 まで）
'*     unit      : KB / MB / GB
'*     extension : png / jpeg / gif / webp
'---------------------------------------------------------------------------------------------------
'* 動作の仕組み：
'   ① dw.WatchStart で Browser.setDownloadBehavior を送信し、ダウンロード先・イベントを設定
'   ② ブラウザからサイトへアクセスしてフォームを操作
'   ③ JS の fetch → Blob → <a download> 方式のダウンロードが発生
'   ④ Browser.downloadWillBegin イベントで開始情報（ファイル名・URL）を取得
'   ⑤ Browser.downloadProgress で進捗を監視 → state="completed" で完了
'   ⑥ dw.DownloadedFilePath で取得
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・WORKSPACE_PATH を各自の環境に合わせて設定してください
'   ・ダウンロード先フォルダは WatchStart が自動作成します
'   ・大容量ファイルは時間がかかります（目安: 10MB ≒ 3?10秒）
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
'* 確認ポイント：
'   - WatchStart → フォーム送信 → WaitCompleted の流れで完了待機できること
'   - DownloadedFilePath に正しい保存先が格納されること
'   - 5MB のファイルでも SuggestedFilename / DownloadedFilePath が正しく取れること
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

    '--- 3. フォームを操作：5MB / PNG ---
    Call SetDownloadForm(browser, Size:=5, Unit:="MB", Extension:="png")

    '--- 4. ダウンロードボタンをクリック ---
    Debug.Print "[Demo01] ダウンロードボタンをクリック..."
    browser.getElementByID("submit-button").click

    '--- 5. 完了まで待機（最大 60 秒） ---
    Debug.Print "[Demo01] ダウンロード完了待ち..."
    If dw.WaitCompleted(TimeoutSec:=60) Then
        Debug.Print "[Demo01] ○ 完了！"
        Debug.Print "[Demo01]   SuggestedFilename : " & dw.SuggestedFilename
        Debug.Print "[Demo01]   DownloadedFilePath: " & dw.DownloadedFilePath
        MsgBox "ダウンロード完了！" & vbCrLf & _
               "ファイル名: " & dw.SuggestedFilename & vbCrLf & _
               "保存先: " & dw.DownloadedFilePath, vbInformation
    Else
        Debug.Print "[Demo01] × タイムアウトまたはキャンセル。State=" & dw.state
        MsgBox "ダウンロード失敗。State=" & dw.state, vbCritical
    End If

    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 02：複数サイズ・形式を連続ダウンロード ■■■
'***************************************************************************************************
'* 機能　　：サイズや形式を変えながら複数ファイルを連続ダウンロードするデモです
'---------------------------------------------------------------------------------------------------
'* ダウンロード構成：
'   Round 1: 1MB  JPEG
'   Round 2: 5MB  PNG
'   Round 3: 10MB WEBP
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - 毎ラウンド WatchStart を呼んでスナップショットをリセットしていること
'   - 完了後に次のダウンロードへ進めること
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

        '毎回 WatchStart でスナップショットリセット
        dw.WatchStart outDir

        'フォーム入力
        Call SetDownloadForm(browser, sizes(i), units(i), exts(i))

        'クリック → 完了待ち
        browser.getElementByID("submit-button").click
        Debug.Print "[Demo02]   完了待ち..."

        If dw.WaitCompleted(TimeoutSec:=120) Then
            Debug.Print "[Demo02]   ○ 保存完了: " & dw.DownloadedFilePath _
                      & " (" & Format(dw.TotalBytes / 1024 / 1024, "0.00") & " MB)"
        Else
            Debug.Print "[Demo02]   × 失敗 (State=" & dw.state & ")"
            MsgBox "Round " & i & " でダウンロード失敗。", vbCritical
            Exit For
        End If

        '次のダウンロードの前にボタンが再度有効になるまで少し待つ
        browser.sleep 1.5
    Next i

    Debug.Print "[Demo02] 全ラウンド完了！"
    MsgBox "3ファイルのダウンロードが完了しました！" & vbCrLf & outDir, vbInformation

    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 03：進捗を表示しながら大きめのファイルをダウンロード ■■■
'***************************************************************************************************
'* 機能　　：1GB PNG をダウンロードしながらイミディエイトウィンドウに進捗を表示するデモです
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - Progress / ReceivedBytes / TotalBytes が正しく更新されること
'   - 大容量でも WatchStart?WaitCompleted のフローが正しく動くこと
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

    '--- 3. フォームを操作：1GB / png ---
    Call SetDownloadForm(browser, Size:=1, Unit:="GB", Extension:="png")

    '--- 4. クリック ---
    browser.getElementByID("submit-button").click
    Debug.Print "[Demo03] Download triggered. Waiting with progress display..."

    '--- 5. カスタムポーリングループで進捗を表示 ---
    Dim t As Double: t = Timer
    Dim lastPct As Long: lastPct = -1

    Do
        browser.TakeEvents
        DoEvents

        If dw.HasStarted Then
            Dim pct As Long
            If dw.Progress >= 0 Then
                pct = CLng(dw.Progress * 100)
                If pct <> lastPct Then   '変化があったときだけ出力（ログが多くなりすぎないように）
                    Debug.Print "[Demo03]   " & Format(pct, "00") & "% ? " _
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
            MsgBox "完了！ " & Format(Timer - t, "0.0") & "秒" & vbCrLf & dw.DownloadedFilePath, vbInformation
            Exit Do
        End If

        If dw.state = "canceled" Then
            Debug.Print "[Demo03] × キャンセルされました"
            Exit Do
        End If

        If (Timer - t) > 180 Then
            Debug.Print "[Demo03] × タイムアウト"
            Exit Do
        End If

        browser.sleep 0.3
    Loop

    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 04：ダウンロード後に別名・別フォルダへ保存（SaveAs） ■■■
'***************************************************************************************************
'* 機能　　：3MB WEBP をダウンロードして、タイムスタンプ付きのファイル名で別フォルダへ保存するデモです
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - WaitCompleted 後に dw.SaveAs でリネーム & 移動できること
'***************************************************************************************************
Sub Demo_DownloadWatcher_04_別名で保存()

    Dim outDir As String:  outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"
    Dim SaveDir As String: SaveDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Saved"

    '--- 1. テストサイトを開く ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動(TESTSITE_URL)
    browser.waitForLoad

    '--- 2. DownloadWatcher 初期化 & 監視開始 ---
    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser
    dw.WatchStart outDir

    '--- 3. フォーム操作：3MB / WEBP ---
    Call SetDownloadForm(browser, Size:=3, Unit:="MB", Extension:="webp")

    '--- 4. クリック → 完了待ち ---
    browser.getElementByID("submit-button").click
    Debug.Print "[Demo04] ダウンロード完了待ち..."

    If Not dw.WaitCompleted(TimeoutSec:=60) Then
        MsgBox "ダウンロード失敗。State=" & dw.state, vbCritical
        browser.quit
        Exit Sub
    End If
    Debug.Print "[Demo04] ダウンロード完了: " & dw.DownloadedFilePath

    '--- 5. 日付付きのファイル名で別フォルダに保存 ---
    Dim newName As String: newName = "testimg_" & Format(Now, "yyyymmdd_HHmmss") & ".webp"
    Dim saved As String:   saved = dw.SaveAs(SaveDir, newName)

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
'                           ■■■ ユーティリティ（内部用） ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：custom-img.lb-product.com のダウンロードフォームに値を入力します
'---------------------------------------------------------------------------------------------------
'* 引数　　：browser    CDPBrowser インスタンス
'            Size       ファイルサイズ（1?999、GBの場合は1まで）
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
