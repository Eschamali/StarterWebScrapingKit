Attribute VB_Name = "Demo_DownloadWatcher"
'***************************************************************************************************
'       CDPexpansion_DownloadWatcher 拡張 - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：`CDPexpansion_DownloadWatcher.cls` を使ったダウンロード監視のサンプルコードです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\Download Watcher\CDPexpansion_DownloadWatcher.cls
'---------------------------------------------------------------------------------------------------
'* 動作の仕組み：
'   ① dw.WatchStart で Browser.setDownloadBehavior を送信し、ダウンロード先・イベントを設定
'   ② ダウンロードをトリガー（リンククリック等）
'   ③ Browser.downloadWillBegin イベントで開始情報（ファイル名・URL）を取得
'   ④ Browser.downloadProgress イベントで進捗を監視
'   ⑤ state="completed" になったら dw.WaitCompleted が True を返す
'   ⑥ dw.DownloadedFilePath or dw.SaveAs で取得
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・WORKSPACE_PATH を各自の環境に合わせて設定してください
'   ・ダウンロード先フォルダは WatchStart が自動作成します
'***************************************************************************************************
Option Explicit



'ワークスペースパス
'※ StarterWebScrapingKit のルートフォルダを入力してください
Private Const WORKSPACE_PATH As String = ""



'***************************************************************************************************
'              ■■■ Demo 01：data URL からシンプルにテキストファイルをダウンロード ■■■
'***************************************************************************************************
'* 機能　　：最も基本的なダウンロード監視デモです
'---------------------------------------------------------------------------------------------------
'* テスト方法：data:text/html ページ上のリンクをクリックしてダウンロードします
'* 確認ポイント：
'   - WatchStart → click → WaitCompleted の流れで完了待機できること
'   - DownloadedFilePath に正しい保存先が格納されること
'***************************************************************************************************
Sub Demo_DownloadWatcher_01_シンプルなテキスト保存()

    Dim outDir As String
    outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- 1. ブラウザ起動（data: URL でダウンロードリンクを持つシンプルなページ） ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動( _
        "data:text/html," & _
        "<html><body>" & _
        "<h2>Download Test</h2>" & _
        "<a id='dl' download='hello.txt' href='data:text/plain,Hello%20from%20VBA%20CDP!'>Download hello.txt</a>" & _
        "</body></html>")

    '--- 2. DownloadWatcher 拡張の初期化 ---
    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser

    '--- 3. 監視開始（ダウンロード先フォルダを指定） ---
    '        ★ ダウンロードトリガーの前に必ず呼ぶこと
    dw.WatchStart outDir
    Debug.Print "[Demo01] WatchStart: " & outDir

    '--- 4. ダウンロードリンクをクリック ---
    Debug.Print "[Demo01] Clicking download link..."
    browser.getElementByID("dl").click

    '--- 5. 完了まで待機（最大 30 秒） ---
    If dw.WaitCompleted(TimeoutSec:=30) Then
        Debug.Print "[Demo01] ? 完了！"
        Debug.Print "[Demo01]   SuggestedFilename : " & dw.SuggestedFilename
        Debug.Print "[Demo01]   DownloadedFilePath: " & dw.DownloadedFilePath
        MsgBox "ダウンロード完了！" & vbCrLf & dw.DownloadedFilePath, vbInformation
    Else
        Debug.Print "[Demo01] × タイムアウトまたはキャンセル。State=" & dw.State
        MsgBox "ダウンロード失敗。State=" & dw.State, vbCritical
    End If

    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 02：PDF を別名保存（SaveAs を使ったリネーム保存） ■■■
'***************************************************************************************************
'* 機能　　：ダウンロード完了後、SaveAs で別フォルダに日付付きファイル名で保存するデモです
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - WaitCompleted 後に SaveAs でリネーム & 移動できること
'***************************************************************************************************
Sub Demo_DownloadWatcher_02_別名で保存()

    Dim outDir As String: outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"
    Dim saveDir As String: saveDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Saved"

    '--- 1. ブラウザ起動（CSV形式のdata:URLリンク） ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動( _
        "data:text/html," & _
        "<html><body>" & _
        "<h2>CSV Download Test</h2>" & _
        "<a id='dl' download='data.csv' " & _
        "href='data:text/csv,%E5%90%8D%E5%89%8D%2C%E5%B9%B4%E9%BD%A2%0A%E7%94%B0%E4%B8%AD%2C30%0A%E5%B1%B1%E7%94%B0%2C25'>Download data.csv</a>" & _
        "</body></html>")

    '--- 2. DownloadWatcher 初期化 ---
    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser

    '--- 3. 監視開始 ---
    dw.WatchStart outDir
    Debug.Print "[Demo02] WatchStart: " & outDir

    '--- 4. ダウンロードリンクをクリック ---
    Debug.Print "[Demo02] Clicking CSV download link..."
    browser.getElementByID("dl").click

    '--- 5. 完了まで待機 ---
    If Not dw.WaitCompleted(TimeoutSec:=30) Then
        MsgBox "ダウンロード失敗。State=" & dw.State, vbCritical
        browser.quit
        Exit Sub
    End If
    Debug.Print "[Demo02] Download complete: " & dw.DownloadedFilePath

    '--- 6. 日付付きのファイル名で別フォルダに保存 ---
    Dim newName As String: newName = "report_" & Format(Now, "yyyymmdd_HHmmss") & ".csv"
    Dim saved As String: saved = dw.SaveAs(saveDir, newName)

    If Len(saved) > 0 Then
        Debug.Print "[Demo02] ? 別名保存完了！ → " & saved
        MsgBox "別名保存完了！" & vbCrLf & saved, vbInformation
    Else
        Debug.Print "[Demo02] × 別名保存失敗"
        MsgBox "SaveAs に失敗しました。", vbCritical
    End If

    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 03：複数ファイルを連続してダウンロード ■■■
'***************************************************************************************************
'* 機能　　：ダウンロードを複数回繰り返すシナリオです
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - WatchStart → click → WaitCompleted のサイクルを繰り返せること
'   - 毎回 ResetState が適切に働き、前回の情報が混在しないこと
'***************************************************************************************************
Sub Demo_DownloadWatcher_03_連続ダウンロード()

    Dim outDir As String: outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- ブラウザ起動（3つのダウンロードリンクを持つページ） ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動( _
        "data:text/html," & _
        "<html><body>" & _
        "<h2>Multi Download Test</h2>" & _
        "<a id='dl1' download='file1.txt' href='data:text/plain,File1'>Download file1.txt</a><br>" & _
        "<a id='dl2' download='file2.txt' href='data:text/plain,File2'>Download file2.txt</a><br>" & _
        "<a id='dl3' download='file3.txt' href='data:text/plain,File3'>Download file3.txt</a>" & _
        "</body></html>")

    '--- DownloadWatcher 初期化 ---
    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser

    '--- 3ファイルを順番にダウンロード ---
    Dim dlIds(1 To 3) As String
    dlIds(1) = "dl1": dlIds(2) = "dl2": dlIds(3) = "dl3"

    Dim i As Integer
    For i = 1 To 3
        Debug.Print "[Demo03] ─── ダウンロード " & i & " / 3 ───"

        '毎回 WatchStart でスナップショットをリセット
        dw.WatchStart outDir
        Debug.Print "[Demo03]   WatchStart: " & outDir

        'クリック
        browser.getElementByID(dlIds(i)).click
        Debug.Print "[Demo03]   Clicked: " & dlIds(i)

        '完了待ち
        If dw.WaitCompleted(TimeoutSec:=30) Then
            Debug.Print "[Demo03]   ? 完了: " & dw.DownloadedFilePath
        Else
            Debug.Print "[Demo03]   × 失敗 (State=" & dw.State & ")"
            MsgBox "Round " & i & " でダウンロード失敗。", vbCritical
            Exit For
        End If

        '少し待ってから次へ
        browser.sleep 0.5
    Next i

    Debug.Print "[Demo03] 3ファイルのダウンロード完了！"
    MsgBox "3ファイルのダウンロードが完了しました！" & vbCrLf & outDir, vbInformation

    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 04：進捗率を表示しながらダウンロード（重いファイル向け） ■■■
'***************************************************************************************************
'* 機能　　：WaitCompleted を使わず、独自ループで進捗を表示するデモです
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - Progress プロパティ、ReceivedBytes、TotalBytes が正しく更新されること
'   - 大きなファイルでも完了まで監視できること
'***************************************************************************************************
Sub Demo_DownloadWatcher_04_進捗表示()

    Dim outDir As String: outDir = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\Download Watcher\Downloads"

    '--- ブラウザ起動（ダウンロード用テストページ） ---
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動( _
        "data:text/html," & _
        "<html><body>" & _
        "<h2>Progress Demo</h2>" & _
        "<a id='dl' download='sample.txt' href='data:text/plain,Hello'>Download</a>" & _
        "</body></html>")

    Dim dw As New CDPexpansion_DownloadWatcher
    dw.Init browser
    dw.WatchStart outDir

    browser.getElementByID("dl").click
    Debug.Print "[Demo04] ダウンロード開始待ち..."

    '--- カスタムポーリングループで進捗を表示 ---
    Dim timeoutSec As Long: timeoutSec = 60
    Dim t As Double: t = Timer

    Do
        browser.TakeEvents
        DoEvents

        If dw.HasStarted And Not dw.IsCompleted Then
            If dw.Progress >= 0 Then
                Debug.Print "[Demo04]   進捗: " & Format(dw.Progress * 100, "0.0") & "%" & _
                            " (" & Format(dw.ReceivedBytes / 1024, "0.0") & " KB / " & _
                            Format(dw.TotalBytes / 1024, "0.0") & " KB)"
            Else
                Debug.Print "[Demo04]   受信: " & Format(dw.ReceivedBytes / 1024, "0.0") & " KB (合計不明)"
            End If
        End If

        If dw.IsCompleted Then
            Debug.Print "[Demo04] ? 完了！ " & dw.DownloadedFilePath
            MsgBox "完了！ " & dw.DownloadedFilePath, vbInformation
            Exit Do
        End If

        If dw.State = "canceled" Then
            Debug.Print "[Demo04] × キャンセルされました"
            Exit Do
        End If

        If (Timer - t) > timeoutSec Then
            Debug.Print "[Demo04] × タイムアウト"
            Exit Do
        End If

        browser.sleep 0.3
    Loop

    'browser.quit

End Sub
