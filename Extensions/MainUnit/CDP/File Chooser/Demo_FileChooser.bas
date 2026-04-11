Attribute VB_Name = "Demo_FileChooser"
'***************************************************************************************************
'       CDPexpansion_FileChooser 拡張 - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：`CDPexpansion_FileChooser.cls` を使ったFileChooserインターセプトのサンプルコードです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\File Chooser\CDPexpansion_FileChooser.cls
'* テストHTML：Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html
'---------------------------------------------------------------------------------------------------
'* 動作の仕組み：
'   ① fc.AddFilePath = "path" でファイルパスを事前登録（複数可）
'   ② fc.EnableIntercept でインターセプト有効化
'   ③ ブラウザ上でクリック等のトリガー
'   ④ OS のダイアログが開く前に CDP が横取り → Page.fileChooserOpened イベント発火
'   ⑤ fc.SetFileWait が backendNodeId を受け取り DOM.setFileInputFiles で注入
'   ⑥ input の change イベント発火 → JS の FileReader がファイルを読んでページに表示
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・「Page.fileChooserOpened」はブラウザが前面にある状態でないと発火しません
'   ・WORKSPACE_PATH を各自の環境に合わせて設定してください
'***************************************************************************************************
Option Explicit



'ワークスペースパス
'※ StarterWebScrapingKit のルートフォルダを入力してください
Private Const WORKSPACE_PATH As String = ""



'***************************************************************************************************
'                  ■■■ Demo 01：静的 input[type=file] への注入（新API） ■■■
'***************************************************************************************************
'* 機能　　：AddFilePath + SetFileWait の新しいAPIを使ったデモです
'---------------------------------------------------------------------------------------------------
'* テストページ：Test_FileChooser/index.html の ZONE A
'* 確認ポイント：
'   - fc.AddFilePath でパスを登録してから fc.SetFileWait を呼ぶフローが動くこと
'   - HTML 側の FileReader が change イベント経由でファイルを読み、表示されること
'***************************************************************************************************
Sub Demo_FileChooser_01_静的inputへ注入()

    Dim txtFile As String
    txtFile = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\File Chooser\sample.txt"

    If Dir(txtFile) = "" Then
        Open txtFile For Output As #1
        Print #1, "こんにちは！ FileChooser Interceptor のテストです。"
        Print #1, ""
        Print #1, "このファイルは VBA が注入しました。"
        Print #1, "でもファイルの中身を読んだのは VBA ではなく、"
        Print #1, "ブラウザの JavaScript FileReader API です！"
        Print #1, ""
        Print #1, "実行時刻: " & Now()
        Close #1
    End If

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動("file:///" & Replace(htmlPath, "\", "/"))
    browser.show

    '--- 2. FileChooser拡張の初期化 ---
    Dim fc As New CDPexpansion_FileChooser
    fc.Init browser

    '--- 3. ★新API：ファイルパスを事前登録 ---
    fc.AddFilePath = txtFile
    Debug.Print "[Demo01] 登録ファイル数: " & fc.FilePathCount

    '--- 4. インターセプトを有効化 ---
    fc.EnableEvents = True

    '--- 5. ファイル選択をトリガー ---
    Debug.Print "[Demo01] static-file-input をクリックします..."
    browser.getElementByID("static-file-input").click

    '--- 6. ★新API：パス引数なし、待機 & 注入 ---
    If Not fc.SetFiles(TimeoutSec:=10) Then
        MsgBox "ファイル注入失敗！ブラウザが前面にあるか確認してください。", vbCritical
        browser.quit
        Exit Sub
    End If

    '--- 7. FileReader の読み取り & アニメーション完了を待つ ---
    browser.sleep 1
    Debug.Print "[Demo01] 完了！FileReader がページに表示しています。"

    '--- 8. 後片付け ---
    fc.EnableEvents = False
    'fc.ClearFilePaths  ← 次回も同じファイルを使う場合は不要

End Sub



'***************************************************************************************************
'            ■■■ Demo 02：動的生成ダイアログへの注入（本命） ■■■
'***************************************************************************************************
'* テストページ：Test_FileChooser/index.html の ZONE B
'* 確認ポイント：
'   - DOM に file input が事前に存在しないケースでも AddFilePath + SetFileWait で動くこと
'***************************************************************************************************
Sub Demo_FileChooser_02_動的ダイアログへ注入()

    Dim txtFile As String
    txtFile = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\File Chooser\sample_dynamic.txt"

    If Dir(txtFile) = "" Then
        Open txtFile For Output As #1
        Print #1, "=========================================="
        Print #1, "  FileChooser Interceptor - 動的注入テスト"
        Print #1, "=========================================="
        Print #1, ""
        Print #1, "このファイルは JS で動的に生成されたダイアログへ注入されました。"
        Print #1, "DOM.setFileInputFiles では対応できないケースです！"
        Print #1, ""
        Print #1, "実行時刻: " & Now()
        Close #1
    End If

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動("file:///" & Replace(htmlPath, "\", "/"))
    browser.show

    '--- 2. FileChooser拡張の初期化 ---
    Dim fc As New CDPexpansion_FileChooser
    fc.Init browser

    '--- 3. ★新API：ファイルパスを事前登録 ---
    fc.AddFilePath = txtFile
    Debug.Print "[Demo02] 登録ファイル数: " & fc.FilePathCount

    '--- 4. インターセプトを有効化 ---
    fc.EnableEvents = True

    '--- 5. Zone B の動的ボタンをクリック（JS が createElement して click） ---
    Debug.Print "[Demo02] btn-dynamic をクリックします..."
    browser.getElementByID("btn-dynamic").click

    '--- 6. ★新API：パス引数なし、待機 & 注入 ---
    If Not fc.SetFiles(TimeoutSec:=10) Then
        MsgBox "ファイル注入失敗！ブラウザが前面にあるか確認してください。", vbCritical
        browser.quit
        Exit Sub
    End If

    '--- 7. 待機 ---
    browser.sleep 1
    Debug.Print "[Demo02] 完了！動的inputへの注入 & FileReader による表示に成功！"

    '--- 8. 後片付け ---
    fc.EnableEvents = False

End Sub



'***************************************************************************************************
'            ■■■ Demo 03：複数ファイルを順番に注入するシナリオ ■■■
'***************************************************************************************************
'* 機能　　：ClearFilePaths + AddFilePath で毎回ファイルを差し替えるデモです
'* 確認ポイント：
'   - ClearFilePaths → AddFilePath のサイクルで毎回異なるファイルを注入できること
'   - 毎ラウンド FileReader が新しい内容でページを更新すること
'***************************************************************************************************
Sub Demo_FileChooser_03_複数ファイル連続注入()

    '--- サンプルファイルを3つ用意 ---
    Dim files(1 To 3) As String
    Dim i As Integer
    For i = 1 To 3
        files(i) = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\File Chooser\test_" & i & ".txt"
        If Dir(files(i)) = "" Then
            Open files(i) For Output As #1
            Print #1, "==============================="
            Print #1, "  テストファイル " & i & " 番"
            Print #1, "==============================="
            Print #1, ""
            Print #1, "作成時刻: " & Now()
            Close #1
        End If
    Next i

    '--- ブラウザ起動 ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動("file:///" & Replace(htmlPath, "\", "/"))
    browser.show

    '--- FileChooser拡張初期化 ---
    Dim fc As New CDPexpansion_FileChooser
    fc.Init browser

    '--- 3ラウンド：毎回ファイルを差し替えて注入 ---
    fc.EnableEvents = True
    For i = 1 To 3
        Debug.Print "[Demo03] ラウンド " & i & " / 3 ..."

        browser.getElementByID("static-file-input").click

        '★ 単一用メソッドで、ファイルパスを登録
        Debug.Print "[Demo03]   登録: " & files(i)
        Dim ok As Boolean: ok = fc.SetFile(files(i), TimeoutSec:=10)
        If Not ok Then
            MsgBox "Round " & i & " で失敗しました。", vbCritical
            Exit For
        End If

        '★ FileReader の読み取り完了を待つ
        browser.sleep 2
        Debug.Print "[Demo03] ラウンド " & i & " 完了"
    Next i
    fc.EnableEvents = False

    Debug.Print "[Demo03] 3ファイルの連続注入が完了しました！"

End Sub
