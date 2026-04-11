Attribute VB_Name = "Demo_FileChooser"
'***************************************************************************************************
'       CDPexpansion_FileChooser 拡張 - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：`CDPexpansion_FileChooser.cls` を使ったFileChooserインターセプトのサンプルコードです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\File Chooser\CDPexpansion_FileChooser.cls
'* テストHTML：ForDevelopers\OperationCheck\CDP\TestHtml\Test_FileChooser\index.html
'---------------------------------------------------------------------------------------------------
'* 動作の仕組み：
'   ① VBA が FileChooserインターセプトを有効化
'   ② ブラウザ上でファイル選択をトリガー（click）
'   ③ OS のダイアログが開く前に CDP が横取り → Page.fileChooserOpened イベント発火
'   ④ VBA が「backendNodeId」を受け取り DOM.setFileInputFiles でファイルパスを注入
'   ⑤ input 要素の「change」イベントが発火
'   ⑥ HTML 側の JavaScript「FileReader」がファイルを読み取り、ページに表示
'      ★ ここが「ちゃんと input タグらしい機能」！ VBA はファイル内容を読まない
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・「Page.fileChooserOpened」はブラウザが前面にある状態でないと発火しません
'   ・ブラウザの show メソッドで前面表示してから使用してください
'***************************************************************************************************
Option Explicit



'ワークスペースパス
'※StarterWebScrapingKitのルートフォルダ を入力してください
Private Const WORKSPACE_PATH As String = ""



'***************************************************************************************************
'                  ■■■ Demo 01：静的 input[type=file] へのファイル注入 ■■■
'***************************************************************************************************
'* 機能　　：HTMLに最初から存在する <input type="file"> をCDP経由で操作するデモです
'---------------------------------------------------------------------------------------------------
'* テストページ：Test_FileChooser/index.html の ZONE A を使います
'* 確認ポイント：
'   - ファイル選択ダイアログが表示されずにファイルが注入されること
'   - HTML の FileReader が change イベントを受け取りページにどーん！と表示されること
'   - VBA はファイル内容を一切読まないこと（FileReader が全部やる）
'***************************************************************************************************
Sub Demo_FileChooser_01_静的inputへ注入()

    Dim txtFile As String
    txtFile = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\File Chooser\sample.txt"

    '--- テキストファイルがなければサンプル作成 ---
    If Dir(txtFile) = "" Then
        Open txtFile For Output As #1
        Print #1, "こんにちは！ FileChooser Interceptor のテストです。"
        Print #1, ""
        Print #1, "このファイルは VBA から注入されました。"
        Print #1, "でも内容を読んだのは VBA ではなく、"
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

    '--- 2. ブラウザを前面表示（fileChooserOpened 発火に必要） ---
    browser.show

    '--- 3. FileChooser拡張の初期化 ---
    Dim fc As New CDPexpansion_FileChooser
    fc.Init browser

    '--- 4. インターセプトを有効化 ---
    fc.EnableIntercept

    '--- 5. Zone A の file input をクリックしてダイアログをトリガー ---
    Debug.Print "[Demo01] static-file-input をクリックします..."
    browser.getElementByID("static-file-input").click

    '--- 6. fileChooserOpened を待ち、ファイルパスを注入 ---
    '        → change イベント発火 → HTML側 FileReader が読み取り → ページに表示
    Dim ok As Boolean: ok = fc.SetFile(txtFile, TimeoutSec:=10)

    If Not ok Then
        MsgBox "ファイル注入失敗！ブラウザが前面にあるか確認してください。", vbCritical
        browser.quit
        Exit Sub
    End If

    '--- 7. FileReader の読み取り & アニメーション完了を少し待つ ---
    browser.sleep 1
    Debug.Print "[Demo01] 完了！FileReader がページにどーん！と表示しているはずです。"

    '--- 8. 後片付け ---
    fc.DisableIntercept

    'browser.quit  ← 確認したいのでそのまま残す

End Sub



'***************************************************************************************************
'            ■■■ Demo 02：動的生成ダイアログへのファイル注入（本命） ■■■
'***************************************************************************************************
'* 機能　　：JSでその場で生成されるファイル選択ダイアログにファイルを注入するデモです
'            DOM.setFileInputFiles では対応不可能なケースが、これで解決できます
'---------------------------------------------------------------------------------------------------
'* テストページ：Test_FileChooser/index.html の ZONE B を使います
'* 確認ポイント：
'   - DOM に file input 要素が事前に存在しない状態でも注入できること
'   - 注入後に HTML側 FileReader の change イベントが発火してページに表示されること
'   - VBA はファイル内容を一切読まないこと（FileReader が全部やる）
'***************************************************************************************************
Sub Demo_FileChooser_02_動的ダイアログへ注入()

    Dim txtFile As String
    txtFile = WORKSPACE_PATH & "\Extensions\MainUnit\CDP\File Chooser\sample_dynamic.txt"

    '--- テキストファイルがなければサンプル作成 ---
    If Dir(txtFile) = "" Then
        Open txtFile For Output As #1
        Print #1, "=========================================="
        Print #1, "  FileChooser Interceptor - 動的注入テスト"
        Print #1, "=========================================="
        Print #1, ""
        Print #1, "このファイルはJSで動的に生成されたダイアログへ注入されました。"
        Print #1, ""
        Print #1, "DOM.setFileInputFiles では対応できないケースです。"
        Print #1, ""
        Print #1, "FileChooser Intercept が横取りして、"
        Print #1, "DOM.setFileInputFiles で注入 → change イベント発火"
        Print #1, "→ FileReader が読んでページに表示！"
        Print #1, ""
        Print #1, "実行時刻: " & Now()
        Close #1
    End If

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browser As CDPBrowser
    Set browser = 設定シートからのCDP起動("file:///" & Replace(htmlPath, "\", "/"))

    '--- 2. ブラウザを前面表示 ---
    browser.show

    '--- 3. FileChooser拡張の初期化 ---
    Dim fc As New CDPexpansion_FileChooser
    fc.Init browser

    '--- 4. インターセプトを有効化 ---
    fc.EnableIntercept

    '--- 5. Zone B の「動的生成」ボタンをクリック ---
    '        JS が createElement('input') → click() を実行
    Debug.Print "[Demo02] btn-dynamic をクリックします（動的inputが生成されます）..."
    browser.getElementByID("btn-dynamic").click

    '--- 6. fileChooserOpened を待ち、ファイルパスを注入 ---
    '        → change イベント発火 → event delegation で拾って FileReader 起動
    Dim ok As Boolean: ok = fc.SetFile(txtFile, TimeoutSec:=10)

    If Not ok Then
        MsgBox "ファイル注入失敗！ブラウザが前面にあるか確認してください。", vbCritical
        browser.quit
        Exit Sub
    End If

    '--- 7. FileReader の読み取り & アニメーション完了を少し待つ ---
    browser.sleep 1
    Debug.Print "[Demo02] 完了！動的inputへの注入 & FileReader による表示に成功！"

    '--- 8. 後片付け ---
    fc.DisableIntercept

End Sub



'***************************************************************************************************
'                  ■■■ Demo 03：複数ファイルを順番に注入するシナリオ ■■■
'***************************************************************************************************
'* 機能　　：インターセプトを繰り返し使って、複数ファイルを順番に注入するデモです
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - EnableIntercept → SetFile のサイクルを繰り返せること
'   - 毎回 change イベントが発火して FileReader が新しい内容を表示すること
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

    '--- 3回繰り返し ---
    For i = 1 To 3
        Debug.Print "[Demo03] ラウンド " & i & " / 3 ..."

        fc.EnableIntercept
        browser.getElementByID("static-file-input").click

        Dim ok As Boolean: ok = fc.SetFile(files(i), TimeoutSec:=10)
        If Not ok Then
            MsgBox "Round " & i & " で失敗しました。", vbCritical
            Exit For
        End If

        '★ FileReader の読み取り & ページへの表示を待つ
        browser.sleep 2
        Debug.Print "[Demo03] ラウンド " & i & " 完了 → FileReader が表示中"
    Next i

    fc.DisableIntercept
    Debug.Print "[Demo03] 3ファイルの連続注入が完了しました！"

End Sub
