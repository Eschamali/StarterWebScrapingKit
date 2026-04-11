Attribute VB_Name = "Demo_FileChooser"
'***************************************************************************************************
'       CDPexpansion_FileChooser 拡張 ? デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：`CDPexpansion_FileChooser.cls` を使ったFileChooserインターセプトのサンプルコードです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\CDP\File Chooser\CDPexpansion_FileChooser.cls
'* テストHTML：ForDevelopers\OperationCheck\CDP\TestHtml\Test_FileChooser\index.html
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・実行前に「CDPexpansion_FileChooser.cls」をVBAプロジェクトに取り込んでください
'   ・「Page.fileChooserOpened」はブラウザが前面にある状態でないと発火しません
'   ・ブラウザのshowメソッドで前面表示してから使用してください
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
'   - ページ上に「読み取ったファイルの内容」がどーん！と表示されること
'---------------------------------------------------------------------------------------------------
'* 準備：適当なテキストファイル（sample.txt等）をあらかじめどこかに作成しておいてください
'***************************************************************************************************
Sub Demo_FileChooser_01_静的inputへ注入()

    '--- 設定 ---
    Dim txtFile As String
    txtFile = WORKSPACE_PATH & "\Extensions\CDP\File Chooser\sample.txt"

    '--- テキストファイルがなければサンプル作成 ---
    If Dir(txtFile) = "" Then
        Open txtFile For Output As #1
        Print #1, "こんにちは！ FileChooser Interceptor のテストです。"
        Print #1, "このファイルは VBA から自動的にブラウザへ注入されました。"
        Print #1, "時刻: " & Now()
        Close #1
        Debug.Print "サンプルファイルを作成しました: " & txtFile
    End If

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String: htmlPath = WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_FileChooser\index.html"
    Dim browser As CDPBrowser: Set browser = 設定シートからのCDP起動("file:///" & Replace(htmlPath, "\", "/"))

    '--- 2. ブラウザを前面表示（fileChooserOpened 発火に必要） ---
    browser.show

    '--- 3. FileChooser拡張の初期化 ---
    Dim fc As New CDPexpansion_FileChooser
    fc.Init browser

    '--- 4. インターセプトを有効化 ---
    fc.EnableIntercept

    '--- 5. Zone A の file input をクリックしてダイアログをトリガー ---
    Debug.Print "[Demo01] Clicking static file input..."
    browser.getElementByID("static-file-input").click

    '--- 6. ダイアログを横取りしてファイルを注入（内部でイベント待ち） ---
    Dim ok As Boolean: ok = fc.SetFile(txtFile, TimeoutSec:=10)
    If Not ok Then
        MsgBox "ファイル注入失敗！ブラウザが前面にあるか確認してください。", vbCritical
        browser.quit
        Exit Sub
    End If

    '--- 7. ファイル内容を読み取ってブラウザに表示（どーん！） ---
    Dim content As String: content = ReadTextFile(txtFile)
    Call ShowResultOnPage(browser, txtFile, content)

    '--- 8. 後片付け ---
    fc.DisableIntercept
    Debug.Print "[Demo01] 完了！ ブラウザにファイル内容が表示されました。"

    'ブラウザはそのままにして確認できるようにする（閉じたい場合は↓を有効化）
    'browser.quit

End Sub



'***************************************************************************************************
'            ■■■ Demo 02：動的生成ダイアログへのファイル注入（本命） ■■■
'***************************************************************************************************
'* 機能　　：JSでその場で生成されるファイル選択ダイアログにファイルを注入するデモです。
'            DOM.setFileInputFiles では対応不可能なケースが、これで解決できます。
'---------------------------------------------------------------------------------------------------
'* テストページ：Test_FileChooser/index.html の ZONE B を使います
'* 確認ポイント：
'   - DOM に file input 要素が事前に存在しない状態でも注入できること
'   - ページ上に「読み取ったファイルの内容」がどーん！と表示されること
'***************************************************************************************************
Sub Demo_FileChooser_02_動的ダイアログへ注入()

    '--- 設定 ---
    Dim txtFile As String
    txtFile = WORKSPACE_PATH & "\Extensions\CDP\File Chooser\sample_dynamic.txt"

    '--- テキストファイルがなければサンプル作成 ---
    If Dir(txtFile) = "" Then
        Open txtFile For Output As #1
        Print #1, "=========================================="
        Print #1, "  FileChooser Interceptor ? 動的注入テスト"
        Print #1, "=========================================="
        Print #1, ""
        Print #1, "このファイルはJSで動的生成されたダイアログにを"
        Print #1, "VBAのFileChooserインターセプトで横取りして注入されました。"
        Print #1, ""
        Print #1, "DOM.setFileInputFiles では対応できないケースです！"
        Print #1, ""
        Print #1, "実行時刻: " & Now()
        Close #1
        Debug.Print "サンプルファイルを作成しました: " & txtFile
    End If

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String: htmlPath = WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_FileChooser\index.html"
    Dim browser As CDPBrowser: Set browser = 設定シートからのCDP起動("file:///" & Replace(htmlPath, "\", "/"))

    '--- 2. ブラウザを前面表示（fileChooserOpened 発火に必要） ---
    browser.show

    '--- 3. FileChooser拡張の初期化 ---
    Dim fc As New CDPexpansion_FileChooser
    fc.Init browser

    '--- 4. インターセプトを有効化 ---
    fc.EnableIntercept

    '--- 5. Zone B の「動的生成」ボタンをクリック（JS経由でダイアログが開く） ---
    Debug.Print "[Demo02] Clicking dynamic file dialog button..."
    browser.getElementByID("btn-dynamic").click

    '--- 6. ダイアログを横取りしてファイルを注入（内部でイベント待ち） ---
    Dim ok As Boolean: ok = fc.SetFile(txtFile, TimeoutSec:=10)
    If Not ok Then
        MsgBox "ファイル注入失敗！ブラウザが前面にあるか確認してください。", vbCritical
        browser.quit
        Exit Sub
    End If

    '--- 7. ファイル内容を読み取ってブラウザに表示（どーん！） ---
    Dim content As String: content = ReadTextFile(txtFile)
    Call ShowResultOnPage(browser, txtFile, content)

    '--- 8. 後片付け ---
    fc.DisableIntercept
    Debug.Print "[Demo02] 完了！ 動的ダイアログへの注入成功！"

    'browser.quit

End Sub



'***************************************************************************************************
'                  ■■■ Demo 03：複数ファイルを順番に注入するシナリオ ■■■
'***************************************************************************************************
'* 機能　　：インターセプトを繰り返し使って、複数ファイルを順番に扱うデモです
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - EnableIntercept → SetFile → ResetState のサイクルを繰り返せること
'   - 2回目・3回目も正常に動作すること
'***************************************************************************************************
Sub Demo_FileChooser_03_複数ファイル連続注入()

    '--- サンプルファイルを3つ用意 ---
    Dim files(1 To 3) As String
    Dim i As Integer
    For i = 1 To 3
        files(i) = WORKSPACE_PATH & "\Extensions\CDP\File Chooser\test_" & i & ".txt"
        If Dir(files(i)) = "" Then
            Open files(i) For Output As #1
            Print #1, "テストファイル " & i & " 番"
            Print #1, "作成時刻: " & Now()
            Close #1
        End If
    Next i

    '--- ブラウザ起動 ---
    Dim htmlPath As String: htmlPath = WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_FileChooser\index.html"
    Dim browser As CDPBrowser: Set browser = 設定シートからのCDP起動("file:///" & Replace(htmlPath, "\", "/"))
    browser.show

    '--- FileChooser拡張初期化 ---
    Dim fc As New CDPexpansion_FileChooser
    fc.Init browser

    '--- 3回繰り返し ---
    For i = 1 To 3
        Debug.Print "[Demo03] ラウンド " & i & " / 3 ..."

        '毎回インターセプトを有効化
        fc.EnableIntercept

        'クリックしてダイアログをトリガー
        browser.getElementByID("static-file-input").click

        'ファイル注入
        Dim ok As Boolean: ok = fc.SetFile(files(i), TimeoutSec:=10)
        If Not ok Then
            MsgBox "Round " & i & " で失敗しました。", vbCritical
            Exit For
        End If

        'ブラウザに内容をどーん！と表示
        Dim content As String: content = ReadTextFile(files(i))
        Call ShowResultOnPage(browser, files(i), content)

        '少し待ってから次へ
        browser.sleep 1.5
        Debug.Print "[Demo03] ラウンド " & i & " 完了"
    Next i

    fc.DisableIntercept
    Debug.Print "[Demo03] 3ファイルの連続注入が完了しました！"

End Sub



'***************************************************************************************************
'                           ■■■ ユーティリティ（内部用） ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：テキストファイルをUTF-8で読み取ります
'***************************************************************************************************
Private Function ReadTextFile(filePath As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object:  Set ts = fso.OpenTextFile(filePath, 1, False, -1)  '-1 = TristateTrue (Unicode/UTF-8)
    ReadTextFile = ts.ReadAll
    ts.Close
End Function

'***************************************************************************************************
'* 機能　　：ブラウザのデモHTMLページにファイル内容を「どーん！」と表示します
'            index.html の showFileContentFromVBA() JS関数を呼び出します
'***************************************************************************************************
Private Sub ShowResultOnPage(browser As CDPBrowser, filePath As String, content As String)

    Dim fileName As String: fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    'シングルクォートと改行をエスケープしてJSに渡す
    Dim safeContent As String: safeContent = content
    safeContent = Replace(safeContent, "\", "\\")
    safeContent = Replace(safeContent, "'", "\'")
    safeContent = Replace(safeContent, Chr(13) & Chr(10), "\n")
    safeContent = Replace(safeContent, Chr(10), "\n")
    safeContent = Replace(safeContent, Chr(13), "\n")

    Dim js As String
    js = "window.showFileContentFromVBA('" & fileName & "', '" & safeContent & "')"

    browser.jsEval js
    Debug.Print "[ShowResult] Content displayed on page. fileName=" & fileName

End Sub
