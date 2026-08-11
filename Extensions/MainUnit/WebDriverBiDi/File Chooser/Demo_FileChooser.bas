Attribute VB_Name = "Demo_FileChooser"
'***************************************************************************************************
'       exBiDi_FileChooser 拡張（WebDriver BiDi版） - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：WebDriver BiDiによるファイル添付の2つのアプローチのサンプルコードです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\WebDriverBiDi\File Chooser\exBiDi_FileChooser.cls
'* テストHTML：Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html
'             （CDP版と同じHTMLをそのまま再利用しています。プロトコル依存のJSは含まれていません）
'---------------------------------------------------------------------------------------------------
'* 2つのアプローチ（CDP版と同じ構造）：
'   ①`exBiDi_FileChooser.SetFileInputFiles`（直接注入）
'      ・既に存在する要素をCSSセレクタで直接指定し、ダイアログを介さず添付する
'      ・`browsingContext.locateNodes`で`sharedId`を取得 → `input.setFiles`
'
'   ②`input.fileDialogOpened`イベントの監視（横取り）
'      ・`.click()`等で実際にファイル選択ダイアログが開こうとした瞬間を検知し、
'        代わりに`input.setFiles`で横取り添付する
'      ・input要素を事前に取得できない（動的生成）ケースで必要になる
'      ・キャンセルモード／添付忘れリトライも、CDP版と同様に用意しています
'---------------------------------------------------------------------------------------------------
'* 重要：OSダイアログの抑制について
'   `input.fileDialogOpened`の通知自体は常に届きますが、実際にOSダイアログを開かせず抑制するには、
'   `StartBiDiModeContext`の`sessionCapabilitiesRequest`で`unhandledPromptBehavior.file`を
'   `"ignore"`以外（既定は`"ignore"`）に設定しておく必要があります。
'   このモジュールでは`BuildFileDialogCapabilities`ヘルパーで組み立てています。
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・`input.setFiles` / `input.fileDialogOpened`はBiDiの比較的新しい機能です。ブラウザのバージョン
'     によっては未対応の場合があります
'   ・WORKSPACE_PATH をご自身の環境に合わせて設定してください
'***************************************************************************************************
Option Explicit



'ワークスペースパス
'※ StarterWebScrapingKit のルートフォルダを入力してください
Private Const WORKSPACE_PATH As String = "C:\Users\mokkun\Tools\StarterWebScrapingKit"

Private Const SAMPLE_DIR As String = "\Extensions\MainUnit\WebDriverBiDi\File Chooser"

Private CharConv As New CharacterCodeConversion



'***************************************************************************************************
'      ■■■ Demo 01：input.setFiles による直接注入（SINGLE / MULTI） ■■■
'***************************************************************************************************
'* 機能　　：静的な`<input type="file">`へ、CSSセレクタ指定で直接ファイルを注入するデモです
'---------------------------------------------------------------------------------------------------
'* テストページ：Test_FileChooser/index.html の「1?? SINGLE FILE INJECTION」「2?? MULTI FILE INJECTION」
'* 確認ポイント：
'   - `exBiDi_FileChooser.SetFileInputFiles`（1件のCollection）で singleInput への単一注入が成功すること
'   - 同（2件以上のCollection）で multiInput への複数注入が成功すること
'   - `input.fileDialogOpened`の監視は不要（ダイアログイベントに依存しない）
'***************************************************************************************************
Sub Demo_FileChooser_01_直接注入_単一と複数()

    Dim file1 As String: file1 = EnsureSampleFile("inject_single.txt", "SINGLE FILE INJECTION のテストファイルです。")
    Dim file2 As String: file2 = EnsureSampleFile("inject_multi_1.txt", "MULTI FILE INJECTION テストファイル 1")
    Dim file3 As String: file3 = EnsureSampleFile("inject_multi_2.txt", "MULTI FILE INJECTION テストファイル 2")

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browserTab As WebDriverBiDiContext
    Set browserTab = ShSetting01_StartBrowser.StartBiDiModeContext("file:///" & Replace(htmlPath, "\", "/"))

    '--- 2. FileChooser拡張の準備 ---
    Dim fc As New exBiDi_FileChooser
    fc.Init browserTab

    '--- 3. SINGLE FILE INJECTION：1件のCollectionを直接注入 ---
    Debug.Print "[Demo01] #singleInput へ 1件のファイルを直接注入..."
    Dim singleFiles As New Collection
    singleFiles.Add file1

    If fc.SetFileInputFiles("#singleInput", singleFiles) And GetInputFileCount(browserTab, "#singleInput") = 1 Then
        Debug.Print "[Demo01] ○ SINGLE成功！ files=" & GetInputFileNames(browserTab, "#singleInput")
    Else
        Debug.Print "[Demo01] × SINGLE失敗 files.length=" & GetInputFileCount(browserTab, "#singleInput")
        MsgBox "SINGLE FILE INJECTION に失敗しました。", vbCritical
    End If

    '--- 4. MULTI FILE INJECTION：2件のCollectionを直接注入 ---
    Debug.Print "[Demo01] #multiInput へ 2件のファイルを直接注入..."
    Dim multiFiles As New Collection
    multiFiles.Add file2
    multiFiles.Add file3

    If fc.SetFileInputFiles("#multiInput", multiFiles) And GetInputFileCount(browserTab, "#multiInput") = 2 Then
        Debug.Print "[Demo01] ○ MULTI成功！ files=" & GetInputFileNames(browserTab, "#multiInput")
        MsgBox "SINGLE / MULTI 両方の直接注入が成功しました！", vbInformation
    Else
        Debug.Print "[Demo01] × MULTI失敗 files.length=" & GetInputFileCount(browserTab, "#multiInput")
        MsgBox "MULTI FILE INJECTION に失敗しました。", vbCritical
    End If

    browserTab.InheritanceWebDriverBiDiMode.quit

End Sub



'***************************************************************************************************
'      ■■■ Demo 02：input.fileDialogOpened の横取りによる3種類の添付 ■■■
'***************************************************************************************************
'* 機能　　：`input.fileDialogOpened`イベントの横取りによる3パターンの添付を確認するデモです
'---------------------------------------------------------------------------------------------------
'* 3種類：
'   ① SINGLE FILE INJECTION（静的input・単一ファイル）
'   ② MULTI FILE INJECTION（静的input・複数ファイル）
'   ③ ON-DEMAND DYNAMIC INPUT（JSで動的生成されたinputへの横取り添付）
'* 確認ポイント：
'   - いずれも AddFilePath → click → AutoWaitFileDialogOpened の流れで添付が完了すること
'   - ③（DOM上に事前に存在しないinput、CSSセレクタで特定できない）でも、イベントさえ来れば
'     `sharedId`をイベント自身から取得できるため、添付が成立すること
'***************************************************************************************************
Sub Demo_FileChooser_02_3種類の添付()

    Dim file1 As String: file1 = EnsureSampleFile("intercept_single.txt", "横取り添付：単一ファイルのテストです。")
    Dim file2 As String: file2 = EnsureSampleFile("intercept_multi_1.txt", "横取り添付：複数ファイルのテスト 1")
    Dim file3 As String: file3 = EnsureSampleFile("intercept_multi_2.txt", "横取り添付：複数ファイルのテスト 2")
    Dim file4 As String: file4 = EnsureSampleFile("intercept_dynamic.txt", "横取り添付：動的生成inputへのテストです。")

    '--- 1. テストHTMLをブラウザで開く（OSダイアログ抑制のための capabilities 付き） ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browserTab As WebDriverBiDiContext
    Set browserTab = ShSetting01_StartBrowser.StartBiDiModeContext("file:///" & Replace(htmlPath, "\", "/"), sessionCapabilitiesRequest:=BuildFileDialogCapabilities())

    '--- 2. FileChooser拡張の準備（Init時点でinput.fileDialogOpenedの購読も行われる） ---
    Dim fc As New exBiDi_FileChooser
    fc.Init browserTab

    '--- ① SINGLE FILE INJECTION ---
    Debug.Print "[Demo02] ────── ① 静的input・単一ファイル ──────"
    fc.AddFilePath = file1
    browserTab.jsEval "document.getElementById('singleInput').click()", userActivation:=True, RunAsyncBiDi:=True

    If fc.AutoWaitFileDialogOpened(TimeOutSecond:=10) Then
        Debug.Print "[Demo02] ○ ①成功 files=" & GetInputFileNames(browserTab, "#singleInput")
    Else
        Debug.Print "[Demo02] × ①失敗"
        MsgBox "①（単一ファイル）の横取り添付に失敗しました。", vbCritical
        browserTab.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If

    '--- ② MULTI FILE INJECTION ---
    Debug.Print "[Demo02] ────── ② 静的input・複数ファイル ──────"
    fc.AddFilePath = file2
    fc.AddFilePath = file3
    Debug.Print "[Demo02]   登録ファイル数: " & fc.FilePathCount
    browserTab.jsEval "document.getElementById('multiInput').click()", userActivation:=True, RunAsyncBiDi:=True

    If fc.AutoWaitFileDialogOpened(TimeOutSecond:=10) Then
        Debug.Print "[Demo02] ○ ②成功 files=" & GetInputFileNames(browserTab, "#multiInput")
    Else
        Debug.Print "[Demo02] × ②失敗 files.length=" & GetInputFileCount(browserTab, "#multiInput")
        MsgBox "②（複数ファイル）の横取り添付に失敗しました。", vbCritical
        browserTab.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If

    '--- ③ ON-DEMAND DYNAMIC INPUT ---
    '    ※ DOM上に事前に存在しないinputだが、イベント自身がsharedIdを運んでくるためセレクタ不要
    Debug.Print "[Demo02] ────── ③ 動的生成input ──────"
    fc.AddFilePath = file4
    browserTab.jsEval "document.getElementById('customBtn').click()", userActivation:=True, RunAsyncBiDi:=True

    If fc.AutoWaitFileDialogOpened(TimeOutSecond:=10) Then
        Debug.Print "[Demo02] ○ ③成功（input.setFilesが正常応答）"
    Else
        Debug.Print "[Demo02] × ③失敗"
        MsgBox "③（動的生成input）の横取り添付に失敗しました。", vbCritical
        browserTab.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If

    fc.EnableEvents = False
    Debug.Print "[Demo02] 3種類全ての添付が成功しました！"
    MsgBox "3種類（単一 / 複数 / 動的生成）全ての横取り添付が成功しました！", vbInformation

    browserTab.InheritanceWebDriverBiDiMode.quit

End Sub



'***************************************************************************************************
'                  ■■■ Demo 03：キャンセルモードが機能すること ■■■
'***************************************************************************************************
'* 機能　　：`EnableEvents(cancel:=True)`で、添付を行わずスキップできることを確認します
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - キャンセルモード中は、ファイルを登録していても添付が行われないこと
'   - 対象inputの files.length が 0 のまま（＝標準DOM APIレベルで未添付）であること
'* 注意事項：
'   - このデモは`unhandledPromptBehavior`によるOSダイアログ抑制を前提としています。詳細はReadMe.mdを参照
'***************************************************************************************************
Sub Demo_FileChooser_03_キャンセル機能()

    Dim file1 As String: file1 = EnsureSampleFile("cancel_test.txt", "キャンセルテスト用ファイル（添付されないはず）。")

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browserTab As WebDriverBiDiContext
    Set browserTab = ShSetting01_StartBrowser.StartBiDiModeContext("file:///" & Replace(htmlPath, "\", "/"), sessionCapabilitiesRequest:=BuildFileDialogCapabilities())

    '--- 2. FileChooser拡張の準備 ---
    Dim fc As New exBiDi_FileChooser
    fc.Init browserTab

    '--- 3. キャンセルモードへ切り替え（ファイルは念のため登録しておく） ---
    fc.EnableEvents(cancel:=True) = True
    fc.AddFilePath = file1
    Debug.Print "[Demo03] キャンセルモードON。登録ファイル数=" & fc.FilePathCount & "（本来は使われないはず）"

    '--- 4. ダイアログをトリガー ---
    browserTab.jsEval "document.getElementById('singleInput').click()", userActivation:=True, RunAsyncBiDi:=True
    Debug.Print "[Demo03] singleInput をクリック（添付されないはず）..."

    '--- 5. イベント受信後、実際に添付されていないことを確認 ---
    If fc.AutoWaitFileDialogOpened(TimeOutSecond:=10) Then
        Debug.Print "[Demo03] ○ input.fileDialogOpened 受信（キャンセルモードのため添付スキップ済）"

        Dim fileCount As Variant
        fileCount = GetInputFileCount(browserTab, "#singleInput")
        If CLng(fileCount) = 0 Then
            Debug.Print "[Demo03] ○ files.length=0 を確認（未添付）"
            MsgBox "キャンセルモードが正常に働きました！", vbInformation
        Else
            Debug.Print "[Demo03] × files.length=" & fileCount & "（想定外：添付されてしまった）"
            MsgBox "キャンセルモードのはずですが、ファイルが添付されています。要確認。", vbCritical
        End If
    Else
        Debug.Print "[Demo03] × input.fileDialogOpened を検知できませんでした"
        MsgBox "ダイアログイベントを検知できませんでした。", vbCritical
    End If

    '--- 6. キャンセルモードを解除（登録ファイルも念のためクリア） ---
    fc.EnableEvents = False
    fc.ClearFilePaths

    browserTab.InheritanceWebDriverBiDiMode.quit

End Sub



'***************************************************************************************************
'          ■■■ Demo 04：添付忘れでも RetrySetFileInputFiles でリカバリー（3種類） ■■■
'***************************************************************************************************
'* 機能　　：ファイルを登録せずにダイアログを開いてしまった（添付忘れ）ケースからの復旧を、
'            Demo02と同じ3パターン（単一 / 複数 / 動的生成input）それぞれで確認します
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - 添付忘れの時点で添付は起きず、UnprocessedCount が増えること（クラス自身が持つ状態で判定）
'   - 後から AddFilePath → RetrySetFileInputFiles で添付が成立すること
'   - 各ラウンド後、UnprocessedCount が 0 に戻ること
'***************************************************************************************************
Sub Demo_FileChooser_04_添付忘れからのリトライ()

    Dim file1 As String: file1 = EnsureSampleFile("retry_single.txt", "添付忘れリトライテスト：単一ファイルです。")
    Dim file2 As String: file2 = EnsureSampleFile("retry_multi_1.txt", "添付忘れリトライテスト：複数ファイル 1")
    Dim file3 As String: file3 = EnsureSampleFile("retry_multi_2.txt", "添付忘れリトライテスト：複数ファイル 2")
    Dim file4 As String: file4 = EnsureSampleFile("retry_dynamic.txt", "添付忘れリトライテスト：動的生成inputです。")

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browserTab As WebDriverBiDiContext
    Set browserTab = ShSetting01_StartBrowser.StartBiDiModeContext("file:///" & Replace(htmlPath, "\", "/"), sessionCapabilitiesRequest:=BuildFileDialogCapabilities())

    '--- 2. FileChooser拡張の準備（ファイルは、まだ何も登録しない＝添付忘れの再現） ---
    Dim fc As New exBiDi_FileChooser
    fc.Init browserTab

    '--- ① 静的input・単一ファイル ---
    Debug.Print "[Demo04] ────── ① 静的input・単一ファイルの添付忘れ ──────"
    browserTab.jsEval "document.getElementById('singleInput').click()", userActivation:=True, RunAsyncBiDi:=True

    If Not WaitUntilUnprocessed(browserTab, fc, expectedCount:=1, TimeOutSecond:=10) Then
        Debug.Print "[Demo04]   × input.fileDialogOpened を検知できませんでした"
        MsgBox "①：ダイアログイベントを検知できませんでした。", vbCritical
        browserTab.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If
    Debug.Print "[Demo04]   ○ 添付忘れを再現。UnprocessedCount=" & fc.UnprocessedCount

    fc.AddFilePath = file1
    If fc.RetrySetFileInputFiles() And GetInputFileCount(browserTab, "#singleInput") = 1 Then
        Debug.Print "[Demo04] ○ ①成功！ files=" & GetInputFileNames(browserTab, "#singleInput") & " / UnprocessedCount=" & fc.UnprocessedCount
    Else
        Debug.Print "[Demo04] × ①失敗"
        MsgBox "①（単一ファイル）のリトライに失敗しました。", vbCritical
        browserTab.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If

    '--- ② 静的input・複数ファイル ---
    Debug.Print "[Demo04] ────── ② 静的input・複数ファイルの添付忘れ ──────"
    browserTab.jsEval "document.getElementById('multiInput').click()", userActivation:=True, RunAsyncBiDi:=True

    If Not WaitUntilUnprocessed(browserTab, fc, expectedCount:=1, TimeOutSecond:=10) Then
        Debug.Print "[Demo04]   × input.fileDialogOpened を検知できませんでした"
        MsgBox "②：ダイアログイベントを検知できませんでした。", vbCritical
        browserTab.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If
    Debug.Print "[Demo04]   ○ 添付忘れを再現。UnprocessedCount=" & fc.UnprocessedCount

    fc.AddFilePath = file2
    fc.AddFilePath = file3
    If fc.RetrySetFileInputFiles() And GetInputFileCount(browserTab, "#multiInput") = 2 Then
        Debug.Print "[Demo04] ○ ②成功！ files=" & GetInputFileNames(browserTab, "#multiInput") & " / UnprocessedCount=" & fc.UnprocessedCount
    Else
        Debug.Print "[Demo04] × ②失敗"
        MsgBox "②（複数ファイル）のリトライに失敗しました。", vbCritical
        browserTab.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If

    '--- ③ 動的生成input ---
    '    ※ CDPElementに相当する要素ラップが無く、かつ添付後にDOMから除去されるため
    '      成否は RetrySetFileInputFiles の戻り値と UnprocessedCount のみで判定する
    Debug.Print "[Demo04] ────── ③ 動的生成inputの添付忘れ ──────"
    browserTab.jsEval "document.getElementById('customBtn').click()", userActivation:=True, RunAsyncBiDi:=True

    If Not WaitUntilUnprocessed(browserTab, fc, expectedCount:=1, TimeOutSecond:=10) Then
        Debug.Print "[Demo04]   × input.fileDialogOpened を検知できませんでした"
        MsgBox "③：ダイアログイベントを検知できませんでした。", vbCritical
        browserTab.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If
    Debug.Print "[Demo04]   ○ 添付忘れを再現。UnprocessedCount=" & fc.UnprocessedCount

    fc.AddFilePath = file4
    If fc.RetrySetFileInputFiles() And fc.UnprocessedCount = 0 Then
        Debug.Print "[Demo04] ○ ③成功！（input.setFilesが正常応答） UnprocessedCount=" & fc.UnprocessedCount
    Else
        Debug.Print "[Demo04] × ③失敗"
        MsgBox "③（動的生成input）のリトライに失敗しました。", vbCritical
        browserTab.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If

    fc.EnableEvents = False
    Debug.Print "[Demo04] 3種類全ての添付忘れリトライが成功しました！"
    MsgBox "3種類（単一 / 複数 / 動的生成）全ての添付忘れリトライが成功しました！", vbInformation

    browserTab.InheritanceWebDriverBiDiMode.quit

End Sub



'***************************************************************************************************
'                           ■■■ ユーティリティ（共通利用） ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：`session.new`用の capabilities を構築します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`unhandledPromptBehavior.file`を`"ignore"`以外にしておくことで、
'            `Page.setInterceptFileChooserDialog`（実際のOSダイアログ抑制）が有効になります。
'            省略した場合の既定値は`"ignore"`＝抑制なしのため、`input.fileDialogOpened`の通知は
'            届いても、実際のOSダイアログが開こうとする可能性があります。
'* 注意事項：ここで指定した capabilities は、`StartBiDiModeContext`が新規にブラウザ/BiDiセッションを
'            起動した場合のみ適用されます（既存セッションへの`reattach`時は無効）。
'***************************************************************************************************
Private Function BuildFileDialogCapabilities() As Dictionary
    Dim promptBehavior As New Dictionary
    promptBehavior.Add "file", "dismiss"

    Dim alwaysMatch As New Dictionary
    alwaysMatch.Add "unhandledPromptBehavior", promptBehavior

    Dim capabilities As New Dictionary
    capabilities.Add "alwaysMatch", alwaysMatch

    Dim sessionCaps As New Dictionary
    sessionCaps.Add "capabilities", capabilities

    Set BuildFileDialogCapabilities = sessionCaps
End Function

'***************************************************************************************************
'* 機能　　：サンプルファイルが無ければ作成し、フルパスを返します
'***************************************************************************************************
Private Function EnsureSampleFile(FileName As String, Content As String) As String
    Dim fullPath As String
    fullPath = WORKSPACE_PATH & SAMPLE_DIR & "\" & FileName

    If Dir(fullPath) = "" Then
        CharConv.BytesToSaveFile CharConv.BytesFromString(Content & vbCrLf & "作成日時: " & Now()), WORKSPACE_PATH & SAMPLE_DIR, FileName
    End If

    EnsureSampleFile = fullPath
End Function

'***************************************************************************************************
'* 機能　　：指定したCSSセレクタの`<input type="file">`から、標準の`files.length`を取得します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ページ側の装飾UI（バッジ等）に依存せず、対象input自体のDOM状態だけで判定するための
'            ヘルパーです。
'***************************************************************************************************
Private Function GetInputFileCount(browserTab As WebDriverBiDiContext, cssSelector As String) As Long
    GetInputFileCount = browserTab.jsEval("document.querySelector('" & cssSelector & "').files.length")
End Function

'***************************************************************************************************
'* 機能　　：指定したCSSセレクタの`<input type="file">`から、添付済みファイル名一覧（カンマ区切り）を取得します
'***************************************************************************************************
Private Function GetInputFileNames(browserTab As WebDriverBiDiContext, cssSelector As String) As String
    GetInputFileNames = CStr(browserTab.jsEval("Array.from(document.querySelector('" & cssSelector & "').files).map(function(f){ return f.name }).join(', ')"))
End Function

'***************************************************************************************************
'* 機能　　：`fc.UnprocessedCount`が指定件数に達するまで待機します（添付忘れ検知用）
'---------------------------------------------------------------------------------------------------
'* 返り値　：True で検知、False でタイムアウト
'* 詳細説明：ページのDOMではなく、`exBiDi_FileChooser`クラス自身が持つ状態（保留リスト件数）だけを
'            見ているため、テストページのUIが変わっても影響を受けません。
'***************************************************************************************************
Private Function WaitUntilUnprocessed(browserTab As WebDriverBiDiContext, fc As exBiDi_FileChooser, expectedCount As Long, Optional TimeOutSecond As Double = 15) As Boolean
    Dim t As Double: t = Timer

    Do
        browserTab.InheritanceWebDriverBiDiMode.TakeEvents

        If fc.UnprocessedCount >= expectedCount Then
            WaitUntilUnprocessed = True
            Exit Function
        End If

        If (Timer - t) > TimeOutSecond Then Exit Function

        browserTab.InheritanceWebDriverBiDiMode.sleep 0.2
    Loop
End Function
