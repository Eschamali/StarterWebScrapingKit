Attribute VB_Name = "Demo_FileChooser"
'***************************************************************************************************
'       exCDP_FileChooser 拡張 / CDPElement.SetFileInputFiles - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：ファイル添付の2つのアプローチ（直接注入 / ダイアログ横取り）のサンプルコードです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\File Chooser\exCDP_FileChooser.cls
'* テストHTML：Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html
'---------------------------------------------------------------------------------------------------
'* 2つのアプローチの違い：
'   ①`CDPElement.SetFileInputFiles`（直接注入）
'      ・既に存在する`<input type="file">`のCDPElementを取得できていれば、これだけで完結
'      ・`Page.fileChooserOpened`の監視は不要（objectId経由で直接`DOM.setFileInputFiles`）
'      ・OSダイアログを一切開かせないので最速・最安定
'
'   ②`exCDP_FileChooser.cls`（ダイアログ横取り）
'      ・`.click()`等で実際にOSのファイル選択ダイアログが開こうとする瞬間を
'        `Page.fileChooserOpened`で検知し、代わりに`DOM.setFileInputFiles`で横取り添付する
'      ・input要素を事前に取得できない（動的生成・別ウィンドウ等）ケースで必要になる
'      ・キャンセルモード／添付忘れリカバリー（Retry）など、①には無い付加機能を持つ
'---------------------------------------------------------------------------------------------------
'* `exCDP_FileChooser.cls` の実際の公開API：
'   ・EnableEvents(Optional cancel As Boolean) = True/False  … 横取り監視のON/OFF（cancel:=Trueで全キャンセルモード）
'   ・AddFilePath = "path"                                   … 添付予定ファイルを1件ずつ登録（複数回でマルチ添付）
'   ・FilePathCount / UnprocessedCount
'   ・AutoWaitFileChooserOpened(TimeOutSecond)                … 次のPage.fileChooserOpenedが来るまで待機（Boolean）
'   ・RetrySetFileInputFiles()                                … 添付忘れで保留になった`backendNodeId`へ再添付（Boolean）
'   ・ClearFilePaths / ClearUnprocessed
'---------------------------------------------------------------------------------------------------
'* 検証方針について：
'   ・成功/失敗の判定は、このテストページ固有のバッジ表示（#badge1等の装飾UI）には依存させず、
'     標準の`HTMLInputElement.files`（対象inputのCDPElementに対する要素スコープjsEval）と、
'     各クラス自身が返す戻り値（AutoWaitFileChooserOpened / RetrySetFileInputFiles のBoolean、
'     UnprocessedCountなどのプロパティ）だけで判定しています。ページのUIが変わっても崩れません。
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'   ・「Page.fileChooserOpened」はブラウザが前面にある状態でないと発火しません
'   ・ブラウザのバージョンによっては、CDPElement.SetOptionUserGesture:=Trueで実行しないと発火しません
'   ・WORKSPACE_PATH をご自身の環境に合わせて設定してください
'***************************************************************************************************
Option Explicit



'ワークスペースパス
'※ StarterWebScrapingKit のルートフォルダを入力してください
Private Const WORKSPACE_PATH As String = ""

Private Const SAMPLE_DIR As String = "\Extensions\MainUnit\CDP\File Chooser"

Private CharConv As New CharacterCodeConversion



'***************************************************************************************************
'      ■■■ Demo 01：CDPElement.SetFileInputFiles による直接注入（SINGLE / MULTI） ■■■
'***************************************************************************************************
'* 機能　　：ダイアログ横取りを使わず、`<input type="file">`のCDPElementへ直接ファイルを注入するデモです
'---------------------------------------------------------------------------------------------------
'* テストページ：Test_FileChooser/index.html の「1?? SINGLE FILE INJECTION」「2?? MULTI FILE INJECTION」
'* 確認ポイント：
'   - `CDPElement.SetFileInputFiles`（1件のCollection）で singleInput への単一注入が成功すること
'   - `CDPElement.SetFileInputFiles`（2件以上のCollection）で multiInput への複数注入が成功すること
'   - `exCDP_FileChooser`を一切使わずに完結できること（`Page.fileChooserOpened`監視は不要）
'***************************************************************************************************
Sub Demo_FileChooser_01_直接注入_単一と複数()

    Dim file1 As String: file1 = EnsureSampleFile("inject_single.txt", "SINGLE FILE INJECTION のテストファイルです。")
    Dim file2 As String: file2 = EnsureSampleFile("inject_multi_1.txt", "MULTI FILE INJECTION テストファイル 1")
    Dim file3 As String: file3 = EnsureSampleFile("inject_multi_2.txt", "MULTI FILE INJECTION テストファイル 2")

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("file:///" & Replace(htmlPath, "\", "/"))
    browserTab.show

    '--- 2. SINGLE FILE INJECTION：1件のCollectionを直接注入 ---
    Debug.Print "[Demo01] singleInput へ 1件のファイルを直接注入..."
    Dim singleFiles As New Collection
    singleFiles.Add file1

    Dim singleInput As CDPElement
    Set singleInput = browserTab.getElementByID("singleInput")
    singleInput.SetFileInputFiles singleFiles

    If GetInputFileCount(singleInput) = 1 Then
        Debug.Print "[Demo01] ○ SINGLE成功！ files=" & GetInputFileNames(singleInput)
    Else
        Debug.Print "[Demo01] × SINGLE失敗 files.length=" & GetInputFileCount(singleInput)
        MsgBox "SINGLE FILE INJECTION に失敗しました。", vbCritical
    End If

    '--- 3. MULTI FILE INJECTION：2件のCollectionを直接注入 ---
    Debug.Print "[Demo01] multiInput へ 2件のファイルを直接注入..."
    Dim multiFiles As New Collection
    multiFiles.Add file2
    multiFiles.Add file3

    Dim multiInput As CDPElement
    Set multiInput = browserTab.getElementByID("multiInput")
    multiInput.SetFileInputFiles multiFiles

    If GetInputFileCount(multiInput) = 2 Then
        Debug.Print "[Demo01] ○ MULTI成功！ files=" & GetInputFileNames(multiInput)
        MsgBox "SINGLE / MULTI 両方の直接注入が成功しました！", vbInformation
    Else
        Debug.Print "[Demo01] × MULTI失敗 files.length=" & GetInputFileCount(multiInput)
        MsgBox "MULTI FILE INJECTION に失敗しました。", vbCritical
    End If

    browserTab.InheritanceCDPBrowser.quit

End Sub



'***************************************************************************************************
'      ■■■ Demo 02：exCDP_FileChooser による3種類の添付（全て成功すること） ■■■
'***************************************************************************************************
'* 機能　　：`Page.fileChooserOpened`の横取りによる3パターンの添付を確認するデモです
'---------------------------------------------------------------------------------------------------
'* 3種類：
'   ① SINGLE FILE INJECTION（静的input・単一ファイル）
'   ② MULTI FILE INJECTION（静的input・複数ファイル）
'   ③ ON-DEMAND DYNAMIC INPUT（JSで動的生成されたinputへの横取り添付）
'* 確認ポイント：
'   - いずれも AddFilePath → click → AutoWaitFileChooserOpened の流れで添付が完了すること
'   - ①②は対象inputのCDPElementを保持できているので、files.length/名前まで直接確認する
'   - ③（DOM上に事前に存在しないinput）は要素を保持できないため、AutoWaitFileChooserOpenedの
'     戻り値（＝CDP `DOM.setFileInputFiles`自体の成否）そのものを成功判定として扱う
'* 注意事項：
'   - `Page.fileChooserOpened`の発火には`userGesture`をOnにて、人間による操作の礼儀を尽くす必要があります
'   - `SimpleClick`は非同期実行（SetOptionRunAsyncCDP:=True）と組み合わせているため、
'     `AutoWaitFileChooserOpened`を呼ぶ前に、既にイベント処理（添付）が終わっているケースがあります。
'     その場合は`FilePathCount=0`（登録した分がすでに消費された）を根拠に成功とみなします。
'***************************************************************************************************
Sub Demo_FileChooser_02_3種類の添付()

    Dim file1 As String: file1 = EnsureSampleFile("intercept_single.txt", "横取り添付：単一ファイルのテストです。")
    Dim file2 As String: file2 = EnsureSampleFile("intercept_multi_1.txt", "横取り添付：複数ファイルのテスト 1")
    Dim file3 As String: file3 = EnsureSampleFile("intercept_multi_2.txt", "横取り添付：複数ファイルのテスト 2")
    Dim file4 As String: file4 = EnsureSampleFile("intercept_dynamic.txt", "横取り添付：動的生成inputへのテストです。")

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("file:///" & Replace(htmlPath, "\", "/"))
    browserTab.show

    '--- 2. FileChooser拡張の準備（Init時点で横取り監視は自動的に有効化される） ---
    Dim fc As New exCDP_FileChooser
    fc.Init browserTab

    '--- ① SINGLE FILE INJECTION ---
    Debug.Print "[Demo02] ────── ① 静的input・単一ファイル ──────"
    Dim singleInput As CDPElement: Set singleInput = browserTab.getElementByID("singleInput")
    fc.AddFilePath = file1
    singleInput.SetOptionUserGesture = True     '人間による操作の礼儀を尽くす
    singleInput.SetOptionRunAsyncCDP = True     '`True`にすると、最初のif分岐を想定。`False`だと2番目の`ElseIf`に分岐する想定
    singleInput.SimpleClick

    If fc.AutoWaitFileChooserOpened(TimeOutSecond:=10) Then
        Debug.Print "[Demo02] ○ ①成功 files=" & GetInputFileNames(singleInput)

    ElseIf fc.FilePathCount = 0 Then
        Debug.Print "[Demo02] ○ ①成功(前述の`.SimpleClick`のついでに処理してくれたようだ) files=" & GetInputFileNames(singleInput)

    Else
        Debug.Print "[Demo02] × ①失敗"
        MsgBox "①（単一ファイル）の横取り添付に失敗しました。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If

    '--- ② MULTI FILE INJECTION ---
    Debug.Print "[Demo02] ────── ② 静的input・複数ファイル ──────"
    Dim multiInput As CDPElement: Set multiInput = browserTab.getElementByID("multiInput")
    fc.AddFilePath = file2
    fc.AddFilePath = file3
    Debug.Print "[Demo02]   登録ファイル数: " & fc.FilePathCount
    multiInput.SimpleClick

    If fc.AutoWaitFileChooserOpened(TimeOutSecond:=10) Then
        Debug.Print "[Demo02] ○ ②成功 files=" & GetInputFileNames(multiInput)

    ElseIf fc.FilePathCount = 0 Then
        Debug.Print "[Demo02] ○ ②成功 (前述の`.SimpleClick`のついでに処理してくれたようだ) files=" & GetInputFileNames(multiInput)

    Else
        Debug.Print "[Demo02] × ②失敗 files.length=" & GetInputFileCount(multiInput)
        MsgBox "②（複数ファイル）の横取り添付に失敗しました。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If

    '--- ③ ON-DEMAND DYNAMIC INPUT ---
    '    ※ DOM上に事前に存在しないinputなのでCDPElementを保持できない。
    '      成否は AutoWaitFileChooserOpened（＝CDPコマンド自体の成否）のみで判定する。
    Debug.Print "[Demo02] ────── ③ 動的生成input ──────"
    fc.AddFilePath = file4
    With browserTab.getElementByID("customBtn")
'        .SetOptionUserGesture = True    'ここでも念のため入れておくとクールタイムが空くよう
        .SetOptionRunAsyncCDP = True    '`True`にすると、最初のif分岐を想定。`False`だと2番目の`ElseIf`に分岐する想定

        .SimpleClick
    End With

    If fc.AutoWaitFileChooserOpened(TimeOutSecond:=10) Then
        Debug.Print "[Demo02] ○ ③成功（DOM.setFileInputFilesが正常応答）"

    ElseIf fc.FilePathCount = 0 Then
        Debug.Print "[Demo02] ○ ③成功 (前述の`.SimpleClick`のついでに処理してくれたようだ。DOM.setFileInputFilesが正常応答)"

    Else
        Debug.Print "[Demo02] × ③失敗"
        MsgBox "③（動的生成input）の横取り添付に失敗しました。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If

    fc.EnableEvents = False
    Debug.Print "[Demo02] 3種類全ての添付が成功しました！"
    MsgBox "3種類（単一 / 複数 / 動的生成）全ての横取り添付が成功しました！", vbInformation

    browserTab.InheritanceCDPBrowser.quit

End Sub



'***************************************************************************************************
'                  ■■■ Demo 03：キャンセルモードが機能すること ■■■
'***************************************************************************************************
'* 機能　　：`EnableEvents(cancel:=True)`で、ファイル選択ダイアログを強制キャンセルできることを確認します
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - キャンセルモード中は、ファイルを登録していてもダイアログが自動キャンセルされ、添付は起きないこと
'   - 対象inputの files.length が 0 のまま（＝標準DOM APIレベルで未添付）であること
'* 注意事項：
'   - キャンセル操作の場合は、「非同期でクリック」しないとうまく判定できません。まぁ、おまけ程度として
'***************************************************************************************************
Sub Demo_FileChooser_03_キャンセル機能()

    Dim file1 As String: file1 = EnsureSampleFile("cancel_test.txt", "キャンセルテスト用ファイル（添付されないはず）。")

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("file:///" & Replace(htmlPath, "\", "/"))
    browserTab.show

    '--- 2. FileChooser拡張の準備 ---
    Dim fc As New exCDP_FileChooser
    fc.Init browserTab

    '--- 3. キャンセルモードへ切り替え（ファイルは念のため登録しておく） ---
    fc.EnableEvents(cancel:=True) = True
    fc.AddFilePath = file1
    Debug.Print "[Demo03] キャンセルモードON。登録ファイル数=" & fc.FilePathCount & "（本来は使われないはず）"

    '--- 4. ダイアログを非同期でトリガー ---
    Dim singleInput As CDPElement: Set singleInput = browserTab.getElementByID("singleInput")
    singleInput.SetOptionUserGesture = True
    singleInput.SetOptionRunAsyncCDP = True
    singleInput.SimpleClick
    Debug.Print "[Demo03] singleInput をクリック（キャンセルされるはず）..."

    '--- 5. イベント側で、実際に添付されていないことを確認 ---
    If fc.AutoWaitFileChooserOpened(TimeOutSecond:=10) Then
        If fc.FilePathCount Then
            Debug.Print "[Demo03] ○ キャンセル成功！"
            MsgBox "キャンセル機能が正常に働きました！", vbInformation
        Else
            Debug.Print "[Demo03] △ キャンセルは成功したけど、添付ファイルリストも消えてるようです"
            MsgBox "キャンセル機能が正常に働きましたが、添付ファイルリストも消えてるようです", vbExclamation
        End If
    Else
        Debug.Print "[Demo03] × files.length=" & GetInputFileCount(singleInput) & "（キャンセルイベントがすでに応答済み？）"
        MsgBox "キャンセル機能が働きませんでした。", vbCritical
    End If

    '--- 7. キャンセルモードを解除（登録ファイルも念のためクリア） ---
    fc.EnableEvents = False
    fc.ClearFilePaths

    browserTab.InheritanceCDPBrowser.quit

End Sub



'***************************************************************************************************
'          ■■■ Demo 04：添付忘れでも RetrySetFileInputFiles でリカバリー（3種類） ■■■
'***************************************************************************************************
'* 機能　　：ファイルを登録せずにダイアログを開いてしまった（添付忘れ）ケースからの復旧を、
'            Demo02と同じ3パターン（単一 / 複数 / 動的生成input）それぞれで確認します
'---------------------------------------------------------------------------------------------------
'* 各パターン共通のフロー（1件ずつ、貯めずに即リトライ）：
'   ① ファイル未登録の状態で input をクリック（Page.fileChooserOpened が来るが、
'      FilePathCount=0 のため添付されず、UnprocessedSetFileList に保留される）
'   ② 保留を確認（UnprocessedCount > 0）
'   ③ 忘れていたファイルを AddFilePath で登録
'   ④ RetrySetFileInputFiles で保留分に遡って添付
'* 確認ポイント：
'   - ①の時点で添付は起きず、UnprocessedCount が増えること（クラス自身が持つ状態で判定）
'   - ③④の後、RetrySetFileInputFilesがTrueを返すこと
'   - ①②（静的input）は対象inputのfiles.lengthまで直接確認する
'   - ③（動的生成input）はCDPElementを保持できないため、クラス自身の戻り値のみで判定する
'   - 各ラウンド後、UnprocessedCount が 0 に戻ること（次のラウンドに影響を残さない）
'---------------------------------------------------------------------------------------------------
'* 複数件を先に貯めてから一括処理したいケースは Demo 05 を参照してください
'***************************************************************************************************
Sub Demo_FileChooser_04_添付忘れからのリトライ()

    Dim file1 As String: file1 = EnsureSampleFile("retry_single.txt", "添付忘れリトライテスト：単一ファイルです。")
    Dim file2 As String: file2 = EnsureSampleFile("retry_multi_1.txt", "添付忘れリトライテスト：複数ファイル 1")
    Dim file3 As String: file3 = EnsureSampleFile("retry_multi_2.txt", "添付忘れリトライテスト：複数ファイル 2")
    Dim file4 As String: file4 = EnsureSampleFile("retry_dynamic.txt", "添付忘れリトライテスト：動的生成inputです。")

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("file:///" & Replace(htmlPath, "\", "/"))
    browserTab.show

    '--- 2. FileChooser拡張の準備（ファイルは、まだ何も登録しない＝添付忘れの再現） ---
    Dim fc As New exCDP_FileChooser
    fc.Init browserTab

    '--- ① 静的input・単一ファイル ---
    Debug.Print "[Demo04] ────── ① 静的input・単一ファイルの添付忘れ ──────"
    Dim singleInput As CDPElement: Set singleInput = browserTab.getElementByID("singleInput")

    singleInput.SetOptionUserGesture = True
    singleInput.click
    Debug.Print "[Demo04]   singleInput をクリック（ファイル未登録のまま）..."

    If Not WaitUntilUnprocessed(browserTab, fc, expectedCount:=1, TimeOutSecond:=10) Then
        Debug.Print "[Demo04]   × Page.fileChooserOpened を検知できませんでした"
        MsgBox "①：ダイアログイベントを検知できませんでした。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If
    Debug.Print "[Demo04]   ○ 添付忘れを再現。UnprocessedCount=" & fc.UnprocessedCount & "（files.length=" & GetInputFileCount(singleInput) & "のはず）"

    fc.AddFilePath = file1
    Debug.Print "[Demo04]   今になってファイルを登録: " & file1

    If fc.RetrySetFileInputFiles() And GetInputFileCount(singleInput) = 1 Then
        Debug.Print "[Demo04] ○ ①成功！ files=" & GetInputFileNames(singleInput) & " / UnprocessedCount=" & fc.UnprocessedCount
    Else
        Debug.Print "[Demo04] × ①失敗 files.length=" & GetInputFileCount(singleInput)
        MsgBox "①（単一ファイル）のリトライに失敗しました。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If

    '--- ② 静的input・複数ファイル ---
    Debug.Print "[Demo04] ────── ② 静的input・複数ファイルの添付忘れ ──────"
    Dim multiInput As CDPElement: Set multiInput = browserTab.getElementByID("multiInput")

    multiInput.SetOptionUserGesture = True
    multiInput.click
    Debug.Print "[Demo04]   multiInput をクリック（ファイル未登録のまま）..."

    If Not WaitUntilUnprocessed(browserTab, fc, expectedCount:=1, TimeOutSecond:=10) Then
        Debug.Print "[Demo04]   × Page.fileChooserOpened を検知できませんでした"
        MsgBox "②：ダイアログイベントを検知できませんでした。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If
    Debug.Print "[Demo04]   ○ 添付忘れを再現。UnprocessedCount=" & fc.UnprocessedCount & "（files.length=" & GetInputFileCount(multiInput) & "のはず）"

    fc.AddFilePath = file2
    fc.AddFilePath = file3
    Debug.Print "[Demo04]   今になってファイルを登録: 2件（登録数=" & fc.FilePathCount & "）"

    If fc.RetrySetFileInputFiles() And GetInputFileCount(multiInput) = 2 Then
        Debug.Print "[Demo04] ○ ②成功！ files=" & GetInputFileNames(multiInput) & " / UnprocessedCount=" & fc.UnprocessedCount
    Else
        Debug.Print "[Demo04] × ②失敗 files.length=" & GetInputFileCount(multiInput)
        MsgBox "②（複数ファイル）のリトライに失敗しました。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If

    '--- ③ 動的生成input ---
    '    ※ DOM上に事前に存在しないinputなのでCDPElementを保持できない。
    '      成否は RetrySetFileInputFiles（＝CDPコマンド自体の成否）とUnprocessedCountのみで判定する。
    Debug.Print "[Demo04] ────── ③ 動的生成inputの添付忘れ ──────"
    Dim customBtn As CDPElement: Set customBtn = browserTab.getElementByID("customBtn")

    customBtn.SetOptionUserGesture = True
    customBtn.click   'JS側で<input>を動的生成してclick()する
    Debug.Print "[Demo04]   customBtn をクリック（ファイル未登録のまま）..."

    If Not WaitUntilUnprocessed(browserTab, fc, expectedCount:=1, TimeOutSecond:=10) Then
        Debug.Print "[Demo04]   × Page.fileChooserOpened を検知できませんでした"
        MsgBox "③：ダイアログイベントを検知できませんでした。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If
    Debug.Print "[Demo04]   ○ 添付忘れを再現。UnprocessedCount=" & fc.UnprocessedCount

    fc.AddFilePath = file4
    Debug.Print "[Demo04]   今になってファイルを登録: " & file4

    If fc.RetrySetFileInputFiles() And fc.UnprocessedCount = 0 Then
        Debug.Print "[Demo04] ○ ③成功！（DOM.setFileInputFilesが正常応答） UnprocessedCount=" & fc.UnprocessedCount
    Else
        Debug.Print "[Demo04] × ③失敗"
        MsgBox "③（動的生成input）のリトライに失敗しました。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If

    fc.EnableEvents = False
    Debug.Print "[Demo04] 3種類全ての添付忘れリトライが成功しました！"
    MsgBox "3種類（単一 / 複数 / 動的生成）全ての添付忘れリトライが成功しました！", vbInformation

    browserTab.InheritanceCDPBrowser.quit

End Sub



'***************************************************************************************************
'      ■■■ Demo 05：添付漏れ3連続 → UnprocessedSetFileListに3件貯める → 一括で順次処理 ■■■
'***************************************************************************************************
'* 機能　　：Demo04とは逆に「都度リトライ」せず、まず3連続でファイル未登録のままSimpleClickを重ねて
'            保留を3件貯めてから、まとめて順次リトライするデモです
'---------------------------------------------------------------------------------------------------
'* フロー：
'   ① singleInput / multiInput / customBtn を、ファイル未登録のまま3連続でSimpleClick
'      （3つとも`Page.fileChooserOpened`が来るが、その都度`FilePathCount=0`のため添付されず、
'        `UnprocessedSetFileList`に backendNodeId が3件積み上がるだけになる）
'   ② UnprocessedCount が 3 になるまで待つ
'   ③ 3件を、貯まった順（FIFO）に1件ずつ AddFilePath → RetrySetFileInputFiles で処理
'* 確認ポイント：
'   - 3連続クリックしても、その場では一切添付が起きないこと（全て保留に回る）
'   - UnprocessedCount が正しく 3 まで積み上がること
'   - 3回のRetrySetFileInputFilesで、保留が古い順に1件ずつ正しく消費されること
'     （① singleInput → ② multiInput → ③ customBtn の順でクリックしたので、この順で消費される想定）
'   - 3回とも処理した後、UnprocessedCount が 0 に戻ること
'* 注意事項：
'   - 3連続クリックは`SetOptionRunAsyncCDP:=True`の`SimpleClick`で行う（結果を待たずに次へ進むため）
'***************************************************************************************************
Sub Demo_FileChooser_05_添付漏れ3連続からの一括リトライ()

    Dim fileSingle As String: fileSingle = EnsureSampleFile("batch_single.txt", "3連続添付漏れテスト：単一ファイルです。")
    Dim fileMulti1 As String: fileMulti1 = EnsureSampleFile("batch_multi_1.txt", "3連続添付漏れテスト：複数ファイル 1")
    Dim fileMulti2 As String: fileMulti2 = EnsureSampleFile("batch_multi_2.txt", "3連続添付漏れテスト：複数ファイル 2")
    Dim fileDynamic As String: fileDynamic = EnsureSampleFile("batch_dynamic.txt", "3連続添付漏れテスト：動的生成inputです。")

    '--- 1. テストHTMLをブラウザで開く ---
    Dim htmlPath As String
    htmlPath = WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_FileChooser\index.html"
    Dim browserTab As CDPContext
    Set browserTab = ShSetting01_StartBrowser.StartCDPModeContext("file:///" & Replace(htmlPath, "\", "/"))
    browserTab.show

    '--- 2. FileChooser拡張の準備（ファイルは、まだ何も登録しない＝添付漏れの再現） ---
    Dim fc As New exCDP_FileChooser
    fc.Init browserTab

    Dim singleInput As CDPElement: Set singleInput = browserTab.getElementByID("singleInput")
    Dim multiInput As CDPElement: Set multiInput = browserTab.getElementByID("multiInput")
    Dim customBtn As CDPElement: Set customBtn = browserTab.getElementByID("customBtn")

    '--- 3. ファイル未登録のまま、3連続でSimpleClick（結果は待たない） ---
    Debug.Print "[Demo05] ────── 3連続で添付漏れを発生させる ──────"

    singleInput.SetOptionUserGesture = True
    singleInput.SetOptionRunAsyncCDP = True
    singleInput.SimpleClick
    Debug.Print "[Demo05]   1回目: singleInput をクリック（未登録）"

    multiInput.SetOptionUserGesture = True
    multiInput.SetOptionRunAsyncCDP = True
    multiInput.SimpleClick
    Debug.Print "[Demo05]   2回目: multiInput をクリック（未登録）"

    customBtn.SetOptionUserGesture = True
    customBtn.SetOptionRunAsyncCDP = True
    customBtn.SimpleClick   'JS側で<input>を動的生成してclick()する
    Debug.Print "[Demo05]   3回目: customBtn をクリック（未登録）"

    '--- 4. UnprocessedCountが3件貯まるまで待つ ---
    If Not WaitUntilUnprocessed(browserTab, fc, expectedCount:=3, TimeOutSecond:=15) Then
        Debug.Print "[Demo05] × 3件貯まりませんでした。UnprocessedCount=" & fc.UnprocessedCount
        MsgBox "3連続クリックしても、保留が3件になりませんでした。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If
    Debug.Print "[Demo05] ○ 3件とも保留に回りました！ UnprocessedCount=" & fc.UnprocessedCount _
              & "（この時点でsingleInput.files.length=" & GetInputFileCount(singleInput) _
              & " / multiInput.files.length=" & GetInputFileCount(multiInput) & "）"

    '--- 5. 貯まった3件を、古い順に1件ずつ処理する ---
    Debug.Print "[Demo05] ────── 保留3件を順次処理 ──────"

    '① 1件目（クリック順どおりなら singleInput のはず）
    fc.AddFilePath = fileSingle
    If fc.RetrySetFileInputFiles() Then
        Debug.Print "[Demo05] ○ 1件目リトライ成功 残りUnprocessedCount=" & fc.UnprocessedCount & " singleInput.files=" & GetInputFileNames(singleInput)
    Else
        Debug.Print "[Demo05] × 1件目リトライ失敗"
        MsgBox "1件目のリトライに失敗しました。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If

    '② 2件目（クリック順どおりなら multiInput のはず。複数ファイルを一括登録してから処理）
    fc.AddFilePath = fileMulti1
    fc.AddFilePath = fileMulti2
    If fc.RetrySetFileInputFiles() Then
        Debug.Print "[Demo05] ○ 2件目リトライ成功 残りUnprocessedCount=" & fc.UnprocessedCount & " multiInput.files=" & GetInputFileNames(multiInput)
    Else
        Debug.Print "[Demo05] × 2件目リトライ失敗"
        MsgBox "2件目のリトライに失敗しました。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If

    '③ 3件目（クリック順どおりなら customBtn の動的input のはず。CDPElementを保持できないため戻り値のみで判定）
    fc.AddFilePath = fileDynamic
    If fc.RetrySetFileInputFiles() Then
        Debug.Print "[Demo05] ○ 3件目リトライ成功 残りUnprocessedCount=" & fc.UnprocessedCount
    Else
        Debug.Print "[Demo05] × 3件目リトライ失敗"
        MsgBox "3件目のリトライに失敗しました。", vbCritical
        browserTab.InheritanceCDPBrowser.quit
        Exit Sub
    End If

    '--- 6. 全て処理し終わって、保留が0件に戻ったことを確認 ---
    If fc.UnprocessedCount = 0 Then
        Debug.Print "[Demo05] ○ 3件全て処理完了！ UnprocessedCount=" & fc.UnprocessedCount
        Debug.Print "[Demo05]   最終結果: singleInput.files=" & GetInputFileNames(singleInput) & " / multiInput.files=" & GetInputFileNames(multiInput)
        MsgBox "3連続の添付漏れを、まとめて順次リトライで全て回収できました！", vbInformation
    Else
        Debug.Print "[Demo05] × 処理し残しがあります。UnprocessedCount=" & fc.UnprocessedCount
        MsgBox "処理し残しがあります。UnprocessedCount=" & fc.UnprocessedCount, vbCritical
    End If

    fc.EnableEvents = False
    browserTab.InheritanceCDPBrowser.quit

End Sub



'***************************************************************************************************
'                           ■■■ ユーティリティ（共通利用） ■■■
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
'* 機能　　：指定した`<input type="file">`のCDPElementから、標準の`files.length`を取得します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ページ側の装飾UI（バッジ等）に依存せず、対象input自体のDOM状態だけで判定するための
'            ヘルパーです。要素スコープの`jsEval`（`this`=input自身）で直接読みます。
'***************************************************************************************************
Private Function GetInputFileCount(el As CDPElement) As Long
    GetInputFileCount = el.jsEval("function(){ return this.files.length }")
End Function

'***************************************************************************************************
'* 機能　　：指定した`<input type="file">`のCDPElementから、添付済みファイル名一覧（カンマ区切り）を取得します
'***************************************************************************************************
Private Function GetInputFileNames(el As CDPElement) As String
    GetInputFileNames = CStr(el.jsEval("function(){ return Array.from(this.files).map(function(f){ return f.name }).join(', ') }"))
End Function

'***************************************************************************************************
'* 機能　　：`fc.UnprocessedCount`が指定件数に達するまで待機します（添付忘れ検知用）
'---------------------------------------------------------------------------------------------------
'* 返り値　：True で検知、False でタイムアウト
'* 詳細説明：ページのDOMではなく、`exCDP_FileChooser`クラス自身が持つ状態（保留リスト件数）だけを
'            見ているため、テストページのUIが変わっても影響を受けません。
'***************************************************************************************************
Private Function WaitUntilUnprocessed(browserTab As CDPContext, fc As exCDP_FileChooser, expectedCount As Long, Optional TimeOutSecond As Double = 15) As Boolean
    Dim t As Double: t = Timer

    Do
        browserTab.InheritanceCDPBrowser.TakeEvents

        If fc.UnprocessedCount >= expectedCount Then
            WaitUntilUnprocessed = True
            Exit Function
        End If

        If (Timer - t) > TimeOutSecond Then Exit Function

        browserTab.InheritanceCDPBrowser.sleep 0.2
    Loop
End Function
