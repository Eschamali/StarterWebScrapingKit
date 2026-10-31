Attribute VB_Name = "Test_BiDiMode_GetTab"
'==============================================================================
' WebDriverBiDiMode.getTab 動作確認（幅優先探索と maxDepth）
' ・HTML: ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPiframe\NestedIframes.html（?mode=url で使用）
'         ※ CDP 側の CDPiframe テストと同じページを使います
' ・実行前に WebDriverBiDi が起動可、WORKSPACE_PATH をルートに設定してから RunAll_BiDiMode_GetTabTests を実行
'------------------------------------------------------------------------------
' テストページの構成（?mode=url&layout=3,0,1）
'   ・親ページ直下の iframe（A, B, C）が 1～3 本、それぞれの下に 0～3 段の iframe（A → A1 → A2 → A3）
'   ・mode=url では、各 iframe を固有のURL「data:text/html;frame=A1,...」で作る（親ページとは別オリジン）
'   ・ボタンを押すと、top.postMessage で親ページの window.__clicked に自分の名前を積む
'
' 確認すること
'   1. getTab() 条件なしで、最初のタブが返る
'   2. ページのURLで探すと、トップレベルのタブ自身が返る（maxDepth 省略 = タブのみ）
'   3. 各 iframe について
'        ・maxDepth=-1（無制限）で固有のURLを探し、返ったコンテキストでボタンを押す → 親に自分の名前が届くか
'        ・maxDepth = 深さ なら見つかる / 深さ-1 なら見つからない
'   4. 存在しないURLでは、エラーにならず Nothing が返る
'   5. doRetrySecond=1 で、約1秒待ってから Nothing が返る
'==============================================================================
Option Explicit

'ワークスペースパス（StarterWebScrapingKit ルート）
Private Const WORKSPACE_PATH As String = ""

Private Const TEST_HTML As String = "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPiframe\NestedIframes.html"

'試す構成（親ページ直下の iframe ごとに、その下へ何段ぶら下げるか）
Private Const TEST_CASES As String = "3,3,3|3,0,0|0|1,2"

'存在しないURL（見つからない場合の確認用）
Private Const UNKNOWN_URL As String = "https://never-exists.example/"

Private Function EOk() As String
    EOk = WorksheetFunction.Unichar(9989)
End Function

'==============================================================================
' エントリ
'==============================================================================
Public Sub RunAll_BiDiMode_GetTabTests()
    Debug.Print RunBiDiModeGetTabTests(WORKSPACE_PATH)
End Sub

'------------------------------------------------------------------------------
' 全項目を実行し、結果の一覧を返します（各行 PASS / FAIL、最終行が集計）
'   workspacePath  StarterWebScrapingKit のルート
'------------------------------------------------------------------------------
Public Function RunBiDiModeGetTabTests(workspacePath As String) As String
    If LenB(workspacePath) = 0 Then Err.Raise 5, "Test_BiDiMode_GetTab", "WORKSPACE_PATH を設定してください。"

    Dim pageUrl As String
    pageUrl = "file:///" & Replace(workspacePath & TEST_HTML, "\", "/")

    Dim br As WebDriverBiDiMode
    Set br = ShSetting01_StartBrowser.StartBiDiMode("about:blank")

    Dim lines As New Collection

    '1. 条件なしで、最初のタブ
    Dim topTab As WebDriverBiDiContext
    Set topTab = br.getTab()
    lines.Add IIf(topTab Is Nothing, "FAIL", "PASS") & " | getTab() 条件なしで最初のタブが返る"

    '2-3. 構成ごとの確認
    If Not topTab Is Nothing Then
        Dim caseName As Variant
        For Each caseName In Split(TEST_CASES, "|")
            lines.Add RunOneCase(br, topTab, pageUrl, CStr(caseName))
        Next
    End If

    '4. 存在しないURL → Nothing
    Dim missTab As WebDriverBiDiContext
    Set missTab = br.getTab(UNKNOWN_URL, -1)
    lines.Add IIf(missTab Is Nothing, "PASS", "FAIL") & " | 存在しないURLでは Nothing が返る"

    '5. doRetrySecond=1 → 約1秒待ってから Nothing
    Dim t0 As Double, elapsedMs As Double
    t0 = CDPHelpers.TimerCounter
    Set missTab = br.getTab(UNKNOWN_URL, -1, , 1)
    elapsedMs = CDPHelpers.TimerCounter - t0
    lines.Add IIf(missTab Is Nothing And elapsedMs >= 900, "PASS", "FAIL") & " | doRetrySecond=1 で約1秒待ってから Nothing が返る (elapsedMs=" & Format$(elapsedMs, "0") & ")"

    br.CloseBrowser

    '集計
    Dim report As String, line As Variant, passCount As Long, failCount As Long
    report = String(70, "=") & vbCrLf & "  WebDriverBiDiMode.getTab テスト 開始" & vbCrLf & String(70, "=") & vbCrLf
    For Each line In lines
        If Left$(line, 4) = "PASS" Then
            passCount = passCount + 1
            report = report & "  " & EOk() & " " & line & vbCrLf
        Else
            failCount = failCount + 1
            report = report & "  " & line & vbCrLf
        End If
    Next
    report = report & String(70, "=") & vbCrLf & "  テスト結果: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)
    RunBiDiModeGetTabTests = report
End Function

'==============================================================================
' 1構成分のテスト
'==============================================================================
Private Function RunOneCase(br As WebDriverBiDiMode, topTab As WebDriverBiDiContext, pageUrl As String, caseName As String) As String
    '1. ページを開いて、入れ子の iframe まで全部読み込まれるのを待つ
    topTab.navigate pageUrl & "?mode=url&layout=" & caseName
    Dim t0 As Double: t0 = CDPHelpers.TimerCounter
    Do Until topTab.jsEval("window.__ready===true") = True
        If CDPHelpers.TimerCounter - t0 > 10000 Then RunOneCase = "FAIL | " & caseName & " | page not ready": Exit Function
        CDPHelpers.Sleep 0.1
    Loop

    Dim layoutText As String, ng As String
    layoutText = topTab.jsEval("window.__layout.map(function(x){return x.name+':'+x.depth}).join('|')")

    '2. ページのURLで探すと、トップレベルのタブ自身（maxDepth 省略 = タブのみ）
    Dim found As WebDriverBiDiContext
    Set found = br.getTab("NestedIframes.html")
    If found Is Nothing Then
        ng = ng & " topNotFound"
    ElseIf found.context <> topTab.context Then
        ng = ng & " topWrongContext"
    End If

    '3. iframe ごとに確認
    Dim entry As Variant, frameName As String, frameDepth As Long, marker As String, lastClicked As String
    Dim t1 As Double
    For Each entry In Split(layoutText, "|")
        frameName = Split(entry, ":")(0)
        frameDepth = CLng(Split(entry, ":")(1))
        marker = "data:text/html;frame=" & frameName & ","

        '3-1. 深さ無制限で探し、そのコンテキストでボタンを押す → 親に自分の名前が届くか
        Set found = br.getTab(marker, -1)
        If found Is Nothing Then
            ng = ng & " notFound(" & frameName & ")"
        Else
            found.jsEval "document.getElementById('btn').click()"
            t1 = CDPHelpers.TimerCounter
            Do
                lastClicked = topTab.jsEval("window.__clicked.length?window.__clicked[window.__clicked.length-1]:''")
                If lastClicked = frameName Or CDPHelpers.TimerCounter - t1 > 2000 Then Exit Do
                CDPHelpers.Sleep 0.05
            Loop
            If lastClicked <> frameName Then ng = ng & " wrongContext(" & frameName & "->" & lastClicked & ")"
        End If

        '3-2. maxDepth = 深さ なら見つかる / 深さ-1 なら見つからない
        If br.getTab(marker, frameDepth) Is Nothing Then ng = ng & " missAtOwnLevel(" & frameName & ")"
        If Not br.getTab(marker, frameDepth - 1) Is Nothing Then ng = ng & " foundTooShallow(" & frameName & ")"
    Next

    '4. 押されたボタン数 = iframe 数 か
    Dim frameCount As Long, clickedCount As Long
    frameCount = UBound(Split(layoutText, "|")) + 1
    clickedCount = topTab.jsEval("window.__clicked.length")
    If clickedCount <> frameCount Then ng = ng & " clicked=" & clickedCount & "/" & frameCount

    RunOneCase = IIf(LenB(ng) = 0, "PASS", "FAIL") & " | " & caseName & " | " & frameCount & " iframes | buttons clicked=" & clickedCount & "/" & frameCount & ng
End Function
