Attribute VB_Name = "Test_CDPiframe_Nest"
'==============================================================================
' CDPiframe.cls 動作確認（入れ子 iframe の幅優先探索と NestLevel）
' ・HTML: ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPiframe\NestedIframes.html
' ・実行前に CDP が起動可、WORKSPACE_PATH をルートに設定してから RunAll_CDPiframe_NestTests を実行
'------------------------------------------------------------------------------
' テストページの構成（?layout=3,0,1 で指定。省略時はランダム）
'   ・親ページ直下の iframe（A, B, C）が 1～3 本
'   ・それぞれの下に、さらに 0～3 段の iframe がぶら下がる（A → A1 → A2 → A3）
'   ・全ての iframe にボタンがあり、押すと親ページの window.__clicked に自分の名前を積む
'
'   例) layout=3,0,0
'       A
'       └A1
'         └A2
'           └A3
'       B
'       C
'
' 確認すること（CDPContext.getElementByIframe 経由）
'   1. 深さ無制限(MaxNestLevel=-1)で名前検索 → その中のボタンを押す → 押されたのが狙った iframe 自身か
'   2. MaxNestLevel = その iframe の深さ    なら見つかる
'   3. MaxNestLevel = その iframe の深さ-1  なら見つからない
'   4. MaxNestLevel 省略は、親ページ直下のみ見つかる
'   5. 全ボタンが押されたか（最大 3×4 = 12 個）
'   6. 条件なしなら、最初の親ページ直下の iframe(A) が返る
'
' 注意事項
'   ・テストの iframe は全て同一オリジン（about:blank に document.write）です。
'     別ドメインの iframe（Edge では別プロセス = OOPIF）の入れ子は、このテストの対象外です
'==============================================================================
Option Explicit

'ワークスペースパス（StarterWebScrapingKit ルート）
Private Const WORKSPACE_PATH As String = ""

Private Const TEST_HTML As String = "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPiframe\NestedIframes.html"

'試す構成。"random" はページ側でランダムに組みます
Private Const TEST_CASES As String = "3,3,3|3,0,0|0|1,2|random|random|random|random|random"

Private Function EOk() As String
    EOk = WorksheetFunction.Unichar(9989)
End Function

'==============================================================================
' エントリ
'==============================================================================
Public Sub RunAll_CDPiframe_NestTests()
    Debug.Print RunCDPiframeNestTests(WORKSPACE_PATH)
End Sub

'------------------------------------------------------------------------------
' 全ケースを実行し、結果の一覧を返します（各行 PASS / FAIL、最終行が集計）
'   workspacePath  StarterWebScrapingKit のルート
'------------------------------------------------------------------------------
Public Function RunCDPiframeNestTests(workspacePath As String) As String
    If LenB(workspacePath) = 0 Then Err.Raise 5, "Test_CDPiframe_Nest", "WORKSPACE_PATH を設定してください。"

    Dim pageUrl As String
    pageUrl = "file:///" & Replace(workspacePath & TEST_HTML, "\", "/")

    Dim br As CDPContext
    Set br = ShSetting01_StartBrowser.StartCDPModeContext("about:blank")
    br.show

    Dim report As String, caseName As Variant, line As String
    Dim passCount As Long, failCount As Long
    report = String(70, "=") & vbCrLf & "  CDPiframe 入れ子探索テスト 開始" & vbCrLf & String(70, "=") & vbCrLf

    For Each caseName In Split(TEST_CASES, "|")
        line = RunOneCase(br, pageUrl, CStr(caseName))
        If Left$(line, 4) = "FAIL" Then failCount = failCount + 1 Else passCount = passCount + 1
        report = report & "  " & IIf(Left$(line, 4) = "FAIL", "", EOk() & " ") & line & vbCrLf
    Next

    '最後のページの構成を、ログにも出しておく（目視確認用）
    br.printChildIframes -1
    br.ThisCDPBrowser.CloseBrowser

    report = report & String(70, "=") & vbCrLf & "  テスト結果: PASS=" & passCount & " / FAIL=" & failCount & " / 合計=" & (passCount + failCount)
    RunCDPiframeNestTests = report
End Function

'==============================================================================
' 1構成分のテスト
'==============================================================================
Private Function RunOneCase(br As CDPContext, pageUrl As String, caseName As String) As String
    '1. ページを開いて、iframe が全部できるのを待つ
    br.navigate pageUrl & IIf(caseName = "random", "", "?layout=" & caseName)
    Dim t0 As Double: t0 = CDPHelpers.TimerCounter
    Do Until br.jsEval("window.__ready===true") = True
        If CDPHelpers.TimerCounter - t0 > 10000 Then RunOneCase = "FAIL | " & caseName & " | page not ready": Exit Function
        CDPHelpers.Sleep 0.1
    Loop

    Dim layoutText As String, layoutInfo As String
    layoutText = br.jsEval("window.__layout.map(function(x){return x.name+':'+x.depth}).join('|')")
    layoutInfo = br.jsEval("document.getElementById('info').textContent")

    '2. iframe ごとに確認（NG があれば ng に理由を積む）
    Dim entry As Variant, frameName As String, frameDepth As Long
    Dim ng As String, lastClicked As String
    Dim doc As CDPElement
    For Each entry In Split(layoutText, "|")
        frameName = Split(entry, ":")(0)
        frameDepth = CLng(Split(entry, ":")(1))

        '2-1. 深さ無制限で見つけて、その中のボタンを押す → 押されたのが狙った iframe か
        Set doc = br.getElementByIframe(frameName, , -1)
        If doc Is Nothing Then
            ng = ng & " notFound(" & frameName & ")"
        Else
            doc.getElementByXPath("//button[@id='btn']").SimpleClick
            lastClicked = br.jsEval("window.__clicked.length?window.__clicked[window.__clicked.length-1]:''")
            If lastClicked <> frameName Then ng = ng & " wrongFrame(" & frameName & "->" & lastClicked & ")"
        End If

        '2-2. MaxNestLevel = 深さちょうど なら見つかる
        If br.getElementByIframe(frameName, , frameDepth) Is Nothing Then ng = ng & " missAtOwnLevel(" & frameName & ")"

        '2-3. MaxNestLevel = 深さ-1 なら見つからない
        If frameDepth > 1 Then
            If Not br.getElementByIframe(frameName, , frameDepth - 1) Is Nothing Then ng = ng & " foundTooShallow(" & frameName & ")"
        End If

        '2-4. MaxNestLevel 省略は親ページ直下のみ
        If (Not br.getElementByIframe(frameName) Is Nothing) <> (frameDepth = 1) Then ng = ng & " defaultLevel(" & frameName & ")"
    Next

    '3. 押されたボタン数 = iframe 数 か
    Dim frameCount As Long, clickedCount As Long
    frameCount = UBound(Split(layoutText, "|")) + 1
    clickedCount = br.jsEval("window.__clicked.length")
    If clickedCount <> frameCount Then ng = ng & " clicked=" & clickedCount & "/" & frameCount

    '4. 条件なしなら、最初の親ページ直下の iframe(A)
    Set doc = br.getElementByIframe()
    If doc Is Nothing Then
        ng = ng & " noArgNotFound"
    Else
        doc.getElementByXPath("//button[@id='btn']").SimpleClick
        lastClicked = br.jsEval("window.__clicked[window.__clicked.length-1]")
        If lastClicked <> "A" Then ng = ng & " noArgPicked(" & lastClicked & ")"
    End If

    RunOneCase = IIf(LenB(ng) = 0, "PASS", "FAIL") & " | " & caseName & " | layout=" & layoutInfo & " | buttons clicked=" & clickedCount & "/" & frameCount & ng
End Function
