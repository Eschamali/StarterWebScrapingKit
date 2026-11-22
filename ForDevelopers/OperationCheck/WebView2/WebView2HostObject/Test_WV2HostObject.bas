Attribute VB_Name = "Test_WV2HostObject"
'***************************************************************************************************
'   `CDPCoreViaWebView2.AddHostObjectToScript`の単体テストです(ブリッジ: `Test_HostObjectBridge`)
'
'   JSの呼び出しは「非同期で投げる → 結果を`window.__r`へ → VBAが`jsEval`ポーリングで拾う」形にしています
'   (JSからのhostObjects呼び出しは、VBAがメッセージを汲み上げたときに実行されるため)
'
'   H01 Add                    : AddHostObjectToScript が成功するか
'   H02 非同期proxyでメソッド    : await bridge.Echo('...')
'   H03 数値の受け渡し          : await bridge.Add(1, 2.5) / VBAに届いた型
'   H04 Property 読み書き       : setHostProperty → VBAで確認 / VBAで変更 → JSで読む
'   H05 Public変数              : await bridge.Count
'   H06 Friendは見えない        : bridge.Init(...) が失敗するか
'   H07 同期proxy               : chrome.webview.hostObjects.sync.bridge.Echo(...)
'   H08 同名で再Add             : 置き換わるか / 前のオブジェクトは cleanupSome() 後に解放されるか
'   H09 Remove                  : JSから呼べなくなるか / VBAでNothingにしたら解放されるか
'   H10 フォームを閉じる         : エラー無く閉じられるか / 残っていたブリッジが解放されるか
'
'   結果は %TEMP%\WV2SharedBufferTest\hostobject_result.txt (UTF-8) とイミディエイトウィンドウへ出力します
'***************************************************************************************************
Option Explicit

Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong

Private m_log        As String
Private m_passCount  As Long
Private m_failCount  As Long
Private m_terminated As String  '解放されたブリッジのtag(カンマ区切り)
Private m_logPath    As String  '1行ごとに書き出す(Class_Terminate内のエラーは捕まえられず、途中で止まるため)



'***************************************************************************************************
'* 機能　　：H01～H10をまとめて実行します
'***************************************************************************************************
Public Sub Test_HostObject_All()
    Dim outDir As String: outDir = Environ$("TEMP") & "\WV2SharedBufferTest"
    If Len(Dir$(outDir, vbDirectory)) = 0 Then MkDir outDir

    m_log = vbNullString
    m_passCount = 0
    m_failCount = 0
    m_terminated = vbNullString
    m_logPath = outDir & "\hostobject_result.txt"

    Dim bridgeA As Test_HostObjectBridge, bridgeB As Test_HostObjectBridge
    Dim res As String
    Dim formClosed As Boolean   'H10で閉じた後に`WebView2Form`を参照すると、既定インスタンスが作り直されるため
    On Error GoTo Catch

    '1. WebView2起動(フォームは表示しない)
    Dim br As CDPBrowser
    Set br = WebView2Form.StartCDPModeWebView2

    Dim wv2 As CDPCoreViaWebView2
    Set wv2 = WebView2Form.ThisWebView2

    Dim ctx As CDPContext
    Set ctx = WebView2Form.ThisCDPContext

    '--- H01 ---
    Set bridgeA = New Test_HostObjectBridge
    bridgeA.Init "A"
    LogResult "H01 Add", wv2.AddHostObjectToScript("bridge", bridgeA), "hostObjects.bridge = " & ctx.jsEval("typeof chrome.webview.hostObjects.bridge")

    '--- H02 ---
    res = RunJs(ctx, "chrome.webview.hostObjects.bridge.Echo('こんにちは')")
    LogResult "H02 非同期proxy Echo", (res = "ok:こんにちは(A)") And (bridgeA.Count = 1), "JS=[" & res & "] VBA側Count=" & bridgeA.Count

    '--- H03 ---
    res = RunJs(ctx, "chrome.webview.hostObjects.bridge.Add(1, 2.5)")
    LogResult "H03 数値の受け渡し", (res = "ok:3.5"), "JS=[" & res & "] VBAに届いた型=" & bridgeA.ArgTypes

    '--- H04 ---
    res = RunJs(ctx, "chrome.webview.hostObjects.bridge.setHostProperty('Value', 'JSから')")
    Dim fromJs As String: fromJs = bridgeA.Value
    bridgeA.Value = "VBAから"
    Dim res2 As String: res2 = RunJs(ctx, "chrome.webview.hostObjects.bridge.Value")
    LogResult "H04 Property 読み書き", (fromJs = "JSから") And (res2 = "ok:VBAから"), "set=[" & res & "] VBAで見た値=[" & fromJs & "] JSで読んだ値=[" & res2 & "]"

    '--- H05 ---
    res = RunJs(ctx, "chrome.webview.hostObjects.bridge.Count")
    LogResult "H05 Public変数", (res = "ok:" & bridgeA.Count), "JS=[" & res & "] VBA側Count=" & bridgeA.Count

    '--- H06 ---
    res = RunJs(ctx, "chrome.webview.hostObjects.bridge.Init('X')")
    LogResult "H06 Friendは見えない", (Left$(res, 4) = "err:") And (bridgeA.Tag = "A"), "JS=[" & res & "] tag=" & bridgeA.Tag

    '--- H07 ---
    res = RunJs(ctx, "new Promise((ok, ng) => setTimeout(() => { try { ok(chrome.webview.hostObjects.sync.bridge.Echo('sync')); } catch (e) { ng(e); } }, 0))")
    LogResult "H07 同期proxy Echo", (res = "ok:sync(A)"), "JS=[" & res & "]"

    '--- H08 ---
    Set bridgeB = New Test_HostObjectBridge
    bridgeB.Init "B"
    Dim readd As Boolean: readd = wv2.AddHostObjectToScript("bridge", bridgeB)
    res = RunJs(ctx, "chrome.webview.hostObjects.bridge.Echo('re')")
    LogResult "H08-1 同名で再Add", readd And (res = "ok:re(B)"), "Add=" & readd & " JS=[" & res & "]"

    Set bridgeA = Nothing
    DoEvents
    '置き換えられた側は、JS側のproxyが参照を握っているため、この時点ではまだ解放されない(WebView2の仕様)
    LogLine "[INFO] H08-2 置き換え直後(JSのproxyが握っている) : 解放済み=[" & m_terminated & "]"

    'cleanupSome()でJS側に手放させると、解放される
    res = RunJs(ctx, "(chrome.webview.hostObjects.cleanupSome(), 'cleanupSome')")
    Dim tWait As LongLong: tWait = GetTickCount64()
    Do Until IsTerminated("A") Or GetTickCount64() - tWait > 2000
        DoEvents
    Loop
    LogResult "H08-3 cleanupSome後に置き換えられた側の解放", IsTerminated("A"), "JS=[" & res & "] 解放済み=[" & m_terminated & "]"

    '--- H09 ---
    Dim removed As Boolean: removed = wv2.RemoveHostObjectFromScript("bridge")
    res = RunJs(ctx, "chrome.webview.hostObjects.bridge.Echo('after remove')")
    LogResult "H09-1 Remove後はJSから呼べない", removed And (Left$(res, 4) = "err:"), "Remove=" & removed & " JS=[" & res & "]"

    Set bridgeB = Nothing
    DoEvents
    LogResult "H09-2 Remove後にNothingで解放", IsTerminated("B"), "解放済み=[" & m_terminated & "]"

    '--- H10 ---
    Dim bridgeC As Test_HostObjectBridge
    Set bridgeC = New Test_HostObjectBridge
    bridgeC.Init "C"
    wv2.AddHostObjectToScript "bridgeC", bridgeC
    Set bridgeC = Nothing   'WebView2(と管理Dictionary)だけが持っている状態にする
    Set wv2 = Nothing
    Set ctx = Nothing
    Set br = Nothing

    Dim errNo As Long, errDesc As String
    On Error Resume Next
    Unload WebView2Form
    errNo = Err.Number: errDesc = Err.Description
    formClosed = True
    On Error GoTo Catch
    DoEvents
    LogResult "H10-1 フォームを閉じる", (errNo = 0), "Err=" & errNo & " " & errDesc
    LogResult "H10-2 残っていたブリッジの解放", IsTerminated("C"), "解放済み=[" & m_terminated & "]"
    GoTo Finally

Catch:
    LogResult "RuntimeError", False, Err.Number & ": " & Err.Description
    Resume Finally

Finally:
    On Error Resume Next
    Set bridgeA = Nothing
    Set bridgeB = Nothing
    If Not formClosed Then Unload WebView2Form
    LogLine "=== PASS " & m_passCount & " / FAIL " & m_failCount & " ==="
End Sub

'* ブリッジの`Class_Terminate`から呼ばれる
Public Sub NotifyTerminated(ByVal tag As String)
    m_terminated = m_terminated & IIf(Len(m_terminated) > 0, ",", "") & tag
End Sub

Private Function IsTerminated(ByVal tag As String) As Boolean
    IsTerminated = (InStr("," & m_terminated & ",", "," & tag & ",") > 0)
End Function



'***************************************************************************************************
'                                   ■■■ JS ■■■
'***************************************************************************************************
'* 機能　　：式をawaitして、「ok:結果」か「err:メッセージ」を返します(タイムアウト時は空文字)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：jsEval自体は投げっぱなしにして、結果は`window.__r`をポーリングで拾います
'***************************************************************************************************
Private Function RunJs(ctx As CDPContext, ByVal expr As String, Optional ByVal timeoutSec As Double = 10) As String
    Dim js As String
    js = "window.__r = undefined;"
    js = js & "(async () => { try { const v = await (" & expr & "); window.__r = 'ok:' + String(v); }"
    js = js & " catch (e) { window.__r = 'err:' + ((e && e.message) || e); } })(); 0"
    ctx.jsEval js

    Dim t0 As LongLong: t0 = GetTickCount64()
    Dim v As Variant
    Do
        v = ctx.jsEval("window.__r === undefined ? '' : String(window.__r)")
        If Not IsError(v) Then
            If Len(v) > 0 Then RunJs = v: Exit Function
        End If
        DoEvents
    Loop Until GetTickCount64() - t0 > timeoutSec * 1000
End Function



'***************************************************************************************************
'                                   ■■■ ログ ■■■
'***************************************************************************************************
Private Sub LogResult(ByVal testName As String, ByVal passed As Boolean, ByVal detail As String)
    If passed Then m_passCount = m_passCount + 1 Else m_failCount = m_failCount + 1
    LogLine IIf(passed, "[PASS] ", "[FAIL] ") & testName & " : " & detail
End Sub

Private Sub LogLine(ByVal s As String)
    Debug.Print s
    m_log = m_log & s & vbCrLf
    If Len(m_logPath) > 0 Then WriteLog m_logPath
End Sub

'* UTF-8で保存
Private Sub WriteLog(ByVal filePath As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    With st
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText m_log
        .SaveToFile filePath, 2
        .Close
    End With
End Sub
