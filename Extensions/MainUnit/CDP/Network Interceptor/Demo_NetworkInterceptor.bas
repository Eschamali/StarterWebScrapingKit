Attribute VB_Name = "Demo_NetworkInterceptor"
Option Explicit

' =========================================================
' exCDP_NetworkInterceptor のデモマクロ
' =========================================================
' [前提]
'   ・このモジュールと exCDP_NetworkInterceptor.cls を
'     VBAProject にインポートしてから実行してください。
'
' [テスト一覧]
'   [A] AddBlockedURL    : 特定URLへのfetchをブロック
'   [B] AddMockResponse  : 特定URLへのfetchを偽レスポンスで差し替え
'   [C] WaitForResponse  : 実際に発生したレスポンスをCDPイベントで検知 + Body取得
' =========================================================

Public Sub Demo_NetworkInterceptor_All()
    Dim tb As CDPContext
    Dim ni As exCDP_NetworkInterceptor

    Set tb = 設定シートからのCDP起動ForTab
    tb.navigate "about:blank"
    Set ni = New exCDP_NetworkInterceptor
    ni.Init tb

    tb.printMsg info_, "================================================", "Demo"
    tb.printMsg info_, "  exCDP_NetworkInterceptor デモ開始", "Demo"
    tb.printMsg info_, "================================================", "Demo"
    tb.InheritanceCDPBrowser.sleep 1

    Demo_A_BlockURL tb, ni
    tb.InheritanceCDPBrowser.sleep 1
    Demo_B_MockResponse tb, ni
    tb.InheritanceCDPBrowser.sleep 1
    Demo_C_WaitForResponse tb, ni

    tb.printMsg info_, "================================================", "Demo"
    tb.printMsg info_, "  全テスト完了", "Demo"
    tb.printMsg info_, "================================================", "Demo"
    tb.InheritanceCDPBrowser.sleep 2
    tb.InheritanceCDPBrowser.quit
End Sub


' =========================================================
' [A] URL ブロック テスト
' =========================================================
Private Sub Demo_A_BlockURL(tb As CDPContext, ni As exCDP_NetworkInterceptor)
    tb.printMsg info_, WorksheetFunction.Unichar(9654) & " [A] URLブロック テスト開始", "Demo"

    ' httpbin.org をブロック登録
    ni.AddBlockedURL "httpbin.org"

    ' ページのJSからブロック対象URLへfetchしてみる
    ' fetch(...).then().catch() をそのまま Promise に渡す形式（末尾セミコロンなし）
    Dim js As String
    js = "new Promise(function(resolve){"
    js = js & "fetch('https://httpbin.org/get')"
    js = js & ".then(function(r){ resolve('status:' + r.status); })"
    js = js & ".catch(function(e){ resolve('blocked:' + e.message); });"
    js = js & "});"

    ' awaitPromise=True で結果を待つ
    Dim res As String
    res = tb.jsEval(js, awaitPromise:=True)
    ' 先頭末尾の " を除去
    If Left(res, 1) = """" Then res = Mid(res, 2)
    If Right(res, 1) = """" Then res = Left(res, Len(res) - 1)

    tb.printMsg info_, "  fetch結果: " & res, "Demo"

    If InStr(1, res, "blocked", vbTextCompare) > 0 Or InStr(1, res, "Blocked", vbTextCompare) > 0 Then
        tb.printMsg info_, "  " & WorksheetFunction.Unichar(10004) & " URLブロック成功！fetchがエラーになりました。", "Demo"
        MsgBox "[A] URLブロック成功！" & vbCrLf & "結果: " & res, vbInformation, "Demo"
    Else
        tb.printMsg WARN_, "  " & WorksheetFunction.Unichar(10008) & " URLブロックが効いていません。結果: " & res, "Demo"
        MsgBox "[A] URLブロック未検出。" & vbCrLf & "結果: " & res, vbExclamation, "Demo"
    End If

    ' クリーンアップ
    ni.ClearBlockedURLs
End Sub


' =========================================================
' [B] モックレスポンス テスト
' =========================================================
Private Sub Demo_B_MockResponse(tb As CDPContext, ni As exCDP_NetworkInterceptor)
    tb.printMsg info_, WorksheetFunction.Unichar(9654) & " [B] モックレスポンス テスト開始", "Demo"

    ' /api/user への通信を偽レスポンスで差し替える
    Dim mockJson As String
    mockJson = "{""name"":""Test Taro"",""role"":""admin"",""mock"":true}"
    ni.AddMockResponse "/api/user", 200, mockJson, "application/json"

    ' /api/user へ fetch してレスポンスを取得
    Dim js As String
    js = "new Promise(function(resolve){"
    js = js & "fetch('/api/user')"
    js = js & ".then(function(r){ return r.text(); })"
    js = js & ".then(function(t){ resolve(t); })"
    js = js & ".catch(function(e){ resolve('error:'+e.message); })"
    js = js & "});"

    Dim res As String
    res = tb.jsEval(js, awaitPromise:=True)
    ' 先頭末尾の " を除去
    If Left(res, 1) = """" Then res = Mid(res, 2)
    If Right(res, 1) = """" Then res = Left(res, Len(res) - 1)

    tb.printMsg info_, "  fetchレスポンス: " & res, "Demo"

    If InStr(1, res, "mock", vbTextCompare) > 0 Then
        tb.printMsg info_, "  " & WorksheetFunction.Unichar(10004) & " モックレスポンス成功！", "Demo"
        MsgBox "[B] モックレスポンス成功！" & vbCrLf & "Body: " & res, vbInformation, "Demo"
    Else
        tb.printMsg WARN_, "  " & WorksheetFunction.Unichar(10008) & " モックレスポンスが効いていません。結果: " & res, "Demo"
        MsgBox "[B] モック未検出。" & vbCrLf & "結果: " & res, vbExclamation, "Demo"
    End If

    ' クリーンアップ
    ni.ClearMockResponses
End Sub


' =========================================================
' [C] レスポンス待機 + Body取得 テスト（CDPネイティブ）
' =========================================================
Private Sub Demo_C_WaitForResponse(tb As CDPContext, ni As exCDP_NetworkInterceptor)
    tb.printMsg info_, WorksheetFunction.Unichar(9654) & " [C] WaitForResponse(CDPネイティブ) テスト開始", "Demo"

    ' CDPの Network ドメインを有効化してキャプチャ開始
    ni.StartNetworkCapture
    tb.printMsg info_, "  Network.enable 完了。fetchを発行します...", "Demo"

    ' 実際に外部APIへリクエストを発行（httpbin.org は json を返す無料エンドポイント）
    tb.jsEval "fetch('https://httpbin.org/json');"
    tb.printMsg info_, "  fetch発行完了。レスポンスを待機中...", "Demo"

    Dim reqId As String
    reqId = ni.WaitForResponse("httpbin.org/json", 15)

    If reqId <> "" Then
        tb.printMsg info_, "  " & WorksheetFunction.Unichar(10004) & " レスポンス検出！ requestId=" & reqId, "Demo"

        ' Body を取得
        Dim Body As String
        Body = ni.GetResponseBody(reqId)

        tb.printMsg info_, "  レスポンスBody(先頭100文字): " & Left(Body, 100), "Demo"
        MsgBox "[C] WaitForResponse 成功！" & vbCrLf & _
               "requestId: " & reqId & vbCrLf & vbCrLf & _
               "Body(先頭100文字):" & vbCrLf & Left(Body, 100), _
               vbInformation, "Demo"
    Else
        tb.printMsg WARN_, "  " & WorksheetFunction.Unichar(10008) & " タイムアウト：レスポンスが検出できませんでした。", "Demo"
        MsgBox "[C] WaitForResponse タイムアウト。", vbExclamation, "Demo"
    End If

    ' キャプチャ停止
    ni.StopNetworkCapture
End Sub
