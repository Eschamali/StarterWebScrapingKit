window.docsContent = window.docsContent || {};

window.docsContent['bidi-methods'] = `
    <h1>メソッドの使い方 - <span style="color:var(--accent-color)">WebDriverBiDiCore</span></h1>
    <p>当ライブラリの中核となる <code>WebDriverBiDiCore.cls</code> を利用した、高度なブラウザ操作の具体的な手順を解説します。</p>

    <h2>1. 準備とブラウザの起動</h2>
    <pre><code class="language-vb">Sub StartBiDi()
    Dim bidi As New WebDriverBiDiCore
    
    ' 空のページでブラウザを起動
    bidi.start "about:blank"
    
    ' --- ここに操作コードを記述 ---
    
    ' 操作完了後にブラウザを終了
    bidi.quit
End Sub</code></pre>

    <h2>2. ページへの遷移 (browsingContext.navigate)</h2>
    <p>現在のタブコンテキストを取得し、指定URLに完全ロードされるまで待機する基本フローです。</p>
    <pre><code class="language-vb">Dim resultBiDi As Dictionary
Dim targetContext As String

' 1. 現在のコンテキストツリーを取得
Set resultBiDi = bidi.invokeMethod("browsingContext.getTree")
targetContext = resultBiDi("contexts")(1)("context")

' 2. URLへ遷移するパラメータを構築
Dim paramsBiDi As New Dictionary
paramsBiDi.Add "context", targetContext
paramsBiDi.Add "url", "https://google.com/"
paramsBiDi.Add "wait", "complete" ' DOM構築完了まで待機

' 3. 遷移コマンドの発行
bidi.invokeMethod "browsingContext.navigate", paramsBiDi</code></pre>

    <div class="alert warning">
        <div class="alert-icon">🛡️</div>
        <div class="alert-content">
            <strong>エラーハンドリング (StopError:=False)</strong>
            <p><code>invokeMethod</code> の引数に <code>StopError:=False</code> を渡すと、処理失敗時にVBAが強制終了せず <code>Nothing</code> を返します。その際、<code>bidi.LastBiDiJsonError</code> から詳細なエラー内容を調べ、自力で回復処理（フォールバック）を実装可能です。</p>
        </div>
    </div>

    <h2>3. 非同期イベントの購読</h2>
    <p>BiDiの最大の強みである非同期通信の例（Console Logの傍受）です。</p>
    <pre><code class="language-vb">' ログ関連イベントの購読を開始
Dim subscribeParams As New Dictionary
subscribeParams.Add "events", Array("log.entryAdded")
bidi.invokeMethod "session.subscribe", subscribeParams

' イベントバッファを読み取る
bidi.TakeEvents

' 貯まった非同期イベントをループで確認
Dim evt As Variant
For Each evt In bidi.Events
    If evt("method") = "log.entryAdded" Then
        Debug.Print "JS Console Output: " & evt("params")("text")
    End If
Next evt</code></pre>
`;
