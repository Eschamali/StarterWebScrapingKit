---
description: Promise / async-await とイベントループを前提とした Puppeteer・Playwright に対し、イベントループを持たない VBA で非同期実行をどう成立させているかを比較します。
---

# 非同期実行とイベントループ

ここが3者で**もっとも大きく違う**領域です。ただし違いの原因は設計の優劣ではなく、言語ランタイムそのものにあります。

## 1. Node 側 ―― 「送る」と「待つ」が分離している

Playwright の送信部分は驚くほど短いコードです。

```ts
// playwright-core/src/server/chromium/crConnection.ts
send<T>(method: T, params?): Promise<...> {
  const id = this._connection._rawSend(this._sessionId, method, params);
  return new Promise((resolve, reject) => {
    this._callbacks.set(id, { resolve, reject, error: new ProtocolError('error', method) });
  });
}
```

やっていることは3行です。

1. コマンドを送って `id` を得る
2. その `id` に紐づく `resolve` / `reject` を `Map` に登録する
3. Promise を返してすぐ抜ける

**待たずに戻る**のがポイントです。あとは前ページで見たディスパッチ処理が応答を受け取ったときに `resolve` を呼び、`await` していた箇所が再開されます。Puppeteer も `CallbackRegistry` という専用クラスを挟むだけで、構造は同じです。

これが成立する土台が **libuv のイベントループ**です。ソケットにデータが届いたことを OS が通知し、ランタイムが対応するコールバックを自動で起動してくれます。開発者は「いつ受信処理を回すか」を一切考えなくて済みます。

## 2. VBA 側 ―― イベントループが存在しない

VBA（Excel VBA / STA シングルスレッド）には、これに相当する仕組みがありません。

- OS からの I/O 完了通知を拾って自動でコールバックを起動する層がない
- `async` / `await` 相当の構文がない
- Promise のような「未完了の値」を表す標準の型もない

つまり **「誰かが能動的にパイプを覗きに行かない限り、受信データは永久に処理されない」** という前提から始まります。

## 3. StarterWebScrapingKit の解 ―― 2階建ての API

キットは、この制約に対して**2つの実行モードを用意する**という答えを出しています。

### 3-1. 同期モード（`ExecuteCDP`）― ポンプを回して待つ

```vb
Do
    '1. 受信処理によるパイプを拝見する ※内部にて`DoEvents`が走ります
    TakeEvents StopApiError

    '2. 内部の`RaiseEvent CDPBrowserID`による発火で、Dictionaryへ蓄積されたか？
    AutoWaitTakeResultCDP = TakeResultCDP(commandID)
    If StrPtr(AutoWaitTakeResultCDP) Then Exit Do

    '3. 受信処理にてエラーがあったら、即抜け
    If PipeCore.LastErrorPeekNamedPipe > 0 Or PipeCore.LastErrorReadFile > 0 Then ... Exit Function

    '4. 一定時間経ってもコマンド結果が来ないなら、エラーで停止します
    If TimerCounter - timerStart > timerOut Then Err.Raise CDPCustomErrorCodes.TIMEOUT, ...
Loop
```

イベントループが無いので、**自前でループを回して自分のコマンド ID が返るのを待ちます**。「ポーリング」と言うと素朴に聞こえますが、実際には次の要素が揃っています。

- `TakeEvents` が届いている分を**まとめて**吸い上げ、他のイベントも同時に処理する
- 自分宛て以外の応答も捨てずに `Dictionary` へ蓄積する（後述の非同期モードが機能する理由）
- タイムアウト監視あり
- パイプ側のエラーを毎周チェックし、即座に離脱する

Playwright も Puppeteer も「ID を Map に登録して、応答が来たら取り出す」ことに変わりはありません。違うのは、**その取り出しを誰が駆動するか**だけです。

### 3-2. 非同期モード（`ExecuteCDPAsync`）― 整理券方式

結果を待ちたくない場合は、コマンド ID（＝整理券）だけを受け取って先へ進めます。

```vb
Public Function ExecuteCDPAsync(methodName As String, Optional params As Scripting.Dictionary, _
                                Optional StopApiError As Boolean = True) As Long
    ' ブラウザへ送信し、実行時の commandID をそのまま返す（整理券の発行）
    ExecuteCDPAsync = PipeCore.ReadyRunCDP(CDPcommand, brTab.sessionID)
End Function
```

これは Promise とほぼ同じ発想です。`await` の代わりに、後から `TakeResultCDP(commandID)`（もしくは自動待機版の `AutoWaitTakeResultCDP(commandID)`）で回収します。

```vb
' 3タブに一斉に navigate を投げてから、あとで結果を回収する
Dim id1 As Long: id1 = tabA.ExecuteCDPAsync("Page.navigate", paramsA)
Dim id2 As Long: id2 = tabB.ExecuteCDPAsync("Page.navigate", paramsB)
Dim id3 As Long: id3 = tabC.ExecuteCDPAsync("Page.navigate", paramsC)

Do
    br.TakeEvents                       ' 受信ポンプを回す（CDPCore は共通）
    If LenB(tabA.TakeResultCDP(id1)) Then Exit Do
    DoEvents
Loop
```

3件を送ってから回収するので、**通信としては `Promise.all` と同じく往復が重ならず1回分にまとまります**。VBA 側の処理が並列に走るわけではありませんが、待ち時間の重ね合わせという実利は得られます（画面上のタブが一括で切り替わるのが見える程度には効きます）。詳しい書き方は [生プロトコル拡張](/guides/extend-raw-protocol) にあります。

| | 「待たずに送る」 | 「あとで結果を取る」 | 複数コマンドの往復まとめ |
| --- | --- | --- | --- |
| **Playwright / Puppeteer** | `send()` が Promise を返す | `await` | `Promise.all` |
| **StarterWebScrapingKit** | `ExecuteCDPAsync` が ID を返す | `TakeResultCDP(id)` | 連続 Async 発行 → まとめて回収 |

::: tip Promise との決定的な差
Promise は「解決したら自動で続きが動く」のに対し、整理券は「取りに行かないと何も起きない」点が違います。回収を忘れると結果は `Dictionary` に溜まったままです。
:::

## 4. `DoEvents` ―― 手動のイールド

VBA でループを回し続けると、Excel の UI がフリーズします。それを避けるため、受信ループには `DoEvents` が仕込まれています。

```vb
Private Const RunDoEventsCount As Long = 2 ^ 10   '長いループ中に`DoEvents`を挟む回数

' ...
'一定の回数ごとに`DoEvents`を呼ぶ。※初回ループは、必ず呼ぶ
If LoopCounter Mod RunDoEventsCount = 0 Then DoEvents
```

`DoEvents` は Windows のメッセージキューを処理する呼び出しで、 **イベントループの「手回し版」** にあたります。ただし1回あたりのコストが小さくないため、毎周ではなく **1024 回に1回**という間引きが入っています。無条件に呼ぶ実装よりも実効速度が出るチューニングです。

::: warning DoEvents の副作用
`DoEvents` 中はユーザーのセル操作や他マクロの起動を許してしまいます。長時間の待機処理では、この点を踏まえた設計が必要です。逆に、これがあるおかげで待機中でも「中断」ボタンが効きます。
:::

## 5. 「待つ」の実装比較

ページ読み込み完了を待つ処理を並べると、差がはっきりします。

```ts
// Playwright — 内部でイベントを Promise 化して await
await page.goto(url, { waitUntil: 'load' });
```

```vb
' StarterWebScrapingKit — 状態をポーリングして待つ
Public Sub wait(Optional till As ReadyState = isComplete, Optional dbgState As Boolean = False)
    ' ...
    sleep 0.1   'reduce sleep will speed up but will cost cpu power
    ' ...
End Sub
```

コメントにある通り、**待機間隔は「速度」と「CPU 負荷」のトレードオフ**であり、そこに正解値はありません。イベント駆動なら本質的に発生しないはずのチューニング項目が、ポーリング方式では設計パラメータとして表に出てきます。

## 6. 何が本質的な差なのか

| 観点 | Puppeteer / Playwright | StarterWebScrapingKit | 差の原因 |
| --- | --- | --- | --- |
| コマンド ID と応答の待ち合わせ | `Map` に callback 登録 | `Dictionary` に結果を蓄積 | 同等 |
| 待たずに送る手段 | Promise | 整理券（コマンド ID） | 同等（回収が手動） |
| 受信の駆動 | libuv が自動起動 | `TakeEvents` を手動で呼ぶ | **言語ランタイム** |
| 待機中の他処理 | 他の Promise が普通に進む | `DoEvents` の範囲のみ | **言語ランタイム** |
| 並行実行 | `Promise.all` で自然に書ける | 連続 Async 発行で往復はまとめられる | 記述性の差 |
| キャンセル | `AbortSignal` 等 | タイムアウト到達で `Err.Raise` | 実装投資 |

イベントループの有無だけは、VBA 側でどれだけ工夫しても埋まりません。逆に言えば、**それ以外の要素（ID 管理・整理券・タイムアウト・パイプライン化）はすべて実装で埋められており、実際に埋めてある**というのがこのキットの位置づけです。

## まとめ

| 観点 | 評価 |
| --- | --- |
| 「送る」と「待つ」の分離 | ✅ 整理券方式で実現済み |
| 往復のパイプライン化 | ✅ 連続 Async 発行で可能 |
| タイムアウト・エラー離脱 | ✅ 受信ループに組み込み済み |
| UI をブロックしない配慮 | ✅ `DoEvents` の間引き呼び出し |
| 自動的な受信駆動 | ❌ 言語に存在しないため手動 |
| `await` の記述性 | ❌ 構文が無いため冗長 |

## 次に読む

- [クラス構成の考え方](/core-comparison/classes) — API 表層の設計はどう違うのか
- [ディスパッチとイベント配信](/core-comparison/dispatch) — `TakeEvents` が何をしているか
- [タイムアウト設計](/guides/timeout) — 実運用での待機設定
