---
description: Browser / Page / Element の三層モデルは3者共通。では何が違うのか。クラス分割の粒度、抽象化レイヤーの有無、プロパティ構文による API 表現の差を比較します。
---

# クラス構成の考え方

ここまでは通信層の話でした。このページは**利用者が実際に触る層**の設計を比較します。

## 1. 三層モデルは完全に一致している

まず結論として、3者の骨格は同じです。

| 層 | Puppeteer | Playwright | StarterWebScrapingKit |
| --- | --- | --- | --- |
| ブラウザ | `Browser` / `BrowserContext` | `Browser` / `BrowserContext` | `CDPBrowser` |
| ページ | `Page` / `Frame` | `Page` / `Frame` | `CDPContext` |
| 要素 | `ElementHandle` / `JSHandle` | `ElementHandle` / `JSHandle` | `CDPElement` |
| 通信 | `Connection` / `CDPSession` | `CRConnection` / `CRSession` | `CDPCore` |

これは偶然ではなく、CDP のドメイン構造（`Target` / `Page` / `DOM` / `Runtime`）が自然にこの分割を要求するためです。[アーキテクチャ](/concepts/architecture) で説明している構造が、そのまま Puppeteer / Playwright にも当てはまります。

## 2. 粒度は桁が違う

一方で、ファイル数を並べると差が明白です。

| | 主要な実装ファイル数 |
| --- | --- |
| **Puppeteer** | `api/` 約24 + `cdp/` 約47 + `bidi/` 約27 |
| **Playwright** | `server/` 約50 + `server/chromium/` 約16（他に `firefox/` `webkit/` `bidi/` `android/` `electron/`） |
| **StarterWebScrapingKit** | CDP スタック6クラス + BiDi スタック3クラス |

キット側の実クラスは以下の規模です。

| クラス | 行数 | 役割 |
| --- | ---: | --- |
| `BiDiCDPJson` | 2,998 | JSON パーサ／ビュー |
| `CDPContext` | 2,889 | 1タブ分の操作 |
| `CDPBrowser` | 1,794 | プロセス起動・タブ管理 |
| `CDPCore` | 1,600 | Pipe 送受信 |
| `CDPElement` | 1,590 | 要素操作 |
| `CDPCoreViaWebSocket` | 1,335 | WebSocket 送受信 |

Puppeteer の `cdp/` にある `FrameManager` / `NetworkManager` / `TargetManager` / `LifecycleWatcher` / `IsolatedWorld` / `ExecutionContext` といった中間層が、キットでは `CDPContext` の内部に畳み込まれている、と考えると近いです。

### なぜ分割しないのか

VBA の言語制約が直接効いています。

- **名前空間がない** — すべてのクラスがグローバルに1つのフラットな名前空間へ並ぶ。`Frame` のような一般名は他のライブラリと衝突する
- **フォルダ分けができない** — VBE のプロジェクトエクスプローラは1階層のみ。50クラスを並べても探せない
- **`import` がない** — 依存関係がコード上に現れないため、クラスを増やすほど「誰が誰を呼ぶか」が読めなくなる
- **配布単位がファイル** — 利用者は `.cls` を手でインポートする。数が増えるほど導入手順が壊れやすくなる

::: info トレードオフとして認識すべき点
「1クラス3000行」は、Node 側の基準では明確に大きすぎます。分割しない選択は VBA の制約への適応であって、設計として優れているわけではありません。実際、`CDPContext` の中では責務が混在しています。
:::

## 3. 抽象化レイヤーの有無 ―― ここは本質的な差

Puppeteer と Playwright は、**プロトコル実装から独立した抽象層**を持っています。

```
Puppeteer:  api/Page.ts（抽象クラス）
              ├─ cdp/Page.ts   （CDP 実装）
              └─ bidi/Page.ts  （BiDi 実装）

Playwright: server/page.ts（共通実装）
              ├─ chromium/crPage.ts
              ├─ firefox/ffPage.ts
              └─ webkit/wkPage.ts
```

利用者は `page.click()` と書くだけで、裏が CDP でも BiDi でも Firefox でも同じように動きます。

キットにも CDP スタックと BiDi スタックの両方がありますが、**共通の抽象基底を持たない並列構造**です。

| | CDP スタック | BiDi スタック |
| --- | --- | --- |
| ブラウザ | `CDPBrowser` | `WebDriverBiDiMode` |
| ページ | `CDPContext` | `WebDriverBiDiContext` |
| 要素 | `CDPElement` | （CDP に変換して使う） |

BiDi 側から要素を細かく触りたい場合は `ConvertToCDPContext` で CDP 側へ橋渡しします。これは「抽象化する代わりに、相互変換で繋ぐ」というアプローチです。

VBA にも `Implements` によるインターフェースはあり（キット内でも `IWebAuthenticator.cls` で使われています）、抽象化自体は不可能ではありません。採用していないのは、**BiDi 側の要素操作 API がまだ CDP 側ほど成熟していない**という現実的な事情によるものです。ここは Puppeteer / Playwright に対して明確に劣る点として認識しておくべきところです。

## 4. 逆に VBA が有利な一点 ―― プロパティ構文

要素操作の書き味では、VBA 側が読みやすくなる場面があります。

```ts
// Puppeteer
const text = await el.evaluate(e => e.innerText);
await el.evaluate((e, v) => { e.value = v; }, '入力値');
```

```vb
' StarterWebScrapingKit
Dim text As String: text = el.innerText
el.value = "入力値"
```

`CDPElement` は `innerText` / `innerHTML` / `value` / `checked` / `selected` を `Property Get` と `Property Let` の両方で公開しているため、**DOM の書き味がそのまま VBA の代入文になります**。

これは VBA のプロパティ構文が「読み書き両用の見た目」を作れることによるもので、`await` が必要な JavaScript では原理的に真似できません。同期実行が前提であることが、ここでは記述性の利点として働いています。

要素の待機系も同様にプロパティ／メソッドとして畳み込まれています。

```vb
el.onExist       ' 出現するまで待つ
el.onExistNot    ' 消えるまで待つ
el.ifExist       ' 存在すれば真
```

## 5. Handle のライフサイクル管理

Node 側は `JSHandle` / `ElementHandle` に明示的な `dispose()` を持ち、`Runtime.releaseObject` でブラウザ側のメモリを解放します。`Realm` / `IsolatedWorld` / `ExecutionContext` といったクラスがその管理を担当しています。

キットの `CDPElement` は `objectId` を保持しますが、解放は VBA の参照カウント（`Class_Terminate`）と、ページ遷移時にブラウザ側のコンテキストが破棄されることに依存します。長時間 1 ページに滞在して大量の要素を取得し続けるようなケースでは、Node 側のほうが厳密です。

## 6. まとめ

| 観点 | 評価 |
| --- | --- |
| Browser / Page / Element の三層モデル | ✅ 完全に同じ発想 |
| CDP ドメインの隠蔽（利用者は CDP を知らなくていい） | ✅ 達成している |
| 要素操作の記述性 | ✅ プロパティ構文でむしろ簡潔 |
| クラス分割の粒度 | ❌ VBA の制約により粗い（1クラス3000行級） |
| プロトコル非依存の抽象層 | ❌ CDP / BiDi が並列構造で共通基底なし |
| Handle の明示的な解放 | ❌ 参照カウント任せ |

## 次に読む

- [残る差分と、埋まらない差](/core-comparison/gaps) — テスト・エラー処理・型安全性
- [アーキテクチャ](/concepts/architecture) — キット側のクラス構成の詳細
- [CDP と BiDi](/concepts/cdp-vs-bidi) — 2つのスタックの使い分け
