---
description: 3プロジェクトが選んだ3つの異なるファイル分割軸と、VBAChromeDevProtocol が CDP 公式仕様から 236 ファイルを自動生成した「型安全への別解」を検証します。
---

# クラス構成とコード生成

同じ CDP を相手にしているのに、3プロジェクトは**ファイルを切る軸そのものが違いました**。

| | 分割の軸 | ブラウザ / タブの扱い |
| --- | --- | --- |
| **StarterWebScrapingKit** | スコープ階層 × 機能の柱 | `CDPBrowser` / `CDPContext` / `CDPElement` で明確に別ファイル |
| **VBAChromeDevProtocol** | CDP プロトコルドメイン | `AutomateBrowser` が両方を兼務（セッションは辞書） |
| **vba-cdp-webdriver** | 処理パイプラインの段階 | `ChromeDriver` が兼務（そもそも分離不要） |

## StarterWebScrapingKit：スコープの階層

```
src/classes/
├── CDPCore.cls / CDPCoreViaWebSocket.cls   ← 通信層（Pipe / WebSocket）
├── CDPCoreViaWebView2.cls                  ← 通信層（WebView2、v3.0.0〜）
├── CDPCoreHost.cls                         ← ローカルブラウザの起動・後始末（v3.0.0〜）
├── CDPWebView2Host.cls / CallbackHandler.cls ← WebView2 COMコールバックの受け皿（v3.0.0〜）
├── CDPBrowser.cls                          ← ブラウザ単位
├── CDPContext.cls                          ← タブ単位
├── CDPElement.cls                          ← 要素単位
├── WebDriverBiDiCore / Context / Mode.cls  ← 別プロトコル（BiDi）の柱
├── WebClient / WebRequest / WebResponse.cls ← REST WebAPI の柱
├── WebSocketCommunicator / HTTPCommunicator ← 汎用 WebSocket 通信の柱
└── Dictionary.cls                          ← `Scripting.Dictionary`互換の自作クラス（v2.4.2〜）
```

「ブラウザ → タブ → 要素」というスコープの階層と、「CDP / BiDi / REST / WebSocket / WebView2」という機能の柱の軸です。これは[アーキテクチャのページ](/concepts/architecture)で説明している通りで、Puppeteer / Playwright が採用している軸と一致します。

## vba-cdp-webdriver：処理の段階

```
src/
├── a1_WebSocketCommunicator / a1x1_HTTPCommunicator   ← ① 通信層
├── a2_JSONHandler                                      ← ② JSON 処理層
├── a3_BasicInfos                                       ← ③ 状態保持層
├── a4_ExecuteCDP / a4x1_MessageGenerator / a4x2_MessageHandler ← ④ コマンド実行層
├── a5_CDPEventHandler                                  ← ⑤ イベント処理層
├── a6_ExecuteHelperFunction                            ← ⑥ 補助関数層
├── b0x0_WebElement / c0x0_WebElements                  ← 要素（a 群の上に積む）
└── ChromeDriver.cls / EdgeDriver.cls / IWebDriver.cls  ← Selenium ライクな最終ファサード
```

3つの中で最もユニークで、**`a → b → c` / `1 → 6` という番号で依存の積み上げ順序そのものをファイル名に刻んでいます**。

一見すると読みづらい命名ですが、[前ページで見た排他的1本づけ](/vba-comparison/multi-tab)の設計と繋げると筋が通ります。どの瞬間を切り取っても生きている接続は1本だけなので、**そもそも同時に2つ目のタブが存在しません**。「ブラウザ」と「タブ」を分ける必然性が発生しないので、`ChromeDriver.cls` 1個が両方を兼務すれば足りる。

タブという自然な分割軸が使えないなら、代わりに使えるのは「その1本の接続を、どんな順番で処理しているか」という内部工程の軸だけです。`a1 → a6` の番号付けは、その帰結だったと読めます。仮に複数タブ同時保持に対応しようとすれば、この一式をタブの数だけ複製する必要が出て、結局「まとめて包む `Tab` クラスが欲しい」という話になるはずです。

### 抽象化はあるのに、共有されていない

このプロジェクトは `IWebDriver.cls` / `IWebElement.cls` という VBA の `Implements` 用インターフェースを持っています。ところが実装側を見ると。

| ファイル | 行数 |
| --- | --- |
| `ChromeDriver.cls` | 2,715 |
| `EdgeDriver.cls` | 2,718 |

差分を取ると **73 行、97% 以上が同一**でした。違うのは起動時のレジストリキー（`msedge.exe` / `chrome.exe`）とユーザーデータフォルダ名くらいです。インターフェースは切られているのに、共通実装を基底クラスに寄せる代わりに**ファイルごと複製**されています。Chrome 側だけを直したら Edge 側も同じ修正が要る、という状態です。

## VBAChromeDevProtocol：CDP 仕様そのままの 236 ファイル

```
src/
├── clsCDP.cls          ← 通信層 + ディスパッチ + 全ドメインオブジェクトの保持
├── AutomateBrowser.cls ← ブラウザ単位とタブ / セッション管理が同居
├── clsElement(s).cls   ← 要素単位
└── cdp/                ← 236 ファイル、合計 30,389 行（全て自動生成）
    ├── Accessibility.cls, Animation.cls, Audits.cls, Browser.cls,
    │   CSS.cls, DOM.cls, Network.cls, Page.cls, Target.cls ...
    └── Domain_TypeName.cls（各ドメインのサブタイプごとにも別ファイル）
```

スコープでの分割はほぼなく、**CDP のプロトコル定義そのものをファイル構成に写し取っています**。これは意図的な設計判断というより、入力の構造がそのまま出力に出た結果です。

## コード生成という別解

`src/generator/` に、この 236 ファイルを生み出した仕掛けが残っています。

- **`protocol.txt`（1.1MB）** ―― Chrome DevTools Protocol の公式スキーマ（`"version": {"major":"1","minor":"3"}`）そのもの
- **`convert.bas`（43KB）** ―― それを読んで .cls を吐き出す VBA 製ジェネレーター

生成されるコードは、想像よりずっと本格的です。

```vb
' Navigates current page to the given URL.
Public Function navigate( _
    ByVal url AS string, _
    Optional ByVal referrer AS Variant, _
    Optional ByVal transitionType AS Variant, _
    Optional ByVal frameId AS Variant, _
    Optional ByVal referrerPolicy AS Variant _
) AS Dictionary
    ' url: string URL to navigate the page to.
    ' referrer: string(optional) Referrer URL.
    ' transitionType: TransitionType(optional) Intended transition type.
    ' ...
    Dim params As New Dictionary
    params("url") = CStr(url)
    If Not IsMissing(referrer) Then params("referrer") = CStr(referrer)
    ' ...
    Set results = cdp.InvokeMethod("Page.navigate", params)
```

`cdp.Page.navigate url:="https://example.com"` と書けます。**メソッド名も引数名も本物の VBA シグネチャなので、タイプミスはコンパイルエラーになり、VBE の IntelliSense も効きます。** 仕様書の説明文はコメントとして埋め込まれ、CDP の enum 型も VBA の `Public Enum` として生成されています。

```vb
Public Enum AdFrameType
    AFT_none
    AFT_child
    AFT_root
End Enum
```

::: tip これは StarterWebScrapingKit が持っていないもの
`ExecuteCDP "Page.navigate", params` という文字列ベースの呼び方は、メソッド名を1文字間違えても実行するまで気づけません。[コアロジック比較の「型安全なプロトコル定義」](/core-comparison/gaps)で、Puppeteer / Playwright が `devtools-protocol` パッケージの型定義で解決していると書いた課題を、**VBA 圏で唯一実際に解いたのが VBAChromeDevProtocol** です。しかも公式仕様が更新されたらジェネレーターを再実行するだけで追従できます。
:::

### ただし、型が付いたのはコマンド方向だけ

生成された `Page.cls` を調べると、`Public Event` の宣言はひとつもありません。CDP の**コマンド（送信）とその型・enum は生成対象ですが、イベント（受信）は対象外**です。イベントは相変わらず `registerEventHandler "Page.downloadWillBegin", handler` という文字列指定のままで、[前ページで見た命名規約ベースのコールバック](/vba-comparison/events)に戻ります。

### そして、繋がっているのは 45 ドメイン中 10 個

もうひとつ。生成されたドメインクラスは 45 個ありますが、`clsCDP.Class_Initialize` が `cdp.〇〇` として実際に配線しているのは 10 個だけです。

```vb
Public Accessibility As cdpAccessibility
Public Browser       As cdpBrowser
Public CSS           As cdpCSS
Public DOM           As cdpDOM
Public SimulateInput As cdpInput
Public Network       As cdpNetwork
Public Overlay       As cdpOverlay
Public Page          As cdpPage
Public Runtime       As cdpRuntime
Public Target        As cdpTarget
```

残る 35 ドメイン（`WebAuthn`、`Fetch`、`Storage`、`Tracing` など）はファイルとしては存在するので、`New cdpWebAuthn` して `.init cdp` すれば使えます。ただし**そのひと手間は利用者側の仕事**で、`cdp.` と打ったときの IntelliSense には出てきません。

enum の扱いにも同じ「あと一歩」があります。`Public Enum` は生成されるものの、コマンドの引数側は多くが `CStr()` 渡しのままで、ジェネレーターのソースにも TODO が残っています。

生成という手段が効いた範囲と、その先の配線・仕上げが追いつかなかった範囲が、きれいに分かれています。

## 余談：JSON をどう読むか

VBA には JSON パーサーが標準搭載されていません。CDP は全てのやりとりが JSON なので、ここも各プロジェクトが自分で調達する必要がありました。選択は三者三様です。

| | 採用したもの | 方式 |
| --- | --- | --- |
| **StarterWebScrapingKit** | `BiDiCDPJson.cls`（UesleiDev 製、約 3,000 行） | トークンツリー + 遅延ノード。パース時に `Dictionary` を作らない |
| **VBAChromeDevProtocol** | `JsonConverter.bas`（VBA-JSON / Tim Hall） | フルパースして `Dictionary` / `Collection` を構築 |
| **vba-cdp-webdriver** | `htmlfile` COM に JScript を流し込む | `JSON.parse` をブラウザエンジンに委譲 |

3つ目が一番変わっています。`CreateObject("htmlfile")` で IE エンジンを起こし、そこへ JScript 関数を書き込んで `JSON.parse` を呼ぶ、という組み立てです。

```vb
Private Sub Class_Initialize()
    Set html = CreateObject("htmlfile")
    html.Write "<meta http-equiv='X-UA-Compatible' content='IE=edge' />"
```

```js
document.getValueFromObjectBySetKey = function(targetJson, key) {
    var vals = JSON.parse(targetJson);
    return vals[key];
}
```

自前でパーサーを書かずに済むのは賢い割り切りですが、代償があります。`GetValue(json, "result", "sessionId")` のようにパスを辿るとき、**中間の階層ごとに `JSON.parse` と `JSON.stringify` が走り直します**。CDP のように1回の往復で何度も値を取り出す用途では、この再パースが積み上がります。

一方 `BiDiCDPJson.cls` は「パース時にノードごとのオブジェクト割り当てをしない」ことを設計目標に掲げていて、[バッファ管理](/core-comparison/transport)と同じく**割り当てを減らす方向**に振られています。JSON という同じ問題に対しても、力の入れどころが違っていたことが見えます。

## 236 分割は「良くない仕組み」だったのか

自動生成そのものは、土台としてはむしろ合理的です。全コマンドを漏れなく網羅でき、仕様更新に追従でき、型が付く。Playwright も `page.context().newCDPSession()` という**生 CDP への抜け道**を用意していて、「便利メソッドで足りないときに生のドメインへ直接アクセスできる」最終手段はどのツールにも必要です。

問題は、その**上に何も積まれなかった**ことでした。

`clsElement.cls` の公開メンバーは8個です。

```vb
Public Property Let / Get value
Public Property Get InnerText
Public Property Get outerHTML
Public Property Get getAttribute
Public Sub setAttribute
Public Sub Click
Public Function Click_Download
```

StarterWebScrapingKit の `CDPElement.cls` は、`Public` な Sub / Function が 34 個に加えてプロパティが十数個。`onExist(timeout)` のようなタイムアウト付き待機、`getElementByXPath`、Shadow Root 対応、`getParent` / `getChildren` といった DOM 走査が揃っています。

### なぜ上乗せ層が要るのか ―― ドメインは必ず混ざる

`AutomateBrowser.cls` の `Click()` を見ると、その理由がわかります。「クリックする」という人間にとって1つの操作の中に、**4つの CDP ドメインが混在**しています。

```vb
Public Sub Click(...)
    Focus nodeId, backendNodeId                       ' ← DOM 系（可視範囲までスクロール / フォーカス）
    getNodeCenter x, y, nodeId:=nodeId, ...           ' ← DOM.getBoxModel 系（座標を取得）

    If strategy = Normal Or eager Then cdp.Page.enable      ' ← Page ドメイン（遷移検知用）
    If strategy = NetworkIdle Then cdp.Network.enable       ' ← Network ドメイン（通信の落ち着き検知用）

    cdp.SimulateInput.dispatchMouseEvent "mousePressed", x, y, ...   ' ← Input ドメイン（実際のクリック）
    cdp.SimulateInput.dispatchMouseEvent "mouseReleased", x, y, ...

    waitForPageToLoad strategy                        ' ← Page / Network / DOM イベントを跨いで待つ
End Sub
```

CDP のドメイン境界と、利用者がやりたいことの境界は一致しません。「ファイル操作なら DOM、クリックなら Input」と分けられるように見えて、実際には常に混ざります。だから **CDP のドメイン境界を一度分解して、人間のスクリプトの書き方に合わせて再構築する**という追加の設計判断が必要になる ―― `ElementHandle` や `CDPElement` は、その再構築の産物です。

VBAChromeDevProtocol にもこの役割の層（`AutomateBrowser.cls`）は一応あります。ただ `clsElement` が8メソッドで止まっているように、投資が薄いまま更新が途絶えました。結果として、236 ファイルの生ドメインは完成しているのに、その上の「人間が実際に使う層」が育たず、利用者が毎回 `cdp.DOM` / `cdp.Input` / `cdp.Page` を自分で組み合わせることになります。

**メッセージをどう捌くかは正解に辿り着けても、利用者にどう見せるかは誰かが意図的に設計しないと生まれません。** 自動生成は下ごしらえとして正しく、ただそこがゴールではなかった、というのが実態でした。

## まとめ

| | 生 CDP へのアクセス | 人間目線の抽象層 | 型安全 |
| --- | --- | --- | --- |
| **StarterWebScrapingKit** | `ExecuteCDP` に文字列で | Browser / Context / Element が厚い | なし（文字列 + `Dictionary`） |
| **VBAChromeDevProtocol** | 236 ファイルの型付きドメイン層 | 薄い（`clsElement` は8メソッド） | コマンド方向のみ、あり |
| **vba-cdp-webdriver** | 提供されていない | Selenium 風 API が厚い | なし |

3者とも一長一短で、**VBAChromeDevProtocol の生成層と StarterWebScrapingKit の抽象層が揃っていれば理想形だった**、というのが正直なところです。

## 関連

- [アーキテクチャ](/concepts/architecture) — このキットのクラス構成
- [クラス構成の比較](/core-comparison/classes) — Puppeteer / Playwright の3層モデルとの比較
- [残る差分](/core-comparison/gaps) — 型安全性を含む、実装投資で埋まる差
