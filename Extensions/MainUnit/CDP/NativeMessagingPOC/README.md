# chrome.debugger × Native Messaging CDP制御 (PoC・技術検証記録)

> [!NOTE]
> これは、Chrome拡張機能の`chrome.debugger` API × Native Messagingを使って、既存の`StarterWebScrapingKit`(Pipe / WebSocket / WebView2)に**4本目のCDP transport**を追加できるか検証したPoC(Proof of Concept)です。\
> **メインの自動化フローとしては使用していません**。後述の通り実用上のメリットは無く、**純粋な技術的好奇心による検証**として動作確認まで完走した記録です。

***

## 🎯 これは何？（＆ 後日判明した前提の崩壊）

このPoCを始めた当初のモチベーションは、以下でした。

* **ユーザーが普段使っている通常のブラウザ・既存タブ**を、`--remote-debugging-port`等を一切使わずに制御できないか？
* Chrome拡張機能の`chrome.debugger` APIなら、既存タブへの明示的アタッチだけでCDP制御できるはず
* 拡張機能↔ローカルアプリ間の通信は、ブラウザ標準の**Native Messaging**を使う
* そのNative Messagingの「ネイティブホスト」として、**Excel自身**を直接登録できないか？

> [!WARNING]
> しかし、このモチベーション自体が**そもそも成立していませんでした**。
>
> `StarterWebScrapingKit`には既に、`CDPCoreViaWebSocket.AutoConnectDevToolsActivePort`(`Demo_CDP.bas`の`AutoConnectDevToolsActivePort`)という機能があり、`edge://inspect/#remote-debugging`から「このブラウザインスタンスのリモートデバッグを許可」するだけで、**「普段使いのブラウザの、今開いてるタブ」をWebSocket transport経由で全権限(chrome.debuggerのようなブラウザスコープ制限なし)制御できます**。
>
> つまり、このPoCが目指していたゴールは、既にこのプロジェクト自身がもっとシンプルな形で持っていました。参考: [Chrome DevTools MCPのセットアップ記事](https://dev.classmethod.jp/articles/chrome-devtools-mcp-setup/)（同じ`edge://inspect/#remote-debugging`の仕組みに言及）
>
> したがって、このPoCに**実用上のメリットは無く**、以下は「chrome.debugger×Native Messagingという組み合わせが技術的に実現可能か」を検証した記録として残しています。

***

## 🧩 全体構成

```
┌─────────────────────────┐   chrome.debugger    ┌──────────────┐
│  対象タブ (Chrome/Edge)   │ ◄──────────────────► │  拡張機能        │
└─────────────────────────┘                       │ (background.js) │
                                                    └────────┬─────────┘
                                                             │ Native Messaging
                                                             │ (stdin/stdout,
                                                             │  4byte長 + UTF-8 JSON)
                                                             ▼
                                                    ┌──────────────────────────┐
                                                    │ Excel (VBA)               │
                                                    │  = ブラウザに起動された    │
                                                    │    ネイティブホスト本体    │
                                                    │                            │
                                                    │ CDPCoreViaNativeMessaging  │
                                                    │  → CDPCore → CDPBrowser   │
                                                    │  → CDPContext             │
                                                    └──────────────────────────┘
```

Chromeの`Native Messaging`は仕様上、拡張機能が`chrome.runtime.connectNative`した瞬間に、OS側(レジストリ)に登録されたホスト実行ファイルをブラウザ自身が起動し、そのプロセスの標準入出力を直接パイプで掴みます。このPoCでは、そのホスト実行ファイルとして**Excel自身**(`EXCEL.EXE`直接、または`/x`起動するラッパー`.bat`経由)を登録しています。

***

## 📁 ファイル構成

```
NativeMessagingPOC/
├── README.md                  ← このファイル
├── 手順書.md                   ← セットアップ手順(拡張機能読み込み〜レジストリ登録〜動作確認)
└── Extension/
    ├── manifest.json           ← 拡張機能マニフェスト(MV3)
    └── background.js           ← chrome.debugger中継ロジック本体
└── Demo/
├── NativeMessagingSetup.bas       ← ホストマニフェスト(.json)/ラッパーBAT生成ヘルパー
└── Demo_NativeMessaging.bas       ← セットアップ用マクロ + XLStart側の待受ループ

src/classes/
├── CDPCoreViaNativeMessaging.cls  ← 標準入出力(stdio)経由のtransport本体
├── CDPCore.cls                    ← 既存のPipe/WebSocket/WebView2と同じ枠組みに結線
└── CDPBrowser.cls                 ← `reattachNativeMessaging`を追加(既存の`reattachWebSocket`等と対称)
```

セットアップの詳細手順は **[手順書.md](./手順書.md)** を参照してください。

***

## ⚙️ VBA側の設計方針

既存の3 transport(Pipe/WebSocket/WebView2)と同じ`CDPCore.cls`の枠組みに、4本目として素直に結線しています。

| メンバー | 役割 |
|---|---|
| `CDPCoreViaNativeMessaging.cls` | `CDPCoreViaWebSocket.cls`と対称設計。標準入出力(`GetStdHandle`/`ReadFile`/`WriteFile`/`PeekNamedPipe`)を直叩きし、4バイト長プレフィックス+UTF-8 JSONのフレーミングを1呼び出し=1回のノンブロッキング読み取りで処理する |
| `CDPCore.cls` | `Property Set NativeMessagingMode`で切り替え。`viaWebSocket_WebSocketReceive`と同じパターンで、生バイトを`CDPPipeStream`に蓄積→既存の`ReadyCDPMessage`でテキスト化(**WebView2のような経路バイパスはしない**) |
| `CDPBrowser.reattachNativeMessaging` | `reattachWebSocket`/`reattachWebView2`と対称。接続確認を兼ねて`Browser.getVersion`を1回実行する |

**ポイント**: `CDPCore.cls`/`CDPBrowser.cls`より上のレイヤーは、Pipe/WebSocket/WebView2の時と**全く同じAPI**(`CDPBrowser`/`CDPContext`の`ExecuteCDP`/`navigate`/`jsEval`等)がそのまま使えます。差異の吸収は、すべて`CDPCoreViaNativeMessaging.cls`と、後述の拡張機能側で行っています。

***

## 🌐 拡張機能側(`background.js`)の設計方針

`chrome.debugger`でタブにアタッチしたセッションは、**ブラウザ全体スコープのコマンドを受け付けない**(`-32000 Not allowed` / `-32601 Method not found`)という制約があります。これは`chrome.debugger`の仕様上の制約で、Pipe/WebSocket/WebView2版では起こらない、Native Messaging固有の問題です。

これに対処するため、`background.js`は以下のようなコマンドを**chrome.debuggerに中継せず、ローカルで応答を合成**しています。

| コマンド | 対処内容 |
|---|---|
| `Browser.getVersion` | `navigator.userAgent`からバージョン文字列を合成(取れなければダミー値で成功扱い) |
| `Target.getTargets` | 自動アタッチ(`Target.attachedToTarget`)イベントから蓄積した`targetInfo`キャッシュを返す(※このタブ配下のみ。ブラウザ全体の一覧ではない) |
| `Target.attachToTarget` | 自動アタッチ済みの`targetId → sessionId`対応表から解決して返す(ルートタブ自身には仮想`sessionId`(`"root:" + tabId`)を払い出す) |
| `Target.detachFromTarget` | ルートの仮想`sessionId`宛てはローカルで成功扱いにする |

また、`chrome.debugger.sendCommand`の`lastError.message`が実はCDPエラーをJSON文字列化したものである点に対応し、二重ネストせず`{code, message}`に正規化しています。

> [!NOTE]
> `Target.setDiscoverTargets`も同様に`Not allowed`になりますが、こちらはVBA側(`CDPBrowser.EnableDiscoverTargets`)が元々失敗を許容する設計のため、対応不要でした。
>
> このように、**Not allowedになるコマンドは今回網羅した以外にも存在し得ます**。本格運用する気は無いため、以降は気が向いた時にのみ、同じパターン(自動アタッチイベント等から拾える情報でローカル合成する)で対処する想定です。

***

## ✅ 動作確認できたこと

実際に、以下の一連のフローがエンドツーエンドで動作することを確認済みです(`.cdp.log`のログより)。

1. Excel(NativeMessagingホストとして起動) ⇔ 拡張機能の接続確立
2. `Browser.getVersion`（ローカル合成）
3. `Target.getTargets` → `Target.attachToTarget`（ルートタブのセッション獲得）
4. `Runtime.evaluate`でページ情報取得
5. `Page.navigate`で実際に別ページへ遷移
6. `document.readyState`ポーリングで`complete`まで待機
7. ページ遷移に伴う子ターゲット(Service Worker)の`Target.attachedToTarget`/`Target.detachedFromTarget`追従
8. `CDPContext`破棄時の`Target.detachFromTarget`（ローカル応答）

つまり、`CDPBrowser`/`CDPContext`の一般的なAPI(`navigate`、`jsEval`、`wait`等)が、Native Messaging transport経由でも他のtransportと同じように使えることを確認しています。

***

## ⚠️ 既知の制限事項

* **`chrome.debugger`のブラウザスコープ制限**: 前述の通り、`Browser.*`/一部の`Target.*`コマンドは拡張機能側での合成が必要。未対応のコマンドに遭遇したら都度対応
* **ホストプロセスの起動リスク**: Excel自身をNative Messagingホストとして登録する構成には、Chromeが渡す不正なコマンドライン引数をExcelがファイルパスと誤解釈する可能性、およびExcelの既定DDEシングルインスタンス動作により新規プロセスが即終了する可能性がある(詳細は手順書§0/§3参照)
* **デバッグ中インフォバー**: `chrome.debugger.attach`により、ブラウザ標準の「デバッグ中」インフォバーが必ず表示される(非表示不可)
* **1メッセージ最大1MiB**: Native Messagingのホスト→ブラウザ方向の上限。大きな結果を返すCDPコマンド(`Page.captureScreenshot`等)で超過する可能性あり
* **セッションごとに新規Excelプロセス**: デバッガセッション(タブ)ごとに新規Excelプロセスが起動する構成のため、複数タブの同時制御は、その分だけExcelプロセスが並行起動する
* **`Target.getTargets`のスコープ**: このタブ配下(自動アタッチ済みの子ターゲット)のみ。ブラウザ全体の他タブ一覧は取得不可

***

## 🤔 なぜメイン自動化フローに採用しないのか

* 冒頭の通り、このPoCが唯一のセールスポイントとしていた「普段使いのブラウザの、今開いてるタブをそのまま制御できる」は、**`edge://inspect/#remote-debugging` + 既存の`CDPCoreViaWebSocket.AutoConnectDevToolsActivePort`で、より簡単かつブラウザスコープ制限なしに既に実現できていた**ため、そもそも存在意義が無い
* その上さらに、上記の制限事項(ブラウザスコープ制限・ホスト起動リスク・セッションごとの新規Excelプロセス)により、Pipe/WebSocket/WebView2はもちろん、同じゴールを実現する`AutoConnectDevToolsActivePort`と比べても、挙動の予測可能性・安定性・導入コストのすべてで劣る
* つまり実用上のメリットは無く、「chrome.debugger×Native Messagingという組み合わせを、VBAだけでどこまで動かせるか」という技術的好奇心のみが、このPoCを作った理由であり、存在価値です

### 🕳️ 強いて言えば意味が残る、唯一の例外ケース

`edge://inspect/#remote-debugging`は「リモートデバッグ用のポートを開く」操作そのものです。つまり`AutoConnectDevToolsActivePort`(WebSocket transport全般)は、**ポートを開ける権限がある**ことが大前提になります。

一方`chrome.debugger`は、TCPポートを一切開きません(`--remote-debugging-port`も`edge://inspect`のトグルも不要)。拡張機能APIとNative Messaging(OSのstdioパイプ)だけで完結します。

したがって、

* 組織のポリシー等で**リモートデバッグ用ポートの開放自体が禁止**されている
* それでも**どうしても今目の前のブラウザ(今開いてるタブ)を制御したい**

という、かなり限定的な条件が両方揃った場合に限り、このPoCの構成が唯一の選択肢になり得ます。とはいえ、そのような環境では「拡張機能をデベロッパーモードで読み込む」こと自体もポリシーで制限されている可能性が高く、実際に役立つ場面はさらに狭いと思われます。
