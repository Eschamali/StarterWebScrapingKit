---
description: 起動時の引数が増えすぎて悲鳴を上げていた頃、SeleniumVBA の Capabilities クラスを覗いて気付いた「オプションオブジェクト」という答え。3度目のメジャーアップデートで入口周りを刷新した経緯と、WebSocket 機能を手放した話です。
---

# オプションオブジェクト、登場秘話

::: tip 登場秘話
〜起動の引数が増えすぎて、他の人の設計を覗いてみたら、同じ痛みの答えが落ちていた〜
:::

v1.0.0 を公開したのは 2025年12月20日。そこから 295 日、ついに **3度目のメジャーアップデート（v4.0.0）** を迎えました。なんやかんやで飽きずにここまで来れたのは、AI と一緒に開発してこれたおかげです🥹

今回のアップデートは、このツールの「入口」――ブラウザを**起動する**ところと、**再接続する**ところ――を、がっつり作り直したものです。ここでは、その発端と、一緒に手放したものの話を振り返ります。

---

## 増えすぎた「経路」と、増え続ける引数

Web スクレイピング機能そのものは、機能数をある程度抑えてあります。ところが、その土台にある**制御経路のパターン**は、気付けば大変なことになっていました🫠

- CDP
  - Pipe
  - WebSocket
  - COM（WebView2 の `CallDevToolsProtocolMethodForSession`）
- WebDriver BiDi
  - Pipe
  - WebSocket
- Edge / Chrome
  - 安定版
  - Beta
  - Dev
  - Canary

これまでは、「設定シート」という便利なセルの集まりに頼って、なんとか耐えてきました。ところが、起動時の設定は増える一方です。ユーザーデータフォルダ、追加の起動引数、初期 URL、ウィンドウの表示状態、BiDi の `capabilities`……。起動のたびに、あっちのセルを読み、こっちの引数を渡し、というコードがどんどんごちゃごちゃになっていきました。

そしてある日、さすがに嫌気がさしました🫠🫠🫠

---

## 他人のコードを覗いてみた

ふと、同じ「VBA でブラウザを動かす」ことをしている先人たちのコードを覗いてみました。

- [TinySeleniumVBA の `Capabilities.cls`](https://github.com/uezo/TinySeleniumVBA/blob/main/Capabilities.cls)
- [SeleniumVBA の `WebCapabilities.cls`](https://github.com/GCuser99/SeleniumVBA/blob/main/src/VBA/WebCapabilities.cls)

どちらも、複雑な設定は、**設定の保持だけを担当する専用クラス 1 つ**にまとまっていました。

> 「ブラウザのオプション多すぎて、引数で渡してられんわ！　Capabilities っていう専用クラス作ろ！！」

きっと、個人開発者のお二人も、そして Microsoft が開発している WebView2 も、まったく同じ痛みを経て、まったく同じ答えにたどり着いたのだと思います。これが、**オプションオブジェクト**という考え方でした。

私もようやく、そこにたどり着くことができました😳

---

## 3 日間の大工事

2026年10月8日から10日にかけて、この考え方を一気に取り込みました。

```text
2026-10-08 `SeleniumVBA`の名称を借りたリネーム処理
2026-10-08 `reattach`1つのメソッドに集約
2026-10-08 `ReattachOptions`による新フローに変更
2026-10-10 初期起動も設定Classオブジェクトを渡す方式に刷新
2026-10-10 `BrowserCapabilities`に関する説明追加
```

起動時の設定は `BrowserCapabilities`、再接続時の設定は `ReattachOptions`。どちらも設定を保持するだけのクラスで、起動／再接続のメソッドに渡します。メソッドの名前も、SeleniumVBA の名称を借りて、見直しました。

起動は、こんなに短くなりました。設定の既定値はワークシートから読み込まれるので、シートの内容のままでよければこれだけです。

```vb
Dim t As CDPContext
Set t = New CDPContext
t.OpenBrowserAndConnectTab BrowserCapabilities
```

初期 URL を渡したければ、起動の前に 1 行足すだけです。

```vb
BrowserCapabilities.Set_StartURL = "https://example.com"
```

どうです？ [SeleniumVBA の起動の儀式](https://note.com/sele_chan/n/n61a67c4379de#42fc70b1-7e0b-489f-9317-b47c9255d1e2)と、似ているのが分かると思います。

### 再接続は「1 つのメソッド」に

もうひとつ、ちょっと感動したことがあります。

ご先祖様である[本家ライブラリ](https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA/blob/main/Class%20Modules/CDPBrowser.cls)の `CDPBrowser` にも、再接続のためのメソッド **`reattach` がありました。しかも引数は、識別名（プロファイル名）だけです。

```vb
Public Function reattach(Optional userProfile As String = "CDP1") As Boolean
```

その後、このツールは Port（WebSocket）や WebView2 といった制御経路を増やしていきました。経路ごとに、再接続のやり方が違います。そこで 2026年8月25日、「責務分離化」と称して、`reattach` を経路ごとに分けました。

```text
2026-08-25 `reattach`の責務分離化
```

Pipe・WebSocket・WebView2 ごとに専用の `reattach` 系メソッドがあり、ブラウザ用とタブ用で、さらに倍になりました。経路が増えるたびに、メソッドも増えていきます🫠

そして 10 月、オプションオブジェクトを取り入れたことで、その答えが見えました。「どの経路で再接続するか」は、メソッドの数ではなく、**渡すオブジェクトの中身**で表せばよかったのです。

```text
2026-10-08 `reattach`1つのメソッドに集約
2026-10-08 `UserName`を渡して通信経路を設定するフローに変更
```

```vb
ReattachOptions.Set_PipeMode = UserName    ' 経路と識別名を指定
b.reattach ReattachOptions
```

気付けば、メソッドは再び **`reattach` ただ 1 つ**。識別名を渡して、つなぎ直すだけ。経路を増やすという遠回りをして、本家が最初から持っていた、あのシンプルな形に、自然と戻ってきていました。

何度も作り直して、結局いちばん素直な形に辿り着く。**先人の設計は、やっぱり偉大でした**🥹

### 使い切りの落とし穴

もうひとつ、このオブジェクトには「使い切り」という設計があります。使い終わったら `Nothing` にして、次の起動に設定が持ち越されないようにするものです。ところが、これが思いのほか厄介でした。

```text
2026-10-10 不意なタイミングで`BrowserCapabilities - Class_Terminate`が走る問題を修正
2026-10-10 簡易起動版にて、`Application.OnTime`が絡むと、うまく機能しない問題を修正
```

VBA のクラスには、名前そのものがそのまま使える「既定インスタンス」があります。裸の `BrowserCapabilities` と書けるのは、そのおかげです。この既定インスタンスの破棄と、参照カウントの扱いを正しく揃えるのに、ちょっと苦労しました🫠

---

## 純WebSocket機能を、手放す

今回のメジャーアップデートを機に、もうひとつ決断をしました。Microsoft が提供している `WinHttp.dll` の WebSocket 機能を使った、**純 WebSocket 機能の廃止**です。

```text
2026-10-10 純WebSocket機能を廃止
```

理由は、お察しの通りです。

- ほぼデモサイトでしか機能しない
  - 実用性を考えると、認証系も考慮しないと意味がない。できたとしても、複雑なパラメーター要求やロジックを強いられる
- ハンドルの儀式が多すぎて、コードが複雑になる

ほぼ POC 止まりの状態で、これ以上進化する兆しもありませんでした。

さらに言えば、WebSocket 機能に特化したライブラリは、もう世の中に出回っています。

- [wasabi](https://github.com/vbacollective/wasabi)

もし WebSocket 機能をお探しなら、こちらをおすすめします。

なお、ここでいう「純 WebSocket 機能」は、`WinHttp.dll` ベースの汎用的な WebSocket 機能のことです。このツールの Port 経路（`--remote-debugging-port` による CDP 制御）は、[別の経緯](/stories/birth-story)で自作した WinSock ベースの仕組みで動いており、今回の廃止とは関係なく、そのまま残っています。

---

## 結論 — 引き算で見えてきたもの

引数を 1 つのオブジェクトにまとめ、メソッドを 1 つに集約し、使い道の薄れた機能を手放しました。足した機能は、ほぼありません。それなのに、入口がずいぶんと見通しよくなりました。

::: tip フィナーレ
ライブラリの設計は、足し算ばかりしていると、気付かないうちに迷路になります。他の人の設計を覗いてみる、というだけのことが、一番の近道でした。先人たちのコードに、改めて感謝します🙏
:::

::: details 開発者の独り言

Ver1.0.0 の時は、ネットの世界へ向かう 3 パターンを載せたようなものでした。http（REST WebAPI）、CDP、WebSocket……。

そして作っていくうちに、気付いたのです。

あれ、CDP 特化系のライブラリってないよな🤔

バージョン履歴を見ればわかる通り、ほぼ `ChromeDevtoolsProtocol` です🫠

そして、時の流れが過ぎると、[http](https://github.com/harumiWeb/VBA-HTTP) / [WebSocket](https://github.com/vbacollective/wasabi) に特化したライブラリも登場し、やっと気付きました。

そうか、わい `ChromeDevtoolsProtocol` 特化を作ってたのか😳

WebSocket は退部させ、http 通信はおまけに降格。

気付いたら、あの頃の IE の 1 行

`CreateObject("InternetExplorer.Application")`

を、丸ごと作っていたんだな〜と。あの 1 行には、あんな処理が詰まっていて、それを VBA コードとしてエミュレートできているんだな〜と、思いました。

:::

## 次に読む

- [`BrowserCapabilities`](/api/cdp/BrowserCapabilities) — 起動時のオプションオブジェクト
- [`ReattachOptions`](/api/cdp/ReattachOptions) — 再接続時のオプションオブジェクト
- [再接続 (reattach)](/guides/reattach)
- [このツールの誕生秘話](/stories/birth-story) — WinHTTP を諦めて WinSock を自作した話
- [WebView2、登場秘話](/stories/webview2-story)
