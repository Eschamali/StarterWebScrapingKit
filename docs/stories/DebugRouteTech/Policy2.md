---
description: Power Automate Desktop の「Web ページで JavaScript 関数を実行」が Edge でだけ動かなくなった原因を追った記録です。たどり着いたのは、拡張機能のポリシー ExtensionSettings と chrome.debugger をめぐる Chrome 155 の変更でした。Chrome と Edge の違い、回避の方法、2026年の CDP まわりのポリシーの動きをまとめます。
---

# Power Automate Desktop のブラウザ自動化が動かない原因を追ったら、今の CDP 関連のポリシー動向が分かった話

<PromoMovie src="/promo/ReiwaIE_Policy2.html" title="令和のIE 技術秘話 ― ポリシーの壁 #2 1分の沈黙" />

[ポリシーの壁](/stories/DebugRouteTech/Policy)の続きです。

今回の主役はこのツールではなく、Power Automate Desktop（以下 PAD）です。ただ、原因をたどっていくと、最後は CDP とポリシーの話に行き着きました。ポリシー続きということで、ここに残しておきます。

## はじめに

普段、PAD でブラウザの自動処理をしているユーザーから、ある報告を受けました。

* なんか突然、「Web ページで JavaScript 関数を実行」のところだけ動作しないんだが...

その一報をきっかけに、次のような声が一気に流れました。

* うちも！ 再インストールしてもだめ

* いったん古いバージョンの PAD にしても、だめだった

* 新しい PAD が来たけど、直ってない

さらに奇妙なことに、途中からこんなやりとりも出てきました。

```text
A ふともう一回、「Web ページで JavaScript 関数を実行」したら直った！？
B あれ、でもうちはまだだめだな...
A あ、Chrome で確認しました！
B ほんとうだ、Edge から Chrome に変えただけで動いたわ
```

調べても調べても、PAD まわりにそういった不具合の情報は見つかりませんでした。

CDP まわりに詳しい私は思いました。

「詳しい原因は調査中...とな？ じゃ～、どっちが早く見つかるか勝負するか🤠」

## まずは再現

まずは、同じ現象が出る PC で、PAD の「Web ページで JavaScript 関数を実行」を実際に動かしてみました。

確かに動きません。いや、動くには動くのですが、挙動が変です。

* 1分ほど、な～んにも反応がない

* PAD 側では成功の扱い。でも結果は、どんな JavaScript でも空欄（`function ExecuteScript() {return document.title;}` ですら空欄）

なんとも変な挙動です。

## CDP を直接使うと？

うちの自慢の CDP ツールで、同じことを試しました。`edge://inspect/#remote-debugging` 経由の CDP でつなぎ、`Runtime.evaluate` を実行しました。

こちらは動きました🤠

つまり、CDP そのものが塞がれているわけではないようです。

## Copilot と原因究明

CDP には詳しくても、PAD そのものはチンプンカンプンです。そこで、Copilot と一緒に原因を調べることにしました。

### ポリシーの確認

はじめに Copilot は、ポリシーを確認するように言ってきました。

幸い、`edge://policy/` は塞がれていません。これを開けば、どのポリシーが入っているかが一目瞭然です。

[前回](/stories/DebugRouteTech/Policy)調べた、CDP に直接かかわるポリシー（`RemoteDebuggingAllowed`、`DeveloperToolsAvailability`、`UserDataDir`）は見当たりませんでした。しかし Copilot は、ちゃんと怪しい候補を見つけてくれました。

### ExtensionSettings

[ExtensionSettings](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/extensionsettings) は、拡張機能をまとめて管理するためのポリシーです。値は JSON で、拡張機能の ID ごとに「入れてよいか」「どの権限を止めるか」「どのサイトに触らせないか」といった設定を書きます。ID の代わりに `*` と書くと、個別の設定を持たない拡張機能すべてに効く、既定の設定になります。

今回かかわるのは、その中の `runtime_blocked_hosts` です。[Microsoft のガイド](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-manage-extensions-ref-guide)によると、ここに書いたサイトでは、拡張機能がページとやり取りしたり、ページを書き換えたりできなくなります。止まるものとして挙がっているのは、JavaScript の挿入、Cookie へのアクセス、Web リクエストの書き換えです。

たとえば次のように書くと、「どの拡張機能も `example.com` には手を出すな」という意味になります。

```json
{"*":{"runtime_blocked_hosts":["*://*.example.com"]}}
```

特定のサイトだけは拡張機能に触らせたくない、というときに使う設定です。逆に、禁止したサイトの中から一部だけ許す `runtime_allowed_hosts` もあります。

CDP を直接使う私には関係のないポリシーです。しかし、拡張機能を経由してブラウザを操る PAD には、まともに効いてきます。

### でも...

しかし、PAD で操作しているサイトは、`runtime_blocked_hosts` に書かれた URL に引っかかっていません。さらに言うと、ほかのブラウザ操作は動いています（確かめたのは1つだけですが、新しいタブを開いて URL へ移動する操作はできました）。

PAD のブラウザ操作が全部できなくなったわけではありません。とにかく、「Web ページで JavaScript 関数を実行」だけが変なのです。

### 原因判明？

次に Copilot から来た提案は、「『ビューの検査: サービス ワーカー』に、それっぽいログはない？」でした。

`edge://extensions/?id=kagpabjoboikccfdghpdlaaopmgpgfdc` を開いて「サービス ワーカー」をクリックすると、PAD の拡張機能そのもののログを見られます。

すると...！

```log
Unchecked runtime.lastError: Host access is restricted by policy.
Unchecked runtime.lastError: Debugger is not attached to the tab with id: 1123621967.
Error handling response: TypeError: Cannot read properties of undefined (reading 'result')
    at chrome-extension://kagpabjoboikccfdghpdlaaopmgpgfdc/background.js:403:70
Unchecked runtime.lastError: Debugger is not attached to the tab with id: 1123621967.
```

特に1行目は、明らかに「ポリシーによる拒否」と読めます。

原因が見えてきた...と思ったのですが、違和感があります。

#### 矛盾

たとえば、ExtensionSettings が次の値だったとします。

```json
{"*":{"runtime_blocked_hosts":["https://XXX.com"]}}
```

`https://XXX.com` で PAD の「Web ページで JavaScript 関数を実行」をすると、当然、さっきの現象（1分の沈黙）になります🤠

では、`https://www.google.com` でやったらどうなるか。なんとこれも、同じく1分の沈黙なのです🙃

`runtime_blocked_hosts` に載っていないのに、なぜか実行できません。これでは、まだ納得できません。

### 最小構成で再現

「Edge のポリシーよりさらに強いポリシーがあるのか？（会社独自のポリシーによる上書きなど）」という疑いも浮かびました。しかし、その領域になると管理者権限が要るので、そこで詰んで迷宮入りになってしまいます😵

「なんとか、まだ試せるルートはないか...」と必死に考えました。

Copilot といろいろ議論を重ねるうちに、こういう回答をもらえました。

> `chrome.debugger.attach` あたりが怪しそうだね

そして、ひらめきました😳

「これは PAD 特有の問題なのか？ それとも、最小構成の自作の拡張機能でも起こるのか？」

とはいえ、検証用の拡張機能を作る気力までは出ませんでした...が、ここで Copilot の画面に、次の質問の候補が出ていました🫨

> 自作拡張で Debugger attach を試す手順を整理して

もしやと思って押しました。すると、なんということでしょう🥺

* manifest.json

* background.js

この2つを一瞬で用意してくれました。そして手順どおりにやると...

```log
Unchecked runtime.lastError: Host access is restricted by policy.
```

で、出た🫨 PAD じゃなくても出た🫨

### そしてついに...

これを Copilot に報告したら、ついに真の原因が返ってきました。

[Chrome 155 で chrome.debugger のエンタープライズ ポリシーの適用を強化](https://developer.chrome.com/blog/debugger-enterprise-policy-restrictions?hl=ja)

めちゃくちゃ簡単に言うと、こうです。

* `runtime_blocked_hosts` に1つでも登録があったら、その拡張機能は、どのサイトに対しても `chrome.debugger`（拡張機能から使う CDP）を使えない

もう少し正確に書きます。発表によると、変わったのは Chrome 155 からです。ExtensionSettings で、ある拡張機能に空でない `runtime_blocked_hosts` が設定されていると、その拡張機能の `chrome.debugger.attach()` はすべての対象で拒否されます。そのときのエラーが、あの `Host access is restricted by policy.` です。`runtime_allowed_hosts` で個別にサイトを許可していても、結果は変わりません。発表はこれを「全か無か」のモデルと呼んでいます。

それまでのルールは、アタッチしようとしているページの URL が禁止リストに載っているかどうかを見るだけでした。`XXX.com` だけを禁止していれば、`www.google.com` では動きます。新しいルールでは、禁止リストに何か1つでも書いてあれば、全サイトで拒否されます。

さっきの矛盾は、これで説明がつきます。

なお、スクリーンショットをポリシーで禁止している場合（`DisableScreenshots` や、データ損失防止（DLP）のルール）も、同じように `chrome.debugger.attach()` が拒否されます。こちらのエラーは `Screenshot capture is restricted by policy.` です。

一方で、ポリシーで管理されていないブラウザや、これらの設定を入れていない環境には影響がありません。こうした設定が入っていない PC なら、同じ PAD でも影響を受けないことになります。

Chromium のソースと変更の履歴を追うと、時系列は次のとおりです。

| 日付         | できごと                                                                                        |
| ---------- | ------------------------------------------------------------------------------------------- |
| 2026年8月20日 | Chromium のソースに、この制限が入る                                                                      |
| 9月8日       | Chrome の発表。同じ日に、Chrome 154 には入れないよう、いったん取り下げ（企業が準備する時間を取るため、とコミットにあります）                     |
| 9月11日      | Chrome 155 向けに入れ直し                                                                          |
| 9月16日      | 制限を止めるためのフラグ `ExtensionDebuggerStrictPolicyRestrictions` を追加（この時点の既定は「有効」）。Chrome 155 のベータ版 |
| 10月6日      | Chrome 155 の安定版                                                                             |
| 10月8日      | Edge 155 の安定版                                                                               |
| 11月3日      | Chromium のソースで、フラグの既定を「無効」に変更                                                               |
| 11月4日      | Chrome の発表に、ロールバックの追記                                                                       |

表の最後の2行のとおり、今の Chrome では、このルールはいったん撤回されています。いざ実装されると、「厳しすぎる」という声が届いたようです。発表の追記には、お客様からのフィードバックを受けて、より柔軟な方式を設計する間はロールバックする、とあります。

撤回のされ方は、バージョンによって2通りあります。

* Chrome 158（158.0.8127.24 以降）と Chrome 159：最初から無効

* それより前の Chrome 155〜158：Google がブラウザに配る設定の更新で無効にする。Chrome が自動で受け取り、再起動のあとに効く

少し前に挙げたやりとりと、辻褄が合います🤠

```text
A ふともう一回、「Web ページで JavaScript 関数を実行」したら直った！？　← Chrome がちょうど旧ルールに戻した時期
B あれ、でもうち（の Edge）はまだだめだな...
A あ、Chrome で確認しました！
```

### じゃあ Edge は...

では、Edge はどうなのか。

なんと現時点（2026年11月12日）では、Chrome と違って、撤回するという話が見当たりません。

* Edge の [ExtensionSettings の説明](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/extensionsettings)には、バージョン 154 以降、`debugger` の権限を持つ拡張機能に `runtime_blocked_hosts` を適用すると、すべての対象で `chrome.debugger.attach()` が完全に無効になる、と書かれたままです

* Edge の[リリースノート](https://learn.microsoft.com/en-us/deployedge/microsoft-edge-relnote-stable-channel)は、安定版の 153〜157 と、ベータ版の 158 を確かめました。この制限に触れた記述も、撤回に触れた記述もありませんでした

Chrome の発表は 155 から、Edge の説明は 154 からで、バージョンが1つずれています。先ほどの表のとおり、Chromium ではもともと 154 に入るはずだったものを 155 に延ばしています。Edge 154 の時点で実際に効いていたかどうかは、確かめていません。

つまり、今はこういう差があります。

| ブラウザ   | ExtensionSettings と chrome.debugger |
| ------ | ----------------------------------- |
| Chrome | 旧ルール（禁止したサイトだけ拒否）                   |
| Edge   | 新ルール（1つでも禁止していれば全サイトで拒否）            |

この差が分かったことで、次の謎も一気に解けました。

```text
B ほんとうだ、Edge から Chrome に変えただけで動いたわ
```

PAD のバージョン、拡張機能のバージョン、OS のバージョン、[SeleniumBasic のときにあった Windows の更新プログラム](/stories/DevHistory/gettcptable-story)、会社独自のポリシー...

どれも違いました。

原因は、ブラウザのアップデートで、ポリシーまわりの判定のしかたが変わったことでした。そのせいで、PAD の「Web ページで JavaScript 関数を実行」が動かなくなっていたのです。

自作の拡張機能という最小構成のテスト結果を Copilot に報告したことで、Copilot もその範囲に絞って考えてくれました。そのおかげで、[この発表](https://developer.chrome.com/blog/debugger-enterprise-policy-restrictions?hl=ja)にたどり着けたのだと思います🥺

#### Edge でも動かすには？

現時点では、次のように起動し直すと、Edge でも「Web ページで JavaScript 関数を実行」が動きました。

```text
msedge.exe --disable-features=ExtensionDebuggerStrictPolicyRestrictions
```

Chrome と並べると、次のとおりです。

| ブラウザ   | 旧ルールにするには                                                                   |
| ------ | --------------------------------------------------------------------------- |
| Chrome | 何もしなくてよい                                                                    |
| Edge   | `--disable-features=ExtensionDebuggerStrictPolicyRestrictions` を明示的に付けて起動する |

Edge にも旧ルールに戻す手段はあり、それが既定ではないだけ、ということのようです。

いくつか補足します。

* このフラグは、Chrome の発表にも載っている回避策です。Chrome でも、Google からの設定の更新が届かない環境（`ChromeVariations` ポリシーなどで、すべてのバリエーションを無効にしている環境）では、このフラグを付けて起動するよう案内されています。「Chrome なら何もしなくてよい」には、この例外があります

* Chromium のソースでは、このフラグの説明に「無効にするとセキュリティ上の影響がある」という注意書きがあります。会社の PC で使うなら、先に管理者へ相談したほうがよいと思います

* 発表にある管理者向けの案内は、「`chrome.debugger` が必要な拡張機能には、禁止サイトのリストを設定しないこと」です。ExtensionSettings をどう書けば PAD の拡張機能だけを外せるかは、試していません

* Chrome の撤回のうち 155〜158 の分は、ソースの書き換えではなく、Google が配る設定の更新です。Edge で制限が効いたままなのは、この更新が Edge には届かないからではないか、と見ています。これは推測です

* Chromium のソースでは、158 から既定が「無効」です。Edge 158 で変わるかもしれませんが、出てみないと分かりません

## 答え合わせ: 拡張機能の中で起きていたこと

原因が分かったところで、PAD の拡張機能（バージョン 2.73.0.36）のソースを読んでみました。すると、最初に見たログの4行とあの変な挙動が、そのまま説明できました。

「Web ページで JavaScript 関数を実行」を担当しているのは、`background.js` の `onRunJavaScript` です。流れは次のとおりです。

1. `chrome.debugger.attach` でタブにアタッチする
2. `chrome.debugger.sendCommand` で `Runtime.evaluate` を送る。`expression` には、PAD で書いた JavaScript がそのまま入る
3. 結果を受け取ったら `chrome.debugger.detach` で離れ、`result.result.value` を PAD に返す

ここに、ログの4行を当てはめます。

| ログ                                                                               | 何が起きたか                                                                                 |
| -------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------- |
| `Host access is restricted by policy.`                                           | 1 のアタッチが、ポリシーで拒否された                                                                    |
| `Debugger is not attached to the tab`                                            | アタッチできていないのに、2 の `Runtime.evaluate` を送った                                               |
| `Cannot read properties of undefined (reading 'result')`（`background.js:403:70`） | 2 が失敗したので、結果は空（`undefined`）。その `result.result.value` を読もうとして例外になった。403行目の70文字目は、まさにこの式 |
| `Debugger is not attached to the tab`                                            | 3 で離れようとしたが、そもそもアタッチしていない                                                              |

ポイントは、4行のうち3行に `Unchecked runtime.lastError` と付いていることです。拡張機能の API は、失敗すると `chrome.runtime.lastError` にエラーを入れます。`Unchecked` は、「それを誰も確かめなかった」という意味です。実際、この拡張機能の `background.js` には、`lastError` という単語が1つも出てきません。アタッチに失敗しても、成功したものとして先へ進む作りになっています。

そして、例外が起きた場所は、PAD に結果を返す処理の手前です。結果を返さないまま止まるので、PAD は返事を待ち続けることになります。1分の沈黙と、成功の扱いなのに結果が空欄という挙動は、ここから来ているのだと思います（PAD 本体が待ちきれなくなったときにどう扱うかは、拡張機能のソースからは分かりません）。

「ほかのブラウザ操作は動く」理由も、ソースで分かります。`chrome.debugger` を使っているのは `onRunJavaScript` と、そのためのアタッチの処理だけです。それ以外でページにスクリプトを入れるところは、`chrome.scripting.executeScript` で書かれています。今回の変更が止めたのは `chrome.debugger` だけなので、JavaScript の実行だけが巻き込まれました。

Microsoft の[マニフェスト V3 への移行のページ](https://learn.microsoft.com/ja-jp/power-automate/desktop-flows/manifest)にも、新しい拡張機能はデバッガーの機能でこのアクションの JavaScript を実行する、と書いてあります。同じページには、開発者ツールがグループ ポリシーで無効だと、このアクションは動かないともあります。つまり PAD のこのアクションは、[前回](/stories/DebugRouteTech/Policy)の `DeveloperToolsAvailability` の影響も受けます。

一方、Microsoft のトラブルシューティングには、[「実行されないのに、エラーも報告されない」という、今回とそっくりな症状のページ](https://learn.microsoft.com/ja-jp/troubleshoot/power-platform/power-automate/desktop-flows/browser-automation/js-not-executed)があります。ただし、原因として挙がっているのは、ほかの拡張機能による妨害だけです。ポリシーの話は出てきません。

## 振り返り

極論、まとめるとこうです。

```text
「Web ページで JavaScript 関数を実行」が動かない原因は、下記です。
https://developer.chrome.com/blog/debugger-enterprise-policy-restrictions?hl=ja

Chrome なら何もしなくていい。Edge は
「--disable-features=ExtensionDebuggerStrictPolicyRestrictions」を付けて起動し直せば直ります
```

どうです？ PAD と接点がありますか？ なかなか無いと思います。CDP に精通している私でも、ここにたどり着くまでに、結構 Copilot と対話しましたからね🙃

仮に「Chrome for Developers」を毎日見ていて、「Chrome 155 で chrome.debugger のエンタープライズ ポリシーの適用を強化」を見かけたとしても、「ふ～ん、CDP を直接使う私には関係ないわ」でスルーしていたかもしれません。

ただ、拡張機能まわりのポリシーへのテコ入れだからこそ、結果として PAD にも影響が出ちゃった、ということです。

### PAD だけではなかった

GitHub を見ると、ブラウザを操るほかの拡張機能でも、同じエラーの報告が上がっていました。AI がブラウザを操作するための拡張機能が目立ちます。

* [Claude in Chrome で、デバッガーを使うツールがすべて失敗する](https://github.com/anthropics/claude-code/issues/103981)

* [BrowserSkill が、CDP のアタッチを拒否される](https://github.com/Tencent/BrowserSkill/issues/490)

* [oh-my-pi のブラウザ中継が、Chromium 155 ですべてのアタッチを拒否される](https://github.com/can1357/oh-my-pi/issues/16748)

3つ目の報告によると、拡張機能は拒否されているのに、利用者にはただのタイムアウトとしか見えません。PAD の1分の沈黙と同じ形です。

### 何分で答えに？

どうやら、30分前後でこの答え（[Chrome の発表](https://developer.chrome.com/blog/debugger-enterprise-policy-restrictions?hl=ja)）にたどり着いたようです。

一方で、この「Web ページで JavaScript 関数を実行」の動作不良は、発生から20日たっても、まだ誰も明確な原因を伝えていませんでした。

ということで、私がこの原因を報告しました🤠

そして、いっぱいの感謝のリアクションスタンプをもらえました。たぶん、今年の中で最もうれしかった日です。

「拡張機能側の制約までは視野に入ってなかった😵」という声もあり、やはりブラウザって、ほんっとうに複雑なんだな～と改めて感じました。

とはいえ、CDP × VBA のツールを作ったことで、「ブラウザ自動化が、どの層で動いているか」が自然と身についたのは事実です。

「Web ページで JavaScript 関数を実行」とだけ見ると、一般の方からすれば、1つの大きなブラックボックスです。でも、私には中の候補が見えます。

```text
「Web ページで JavaScript 関数を実行」
┗ JS の実行は、Runtime.evaluate？ Runtime.callFunctionOn？ chrome.scripting.executeScript？
┗ ポリシーの制限？
┗ JS 自体が変？
┗ PAD 自体が変？
┗ Native Messaging が変？
┗ 標準入出力が変？
┗ etc...
```

「Web ページで JavaScript 関数を実行」という1つの言葉だけで、「たぶんこの辺が調査の候補だろうな～」という選択肢が、すぐに浮かびます。だからこそ、Copilot との対話もスムーズに進んだのかなと思います。

CDP × VBA のツール自体は、確かに、ほぼ自分しかいじれないほどの低レイヤーの集まりです。でも、その知識がこの技術トラブルの素早い解決につながったと思うと、とても気持ちいいです🤠

## そもそも、なぜ ExtensionSettings にテコ入れ？

ExtensionSettings 自体は、昔からあるポリシーです。では、なぜ今さらテコ入れが入ったのか。Copilot にたずねたら、CDP に慣れた私はすぐに納得しました。

* A サイトだけ、拡張機能に触らせたくない

* B サイトにアタッチして、CDP の準備 OK → その B サイトで使える CDP が、A サイトに届かない保証はあるか？

これが答えです。

ご存じのとおり、CDP はひじょ～に強力な、デバッグ用のブラウザ制御プロトコルです。

「もしかしたら、CDP を駆使（Page、DOM、Network、etc...）すれば、禁止しているはずの A サイトも間接的に操れてしまうかも🫠 よし、`runtime_blocked_hosts` に1つでもあったら、一律で CDP 禁止じゃ～🤪」

と、まぁ、乱暴というか、でも理解できちゃう背景です。

Chromium 側の説明も、同じ方向です。制限を入れたコミットには、デバッガーをアタッチすると生の CDP が使えるようになり、それを特定のサイトだけに安全に絞ることはできない、とあります。そのため、ポリシーの迂回を防ぐ目的でアタッチを拒否する、という説明です。Chrome の発表も、CDP は Web の「オリジン」（サイトごとの区切り）の仕組みより下の層で動くので、オリジン単位の絞り込みでは安全に制限できない、と説明しています。

しかも、今は AI 時代です。下手したら、AI に聞けば、いくらでも迂回ルートが出てきそうです🙃 だからこそ、最初はこれでリリースしたんだなと。

ただ、フィードバックを受けて、いったん旧ルールに戻して見直しているのが、今の Chrome です。

## 考察

さて、ここからは、私の意見も踏まえた考察です。

### chrome.debugger のポリシーについて

`chrome.debugger` には、入口（アタッチ）で止める仕組みはあります。しかし、その先の細かいコマンド単位の制御（`Runtime.evaluate` だけ許可する、など）は、今のところありません。

コマンド単位で絞れたら、それでいい気もします。ただ、Chrome の発表が書いているのは、先ほどの「サイト単位では安全に絞れない」というところまでです。コマンド単位で絞る話は、発表には出てきません。

発表の追記には「より柔軟な方式を設計する」とあるので、形を変えて戻ってくる可能性はあります。それがコマンド単位まで細かくなるかというと、私はなさそうだと見ています。理由は次のとおりです。

#### Runtime.evaluate だけにしても、さらなる沼

仮に、`Runtime.evaluate` だけを許可する形にしたとします。

ブラウザは日々進化します。昨日は使えなかった JavaScript が、今日は使える、なんてこともあります。

`Runtime.evaluate` に渡す中身を正規表現で検査しようにも、AI で迂回ルートが出やすい時代です。現実的ではありません。

メソッドを規制したら、次は params を規制し、その次は params の中身を規制し...時間と人がいくらあっても足りません🙃 いや、今なら AI がいますが、迂回ルートも AI で出ちゃいそうなので、意味がないよな～と。ということで、今後も入口で止める程度にとどまりそうです。

#### 似た機能の chrome.scripting.executeScript

`chrome.scripting.executeScript` も、JavaScript を実行するという機能自体は同じです。先ほどの `chrome.debugger` × `Runtime.evaluate` との違いは、次の点です。

* あらかじめ用意された JavaScript を実行するのか、自由な JavaScript を実行できるのか

マニフェスト V3 の拡張機能は、使うコードをすべて拡張機能の中に同梱しておく必要があります（[Chrome の説明](https://developer.chrome.com/docs/extensions/develop/migrate/remote-hosted-code?hl=ja)）。`chrome.scripting.executeScript` で動かせるのは、基本的に、拡張機能があらかじめ持っているコードです。

PAD の「Web ページで JavaScript 関数を実行」は、後者（自由な JavaScript）です。フローを作る人が、その場で好きな JavaScript を書きます。

PAD 側で、よく使う JavaScript を一式用意したとしても、次のような細か～い要望は必ず来ます。

* 「そうじゃないんだ、うちはこんな JS じゃないと無理なんだ🙃」

* 「ここの JS がああなら、もっといい感じにいけるんだけどな～」

* 「サイトの作りが変わったから、JS を1行だけ書き換えたいんだけど...」

だからこそ、ユーザーに委ねる自由な形、`chrome.debugger` × `Runtime.evaluate` になったのかなと思います。Chrome の発表は、`chrome.debugger` がどうしても必要でなければ `chrome.scripting` などへ移るよう勧めています。しかし PAD のこのアクションは、その「どうしても必要」な側です。

### 今年は結構、CDP 系のポリシーにテコ入れが

今回は拡張機能側の CDP まわりの変更でしたが、意外にも今年は、CDP まわりのポリシーや制限へのテコ入れが盛んでした。確かめられたものを並べます（Chromium の日付は、ソースに入った日です）。

| 時期      | 変更                                                                                                                                          |
| ------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| 2026年5月 | Edge 148 で、`DeveloperToolsAvailabilityAllowlist` と `DeveloperToolsAvailabilityBlocklist` が追加（[前回](/stories/DebugRouteTech/Policy)の記事で触れたもの） |
| 5月      | Chromium：CDP の子セッションを使って、アタッチや Cookie の制限をすり抜けられる問題への対処                                                                                     |
| 7月      | Chromium：拡張機能のデバッガー API で、Web Worker を使った制限の迂回を修正                                                                                           |
| 8月      | Chromium：`chrome.debugger` の一覧に特権のあるページを出さず、アタッチもできなくする                                                                                     |
| 8月〜11月  | 今回の `runtime_blocked_hosts` とスクリーンショット禁止の件（導入 → 延期 → 再導入 → 撤回）                                                                              |
| 9月      | Chromium：`DeveloperToolsAvailability` の説明に、CDP の直接接続と `chrome.debugger` にも効くことを明記                                                           |
| 10月     | Chromium：`chrome.debugger` から、WebUI の中の `<webview>` や、PDF の内側のフレームへのアクセスを止める                                                                |

こうして並べると、`chrome.debugger` という入口に、5月以降、何度も制限が足されているのが分かります。

[前回](/stories/DebugRouteTech/Policy)でも触れた [DeveloperToolsAvailability](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailability) も、テコ入れが入った1つです。

むか～しは、ただただ F12 を禁止するだけのポリシーでした。それが、いつの間にか、説明に次のものが並ぶようになっています。

* ページのソースの表示（Chrome 99 から）

* `--remote-debugging-port` と `--remote-debugging-pipe` による CDP の直接接続

* `chrome.debugger` 経由の CDP

下の2つは、2026年9月に Chromium のポリシーの説明へ書き足されたものです。コミットの題は「説明を明確にする」なので、動作そのものがいつからそうなのかは、確かめていません。

しばらくは、CDP 系のポリシーにも目を配らないといけなさそうですね。

## おわりに

今、CDP でブラウザを操るルートは、大きく2つあります。

* 直通（`--remote-debugging-port`、`--remote-debugging-pipe`、`edge://inspect/#remote-debugging`）

* 拡張機能経由（`chrome.debugger`）

直通ルートの CDP を制限するポリシーは、[前回](/stories/DebugRouteTech/Policy)触れたとおり、現時点では次の3つです。

* `RemoteDebuggingAllowed`

* `UserDataDir`（値が空のときは、起動引数を付けて始める CDP に限り、`RemoteDebuggingAllowed` を無効にしたのと同じ。値があっても、少しデバッグしづらくなる程度）

* `DeveloperToolsAvailability`

見てのとおり、非常にシンプルです。言ってしまえば、「CDP」のことだけを考えればいい話ですからね。

なお、`#remote-debugging` の経路も、Chromium のソースを見るかぎり、`RemoteDebuggingAllowed` が無効だと始まりません。一方で、普段使いのフォルダーかどうかの確認は、この経路には入っていません。実際に試したところ、`UserDataDir` が空でも、`#remote-debugging` の経路は使えました。`UserDataDir` が空のときに影響を受けるのは、`--remote-debugging-port` や `--remote-debugging-pipe` を付けて起動する場合のほうです。

一方、拡張機能はというと、こうなります。

| 何を決めるか | かかわるポリシー |
| --- | --- |
| 拡張機能のインストールの規制 | `ExtensionInstallSources`、`ExtensionAllowedTypes`、`BlockExternalExtensions`、`ExtensionInstallTypeBlocklist`、`ExtensionInstallForcelist`（強制インストール） |
| 拡張機能の許可リスト・禁止リスト | `ExtensionInstallAllowlist`、`ExtensionInstallBlocklist`（`ExtensionSettings` の `installation_mode` でも書ける） |
| 拡張機能そのものの無効化 | `ExtensionInstallBlocklist`（禁止リストに載ると、インストール済みでも無効になる）、`ExtensionSettings` の `blocked_permissions`（禁止した権限を求める拡張機能は、読み込まれない） |
| 拡張機能を動かしてもいいサイト | `ExtensionSettings` の `runtime_blocked_hosts`、`runtime_allowed_hosts` |
| 拡張機能が CDP を使ってもいいサイト | `DeveloperToolsAvailability`、`DeveloperToolsAvailabilityAllowlist`、`DeveloperToolsAvailabilityBlocklist` |
| 拡張機能が PC 側のアプリと通信してよいか | `NativeMessagingBlocklist`、`NativeMessagingAllowlist`（PAD の拡張機能は、Native Messaging で PAD 本体とやり取りします） |

etc...🫠

Edge のポリシー一覧を見ると、「拡張機能」のグループだけで14件あります（2026年11月12日時点。うち1件は非推奨）。たとえば `blocked_permissions` で `debugger` を禁止されると、ガイドの説明どおりなら、その権限を求める PAD の拡張機能は、JavaScript の実行どころか、拡張機能ごと読み込まれなくなります。

言ってしまえば、「拡張機能そのもののポリシー」と「CDP 系のポリシー」の組み合わせ（数学 A で習う、C のあれ）になるぶん、制限に引っかかりやすい、ということです。

少し前に書いたとおり、PAD 以外の、拡張機能を使った CDP（Claude など）も同じ変更を食らいました。それも、そういうことなのです。

拡張機能は、ストアから手軽にインストールできてしまいます。しかも、普段使っているブラウザにです。

普段使いのブラウザには、ログインしたままのサイトがたくさんあります。そして拡張機能は、権限さえもらえば、開いているページに JavaScript を入れたり、Cookie を読んだり、通信を書き換えたりできます。PAD の拡張機能も、すべてのサイトに対する権限（`<all_urls>`）を求めています。

これを悪用した実例もあります。2026年6月に Microsoft が公表した「StegoAd」です。広告ブロッカーや VPN、翻訳、動画ダウンローダーを装った拡張機能が、Edge の公式ストアに119個並んでいました。インストール数は、合わせて最大260万です。どれも表向きの機能はちゃんと動き、数日たってから、画像ファイルの中に隠したコードを動かします。Microsoft によると、Google のログイン情報や2段階認証のコード、ブラウザの Cookie が盗まれ、ページには勝手に広告が差し込まれていました。

公式ストアにあって表向きはちゃんと動く拡張機能でも、中身までは見た目で分かりません。

だからこそ、いろ～んなセキュリティ対策が施されているのです。

そう思うと、同じ「普段使いのブラウザでの CDP」なのに、`#remote-debugging` での直通 CDP には拡張機能系のポリシーがかからない、というのも、ある意味強力ですね🙃  
まぁ、こっちはこっちで別のポリシーがあるので、何とも言えないですが🙄

うちの会社の PC では、拡張機能系のポリシーはいっぱい入っていましたが、この CDP 直通系のポリシーは1つも入っていませんでした🫥

知ったうえで入れていないのか、はたまた、知らないから入れていないのか...

...まぁ、[WebView2 ルート](/stories/DebugRouteTech/Policy#webview2-ルートというチート)なら関係ないんですけどね🤪

## 参考資料

### 今回の変更（Chrome・Chromium）

* [Chrome 155 で chrome.debugger のエンタープライズ ポリシーの適用を強化](https://developer.chrome.com/blog/debugger-enterprise-policy-restrictions?hl=ja)（2026年9月8日公開、11月6日更新。11月4日付のロールバックの追記、Chrome 155 のベータ版・安定版の日付、2つのエラー文、回避用のフラグ）

* [chrome/browser/extensions/api/debugger/debugger\_api.cc](https://source.chromium.org/chromium/chromium/src/+/main:chrome/browser/extensions/api/debugger/debugger_api.cc)（`Host access is restricted by policy.` を返す箇所。フラグが無効のときは、URL ごとの判定に戻る。例外扱いの拡張機能は Perfetto UI だけ）

* [extensions/common/extension\_features.h](https://source.chromium.org/chromium/chromium/src/+/main:extensions/common/extension_features.h)（`kExtensionDebuggerStrictPolicyRestrictions` の説明。無効が既定で、無効にするとセキュリティ上の影響があるという注意書き）

* [\[Extensions\] Prevent attaching debugger when policy blocks some hosts](https://chromium.googlesource.com/chromium/src/+/e6e507e0011cf2357b7b053eedc86268e1403259)（2026年8月20日。制限を入れたコミットと、その理由）

* [Rollback of chrome.debugger enterprise policy restrictions](https://chromium.googlesource.com/chromium/src/+/403b75eb9b9617592da3a986fc3ebfceb8063147)（9月8日。154 には入れず、155 で入れ直す、という経緯）

* [\[extensions\] Add kill switch for strict debugger policy restrictions](https://chromium.googlesource.com/chromium/src/+/e0a733be75464d9cca4a444e15058a0ffecbc7f2)（9月16日。フラグの追加。この時点の既定は有効）

* [\[Extensions\] Disable strict debugger policy restrictions by default](https://chromium.googlesource.com/chromium/src/+/3c17d4ab3af1b2729b84647fd0114e79f5e3a817)（11月3日。フラグの既定を無効に）

* [components/policy/.../Extensions/ExtensionSettings.yaml](https://source.chromium.org/chromium/chromium/src/+/main:components/policy/resources/templates/policy_definitions/Extensions/ExtensionSettings.yaml)（Chrome 側の ExtensionSettings の説明。155〜158 で起こりうる、という書き方に変わっている）

### Edge 側

* [ExtensionSettings](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/extensionsettings)（バージョン 154 以降、`debugger` の権限を持つ拡張機能に `runtime_blocked_hosts` を適用すると、すべての対象で `chrome.debugger.attach()` が無効になる）

* [ExtensionSettings ポリシーの詳細なガイド](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-manage-extensions-ref-guide)（`runtime_blocked_hosts` と `runtime_allowed_hosts` の意味、`*` の既定の設定、書き方の例）

* [Microsoft Edge release notes for Stable Channel](https://learn.microsoft.com/en-us/deployedge/microsoft-edge-relnote-stable-channel)（156〜157。この制限や撤回に触れた記述はなし。2026年11月12日に確認）

* [Archived release notes for Microsoft Edge Stable Channel](https://learn.microsoft.com/en-us/deployedge/microsoft-edge-relnote-archive-stable-channel)（153〜155。同じく記述なし。Edge 155 の安定版は 2026年10月8日。Edge 148 で `DeveloperToolsAvailabilityAllowlist` / `DeveloperToolsAvailabilityBlocklist` を追加）

### Power Automate Desktop

* [マニフェスト V3 への移行](https://learn.microsoft.com/ja-jp/power-automate/desktop-flows/manifest)（新しい拡張機能は、デバッガーの機能で「Web ページで JavaScript 関数を実行」の JavaScript を実行する。開発者ツールが無効だと動かない）

* [Web ページでの JavaScript 関数の実行アクションの実行に失敗する](https://learn.microsoft.com/ja-jp/troubleshoot/power-platform/power-automate/desktop-flows/browser-automation/js-not-executed)（実行されないのにエラーも出ない症状。原因として挙がっているのは、ほかの拡張機能による妨害だけ）

* [ブラウザー自動化アクション リファレンス](https://learn.microsoft.com/ja-jp/power-automate/desktop-flows/actions-reference/webautomation)（「Web ページで JavaScript 関数を実行」アクションの説明）

* PAD の拡張機能（ID `kagpabjoboikccfdghpdlaaopmgpgfdc`、バージョン 2.73.0.36）の `manifest.json` と `background.js`（権限に `debugger` と `scripting`。`chrome.debugger` を使うのは `onRunJavaScript` だけ）

### ほかの拡張機能での報告

* [Claude in Chrome: all debugger tools fail with "Host access is restricted by policy" in Comet 155 (Chrome 156 unaffected)](https://github.com/anthropics/claude-code/issues/103981)（2026年10月21日）

* [Comet 155.0.8059.148 rejects BrowserSkill CDP attachment with no configured policies](https://github.com/Tencent/BrowserSkill/issues/490)（10月24日。同じフラグでの回避に触れている）

* [bug(browser-relay): Chromium 155 rejects every chrome.debugger.attach ("Host access is restricted by policy"); relay only reports a generic timeout](https://github.com/can1357/oh-my-pi/issues/16748)（10月22日）

* [Chrome/Edge 155: non-empty runtime\_blocked\_hosts now blocks chrome.debugger on every site](https://github.com/drpwchen/kimi-webbridge-lockdown/issues/3)（10月12日。Chrome 155 で、禁止していないサイトでも失敗することを実測）

### CDP まわりのほかの変更

* [DeveloperToolsAvailability](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailability)（CDP の直接接続にも、`chrome.debugger` 経由の CDP にも適用される）

* [Update Developer Tools policy descriptions to clarify CDP and remote debugging behavior](https://chromium.googlesource.com/chromium/src/+/89ae2af6173e)（2026年9月8日。`DeveloperToolsAvailability` の説明に、CDP の直接接続と `chrome.debugger` を明記）

* [debugger\_api.cc の変更履歴](https://github.com/chromium/chromium/commits/main/chrome/browser/extensions/api/debugger/debugger_api.cc)（2026年5月〜10月の、子セッション・Web Worker・特権のあるページ・`<webview>`・PDF のフレームに関する変更）

* [リモートホスト型コード違反に対処する](https://developer.chrome.com/docs/extensions/develop/migrate/remote-hosted-code?hl=ja)（マニフェスト V3 の拡張機能は、使うコードをすべて同梱する必要がある）

### 拡張機能系のポリシーと、直通ルートのポリシー

* [Microsoft Edge ブラウザー ポリシーに関するドキュメント](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies)（「拡張機能」のグループは14件、「ネイティブ メッセージング」のグループは3件。2026年11月12日に確認）

* [グループ ポリシーを使用して Microsoft Edge の拡張機能を管理する](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-manage-extensions-policies)（権限によるブロック、サイト単位のブロック、許可と禁止、強制インストール）

* [ExtensionInstallBlocklist](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/extensioninstallblocklist)、[ExtensionInstallAllowlist](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/extensioninstallallowlist)、[ExtensionInstallForcelist](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/extensioninstallforcelist)（禁止リストに載ると、インストール済みの拡張機能も無効になる。`*` はすべて禁止。強制インストールは禁止リストより優先）

* [ExtensionInstallSources](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/extensioninstallsources)、[ExtensionAllowedTypes](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/extensionallowedtypes)、[BlockExternalExtensions](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/blockexternalextensions)、[ExtensionInstallTypeBlocklist](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/extensioninstalltypeblocklist)（インストール元、種類、外部からのインストール、コマンド ラインからの読み込みの制限）

* [マニフェストで API のアクセス許可を宣言する](https://learn.microsoft.com/ja-jp/microsoft-edge/extensions/developer-guide/declare-permissions)（`debugger` は、chrome.debugger API へのアクセスを与える権限）

* [DeveloperToolsAvailabilityAllowlist](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailabilityallowlist)（対象ごとに URL で判定する。CDP の直接接続にも、`chrome.debugger` 経由の CDP にも適用）

* [NativeMessagingBlocklist](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/nativemessagingblocklist)（読み込んではいけないネイティブ メッセージング ホストの指定）

* [Inside StegoAd: How We Disrupted a Massive Malicious Extension Campaign](https://microsoftedge.github.io/edgevr/posts/Inside-StegoAd-How-We-Disrupted-a-Massive-Malicious-Extension-Campaign/)（2026年6月16日、Microsoft Edge Extensions Security Team。Edge の公式ストアから悪意のある拡張機能119個を削除。インストール数は最大260万）

* [RemoteDebuggingAllowed](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/remotedebuggingallowed)（無効にすると、リモート デバッグを使えない）

* [chrome/browser/devtools/remote\_debugging\_server.cc](https://source.chromium.org/chromium/chromium/src/+/main:chrome/browser/devtools/remote_debugging_server.cc)（`chrome://inspect` で有効にする承認つきのモードも、`RemoteDebuggingAllowed` が無効だと始まらない。既定のフォルダーかどうかを確かめるのは、起動引数の経路だけ）
