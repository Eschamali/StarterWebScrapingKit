# ポリシーに関する研究レポート

<PromoMovie src="/promo/ReiwaIE_Policy.html" title="令和のIE 技術秘話 ― ポリシーの壁" />

さて、これでようやく3つの経路が開通し、好きな経路で Edge（Chromium）を制御で...きないことがあるのが今の時代です。

IE の頃とは違い、セキュリティ意識が強くなってしまった今は、ある概念も考えないといけません。

企業ポリシーです。

コンピューターを少し高度にいじろうとするとき、「ローカルグループポリシーエディター」や「レジストリエディター」で色々と設定を変えることがあると思います。その設定の中に、VBA から Edge を操ろうとすると引っかかって、まともに動かなくなるものがあります。

ここでは、現時点で分かっているポリシーについてまとめます。

当たり前ですが、ローカルグループポリシーエディターやレジストリエディターでの変更は、管理者権限がないと無理ですよ？

## ポリシーの種類

VBA で Edge を操るとき、ポリシーの壁は大きく分けて3種類あります。

- Office 系のポリシー
- ASR ルール
- Edge/Chrome 系のポリシー

ここでは、厳しい順に見ていきます。

## Office 系のポリシー

### Office アプリケーションの VBA を無効にする

![「Office アプリケーションの VBA を無効にする」ポリシー](/img/Policy/1.png)

おそらく、私が見つけた中で最も厳しいポリシーです。

内容からお察しのとおり、VBA の機能そのものを無効にします。うちのツールはもちろん、SeleniumVBA ですら使えなくなる最強のポリシーです。

これが設定されていたら...もうどうしようもないです。ただ、こういった環境は「特定の作業でしか使えない PC」であることが多いです。用意された専用ツール（ps1 や bat、社内製の exe、説明書など）しか使えなくても、必要な目的は果たせる状態になっていると思います。

ここはもう、用意されたツールだけで乗り切りましょう。

ちなみに、このポリシーの下で Access を起動すると、次のようなエラーが出て、そもそも VBE すら開けません。

![Access で Vbe7.DLL を読み込めないエラー](/img/Policy/2.png)

Excel の場合は、マクロの本体である `vbaProject.bin` が消され、VBE のアイコンも押せなくなります。

![Excel で vbaProject.bin が削除され、VBE のボタンが押せない様子](/img/Policy/3.png)

### そのほかの Office 系ポリシー

Office 系のポリシーで強力なのは先ほどの1つだけで、ほかはあまり脅威ではありません。強いて言うなら、次の2つだと思っています。

- マクロ実行時のスキャンの範囲
- 自動セキュリティ

![「マクロ実行時のスキャンの範囲」ポリシー](/img/Policy/4.png)

![「自動セキュリティ」ポリシー](/img/Policy/5.png)

「マクロ実行時のスキャンの範囲」が有効になっていると、一部のマクロの処理が遅くなる事例が確認されています。例えば、`Scripting.Dictionary` の中に `Scripting.Dictionary` を入れる、入れ子の処理が遅くなります。ただし、これは Office 2016 の頃は特に起きておらず、365 で起きた、という感じです。

色々と調べると、[VBA と AMSI](https://www.microsoft.com/en-us/security/blog/2018/09/12/office-vba-amsi-parting-the-veil-on-malicious-macros/) が連携していて、マクロの実行と同時に、悪意のある動きをしていないかをリアルタイムで確かめています。しかもこの監視は特定の関数に限らず、Win32 API や COM メソッド（`Scripting.Dictionary` も COM です）の呼び出しを全部記録する作りです（[Microsoft のドキュメント](https://learn.microsoft.com/en-us/windows/win32/amsi/how-amsi-helps)）。先ほどのグループポリシーの説明にも「影響を受ける VBA プロジェクトおよび XLM シートのランタイム パフォーマンスが低下する可能性があります」とあります。

そして、AMSI との連携は Office 365（今の Microsoft 365）のアプリにだけ入った機能です。同じ説明に「Office のサブスクリプション バージョンにのみ適用されます」ともあるので、2016 では起きず 365 で起きたことにも合致します。

実際、CopyMemory を呼ぶ速さを測った[この記事](https://web.archive.org/web/20250807224133/https://kuroihako.com/vba/vba-avast/)を見ると、マクロの実行中に Windows Defender が監視しているのが分かります。Windows Defender 以外の製品（記事では Avira Antivirus）を使えば、速さの低下はある程度抑えられるようですが、それはセキュリティ製品によります。

でも実は、これを避ける手法が最近出てきました。COM 経由で起動する方法です。

一番手軽に試せるのは、`Workbooks.Open` でマクロのブックを開くことです。これだけで、あら不思議、処理が速くなります。

でも、これにもちゃんとポリシーのメスが入っています。それが先ほどの「自動セキュリティ」です。これが「マクロを既定で無効にする」になっていると、この高速化のテクニックは使えず、監視付きの遅い VBA の処理のままになります。実行自体はできますけどね。

遅くてもまぁいいか、というなら、別に脅威ではありません。

## ASR ルール

ここからは [ASR ルール](https://learn.microsoft.com/ja-jp/defender-endpoint/attack-surface-reduction-rules-overview)です。

簡単に言うと、「こういう動きをするものは、どんなツールであろうとブロックする」という決め打ちのルールです。

### Office マクロからの Win32 API 呼び出しをブロックする

「Office マクロからの Win32 API」...そうです。`Declare` で簡単に Win32 API を呼べてしまう、あれです。

このルールが適用されると、次のようになります。

- Excel の中で完結する VBA の操作は OK
- Windows の機能を借りようとするとエラー

最初に紹介した「VBA を無効にする」よりは、ましなルールです。

ただし今のところ、「どの Windows API を禁止するか」を選ぶ仕組みはありません。ASR で管理者が除外を設定できるのも、[ファイルとフォルダー単位](https://learn.microsoft.com/ja-jp/defender-endpoint/attack-surface-reduction-rules-overview#file-and-folder-exclusions-for-asr-rules)だけです。`Declare` の宣言があるものは、一律に使えないという感じです。

![Sleep を使うマクロが ASR にブロックされた様子](/img/Policy/6.png)

上の画像は、VBA でおそらく最もよく使われる `Sleep` の例です。こんな単純なコードですら、ブロックの対象になります。

なお、動かすことはできます。ただし、保存も、開くこともできません。

#### 奇妙な現象

この Win32 API のブロックを調べていくうちに、矛盾を見つけました。次の2つは、Win32 API をふんだんに使っているのに、なぜかブロックされません。

- [WebDriver-BiDi-for-SeleniumVBA から落とせるマクロブック](https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA/releases/tag/v8.2)
- Access のマクロ

うちのツールでも、xlsm はブロックされますが、Access なら使えます。

個人で Access を持っている方は少ないと思いますが、最近の企業なら Microsoft 365 に Access も付いてくるので、こういう回避の手も使えてしまいます。

なぜ SeleniumVBA 系だけ見逃されるのかは分かりません。違いを挙げるなら、SeleniumVBA は歴史のあるツールだということです。そのため、ある程度信頼されたマクロとして扱われ、Win32 API の呼び出しをブロックする ASR の下でも使えるのかもしれません。これは推測です。

### すべての Office アプリケーションによる子プロセスの作成をブロックする

VBA だけでは厳しい処理や難しい処理を、ps1 や bat に任せる手法があります。[VBA から PowerShell を呼んで ZIP を解凍する手法](https://www.limecode.jp/entry/utility/extract-zip-file-with-macro)が、いい例ですね。

しかし、これをやった途端に、このルールが働きます。

![msedgedriver.exe の起動が ASR にブロックされた様子](/img/Policy/7.png)

そうなんです。SeleniumVBA を動かすのに欠かせない WebDriver の exe を起動しようとすると、このルールに引っかかって、まともに動かなくなります。

ただ、名前からお察しのとおり、このルールが止めるのは、Office から外部のプログラムを起動する（Shell など）経路だけです。実際、[この記事](https://note.com/sele_chan/n/nf15e30e08c94)のとおり、WebDriver の起動だけを手動でやれば回避できます。

#### 奇妙な現象

はい、またしても奇妙な現象のコーナーです。たしかに、VBA の `Shell` で起動すると、このルールにブロックされます。

でも、`Shell "msedge.exe"` はどうでしょう？ なんと、普通に起動できます！ 実際、このルールを適用した実験環境でも、ふつ～うに CDP で制御できます。

Office からブラウザを起動する場面は、VBA だけではありません。

- セル内のハイパーリンク
- リボンにある外部リンク

このルールは、VBA から起動したのか、セル内のハイパーリンクから起動したのかを区別していません。主に、起動される側の exe がどこにあるかで判断しています。

実際、Defender の中身（Lua で書かれたルール）を抜き出して公開している[リポジトリ](https://github.com/HackingLZ/ExtractedDefender/blob/main/asr/d4f940ab-401b-4efc-aadc-ad5f3c50688a)を見ると、このルールには除外するパスの一覧（`GetPathExclusions`）があり、`msedge.exe` や `chrome.exe`、Firefox といったブラウザがずらっと並んでいます。

つまり内部では、ホワイトリストで起動してよいかどうかを判断しているわけです。有名なブラウザで何もエラーが出ないのは、このためです。ただし、この一覧は 2022 年 6 月に抜き出されたもので、今の Defender とまったく同じとは限りません。

ちなみに、Outlook 向けの別のルール（「Office の通信アプリケーションによる子プロセスの作成をブロックする」）では、[Opera だけリンクを開けない](https://forums.opera.com/topic/71122/asr-attack-surface-reduction-gp-rule-block-only-office-communication-applications-from-creating-child-processes-blocks-opera-from-opening-any-links-other-browsers-work-just-fine)という報告もありました。こちらも、ブラウザごとに扱いが違うことが分かります。

## ここまでのまとめ

Win32 API を直接触れる `Declare` というチートのせいで、VBA を使った企業への攻撃が長年続きました。

- 実在の組織や人物になりすましたメールの添付ファイルを開き、「コンテンツの有効化」を押すだけで感染する（[Emotet](https://www.jpcert.or.jp/at/2019/at190044.html)。Word の設定によっては警告すら出ない）
- 実際の組織間のメールのやりとりを転用して、取引先からの返信を装う（同じく Emotet）
- マクロから Win32 API を呼び、ディスクに何も書かずに悪意のあるシェルコードを動かす（[ASR ルールの説明](https://learn.microsoft.com/ja-jp/defender-endpoint/attack-surface-reduction-rules-reference)）

しかし、だからといって VBE そのものを無効にしたり、廃止したりすることは、もうできません。それほど VBA は企業にしみわたりました。正しく使えば時短になり、業務の効率化につながり、貢献もできます。

一方で、悪い方向に使われてしまうのも事実です。

こういった攻撃を最小限に抑えつつ、健全な VBA 民も生かすために生み出されたのが、ここまで見てきたポリシーなのです。全部の PC に一律の規制をかけるのではなく、この PC は未設定、こっちは少し制限、あっちはもっと強めに...と、場面に応じて制限を強める、これが Microsoft の出した答えかなと思っています。

## Edge のポリシー

さて、ようやく[ブラウザポリシー](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies)のコーナーです。

VBA そのものは使えても、Edge 側の規制で使えない、なんてこともあります。

### [RemoteDebuggingAllowed](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/remotedebuggingallowed)

ブラウザポリシーで最も強力なのが `RemoteDebuggingAllowed` です。

名前のとおり、CDP による制御そのものができなくなります。

しかもやっかいなことに、ブラウザはこの警告を標準エラー出力にしか出しません（[Chromium のソース](https://source.chromium.org/chromium/chromium/src/+/main:chrome/browser/browser_process_impl.cc)でも、`stderr` に書き出すだけです）。そのため、ツールの問題だと思っていたら実はポリシーのせいだった、と気付くまでに時間をロスすることも珍しくありません。

```log
[2026/11/01 13:09:38.964] ERROR | CDPCore.viewStdoutAndStderr                         | COID580 | stderrout = DevTools remote debugging is disallowed by the system admin.
```

そういったこともあって、このツールは制御の前にポリシーを確かめ、設定されていればユーザーに警告を出します。

![ポリシーを検知したときの VBE のエラー](/img/Policy/8.png)

...何やら気になる文言がありますが、それは後ほど...

### [UserDataDir](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/userdatadir)

簡単に言うと、`--user-data-dir` 引数を無視して、ポリシーで決めたフォルダーを強制するものです。

このツールは基本的に、普段使いのブラウザデータと、自動制御用のブラウザデータを分けて使います。このポリシーがあると、その2つを共用することになってしまいます。

そして、すでに使用中のブラウザデータに対して、あとから CreateProcess で起動して CDP 接続することはできません。そのため、普段使いのブラウザを一度閉じてから CDP 接続すれば、なんとか自動制御できる、という状態です。

しかし、このポリシーにはちょっとした罠があります。理論上は、ポリシーは有効なのに値は空欄、という設定もできます。

すると、どうなるか？ ブラウザデータが、普段使いの "User Data" に固定されます。

[Chromium のソース](https://source.chromium.org/chromium/chromium/src/+/main:chrome/install_static/user_data_dir.cc)を見ても、ポリシーの値が空だと、既定のフォルダー（普段使いの "User Data"）に戻る作りになっています。

...思い出しましたか？ 少し前に、[Chromium 系で、普段使いのブラウザデータに対するデバッグを一律に禁止する](https://developer.chrome.com/blog/remote-debugging-port?hl=ja)アップデートが入ったことを。Chrome 136 からの変更で、Edge も同じ挙動です（Edge 側の公式発表は見当たりません）。

そうなんです。UserDataDir ポリシーを空で設定すると、`RemoteDebuggingAllowed:=0` と同じことになってしまうのです。なお、このときのエラーは `DevTools remote debugging requires a non-default data directory. Specify this using --user-data-dir.` で、[ソース](https://source.chromium.org/chromium/chromium/src/+/main:chrome/browser/browser_process_impl.cc)上は `RemoteDebuggingAllowed` のときとは別の文言です。

UserDataDir もまた、条件しだいで CDP 接続できなくなるポリシーなのです。でも...

### [DeveloperToolsAvailability](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailability)

最後はこれです。簡単に言うと、F12 が使えなくなります。

![「このページでは DevTools は許可されていません」のダイアログ](/img/Policy/9.png)

エラーメッセージだけを見ると F12 だけの規制かと思いきや、CDP による制御（`--remote-debugging-port` と `--remote-debugging-pipe`）も規制の対象になるのです。[Microsoft のドキュメント](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailability)にも「Chrome DevTools プロトコル(CDP)の直接接続(--remote-debugging-port や --remote-debugging-pipe を使用する接続など)にも適用され」るとはっきり書いてあります。

やっかいなのは、このツールで例えると、次のような止まり方をすることです。

- CDPBrowser 単位：CDP 可能
- CDPContext 単位：CDP 不可能

つまり、ブラウザそのものは制御できても、タブは制御できません。ドキュメントにも「リモート デバッグが許可されている場合でも、このポリシーは検査できるターゲットを制御します」とあります。Chromium のソースでは、接続（attach）の前に[ターゲットごとにポリシーを確かめて](https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/devtools_agent_host_impl.cc)いて、[プロファイルを持たないブラウザ全体は素通り](https://source.chromium.org/chromium/chromium/src/+/main:chrome/browser/devtools/chrome_devtools_manager_delegate.cc)、タブ（ページ）だけが止められる作りです。

当然、ブラウザそのものを操れるだけでは、ほとんど意味がありません。タブも制御できてこそ意味があります。

今のところ、このツールにはこのポリシーを確かめる機能はありません。ただ、実装するとなると割と面倒です。

このポリシーは、さっきの2つと違って、`DeveloperToolsAvailability:=2` かどうかを見るだけでは済みません。ほかのポリシーとも照らし合わせないと、正しく判定できないのです。

[Microsoft のドキュメント](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailability)の説明をそのまま借りると、こうです。

1. URL が DeveloperToolsAvailabilityAllowlist のパターンと一致する場合、開発者ツールは許可されます。
2. 許可リストが構成されていて、ブロックリストが構成されていない場合、許可リストにない URL はブロックされます。
3. URL が DeveloperToolsAvailabilityBlocklist のパターンと一致する場合、開発者ツールはブロックされます。
4. それ以外の場合で、どちらの URL リスト ポリシーも構成されていない場合、または URL がどちらのリストにも一致しない場合、このポリシー (DeveloperToolsAvailability) がフォールバックとして適用されます。

このとおり、2つのリストを照らし合わせないといけないので、まあまあ難易度が高いです。ちなみに、この2つのリスト（`DeveloperToolsAvailabilityAllowlist` / `DeveloperToolsAvailabilityBlocklist`）は、[Edge 148（2026年5月）](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-relnote-archive-stable-channel)で追加されたばかりのポリシーです。

許可リストに入っている URL なら CDP は普通に動くので、ポリシーがあるからといって門前払い...というわけにもいきません。

CDPBrowser-attachToTab をするときにこのポリシーが働くと、`{"id":3,"error":{"code":-32000,"message":"Not allowed"}}` というエラーが返ります（[attach に失敗したときの "Not allowed"](https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/protocol/target_handler.cc)です）。そのときに、どのポリシーに引っかかっているかが分かるエラーメッセージにするのがいいかな、と考えています。

## WebView2 ルートというチート

さて、`RemoteDebuggingAllowed` の VBE エラーに、ﾁｮｯﾄｷﾆﾅﾙ文言がありましたよね。

「WebView2ルートでの制御をお試しください。」

そうなんです。WebView2 ルートなら、先ほど挙げたポリシーをすべてバイパスできてしまうのです。

![WebView2 ルートでの制御](/img/Policy/10.gif)

`RemoteDebuggingAllowed` と `DeveloperToolsAvailability` が拒否の値になっていても、ふつ～に CDP で制御できているのが分かると思います。

実は、先ほど挙げた3つのポリシーは、あくまでも Edge.exe に対してしか効果がありません。[Microsoft のドキュメント](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/concepts/enterprise)にも、「Microsoft Edge のポリシーは、WebView2 ランタイム アプリと WebView2 アプリには適用されません」とはっきり書いてあります。

その代わり、WebView2 には[独自のポリシー](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-webview-policies)があります。しかし、その中身は次のとおりで、Edge.exe と比べるとはるかに少ないです。

- WebView2 を動かす本体（msedgewebview2.exe）の場所
- 通信まわり
- その他（PDF ビューアーや XSLT の切り替えなど）

そして、デバッグ系のポリシーは...何ひとつありません。これが、WebView2 なら CDP で制御できてしまった理由です。

## おわりに

Edge.exe と msedgewebview2.exe は、どちらもエンジン自体は Chromium 系ですが、何を優先するべきかが違います。Microsoft も、アプリとブラウザーでは使い方が違うので、Edge のポリシーを WebView2 に適用しないのは[仕様](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/concepts/enterprise)だと説明しています。

- Edge.exe は、企業ファースト（引き算）
- msedgewebview2.exe は、開発者ファースト（足し算）

そんな印象を受けました。

Edge.exe は、基本的にはウェブサイトを閲覧するためのものです。しかし時代とともに、ブラウザでできることがいっぱい増えました。たとえば、次のようなものです。

- パスワード マネージャーと保護
- ゲーム設定
- デバッグ機能
- 個人用設定

しかし企業から見れば、「たかが社内で使うブラウザに、明らかに要らない機能が多すぎる。そのイラン機能のせいで、社員がさぼったり、何か Web ハックしようとしたりしたら、たまったもんじゃない」という声が聞こえてきそうです。

そこで、必要な機能だけに絞って規制できるように、Edge には大量の制限ポリシーが生まれました。

一方、WebView2 は本来、ネイティブ アプリのさまざまな部分に Web のコードを埋め込んだり、画面のデザインに Web の描画を使ったりするための、アプリ開発者向けのものです。

自分の環境ではうまくいくのに、いざ別の環境に持っていったら Edge のポリシーのせいでぜんぜん動かない...なんてことになったら、開発者もぷんすこぷんぷんです。ポリシーを考えながら開発しようとしたら、頭がパンクします。実際、Microsoft のドキュメントでも、ブラウザーで JavaScript をブロックしたら JavaScript を使う WebView2 アプリが動かなくなる、という例を挙げています。

だからこそ、WebView2 のポリシーは必要最小限に抑えられています。

では、WebView2 には制限する手段がないのかというと、そうではありません。Edge なら企業がポリシーで決めることの多くを、WebView2 では開発者が API で決めるようになっています。

たとえば、F12 の禁止です。Edge には `DeveloperToolsAvailability` ポリシーがあり、WebView2 には `ICoreWebView2Settings` の `AreDevToolsEnabled` があります。どちらも F12 で DevTools を開けるかどうかを決める設定で、どちらにもあります。違うのは、企業が設定するか、開発者が設定するかです。

ただし、効く範囲は同じではありません。Edge のポリシーは CDP の接続まで止めますが、[AreDevToolsEnabled の説明](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2settings)にあるのは、右クリックメニューやショートカットキーで DevTools の窓を開けるかどうかだけです。

実際に、このツールで `AreDevToolsEnabled` を False にしてから、CDP でページ遷移させてみました。最初の遷移も、UserForm を表示したあとにボタンから行う遷移も、どちらも普通に動きました。WebView2 では、DevTools の窓を禁止しても CDP は使えます。

ほかにも、次のような組み合わせがあります。

| 制限したいこと | Edge（企業がポリシーで設定） | WebView2（開発者が API で設定） |
| --- | --- | --- |
| F12（DevTools） | `DeveloperToolsAvailability` | `AreDevToolsEnabled` |
| JavaScript | `DefaultJavaScriptSetting` | `IsScriptEnabled` |
| ショートカットキー | `ConfigureKeyboardShortcuts` | `AreBrowserAcceleratorKeysEnabled` |
| パスワードの保存 | `PasswordManagerEnabled` | `IsPasswordAutosaveEnabled` |
| InPrivate | `InPrivateModeAvailability` | `IsInPrivateModeEnabled` |
| ダウンロード先 | `DownloadDirectory` | `DefaultDownloadFolderPath` |
| 拡張機能 | `ExtensionInstallBlocklist` | `AreBrowserExtensionsEnabled` |
| 追跡防止 | `TrackingPrevention` | `PreferredTrackingPreventionLevel` |

WebView2 の設定系の API（`ICoreWebView2Settings`・`ICoreWebView2Profile`・`ICoreWebView2EnvironmentOptions`）を数えると、全部で40項目ほどあります。そのうち18項目に、同じ機能を扱う Edge のポリシーがありました。対応は名前と説明文から判断したもので、ポリシーのページを1件ずつ確かめたわけではありません。

ちなみに、ポリシーの数は Edge が797件、WebView2 が15件です（どちらも2026年10月更新の [Microsoft Learn の一覧](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies)で数えました）。

::: details 40項目の対応表

**ICoreWebView2Settings**（18項目のうち、対応あり6）

| WebView2 | Edge のポリシー |
| --- | --- |
| `IsScriptEnabled` | `DefaultJavaScriptSetting` / `JavaScriptAllowedForUrls` / `JavaScriptBlockedForUrls` |
| `IsWebMessageEnabled` | なし（WebView2 固有） |
| `AreDefaultScriptDialogsEnabled` | なし |
| `IsStatusBarEnabled` | なし |
| `AreDevToolsEnabled` | `DeveloperToolsAvailability` |
| `AreDefaultContextMenusEnabled` | なし |
| `AreHostObjectsAllowed` | なし（WebView2 固有） |
| `IsZoomControlEnabled` | なし |
| `IsBuiltInErrorPageEnabled` | なし（`AlternateErrorPagesEnabled` は別物） |
| `UserAgent` | なし（`UserAgentReduction` などでは文字列を決められない） |
| `AreBrowserAcceleratorKeysEnabled` | `ConfigureKeyboardShortcuts` |
| `IsPasswordAutosaveEnabled` | `PasswordManagerEnabled` |
| `IsGeneralAutofillEnabled` | `AutofillAddressEnabled` / `AutofillCreditCardEnabled` |
| `IsPinchZoomEnabled` | なし |
| `IsSwipeNavigationEnabled` | なし（`KioskSwipeGesturesEnabled` はキオスク用） |
| `HiddenPdfToolbarItems` | なし |
| `IsReputationCheckingRequired` | `SmartScreenEnabled`（WebView2 側は Windows の SmartScreen の設定が優先） |
| `IsNonClientRegionSupportEnabled` | なし |

**ICoreWebView2Profile**（10項目のうち、対応あり7。`ICoreWebView2Profile6` のパスワード・自動入力は Settings と重なるので数えていません）

| WebView2 | Edge のポリシー |
| --- | --- |
| `DefaultDownloadFolderPath` | `DownloadDirectory` / `DefaultDownloadDirectory` |
| `IsInPrivateModeEnabled` | `InPrivateModeAvailability` |
| `PreferredColorScheme` | なし |
| `ClearBrowsingData` | `ClearBrowsingDataOnExit` / `AllowDeletingBrowserHistory` |
| `PreferredTrackingPreventionLevel` | `TrackingPrevention` |
| `SetPermissionState` | `DefaultGeolocationSetting`・`DefaultNotificationsSetting` などの `Default*Setting` |
| `CookieManager` | `DefaultCookiesSetting` / `CookiesAllowedForUrls` / `BlockThirdPartyCookies` など |
| `AddBrowserExtension` | `ExtensionInstallForcelist` / `ExtensionInstallBlocklist` |
| `Delete` | なし |
| `AreWebViewScriptApisEnabledForServiceWorkers` | なし |

**ICoreWebView2EnvironmentOptions**（12項目のうち、対応あり5。うち2つは近いもの）

| WebView2 | Edge のポリシー |
| --- | --- |
| `AdditionalBrowserArguments` | なし（Edge は `RemoteDebuggingAllowed` などで、起動引数を止める側） |
| `AllowSingleSignOnUsingOSPrimaryAccount` | `AADWebSiteSSOUsingThisProfileEnabled` / `MSAWebSiteSSOUsingThisProfileAllowed`（近いもの） |
| `Language` | `ApplicationLocaleValue` |
| `TargetCompatibleBrowserVersion` | なし |
| `ExclusiveUserDataFolderAccess` | なし |
| `IsCustomCrashReportingEnabled` | `MetricsReportingEnabled` / `DiagnosticData`（近いもの） |
| `CustomSchemeRegistrations` | なし |
| `EnableTrackingPrevention` | `TrackingPrevention` |
| `AreBrowserExtensionsEnabled` | `ExtensionInstallBlocklist` |
| `ChannelSearchKind` | なし（WebView2 のポリシーに同じ名前のものがある） |
| `ReleaseChannels` | なし（WebView2 のポリシーに同じ名前のものがある） |
| `ScrollBarStyle` | なし |

:::

Microsoft も、[WebView2 のエンタープライズ管理](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/concepts/enterprise)の中で、IT 管理者が WebView2 ランタイムを直接管理するより、WebView2 アプリの開発者が独自のグループポリシーを実装することを勧めています。理由は、IT 管理者にとっては、WebView2 ランタイムよりもアプリを管理するほうが簡単だから、とあります。

とはいえ、中身は所詮 Chromium です。外部サイトの URL につなげば、ふつ～にネットサーフィンもできます。

本来なら、Visual Studio を用意して、WebView2 を埋め込んだ exe を1つ作らないと、msedgewebview2.exe でネットサーフィンはできませんでした。

しかし、人間には面倒な次のような土台を、AI がさくっと用意してくれました。

- WebView2 の vtable やサンクの配線
  - `CallDevToolsProtocolMethod(ForSession)` など

おかげで作者は、CDP の制御をさばく流れだけを考えればよくなりました。

そうして生まれちゃったのが、第三のルートである WebView2 なのです。

- Access なら、「Office アプリケーションの VBA を無効にする」がない限り、WebView2 でスクレイピングできます。
- Excel なら、それに加えて「Office マクロからの Win32 API 呼び出しをブロックする」がない限り、WebView2 でスクレイピングできます。

Edge のポリシー設定は、WebView2 では無力化されます。

AI が登場してしまったことで、このチートルートが生まれてしまいました。

もしかしたら、今までのセキュリティは「理論上はできるが、実際に手を動かそうとすると超面倒」ということで、かなり成り立っていたのかもしれません。しかし今なら、その理論さえひらめいてしまえば...

あとは AI が土台を作ってくれます。土台さえもらえれば、不思議と人間の創作意欲にバフがかかって、ちゃっちゃと完成してしまいます。

その完成品が、良となるか悪となるか...

結局は、作り手しだいなのです。

## ちなみに

SeleniumVBAならぬ、WebDriver.exeルートでのポリシー周りはどうなっているのか？

結論から言うと、どれもポリシーチェックをしていないかのようなエラーで返る感じでした

### RemoteDebuggingAllowed
![seleniumvbaのポリシーRemoteDebuggingAllowedエラー](/img/Policy/seleniumvba-RemoteDebuggingAllowed.png)

タイムアウトでエラーになっちゃった

### UserDataDir
![seleniumvbaのポリシーUserDataDirエラー](/img/Policy/seleniumvba-UserDataDir.png)

こちらは、クラッシュエラーです。`DevToolsActivePort`という、このツールでは見慣れたワードが出てきましたね

### DeveloperToolsAvailability
![seleniumvbaのポリシーDeveloperToolsAvailabilityエラー](/img/Policy/seleniumvba-DeveloperToolsAvailability.png)

こっちは、CDPの生結果が返ってますね

## 参考資料

### Office 系のポリシー

* [How to turn off Visual Basic for Applications when you deploy Office](https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/turn-off-visual-basic-for-application)（「Disable VBA for Office applications」で `VBAOFF` が 1 になる）

* [Office VBA + AMSI: Parting the veil on malicious macros](https://www.microsoft.com/en-us/security/blog/2018/09/12/office-vba-amsi-parting-the-veil-on-malicious-macros/)（2018年9月。Office 365 のアプリだけ、Win32 API と COM の呼び出しを記録、「Macro Runtime Scan Scope」、性能が下がることがある）

* [How the Antimalware Scan Interface (AMSI) helps you defend against malware](https://learn.microsoft.com/en-us/windows/win32/amsi/how-amsi-helps)（VBA ランタイムが Win32・COM・VBA の API 呼び出しを循環バッファーに記録する）

* [Malicious macros were found](https://support.microsoft.com/en-us/office/vba/malicious-macros-were-found)（実行時スキャンが外れる例外。信頼できる場所と、信頼できる署名）

* [Application.AutomationSecurity property (Excel)](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)（プログラムから開くときの既定は `msoAutomationSecurityLow`＝すべてのマクロを有効）

* [[VBA]AvastがVBAから呼び出したCopyMemory(WinAPI)を監視しまくる件](https://web.archive.org/web/20250807224133/https://kuroihako.com/vba/vba-avast/)（2020年。CopyMemory 1回あたり Avast 1970ns、Windows Defender 447ns、Avira 3ns）

* [Heap-based AMSI bypass for MS Excel VBA and others](https://code-white.com/blog/2019-07-heap-based-amsi-bypass-in-vba/)（2019年。Win32 API を約12万回呼ぶループが、Defender ありで15分以上、なしで1〜2分）

### ASR ルール

* [攻撃面の減少 (ASR) ルールの参照](https://learn.microsoft.com/ja-jp/defender-endpoint/attack-surface-reduction-rules-reference)（子プロセスのルールの対象は Word・Excel・PowerPoint・OneNote・Access）

* [ASR ルールのファイルとフォルダーの除外](https://learn.microsoft.com/ja-jp/defender-endpoint/attack-surface-reduction-rules-overview#file-and-folder-exclusions-for-asr-rules)

* [HackingLZ/ExtractedDefender asr/d4f940ab-401b-4efc-aadc-ad5f3c50688a](https://github.com/HackingLZ/ExtractedDefender/blob/main/asr/d4f940ab-401b-4efc-aadc-ad5f3c50688a)（2022年6月に Defender から抜き出した子プロセスのルール。`GetPathExclusions` にブラウザが並ぶ）

* [commial/experiments windows-defender/ASR](https://github.com/commial/experiments/tree/master/windows-defender/ASR)（ASR のルールが Lua で書かれていることと、その読み方）

* [【SeleniumVBA】VBAからWebDriverが起動できない？ Microsoft Defender ASRを実機検証して対応版を作ってみた](https://note.com/sele_chan/n/nf15e30e08c94)（2026年9月。`msedgedriver.exe` の起動がブロックされ、手動起動で回避）

* [ASR GP rule (Block only Office communication applications from creating child processes) blocks Opera from opening any links!](https://forums.opera.com/topic/71122/asr-attack-surface-reduction-gp-rule-block-only-office-communication-applications-from-creating-child-processes-blocks-opera-from-opening-any-links-other-browsers-work-just-fine)（Outlook 向けのルールで Opera だけリンクを開けない報告）

### Edge のポリシー

* [RemoteDebuggingAllowed](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/remotedebuggingallowed)（無効にすると `--remote-debugging-port` も `--remote-debugging-pipe` も使えない）

* [UserDataDir](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/userdatadir)（`--user-data-dir` を指定しても、ポリシーのフォルダーが使われる）

* [DeveloperToolsAvailability](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies/developertoolsavailability)（CDP の直接接続にも適用。許可リスト・ブロックリストとの判定順）

* [Microsoft Edge Stable チャネルのアーカイブされたリリース ノート](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-relnote-archive-stable-channel)（Version 148.0.3967.54 で `DeveloperToolsAvailabilityAllowlist` / `DeveloperToolsAvailabilityBlocklist` を追加）

* [Changes to remote debugging switches to improve security](https://developer.chrome.com/blog/remote-debugging-port?hl=ja)（2025年3月。Chrome 136 から、既定のデータフォルダーではリモートデバッグのスイッチが効かない）

* [Microsoft Edge Remote Debugging is Not Working When Using the Default User Profile](https://learn.microsoft.com/en-us/answers/questions/5976829/microsoft-edge-remote-debugging-is-not-working-whe)（2026年8月の Microsoft Q&A。Edge でも同じ挙動で、Edge 固有の発表は見当たらないという回答）

* [chrome/browser/browser\_process\_impl.cc](https://source.chromium.org/chromium/chromium/src/+/main:chrome/browser/browser_process_impl.cc)（`DevTools remote debugging is disallowed by the system admin.` などを `stderr` に出す）

* [chrome/browser/devtools/remote\_debugging\_server.cc](https://source.chromium.org/chromium/chromium/src/+/main:chrome/browser/devtools/remote_debugging_server.cc)（`kDisabledByPolicy` / `kDisabledByDefaultUserDataDir`）

* [chrome/install\_static/user\_data\_dir.cc](https://source.chromium.org/chromium/chromium/src/+/main:chrome/install_static/user_data_dir.cc)（`UserDataDir` の値が空なら既定のフォルダーに戻る）

* [content/browser/devtools/devtools\_agent\_host\_impl.cc](https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/devtools_agent_host_impl.cc)（attach の前の `AllowInspectingTarget`）

* [chrome/browser/devtools/chrome\_devtools\_manager\_delegate.cc](https://source.chromium.org/chromium/chromium/src/+/main:chrome/browser/devtools/chrome_devtools_manager_delegate.cc)（プロファイルの無いターゲットは許可）

* [content/browser/devtools/protocol/target\_handler.cc](https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/protocol/target_handler.cc)（`kNotAllowedError = "Not allowed"`）

### WebView2 のポリシー

* [WebView2 のエンタープライズ管理](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/concepts/enterprise)（Microsoft Edge のポリシーは WebView2 には適用されない。仕様）

* [Microsoft Edge WebView2 - ポリシー](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-webview-policies)（ローダーの上書き設定・ネットワーク設定・その他の3分類。2026年10月7日更新の版で15件）

* [Microsoft Edge - ポリシー](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-policies)（2026年10月27日更新の版で797件）

* [ICoreWebView2Settings](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2settings)（`ICoreWebView2Settings2`〜`9` も含めて18項目。`AreDevToolsEnabled` は DevTools の窓を開けるかどうか。`ICoreWebView2Settings8` の `IsReputationCheckingRequired` は、Windows の SmartScreen の設定が無効だと上書きされる）

* [ICoreWebView2Profile](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2profile)（`ICoreWebView2Profile2`〜`9` も含む）

* [ICoreWebView2EnvironmentOptions](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2environmentoptions)（`ICoreWebView2EnvironmentOptions2`〜`8` も含む）

### マクロを使った攻撃

* [マルウエア Emotet の感染に関する注意喚起](https://www.jpcert.or.jp/at/2019/at190044.html)（2019年、JPCERT/CC）

* [インターネットからのマクロは、Office で既定でブロックされます](https://learn.microsoft.com/ja-jp/microsoft-365-apps/security/internet-macros-blocked)
