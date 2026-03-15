window.docsContent = window.docsContent || {};

window.docsContent['bidi-story'] = `
  <article class="story-article">
    <header class="story-header">
      <div class="story-tag">👾 登場秘話</div>
      <h1>なぜ <code>WebDriverBiDi.exe</code> なしで<br>ブラウザ自動化ができると<span class="story-highlight">わかったのか</span></h1>
      <p class="story-subtitle">〜公式ドライバーの闇を暴いた、土日の記録〜</p>
    </header>

    <section class="story-section">
      <img src="img/story/1.png" alt="海外フォーラムでの議論" class="story-img" loading="lazy">
      <p>すべては、ある海外の議論で見かけたこの意見から始まりました。</p>
      <blockquote class="story-quote">「CDP（Chrome DevTools Protocol）はChrome専用の独自仕様。将来性は WebDriver BiDi が上だ。」</blockquote>
      <p>言っていることは正しい。でも、現場のリアルは少し違います。企業環境において自動化を阻む本当の壁は「ブラウザのインストール禁止」よりも、<strong>「プリインストール以外のexeやNode.jsを情シスが許可しない」</strong>という壁でした。</p>
      <p>Windowsの標準搭載とExcelのインフラ化が後押しし、以下の組み合わせで環境はすでに整っていました。</p>
      <ul class="story-list">
        <li>REST WebAPI → <code>WinHTTP 5.1</code></li>
        <li>ブラウザ自動操作 → <code>Edge-CDP via Pipe</code></li>
        <li>WebSocket通信 → <code>Winhttp.dll</code></li>
      </ul>
      <p>それでも……<strong>「将来性はBiDiが上」</strong>という言葉が頭から離れなかった。</p>
    </section>

    <hr class="story-divider">

    <section class="story-section">
      <h2><span class="story-num">壁</span> <q>msedgedriver.exe</q> という巨大な存在</h2>
      <img src="img/story/2.png" alt="msedgedriverの壁" class="story-img" loading="lazy">
      <p>ネットを調べるたびに、同じ言葉が出てくる。</p>
      <blockquote class="story-quote">「まず、対応するWebDriver — Edgeなら <code>msedgedriver.exe</code> をダウンロードします」</blockquote>
      <p>ExcelでBiDiを使う記事があっても、結局は <em>exe × WebSocket</em> の構成に依存してしまう🫠<br>「やっぱりexeが要る。情シスの壁は越えられない……」と一度は諦めかけた。</p>
      <p>それでも手が止まらず調査を続けた数日後、あるリポジトリに辿り着いた。</p>
      <img src="img/story/3.png" alt="核となるリポジトリ" class="story-img" loading="lazy">
      <p class="story-link-note">🔗 <a href="https://github.com/GoogleChromeLabs/chromium-bidi" target="_blank">GoogleChromeLabs / chromium-bidi</a></p>
      <p>いかにもChromium公式のBiDiリポジトリ。しかしREADMEには <code>Node.js</code>、<code>npm</code>…「結局これも外部依存か」と落胆した。</p>
    </section>

    <hr class="story-divider">

    <section class="story-section">
      <h2><span class="story-num">転機</span> AIが明かした真実 — 主役は「exe」ではなかった</h2>
      <img src="img/story/4.png" alt="Google AI Pro 特典メール" class="story-img-half" loading="lazy">
      <p>そんな中、「<strong>Google AI Pro を3か月お試し</strong>」というメールが届いた。早速有効化し、<a href="https://antigravity.google/" target="_blank">Antigravity</a> といったAIにコードを読み込ませて解説してもらう機能を発見した。</p>
      <img src="img/story/5.png" alt="Antigravityのホームページ" class="story-img-half" loading="lazy">
      <div class="story-idea-box">
        💡 閃き：<strong>「AIに chromium-bidi のソースを読ませて、Excelで完全再現できないか？」</strong>
      </div>
      <p>そして、衝撃の事実が判明した。<code>Node.js</code> や <code>.exe</code> は単なる「<strong>運び屋（橋渡し役）</strong>」に過ぎなかった。AIは言った——</p>
      <blockquote class="story-quote">「<code>mapperTab.js</code> という巨大なJavaScriptファイルこそが、BiDiの心臓部です。」</blockquote>
      <img src="img/story/6.png" alt="BiDiの心臓部ソースコード" class="story-img" loading="lazy">
    </section>

    <hr class="story-divider">

    <section class="story-section">
      <h2><span class="story-num">手順</span> 実現のための「5ステップ」</h2>
      <img src="img/story/7.png" alt="大まかな手順" class="story-img" loading="lazy">
      <ol class="story-steps">
        <li><strong>JSの入手:</strong> <code>npm</code> や <code>JSDelivr</code> などのCDNから <code>mapperTab.js</code> を取得する。</li>
        <li><strong>特権の付与:</strong> CDPコマンド <code>Target.exposeDevToolsProtocol</code> でタブにブラウザ操作の特権を与える。</li>
        <li><strong>窓口の確保:</strong> <code>Runtime.addBinding</code> でVBAと通信するための受け取り口を確保。</li>
        <li><strong>注入と起動:</strong> <code>mapperTab.js</code> をタブに注入し、<code>Runtime.evaluate</code> でBiDiを起動する。</li>
        <li><strong>非同期通信:</strong> <code>Runtime.bindingCalled</code> イベントをキャプチャし、BiDiの非同期レスポンスを受け取る。</li>
      </ol>
      <p>もうお分かりだろう。<strong>「mapperTab.js というブラウザ内で動くプログラムが、BiDiコマンドをCDPに翻訳する作業をぜ〜〜んぶ肩代わりしていたのだ。」</strong></p>
    </section>

    <hr class="story-divider">

    <section class="story-section">
      <h2><span class="story-num">封印</span> Excelのセルにブラウザの心臓部を閉じ込める</h2>
      <p>主役はバイナリ（exe）ではなく、テキストデータ（js）だった。<br><strong>そうです。テキストなら、Excelのセルに置けちゃうのです🥳</strong></p>
      <img src="img/story/8.png" alt="Excelのセルにブラウザの心臓部を封印" class="story-img" loading="lazy">
      <p>数万行のスクリプトでも複数セルに分割して格納し、VBAの <code>Join</code> 関数で結合してブラウザへ注入できる。バイナリ（exe）は情シスに即ブロックされるが、テキストデータなら<strong>「ただのExcelファイル」</strong>としてパスできるのだ。</p>

      <div class="story-achievement">
        🏆 <strong>ついに、Excel単体でBiDiコマンドが実行できるようになりました！🥳🥳🥳</strong>
      </div>

      <p>土日を完全に溶かし、BiDi版の低レベル制御機能（<code>invokeMethod</code>, <code>invokeMethodAsync</code>, <code>TakeEvents</code>）を作り上げた。</p>
      <img src="img/story/12.png" alt="完成！" class="story-img" loading="lazy">

      <p>さらに、SeleniumVBAにあった「自動更新機能」も独自実装。<code>jsdelivr.com</code> のAPIを叩くことで:</p>
      <img src="img/story/9.png" alt="自動更新のサイト" class="story-img-half" loading="lazy">
      <ul class="story-list">
        <li>最新バージョンのチェック</li>
        <li><code>mapperTab.js</code> 自体の自動ダウンロード</li>
      </ul>
      <p>がVBA単体で完結。SeleniumVBAが「フォルダに <code>webdriver.exe</code> を配置」するのに対し、このツールは <strong>「ExcelのテーブルにJSのテキストを上書き🙂」</strong> するだけ。情シスの監視をすり抜ける<em>究極のステルス仕様</em>だ。</p>
    </section>

    <hr class="story-divider">

    <section class="story-section">
      <h2><span class="story-num">疑惑</span> 公式「msedgedriver.exe」も同じハックをしているのか？</h2>
      <p>夢は叶った。しかし一つの怖い疑問が湧いた。</p>
      <blockquote class="story-quote">「本当に公式の msedgedriver.exe も、ただの橋渡し役なのか？😰」</blockquote>
      <p>もしexeがもっと高度なネイティブロジックで動いていたら、自分の作ったものは「非公式の迂回ルート」になってしまう。</p>
      <p>そこで、<a href="https://github.com/hanamichi77777/WebDriverBiDi-via-VBA-test" target="_blank">SeleniumVBAのBiDi拡張版</a>を使い、<strong>公式ドライバーが裏で何をやっているか、この目で確認することにした。</strong></p>
      <img src="img/story/10.png" alt="SeleniumVBAにWebDriver BiDi機能を付けた拡張版" class="story-img" loading="lazy">

      <p>起動してみたが、画面に「BiDi-CDP Mapper is controlling this tab」というタブは見当たらない🥲<br>「やっぱり違う戦法か😩」と絶望しかけたが、AIからヒントが届いた。</p>
      <img src="img/story/11.png" alt="AIからの新たな手がかり" class="story-img" loading="lazy">

      <blockquote class="story-quote">「新しいタブを <strong>非表示（type: other）</strong> として生成している可能性があります。<code>edge://inspect/#devices</code> で確認できるはずです😋」</blockquote>

      <p>早速ブラウザのデバッグ画面を開くと……</p>
      <img src="img/story/13.png" alt="デバッグ画面を開くと" class="story-img" loading="lazy">
      <p>1つしか開いてないのにターゲットが3つ？🤔 クリックしてみると……<strong>あるじゃありませんか〜〜！🥹</strong></p>
      <img src="img/story/14.png" alt="あるじゃありませんか！" class="story-img" loading="lazy">

      <p>非表示タブなので画面は描画されないが、<code>outerHTML</code>をコピーしてファイル化してみると、BiDiコマンドが処理されている痕跡が残っていた。</p>
      <div class="story-img-grid">
        <img src="img/story/15.png" alt="outerHTMLをコピー" loading="lazy">
        <img src="img/story/16.png" alt="ファイル化してみると" loading="lazy">
        <img src="img/story/17.png" alt="BiDiコマンドを処理している痕跡" loading="lazy">
      </div>

      <p>さらにAIからの助言が続いた。<strong>「exe をテキストエディタで強引に開けば証拠が見つかるかもしれません」</strong></p>
      <p>「どうかプレーンテキストで残ってますように……！」と祈りながら、<code>msedgedriver.exe</code> をバイナリエディタで強引に開き、検索をかけた。</p>

      <div class="story-achievement">
        🔍 <strong>ヒットした！🥹😂🥳</strong>
      </div>

      <p>著作権表示と共に <code>&lt;!DOCTYPE html&gt;&lt;title&gt;BiDi-CDP Mapper&lt;/title&gt;...</code> という生々しいコードが、数万行にわたってハードコードされていた！</p>
      <img src="img/story/18.png" alt="バイナリエディタで検索結果" class="story-img" loading="lazy">

      <p>あの重厚な公式ドライバーも、裏では私がVBAでやったのと全く同じ <strong>「JSの翻訳機を隠しタブに注入する」</strong> という泥臭いハックをやっていた。「この土日の作業は無駄ではなかった😭」と心から納得できた。</p>
    </section>

    <hr class="story-divider">

    <section class="story-section">
      <h2><span class="story-num">結論</span> そして、未来へ</h2>
      <p>WebDriver BiDi 自体はまだβ版。将来的には <code>mapperTab.js</code> が不要になる日も来るかもしれない。しかし、AIをフル活用してこの「ロマンティックなツール」を作り上げたことは、最高の勉強になった。</p>
      <p>言語も環境（バイナリ、Node.js、Excelマクロ）も全く違うのに、<strong>ブラウザの裏口（CDP）を開けて翻訳機を忍ばせる</strong>という本質的なアプローチは、見事に共通していた。</p>

      <div class="story-platforms">
        <div class="story-platform-item">
          <span class="platform-icon">⚙️</span>
          <strong>exe によるオートメーション:</strong>
          <span>msedgedriver.exeなどが内部C++の文字列として隠し持ち、起動時に注入する</span>
        </div>
        <div class="story-platform-item">
          <span class="platform-icon">💚</span>
          <strong>Node.js によるオートメーション:</strong>
          <span>Google Chrome Labs の chromium-bidi リポジトリから直接呼び出される</span>
        </div>
        <div class="story-platform-item">
          <span class="platform-icon">📊</span>
          <strong>VBA によるオートメーション:</strong>
          <span>Excelのセルにテキストとして封印され、VBAマクロから直接ブラウザへ送り込まれる</span>
        </div>
      </div>

      <p>最新のWeb標準規格の裏側を支えているのが、たったひとつの巨大なJavaScriptファイル。それが「Excel VBA」というレガシーな環境にもピタリと当てはまる。</p>
      <div class="story-finale">
        技術の最先端と普遍性を同時に味わえた、<br>最高にエキサイティングで、<strong>ハッカーとしてのロマンに溢れた週末だった！</strong>
      </div>
    </section>
  </article>
`;
