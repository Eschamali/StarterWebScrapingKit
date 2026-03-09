window.docsContent = window.docsContent || {};

window.docsContent['technical-journey'] = `
    <h1>技術的な道のり：<span style="color:var(--accent-color)">EXEレスBiDiの実現</span></h1>
    
    <p>なぜ、WebDriverBiDi.exe なしで動くことがわかったのか。その「生々しい」実現の記録です。</p>

    <div class="card mb-4">
        <h3>1. 従来の常識と「壁」</h3>
        <p>通常、Selenium や初期の WebDriver BiDi 実装では、ブラウザを操作するために <code>chromedriver.exe</code> や <code>WebDriverBiDi.exe</code> といった「中間バイナリ（EXE）」が必要でした。これらは：</p>
        <ul>
            <li>HTTPやWebSocketのプロトコル変換を行う代理人（Proxy）として機能</li>
            <li>ブラウザの起動オプションやセッション管理を統括</li>
        </ul>
        <p>という役割を担っていましたが、VBA環境においては「外部EXEの配布・管理」が最大のネックとなっていました。</p>
    </div>

    <div class="card mb-4">
        <h3>2. 突破口：BiDiPoc.bas が証明した「セルフ・プロキシ」</h3>
        <p>本プロジェクトの核心は、<strong>「EXEがやっている変換処理を、ブラウザ内部のJavaScriptに行わせる」</strong>という逆転の発想にあります。</p>
        <p>この着想の原型は、リポジトリ直下の <code>BiDiPoc.bas</code>（Proof of Concept）に記録されています。</p>
        
        <h4 style="margin-top:1.5rem;">実現のための「3つの神器」</h4>
        
        <div class="alert">
            <div class="alert-content">
                <strong>① Target.exposeDevToolsProtocol</strong>
                <p>ブラウザ内部の特定のタブ（Mapper用タブ）に対して、JSから直接CDPを叩ける特別なオブジェクト <code>window.cdp</code> を露出させます。</p>
            </div>
        </div>

        <div class="alert">
            <div class="alert-content">
                <strong>② Runtime.addBinding</strong>
                <p>JSからVBA（ホスト側）へメッセージを「逆流」させるためのブリッジ関数を作ります。これにより、BiDiのイベントがVBAへ通知されるようになります。</p>
            </div>
        </div>

        <div class="alert">
            <div class="alert-content">
                <strong>③ Runtime.evaluate による mapperTab.js の注入</strong>
                <p>Chromium公式チームが開発している <code>chromium-bidi</code> のコアロジックをJS文字列としてブラウザに流し込み、インスタンスを起動します。</p>
            </div>
        </div>
    </div>

    <div class="card mb-4">
        <h3>3. 実現のメカニズム</h3>
        <div style="text-align:center; padding: 1rem;">
            <div style="display:inline-block; text-align:left; border:1px solid rgba(255,255,255,0.1); padding:1rem; border-radius:8px; background:rgba(255,255,255,0.02);">
                <p><strong>1. VBA側:</strong> BiDi形式のJSONを <code>evaluate</code> でJSへ投下</p>
                <div style="text-align:center; color:var(--accent-color)">↓</div>
                <p><strong>2. ブラウザ内(JS):</strong> <code>window.cdp</code> を通じて自分自身へCDP命令</p>
                <div style="text-align:center; color:var(--accent-color)">↓</div>
                <p><strong>3. ブラウザ内(JS):</strong> 結果を <code>sendBidiResponse</code> (Binding) で返却</p>
                <div style="text-align:center; color:var(--accent-color)">↓</div>
                <p><strong>4. VBA側:</strong> <code>TakeEvents</code> でメッセージを回収</p>
            </div>
        </div>
    </div>

    <div class="alert success">
        <div class="alert-content">
            <strong>結論：ポータビリティの極致へ</strong>
            <p>この方法の確立により、「Excelファイル1つとブラウザさえあれば、世界標準の次世代プロトコルをフルパワーで扱える」という環境が誕生しました。<code>BiDiPoc.bas</code> での生々しい実験こそが、技術的ブレイクスルーの瞬間でした。</p>
        </div>
    </div>
`;
