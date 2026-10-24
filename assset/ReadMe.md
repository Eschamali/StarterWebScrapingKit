# このフォルダの役割

Chrome for Testing（CfT）のチャンネル（Stable / Beta / Dev / Canary）に対応する、chromium-bidiの `mapperTab.js` を取得するツールです。

## これまでとの違い

| | これまで | これから |
|---|---|---|
| バージョンの決め方 | GitHubのリリースタグ（`chromium-bidi-vX.Y.Z`）を追う | [CfTのAPI](https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions.json) の `channels.<CHANNEL>.version` から決める |
| Chromeとの対応 | タグとChromeの対応は、`.browser` に書かれた検証用のChromeだけ | **Chromeのビルドとchromium-bidiの版が1対1で対応する** |
| 取得元 | npmの `lib/iife/mapperTab.js` | npm（unpkg経由）の `out/Default/gen/src/mapperTab.js` |

GitHubのリリースタグは `v17.0.2`（2026-07-06）で止まっていて、今後は更新されません。CfTのAPIを使うと、**使うChromeと同じビルドから作られた `mapperTab.js`** を取れるので、バージョンの食い違いが起きません。

## 背景：なぜnpmのバージョンが `158.0.8110-0` になったのか

### Chromiumへの移行

| 日付 | 出来事 |
|---|---|
| 2026-07-06 | release-pleaseで `v17.0.2` をリリース。セマンティックバージョニングでのリリースはこれが最後 |
| 2026-08-24 | `fc7df403` で `.github/workflows` やrelease-pleaseの設定を削除。ソースの本体が `chromium/src/third_party/chromium-bidi` に移る |
| 2026-08-27 | Copybaraの設定を追加。GitHubのリポジトリはChromiumからの一方向ミラーになる |
| 2026-09-28 | `0188a647` で `publish.yml` を削除 |
| 2026-09-30 | npmに初めてChromeのバージョン番号の `157.0.8080-0` が公開される |
| 2026-10-14 | `1c6f4363` でREADMEにnpmへの公開手順が書かれる |

### npmへの公開の仕組み（chromium-bidiのREADME「Publish new `npm` release」より）

- CfTの公式ビルダー `linux64-cft` が、**Chromeのリリースごとに1回** publishする
- `package.json` の `version`（`17.0.2` のまま）は使われない
- **Chromeのバージョンの最後の `.` を `-` に変えたもの**がnpmのバージョンになる
  - 例：Chrome `157.0.8103.6` → npm `157.0.8103-6`
- 毎回 `latest` タグが付く

### 番号の読み方

- `XXX.0.YYYY-0`（末尾が `-0` で、YYYYが毎日増える）：trunkのビルド。Canary / Dev相当
- `XXX.0.YYYY-N`（YYYYが固定で、Nだけ増える）：リリースブランチのビルド。Beta、いずれStable相当

## npmの `latest` が当てにならない理由

- trunkとリリースブランチのビルドが同じパイプラインから順番に出てきて、そのたびに `latest` が付け替えられる
- そのため、`latest` は「最新」でも「安定版」でもなく、 **「最後にCIを通ったもの」** になる。番号が戻ることもある

```
10/15 15:39  158.0.8107-0  ← latest（Canary）
10/16 02:26  157.0.8103-8  ← latest（Beta。番号が戻る）
10/16 04:22  158.0.8108-0  ← latest（Canary）
```

- プレリリース扱いなので、`^` や `~` の範囲指定も使えない
- `stable` / `beta` のようなdist-tagはない

## Puppeteerからたどる方法（つなぎ）

自動publishは `157.0.8080-0` から始まったので、M155とM156の版はnpmに**存在しません**。その間は、Puppeteerが固定している版を「事実上の安定版」として使います。

```bash
npm view puppeteer-core dependencies.chromium-bidi
```

2026-10-19時点では、`puppeteer-core@25.14.0` が以下の組み合わせになっています。

- 同梱しているChrome：`155.0.8059.84`（Stable）
- 依存しているchromium-bidi：`157.0.8090-0`

npmのバージョン別のダウンロード数を見ても、`157.0.8090-0` だけが突出しています（週140万回）。

## 使い方

### メニューから選ぶ

`track-bidi.bat` をダブルクリックするか、引数なしで実行します。

```
  1. Stable
  2. Beta
  3. Dev
  4. Canary
Channel [1-4]: 3
Chrome Dev 158.0.8106.0 / chromium-bidi 158.0.8106-0 (cft)
Updated:  -> 158.0.8106-0
Press any key to continue . . .
```

数字キーを押すだけで決まります（Enterは要りません）。終わったら `pause` で止まります。

### 引数で指定する

```bat
track-bidi.bat Stable
```

メニューも `pause` も出ません。タスクスケジューラで定期的に実行する場合はこちらを使います。

## 処理の流れ

1. CfTのAPIから、選んだチャンネルのChromeのバージョンを取得する
2. 最後の `.` を `-` に変えて、npmのバージョンにする
3. `https://registry.npmjs.org/chromium-bidi/<版>` でnpmにあるか確かめる
   - **ない場合**は、最新の `puppeteer-core` が依存している版を代わりに使う
4. 前回と同じ版なら何もしない
5. `https://unpkg.com/chromium-bidi@<版>/out/Default/gen/src/mapperTab.js` を保存して、記録する

## 保存されるファイル（batと同じフォルダ）

| ファイル | 中身 |
|---|---|
| `<CHANNEL>-version.txt` | そのチャンネルで前回取得した版（例：`Stable-version.txt`） |
| `history.log` | 版が変わった日時、チャンネル、Chromeとchromium-bidiの版、取得元 |
| `mapperTab.js` | その版の `mapperTab.js` 本体（約1MB） Gitの差分確認をしやすくするため、同名として保存します |

`history.log` の例です。最後の括弧が取得元で、`(cft)` ならCfTから直接、`(puppeteer-core …)` ならPuppeteerの版で代用したことを表します。

```
2026/10/19 17:30:00 Stable Chrome 155.0.8059.84 chromium-bidi 157.0.8090-0 (puppeteer-core 25.14.0)
2026/10/19 17:30:07 Beta Chrome 157.0.8103.6 chromium-bidi 157.0.8103-6 (cft)
```

## いつから恩恵があるか

ChromiumDashのリリース日程（2026-10-19時点）です。今はStableが2週間ごとに上がっています。

| マイルストーン | ブランチを切った日 | Stableになる日 | npmの版 |
|---|---|---|---|
| M156 | 2026-09-28 | 2026-10-20 | なし（Puppeteerの版で代用） |
| M157 | 2026-10-12 | 2026-11-03 | `157.0.8103-N` |

**2026-11-03以降**にStableで実行して、`history.log` の最後が `(cft)` になっていれば、CfTから直接取れるようになっています。

## 補足

### BiDiを試すだけならnpmは要らない

ChromeDriverには、同じバージョンのChromeに合わせたMapperがビルド時に同梱されています。WebDriver BiDiを使うだけなら、ChromeとChromeDriverのバージョンをそろえれば十分です。

```bash
npx @puppeteer/browsers install chrome@stable chromedriver@stable
```

npmのchromium-bidiは、Puppeteerのように自分でMapperを注入するライブラリ向けの部品です。

### ChromeはネイティブでBiDiを話せない

FirefoxはブラウザがWebDriver BiDiを直接話せますが、Chromeが直接話せるのはCDP（Chrome DevTools Protocol）だけです。ChromeのBiDiは、CDPの上に `mapperTab.js` という翻訳層をかぶせて実現しています。

| | Firefox | Chrome |
|---|---|---|
| BiDiの実装場所 | ブラウザ本体（Remote Agent） | ブラウザの外。翻訳層の `mapperTab.js` |
| ブラウザ本体が話すプロトコル | WebDriver BiDi（CDP対応はすでに廃止） | CDP |
| BiDiの入口 | `firefox --remote-debugging-port=9222` → `ws://127.0.0.1:9222/session` に直接つなげる | ChromeDriverかPuppeteerを経由する |

chromium-bidiのREADME（「How does it work?」）によると、Chromeでは次のように動いています。

```
クライアント
   │  WebDriver BiDi（WebSocket）
   ▼
ChromeDriver（C++）  ← BiDiのWebSocketを受け付ける
   │  CDP
   ▼
Chrome
 └─ 隠しタブで mapperTab.js が動く  ← BiDiをCDPに翻訳する
        │  CDP
        ▼
     ページやターゲットを操作
```

- **ChromeDriver**：BiDiのWebSocketとHTTPの入口を持ち、ブラウザを起動・管理する
- **mapperTab.js**：Chromeの中の隠しタブで動き、BiDiのコマンドをCDPのコマンドに翻訳する
- **Puppeteer**：ChromeDriverを使わず、自分でCDPを通して `mapperTab.js` を注入する。npmのchromium-bidiが必要なのはこのため

`chrome --remote-debugging-port=9222` で開くポートはCDP専用なので、Firefoxのように直接BiDiでつなぐことはできません。

ChromeにBiDiをネイティブで搭載するという話は、2026-10-19時点のchromium-bidiのリポジトリやREADMEには見当たりませんでした。

より詳細を知りたい方は[こちらの開発秘話](https://eschamali.github.io/StarterWebScrapingKit/stories/DevHistory/bidi-story)をご覧ください。  
動画付です🤠

### ソースコードをChromeのバージョンに合わせて読む

Chromiumに移ってからのコミットには、`Cr-Commit-Position: refs/heads/main@{#NNNNNNN}` が付いています。CfTのAPIの `revision` 以下で一番新しいコミットを探せば、そのChromeに入っているMapperのコードがわかります。

```bash
git log --format="%h|%ad|%s|%(trailers:key=Cr-Commit-Position,valueonly,separator=)" --date=short | awk -F"|" '{n=$4; sub(/.*#/,"",n); sub(/}.*/,"",n); if (n+0>0 && n+0<=1697595){print; exit}}'
```

Chrome 155 Stable（revision `1697595`）の場合、結果は `c372e57a`（2026-09-13）でした。

ただし、GitHubのミラーはChromiumの `main` だけを同期しているので、リリースブランチに後から入った修正（cherry-pick）は反映されません。

### batの文字コード

batの中は英語だけで書いています。日本語を書くと、cmdの文字コード（CP932）とファイルの文字コード（UTF-8）が食い違い、文字化けや構文エラーの原因になるためです。

### 運用ルール

git追跡対象は`Stable`のみとする
