---
description: テスト網羅・エラー分類・型安全性・マルチブラウザ対応など、コアロジック以外に残る差分を整理し、「投資で埋まる差」と「言語の制約で埋まらない差」に切り分けます。
---

# 残る差分と、埋まらない差

ここまでの4ページで、**コアロジック（通信・振り分け・非同期・クラス構成）については構造的にかなり近い**ことを見てきました。

しかし当然ながら、Puppeteer / Playwright とキットの間には大きな差があります。このページではその差を洗い出し、 **「実装投資で埋まる差」と「言語の制約で原理的に埋まらない差」** に分類します。

## 1. テスト ―― 差が最も大きい領域

| | 自動テスト |
| --- | ---: |
| **Playwright** | `*.spec.ts` 約 556 ファイル（`tests/` 配下だけで約 575 の TS ファイル） |
| **Puppeteer** | `*.test.ts` 約 138 ファイル + `test/` 配下 約 91 ファイル |
| **StarterWebScrapingKit** | `ForDevelopers/OperationCheck/` に 8 モジュール（アサーション約 78 件） |

Playwright は3ブラウザエンジン × 数百本のテストを CI で常時回しています。「Chrome の更新でこの操作が壊れた」を機械的に検出できる体制です。

### キット側にもテストはある

数こそ少ないものの、`dev` ブランチの [`ForDevelopers/OperationCheck/`](https://github.com/Eschamali/StarterWebScrapingKit/tree/dev/ForDevelopers/OperationCheck) には検証コード一式が置かれています。

| モジュール | 規模 | 内容 |
| --- | --- | --- |
| `CDP/TestVBA/Test_CDPElement.bas` | 510 行 / 17 テスト / 56 アサーション | `value` / `innerText` / `checked` / Shadow DOM / iframe / ファイル入力など要素操作の全面検証 |
| `CDP/TestVBA/Test_jsEval.bas` | 618 行 / 16 テスト | `Runtime.evaluate` と `callFunctionOn`、`awaitPromise`、Unicode、JS 例外処理 |
| `CDP/TestVBA/Test_AsyncBenchmark.bas` | 479 行 | マルチタブ非同期ラウンドのレイテンシ計測（インライン方式と拡張クラス方式の2パターン） |
| `WebDriverBiDi/TestVBA/Test_BiDiAlertRace.bas` | 218 行 | BiDi でのアラート競合という再現しにくい条件の検証 |
| `common/RPAChallenge/Test_RPAChallenge.bas` | 74 行 | rpachallenge.com を使った E2E |

しかも「実行して目視」ではなく、**アサーションと PASS / FAIL 集計を持つ手作りのテストハーネス**になっています。

```vb
el.value = "Hello VBA"
AssertEq "value(LetとGet)", el.value, "Hello VBA"

el.sendString "Real Key Input"
AssertEq "sendString後のvalue", el.value, "Real Key Input"

el.clearValue
AssertEq "clearValue後のvalue", el.value, ""
```

```vb
PrintHeader "テスト完了: PASS=" & passCount & " / FAIL=" & failCount & _
            " / 合計=" & (passCount + failCount)
```

検証対象の HTML フィクスチャ（`TestHtml/` 配下）も同梱されており、外部サイトの変更に左右されずにローカルファイルで再現できるようになっています。結果は `jsEval` でページ側にも書き戻され、ブラウザ上で PASS / FAIL が視覚的に確認できます。

これとは別に、配布物に同梱の `Demo_CDP` / `Demo_WebDriverBiDi` / `Demo_WebSocket` が使用例兼スモークテストとして機能しています。

### それでも残る差

| 観点 | Playwright / Puppeteer | StarterWebScrapingKit |
| --- | --- | --- |
| テストの形式 | アサーションベース | ✅ アサーションベース（自作ハーネス） |
| フィクスチャの同梱 | ✅ | ✅ `TestHtml/` |
| カバー範囲 | ほぼ全 API | `CDPElement` / `jsEval` が中心。`CDPCore` のバッファ管理やディスパッチ、`CDPBrowser`、BiDi の大半は未カバー |
| 実行 | CI で自動 | 人間が VBE から `RunAll_○○_Tests` を叩く |
| リグレッション検出 | プルリクごとに自動 | 気づいた人が実行したときだけ |

つまり差は「テストがあるか無いか」ではなく、**カバー範囲と、自動で回る仕組みがあるかどうか**です。

::: warning ここは正直に劣る点
一番テストが欲しいのは、皮肉にも[このコーナーで一番自信を持って紹介したバッファ管理やディスパッチ](/core-comparison/transport)の層です。要素操作という「壊れたらすぐ気づける層」は固めてある一方、コアの回帰検出は手薄なままです。

そして VBA には CI で回す標準的な手段がありません。ブラウザとの実通信が前提なので、そもそもヘッドレス環境での自動実行と相性が悪いという事情もあります。とはいえ Rubberduck のようなユニットテスト基盤は存在するので、**これは言語の制約ではなく投資量の問題**です。
:::

## 2. エラー処理 ―― 分類の粒度

Puppeteer は継承関係を持つエラー階層を定義しています。

```ts
// puppeteer-core/src/common/Errors.ts
export class PuppeteerError extends Error {}
export class TimeoutError extends PuppeteerError {}
export class TouchError extends PuppeteerError {}
export class ProtocolError extends PuppeteerError {}
export class UnsupportedOperation extends PuppeteerError {}
export class TargetCloseError extends ProtocolError {}
export class ConnectionClosedError extends ProtocolError {}
```

`catch (e) { if (e instanceof TimeoutError) ... }` のように、**型で分岐して回復処理を書けます**。Playwright も `TimeoutError` / `TargetClosedError` / `AbortError` / `ProtocolError` などを持ちます。

キット側は数値のエラーコードです。

```vb
Public Enum CDPCustomErrorCodes
    TIMEOUT = 900   '時間内に、CDPから応答がありませんでした
    PIPE = 901      'Pipe通信周り(`PeekNamedPipe`など)でエラー
    Protocol = 902  'CDPコマンド自体のエラー
End Enum
```

VBA の `Err.Raise` には例外クラスという概念がないため、`On Error` + `Err.Number` で分岐します。**分類の粒度は粗いものの、「タイムアウト / 通信 / プロトコル」という最重要の3分類は押さえてあります**。

一方で、ログ出力はむしろ整備されています。

```vb
Public Enum LogLevelName
    Trace_
    Debug_
    info_
    WARN_
    ERROR_
End Enum
```

`printMsg` が全クラスに行き渡っており、`FromProcedureName` で発生箇所も記録されます。Node 側が `debug` パッケージで行っていることと役割は同じです。

| | エラー分類 | ログ | 回復処理 |
| --- | --- | --- | --- |
| **Puppeteer / Playwright** | 型階層で細分化 | `debug` / `debugLogger` | `instanceof` で分岐 |
| **StarterWebScrapingKit** | 数値コード3種 | 5段階のレベル付きログ | `Err.Number` で分岐 |

## 3. 型安全性 ―― CDP メソッド名の扱い

Puppeteer / Playwright は CDP の JSON スキーマから `protocol.d.ts` を生成しており、**存在しないメソッド名やパラメータの誤りがコンパイル時に落ちます**。

```ts
await session.send('Page.navigate', { url });   // メソッド名も params も型チェック済み
```

キットでは文字列と `Dictionary` です。

```vb
tab.ExecuteCDP "Page.navigate", params   ' タイポは実行時まで分からない
```

これも原理的に埋まらない差ではありません。CDP のスキーマから VBA の `Enum` や定数モジュールを生成すれば、少なくともメソッド名のタイポは防げます。ただし `params` の構造まで型で縛るのは、`Dictionary` ベースである以上かなり難しくなります。

なお `WithEvents` によるイベントハンドラのシグネチャ検証（[ディスパッチのページ](/core-comparison/dispatch)参照）のように、**VBA でも型チェックが効いている箇所はあります**。型安全性がゼロというわけではなく、CDP コマンドの層に穴がある、というのが正確です。

## 4. マルチブラウザ対応

| | 対応エンジン |
| --- | --- |
| **Playwright** | Chromium / Firefox / WebKit（＋ Android / Electron） |
| **Puppeteer** | Chromium（＋ Firefox を BiDi 経由で） |
| **StarterWebScrapingKit** | Chromium 系のみ（Chrome / Edge） |

Playwright は Firefox と WebKit に**独自のパッチを当てたビルドを配布**することでこれを実現しています。ブラウザバイナリのメンテナンスを含む取り組みであり、個人開発で追随できる範囲を超えています。

キットは BiDi スタックを持っていますが、これは `mapperTab.js`（chromium-bidi）を CDP の上に載せる方式であり、あくまで Chromium 上での BiDi です。Firefox の BiDi エンドポイントへ直接繋ぐには、WebSocket 経路の成熟が前提になります（[Transport のページ](/core-comparison/transport)参照）。

## 5. エコシステム ―― 差の原因が逆転する領域

Node 側では、以下はすべて既製品を `import` するだけです。

| 必要な機能 | Node | StarterWebScrapingKit |
| --- | --- | --- |
| WebSocket | `ws` パッケージ | `CDPCoreViaWebSocket`（1,335行、WinSock から自作） |
| JSON | `JSON.parse` / `stringify`（ネイティブ） | `BiDiCDPJson`（2,998行、自作パーサ） |
| UTF-8 変換 | `Buffer` | `CharacterCodeConversion`（284行） |
| 非同期 I/O | libuv | `PeekNamedPipe` / `ioctlsocket` を直接呼ぶ |
| Windows API エラー | 不要 | `WinApiError` |

**この表は「キットが劣っている」ことを示していません。** むしろ逆で、Node 側が数行で済ませている部分に、キットは 4,000 行以上を投じて同等の機能を用意しているということです。

同時に、この自作分がそのまま**保守負担**でもあります。`ws` パッケージは世界中で使われ続けてバグが潰されていますが、`CDPCoreViaWebSocket` の品質を担保するのは作者ひとりです。

## 6. キット側にしかない強み

比較の公平のために、逆方向も挙げておきます。

- **Node.js/WebDriver.exe のインストールが不要** — Excel があれば動く。ソフトウェア導入が制限された環境で決定的
- **Excel との一体化** — スクレイピング結果をそのままセルへ。データの受け渡し層が存在しない
- **既存ブラウザへの再接続（reattach）** — 手動でログイン済みのブラウザを掴んで自動化を継続できる（[再接続](/guides/reattach)）
- **同期実行による記述の素直さ** — `await` が無い分、上から下に読める（[非同期のページ](/core-comparison/async)）
- **VBE デバッガでのステップ実行** — ブレークポイントを置いて CDP の生 JSON をイミディエイトで覗ける

## 7. 差分の総括

### 実装投資で埋まる差

| 項目 | 必要なもの |
| --- | --- |
| 自動テストの拡充 | `OperationCheck` のカバー範囲拡大 + 自動実行の仕組み |
| エラー分類の細分化 | エラーコードの追加と `On Error` 分岐の整理 |
| CDP メソッド名の型付け | スキーマからの `Enum` 生成 |
| Handle の明示的解放 | `Runtime.releaseObject` の呼び出し設計 |
| CDP / BiDi の抽象化 | `Implements` による共通インターフェース |
| WebSocket の堅牢化 | Ping/Pong 応答、ハンドシェイク検証の厳密化 |

### 言語・環境の制約で埋まらない差

| 項目 | 理由 |
| --- | --- |
| イベントループによる自動的な受信駆動 | VBA に存在しない |
| `async` / `await` の記述性 | 構文が存在しない |
| マルチブラウザエンジン対応 | ブラウザバイナリのメンテナンスが必要 |
| パッケージマネージャによる配布・更新 | `.cls` の手動インポートが前提 |
| クロスプラットフォーム | Windows API に直接依存 |
| 大規模なクラス分割 | 名前空間・フォルダ・`import` が無い |

## 結論

**コアロジックそのものは、驚くほど近い。** メッセージの切り出し方、`method` / `id` / `sessionId` による振り分け、セッション多重化、コマンド ID と応答の待ち合わせ、三層のオブジェクトモデル ―― これらは3者ともほぼ同じ解に到達しています。CDP という共通のプロトコルを相手にする以上、行き着く先が同じになるのは自然なことです。

差が生まれているのは、その周辺です。テスト・型・エラー分類・マルチブラウザ・エコシステム。これらの多くは**言語の限界ではなく、投じられたリソースの差**です。そして残る一部（イベントループ、`await`、配布形態）だけが、VBA という選択に付随して動かせない制約です。

つまり **「VBA だからここまでしかできない」という部分は、思っていたより小さい**というのが、このコーナーの結論になります。

## 次に読む

- [概要と結論](/core-comparison/) — 比較全体のサマリ
- [設計思想](/concepts/design-philosophy) — なぜこのスコープを選んだのか
- [BiDi 対応の物語](/stories/bidi-story) — 実際にどう作られたか
