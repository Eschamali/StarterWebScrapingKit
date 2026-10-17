---
description: WebDriver BiDi にも要素操作ヘルパーを。1つのクラスに統合しかけて、やっぱり分けた WebDriverBiDiElement 誕生の3日間と、sharedId と handle、Core を隠した拡張イベントの話です。
---

# BiDi要素ヘルパー、登場秘話

::: tip 登場秘話
〜1つにまとめようとして、やっぱり分けた。BiDi にも「要素」がやってきた3日間〜
:::

v4.0.0 で入口を作り直したあと、次に手を付けたのは、ずっと残っていた宿題でした。

**WebDriver BiDi には、要素操作のヘルパーがない。**

CDP には `CDPElement` があります。`getElementByQuery(...)` で要素を取って、`.click` や `.value = "..."` と書くだけです。一方 BiDi は、[BiDi 登場秘話](/stories/bidi-story)のとおり、通信の土台はできても、要素を触るには `jsEval` で JavaScript を書くか、`UpgradeBiDiPlus` で CDP に変換するしかありませんでした。

---

## まずは「結果」を正しく受け取る

要素ヘルパーを作るには、検索した要素を**参照として受け取る**必要があります。ところが BiDi の JavaScript の結果には、参照が2種類ありました。

- `script.SharedId` … DOM の要素（node）に付く ID。ブラウザ側で保持されないので、解放不要
- `script.Handle` … 配列やオブジェクトなど、要素以外を `resultOwnership: "root"` で受け取ったときの ID。使い終わったら解放（`script.disown`）が必要

値と `VarType` だけを返す今の `jsEval` では、返ってきた文字列がどちらの ID なのか、区別がつきません😇

そこで最初に作ったのが、JavaScript の実行結果を丸ごと記録する `JsEvalResult` です。

```text
2026-10-15 JavaScriptの実行結果を詳細に返す仕組みを導入
2026-10-16 jsevalのBiDiは、`script.SharedId`を最優先に返す
2026-10-16 `script.SharedId`ベースになり、ブラウザ側で保持されないため、解放不要
2026-10-16 複数形は`handle`取得に
```

`jsEval` に `returnByDetails:=True` を付けると、値の代わりにこの `JsEvalResult` が返ります。型・`sharedId`・`handle`・例外の行番号まで、まとめて見られるようになりました。単数の検索は `sharedId`、複数の検索はコレクションを `handle` で受けて1個ずつ `sharedId` に、という使い分けも、ここで決まりました。

---

## 1つにまとめようとした

次は要素のクラスです。最初は、`CDPElement` に BiDi も扱わせる方向で進めました。

```text
2026-10-15 BiDi判別サインを追加
2026-10-15 BiDi用要素検索プロシージャを用意
2026-10-16 CDP依存コマンドをBiDiコマンドでも行けるようにする
2026-10-16 CDP専用じゃなくなったので、リネーム
```

CDP 専用ではなくなったので、`CDPElement` という名前もやめて、`PageElement` に変えました。初期化は `InitCDP` と `InitBiDi` の2本立て。1つのクラスで両方を扱えれば、利用者は何も意識しなくて済みます。

……と思ったのですが。

---

## ……やっぱり分けます

```text
2026-10-17 ...やっぱり分けます
2026-10-17 統合のやつを削除
2026-10-17 `CDPElement.cls`の書き方に合わせる
2026-10-17 v4.0.0頃の命名に戻す
```

ただ、CDP と BiDi には、仕様そのものの差があります。CDP なら iframe を要素から辿れますが、BiDi では iframe は別の `browsingContext` です。Shadow DOM も、BiDi では要素ごとに1つずつしか取れません。`PageElement` には、「BiDi モードでは使えません」とエラーを出すメソッドが混ざることになりました。中で使う ID も、`objectId` と `sharedId` で別物です。

結局、統合したクラスは削除して、`CDPElement` は元の名前に戻し、BiDi 版は **`WebDriverBiDiElement`** として独立させました。名前こそ違いますが、呼び出すメソッドは `CDPElement` とほぼ同じです。BiDi でできない `getIFrame` と `GetShadowRoots` だけは、無理に用意していません。

```text
2026-10-17 無理にCDP要素を出さない
```

ついでに、`CDPElement` の `sendClick` も、BiDi に合わせて「マウスを動かしてから、押して、離す」という人間らしい流れに直しました。

```text
2026-10-16 `sendClick`を人間らしいフローに改良
```

---

## Core には、触らせない

最後にもう1つ。拡張用のクラスを作るとき、これまでは通信コア（`CDPCore`）を `ExEventCore` で取り出して、`WithEvents` で受け取っていました。

```text
2026-10-17 Core系のオブジェクトを取れないように調整
```

コアは、送受信と解析を担う心臓部です。設計思想として、利用者や拡張クラスからは、なるべく Core 系のクラスに触れないようにしたい。そこで、`ExEventCore` は廃止しました。代わりに `CDPBrowser` / `CDPContext` / `WebDriverBiDiMode` が、自分の処理を済ませたあとで、拡張用のイベント `ExEvent` を発信します。

```vb
Private WithEvents ex_CDPContext As CDPContext

Private Sub ex_CDPContext_ExEvent(methodName As String, RawJson As String)
    ' このタブ宛てのイベントだけが届く
End Sub
```

拡張クラスで用意するものは、`CDPCore` / `CDPBrowser` / `CDPContext` の3つから、`CDPContext`（必要なら `CDPBrowser`）だけになりました。

---

## 結論 — 似ているからこそ、分ける

CDP と BiDi は、どちらも「ブラウザを動かすプロトコル」で、やっていることはよく似ています。だからこそ1つにまとめたくなります。でも、仕様の差まで1つのクラスに押し込めると、「このモードでは使えません」が混ざったクラスになってしまいました。

::: tip フィナーレ
同じ名前のメソッドを持つ、別々のクラス。使う側から見ればほぼ同じ書き方で、中身はそれぞれのプロトコルに素直なまま。3日間、遠回りして辿り着いたのは、そんな形でした🤠
:::

## 次に読む

- [`WebDriverBiDiElement`](/api/bidi/WebDriverBiDiElement) — BiDi の要素ヘルパー
- [`JsEvalResult`](/api/cdp/JsEvalResult) — JavaScript の詳細な実行結果
- [イベント購読](/guides/events) — 拡張クラスでの `ExEvent` の受け取り方
- [BiDi 登場秘話](/stories/bidi-story)
