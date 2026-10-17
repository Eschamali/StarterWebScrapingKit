---
description: WebDriver BiDi で取得した DOM 要素を操作する WebDriverBiDiElement。CDPElement とほぼ同じメソッドで、クリック・入力・存在待ち・ツリー走査・Shadow DOM を扱います。
---

# WebDriverBiDiElement

WebDriver BiDi 版の DOM 要素ハンドルです。[`CDPElement`](/api/cdp/CDPElement) の BiDi 版で、**呼び出すメソッドはほぼ同じ**です。[`WebDriverBiDiContext`](./WebDriverBiDiContext) の `getElementBy*` から取得します。

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext("https://example.com")

t.getElementByQuery("input[name='q']").value = "検索語"
t.getElementByQuery("button#go").click
```

要素は、BiDi の `script.SharedId`（要素の共有 ID）で参照しています。ブラウザ側で保持されるオブジェクトではないため、解放処理は不要です。

## 取得方法

[`WebDriverBiDiContext`](./WebDriverBiDiContext) に、次の検索メソッドがあります。

| メソッド | 戻り値 |
| --- | --- |
| `getElementByID(strID As String)` | `WebDriverBiDiElement` |
| `getElementByQuery(strQuery As String)` | `WebDriverBiDiElement` |
| `getElementsByQuery(strQuery As String)` | `Collection`（1始まり） |
| `getElementByXPath(strXPath As String)` | `WebDriverBiDiElement` |
| `getElementsByXPath(strXPath As String)` | `Collection`（1始まり） |

見つからない場合でも、単数系は `Nothing` ではなく「存在しない要素」を返します。[`isExist`](#存在確認) で判定してください。複数系は、検索自体に失敗したときに `Nothing` を返します。

## メソッド一覧

使い方・引数は [`CDPElement`](/api/cdp/CDPElement) と同じです。各メソッドの詳しい説明は、CDPElement の同名の節を参照してください。

| 分類 | メンバー |
| --- | --- |
| 値・状態 | `value` / `innerText` / `innerHTML` / `checked` / `selected` |
| 操作 | `click` / `SimpleClick` / `submit` / `sendString` / `sendClick` / `sendHover` / `sendKey` / `clearValue` / `focus` / `selectText` / `fireEvent` / `getAttribute` / `setAttribute` / `setSelection` / `SetFileInputFiles` |
| 存在確認 | `isExist` / `ifExist` / `onExist` / `onExistNot` |
| ツリー走査 | `getParent` / `getNextSibling` / `getPrevSibling` / `getFirstChild` / `getChildren` |
| 入れ子検索 | `getElementByID` / `getElementByQuery` / `getElementsByQuery` / `getElementByXPath` / `getElementsByXPath` |
| Shadow DOM | `GetShadowRoot` |
| その他 | `jsEval` / `ExposeDevTools` / `CurrentObjectId` / `CurrentNodeType` / `SetOptionStopException` / `SetOptionRunAsyncBiDi` / `SetOptionUserGesture` |

`sendClick` / `sendHover` / `sendKey` は、BiDi の `input.performActions` で、要素を起点にした実際のマウス・キー操作を送ります。

## CDPElement との違い

| 項目 | CDPElement | WebDriverBiDiElement |
| --- | --- | --- |
| 要素の参照 | `Runtime.RemoteObjectId`（`CurrentObjectId`） | `script.SharedId`（`CurrentObjectId` で取得。名前は CDP 版と揃えています） |
| iframe | `getIFrame` | **なし**。iframe は独立した `browsingContext` なので、[`WebDriverBiDiMode.getTab`](./WebDriverBiDiMode) の `maxDepth:=1` で取得します |
| Shadow DOM | `GetShadowRoot(Optional Index)` / `GetShadowRoots` | `GetShadowRoot()` のみ（Index 引数なし）。1段ずつ辿ります |
| `jsEval` の引数 | `returnByValue` / `generatePreview` あり | `returnByValue` の代わりに `Ownership`、`generatePreview` はなし |
| 非同期実行オプション | `SetOptionRunAsyncCDP` | `SetOptionRunAsyncBiDi` |
| 後始末 | オブジェクトの解放あり | 解放不要 |

`keyboardCode`（`sendKey` のキー）と `NodeNodeType`（`CurrentNodeType` の値）の列挙型は、`CDPElement.cls` 側で定義しています。両方のクラスをインポートしてください。

## Shadow DOM

`GetShadowRoot` で、その要素の直下の Shadow Root を取得します。open / closed のどちらにも対応しています。入れ子の Shadow DOM は、ホスト要素 → `GetShadowRoot` → 検索、を1段ずつ繰り返します。

```vb
Function GetShadowRoot() As WebDriverBiDiElement
```

```vb
Dim host As WebDriverBiDiElement, root As WebDriverBiDiElement
Set host = t.getElementByQuery("my-widget")
Set root = host.GetShadowRoot
If Not root Is Nothing Then root.getElementByQuery("button").click
```

- 取得できない場合は `Nothing` を返します
- BiDi では、要素ごとに Shadow Root は1つです。そのため `GetShadowRoots`（複数取得）はありません。実務では、1段ずつで十分足ります

## iframe

iframe 内の要素は、iframe 自体を別のタブ（`browsingContext`）として取得してから検索します。

```vb
Dim frame As WebDriverBiDiContext
Set frame = t.ThisWebDriverBiDiMode.getTab("https://example.com/frame.html", maxDepth:=1)
frame.getElementByQuery("#inner").click
```

## `onExist` / `onExistNot` が使えない要素

`onExist` / `onExistNot` は、要素を見つけたときの検索を再実行して待ちます。次の要素は検索の情報を持たないため、待機できません。

- `getElementsByQuery` / `getElementsByXPath` で取得した、コレクションの各要素
- `GetShadowRoot` で取得した Shadow Root

## `jsEval`

要素を `this` にして JavaScript を実行します。引数は [`script.LocalValue`](https://w3c.github.io/webdriver-bidi/#type-script-LocalValue) 形式（`{"type":…,"value":…}`）で渡します。

```vb
Function jsEval(JavaScriptStr As String, Optional objectArguments As Variant, Optional IFEXCEPTION As Variant, _
    Optional Ownership As Boolean, Optional awaitPromise As Boolean, Optional serializationOptions As Dictionary, _
    Optional returnByDetails As Boolean, Optional StopApiError As Boolean = True) As Variant
```

```vb
Debug.Print t.getElementByQuery("h1").jsEval("function(){ return this.textContent }")
```

`returnByDetails:=True` にすると、結果の詳細を [`JsEvalResult`](/api/cdp/JsEvalResult) で受け取れます（`Set` で受けてください）。

## 関連

- [`WebDriverBiDiContext`](./WebDriverBiDiContext) — 要素の取得元
- [`CDPElement`](/api/cdp/CDPElement) — 各メソッドの詳しい説明
- [`JsEvalResult`](/api/cdp/JsEvalResult) — JavaScript の詳細な実行結果
- [要素の取得](/guides/selectors)
