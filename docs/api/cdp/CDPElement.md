# CDPElement

DOM 要素ハンドルです。Playwright の Locator / ElementHandle に近い役割です。[`CDPContext`](./CDPContext) の `getElementBy*` から取得します。

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext("https://example.com")
t.getElementByQuery("button#go").click
```

## 操作

| メソッド | 説明 |
| --- | --- |
| `click` | クリック（任意で読み込み待ち） |
| `submit` | フォーム送信 |
| `sendString` | テキスト入力 |
| `sendClick` | クリック送信系 |
| `sendKey` | キーコード送信 |
| `clearValue` | 値クリア |
| `focus` / `selectText` | フォーカス・選択 |
| `fireEvent` | DOM イベント発火（React 等で有用） |
| `getAttribute` / `setAttribute` | 属性 |
| `setSelection` | `<select>` 選択 |

[入力とクリック](/guides/input)

## 存在確認

```vb
Public Function isExist() As Boolean
Public Function ifExist() As CDPElement
Public Function onExist(Optional timeOutInSeconds As Double = 30, _
    Optional raiseTimeoutError As Boolean = True) As CDPElement
Public Function onExistNot(Optional timeOutInSeconds As Double = 30) As Boolean
```

`onExist` は見つかるまで待ってからチェーンできます。

```vb
t.getElementByID("async-btn").onExist.click
```

## ツリー走査

| メソッド | 説明 |
| --- | --- |
| `getParent` | 親 |
| `getNextSibling` / `getPrevSibling` | 兄弟 |
| `getFirstChild` / `getChildren` | 子 |

## 入れ子検索

要素スコープでも Context と同様の `getElementByID` / `Query` / `XPath`（単数・複数）が使えます。

## Shadow / iframe

```vb
Public Function getIFrame() As CDPElement
Public Function GetShadowRoot(Optional Index As Long) As CDPElement
Public Function GetShadowRoots() As Collection
```

[要素の取得](/guides/selectors)

## その他

### `ExposeDevTools`

```vb
Public Sub ExposeDevTools(varName As String)
```

DevTools 上で要素を変数公開（デバッグ用）。

### `Init`

フレームワーク内部の初期化。ユーザーコードから直接呼ぶ必要は通常ありません。

## 関連

- [`CDPContext`](./CDPContext)
- [入力とクリック](/guides/input)
