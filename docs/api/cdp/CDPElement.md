---
description: クリック・入力・属性・存在待ち・ツリー走査・Shadow/iframe など、DOM 要素ハンドルの操作メソッドを解説します。
---

# CDPElement

DOM 要素ハンドルです。Playwright の Locator / ElementHandle に近い役割です。[`CDPContext`](./CDPContext) の `getElementBy*` から取得します。

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext("https://example.com")
t.getElementByQuery("button#go").click
```

詳細なセレクタ戦略は [要素の取得](/guides/selectors)、入力パターンは [入力とクリック](/guides/input) を参照してください。

## 値・状態

### `value`

```vb
Property Get value() As String
Property Let value(strTextVal As String)
```

要素の `value` です。代入時は React 系フィールドも考慮して書き込み、続けて `input` イベントを発火します。

```vb
Dim box As CDPElement
Set box = t.getElementByQuery("input[name='q']")
Debug.Print box.value
box.value = "検索語"
```

### `innerText` / `innerHTML`

```vb
Property Get innerText() As String
Property Let innerText(strTextVal As String)

Property Get innerHTML() As String
Property Let innerHTML(strTextVal As String)
```

表示テキスト／内側 HTML の取得・書き換えです。

```vb
Debug.Print el.innerText
el.innerHTML = "<span>更新</span>"
```

### `checked`

```vb
Property Get checked() As Boolean
Property Let checked(boolChecked As Boolean)
```

チェックボックス／ラジオのオンオフです。

```vb
t.getElementByQuery("input[type='checkbox']").checked = True
```

### `selected`

```vb
Property Get selected() As String
Property Let selected(selectedOption As String)
```

`<select>` の選択状態です。取得は先頭の `selectedOptions[0]`、代入はオプション値の設定です。選択肢の切り替えには [`setSelection`](#setselection) も使えます。

```vb
Debug.Print t.getElementByQuery("select#country").selected
```

## 操作

### `click`

```vb
Public Function click(Optional till As ReadyState = isComplete) As Boolean
```

要素をクリックします（スクロールインビュー → クリック → 待機）。

| 引数 | 意味 |
| --- | --- |
| `till` | クリック後に待つ [`ReadyState`](./CDPContext#readystate)。既定は `isComplete` |

```vb
t.getElementByID("submit").click
t.getElementByID("submit").click isInteractive   ' 待ちを短縮
```

### `submit`

```vb
Public Function submit(Optional till As ReadyState = isComplete) As Boolean
```

所属フォームを送信します（`this.form.submit()`）。

| 引数 | 意味 |
| --- | --- |
| `till` | 送信後に待つ ReadyState。既定は `isComplete` |

```vb
t.getElementByQuery("form").submit
```

### `sendString`

```vb
Public Function sendString(textToSend As String) As Boolean
```

値をクリアしたうえで、CDP の `Input.insertText` でテキストを送ります。日本語・絵文字も送れます（UTF-8 送信設定を推奨）。

| 引数 | 意味 |
| --- | --- |
| `textToSend` | 入力する文字列 |

```vb
box.clearValue
box.sendString "検索キーワード"
```

デモ: `Demo_CDP.JapaneseElementTest`

### `sendClick`

```vb
Public Function sendClick() As Boolean
```

座標ベースの物理クリック相当です（`Input.dispatchMouseEvent`）。JS の `dispatchEvent` では反応しない UI 向け。  
中心としてクリックします。

```vb
t.getElementByQuery("canvas#map").sendClick
```

### `sendKey`

```vb
Public Function sendKey(Key As keyboardCode, Optional altKey As Boolean = False) As Boolean
```

仮想キーコードを 1 回送ります（`Input.dispatchKeyEvent`）。

| 引数 | 意味 |
| --- | --- |
| `Key` | `keyboardCode` 列挙（下記） |
| `altKey` | `True` で Alt 修飾付き |

| `keyboardCode` | 意味 |
| --- | --- |
| `keyEnter` | Enter |
| `keyTab` | Tab |
| `keyEsc` | Esc |
| `keyBackspace` | Backspace |
| `keyDelete` | Delete |

```vb
box.sendKey keyEnter
box.sendKey keyTab
```

### `clearValue`

```vb
Public Function clearValue() As Boolean
```

入力欄の値を空にします。React フィールドも考慮し、最後に `input` イベントを発火します。

```vb
box.clearValue
box.sendString "新しい値"
```

### `focus` / `selectText`

```vb
Public Function focus()
Public Function selectText()
```

フォーカス付与／要素内テキスト全選択です。

```vb
box.focus
box.selectText
```

### `fireEvent`

```vb
Public Function fireEvent(strEventName As String, Optional till As ReadyState = isComplete) As Boolean
```

DOM イベントを発火します。React 等で `setAttribute` だけでは状態が同期されないときに使います。名前に `on` を付けても自動で除去されます（`"onchange"` → `"change"`）。

| 引数 | 意味 |
| --- | --- |
| `strEventName` | イベント名（`"input"` / `"change"` / `"blur"` など） |
| `till` | 発火後に待つ ReadyState。既定は `isComplete` |

```vb
el.focus
el.sendString "value"
el.fireEvent "input"
```

デモ: `Demo_CDP.fillReactForm`

### `getAttribute` / `setAttribute`

```vb
Public Function getAttribute(strAttributeName As String) As String
Public Function setAttribute(strAttributeName As String, strValue As String)
```

属性の取得／設定です。

| 引数 | 意味 |
| --- | --- |
| `strAttributeName` | 属性名（例: `"href"` / `"data-x"`） |
| `strValue` | 設定する値（`setAttribute` のみ） |

```vb
Debug.Print el.getAttribute("href")
el.setAttribute "data-x", "1"
```

### `setSelection`

```vb
Public Function setSelection(strOptionName As String)
```

`<select>` でオプションを選び、`change` を発火します。

| 引数 | 意味 |
| --- | --- |
| `strOptionName` | 選択する option の value |

```vb
t.getElementByQuery("select#country").setSelection "JP"
```

## 存在確認

### `isExist`

```vb
Public Function isExist() As Boolean
```

要素が取れているか（内部の `objectId` 有無）です。

```vb
If t.getElementByQuery(".toast").isExist Then
    Debug.Print "shown"
End If
```

### `ifExist`

```vb
Public Function ifExist() As CDPElement
```

チェーン用の存在ガードです。存在しなければ、続く操作はスキップされます。

```vb
' if el.isExist Then el.focus の代わり
t.getElementByQuery(".optional").ifExist.focus
```

### `onExist`

```vb
Public Function onExist(Optional timeOutInSeconds As Double = 30, _
    Optional raiseTimeoutError As Boolean = True) As CDPElement
```

要素が現れるまでポーリングし、見つかったら自身を返してチェーンできます。

| 引数 | 意味 |
| --- | --- |
| `timeOutInSeconds` | 待ち上限秒。既定は `30` |
| `raiseTimeoutError` | タイムアウト時にエラー停止するか。既定は `True` |

```vb
t.getElementByID("async-btn").onExist.click
t.getElementByID("ready").onExist(10, False).click   ' 10 秒・タイムアウトでも止めない
```

### `onExistNot`

```vb
Public Function onExistNot(Optional timeOutInSeconds As Double = 30) As Boolean
```

要素が消えるまで待ちます。消えたら `True`、上限まで残っていたら `False` です。

| 引数 | 意味 |
| --- | --- |
| `timeOutInSeconds` | 待ち上限秒。既定は `30` |

```vb
If t.getElementByQuery(".spinner").onExistNot(15) Then
    Debug.Print "読み込み完了"
End If
```

## ツリー走査

いずれも**現在の要素を基準**に相対移動します。見つからない場合は空の `CDPElement`（または `Nothing`）になります。

### `getParent`

```vb
Public Function getParent() As CDPElement
```

親要素（`parentElement`）です。

```vb
Set el = t.getElementByQuery("span.label").getParent
```

### `getNextSibling` / `getPrevSibling`

```vb
Public Function getNextSibling() As CDPElement
Public Function getPrevSibling() As CDPElement
```

次／前の兄弟要素です。

```vb
Set nextEl = el.getNextSibling
Set prevEl = el.getPrevSibling
```

### `getFirstChild` / `getChildren`

```vb
Public Function getFirstChild() As CDPElement
Public Function getChildren() As Collection
```

先頭の子／子要素のコレクションです。コレクションの各要素は `CDPElement` です。

```vb
Dim child As CDPElement
Set child = el.getFirstChild

Dim kids As Collection
Set kids = el.getChildren
For Each child In kids
    Debug.Print child.innerText
Next
```

## 入れ子検索

要素スコープでも Context と同様に検索できます（主に Shadow Root 内や部分 DOM 向け）。

### `getElementByID`

```vb
Public Function getElementByID(strID As String) As CDPElement
```

| 引数 | 意味 |
| --- | --- |
| `strID` | `id` 属性（`#` は付けない） |

```vb
root.getElementByID("ok").click
```

### `getElementByQuery` / `getElementsByQuery`

```vb
Public Function getElementByQuery(strQuery As String) As CDPElement
Public Function getElementsByQuery(strQuery As String) As Collection
```

| 引数 | 意味 |
| --- | --- |
| `strQuery` | CSS セレクタ |

```vb
Set el = host.getElementByQuery("button.primary")

Dim list As Collection
Set list = host.getElementsByQuery("a.item")
```

### `getElementByXPath` / `getElementsByXPath`

```vb
Public Function getElementByXPath(strXPath As String) As CDPElement
Public Function getElementsByXPath(strXPath As String) As Collection
```

現在要素を contextNode にした XPath 検索です。先頭の `//` は無効構文として自動除去されます。

| 引数 | 意味 |
| --- | --- |
| `strXPath` | XPath（相対パス想定） |

```vb
Set el = host.getElementByXPath(".//button[contains(.,'送信')]")
```

## Shadow / iframe

### `getIFrame`

```vb
Public Function getIFrame() As CDPElement
```

`<iframe>` 要素の `contentDocument` を `CDPElement` として返します。**同一オリジン**の iframe 向けです。別ドメインの場合は機能しないため、[`CDPBrowser.getTab`](./CDPBrowser#gettab) でターゲット接続してください。

```vb
Dim frame As CDPElement
Set frame = t.getElementByQuery("iframe#app").getIFrame
frame.getElementByID("inner").click
```

### `GetShadowRoot` / `GetShadowRoots`

```vb
Public Function GetShadowRoot(Optional Index As Long) As CDPElement
Public Function GetShadowRoots() As Collection
```

現在要素の直下にある Shadow Root に入ります。ホスト要素を取ってから呼び出す「通り抜け」イメージです。

| 引数 | 意味 |
| --- | --- |
| `Index` | 複数あるときの 0 始まりインデックス（`GetShadowRoot` のみ。既定 `0`） |

```vb
Dim host As CDPElement
Set host = t.getElementByQuery("my-widget")

Dim root As CDPElement
Set root = host.GetShadowRoot
root.getElementByQuery("button").click

' 2 つ目の Shadow Root
Set root = host.GetShadowRoot(1)
```

失敗時は `Nothing` です。デモ: `Demo_CDP.SimpleShadowRootTest` / `iframeShadowRootTest` / `runIFrame`

## その他

### `ExposeDevTools`

```vb
Public Sub ExposeDevTools(varName As String)
```

保持中の要素を DevTools コンソールのグローバル変数として公開します（デバッグ用）。

| 引数 | 意味 |
| --- | --- |
| `varName` | コンソール上の変数名（例: `"el"`） |

```vb
el.ExposeDevTools "el"
' DevTools コンソールで el を参照可能
```

::: tip 注意
サイトによっては不自然なグローバル変数を検知する対策があります。本番スクレイピングでは使わないでください。
:::

### `CurrentObjectId`

```vb
Property Get CurrentObjectId() As String
```

内部で保持している CDP の `objectId` です。空なら検索未ヒットです。日常利用では通常不要です。

### `StopException`

```vb
Public StopException As Boolean
```

この要素経由の `jsEval` で JS 例外時に停止するかのスイッチです。開発時のみ `True` を推奨します。

### `Init`

フレームワーク内部の初期化です。ユーザーコードから直接呼ぶ必要は通常ありません。

## 関連

- [`CDPContext`](./CDPContext)
- [要素の取得](/guides/selectors)
- [入力とクリック](/guides/input)
