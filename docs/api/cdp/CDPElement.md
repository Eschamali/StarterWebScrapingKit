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

`<select>` の選択状態です。取得は先頭の `selectedOptions[0]`（該当なしなら空扱い）、代入は `selectedIndex`（0 始まりの**位置**）での切り替えです。option の `value` 属性で選びたい場合は [`setSelection`](#setselection) を使ってください。

```vb
Debug.Print t.getElementByQuery("select#country").selected

' 2番目（index=1）の option を選ぶ
t.getElementByQuery("select#country").selected = 1
```

## 操作

### `click`

```vb
Public Function click() As Boolean
```

要素をクリックします（スクロールインビュー → 合成 `MouseEvent('click')` を `dispatchEvent`）。

```vb
t.getElementByID("submit").click
```

::: tip 注意
クリック後の画面遷移待ちは自動では行いません。必要なら呼び出し側で [`CDPContext.wait`](./CDPContext#wait) を呼んでください。
:::

### `SimpleClick`

```vb
Public Function SimpleClick() As Boolean
```

要素の `this.click()` をそのまま呼ぶ、素朴なクリックです。合成 `MouseEvent` を経由する [`click`](#click) と違い、ブラウザ標準のクリック処理（フォーム送信ボタンの既定動作など）にそのまま乗せたいときに使います。

```vb
t.getElementByID("submit").SimpleClick
```

### `submit`

```vb
Public Function submit() As Boolean
```

所属フォームを送信します（`this.form.submit()`）。

```vb
t.getElementByQuery("form").submit
```

::: tip 注意
送信後の画面遷移待ちは自動では行いません。必要なら呼び出し側で [`CDPContext.wait`](./CDPContext#wait) を呼んでください。
:::

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

### `sendHover`

```vb
Public Function sendHover() As Boolean
```

要素の中心座標へ物理的なマウスホバーをシミュレートします（`Input.dispatchMouseEvent` の `mouseMoved`）。JS の `dispatchEvent` では反応しない、CSS の `:hover` で出現するメニューやツールチップなど向けです。呼び出し前に `scrollIntoView` で自動的にビューポート内へ収めます。

```vb
t.getElementByQuery(".menu-item").sendHover
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
Public Function fireEvent(strEventName As String) As Boolean
```

DOM イベントを発火します。React 等で `setAttribute` だけでは状態が同期されないときに使います。

| 引数 | 意味 |
| --- | --- |
| `strEventName` | JS のイベント名そのもの（`"input"` / `"change"` / `"blur"` など） |

```vb
el.focus
el.sendString "value"
el.fireEvent "input"
```

::: tip 注意
`"on"` プレフィックス（`"onchange"` など）は自動除去されません。IE 時代の命名（`onchange`）ではなく、JS の正しいイベント名（`change`）をそのまま渡してください。
:::

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

### `SetFileInputFiles`

```vb
Public Sub SetFileInputFiles(files As Collection)
```

`<input type="file">` へ、ダイアログ操作なしでファイルを添付します（CDP の `DOM.setFileInputFiles`）。

| 引数 | 意味 |
| --- | --- |
| `files` | 添付したいファイルの**フルパス**を格納した `Collection` |

```vb
Dim files As New Collection
files.Add "C:\path\to\image.png"

t.getElementByQuery("input[type='file']").SetFileInputFiles files
```

::: tip 注意
`files` が `Nothing` または空の場合は、警告ログを出して何もしません。
:::

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

::: warning 使えないケース
`onExist` / `onExistNot` は、要素を取得したときの検索コード（JavaScript）を内部で再実行してポーリングします。そのため、そもそも単一の検索コードを持たない次の取得方法では**使用できません**（要素自体は問題なく操作できますが、ポーリングは効きません）。

- 複数取得系（[`getChildren`](#getfirstchild--getchildren) / [`getElementsByQuery`](#getelementbyquery--getelementsbyquery) / [`getElementsByXPath`](#getelementbyxpath--getelementsbyxpath)）で得られた各要素
- [`GetShadowRoot`](#getshadowroot--getshadowroots) / `GetShadowRoots` で得られた Shadow Root
:::

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

::: tip
判定は「保持中のこの要素（`objectId`）が `document` から外れたか（`this.isConnected`）」で行います。同じセレクタに一致する**別の**要素が現れても、元の要素自体が外れていれば消滅とみなします。
:::

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

`<iframe>` 要素の `contentDocument` を `CDPElement` として返します。**同一オリジン**の iframe 向けです。別ドメインの場合は機能しないため、[`CDPBrowser.getTab`](./CDPBrowser#gettab) でターゲット接続するか、Context 側の [`CDPContext.getIFrameContextID`](./CDPContext#getiframecontextid) / [`getIFrame`](./CDPContext#getiframe) を検討してください。

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

### `jsEval`

```vb
Public Function jsEval(JavaScriptStr As String, Optional objectArguments As Variant, _
    Optional IFEXCEPTION As Variant, Optional returnByValue As Boolean, _
    Optional awaitPromise As Boolean, Optional serializationOptions As Dictionary, _
    Optional generatePreview As Boolean, Optional StopApiError As Boolean = True) As Variant
```

この要素の `objectId` を `this` として JavaScript を評価します。[`CDPContext.jsEval`](./CDPContext#jseval) の要素スコープ版で、クラス内の他メソッドも内部的にこれ経由で実装されています。用意されたメソッドで足りない操作をしたいときの逃げ道として使えます。

`objectArguments` は `Collection` / `Array(...)` / 固定長の `Dictionary` 型 1 次元配列のいずれでも渡せます（詳細は [`CDPContext.jsEval`](./CDPContext#jseval) 参照）。

```vb
Dim el As CDPElement
Set el = t.getElementByQuery("#price")

Debug.Print el.jsEval("function(){ return this.dataset.raw }")
```

### 実行オプション（`SetOptionStopException` / `SetOptionRunAsyncCDP` / `SetOptionUserGesture`）

```vb
Property Let SetOptionStopException(v As Boolean)
Property Let SetOptionRunAsyncCDP(v As Boolean)
Property Let SetOptionUserGesture(v As Boolean)
```

この要素経由の [`jsEval`](#jseval)（および内部で `jsEval` を使う各メソッド）の実行方法を切り替える、Let 専用のスイッチです。

| プロパティ | 意味 |
| --- | --- |
| `SetOptionStopException` | `True` で、JS 例外発生時に `Err.Raise` で停止。基本は開発時のみ `True` を推奨 |
| `SetOptionRunAsyncCDP` | `True` で、結果を待たない非同期実行に切り替える。使い終わったら必ず `False` に戻すこと |
| `SetOptionUserGesture` | `True` で、人間の操作であるかのように偽装する。スクレイピング対策の回避に有効な場合がある |

```vb
el.SetOptionStopException = True   ' デバッグ中だけ例外で止める
el.SetOptionStopException = False
```

### `Init`

フレームワーク内部の初期化です。ユーザーコードから直接呼ぶ必要は通常ありません。

## 関連

- [`CDPContext`](./CDPContext)
- [要素の取得](/guides/selectors)
- [入力とクリック](/guides/input)
