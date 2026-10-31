---
description: jsEval の returnByDetails:=True で受け取る、JavaScript 実行結果の詳細オブジェクト JsEvalResult。型・例外情報・objectId／sharedId／handle などを取り出せます。
---

# JsEvalResult

`jsEval` の実行結果を、**詳細に受け取るためのクラス**です。CDP と WebDriver BiDi の両方で使えます。

ふだんの `jsEval` は、結果の「値」だけを返します。値や `VarType` だけでは判断しにくい場合に、`returnByDetails:=True` を付けると、この `JsEvalResult` が返ります。

- CDP: `Runtime.evaluate` / `Runtime.callFunctionOn` の `Runtime.RemoteObject` と `exceptionDetails`
- BiDi: `script.evaluate` / `script.callFunction` の `script.RemoteValue` と `script.ExceptionDetails`

## 使い方

戻り値はオブジェクトなので、`Set` で受け取ります。

```vb
Dim r As JsEvalResult
Set r = t.jsEval("document.title", returnByDetails:=True)

If r.isException Then
    Debug.Print "例外: " & r.ExceptionText
Else
    Debug.Print r.ResultType, r.value
End If
```

`returnByDetails` は、次の `jsEval` で使えます。

| クラス | メソッド |
| --- | --- |
| [`CDPContext`](./CDPContext) | `jsEval(..., Optional returnByDetails As Boolean, ...)` |
| [`CDPElement`](./CDPElement) | `jsEval(..., Optional returnByDetails As Boolean, ...)` |
| [`WebDriverBiDiContext`](/api/bidi/WebDriverBiDiContext) | `jsEval(..., Optional returnByDetails As Boolean, ...)` |
| [`WebDriverBiDiElement`](/api/bidi/WebDriverBiDiElement) | `jsEval(..., Optional returnByDetails As Boolean, ...)` |

- 例外が起きた場合も、`StopException` を指定していなければ `JsEvalResult` が返ります。`isException` で判定してください
- 非同期実行（`RunAsyncCDP` / `RunAsyncBiDi`）のときは無視され、コマンド ID が返ります

::: tip 取り出すたびに解析します
`value` 以外のプロパティは、呼ぶたびに JSON を解析します。同じ値を何度も使う場合は、変数に控えてから使ってください。
:::

## 共通のプロパティ

| プロパティ | 型 | 意味 |
| --- | --- | --- |
| `value`（既定メンバー） | `Variant` | 結果の値。ふだんの `jsEval` が返す値と同じです（下記「`value` の決まり方」） |
| `rawValue` | `Variant` | 加工前の `value`（JSON ノード、またはそのままの値） |
| `ResultType` | `String` | JavaScript 側の型（`string` / `number` / `object` / `node` など）。例外時は `exception` |
| `ResultTypeForVBA` | `VbVarType` | VBA 側の型。`vbDataObject` は、`value` が objectId / sharedId / handle などの ID 文字列であることを表します |
| `isException` | `Boolean` | JavaScript の例外が起きたか |
| `ExceptionText` | `String` | 例外のメッセージ |
| `ExceptionLineNumber` / `ExceptionColumnNumber` | `Long` | 例外の行・列 |
| `ExceptionStackTrace` | `BiDiCDPJConv` | スタックトレース |
| `ExceptionResultObject` | `JsEvalResult` | 投げられた例外オブジェクトそのものを、`JsEvalResult` として |
| `rawException` | `BiDiCDPJConv` | 例外情報の JSON ノード |

## CDP だけのプロパティ

| プロパティ | 型 | 意味 |
| --- | --- | --- |
| `objectId` | `String` | `Runtime.RemoteObjectId` |
| `ClassName` | `String` | オブジェクトのクラス名 |
| `Description` | `String` | 文字列表現 |
| `ResultSubtype` | `String` | `node` / `array` / `error` などのサブタイプ |
| `UnserializableValue` | `String` | `NaN` / `Infinity` / `-0` / BigInt など、JSON にできない値 |
| `Preview` / `CustomPreview` | `BiDiCDPJConv` | オブジェクトのプレビュー（`generatePreview:=True` のとき） |
| `DeepSerializedValue` | `BiDiCDPJConv` | `serializationOptions` でディープシリアライズした値 |

## WebDriver BiDi だけのプロパティ

| プロパティ | 型 | 意味 |
| --- | --- | --- |
| `sharedId` | `String` | `script.SharedId`。DOM ノードを返したときに入ります（要素の参照。解放不要） |
| `handle` | `String` | `script.Handle`。`Ownership:=True` のとき、ノード以外のオブジェクトに入ります（不要になったら `script.disown` で解放） |
| `internalId` | `String` | `script.InternalId` |

::: info BiDi は受け取れる情報が少なめです
WebDriver BiDi（W3C 仕様）の `script.RemoteValue` には、CDP の `className` / `description` / `objectId` などに当たる情報がありません。CDP だけのプロパティは、BiDi では空になります。
:::

BiDi では、要素は `sharedId`、配列やオブジェクトは `handle` と、2種類の参照を使い分けます。どちらが返ったかを確かめたいときに `returnByDetails:=True` が役立ちます。

## `value` の決まり方

- **CDP**: `deepSerializedValue` → `objectId` → `value` → `unserializableValue` の順に、見つかったものを返します。`null` は `Null`、`undefined` は `Empty` になります
- **BiDi**: `sharedId` → `handle` → `value` の順です。`date` は VBA の `Date` に変換します
- 例外時は、`jsEval` の `IFEXCEPTION` に指定した値（未指定なら `#VALUE!` エラー値）になります

## 関連

- [JavaScript 実行](/guides/javascript)
- [`CDPContext.jsEval`](./CDPContext) / [`WebDriverBiDiContext.jsEval`](/api/bidi/WebDriverBiDiContext)
