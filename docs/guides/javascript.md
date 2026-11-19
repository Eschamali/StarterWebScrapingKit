---
description: CDP / BiDi の jsEval でページ上の JavaScript を実行する方法。例外処理、handle / objectId、Promise 待ちなどの引数を整理します。
---

# JavaScript 実行

ページ上で任意の JS を評価します。JavaScript 版の低レベル操作としての位置付けで、スクレイピングに関わる引数を一通り用意しています。

入口は次の 2 つです（いずれもタブ／コンテキスト単位）。

| プロトコル | メソッド | 内部プロトコル |
| --- | --- | --- |
| CDP | [`CDPContext.jsEval`](/api/cdp/CDPContext#jseval) | `Runtime.evaluate` / `Runtime.callFunctionOn` |
| BiDi | [`WebDriverBiDiContext.jsEval`](/api/bidi/WebDriverBiDiContext#jseval) | `script.evaluate` / `script.callFunction` |

::: tip
利用時は **名前付き引数**（`xxx:=yyy`）での呼び出しを推奨します。
:::

## 基本

::: code-group

```vb [CDP]
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext("https://example.com")

Dim result As Variant
result = t.jsEval("document.title")
Debug.Print result

' 例外時の代替値（IFERROR チック）
result = t.jsEval("notDefined.x", IFEXCEPTION:="fallback")
```

```vb [BiDi]
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext("https://example.com")

Dim result As Variant
result = t.jsEval("document.title")
Debug.Print result

result = t.jsEval("notDefined.x", IFEXCEPTION:="fallback")
```

:::

`JavaScriptStr` は As Is で実行されます。エスケープの大半は `vbacollective-json` が担うため、手動エスケープは基本不要です。

## 3 つの実行パターン

`objectId` / `scriptSharedId` や `objectArguments` の有無で、内部メソッドが切り替わります。

| パターン | CDP | BiDi | `JavaScriptStr` の書き方 |
| --- | --- | --- | --- |
| 1. 基準オブジェクトあり | `Runtime.callFunctionOn`（`objectId`） | `script.callFunction`（`scriptSharedId`。要素のみ） | `function () { this... }` 形式 |
| 2. 引数だけ安全に渡す | `Runtime.callFunctionOn`（`objectArguments`） | `script.callFunction`（`objectArguments`） | 同上 |
| 3. ただのコード実行 | `Runtime.evaluate` | `script.evaluate` | 通常の式／文 |

```vb
' パターン1: 既存 objectId / sharedId に対して関数実行
result = t.jsEval( _
    "function () { return this.textContent; }", _
    objectId:=oid)          ' BiDi なら scriptSharedId:=sid（要素のみ）

' パターン3: 単純評価
result = t.jsEval("1 + 1")
```

`objectArguments` の組み立ては次を参照してください。

- CDP: [Runtime.CallArgument](https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#type-CallArgument)
- BiDi: [script.LocalValue](https://w3c.github.io/webdriver-bidi/#cddl-type-scriptlocalvalue)

## 主な引数

### 共通イメージ

| 目的 | CDP | BiDi |
| --- | --- | --- |
| 実行コード | `JavaScriptStr` | 同左 |
| 基準オブジェクト | `objectId` | `scriptSharedId`（要素のみ） |
| 関数引数 | `objectArguments` | 同左 |
| 例外時の代替値 | `IFEXCEPTION` | 同左 |
| iframe など実行場所 | `contextId`（`objectId` 指定時は無視） | `RealmTarget` |
| Promise 完了待ち | `awaitPromise` | 同左 |
| ユーザー操作があった扱いにする | `userGesture` | `userActivation` |
| シリアライズ細調整 | `serializationOptions` | 同左 |
| 結果を待たない実行 | `RunAsyncCDP`（戻り値は実行 id） | `RunAsyncBiDi`（戻り値は実行 id） |
| 詳細な結果を受け取る | `returnByDetails`（[`JsEvalResult`](/api/cdp/JsEvalResult) を返す） | 同左 |
| JS 例外で VBE 停止 | `StopException`（開発時向け） | 同左 |
| 通信エラーで停止 | `StopApiError`（既定 `True`） | `StopBiDiError`（既定 `True`） |

### オブジェクト結果の受け取り方（CDP: `returnByValue` / BiDi: `Ownership`）

オブジェクト型の結果を「中身」で取るか、「ブラウザ内の参照 id」で取るかのスイッチです。**CDP と BiDi で True/False の意味が逆寄り**なので注意してください。

| | CDP `returnByValue` | BiDi `Ownership` |
| --- | --- | --- |
| 参照 id を得る（次回 `jsEval` に流用） | `False`（既定）→ `objectId` 文字列 | 要素は既定で `sharedId` 文字列。要素以外は `True` → `handle` 文字列（`resultOwnership: "root"`） |
| 値として受け取る | `True` → `value` を試みる | `False`（既定）→ `value` を試みる（`resultOwnership: "none"`） |

- プリミティブ型の結果は、上記設定に関わらず中身が返ります
- 常に `objectId` / `handle` が欲しい場合は、JS 側の書き方を工夫してください
- 値受け取りで空オブジェクトになる場合は、対応していないデータ型のことがあります。CDP なら `returnByValue:=False`、BiDi なら `Ownership:=True` に切り替えてください
- `returnByValue:=True`（または BiDi で値受け取り）時、結果によって `Set` / `Let` が変わります。静的に決められないときは `Array` で受けて `VarType` 判定する手もあります

```vb
' CDP: オブジェクト参照を保持して次の jsEval に渡す
Dim oid As String
oid = t.jsEval("document.body", returnByValue:=False)
result = t.jsEval("function () { return this.tagName; }", objectId:=oid)

' BiDi: 要素は sharedId を保持（解放不要）
Dim sid As String
sid = t.jsEval("document.body")
result = t.jsEval("function () { return this.tagName; }", scriptSharedId:=sid)

' BiDi: 要素以外は handle を保持し、引数（script.RemoteReference）で渡す
Dim hid As String, args(0) As Dictionary
hid = t.jsEval("[1, 2, 3]", Ownership:=True)
Set args(0) = New Dictionary: args(0).Add "handle", hid
result = t.jsEval("function (a) { return a.length; }", objectArguments:=args)
```

BiDi では、要素は `script.SharedId`、それ以外のオブジェクトは `script.Handle` と、参照が2種類あります。`scriptSharedId` に渡せるのは要素の `sharedId` だけです。

### その他のスイッチ

- **`awaitPromise`**: JS 結果が Promise でも完了まで待つ（`True`）
- **`userGesture` / `userActivation`**: ページを「ユーザーが操作した直後」の状態にしてから実行。ファイル選択やポップアップなどで必要になります。詳しくは[ユーザー操作の扱い](#user-activation)を参照
- **`allowUnsafeEvalBlockedByCSP`（CDP）**: CSP による外部 JS 規制を無視して実行許可
- **`serializationOptions`**: 意地でも `value` で受け取るための細かい調整。指定時は CDP の `returnByValue` は無視されます
- **`generatePreview`（CDP）**: `objectId` の中身をプレーンテキストでプレビュー。ログレベルを DEBUG 以下に。開発時向け
- **`RunAsyncCDP` / `RunAsyncBiDi`**: 結果を待たず実行 id のみ返す。「この操作、`alert` が発動するな…」という場面向け。回収は [低レイヤーガイド](/guides/extend-raw-protocol) / [イベント購読](/guides/events) 参照

## ユーザー操作の扱い {#user-activation}

ブラウザには、ユーザーがクリックやキー入力をしたときだけ使える機能があります。ページが勝手にポップアップを開いたり、ファイル選択ダイアログを出したりするのを防ぐためです。この「ユーザーが操作した」という状態を **user activation** と呼びます。

JS の `element.click()` は、ユーザーの操作ではないので user activation になりません。そのため、`jsEval` でボタンを押しても、ファイル選択ダイアログやポップアップが出ないことがあります。`userGesture:=True` は、この状態を作ってから JS を実行する引数です。

::: info Chromium の中での動き
CDP の `Runtime.evaluate` / `Runtime.callFunctionOn` は、`userGesture` が `True` のとき、実行の前に V8 の `pretendUserGesture` を呼びます（[v8-runtime-agent-impl.cc](https://chromium.googlesource.com/v8/v8/+/main/src/inspector/v8-runtime-agent-impl.cc)）。これを受けて、Blink がそのフレームに user activation を通知します（[thread_debugger_common_impl.cc](https://source.chromium.org/chromium/chromium/src/+/main:third_party/blink/renderer/core/inspector/thread_debugger_common_impl.cc) の `beginUserGesture`）。

BiDi の `userActivation` も、chromium-bidi の中で CDP の `userGesture` に置き換えられます（chromium-bidi の `Realm.ts`）。
:::

::: warning ボット対策をすり抜けるものではありません
変わるのは user activation の状態だけです。`element.click()` で起きるイベントの `isTrusted` は `false` のままなので、`isTrusted` を見るボット対策には効きません。`isTrusted` を `true` にしたいときは、[`isTrusted` との関係](#istrusted)を参照してください。
:::

### `isTrusted` との関係 {#istrusted}

イベントの `isTrusted` は、そのイベントをブラウザ自身が起こしたか（`true`）、JS が作ったか（`false`）を表します。ページ側の JS で読めるので、ボット対策に使われることがあります。

`jsEval`（`Runtime.evaluate` / `Runtime.callFunctionOn`）で起こしたイベントは、`userGesture:=True` を付けても `isTrusted` が `false` になります。`true` にしたいときは、`jsEval` ではなく、入力系のメソッドを使います。

| 操作の方法 | 使うプロトコル | `isTrusted` | user activation |
| --- | --- | --- | --- |
| `jsEval` で `element.click()` や `dispatchEvent` | `Runtime.evaluate` / `Runtime.callFunctionOn` | `false` | `userGesture:=True` のときだけ付く |
| [`sendClick`](/api/cdp/CDPElement#sendclick) / [`sendHover`](/api/cdp/CDPElement#sendhover) / [`sendKey`](/api/cdp/CDPElement#sendkey) など | CDP: `Input.dispatchMouseEvent` / `Input.dispatchKeyEvent`<br>BiDi: `input.performActions` | `true` | マウスの押下やキー入力で自然に付く |

入力系のメソッドは、ブラウザが本物のマウスやキーボードと同じ入口から入力を送るので、ページからは区別がつきません。Puppeteer や Playwright のクリックも、同じ方法で動いています。

ただし、座標で押すので、その座標でいちばん手前に表示されている要素が押されます。要素の上に別のもの（広告など）が重なっていると、そちらが押されます。

::: details Chromium のソースでの確認
`isTrusted` は、ページ側（Blink）がイベントを発火する直前に決まります。

- **入力系（`true`）**：`Input.dispatchMouseEvent` は、DevTools の `InputHandler` から `ForwardMouseEvent` でページに送られます（[input_handler.cc](https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/protocol/input_handler.cc)）。`ForwardMouseEvent` は、本物のマウス入力と同じ入口です。ページ側では `MouseEventManager` が `target->DispatchEvent(*event)` で発火し（[mouse_event_manager.cc](https://source.chromium.org/chromium/chromium/src/+/main:third_party/blink/renderer/core/input/mouse_event_manager.cc)）、この `DispatchEvent` の中で `event.SetTrusted(true)` が呼ばれます（[event_target.cc](https://source.chromium.org/chromium/chromium/src/+/main:third_party/blink/renderer/core/dom/events/event_target.cc)）。
- **`dispatchEvent`（`false`）**：JS から呼ぶと `dispatchEventForBindings` を通り、`event->SetTrusted(false)` が呼ばれます（[event_target.cc](https://source.chromium.org/chromium/chromium/src/+/main:third_party/blink/renderer/core/dom/events/event_target.cc)）。
- **`element.click()`（`false`）**：`HTMLElement::click()` は `SimulatedClickCreationScope::kFromScript` でクリックを作ります（[html_element.cc](https://source.chromium.org/chromium/chromium/src/+/main:third_party/blink/renderer/core/html/html_element.cc)）。`isTrusted` が `true` になるのは `kFromAccessibility` か `kFromUserAgent` のときだけなので、`false` になります（[simulated_event_util.cc](https://source.chromium.org/chromium/chromium/src/+/main:third_party/blink/renderer/core/events/simulated_event_util.cc)）。
:::

### 2 種類の user activation

user activation には 2 種類あります（[MDN: User activation](https://developer.mozilla.org/ja/docs/Web/Security/Defenses/User_activation)）。

| 種類 | 続く期間 | 主な用途 |
| --- | --- | --- |
| 一時的（transient） | 操作から短い時間だけ。一部の機能を使うと、その時点で消える（消費） | ポップアップ、ファイル選択など |
| 持続的（sticky） | 一度でも操作があれば、そのページにいる間ずっと | ページを離れるときの確認ダイアログ、音の自動再生など |

### `userGesture` が要る場面

一時的（transient）な user activation が要る主なものです。

| 場面 | 主な API | user activation を消費するか |
| --- | --- | --- |
| ファイル選択ダイアログ | `<input type="file">` の `click()`、`showPicker()`、`showOpenFilePicker()` | `<input type="file">` は消費する |
| 別ウィンドウ（ポップアップ） | `window.open()`、`target="_blank"` のリンクの `click()` | 消費する |
| 全画面表示 | `requestFullscreen()` | |
| クリップボード | `navigator.clipboard.writeText()` / `readText()` など | |
| 共有・支払い | `navigator.share()`、`PaymentRequest.show()` | |
| 画面共有 | `getDisplayMedia()` | |
| デバイスの選択 | `USB.requestDevice()`、`HID.requestDevice()`、`Serial.requestPort()` | |

持続的（sticky）な user activation が要る主なものは、次のとおりです。

- ページを離れるときの確認ダイアログ（`beforeunload`）
- 音の自動再生（`AudioContext` など）
- `navigator.vibrate()`

ファイル選択ダイアログは、user activation がないと開かず、コンソールに `File chooser dialog can only be shown with a user activation.` という警告が出ます（[file_input_type.cc](https://source.chromium.org/chromium/chromium/src/+/main:third_party/blink/renderer/core/html/forms/file_input_type.cc)）。

::: warning ページを離れるときの確認ダイアログ
一度 `userGesture:=True` で実行すると、そのページは持続的（sticky）な user activation を持ちます。そのため、ページ側が `beforeunload` を設定していると、以降のページ遷移でページを離れるときの確認ダイアログが出ることがあります。遷移が止まったら、[ダイアログ](#ダイアログ)の処理を入れてください。
:::

### 一度付けると、しばらく続く

`userGesture:=True` で実行したあと、続けて `userGesture` を省略して実行しても、しばらくは user activation がある扱いになります。

これは、一時的（transient）な user activation の有効期間が **5 秒** だからです。Chromium のソースで、`kActivationLifespan = base::Seconds(5)` と決まっています（[user_activation_state.h](https://source.chromium.org/chromium/chromium/src/+/main:third_party/blink/public/common/frame/user_activation_state.h)）。5 秒は、`userGesture:=True` の JS を実行し始めた時点から数えます。また、JS の実行が終わっても、user activation は取り消されません。

ただし、5 秒以内でも、消費する機能を一度使うと切れます。たとえばポップアップを 2 つ続けて開くときは、2 つ目にも `userGesture:=True` を付けてください。

今の状態は `navigator.userActivation` で確かめられます。

```vb
Debug.Print t.jsEval("navigator.userActivation.isActive", userGesture:=True)  ' True
Debug.Print t.jsEval("navigator.userActivation.isActive")                     ' 5秒以内なら True
Application.Wait Now + TimeValue("0:00:06")
Debug.Print t.jsEval("navigator.userActivation.isActive")                     ' False（一時的なものは切れた）
Debug.Print t.jsEval("navigator.userActivation.hasBeenActive")                ' True（持続的なものは残る）
```

## 詳細な結果を受け取る（`returnByDetails`）

ふだんの `jsEval` は値だけを返します。値や `VarType` だけでは判断しにくいとき（BiDi で返ったのが `sharedId` か `handle` か、JS 側の型は何か、例外の行番号は、など）は、`returnByDetails:=True` を付けます。戻り値が [`JsEvalResult`](/api/cdp/JsEvalResult) になるので、`Set` で受けます。

```vb
Dim r As JsEvalResult
Set r = t.jsEval("notDefined.x", returnByDetails:=True)

If r.isException Then
    Debug.Print r.ExceptionText, r.ExceptionLineNumber
Else
    Debug.Print r.ResultType, r.value
End If
```

- `value` は、ふだんの `jsEval` の戻り値と同じです
- BiDi は W3C 仕様の都合で、CDP より受け取れる情報が少なめです（`ClassName` / `Description` / `objectId` などは空）
- 非同期実行（`RunAsyncCDP` / `RunAsyncBiDi`）のときは無視されます

## 例外とエラーの見分け

| 状況 | 確認方法 |
| --- | --- |
| JS 実行中の例外 | `IsError(result)`。詳細は `LastJavaScriptException`。面倒なら `IFEXCEPTION` |
| JS は成功したが結果がエラー値 | `IsError` で対処（`IFEXCEPTION` は「例外」時のみ機能） |
| 運用時の `StopException` | JS は例外を処理に使うことが多いため、**省略（止めない）を推奨**。止めるのは開発時向け |
| CDP / BiDi 通信自体のエラー | 既定で停止。最後の引数で制御 |

「この JS ならこういう返り値」という感覚がないと戻り値に困惑しやすいです。開発時はデバッグログで試行錯誤してください。

## スクリプトの追加・通知（CDP）

```vb
t.jsAddLib "https://cdn.example.com/lib.js"                ' URL
t.jsAddScript "C:\scripts", "helper.js"                    ' ローカルファイル（フォルダ, ファイル名）
t.notify "処理が完了しました", 5                            ' 表示秒数
```

## ダイアログ

アラート処理のデモは CDP / BiDi 双方にあります（`Demo_CDP.TestAlert` / `Demo_WebDriverBiDi.TestAlert`）。

BiDi ではセッション capability で自動 dismiss を無効化してから、イベント駆動で応答するパターンが典型です（[はじめに](/getting-started#sessioncapabilitiesrequest-とは) / [イベント購読](/guides/events)）。

CDP:

```vb
t.handleDialog True            ' Accept
t.handleDialog False, "入力"   ' prompt 用テキスト
```

## 関連

- [`CDPContext.jsEval`](/api/cdp/CDPContext#jseval) — [Runtime.evaluate](https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#method-evaluate)
- [`CDPContext.LastJavaScriptException`](/api/cdp/CDPContext#lastjavascriptexception)
- [`WebDriverBiDiContext.jsEval`](/api/bidi/WebDriverBiDiContext#jseval) — [script.evaluate](https://w3c.github.io/webdriver-bidi/#command-script-evaluate)
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)
- [イベント購読](/guides/events)
