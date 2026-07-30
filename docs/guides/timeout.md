# タイムアウト設定方法について

CDP / BiDi のコマンド結果待ちや、起動直後の遷移完了判定などに使う待機上限です。デフォルトは **30 秒**です。

「この処理、時間がかかるな」と分かっているときは、意図的に伸ばして使ってください。

## `TimeOutSecond`

```vb
Property Get TimeOutSecond() As Double   ' 設定中のタイムアウト秒数
Property Let TimeOutSecond(TimeSec As Double)
```

| | 意味 |
| --- | --- |
| **GET** | 設定中のタイムアウト秒数 |
| **LET** | 置き換えるタイムアウト秒数 |

主な用途:

- 最初のブラウザ起動時の遷移完了判定
- CDP-Json / BiDi-Json コマンド結果を得るまでの待ち時間

設定できる主なクラス:

| クラス | 用途 |
| --- | --- |
| [`CDPBrowser`](/api/cdp/CDPBrowser) | ブラウザ単位のコマンド待ち |
| [`CDPContext`](/api/cdp/CDPContext) | タブ単位のコマンド待ち・起動遷移など |
| [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode) | BiDi コマンド結果待ち |

## 使い方

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext

' 重いページや遅い応答が分かっているとき
t.TimeOutSecond = 60

Debug.Print t.TimeOutSecond   ' → 60

t.navigate "https://example.com/heavy"
' ... 通常どおり操作 ...

t.InheritanceCDPBrowser.quit
```

ブラウザ側にも同様に設定できます。

```vb
Dim b As CDPBrowser
Set b = ShSetting01_StartBrowser.StartCDPMode
b.TimeOutSecond = 90
```

BiDi の場合:

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext
t.InheritanceWebDriverBiDiMode.TimeOutSecond = 60
```

::: tip
タイムアウトを伸ばしても、実際にハングしているパイプ／WebSocket までは救えません。接続自体が死んでいる場合は [再接続 (reattach)](/guides/reattach) を検討してください。
:::

## `TimerCounter`（自前ループ用）

VBA 組み込みの `Timer` 関数は、日付が変わると 0 に戻るため、深夜をまたぐと経過時間判定がおかしくなります。

`TimerCounter` はその代わりにどうぞ。**単調増加の経過ミリ秒**を返すので、自前ループのタイムアウト判定に使えます。

```vb
Public Function TimerCounter() As Double   ' 経過ミリ秒
```

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext

Dim startMs As Double
startMs = t.InheritanceCDPBrowser.TimerCounter

Do
    ' ... ポーリングなど ...
    If t.InheritanceCDPBrowser.TimerCounter - startMs > 5000 Then Exit Do  ' 5 秒で打ち切り
    DoEvents
Loop
```

[`CDPBrowser`](/api/cdp/CDPBrowser) / [`CDPContext`](/api/cdp/CDPContext) / [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode) から利用できます。

## 関連

- [ページ遷移](/guides/navigation)
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)
- [再接続 (reattach)](/guides/reattach)
