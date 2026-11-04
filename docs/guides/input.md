---
description: CDPElement のクリック・フォーム送信・テキスト入力・キー送信・属性操作。React など仮想 DOM 向けの fireEvent パターンも紹介します。
---

# 入力とクリック

`CDPElement` の操作メソッドです。BiDi のみのコンテキストでは [要素の取得](/guides/selectors) のとおり CDP に変換してから使います。

## クリック・送信

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext(StartURL:="https://example.com/form")

t.getElementByID("submit").click
t.getElementByQuery("form").submit
```

`click` は省略可能な `ReadyState` で、クリック後の読み込み待ちを指定できます。

## テキスト入力

```vb
Dim box As CDPElement
Set box = t.getElementByQuery("input[name='q']")
box.clearValue
box.sendString "検索キーワード"
box.focus
box.selectText
```

日本語や絵文字も `sendString` で送れます。  
デモ: `Demo_CDP.JapaneseElementTest`

## 属性・選択

```vb
Debug.Print box.PageAttribute("value")
box.PageAttribute("data-x") = "1"

' <select>：value 属性 / 表示文字 / 番号 のどれか1つで選び、同じ基準で取り出す
Dim sel As CDPElement
Set sel = t.getElementByQuery("select#country")
sel.SelectOption "JP"
sel.SelectOption Label:="日本"
Debug.Print sel.GetSelection(seLabel)
```

`<select>` の詳細は [`CDPElement.SelectOption` / `GetSelection`](/api/cdp/CDPElement#selectoption--getselection) を参照してください。

## キー入力

```vb
box.sendKey keyEnter   ' keyboardCode 列挙を参照
box.sendClick          ' 座標系クリック相当
box.sendHover          ' 座標系ホバー相当（:hover で出現するメニューなど）
```

## React など仮想 DOM

単純な `PageAttribute` などの代入では状態が同期されないことがあります。デモ `Demo_CDP.fillReactForm` のように、フォーカスやイベント発火（`fireEvent`）を組み合わせてください。

```vb
el.focus
el.sendString "value"
el.fireEvent "input"
```

## 関連

- [`CDPElement`](/api/cdp/CDPElement)
- [JavaScript 実行](/guides/javascript)
