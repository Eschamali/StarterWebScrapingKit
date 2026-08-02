---
description: 1 つのブラウザで複数タブ・ウィンドウを扱う方法。CDP / BiDi の newTab・getTab と並列操作のパターンをコード例で紹介します。
---

# マルチタブ

1 つのブラウザプロセス内で複数タブ（またはウィンドウ）を扱います。  
CDP デモ: `runTabsAsOne` / `runTabsAsMany` / `runNewTab`。

## Context から newTab（CDP）

```vb
Dim main As CDPContext
Set main = ShSetting01_StartBrowser.StartCDPModeContext
main.navigate "https://google.com"

main.InheritanceCDPBrowser.newTab "https://example.com"
main.InheritanceCDPBrowser.newTab "https://bing.com"
```

## Browser からタブを割り当て（CDP）

```vb
Dim chrome As CDPBrowser
Set chrome = ShSetting01_StartBrowser.StartCDPMode

Dim tab1 As CDPContext, tab2 As CDPContext, tab3 As CDPContext
Set tab1 = chrome.getTab(setMain:=True)
Set tab2 = chrome.newTab(newWindow:=True)
Set tab3 = chrome.newTab(newWindow:=True)

tab1.navigate "https://google.com"
tab2.navigate "https://example.com"
tab3.navigate "https://bing.com"

chrome.quit
```

`getTab` / `newTab` で操作対象をメインにするときは **`setMain:=True`** を付けます（reattach 後も同様）。

## BiDi

```vb
Dim mode As WebDriverBiDiMode
Set mode = ShSetting01_StartBrowser.StartBiDiMode("https://news.google.com/home")

Dim tab As WebDriverBiDiContext
Set tab = mode.getTab("https://news.google.com/", setMain:=True)
tab.navigate "https://example.com"

' 新規
' Set tab = mode.newTab(setMain:=True)

mode.quit
```

## マルチプロファイル（別プロセス）

同じプロファイルでは非同期並列に限界があります。別ユーザ名で起動すると独立インスタンスになります。

```vb
Set e1 = ShSetting01_StartBrowser.StartCDPModeContext
Set e2 = ShSetting01_StartBrowser.StartCDPModeContext(SwitchUser:="CDP2")
```

デモ: `Demo_CDP.demoMultiProfileOperation`

## 関連

- [`CDPBrowser.newTab` / `getTab`](/api/cdp/CDPBrowser)
- [`WebDriverBiDiMode.newTab` / `getTab`](/api/bidi/WebDriverBiDiMode)
- [再接続](/guides/reattach)
