# 技術解説：何故、WebDriverBiDi.exe なしで動くことがわかったか

## 1. はじめに：従来の常識と「壁」
通常、Selenium や初期の WebDriver BiDi 実装では、ブラウザを操作するために `chromedriver.exe` や `WebDriverBiDi.exe` といった「中間バイナリ（EXE）」が必要でした。これらは：
- HTTPやWebSocketのプロトコル変換を行う代理人（Proxy）として機能
- ブラウザの起動オプションやセッション管理を統括
という役割を担っていましたが、VBA環境においては「外部EXEの配布・管理」が最大のネックとなっていました。

## 2. 突破口：`BiDiPoc.bas` が証明した「セルフ・プロキシ」の概念
本プロジェクトの核心は、**「EXEがやっている変換処理を、ブラウザ内部のJavaScriptに行わせる」**という逆転の発想にあります。

この着想の原型は、リポジトリ直下の `BiDiPoc.bas`（Proof of Concept）に生々しく記録されています。

### 実現のための「3つの神器」
`BiDiPoc.bas` のコード内で、以下の3つのCDPメソッドを組み合わせることで、EXEなしのBiDi制御が成立することが証明されました。

1.  **`Target.exposeDevToolsProtocol`**
    - **役割**: ブラウザ内部の特定のタブ（Mapper用タブ）に対して、JSから直接CDPを叩ける特別なオブジェクト `window.cdp` を露出させます。
    - **コード抜粋**:
      ```vb
      paramsCDP.Add "bindingName", "cdp"
      Set ResultCDP = .invokeMethod("Target.exposeDevToolsProtocol", paramsCDP, True)
      ```

2.  **`Runtime.addBinding`**
    - **役割**: JSからVBA（ホスト側）へメッセージを「逆流」させるためのブリッジ関数を作ります。これにより、BiDiのイベントがVBAへ通知されるようになります。
    - **コード抜粋**:
      ```vb
      paramsCDP.Add "name", "sendBidiResponse"
      .invokeMethod "Runtime.addBinding", paramsCDP
      ```

3.  **`Runtime.evaluate` による `mapperTab.js` の注入**
    - **役割**: Chromium公式チームが開発している `chromium-bidi` のコアロジックをJS文字列としてブラウザに流し込みます。
    - **コード抜粋**:
      ```vb
      .jsEval CharConv.BytesToString(BiDiMapperScript)
      .jsEval "window.runMapperInstance('" & current_targetID & "')"
      ```

## 3. 実現のメカニズム
これにより、以下のような「EXEレス」の通信フローが完成しました。

1.  **VBA側**: BiDi形式のJSONを `Runtime.evaluate` を通じてブラウザ内のJSに投下。
2.  **ブラウザ内 (JS/Mapper)**: 届いたJSONを解釈し、`window.cdp` を通じて自分自身や他のタブへ命令（CDP）を飛ばす。
3.  **ブラウザ内 (JS/Mapper)**: 実行結果や非同期イベントを受け取り、`sendBidiResponse`（Binding）を通じてVBAへ返す。
4.  **VBA側**: `TakeEvents` メソッドで、Binding経由で届いたメッセージを回収し、ユーザーに引き渡す。

## 4. 結論：何が変わったのか
この方法の確立により、**「Excelファイル1つとChromiumブラウザさえあれば、世界標準の次世代プロトコル（BiDi）をフルパワーで扱える」**という、極めてポータビリティの高いスクレイピング環境が誕生しました。

`BiDiPoc.bas` で行われたこの「生々しい実験」こそが、現在の洗練された `WebDriverBiDiCore.cls` の母体であり、技術的なブレイクスルーの瞬間でした。
