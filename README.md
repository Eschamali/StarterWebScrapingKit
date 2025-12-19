# Excel VBA Web Automation Starter Kit
## インターネットの世界を、その手に

スクレイピングに必要なすべての要素を、このマクロブック「1つ」に詰め込みました。(配布はもうしばらくお待ち下さい)  
面倒な環境構築はもう必要ありません。このファイルを開いたその瞬間から、あなたの業務効率化とインターネット自動操作への旅が始まります。

本ツールは、現代のWeb技術を攻略するために必須となる「3つの神器」を実装しています。

1.  **🚀 REST WebAPI (WinHTTP 5.1)**
    *   高速・軽量なデータ収集の王道。参照設定のみで完結する堅牢な実装です。
2.  **🤖 ブラウザ自動操作 (CDP via Pipe)**
    *   Chromiumベースのブラウザ（Edge/Chrome）を自在に操ります。外部ドライバー(exe)を必要としない、パイプ通信によるモダンな実装です。
3.  **⚡ WebSocket 通信 (Beta)**
    *   リアルタイム通信への挑戦。WinAPIを駆使し、最低限の接続・送受信機能を搭載しました。VBAの限界を押し広げる、発展途上の機能です。

#### 【Credits & Acknowledgments】
このツールは、世界中のVBA職人が公開してくれた素晴らしいライブラリの数々を、実務で使いやすい形に統合（マッシュアップ）したものです。
偉大な先人たちの知恵とコードに、心からの敬意と感謝を表します。

*   **WebSocket実装のコアロジック**
    *   [ChromeControler-No-Selenium-WebDriver-VBAJSON](https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON)
*   **CDP制御・パイプ通信の基盤**
    *   [Chromium-Automation-with-CDP-for-VBA](https://github.com/GCuser99/Chromium-Automation-with-CDP-for-VBA)
*   **WinHTTP 5.1 ラッパー**
    *   [VBA-Web](https://github.com/VBA-Tools-v2/VBA-Web)
*   **高速・高機能なJSONパーサー**
    *   [WebJsonConverter.cls (from SeleniumVBA)](https://github.com/GCuser99/SeleniumVBA/blob/main/src/VBA/WebJsonConverter.cls)
    *   ※メンテナンス性を考慮し、既存のJsonConverterからこちらへ換装済み

※各機能の詳細な使用方法やメソッドについては、上記オリジナルライブラリのドキュメントをご参照ください。
