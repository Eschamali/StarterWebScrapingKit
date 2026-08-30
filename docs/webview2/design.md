---
description: WebView2 モードを増設した理由と設計。COM/vtable経由でCDPをやり取りする仕組みと、移植元プロジェクトへのクレジットを説明します。
---

# 設計思想について

Pipe・WebSocket に続く3つ目の通信路として、v3.0.0 で **WebView2（`ICoreWebView2`）を直接叩くモード**を増設しました。

Pipe・WebSocket がどちらも「**外にいるブラウザプロセス**」を相手にするのに対し、WebView2 モードは「**Excel自身のUserFormに埋め込んだブラウザ**」を相手にします。デバッグポートも名前付きパイプも一切開かず、WebView2 SDK が公開する COM インターフェースを直接呼び出す、まったく別の経路です。

## 既存のCDPスタックへの繋ぎ込み方

`CDPCoreViaWebSocket.cls` と同じ「`CDPCore.cls` に `RaiseEvent` で繋ぎ込む」拡張パターンを踏襲しています。実装対象は WebView2 SDK の3メソッドだけです（このツールは CDP フォーマットしか扱わないため、ナビゲーション制御や DOM 操作などの WebView2 本来の機能は一切実装しません）。

- `CallDevToolsProtocolMethodAsync` → `ICoreWebView2::CallDevToolsProtocolMethod`
- `CallDevToolsProtocolMethodForSessionAsync` → `ICoreWebView2_11::CallDevToolsProtocolMethodForSession`
- `GetDevToolsProtocolEventReceiver` → `ICoreWebView2::GetDevToolsProtocolEventReceiver`

WebSocket 版は「バイト列の断片を都度 `RaiseEvent` で流す」設計でしたが、WebView2 は COM コールバック経由で「UTF-16 デコード済み・欠けのない完成 JSON 文字列」を1件ずつ届けてくれます。そのため `CDPCoreViaWebView2` は `RaiseEvent CDPMessageReceived(RawJson As String)` という、文字列1件そのままの、より単純な形のイベントを発火するだけで済みます。

## 既知の制約

::: warning WebView2独自のイベント購読モデル
CDP-over-Pipe / CDP-over-WebSocket は「ドメインを `enable` すれば、以後そのドメインの全イベントが自動で流れてくる」モデルです。しかし WebView2 の `GetDevToolsProtocolEventReceiver` は **「イベント名ごとに個別登録」** が必要なモデルです。この違いを隠さず、`SubscribeCdpEvent` / `UnsubscribeCdpEvent` という明示的な API として公開しています（一括購読の概念はWebView2側に無いため未対応。一括解除のみ `UnsubscribeAllCdpEvents` として提供）。
:::

::: warning VBEでのブレークに注意
すべてのCOMコールバックは、機械語で書かれたサンク（後述）を経由します。コールバック待ち中（コマンド送信〜完了、イベント購読中）にVBEでブレーク／ステップ実行すると、Excelがクラッシュする可能性があります。
:::

::: info 複数タブは「別ウィンドウ」として開きます
`CDPBrowser.newTab`（内部的には`Target.createTarget`）は、WebView2モードでも他のtransportと同じように使えます。ただし、WebView2の`ICoreWebView2`は1インスタンス=1ページのため、新しく作られたタブをUserFormの中に**埋め込む**仕組みまでは用意していません。そのため、2つ目以降のタブは独立した新規ウィンドウとして立ち上がります。

タブ（＝ウィンドウ）をまたいだCDPコマンドの送受信自体は、`ICoreWebView2_11::CallDevToolsProtocolMethodForSession`（`SendCommandCDP`が`sessionId`の有無で自動的に切り分けます）で正しくルーティングされるため、`CDPContext`を複数持って並行操作すること自体は可能です。「UserFormの中に複数タブ分のビューを並べて表示する」機能が無いだけです。
:::

## 低レイヤの実装：機械語サンクとvtable

WebView2は`IUnknown`ベースのCOMオブジェクトで、VBAの`Object`変数（IDispatchベース）としては直接扱えません。関数の呼び出しには`DispCallFunc`（vtableのインデックスを直接指定して実行するWindows API）を使い、コールバックを受け取るには、`AddressOf`で取得した関数ポインタをメモリ上に構造体として詰め込み、「COMオブジェクトのフリをしたデータ」を構築する（**vtable偽造**）必要があります。

::: tip 移植元へのクレジット
この機械語サンク・vtable呼び出し・SAFEARRAYメモリプリミティブの心臓部（`CDPWebView2Thunks.bas`）は、[**WebView2-For-Excel-VBA**](https://github.com/tarboh/WebView2-For-Excel-VBA)（作者：たーぼー(インコ) 氏、MIT License）の `Wv2Thunks.bas` を、バイト列やオフセット値を一切変更せずそのまま移植したものです。

このツール向けに追加/変更したのはCDP専用の薄い層だけです。

- `HandlerKind` を、CDP用の4種類（`HK_EnvironmentCompleted` / `HK_ControllerCompleted` / `HK_CdpMethodCompleted` / `HK_CdpEventReceived`）に絞り込み
- `CallDevToolsProtocolMethodCompletedHandler` と `DevToolsProtocolEventReceivedEventHandler` の実IIDを、内部のIIDテーブルへ追加

低レベルの死闘（QueryInterfaceがE_NOINTERFACEを返す、vtableオフセットが1つずれるだけでクラッシュする等）をすでに乗り越えてくれていた先人の実装があったからこそ、CDP側は安心して「その上に何を乗せるか」だけに集中できました。改めて感謝します。
:::

## 関連

- [WebView2モードでできること](/webview2/capabilities)
- [Excel単独で「真のWebView2」を完全制御する](/userform/vba-only) — UserFormへの埋め込み手順
- [アーキテクチャ](/concepts/architecture)
- [このツールの誕生秘話](/stories/birth-story) — WebView2ブランチが辿った半年
