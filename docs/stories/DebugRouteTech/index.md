---
description: Excel VBA からブラウザへつながる3つの経路（Pipe・WebSocket・WebView2）を、映像でひと通り見られるページです。各経路の詳しい仕組みは、続く3本の技術資料で扱います。
---

# 技術資料 ― ブラウザへ続く3つの道

このツールは、Excel VBA と Windows 標準の API だけで、Chromium 系のブラウザを CDP（Chrome DevTools Protocol）で操ります。ブラウザへつながる通り道は3つあり、それぞれ渡し方も、ぶつかった壁も違います。

まずは映像で、3つの道をひと通り見てください。

<PromoMovie src="/promo/ReiwaIE_DebugRoute.html" title="令和のIE 技術秘話 ― Pipe / Port / WebView2" />

## この資料で扱うこと

| 経路 | ひとことで | 詳しく |
| --- | --- | --- |
| Pipe（`--remote-debugging-pipe`） | fd3/4 に匿名パイプをつなぎ、`\0` 区切りで CDP をやり取りする | [Pipe 経路](/stories/DebugRouteTech/pipe-tech) |
| WebSocket（`--remote-debugging-port`） | WinSock でハンドシェイクとフレームを自作し、Pipe と同じ受信処理に流し込む | [WebSocket 経路](/stories/DebugRouteTech/websocket-tech) |
| WebView2（`CallDevToolsProtocolMethod`） | vtable の番号とコールバックを VBA で組み立て、COM 経由で CDP を受け取る | [WebView2 経路](/stories/DebugRouteTech/webview2-tech) |

3つの道は、最後はどれも「JSON ＋ `\0`」という同じ形のメッセージにそろい、同じ受信処理（`RaiseEvent`）に合流します。

::: info 映像の終盤について
映像の終わりでは、[WebView2 経路](/stories/DebugRouteTech/webview2-tech)の「AI が越えた実現の壁」の節を受けて、AI が絡んだ個人情報の漏えいや悪用の事例（2023〜2026年）にも触れています。事例の出典は、映像のエンドクレジットに載せています。
:::
