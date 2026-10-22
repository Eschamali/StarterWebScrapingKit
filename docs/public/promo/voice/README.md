# ナレーション音声の置き場所

`docs/public/promo/` にある HTML 製ムービーのナレーション（VOICEVOX）を、**動画ごとに 1 フォルダ**で管理します。

## フォルダ構成

```text
promo/
├─ ReiwaIE_60s.html
├─ ReiwaIE_History.html
├─ ReiwaIE_CoreCompare.html
├─ ReiwaIE_VbaCompare.html
├─ ReiwaIE_DebugRoute.html
└─ voice/
   ├─ README.md            … このファイル
   ├─ ReiwaIE_60s/         … ReiwaIE_60s.html の声
   │  ├─ narration.json
   │  └─ A1.wav, A2.wav, …
   ├─ ReiwaIE_History/     … ReiwaIE_History.html の声
   │  ├─ narration.json
   │  └─ P1.wav, P2.wav, …
   ├─ ReiwaIE_CoreCompare/ … ReiwaIE_CoreCompare.html の声
   │  ├─ narration.json
   │  └─ I1.wav, I2.wav, …
   ├─ ReiwaIE_VbaCompare/  … ReiwaIE_VbaCompare.html の声
   │  ├─ narration.json
   │  └─ O1.wav, O2.wav, …
   └─ ReiwaIE_DebugRoute/  … ReiwaIE_DebugRoute.html の声と BGM
      ├─ narration.json
      ├─ o1.wav, p1.wav, …
      └─ bgm.mp3
```

## ルール

- フォルダ名は、**動画 HTML のファイル名（拡張子なし）と同じ**にします
  - 例：`promo/ReiwaIE_Foo.html` の声は `voice/ReiwaIE_Foo/`
- 各フォルダには、台本 `narration.json` と、セリフごとの `<id>.wav` を置きます
  - ファイル名は固定で `narration.json`（動画名を付けない）
  - `<id>` は `narration.json` の `lines[].id` と一致させます
- 動画 HTML 側は、冒頭付近の定数 `VOICE_DIR` でフォルダを指定します

  ```js
  const VOICE_DIR = 'voice/ReiwaIE_Foo/';
  ```

- 声が読み込めなかった場合、ムービーはナレーションなし（BGM・効果音・字幕のみ）で再生されます
- `ReiwaIE_DebugRoute` だけは作り方が違います。HTML は Canvas に描く方式で、台本（字幕と開始秒）と描画コードを HTML の中に埋め込んでいます。このフォルダには、`narration.json`・`<id>.wav` に加えて、BGM と効果音を事前に合成した `bgm.mp3` を置いています

## narration.json の書き方

```json
{
  "movie": "ReiwaIE_Foo.html",
  "speaker": "ずんだもん",
  "credit": "VOICEVOX:ずんだもん",
  "lines": [
    { "id": "A1", "scene": "シーン名", "at": 5.5,
      "text": "画面に出す字幕", "say": "VOICEVOX に読ませる文章",
      "style": "ノーマル", "speed": 1.15, "pitch": 0, "intonation": 1.2 }
  ]
}
```

| 項目 | 意味 |
| --- | --- |
| `text` | 画面の字幕（の元） |
| `say` | VOICEVOX に読ませる文章。英字や数字は読み仮名に直す（例：`IE` → `アイイー`、`v4.0.0` → `バージョン4`） |
| `at` | ムービー内の開始秒。`ReiwaIE_History.html` のようにシーン表から自動で決める動画では参考値 |
| `style` | ずんだもんのスタイル名（下表） |
| `speed` / `pitch` / `intonation` / `volume` | 話速・音高・抑揚・音量。`volume` は再生時の音量倍率で、省略時は 1 |
| `credit` | エンドクレジットに表示する文字列。VOICEVOX の利用規約上、クレジット表記は必須です |

### ずんだもんのスタイル

| スタイル | ID |
| --- | ---: |
| ノーマル | 3 |
| あまあま | 1 |
| ツンツン | 7 |
| セクシー | 5 |
| ささやき | 22 |
| ヒソヒソ | 38 |
| ヘロヘロ | 75 |
| なみだめ | 76 |

※ ID は VOICEVOX ENGINE `0.25.2` 時点のものです。`http://127.0.0.1:50021/speakers` で確認できます。

## 声の作り方（VOICEVOX ENGINE）

VOICEVOX を起動しておくと、`http://127.0.0.1:50021` で ENGINE の API が使えます。セリフごとに次の 2 手順で wav を作ります。

1. `POST /audio_query?text=<say>&speaker=<スタイルID>` で合成用のクエリ(JSON)を受け取る
2. 受け取った JSON の `speedScale` / `pitchScale` / `intonationScale` を `narration.json` の値で上書きし、`POST /synthesis?speaker=<スタイルID>` に送る。返ってきた音声を `<id>.wav` として保存する

`ReiwaIE_History`・`ReiwaIE_CoreCompare`・`ReiwaIE_VbaCompare` では、間延びを防ぐため `prePhonemeLength = 0.05`、`postPhonemeLength = 0.1` にしています。

> [!WARNING]
> `ReiwaIE_History.html`・`ReiwaIE_CoreCompare.html`・`ReiwaIE_VbaCompare.html` は、各 wav の長さ（秒）を HTML 内の `VO` テーブルに持っていて、そこからシーンの尺を決めています。wav を作り直したときは、`VO` テーブルの `d`（秒）も更新してください。
