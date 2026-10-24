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
├─ ReiwaIE_BiDiStory.html
└─ voice/
   ├─ README.md            … このファイル
   ├─ ReiwaIE_60s/         … ReiwaIE_60s.html の声
   │  ├─ narration.json
   │  └─ A1.wav, A2.wav, …
   ├─ ReiwaIE_History/     … ReiwaIE_History.html の声
   │  ├─ narration.json
   │  └─ P1.wav, P2.wav, …
   ├─ ReiwaIE_CoreCompare/ … ReiwaIE_CoreCompare.html の声と BGM
   │  ├─ narration.json
   │  ├─ op1.wav, tr1.wav, …
   │  └─ bgm.mp3
   ├─ ReiwaIE_VbaCompare/  … ReiwaIE_VbaCompare.html の声
   │  ├─ narration.json
   │  └─ O1.wav, O2.wav, …
   ├─ ReiwaIE_DebugRoute/  … ReiwaIE_DebugRoute.html の声と BGM
   │  ├─ narration.json
   │  ├─ o1.wav, p1.wav, …
   │  └─ bgm.mp3
   └─ ReiwaIE_BiDiStory/   … ReiwaIE_BiDiStory.html の声と BGM
      ├─ narration.json
      ├─ o1.wav, s1.wav, …
      └─ bgm.mp3
```

## ルール

- フォルダ名は、**動画 HTML のファイル名（拡張子なし）と同じ**にします
  - 例：`promo/ReiwaIE_Foo.html` の声は `voice/ReiwaIE_Foo/`
- 各フォルダには、台本 `narration.json` と、セリフごとの `<id>.wav` を置きます
  - ファイル名は固定で `narration.json`（動画名を付けない）
  - `<id>` は `narration.json` の `lines[].id` と一致させます
  - 作り直すときは、前の版のファイルと大文字・小文字だけが違う名前（`A1.wav` と `a1.wav` など）にしません。Windows と、大文字・小文字を区別しない設定の git では同じファイルとして扱われ、公開先（大文字・小文字を区別する）で声が読めなくなるためです
- 動画 HTML 側は、冒頭付近の定数 `VOICE_DIR` でフォルダを指定します

  ```js
  const VOICE_DIR = 'voice/ReiwaIE_Foo/';
  ```

- 声が読み込めなかった場合、ムービーはナレーションなし（BGM・効果音・字幕のみ）で再生されます
- `ReiwaIE_DebugRoute`・`ReiwaIE_BiDiStory`・`ReiwaIE_CoreCompare` は作り方が違います。HTML は Canvas に描く方式で、台本（字幕と開始秒）と描画コードを HTML の中に埋め込んでいます。このフォルダには、`narration.json`・`<id>.wav` に加えて `bgm.mp3` を置いています
  - `ReiwaIE_DebugRoute` の `bgm.mp3` は、BGM と効果音を事前に合成したものです
  - `ReiwaIE_BiDiStory` と `ReiwaIE_CoreCompare` の `bgm.mp3` は曲だけです。効果音は mp3 に入れず、HTML の中で Web Audio API を使って鳴らしています。絵と同じ時刻表から鳴らすので、効果音の時刻を直したときは HTML を作り直すだけで済みます
  - `ReiwaIE_BiDiStory.html` は、記事の画像を `img/story/` から相対パス（`../img/story/1.png` など）で読み込みます。画像を移動・改名したときは、HTML 内の `IMGSRC` も直してください

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

`ReiwaIE_History`・`ReiwaIE_VbaCompare` では、間延びを防ぐため `prePhonemeLength = 0.05`、`postPhonemeLength = 0.1` にしています。

> [!WARNING]
> `ReiwaIE_History.html`・`ReiwaIE_VbaCompare.html` は、各 wav の長さ（秒）を HTML 内の `VO` テーブルに持っていて、そこからシーンの尺を決めています。wav を作り直したときは、`VO` テーブルの `d`（秒）も更新してください。
