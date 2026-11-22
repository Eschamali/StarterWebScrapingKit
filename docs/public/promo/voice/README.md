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
├─ ReiwaIE_Policy.html
├─ ReiwaIE_WebView2Story.html
├─ ReiwaIE_GetTcpTable.html
├─ ReiwaIE_OfficeScripts.html
├─ ReiwaIE_Policy2.html
├─ ReiwaIE_BiDiStory2.html
├─ ReiwaIE_TwoWay.html
└─ voice/
   ├─ README.md            … このファイル
   ├─ ReiwaIE_60s/         … ReiwaIE_60s.html の声
   │  ├─ narration.json
   │  └─ A1.wav, A2.wav, …
   ├─ ReiwaIE_History/     … ReiwaIE_History.html の声と BGM
   │  ├─ narration.json
   │  ├─ op1.wav, yo1.wav, …
   │  └─ bgm.mp3
   ├─ ReiwaIE_CoreCompare/ … ReiwaIE_CoreCompare.html の声と BGM
   │  ├─ narration.json
   │  ├─ op1.wav, tr1.wav, …
   │  └─ bgm.mp3
   ├─ ReiwaIE_VbaCompare/  … ReiwaIE_VbaCompare.html の声と BGM
   │  ├─ narration.json
   │  ├─ op1.wav, ln1.wav, …
   │  └─ bgm.mp3
   ├─ ReiwaIE_DebugRoute/  … ReiwaIE_DebugRoute.html の声と BGM
   │  ├─ narration.json
   │  ├─ o1.wav, p1.wav, …
   │  └─ bgm.mp3
   ├─ ReiwaIE_BiDiStory/   … ReiwaIE_BiDiStory.html の声と BGM
   │  ├─ narration.json
   │  ├─ o1.wav, s1.wav, …
   │  └─ bgm.mp3
   ├─ ReiwaIE_Policy/      … ReiwaIE_Policy.html の声（2人の会話）と BGM
   │  ├─ narration.json
   │  ├─ op1.wav, wl1.wav, …
   │  └─ bgm.mp3
   ├─ ReiwaIE_WebView2Story/ … ReiwaIE_WebView2Story.html の声と BGM
   │  ├─ narration.json
   │  ├─ fo1.wav, ac1.wav, …, ts1.wav
   │  └─ bgm.mp3
   ├─ ReiwaIE_GetTcpTable/ … ReiwaIE_GetTcpTable.html の声（2人の会話）と BGM
   │  ├─ narration.json
   │  ├─ pr1.wav, sn1.wav, …, ro8.wav
   │  └─ bgm.mp3
   ├─ ReiwaIE_OfficeScripts/ … ReiwaIE_OfficeScripts.html の声（2人の会話）と BGM、実機の録画のクリップ
   │  ├─ narration.json
   │  ├─ op1.wav, in1.wav, …, cl11.wav
   │  ├─ bgm.mp3
   │  └─ clips/
   │     └─ desk.mp4, web.mp4, help.mp4, mirror.mp4
   ├─ ReiwaIE_Policy2/     … ReiwaIE_Policy2.html の声（2人の会話）と BGM、実機の録画のクリップと静止画
   │  ├─ narration.json
   │  ├─ op1.wav, rp1.wav, …, en14.wav
   │  ├─ bgm.mp3
   │  └─ clips/
   │     ├─ bug.mp4, cdp.mp4, canary.mp4
   │     └─ reg.jpg, console.jpg, …（録画から切り抜いた静止画）
   ├─ ReiwaIE_BiDiStory2/  … ReiwaIE_BiDiStory2.html の声（2人の会話）と BGM、記事の画面写真
   │  ├─ narration.json
   │  ├─ op1.wav, nt1.wav, …, en6.wav
   │  ├─ bgm.mp3
   │  └─ img/
   │     └─ chrome.png, edge.png, edge2.png
   └─ ReiwaIE_TwoWay/      … ReiwaIE_TwoWay.html の声（2人の会話）と BGM、リリースノートと Power Query の画面
      ├─ narration.json
      ├─ in1.wav, pl1.wav, …, g14.wav
      ├─ bgm.mp3
      └─ img/
         └─ intro.png, powerquery.png, debug.png
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
- `ReiwaIE_History`・`ReiwaIE_DebugRoute`・`ReiwaIE_BiDiStory`・`ReiwaIE_CoreCompare`・`ReiwaIE_VbaCompare`・`ReiwaIE_Policy`・`ReiwaIE_WebView2Story`・`ReiwaIE_GetTcpTable`・`ReiwaIE_OfficeScripts`・`ReiwaIE_Policy2`・`ReiwaIE_BiDiStory2`・`ReiwaIE_TwoWay` は作り方が違います。HTML は Canvas に描く方式で、台本（字幕と開始秒）と描画コードを HTML の中に埋め込んでいます。このフォルダには、`narration.json`・`<id>.wav` に加えて `bgm.mp3` を置いています
  - `ReiwaIE_DebugRoute` の `bgm.mp3` は、BGM と効果音を事前に合成したものです
  - `ReiwaIE_History`・`ReiwaIE_BiDiStory`・`ReiwaIE_CoreCompare`・`ReiwaIE_VbaCompare`・`ReiwaIE_Policy`・`ReiwaIE_WebView2Story`・`ReiwaIE_GetTcpTable`・`ReiwaIE_OfficeScripts`・`ReiwaIE_Policy2`・`ReiwaIE_BiDiStory2`・`ReiwaIE_TwoWay` の `bgm.mp3` は曲だけです。効果音は mp3 に入れず、HTML の中で Web Audio API を使って鳴らしています。絵と同じ時刻表から鳴らすので、効果音の時刻を直したときは HTML を作り直すだけで済みます
  - `ReiwaIE_BiDiStory.html` は、記事の画像を `img/story/` から相対パス（`../img/story/1.png` など）で読み込みます。画像を移動・改名したときは、HTML 内の `IMGSRC` も直してください
  - `ReiwaIE_Policy.html` も、記事の画像を `img/Policy/` から相対パス（`../img/Policy/1.png` など）で読み込みます。画像を移動・改名したときは、HTML 内の `IMG_BASE` も直してください
  - `ReiwaIE_WebView2Story.html` は、記事の画像を `img/` の直下と `img/story/human/` から相対パス（`../img/化石コントロール.png` など）で読み込みます。画像を移動・改名したときは、HTML 内の `IMG_FILES` も直してください
- `ReiwaIE_Policy` は2人の会話です。`narration.json` の各行に話し手（`speaker`）を書いています。春日部つむぎ（VOICEVOX のスタイル ID は 8）が案内役、ずんだもんが聞き役です
- `ReiwaIE_GetTcpTable` は2人の会話です。東北イタコ（VOICEVOX のスタイル ID は 109）が案内役、ずんだもんが一緒に調べる聞き役です。記事の画像は使わず、絵はすべて Canvas で描いています。`ep1`〜`ep27`（`ep12` はなし）は、記事の「後日談: 本当にしれっと直った」の声です
- `ReiwaIE_OfficeScripts` は2人の会話です。春日部つむぎ（VOICEVOX のスタイル ID は 8）が案内役、ずんだもんが聞き役です。最後の `cl0`〜`cl11` は、記事の「CDPは、VBA自身の機能拡張でもあった」を要約したものです
  - `clips/` の mp4 は、`img/breakthrough/` にある実機の録画から、使う区間だけを切り出したものです（音なし・15fps・横 1920）。HTML の中の `CLIP_DIR` で場所を指定し、video 要素で読み込んで Canvas に描きます。再生中は流し、止めているときやシークでは、その時刻のコマに合わせます
  - タスクマネージャーのスクリーンショット（`img/breakthrough/1.png`）は、HTML の中に埋め込んでいます
- `ReiwaIE_Policy2` は2人の会話です。春日部つむぎ（VOICEVOX のスタイル ID は 8）が捜査の主任、ずんだもんが見習いの聞き役です。記事「ポリシーの壁」の第2話で、絵は Canvas で描いた捜査ボードです
  - `clips/` の mp4 は、実機の録画3本から作ったものです（音なし・15fps・横 1920）。PAD が1分黙る様子（`bug`）、同じ PC でこのツールの CDP が動く様子（`cdp`）、Edge Canary で PAD が動く様子（`canary`）です。録画の中のブラウザには、よそのサイトの絵が映っているため、ページの部分をぼかしています
  - `clips/` の jpg は、同じ録画から切り抜いた静止画です。HTML の中の `CLIP_DIR` で場所を指定しています
- `ReiwaIE_BiDiStory2` は2人の会話です。四国めたん（VOICEVOX のスタイル ID は 2。ささやき・あまあまの行は 36・0）が「Google社員みたいなもの」を名乗る現場の案内役、ずんだもんが見学者です。記事「BiDi 登場秘話・その後」の映像で、絵は Canvas で描いた工事現場（Chrome 本体ビルの断面）です
  - 記事は日付順に記録を書き足していく形です。映像も、記録ごとに章（工区）をまとめ、工事の状態（済み・工事中・試作のみ など）は日付つきの表から描いています。記録が増えたら、その記録の声（新しい `<id>.wav`）と章を足し、`narration.json` と HTML を作り直します。前の記録の章は、その日の状態のまま描かれます
  - `img/` の png は、記事の画面写真（Chrome Canary と Edge Canary に `/bidi-mapper.js` を送った画面）です。HTML の中の `IMG_DIR` で場所を指定しています
- `ReiwaIE_TwoWay` は2人の会話です。春歌ナナ（VOICEVOX のスタイル ID は 54）が輪廻の案内人、ずんだもんが聞き役です。記事「Web とローカルの双方向やり取り、30年の輪廻」の映像で、「このツールの場合」の補足はガイド「ブラウザと VBA の双方向やり取り」から取っています
  - `img/intro.png` は v4.2.2 のリリースノート、`img/powerquery.png` は Power Query エディターに IE の F12 開発者ツールを開いた画面、`img/debug.png` はリモート デバッグの許可を求める Edge のダイアログです（後ろのページは映さず、ダイアログだけを切り抜いて使います）。HTML の中の `IMG_DIR` で場所を指定しています
  - HTA のコード（`ForAI/材料/SampleHTA.hta` の一部）は、画像ではなく HTML の中に文字で書いています
- `ReiwaIE_WebView2Story` は、ずんだもんのひとり語りです。最後の4行（`ts1`〜`ts4`）だけ春日部つむぎがささやきます。つむぎにはささやきのスタイルがないため、合成した声に ffmpeg で響きと1オクターブ下の薄い重ねを足しています

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
| `at` | ムービー内の開始秒 |
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

`ReiwaIE_History`・`ReiwaIE_CoreCompare`・`ReiwaIE_VbaCompare`・`ReiwaIE_Policy`・`ReiwaIE_WebView2Story` では、間延びを防ぐため `prePhonemeLength = 0.05`、`postPhonemeLength = 0.12` にしています。
