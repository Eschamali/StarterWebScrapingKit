# Excel単独で「真のWebView2」を完全制御する

> さぁ、ここが本当の目的。  
> Excelだけで、しかもPowershellのような外部source依存もなしで、真WebView2をUserformに。

ついに辿り着きました。これが本プロジェクトの真の目的であり、到達点です。  
**外部プロセス（PowerShellなど）に頼らず、Excel VBAのメモリ空間上だけで WebView2 を直接起動・制御します。**

> ふはははははは！！！Excel VBAのユーザーフォーム上でWebView2を動作させて、イベントを検知することに成功したぞ  
> — [たーぼー（インコ） @fenblen_puyo](https://twitter.com/fenblen_puyo/status/2032821182924468312)

一見するとLv.1（Edge埋め込み）と似ていますが、タスクマネージャーを見ればその違いは一目瞭然です。

![Excel直からWebView2](/img/Excel直からWebView2.png)

*▲ Excel.exe の配下に直接 WebView2 プロセスが生成されています*

## 禁断の魔導：VBAによるCOM直接制御の仕組み

なぜ、これまでこれが不可能だと思われていたのか。それは WebView2 が **IUnknown ベース** という、VBAにとっては非常に「扱いづらい」設計になっているからです。

::: info
通常、VBAで `Object` として扱えるものは **IDispatch** という「親切な案内板」を持っています。しかし、WebView2にはそれがありません。
:::

### 1. `DispCallFunc` による関数の強行突破

「オブジェクト.メソッド」という通常の呼び出しができないため、Windows APIの `DispCallFunc` を使用します。  
これは、メモリ上の「関数の住所（vtableのインデックス）」を直接指定して実行する、いわば **VBA界の狙撃術**です。

### 2. `vtable` 偽造：自作のCOMオブジェクトを作る

WebView2は、処理が終わると「終わったよ！」とコールバック（通知）を返してきます。この通知を受け取るには「WebView2が理解できる形式のオブジェクト」である必要があります。  
VBAの **AddressOf** で取得した関数のポインタを構造体に詰め込み、メモリ上に「COMオブジェクトのフリをしたデータ」を構築（vtable偽造）することで、WebView2からの通信を直接受け止めます。

::: tip 処理のリレー（非同期通信）
VBAが制御を投げる → WebView2が処理 → 偽造オブジェクト経由でVBAのハンドラを叩く → VBAが次の指示を出す……という、極めて高度な連携によって動作しています。
:::

## 実装のエッセンス

::: details vtable構築の核心部分

```vb
' COM vtable 偽造オブジェクトの定義
Private Type VtblObj
    pVTable As LongPtr
End Type

' 4つのエントリ（QI, AddRef, Release, Invoke）を持つvtable
Private Type VtblData4
    fn(0 To 3) As LongPtr
End Type

' 実行時にVBA標準モジュールのメソッドを AddressOf で登録
Private Sub BuildVtables()
    Dim envFn()  As LongPtr
    WV2_FillFunctionPointers envFn

    m_EnvHandlerVT.fn(0) = envFn(0) ' QueryInterface
    m_EnvHandlerVT.fn(1) = envFn(1) ' AddRef
    m_EnvHandlerVT.fn(2) = envFn(2) ' Release
    m_EnvHandlerVT.fn(3) = envFn(3) ' Invoke (本命のコールバック)

    ' これがWebView2側に渡す「偽オブジェクト」のポインタになる
    m_EnvHandlerThis.pVTable = VarPtr(m_EnvHandlerVT.fn(0))
End Sub
```

:::

正直に申し上げて、これは **VBAの限界を突破した「ハック」** に近いです。  
しかし、この手法をマスターすれば、TLBや外部DLLに一切頼ることなく、最新のブラウザエンジンをExcelのUserFormに完全に支配下に置くことができます。

::: warning
仕組みの探求の覚悟があるなら、ぜひリポジトリを覗いてみてください：  
[WebView2-For-Excel-VBA (GitHub)](https://github.com/tarboh/WebView2-For-Excel-VBA)
:::

最後に、このロマンな技術テクニックを公開、教えてくれた「たーぼー(インコ)」さんに感謝します。

## 次へ

- [総括：3つの手法の比較](./summary)
- [PowerShell 経由](./powershell)
- [はじめに](./intro)
