VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2SharedBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'                   `ICoreWebView2SharedBuffer`の機能管理です
'   Page内で起こったデータをこの`SharedBuffer`を利用して、VBAにデータ参照させる仕組みです
'
'   CDP完結でも出来るっちゃできますが以下の問題があります
'   ・`Runtime.addBinding`による非同期受け取り：大容量には向かない。軽量データ向け
'   ・`Browser.downloadProgress`：blob URL を DLする仕組みを応用。大容量として行けるが、一旦Disk保存がいる
'
'           大容量をVBA内で完結したいというニッチな問題解決に使えるかもしれません
'
'   このClass単体では、終端目印みたいな、JavaScript側の書き込み完了サイン等は設けません。別途拡張Classか、CDPContext-BrowserEventsによる`Runtime.bindingCalled`受け取りを作ってもらってください。
'   あるいは、`SharedBuffer`の先頭にサイズ情報を載せる`jsEval`で同期処理による書き込み待機でも良いです。上記は、非同期でもらう場合です。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'COREWEBVIEW2_SHARED_BUFFER_ACCESS
Public Enum ExSharedBufferAccess
    SharedBuffer_ReadOnly = 0
    SharedBuffer_ReadWrite = 1
End Enum

'IStream(objidl.h)のvtableスロット番号
Private Const ISTREAM_READ  As Long = 3     'Read(void* pv, ULONG cb, ULONG* pcbRead)
Private Const ISTREAM_WRITE As Long = 4     'Write(const void* pv, ULONG cb, ULONG* pcbWritten)
Private Const ISTREAM_SEEK  As Long = 5     'Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)

'STREAM_SEEK
Private Const STREAM_SEEK_SET As Long = 0   'ストリーム先頭からの位置指定
Private Const STREAM_SEEK_CUR As Long = 1   '現在位置からの相対指定



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'SharedBuffer本体
Private m_pSharedBuffer As LongPtr  'ICoreWebView2SharedBuffer*(このClassが参照を1つ所有)
Private m_pData         As LongPtr  'BYTE*(共有メモリ先頭)
Private m_size          As LongLong '確保時のバイト数
Private m_pStream       As LongPtr  'IStream*(共有メモリ上の読み書き口。このClassが参照を1つ所有)

'ICoreWebView2_17:PostSharedBufferToScript を実行するためのポインター
Private m_pView         As LongPtr



'***************************************************************************************************
'                                   ■■■ 初期化系 ■■■
'***************************************************************************************************
'* 機能　　：`ICoreWebView2SharedBuffer`をアクティブにするための準備です
'---------------------------------------------------------------------------------------------------
'* 引数　　：pBuffer    ICoreWebView2Environment12.CreateSharedBuffer を呼び出した際の`ICoreWebView2SharedBuffer`ポインター
'            pView      ICoreWebView2_17:PostSharedBufferToScript を実行するための借用ポインター
'***************************************************************************************************
Friend Sub Activation(pBuffer As LongPtr, pView As LongPtr)
    '1. Class内変数に情報を確保
    m_pSharedBuffer = pBuffer
    m_pView = pView

    '2. 予め追加の必要情報も取っておく
    m_size = get_SizeInternal       'サイズ情報
    m_pData = get_BufferInternal    '先頭アドレス
    m_pStream = OpenStreamInternal  '読み書き用のIStream(1本だけ開いて使い回す)
End Sub

'***************************************************************************************************
'* 機能　　：共有メモリをJSへ渡します
'---------------------------------------------------------------------------------------------------
'* 引数　　：access                JS側の権限
'            additionalDataAsJson  JS側の`e.additionalData`になるJSON文字列(省略時はnull)
'* 返り値  ：成功可否
'---------------------------------------------------------------------------------------------------
'* 機能説明：JS側は`chrome.webview`の'sharedbufferreceived'で受け取れます
'* 注意事項：・受け取り側のリスナーは、これを実行前に登録しておくこと(Post時点でリスナーが無いと届かない)
'            ・トップレベルのページ(UserFormに写ってるWebページ)のみ有効です
'***************************************************************************************************
Public Function PostToScript(Optional ByVal access As ExSharedBufferAccess = SharedBuffer_ReadWrite, Optional ByVal additionalDataAsJson As String) As Boolean
    '実行
    Dim result As Long
    result = WebView2Thunks.dcf(m_pView, 116, "View17.PostSharedBufferToScript", m_pSharedBuffer, CLng(access), StrPtr(additionalDataAsJson))

    '結果を返却
    PostToScript = (result = 0)
End Function



'***************************************************************************************************
'                      ■■■ ADODB.Streamライクなメソッド(実体はIStream) ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] cursor操作
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：読み書きの現在位置(先頭からのバイト数)を取得/設定します
'---------------------------------------------------------------------------------------------------
'* 機能説明：`ADODB.Stream.Position`相当。中身はIStreamの`Seek`です
'            ・Get：`Seek(0, STREAM_SEEK_CUR)`で、IStreamが今いる位置を問い合わせる
'            ・Let：`Seek(位置, STREAM_SEEK_SET)`で、先頭からの位置へ移動する
'* 注意事項：・Close済み/Activation前は、Getは-1を返し、Letは何もしません
'            ・Letで0～Sizeの範囲外を指定した場合は、何もしません(位置は変わらない)
'            ・`WriteBytes`で書いた分だけ、位置は自動で進みます
'***************************************************************************************************
Public Property Get Position() As LongLong
    Position = -1
    If m_pStream = 0 Then Exit Property

    Dim newPosition As LongLong
    If WebView2Thunks.dcf(m_pStream, ISTREAM_SEEK, "Stream.Seek", CLngLng(0), STREAM_SEEK_CUR, VarPtr(newPosition)) = 0 Then Position = newPosition
End Property

Public Property Let Position(ByVal newValue As LongLong)
    If m_pStream = 0 Then Exit Property
    If newValue < 0 Or newValue > m_size Then Exit Property

    Dim newPosition As LongLong
    WebView2Thunks.dcf m_pStream, ISTREAM_SEEK, "Stream.Seek", newValue, STREAM_SEEK_SET, VarPtr(newPosition)
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 読み書き
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'* 機能　　：Byte配列を、現在位置(`Position`)から共有メモリへ書き込みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：Data    書き込むバイト配列
'* 返り値  ：成功可否(未初期化の配列、範囲外、Close済みの場合はFalse。その場合は何も書かない)
'---------------------------------------------------------------------------------------------------
'* 機能説明：`ADODB.Stream.Write`相当。コピー自体はWebView2側のネイティブ処理(IStream.Write)で、
'            VBA側でループもRtlMoveMemoryも使いません。書いた分だけ`Position`が進みます
'* 注意事項：書き込み位置を指定したい場合は、先に`Position`を設定してください
'            例) sb.Position = 0: sb.WriteBytes header: sb.WriteBytes body  ' headerの直後にbodyが続く
'***************************************************************************************************
Public Function WriteBytes(Data() As Byte) As Boolean
    '1. 書き込み口が無ければ即抜け(Activation前、またはClose済み)
    If m_pStream = 0 Then Exit Function

    '2. データ容量を確認(未初期化の配列はFalse)
    Dim WriteSize As LongLong: WriteSize = ByteArrayLength(Data)
    If WriteSize = 0 Then Exit Function

    '3. 現在位置から書いて、共有メモリの範囲に収まるか確認(はみ出すなら1バイトも書かない)
    Dim currentPosition As LongLong: currentPosition = Position
    If currentPosition < 0 Or currentPosition + WriteSize > m_size Then Exit Function

    '4. まとめて書き込み(位置は書いた分だけ進む)
    Dim hr As Long, written As Long
    hr = WebView2Thunks.dcf(m_pStream, ISTREAM_WRITE, "Stream.Write", VarPtr(Data(LBound(Data))), CLng(WriteSize), VarPtr(written))

    '5. 全量書けたかで判定
    WriteBytes = (hr = 0 And written = WriteSize)
End Function

'***************************************************************************************************
'* 機能　　：現在位置(`Position`)から、共有メモリの内容をByte配列として読み出します
'---------------------------------------------------------------------------------------------------
'* 引数　　：NumBytes  読み出すバイト数(省略/負数なら、現在位置から終端まで全部)
'* 返り値  ：読み出したバイト配列(読めるものが無い場合は未初期化の配列)
'---------------------------------------------------------------------------------------------------
'* 機能説明：`ADODB.Stream.Read`相当。コピー自体はWebView2側のネイティブ処理(IStream.Read)です。
'            読んだ分だけ`Position`が進みます
'            ・「終端」は`SetEOS`で記録した位置です(未設定なら確保サイズの末尾=全データ)
'            ・終端までの残りより多く指定した場合は、残りの分だけ返します(`ADODB.Stream`と同じ)
'            ・EOS(終端)で呼んだ場合、Close済み/Activation前の場合は、未初期化の配列を返します
'* 注意事項：・VBAのByte配列として丸ごと確保するため、数百MB以上を一度に読むとメモリを圧迫します。
'              大容量は`NumBytes`で区切って読むか、別の手段(ファイルへ直接書き出す等)を検討してください
'***************************************************************************************************
Public Function ReadBytes(Optional NumBytes As Long) As Byte()
    Dim result() As Byte
    ReadBytes = result

    '1. 読み出し口が無ければ即抜け(Activation前、またはClose済み)
    If m_pStream = 0 Then Exit Function

    '2. 省略の場合は、末尾まで読むようにする
    If NumBytes <= 0 Then NumBytes = CLng(m_size - Position)  '省略による大きすぎるサイズの場合はここでオーバフローエラーになります。引数で調整してください

    '3. 既に末尾の場合は抜ける
    If NumBytes <= 0 Then Exit Function

    '4. 受け皿を用意して、まとめて読み出し(位置は読んだ分だけ進む)
    ReDim result(0 To NumBytes - 1)
    Dim hr As Long, readCount As Long
    hr = WebView2Thunks.dcf(m_pStream, ISTREAM_READ, "Stream.Read", VarPtr(result(0)), NumBytes, VarPtr(readCount))

    '5. 失敗(HRESULTが負)なら未初期化のまま返す。S_FALSE(=1)は「要求より少なく読めた」なので成功扱い
    If hr < 0 Or readCount <= 0 Then Exit Function

    '6. 要求より少なかった場合は、読めた分に切り詰める
    If readCount < NumBytes Then ReDim Preserve result(0 To readCount - 1)
    ReadBytes = result
End Function

'***************************************************************************************************
'* 機能　　：バイト配列の要素数を返します(未初期化なら0)
'***************************************************************************************************
Private Function ByteArrayLength(Data() As Byte) As LongLong
    On Error GoTo NotInitialized
    ByteArrayLength = UBound(Data) - LBound(Data) + 1
    Exit Function

NotInitialized:
    ByteArrayLength = 0
End Function



'***************************************************************************************************
'                              ■■■ ICoreWebView2SharedBuffer 一式 ■■■
'***************************************************************************************************
'* 機能　　：共有バッファのサイズをバイト数で受け取ります
'---------------------------------------------------------------------------------------------------
'* 注意事項：基本的には、`Activation`からの初回呼び出しのみに使います。値の確認は次の`Property`を使ってください
'***************************************************************************************************
Private Function get_SizeInternal() As LongLong
    get_SizeInternal = WebView2Thunks.GetLongLongProperty(m_pSharedBuffer, 3, "SharedBuffer.get_Size")
End Function
Property Get get_Size() As LongLong
    get_Size = m_size
End Property

'***************************************************************************************************
'* 機能　　：共有バッファの先頭を指すメモリポインタを受け取ります
'---------------------------------------------------------------------------------------------------
'* 注意事項：基本的には、`Activation`からの初回呼び出しのみに使います。値の確認は次の`Property`を使ってください
'***************************************************************************************************
Private Function get_BufferInternal() As LongPtr
    get_BufferInternal = WebView2Thunks.GetPtrProperty(m_pSharedBuffer, 4, "SharedBuffer.get_Buffer")
End Function
Property Get get_Buffer() As LongPtr
    get_Buffer = m_pData
End Property

'***************************************************************************************************
'* 機能　　：共有バッファ上に開いた IStream インターフェイスへのポインタを受け取ります
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`Activation`で1回だけ呼び、`m_pStream`に控えて使い回します(外部には公開しない)
'            ・新しい参照が返るため、`CLOSE_`でReleaseすること
'***************************************************************************************************
Private Function OpenStreamInternal() As LongPtr
    OpenStreamInternal = WebView2Thunks.GetPtrProperty(m_pSharedBuffer, 5, "SharedBuffer.OpenStream")
End Function

'***************************************************************************************************
'* 機能　　：共有バッファの基になるファイルマッピングオブジェクトのハンドルを受け取ります
'---------------------------------------------------------------------------------------------------
'* 注意事項：毎回、Vtableからの実行になるため、呼び出し側で確保するように
'***************************************************************************************************
Public Function get_FileMappingHandle() As LongPtr
    get_FileMappingHandle = WebView2Thunks.GetPtrProperty(m_pSharedBuffer, 6, "SharedBuffer.get_FileMappingHandle")
End Function

'***************************************************************************************************
'* 機能　　：共有バッファを閉じ、関連するリソースを解放します
'---------------------------------------------------------------------------------------------------
'* 機能説明：IStream(子) → SharedBuffer本体(親) の順に手放します
'* 注意事項：以降、`m_pData`(先頭アドレス)は無効になるため0に戻します
'***************************************************************************************************
Private Sub CLOSE_()
    '1. 読み書き用のIStreamを手放す
    If m_pStream <> 0 Then WebView2Thunks.ComRelease m_pStream
    m_pStream = 0

    '2. 共有メモリを閉じ、このClassが所有していた参照を手放す
    If m_pSharedBuffer <> 0 Then
        WebView2Thunks.dcf m_pSharedBuffer, 7, "SharedBuffer.Close"
        WebView2Thunks.ComRelease m_pSharedBuffer
    End If
    m_pSharedBuffer = 0

    '3. 控えていた情報も無効化
    m_pData = 0
    m_size = 0
End Sub



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
Private Sub Class_Terminate()
    CLOSE_
End Sub
