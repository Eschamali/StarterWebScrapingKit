Attribute VB_Name = "Test_WV2SharedBufferHLS"
'***************************************************************************************************
'   `WebView2SharedBuffer`の大容量テスト(HLS動画の保存、1bas完結)
'
'   流れ
'     1. JS  : master.m3u8 → 最高画質の映像 + 対応する音声 のプレイリストを解析し、
'              init(EXT-X-MAP) + 全セグメントをfetchしてメモリに保持。合計サイズを`window.__hlsInfo`へ
'     2. VBA : `jsEval`ポーリングでサイズを受け取り、映像+音声の合計で`CreateSharedBuffer`
'     3. VBA : `PostToScript` → JSが 映像 → 音声 の順に詰めて`releaseBuffer`
'     4. VBA : `get_Buffer`のポインターをそのまま`WriteFile`(0コピー)
'                映像 = get_Buffer           から videoBytes
'                音声 = get_Buffer + videoBytes から audioBytes
'
'   ※映像と音声は別ストリーム(fMP4)なので、ファイルも別々です(結合=muxはしない)
'   結果は %UserProfile%\Downloads\WV2SharedBufferTest\hls_result.txt (UTF-8) とイミディエイトウィンドウへ出力します
'***************************************************************************************************
Option Explicit

Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private Const GENERIC_WRITE         As Long = &H40000000
Private Const CREATE_ALWAYS         As Long = 2
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE  As LongPtr = -1

Private Const MASTER_URL As String = "https://video.twimg.com/amplify_video/2109714325640540161/pl/etWspoLgT2yO7oiN.m3u8"

Private m_log       As String
Private m_passCount As Long
Private m_failCount As Long



'***************************************************************************************************
'* 機能　　：HLS動画を、SharedBuffer経由でディスクへ保存します
'***************************************************************************************************
Public Sub Test_SharedBuffer_HLS()
    Dim outDir As String: outDir = Environ("UserProfile") & "\Downloads" & "\WV2SharedBufferTest"
    If Len(Dir$(outDir, vbDirectory)) = 0 Then MkDir outDir

    m_log = vbNullString
    m_passCount = 0
    m_failCount = 0

    Dim sb As WebView2SharedBuffer
    On Error GoTo Catch

    '1. WebView2起動(フォームは表示しない)
    Dim br As CDPBrowser
    Set br = WebView2Form.StartCDPModeWebView2

    Dim wv2 As CDPCoreViaWebView2
    Set wv2 = WebView2Form.ThisWebView2

    Dim ctx As CDPContext
    Set ctx = WebView2Form.ThisCDPContext

    LogLine "location = " & ctx.jsEval("location.href")
    LogLine "listener = " & ctx.jsEval(ListenerScript())

    '2. JSでダウンロード開始(完了を待たずに戻る) → 完了をポーリング
    Dim t0 As LongLong: t0 = GetTickCount64()
    LogLine "download = " & ctx.jsEval(DownloadScript(MASTER_URL))

    Dim info As String: info = WaitJs(ctx, "window.__hlsInfo", 180)
    Dim f() As String: f = Split(info, "|")
    If UBound(f) < 5 Then LogResult "1 ダウンロード", False, "JS=[" & info & "]": GoTo Finally

    'videoBytes|audioBytes|videoUrls|audioUrls|resolution|ms
    Dim videoBytes As LongLong: videoBytes = CLngLng(f(0))
    Dim audioBytes As LongLong: audioBytes = CLngLng(f(1))
    Dim totalBytes As LongLong: totalBytes = videoBytes + audioBytes
    LogResult "1 ダウンロード(JS)", (videoBytes > 0), "解像度=" & f(4) & " 映像=" & videoBytes & "B(" & f(2) & "本) 音声=" & audioBytes & "B(" & f(3) & "本) JS計測=" & f(5) & "ms VBA計測=" & (GetTickCount64() - t0) & "ms"

    '3. 合計サイズで確保 → JSへ → JSが詰める
    t0 = GetTickCount64()
    Set sb = wv2.CreateSharedBuffer(totalBytes)
    If sb Is Nothing Then LogResult "2 CreateSharedBuffer", False, "Nothing size=" & totalBytes: GoTo Finally

    ctx.jsEval "window.__sbRes = undefined; 0"
    If Not sb.PostToScript(SharedBuffer_ReadWrite, "{""mode"":""hls""}") Then LogResult "2 PostToScript", False, "失敗": GoTo Finally

    Dim res As String: res = WaitJs(ctx, "window.__sbRes", 60)
    Dim r() As String: r = Split(res, "|")
    Dim copied As Boolean: copied = (UBound(r) >= 1)
    If copied Then copied = (CLngLng(r(0)) = videoBytes) And (CLngLng(r(1)) = audioBytes)
    LogResult "2 Create+Post+JSコピー", copied, "JS=[" & res & "] size=" & sb.get_Size & " " & (GetTickCount64() - t0) & "ms"
    If Not copied Then GoTo Finally

    '4. ポインターからそのままWriteFile(映像 / 音声)
    Dim videoPath As String: videoPath = outDir & "\hls_video.mp4"
    Dim audioPath As String: audioPath = outDir & "\hls_audio.m4a"

    t0 = GetTickCount64()
    Dim okV As Boolean: okV = WritePointerToFile(videoPath, sb.get_Buffer, CLng(videoBytes))
    Dim okA As Boolean: okA = WritePointerToFile(audioPath, sb.get_Buffer + CLngPtr(videoBytes), CLng(audioBytes))
    Dim tWrite As LongLong: tWrite = GetTickCount64() - t0

    LogResult "3 get_Buffer→WriteFile", okV And okA And (FileLen(videoPath) = videoBytes) And (FileLen(audioPath) = audioBytes), _
              "video=" & FileLen(videoPath) & "B audio=" & FileLen(audioPath) & "B " & tWrite & "ms"

    '5. 先頭の箱が'ftyp'か(fMP4のinitセグメント)をIStream側から確認
    LogResult "4 映像の先頭がftyp", (BoxTypeAt(sb, 0) = "ftyp"), "box=" & BoxTypeAt(sb, 0)
    LogResult "4 音声の先頭がftyp", (BoxTypeAt(sb, videoBytes) = "ftyp"), "box=" & BoxTypeAt(sb, videoBytes)
    GoTo Finally

Catch:
    LogResult "RuntimeError", False, Err.Number & ": " & Err.Description
    Resume Finally

Finally:
    On Error Resume Next
    Set sb = Nothing
    Unload WebView2Form
    LogLine "=== PASS " & m_passCount & " / FAIL " & m_failCount & " ==="
    WriteLog outDir & "\hls_result.txt"
End Sub

'* MP4の箱(size 4byte + type 4byte)のtypeを、指定位置から読む
Private Function BoxTypeAt(sb As WebView2SharedBuffer, ByVal offset As LongLong) As String
    sb.Position = offset + 4
    Dim b() As Byte: b = sb.ReadBytes(4)
    BoxTypeAt = StrConv(b, vbUnicode)
End Function



'***************************************************************************************************
'                                   ■■■ JS ■■■
'***************************************************************************************************
'* 機能　　：HLSのダウンロードを開始するJSを返します(完了を待たずに'started'を返す)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：完了すると`window.__hlsInfo`に「映像バイト|音声バイト|映像URL数|音声URL数|解像度|ms」、
'            失敗時は「error:...」を置きます。データ本体は`window.__hls`に保持
'***************************************************************************************************
Private Function DownloadScript(ByVal masterUrl As String) As String
    Dim js As String
    js = "(() => {"
    js = js & "window.__hlsInfo = undefined; window.__hls = null;"
    js = js & "const master = '" & masterUrl & "';"
    js = js & "const abs = u => new URL(u, master).href;"
    js = js & "const get = async u => { const r = await fetch(u); if (!r.ok) throw new Error(r.status + ' ' + u); return r; };"
    js = js & "const text = async u => (await get(u)).text();"
    js = js & "const bin = async u => new Uint8Array(await (await get(u)).arrayBuffer());"
    'メディアプレイリスト → [init, seg1, seg2, ...]
    js = js & "const media = t => { const urls = []; const m = /#EXT-X-MAP:URI=""([^""]+)""/.exec(t); if (m) urls.push(m[1]);"
    js = js & "  for (const l of t.split(/\r?\n/)) { const s = l.trim(); if (s && !s.startsWith('#')) urls.push(s); } return urls.map(abs); };"
    '同時6本で取得(順番は保持)
    js = js & "const pool = async (urls, n) => { const out = new Array(urls.length); let i = 0;"
    js = js & "  const w = async () => { while (i < urls.length) { const k = i++; out[k] = await bin(urls[k]); } };"
    js = js & "  await Promise.all(Array.from({ length: n }, w)); return out; };"
    js = js & "const sum = a => a.reduce((s, x) => s + x.length, 0);"
    js = js & "(async () => {"
    js = js & "  const t0 = performance.now();"
    js = js & "  const mt = await text(master); const lines = mt.split(/\r?\n/);"
    'BANDWIDTH最大のバリアントを選ぶ(AVERAGE-BANDWIDTHは除外)
    js = js & "  let best = null;"
    js = js & "  for (let i = 0; i < lines.length; i++) { if (!lines[i].startsWith('#EXT-X-STREAM-INF:')) continue;"
    js = js & "    const bw = +((/[:,]BANDWIDTH=(\d+)/.exec(lines[i]) || [0, 0])[1]);"
    js = js & "    const res = (/RESOLUTION=([^,]+)/.exec(lines[i]) || [0, ''])[1];"
    js = js & "    const au = (/AUDIO=""([^""]+)""/.exec(lines[i]) || [0, ''])[1];"
    js = js & "    if (!best || bw > best.bw) best = { bw, res, au, uri: lines[i + 1].trim() }; }"
    'そのバリアントのAUDIOグループの音声プレイリスト
    js = js & "  let audioUri = '';"
    js = js & "  for (const l of lines) { if (best.au && l.startsWith('#EXT-X-MEDIA:') && l.includes('GROUP-ID=""' + best.au + '""')) { const m = /URI=""([^""]+)""/.exec(l); if (m) audioUri = m[1]; } }"
    js = js & "  const vUrls = media(await text(abs(best.uri)));"
    js = js & "  const aUrls = audioUri ? media(await text(abs(audioUri))) : [];"
    js = js & "  const v = await pool(vUrls, 6); const a = await pool(aUrls, 6);"
    js = js & "  window.__hls = { v, a };"
    js = js & "  window.__hlsInfo = [sum(v), sum(a), vUrls.length, aUrls.length, best.res, Math.round(performance.now() - t0)].join('|');"
    js = js & "})().catch(e => { window.__hlsInfo = 'error:' + e; });"
    js = js & "return 'started'; })()"
    DownloadScript = js
End Function

'* 機能　　：'sharedbufferreceived'のリスナー。mode=hls なら 映像 → 音声 の順に詰める
Private Function ListenerScript() As String
    Dim js As String
    js = "(() => {"
    js = js & "if (!window.chrome || !chrome.webview) return 'no chrome.webview';"
    js = js & "if (window.__sbHlsListener) return 'already';"
    js = js & "window.__sbHlsListener = true;"
    js = js & "chrome.webview.addEventListener('sharedbufferreceived', e => {"
    js = js & "  const info = e.additionalData || {};"
    js = js & "  if (info.mode !== 'hls') return;"
    js = js & "  const buf = e.getBuffer(); const u8 = new Uint8Array(buf);"
    js = js & "  let res = '';"
    js = js & "  try {"
    js = js & "    let off = 0;"
    js = js & "    for (const p of window.__hls.v) { u8.set(p, off); off += p.length; }"
    js = js & "    const vLen = off;"
    js = js & "    for (const p of window.__hls.a) { u8.set(p, off); off += p.length; }"
    js = js & "    res = vLen + '|' + (off - vLen);"
    js = js & "    window.__hls = null;"
    js = js & "  } catch (err) { res = 'error:' + err; }"
    js = js & "  chrome.webview.releaseBuffer(buf);"
    js = js & "  window.__sbRes = res;"
    js = js & "});"
    js = js & "return 'installed'; })()"
    ListenerScript = js
End Function

'* JS側の変数に値が入るまで待つ(タイムアウト時は空文字)
Private Function WaitJs(ctx As CDPContext, ByVal jsVar As String, ByVal timeoutSec As Double) As String
    Dim t0 As LongLong: t0 = GetTickCount64()
    Dim v As Variant
    Do
        v = ctx.jsEval(jsVar & " === undefined ? '' : String(" & jsVar & ")")
        If Not IsError(v) Then
            If Len(v) > 0 Then WaitJs = v: Exit Function
        End If
        DoEvents
    Loop Until GetTickCount64() - t0 > timeoutSec * 1000
End Function



'***************************************************************************************************
'                                   ■■■ 補助 ■■■
'***************************************************************************************************
'* ポインターの指す先を、そのままファイルへ書き出す
Private Function WritePointerToFile(ByVal filePath As String, ByVal pData As LongPtr, ByVal sizeBytes As Long) As Boolean
    Dim h As LongPtr
    h = CreateFileW(StrPtr(filePath), GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If h = INVALID_HANDLE_VALUE Then Exit Function

    Dim written As Long
    Dim ok As Boolean: ok = (WriteFile(h, pData, sizeBytes, written, 0) <> 0) And (written = sizeBytes)
    CloseHandle h

    WritePointerToFile = ok
End Function

Private Sub LogResult(ByVal testName As String, ByVal passed As Boolean, ByVal detail As String)
    If passed Then m_passCount = m_passCount + 1 Else m_failCount = m_failCount + 1
    LogLine IIf(passed, "[PASS] ", "[FAIL] ") & testName & " : " & detail
End Sub

Private Sub LogLine(ByVal s As String)
    Debug.Print s
    m_log = m_log & s & vbCrLf
End Sub

'* UTF-8で保存
Private Sub WriteLog(ByVal filePath As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    With st
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText m_log
        .SaveToFile filePath, 2
        .Close
    End With
End Sub
