Attribute VB_Name = "Test_WebView2SharedBuffer"
'***************************************************************************************************
'   `WebView2SharedBuffer`の単体テストです(1bas完結)
'
'   JS側の完了サインは`Runtime.bindingCalled`ではなく、`jsEval`のポーリングで待ちます
'   (JSのリスナーが`window.__sbRes`へ結果を置く → VBAが拾う)
'
'   T01 Create直後の状態      : get_Size / get_Buffer / get_FileMappingHandle / Position
'   T02 Position              : 範囲内は移動、範囲外(負数/Size超)は無視
'   T03 VBA内の読み書き        : WriteBytes → Position → ReadBytes、はみ出し/未初期化/末尾
'   T04 additionalData省略     : Postが通り、JS側で`e.additionalData`が空(実測ではundefined)になるか
'   T05 VBA → JS (ReadOnly)   : VBAがUTF-8で書いた内容をJSが読めるか
'   T06 JS → VBA (ReadWrite)  : JSが書いた内容をVBAが読めるか / JSの書き込みでPositionが動かないか
'   T07 同じバッファで往復      : VBA書き込み → JSが読んで続きに書く → VBAが続きを読む
'   T08 大容量(100MB)          : JS書き込み → ReadBytes一括/分割 → get_Bufferポインターから直接WriteFile
'   T09 確保と解放の繰り返し    : Class_Terminateでの解放が問題なく回るか
'   T10 サイズ上限             : 2GB(2^31)は確保できない(Nothing)こと
'
'   結果は %TEMP%\WV2SharedBufferTest\result.txt (UTF-8) とイミディエイトウィンドウへ出力します
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

Private Const LARGE_SIZE As Long = 100 * 1024& * 1024&  '100MB
Private Const CHUNK_SIZE As Long = 16 * 1024& * 1024&   '分割読み出しの単位

Private m_log       As String
Private m_passCount As Long
Private m_failCount As Long
Private m_conv      As CharacterCodeConversion



'***************************************************************************************************
'* 機能　　：T01～T10をまとめて実行します
'***************************************************************************************************
Public Sub Test_SharedBuffer_All()
    Dim outDir As String: outDir = Environ$("TEMP") & "\WV2SharedBufferTest"
    If Len(Dir$(outDir, vbDirectory)) = 0 Then MkDir outDir

    m_log = vbNullString
    m_passCount = 0
    m_failCount = 0
    Set m_conv = New CharacterCodeConversion

    On Error GoTo Catch

    '1. WebView2起動(フォームは表示しない)
    Dim br As CDPBrowser
    Set br = WebView2Form.StartCDPModeWebView2

    Dim wv2 As CDPCoreViaWebView2
    Set wv2 = WebView2Form.ThisWebView2

    Dim ctx As CDPContext
    Set ctx = WebView2Form.ThisCDPContext

    '2. JS側の受け口を用意
    LogLine "typeof chrome.webview = " & ctx.jsEval("typeof (window.chrome && window.chrome.webview)")
    LogLine "listener = " & ctx.jsEval(ListenerScript())

    '3. 各テスト
    Test01_Create wv2
    Test02_Position wv2
    Test03_VbaReadWrite wv2
    Test04_NoAdditionalData wv2, ctx
    Test05_VbaToJs wv2, ctx
    Test06_JsToVba wv2, ctx
    Test07_PingPong wv2, ctx
    Test08_Large wv2, ctx, outDir
    Test09_Lifecycle wv2
    Test10_SizeLimit wv2
    GoTo Finally

Catch:
    LogResult "RuntimeError", False, Err.Number & ": " & Err.Description
    Resume Finally

Finally:
    On Error Resume Next
    Unload WebView2Form
    LogLine "=== PASS " & m_passCount & " / FAIL " & m_failCount & " ==="
    WriteLog outDir & "\result.txt"
End Sub



'***************************************************************************************************
'                                   ■■■ 各テスト ■■■
'***************************************************************************************************
'* T01: Create直後の状態
Private Sub Test01_Create(wv2 As CDPCoreViaWebView2)
    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(4096))
    If sb Is Nothing Then LogResult "T01 Create", False, "Nothing が返った": Exit Sub

    Dim hMap As LongPtr: hMap = sb.get_FileMappingHandle
    Dim ok As Boolean
    ok = (sb.get_Size = 4096) And (sb.get_Buffer <> 0) And (hMap <> 0) And (sb.Position = 0)

    LogResult "T01 Create", ok, "Size=" & sb.get_Size & " Buffer=" & Hex$(sb.get_Buffer) & " MappingHandle=" & Hex$(hMap) & " Position=" & sb.Position
End Sub

'* T02: Position
Private Sub Test02_Position(wv2 As CDPCoreViaWebView2)
    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(4096))
    If sb Is Nothing Then LogResult "T02 Create", False, "Nothing が返った": Exit Sub

    Dim ok As Boolean: ok = True
    Dim detail As String

    sb.Position = 100
    ok = ok And (sb.Position = 100): detail = detail & "Let100->" & sb.Position

    sb.Position = 4096          '末尾ちょうどはOK
    ok = ok And (sb.Position = 4096): detail = detail & " Let4096->" & sb.Position

    sb.Position = 4097          '範囲外は無視
    ok = ok And (sb.Position = 4096): detail = detail & " Let4097->" & sb.Position

    sb.Position = -1            '負数も無視
    ok = ok And (sb.Position = 4096): detail = detail & " Let-1->" & sb.Position

    sb.Position = 0
    ok = ok And (sb.Position = 0): detail = detail & " Let0->" & sb.Position

    LogResult "T02 Position", ok, detail
End Sub

'* T03: VBA内の読み書き
Private Sub Test03_VbaReadWrite(wv2 As CDPCoreViaWebView2)
    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(4096))
    If sb Is Nothing Then LogResult "T03 Create", False, "Nothing が返った": Exit Sub

    Dim a() As Byte: a = m_conv.BytesFromString("ABC")
    Dim b() As Byte: b = m_conv.BytesFromString("日本語")

    '1. 連続書き込み(位置が進む)
    Dim ok1 As Boolean
    ok1 = sb.WriteBytes(a) And sb.WriteBytes(b) And (sb.Position = 12)
    LogResult "T03-1 WriteBytes連続", ok1, "Position=" & sb.Position

    '2. 先頭に戻して読む
    sb.Position = 0
    Dim r1() As Byte: r1 = sb.ReadBytes(3)
    Dim pos1 As LongLong: pos1 = sb.Position
    Dim r2() As Byte: r2 = sb.ReadBytes(9)
    Dim s1 As String: s1 = m_conv.BytesToString(r1)
    Dim s2 As String: s2 = m_conv.BytesToString(r2)
    LogResult "T03-2 ReadBytes", (s1 = "ABC") And (pos1 = 3) And (s2 = "日本語") And (sb.Position = 12), "[" & s1 & "] pos=" & pos1 & " [" & s2 & "] pos=" & sb.Position

    '3. はみ出す書き込みは1バイトも書かず、位置も動かない
    Dim ten(0 To 9) As Byte
    sb.Position = 4090
    Dim wr As Boolean: wr = sb.WriteBytes(ten)
    LogResult "T03-3 はみ出し書き込み", (Not wr) And (sb.Position = 4090), "WriteBytes=" & wr & " Position=" & sb.Position

    '4. 未初期化の配列はFalse
    Dim emptyArr() As Byte
    LogResult "T03-4 未初期化配列", Not sb.WriteBytes(emptyArr), "Position=" & sb.Position

    '5. 省略時は末尾まで / 残りより多く指定しても残りだけ
    sb.Position = 4090
    Dim r3() As Byte: r3 = sb.ReadBytes()
    sb.Position = 4090
    Dim r4() As Byte: r4 = sb.ReadBytes(100)
    LogResult "T03-5 末尾までの読み出し", (ByteLen(r3) = 6) And (ByteLen(r4) = 6) And (sb.Position = 4096), "省略=" & ByteLen(r3) & " 100指定=" & ByteLen(r4) & " Position=" & sb.Position

    '6. 末尾で読むと未初期化
    Dim r5() As Byte: r5 = sb.ReadBytes()
    LogResult "T03-6 末尾での読み出し", (ByteLen(r5) = 0), "len=" & ByteLen(r5)
End Sub

'* T04: additionalData省略
Private Sub Test04_NoAdditionalData(wv2 As CDPCoreViaWebView2, ctx As CDPContext)
    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(4096))
    If sb Is Nothing Then LogResult "T04 Create", False, "Nothing が返った": Exit Sub

    ResetJsResult ctx
    If Not sb.PostToScript(SharedBuffer_ReadOnly) Then LogResult "T04 Post", False, "PostToScript失敗": Exit Sub

    Dim res As String: res = WaitJsResult(ctx, 10)
    'Post自体は成功し、JS側では空(実測ではundefined)になる。nullでも可とする
    LogResult "T04 additionalData省略", (res = "additionalData=undefined,len=4096") Or (res = "additionalData=null,len=4096"), "JS=[" & res & "]"
End Sub

'* T05: VBA → JS (ReadOnly)
Private Sub Test05_VbaToJs(wv2 As CDPCoreViaWebView2, ctx As CDPContext)
    Dim msg As String
    msg = "こんにちは、SharedBuffer！VBA→JS " & ChrW$(&HD83D) & ChrW$(&HDE43)

    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(4096))
    If sb Is Nothing Then LogResult "T05 Create", False, "Nothing が返った": Exit Sub

    Dim data() As Byte: data = m_conv.BytesFromString(msg)
    If Not sb.WriteBytes(data) Then LogResult "T05 Write", False, "WriteBytes失敗": Exit Sub

    ResetJsResult ctx
    If Not sb.PostToScript(SharedBuffer_ReadOnly, "{""mode"":""echo"",""length"":" & sb.Position & "}") Then LogResult "T05 Post", False, "PostToScript失敗": Exit Sub

    Dim res As String: res = WaitJsResult(ctx, 10)
    LogResult "T05 VBA->JS ReadOnly", (res = msg), "utf8Bytes=" & ByteLen(data) & " echo=[" & res & "]"
End Sub

'* T06: JS → VBA (ReadWrite)
Private Sub Test06_JsToVba(wv2 As CDPCoreViaWebView2, ctx As CDPContext)
    Dim expected As String
    expected = "WebView2→VBA：日本語テキストをTextEncoderで書き込み"

    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(4096))
    If sb Is Nothing Then LogResult "T06 Create", False, "Nothing が返った": Exit Sub

    ResetJsResult ctx
    If Not sb.PostToScript(SharedBuffer_ReadWrite, "{""mode"":""text"",""text"":""" & expected & """}") Then LogResult "T06 Post", False, "PostToScript失敗": Exit Sub

    Dim res As String: res = WaitJsResult(ctx, 10)
    Dim written As Long: written = CLng(Val(res))

    'JSが書いてもIStreamの位置は動かない(=0のまま)
    Dim posAfterJs As LongLong: posAfterJs = sb.Position

    Dim got() As Byte: got = sb.ReadBytes(written)
    Dim s As String: s = m_conv.BytesToString(got)
    LogResult "T06 JS->VBA ReadWrite", (s = expected) And (posAfterJs = 0), "written=" & written & " JS書込後Position=" & posAfterJs & " text=[" & s & "]"
End Sub

'* T07: 同じバッファで往復
Private Sub Test07_PingPong(wv2 As CDPCoreViaWebView2, ctx As CDPContext)
    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(4096))
    If sb Is Nothing Then LogResult "T07 Create", False, "Nothing が返った": Exit Sub

    '1. VBAが先頭に"ping(VBA)"を書く → Positionはその直後
    Dim ping() As Byte: ping = m_conv.BytesFromString("ping(VBA)")
    If Not sb.WriteBytes(ping) Then LogResult "T07 Write", False, "WriteBytes失敗": Exit Sub

    '2. JSは読んで、その続きに"pong(JS)"を書く
    ResetJsResult ctx
    If Not sb.PostToScript(SharedBuffer_ReadWrite, "{""mode"":""pingpong"",""length"":" & sb.Position & "}") Then LogResult "T07 Post", False, "PostToScript失敗": Exit Sub

    Dim res As String: res = WaitJsResult(ctx, 10)
    Dim parts() As String: parts = Split(res, "|")
    If UBound(parts) < 1 Then LogResult "T07 PingPong", False, "JS=[" & res & "]": Exit Sub

    '3. VBAはPositionを動かさずに、そのまま続きを読む
    Dim pong() As Byte: pong = sb.ReadBytes(CLng(Val(parts(1))))
    Dim s As String: s = m_conv.BytesToString(pong)
    LogResult "T07 PingPong", (parts(0) = "ping(VBA)") And (s = "pong(JS)"), "JSが読んだ=[" & parts(0) & "] VBAが読んだ=[" & s & "]"
End Sub

'* T08: 大容量
Private Sub Test08_Large(wv2 As CDPCoreViaWebView2, ctx As CDPContext, ByVal outDir As String)
    Dim t0 As LongLong

    '1. 確保 → JSが全域に書き込み
    t0 = GetTickCount64()
    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(LARGE_SIZE))
    If sb Is Nothing Then LogResult "T08 Create", False, "Nothing が返った": Exit Sub

    ResetJsResult ctx
    If Not sb.PostToScript(SharedBuffer_ReadWrite, "{""mode"":""fill""}") Then LogResult "T08 Post", False, "PostToScript失敗": Exit Sub

    Dim res As String: res = WaitJsResult(ctx, 60)
    LogResult "T08-1 JS書き込み", (Val(res) = LARGE_SIZE), "JS=[" & res & "] 確保+JS書込=" & (GetTickCount64() - t0) & "ms"

    '2. ReadBytes省略で一括 → 全バイト検証(JSは u8[i] = (i*31+7) & 255)
    t0 = GetTickCount64()
    sb.Position = 0
    Dim all() As Byte: all = sb.ReadBytes()
    Dim tRead As LongLong: tRead = GetTickCount64() - t0

    t0 = GetTickCount64()
    Dim i As Long, v As Long, bad As Long: bad = -1
    v = 7
    For i = 0 To ByteLen(all) - 1
        If all(i) <> v Then bad = i: Exit For
        v = (v + 31) And 255
    Next i
    Dim tVerify As LongLong: tVerify = GetTickCount64() - t0
    LogResult "T08-2 ReadBytes一括", (ByteLen(all) = LARGE_SIZE) And (bad = -1), "len=" & ByteLen(all) & " 不一致位置=" & bad & " Read=" & tRead & "ms 全数検証=" & tVerify & "ms"
    Erase all

    '3. 16MBずつ分割で読んでファイルへ(Put)
    Dim path1 As String: path1 = outDir & "\large_chunked.bin"
    If Len(Dir$(path1)) > 0 Then Kill path1

    t0 = GetTickCount64()
    sb.Position = 0
    Dim fn As Integer: fn = FreeFile
    Open path1 For Binary Access Write As #fn
    Dim chunk() As Byte, chunkCount As Long
    Do
        chunk = sb.ReadBytes(CHUNK_SIZE)
        If ByteLen(chunk) = 0 Then Exit Do
        Put #fn, , chunk
        chunkCount = chunkCount + 1
    Loop
    Close #fn
    LogResult "T08-3 ReadBytes分割→Put", (FileLen(path1) = LARGE_SIZE), "chunks=" & chunkCount & " file=" & FileLen(path1) & " " & (GetTickCount64() - t0) & "ms"

    '4. get_Bufferのポインターから直接WriteFile(VBAの配列を経由しない)
    Dim path2 As String: path2 = outDir & "\large_pointer.bin"
    t0 = GetTickCount64()
    Dim wf As Boolean: wf = WritePointerToFile(path2, sb.get_Buffer, LARGE_SIZE)
    LogResult "T08-4 get_Buffer→WriteFile", wf And (FileLen(path2) = LARGE_SIZE), "file=" & FileLen(path2) & " " & (GetTickCount64() - t0) & "ms"

    '5. 両ファイルの数点が期待値と一致するか
    Dim offsets As Variant: offsets = Array(0&, 1&, 255&, 4096&, 12345678, LARGE_SIZE - 1)
    Dim ok As Boolean: ok = True
    Dim k As Long, b1 As Byte, b2 As Byte
    Dim fn1 As Integer: fn1 = FreeFile
    Open path1 For Binary Access Read As #fn1
    Dim fn2 As Integer: fn2 = FreeFile
    Open path2 For Binary Access Read As #fn2
    For k = LBound(offsets) To UBound(offsets)
        Get #fn1, CLng(offsets(k)) + 1, b1
        Get #fn2, CLng(offsets(k)) + 1, b2
        If b1 <> Expected8(CLng(offsets(k))) Or b2 <> b1 Then ok = False
    Next k
    Close #fn1
    Close #fn2
    LogResult "T08-5 ファイル内容(抜き取り)", ok, "offsets=0,1,255,4096,12345678,末尾"
End Sub

Private Function Expected8(ByVal i As Long) As Byte
    Expected8 = CByte((CLngLng(i) * 31 + 7) And 255)
End Function

'* T09: 確保と解放の繰り返し
Private Sub Test09_Lifecycle(wv2 As CDPCoreViaWebView2)
    Dim i As Long, okCount As Long
    Dim sb As WebView2SharedBuffer
    Dim one(0 To 0) As Byte: one(0) = 123

    For i = 1 To 20
        Set sb = wv2.CreateSharedBuffer(CLngLng(10 * 1024& * 1024&))
        If Not sb Is Nothing Then
            sb.Position = sb.get_Size - 1
            If sb.WriteBytes(one) Then okCount = okCount + 1
        End If
        Set sb = Nothing    'Class_Terminate → CLOSE_
    Next i

    LogResult "T09 確保/解放x20(10MB)", (okCount = 20), "ok=" & okCount & "/20"
End Sub

'* T10: サイズ上限
Private Sub Test10_SizeLimit(wv2 As CDPCoreViaWebView2)
    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(2 ^ 31))
    'IIfは両方の式を評価するため、Nothingのときにget_Sizeを読まないようIfで分ける
    If sb Is Nothing Then
        LogResult "T10 2GB(2^31)は確保不可", True, "Nothing"
    Else
        LogResult "T10 2GB(2^31)は確保不可", False, "確保できてしまった Size=" & sb.get_Size
    End If
End Sub



'***************************************************************************************************
'                             ■■■ JS側の受け口(リスナー) ■■■
'***************************************************************************************************
'* 機能　　：'sharedbufferreceived'のリスナーを登録するJSを返します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`additionalData.mode`で処理を切り替え、終わったら`releaseBuffer`してから
'            結果文字列を`window.__sbRes`へ置きます(VBAは`WaitJsResult`で拾う)
'***************************************************************************************************
Private Function ListenerScript() As String
    Dim js As String
    js = "(() => {"
    js = js & "if (!window.chrome || !chrome.webview) return 'no chrome.webview';"
    js = js & "if (window.__sbListener) return 'already';"
    js = js & "window.__sbListener = true;"
    js = js & "chrome.webview.addEventListener('sharedbufferreceived', e => {"
    js = js & "  const info = e.additionalData;"
    js = js & "  const buf = e.getBuffer();"
    js = js & "  const u8 = new Uint8Array(buf);"
    js = js & "  let res = '';"
    js = js & "  try {"
    js = js & "    switch (info ? info.mode : 'nodata') {"
    js = js & "      case 'nodata': { res = 'additionalData=' + (info === null ? 'null' : typeof info) + ',len=' + u8.length; break; }"
    js = js & "      case 'echo': { res = new TextDecoder().decode(u8.subarray(0, info.length)); break; }"
    js = js & "      case 'text': { res = String(new TextEncoder().encodeInto(info.text, u8).written); break; }"
    js = js & "      case 'pingpong': {"
    js = js & "        const ping = new TextDecoder().decode(u8.subarray(0, info.length));"
    js = js & "        const w = new TextEncoder().encodeInto('pong(JS)', u8.subarray(info.length)).written;"
    js = js & "        res = ping + '|' + w; break; }"
    js = js & "      case 'fill': { for (let i = 0; i < u8.length; i++) u8[i] = (i * 31 + 7) & 255; res = String(u8.length); break; }"
    js = js & "      default: res = 'unknown mode';"
    js = js & "    }"
    js = js & "  } catch (err) { res = 'error:' + err; }"
    js = js & "  chrome.webview.releaseBuffer(buf);"
    js = js & "  window.__sbRes = res;"
    js = js & "});"
    js = js & "return 'installed'; })()"
    ListenerScript = js
End Function

'* Post前に、前回の結果を消しておく
Private Sub ResetJsResult(ctx As CDPContext)
    ctx.jsEval "window.__sbRes = undefined; 0"
End Sub

'* JSのリスナーが結果を置くまで待つ(タイムアウト時は空文字)
Private Function WaitJsResult(ctx As CDPContext, ByVal timeoutSec As Double) As String
    Dim t0 As LongLong: t0 = GetTickCount64()
    Dim v As Variant
    Do
        v = ctx.jsEval("window.__sbRes === undefined ? '' : String(window.__sbRes)")
        If Not IsError(v) Then
            If Len(v) > 0 Then WaitJsResult = v: Exit Function
        End If
        DoEvents
    Loop Until GetTickCount64() - t0 > timeoutSec * 1000
End Function



'***************************************************************************************************
'                                   ■■■ 補助 ■■■
'***************************************************************************************************
'* バイト配列の要素数(未初期化なら0)
Private Function ByteLen(Data() As Byte) As Long
    On Error GoTo NotInitialized
    ByteLen = UBound(Data) - LBound(Data) + 1
    Exit Function
NotInitialized:
    ByteLen = 0
End Function

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



'***************************************************************************************************
'                                   ■■■ ログ ■■■
'***************************************************************************************************
Private Sub LogResult(ByVal testName As String, ByVal passed As Boolean, ByVal detail As String)
    If passed Then m_passCount = m_passCount + 1 Else m_failCount = m_failCount + 1
    LogLine IIf(passed, "[PASS] ", "[FAIL] ") & testName & " : " & detail
End Sub

Private Sub LogLine(ByVal s As String)
    Debug.Print s
    m_log = m_log & s & vbCrLf
End Sub

'* UTF-8で保存(絵文字等も確認できるように)
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
