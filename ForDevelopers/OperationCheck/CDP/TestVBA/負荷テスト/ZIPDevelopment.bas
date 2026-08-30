Attribute VB_Name = "ZIPDevelopment"
Option Explicit

Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" _
    Alias "SHCreateDirectoryExA" _
    (ByVal hWnd As LongPtr, _
     ByVal pszPath As String, _
     ByVal psa As LongPtr) As Long

Private Const WebSocketTest As Boolean = True



Sub Webブラウザ操作でZIPテスト()
    '設定シートに基づくブラウザ立ち上げ
    Dim ZIPテスト As CDPContext
    If WebSocketTest Then
        '---- WebSocket版 ----
        '設定セルから、ユーザ名を取得
        Dim UserName As String
        UserName = ShSetting01_StartBrowser.CurrentUserName

        '指定のWebSocketForCDPへ接続
        Dim WebSocketCDP As New CDPCoreViaWebSocket
        Debug.Print WebSocketCDP.AutoConnectBrowserCDP(UserName)

        '繋げたWebSocketオブジェクトを`reattach`メソッドに渡す
        Dim chrome As New CDPBrowser
        If Not chrome.reattach(UserName, WebSocketCDP) Then MsgBox "「" & UserName & "」に接続できませんでした。WebSocket情報がお亡くなりです。", vbCritical, "Chrome DevTools Protocol": Exit Sub
        Set ZIPテスト = chrome.newTab(setMain:=True)
        '---------------------
    Else
        '---- Pipe版 ----
        Set ZIPテスト = ShSetting01_StartBrowser.StartCDPModeContext
        '----------------
    End If

    ' 1. zip.js (UMD版) を動的にロードするJSを実行
    Dim injectCode As String
    injectCode = "var script = document.createElement('script');" & _
                 "script.src = 'https://cdn.jsdelivr.net/npm/@zip.js/zip.js@2.7.34/dist/zip-no-worker.min.js';" & _
                 "document.head.appendChild(script);"

    
    ZIPテスト.jsEval injectCode

    ' 2. ライブラリの読み込み完了を待機
    Dim isLoaded
    Do
        isLoaded = ZIPテスト.jsEval("typeof zip !== 'undefined'")
        If Not IsError(isLoaded) Then Exit Do
        Application.wait (Now + TimeValue("0:00:01"))
    Loop

    ' 3. ローカルのZIPをBase64化する
    Dim b64ZipData As String, ZipDataBin() As Byte
    Dim CharConv As New CharacterCodeConversion
    ZipDataBin = CharConv.BytesFromSavedFile(Environ("UserProfile") & "\Downloads", "twinBASIC_IDE_BETA_983.zip")
    b64ZipData = WebCrypto.Encode(ZipDataBin, edfBase64, efNoFolding)

    ' 4. 即時実行関数 (IIFE) のJSコードを組み立て
    Dim JsCode As String
    JsCode = _
        "(async () => {" & _
        "  const zipBytes = Uint8Array.from(atob('" & b64ZipData & "'), c => c.charCodeAt(0));" & _
        "  const reader = new zip.ZipReader(new zip.Uint8ArrayReader(zipBytes));" & _
        "  const entries = await reader.getEntries();" & _
        "  const results = [];" & _
        "  for (const entry of entries) {" & _
        "    if (!entry.directory) {" & _
        "      const fileBytes = await entry.getData(new zip.Uint8ArrayWriter());" & _
        "      let binary = '';" & _
        "      const len = fileBytes.byteLength;" & _
        "      for (let i = 0; i < len; i++) {" & _
        "        binary += String.fromCharCode(fileBytes[i]);" & _
        "      }" & _
        "      results.push({" & _
        "        filename: entry.filename," & _
        "        base64: btoa(binary)" & _
        "      });" & _
        "    }" & _
        "  }" & _
        "  await reader.close();" & _
        "  return results;" & _
        "})()"

    ' 5. CDP経由で実行し、解凍された全データを一発で受け取る！
    ' ※ awaitPromise:=True でPromiseの解決を待ち、
    '   returnByValue:=True で中身のデータを直接取得します。
    Dim resCDP As BiDiCDPJson
    Set resCDP = ZIPテスト.jsEval(JsCode, awaitPromise:=True, returnByValue:=True)
    ZIPテスト.ThisCDPBrowser.quit

    '6．展開
    Dim ベース展開先 As String
    ベース展開先 = Environ("UserProfile") & "\Downloads\JavaScriptからの結果"
    
    Dim i As Long
    Dim NodeToken As Long
    With resCDP
        NodeToken = .FirstChildToken
        Debug.Print .Count; "個のファイルを展開します..."
        Do While NodeToken > 0
            '保存先用意
            Dim 相対パス As String, ファイル名 As String, 保存先フルフォルダパス As String
            相対パス = "\" & .TokenString(NodeToken, "filename")                  'JavaScript結果から、相対パスを取って...
            ファイル名 = ファイルとフォルダパス分離(相対パス)   'うまいぐあいに、ファイル名とフォルダパスを分離させて...
            保存先フルフォルダパス = ベース展開先 & 相対パス    '保存先フォルダの絶対パスを作って...
            If Len(Dir(保存先フルフォルダパス, vbDirectory)) = 0 Then
                Dim ResultCode As Long
                ResultCode = SHCreateDirectoryEx(0&, 保存先フルフォルダパス, 0&)    'その保存先がなければ作る
    
                '失敗時は終了
                If ResultCode Then
                    MsgBox 保存先フルフォルダパス & vbCrLf & "上記のフォルダ作成にて、エラーが発生しました。" & vbCrLf & "> " & WinApiError.GetMessage(ResultCode), vbCritical, "ErrorCode: " & ResultCode
                    Exit Sub
                End If
            End If
    
            '実際に保存
            Dim 展開後B64 As String, 展開後Bin() As Byte
            展開後B64 = .TokenString(NodeToken, "base64")
            
            '空文字=0バイト判定
            If LenB(展開後B64) > 0 Then 展開後Bin = WebCrypto.Decode(展開後B64, edfBase64) Else 展開後Bin = vbNullString
            CharConv.BytesToSaveFile 展開後Bin, 保存先フルフォルダパス, ファイル名
    
            DoEvents
            i = i + 1
            Debug.Print i; "つめのファイルを展開完了: "; ファイル名
            NodeToken = .NextToken(NodeToken)
        Loop
    End With

End Sub

Function ファイルとフォルダパス分離(ByRef Path As String) As String
    '1. 最も右にある「\」を取得
    Dim EndPos As Long
    EndPos = InStrRev(Path, "/")

    '2. 存在有無に応じた判定
    If EndPos = 0 Then
        'ない場合は、引数をファイル名として返す
        ファイルとフォルダパス分離 = Path

        'クリア
        Path = vbNullString
    Else
        ファイルとフォルダパス分離 = Mid(Path, EndPos + 1, Len(Path) - EndPos)
        Path = Replace(Mid(Path, 1, EndPos - 1), "/", "\")
    End If
End Function
