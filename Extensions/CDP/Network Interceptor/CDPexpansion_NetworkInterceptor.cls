VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPexpansion_NetworkInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =========================================================
' CDPexpansion_NetworkInterceptor
'
' [概要]
' ネットワーク通信の傍受・モック・レスポンス待機を提供する拡張機能クラス。
'
' [できること]
'   A) URL ブロック       ─ 特定URLパターンへの fetch/XHR をブロックする
'   B) モックレスポンス   ─ 特定URLへの通信を偽レスポンスで差し替える
'   C) レスポンス待機     ─ 指定URLを含むレスポンスが届くまで待機する (CDPネイティブ)
'   D) レスポンスBody取得 ─ requestId をキーにレスポンスBody文字列を取得する
'
' [仕組みの違い]
'   A/B : JS で fetch / XHR を乗っ取るハック手法。ページスコープで完結。
'         ※ <img>/<link>/<script> タグが直に起こすリソース読み込みには効かない。
'   C/D : CDP ネイティブの Network.responseReceived イベントを使う正攻法。
'         全ての通信（タグ起因含む）をキャプチャできる。
'
' [使い方]
'   Dim ni As New CDPexpansion_NetworkInterceptor
'   ni.Init br
'
'   ' --- A. URLブロック ---
'   ni.AddBlockedURL "ads.example.com"
'   ' ページを操作... → ads.example.com への fetch/XHR が失敗扱いになる
'   ni.ClearBlockedURLs
'
'   ' --- B. モックレスポンス ---
'   ni.AddMockResponse "/api/user", 200, "{""name"":""テスト太郎""}", "application/json"
'   ' ページを操作... → /api/user への fetch/XHR が即座に偽レスポンスを返す
'   ni.ClearMockResponses
'
'   ' --- C. レスポンス待機 ---
'   ni.StartNetworkCapture          ' ← Network.enable (CDPドメイン有効化)
'   br.jsEval "fetch('/api/data');"
'   Dim reqId As String
'   reqId = ni.WaitForResponse("/api/data", 10)    ' requestId が返る
'   If reqId <> "" Then
'       Dim body As String
'       body = ni.GetResponseBody(reqId)
'       MsgBox body
'   End If
'   ni.StopNetworkCapture           ' ← Network.disable
' =========================================================
Option Explicit


' =========================================================
' 共通定義
' =========================================================
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

Private Const ThisClassName         As String = "CDPexpansion_NetworkInterceptor"

' Network.responseReceived で蓄積するレスポンス情報
' 各要素は Scripting.Dictionary {RequestId, Url, Status, MimeType}
Private capturedResponses_          As Collection

' イベント監視フラグ
Private networkCaptureActive_       As Boolean


' =========================================================
' init
' =========================================================
Friend Sub Init(Inheritance As CDPBrowser)
    Set ex_CDPBrowser = Inheritance
    Set capturedResponses_ = New Collection
End Sub


' =========================================================
'  CDPブラウザイベントハンドラ（CDPネイティブイベントを受け取る場所）
'  ─ Network.enable 後に Network.responseReceived が届く
' =========================================================
Private Sub ex_CDPBrowser_detectionCDPEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPBrowser_detectionCDPEvent"


    ' networkCaptureActive_ が False の間はすべて無視
    If Not networkCaptureActive_ Then Exit Sub

    Select Case methodName

        ' CDPネイティブのレスポンス受信イベント
        ' Network.enable を呼んで初めて届く
        Case "Network.responseReceived"
            Dim d As New Scripting.Dictionary
            d("RequestId") = ExtractJsonStr(EventRawJson, """requestId"":""", """")
            d("Url") = ExtractJsonStr(EventRawJson, """url"":""", """")
            d("Status") = ExtractJsonStr(EventRawJson, """status"":", ",")
            d("MimeType") = ExtractJsonStr(EventRawJson, """mimeType"":""", """")
            capturedResponses_.Add d
            printMsg Debug_, "<NetworkInterceptor> レスポンス受信: [" & d("Status") & "] " & d("Url"), FromProcedureName

        Case Else
            ' 上記以外は無視
    End Select

End Sub


' =========================================================
' [C] レスポンス待機 ─ CDPネイティブ Network ドメイン
' =========================================================

' Network.enable を呼んで CDPイベントのキャプチャを開始する
' WaitForResponse を使う前に必ず呼ぶこと
Public Sub StartNetworkCapture()
    Const FromProcedureName As String = ThisClassName & ".StartNetworkCapture"


    invokeMethod "Network.enable"
    networkCaptureActive_ = True
    Set capturedResponses_ = New Collection
    printMsg Debug_, "<NetworkInterceptor> Network.enable 完了。レスポンスのキャプチャを開始します。", FromProcedureName
End Sub

' Network.disable を呼んでキャプチャを終了する
Public Sub StopNetworkCapture()
    Const FromProcedureName As String = ThisClassName & ".StopNetworkCapture"


    invokeMethod "Network.disable"
    networkCaptureActive_ = False
    printMsg Debug_, "<NetworkInterceptor> Network.disable 完了。キャプチャを停止しました。", FromProcedureName
End Sub

' 蓄積済みのレスポンス一覧をクリアする（待機前にリセットしたいときに使う）
Public Sub ClearCapturedResponses()
    Set capturedResponses_ = New Collection
End Sub

'-----------------------------------------------------------
' 指定URL文字列を含むレスポンスが届くまで待機する
' 成功: requestId 文字列を返す
' 失敗(タイムアウト): "" を返す + WARNログ
'-----------------------------------------------------------
Public Function WaitForResponse(urlContains As String, Optional maxWaitTimeSec As Long = 30) As String
    Const FromProcedureName As String = ThisClassName & ".WaitForResponse"


    printMsg Debug_, "<WaitForResponse> '" & urlContains & "' を含むレスポンスを待機開始。(上限 " & maxWaitTimeSec & "秒)", FromProcedureName

    Dim startTime As Double: startTime = Timer
    Dim i As Long

    Do
        ' Application.Wait + DoEvents で 100ms ずつメッセージキューを消化
        ' (kernel32.Sleep はスレッドをブロックするため、RaiseEvent が届かない)
        For i = 1 To 5
            Application.wait Now + (0.02 / 86400)
            DoEvents
        Next i

        ' 蓄積されたレスポンスを後ろから検索（新しいものを優先）
        Dim j As Long
        For j = capturedResponses_.Count To 1 Step -1
            If InStr(1, CStr(capturedResponses_(j)("Url")), urlContains, vbTextCompare) > 0 Then
                printMsg Debug_, "<WaitForResponse> 一致するレスポンスを検出: " & capturedResponses_(j)("Url"), FromProcedureName
                WaitForResponse = CStr(capturedResponses_(j)("RequestId"))
                Exit Function
            End If
        Next j

        If (Timer - startTime) > maxWaitTimeSec Then
            printMsg WARN_, "<WaitForResponse> " & maxWaitTimeSec & "秒待機しましたが '" & urlContains & "' を含むレスポンスが届きませんでした。", FromProcedureName
            WaitForResponse = ""
            Exit Function
        End If
    Loop
End Function

'-----------------------------------------------------------
' requestId に対応するレスポンスBodyを取得する
' WaitForResponse の戻り値を渡してください
'-----------------------------------------------------------
Public Function GetResponseBody(requestId As String) As String
    Const FromProcedureName As String = ThisClassName & ".GetResponseBody"


    If requestId = "" Then
        printMsg WARN_, "<GetResponseBody> requestId が空です。WaitForResponse に失敗していないか確認してください。", FromProcedureName
        GetResponseBody = ""
        Exit Function
    End If

    Dim params As New Scripting.Dictionary
    params("requestId") = requestId

    Dim res As Scripting.Dictionary
    Set res = invokeMethod("Network.getResponseBody", params, , False)

    If res Is Nothing Then
        printMsg WARN_, "<GetResponseBody> レスポンスBodyの取得に失敗しました。(requestId: " & requestId & ")", FromProcedureName
        GetResponseBody = ""
        Exit Function
    End If

    If res.Exists("result") Then
        Dim result As Scripting.Dictionary
        Set result = res("result")
        If result.Exists("body") Then
            GetResponseBody = CStr(result("body"))
            Exit Function
        End If
    End If

    printMsg WARN_, "<GetResponseBody> レスポンスBodyが見つかりませんでした。", FromProcedureName
    GetResponseBody = ""
End Function


' =========================================================
' [A] URL ブロック ─ JS で fetch/XHR を乗っ取る手法
' =========================================================

'-----------------------------------------------------------
' 指定URLパターンへの fetch/XHR をブロックする
' ※ページに NetworkInterceptor JS が自動インストールされる
' ※<img> / <script> / <link> タグ起因の通信には効かない
'-----------------------------------------------------------
Public Sub AddBlockedURL(urlPattern As String)
    Const FromProcedureName As String = ThisClassName & ".AddBlockedURL"


    Call installInterceptorIfNeeded_

    Dim js As String
    js = "(function(){"
    js = js & "window.__vbaBlockedURLPatterns=window.__vbaBlockedURLPatterns||[];"
    js = js & "var p=" & JsonEscape_(urlPattern) & ";"
    js = js & "if(window.__vbaBlockedURLPatterns.indexOf(p)<0){"
    js = js & "window.__vbaBlockedURLPatterns.push(p);"
    js = js & "}"
    js = js & "return true;"
    js = js & "})();"
    ex_CDPBrowser.jsEval js
    printMsg Debug_, "<NetworkInterceptor> URLブロック追加: " & urlPattern, FromProcedureName
End Sub

' すべてのブロックURLパターンをクリアする
Public Sub ClearBlockedURLs()
    Const FromProcedureName As String = ThisClassName & ".ClearBlockedURLs"


    ex_CDPBrowser.jsEval "(function(){ window.__vbaBlockedURLPatterns=[]; return true; })();"
    printMsg Debug_, "<NetworkInterceptor> URLブロックをすべてクリアしました。", FromProcedureName
End Sub


' =========================================================
' [B] モックレスポンス ─ JS で fetch/XHR を乗っ取る手法
' =========================================================

'-----------------------------------------------------------
' 指定URLパターンへの通信を偽レスポンスで差し替える
'   urlPattern  : URL に含まれる文字列（前方一致ではなく Contains）
'   statusCode  : HTTP ステータスコード (例: 200, 404)
'   body        : レスポンスボディ文字列
'   contentType : Content-Type (例: "application/json", "text/html")
'-----------------------------------------------------------
Public Sub AddMockResponse(urlPattern As String, StatusCode As Long, Body As String, Optional ContentType As String = "application/json")
    Const FromProcedureName As String = ThisClassName & ".AddMockResponse"


    Call installInterceptorIfNeeded_

    Dim js As String
    js = "(function(){"
    js = js & "window.__vbaMockResponses=window.__vbaMockResponses||[];"
    js = js & "window.__vbaMockResponses.push({"
    js = js & "urlPattern:" & JsonEscape_(urlPattern) & ","
    js = js & "statusCode:" & StatusCode & ","
    js = js & "body:" & JsonEscape_(Body) & ","
    js = js & "contentType:" & JsonEscape_(ContentType)
    js = js & "});"
    js = js & "return true;"
    js = js & "})();"
    ex_CDPBrowser.jsEval js
    printMsg Debug_, "<NetworkInterceptor> モックレスポンス追加: [" & StatusCode & "] " & urlPattern, FromProcedureName
End Sub

' すべてのモックレスポンスをクリアする
Public Sub ClearMockResponses()
    Const FromProcedureName As String = ThisClassName & ".ClearMockResponses"


    ex_CDPBrowser.jsEval "(function(){ window.__vbaMockResponses=[]; return true; })();"
    printMsg Debug_, "<NetworkInterceptor> モックレスポンスをすべてクリアしました。", FromProcedureName
End Sub

' fetch / XHR フック両方を完全に取り除く（ページリフレッシュ前後でも呼べる）
Public Sub DisableInterceptor()
    Const FromProcedureName As String = ThisClassName & ".DisableInterceptor"


    Dim js As String
    js = "(function(){"
    js = js & "try{"
    js = js & "if(window.__vbaOriginalFetch){"
    js = js & "window.fetch=window.__vbaOriginalFetch;"
    js = js & "window.__vbaOriginalFetch=null;"
    js = js & "}"
    js = js & "if(window.__vbaOriginalXHROpen){"
    js = js & "XMLHttpRequest.prototype.open=window.__vbaOriginalXHROpen;"
    js = js & "window.__vbaOriginalXHROpen=null;"
    js = js & "}"
    js = js & "if(window.__vbaOriginalXHRSend){"
    js = js & "XMLHttpRequest.prototype.send=window.__vbaOriginalXHRSend;"
    js = js & "window.__vbaOriginalXHRSend=null;"
    js = js & "}"
    js = js & "window.__vbaBlockedURLPatterns=[];"
    js = js & "window.__vbaMockResponses=[];"
    js = js & "window.__vbaNetworkInterceptionInstalled=false;"
    js = js & "return true;"
    js = js & "}catch(e){return false;}"
    js = js & "})();"
    ex_CDPBrowser.jsEval js
    printMsg Debug_, "<NetworkInterceptor> fetch/XHR フックを取り除きました。", FromProcedureName
End Sub


' =========================================================
' プライベートユーティリティ
' =========================================================

' fetch/XHR のフックコアJSをページに一度だけインストールする
Private Sub installInterceptorIfNeeded_()
    Const FromProcedureName As String = ThisClassName & ".installInterceptorIfNeeded_"


    Dim js As String
    js = "(function(){"
    js = js & "if(window.__vbaNetworkInterceptionInstalled){return true;}"
    js = js & "window.__vbaNetworkInterceptionInstalled=true;"
    js = js & "window.__vbaBlockedURLPatterns=window.__vbaBlockedURLPatterns||[];"
    js = js & "window.__vbaMockResponses=window.__vbaMockResponses||[];"

    ' ── URL マッチャー ──
    js = js & "window.__vbaMatchUrl=function(url,pat){try{return String(url).indexOf(String(pat))>=0;}catch(e){return false;}};"
    js = js & "window.__vbaGetMockResponse=function(url){"
    js = js & "var i,m;"
    js = js & "for(i=0;i<window.__vbaMockResponses.length;i++){"
    js = js & "m=window.__vbaMockResponses[i];"
    js = js & "if(window.__vbaMatchUrl(url,m.urlPattern)){return m;}"
    js = js & "}return null;};"
    js = js & "window.__vbaIsBlockedUrl=function(url){"
    js = js & "var i;"
    js = js & "for(i=0;i<window.__vbaBlockedURLPatterns.length;i++){"
    js = js & "if(window.__vbaMatchUrl(url,window.__vbaBlockedURLPatterns[i])){return true;}"
    js = js & "}return false;};"

    ' ── fetch フック ──
    js = js & "if(window.fetch&&!window.__vbaOriginalFetch){"
    js = js & "window.__vbaOriginalFetch=window.fetch.bind(window);"
    js = js & "window.fetch=function(input,init){"
    js = js & "var url=(typeof input==='string')?input:((input&&input.url)?input.url:String(input));"
    js = js & "if(window.__vbaIsBlockedUrl(url)){return Promise.reject(new Error('Blocked by VBA NetworkInterceptor'));}"
    js = js & "var mock=window.__vbaGetMockResponse(url);"
    js = js & "if(mock){return Promise.resolve(new Response(mock.body,{status:mock.statusCode,headers:{'Content-Type':mock.contentType}}));}"
    js = js & "return window.__vbaOriginalFetch(input,init);"
    js = js & "};"
    js = js & "}"

    ' ── XHR フック ──
    js = js & "if(window.XMLHttpRequest&&!window.__vbaOriginalXHROpen){"
    js = js & "window.__vbaOriginalXHROpen=XMLHttpRequest.prototype.open;"
    js = js & "window.__vbaOriginalXHRSend=XMLHttpRequest.prototype.send;"
    js = js & "XMLHttpRequest.prototype.open=function(method,url){this.__vbaUrl=url;return window.__vbaOriginalXHROpen.apply(this,arguments);};"
    js = js & "XMLHttpRequest.prototype.send=function(){"
    js = js & "var url=this.__vbaUrl||'';"
    js = js & "if(window.__vbaIsBlockedUrl(url)){throw new Error('Blocked by VBA NetworkInterceptor');}"
    js = js & "var mock=window.__vbaGetMockResponse(url);"
    js = js & "if(mock){"
    js = js & "var xhr=this;"
    js = js & "Object.defineProperty(xhr,'readyState',{value:4,configurable:true});"
    js = js & "Object.defineProperty(xhr,'status',{value:mock.statusCode,configurable:true});"
    js = js & "Object.defineProperty(xhr,'responseText',{value:String(mock.body),configurable:true});"
    js = js & "Object.defineProperty(xhr,'response',{value:String(mock.body),configurable:true});"
    js = js & "setTimeout(function(){if(xhr.onreadystatechange){xhr.onreadystatechange();}if(xhr.onload){xhr.onload();}},0);"
    js = js & "return;"
    js = js & "}"
    js = js & "return window.__vbaOriginalXHRSend.apply(this,arguments);"
    js = js & "};"
    js = js & "}"

    js = js & "return true;"
    js = js & "})();"

    ex_CDPBrowser.jsEval js
    printMsg Debug_, "<NetworkInterceptor> fetch/XHR フックをインストールしました。", FromProcedureName
End Sub

' 文字列を JS の文字列リテラルとして安全にエスケープする
' ※単純な置換のみ。バイナリや特殊制御文字が含まれる場合は要拡張
Private Function JsonEscape_(s As String) As String
    Dim r As String
    r = Replace(s, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, Chr(13), "\r")
    r = Replace(r, Chr(10), "\n")
    r = Replace(r, Chr(9), "\t")
    JsonEscape_ = """" & r & """"
End Function

' rawJson から keyPrefix と keySuffix に挟まれた値を簡易抽出する
Private Function ExtractJsonStr(rawJson As String, keyPrefix As String, keySuffix As String) As String
    Dim p As Long: p = InStr(rawJson, keyPrefix)
    If p = 0 Then Exit Function
    Dim s As Long: s = p + Len(keyPrefix)
    Dim e As Long: e = InStr(s, rawJson, keySuffix)
    If e > s Then ExtractJsonStr = Mid(rawJson, s, e - s)
End Function

' CDPBrowser直通ラッパー
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set invokeMethod = ex_CDPBrowser.invokeMethod(methodName, params, alwaysBrowserContext, StopError)
End Function
