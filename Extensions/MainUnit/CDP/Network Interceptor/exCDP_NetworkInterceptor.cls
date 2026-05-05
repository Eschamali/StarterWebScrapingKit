VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_NetworkInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =========================================================
' exCDP_NetworkInterceptor
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
'   Dim ni As New exCDP_NetworkInterceptor
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

Private Const ThisClassName         As String = "exCDP_NetworkInterceptor"

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

    Dim reqId As String
    Dim k As Long

    Select Case methodName

        ' CDPネイティブのレスポンス受信イベント
        ' Network.enable を呼んで初めて届く
        Case "Network.responseReceived"
            ' ヘッダー到着。Body はまだなので BodyReady=False
            reqId = ExtractVal_(EventRawJson, """requestId"":""", """")
            Dim d As New Scripting.Dictionary
            d.Add "RequestId", reqId
            d.Add "Url", ExtractVal_(EventRawJson, """url"":""", """")
            d.Add "Status", ExtractVal_(EventRawJson, """status"":", ",")
            d.Add "MimeType", ExtractVal_(EventRawJson, """mimeType"":""", """")
            d.Add "BodyReady", False
            capturedResponses_.Add d
            printMsg Debug_, "<NetworkInterceptor> Header received: [" & d("Status") & "] " & d("Url"), FromProcedureName

        Case "Network.loadingFinished"
            ' Body DL完了 -> 対応エントリの BodyReady を True に更新
            reqId = ExtractVal_(EventRawJson, """requestId"":""", """")
            For k = 1 To capturedResponses_.Count
                If capturedResponses_(k)("RequestId") = reqId Then
                    capturedResponses_(k)("BodyReady") = True
                    printMsg Debug_, "<NetworkInterceptor> Body ready: " & capturedResponses_(k)("Url"), FromProcedureName
                    Exit For
                End If
            Next k

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
    printMsg Debug_, "<NetworkInterceptor> Network.enable - capture started.", FromProcedureName
End Sub

' Network.disable を呼んでキャプチャを終了する
Public Sub StopNetworkCapture()
    Const FromProcedureName As String = ThisClassName & ".StopNetworkCapture"


    invokeMethod "Network.disable"
    networkCaptureActive_ = False
    printMsg Debug_, "<NetworkInterceptor> Network.disable - capture stopped.", FromProcedureName
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


    printMsg Debug_, "<WaitForResponse> Waiting for '" & urlContains & "' (BodyReady). Timeout=" & maxWaitTimeSec & "s", FromProcedureName

    Dim startTime As Double: startTime = Timer

    Do
        ' TakeEvents を呼ぶことでPipeを読み込み、RaiseEvent を発火させる
        ' この呼び出しがないと detectionCDPEvent が一切トリガーされない
        ex_CDPBrowser.TakeEvents

        ' 蓄積されたレスポンスを後ろから検索（新しいものを優先）
        Dim j As Long
        For j = capturedResponses_.Count To 1 Step -1
            ' BodyReady=True のものだけが対象
            If capturedResponses_(j)("BodyReady") = True Then
                If InStr(1, CStr(capturedResponses_(j)("Url")), urlContains, vbTextCompare) > 0 Then
                    printMsg Debug_, "<WaitForResponse> Found: " & capturedResponses_(j)("Url"), FromProcedureName
                    WaitForResponse = CStr(capturedResponses_(j)("RequestId"))
                    Exit Function
                End If
            End If
        Next j

        If (Timer - startTime) > maxWaitTimeSec Then
            printMsg WARN_, "<WaitForResponse> Timeout: '" & urlContains & "' Body not ready within " & maxWaitTimeSec & "s.", FromProcedureName
            WaitForResponse = ""
            Exit Function
        End If

        DoEvents
    Loop
End Function

'-----------------------------------------------------------
' requestId に対応するレスポンスBodyを取得する
' WaitForResponse の戻り値を渡してください
'-----------------------------------------------------------
Public Function GetResponseBody(requestId As String) As String
    Const FromProcedureName As String = ThisClassName & ".GetResponseBody"


    If requestId = "" Then
        printMsg WARN_, "<GetResponseBody> requestId is empty.", FromProcedureName
        GetResponseBody = ""
        Exit Function
    End If

    Dim params As New Scripting.Dictionary
    params("requestId") = requestId

    Dim res As Scripting.Dictionary
    Set res = invokeMethod("Network.getResponseBody", params, , False)

    If res Is Nothing Then
        printMsg WARN_, "<GetResponseBody> Failed. (requestId: " & requestId & ")", FromProcedureName
        GetResponseBody = ""
        Exit Function
    End If

    If res.Exists("body") Then
        GetResponseBody = CStr(res("body"))
        Exit Function
    End If

    printMsg WARN_, "<GetResponseBody> Body not found in result.", FromProcedureName
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
    js = js & "var p=" & JsStr_(urlPattern) & ";"
    js = js & "if(window.__vbaBlockedURLPatterns.indexOf(p)<0){window.__vbaBlockedURLPatterns.push(p);}"
    js = js & "return true;})();"
    ex_CDPBrowser.jsEval js
    printMsg Debug_, "<NetworkInterceptor> Block added: " & urlPattern, FromProcedureName
End Sub

' すべてのブロックURLパターンをクリアする
Public Sub ClearBlockedURLs()
    Const FromProcedureName As String = ThisClassName & ".ClearBlockedURLs"


    ex_CDPBrowser.jsEval "(function(){window.__vbaBlockedURLPatterns=[];return true;})();"
    printMsg Debug_, "<NetworkInterceptor> Blocked URLs cleared.", FromProcedureName
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
    js = js & "urlPattern:" & JsStr_(urlPattern) & ","
    js = js & "statusCode:" & StatusCode & ","
    js = js & "body:" & JsStr_(Body) & ","
    js = js & "contentType:" & JsStr_(ContentType)
    js = js & "});return true;})();"
    ex_CDPBrowser.jsEval js
    printMsg Debug_, "<NetworkInterceptor> Mock added: [" & StatusCode & "] " & urlPattern, FromProcedureName
End Sub

' すべてのモックレスポンスをクリアする
Public Sub ClearMockResponses()
    Const FromProcedureName As String = ThisClassName & ".ClearMockResponses"


    ex_CDPBrowser.jsEval "(function(){window.__vbaMockResponses=[];return true;})();"
    printMsg Debug_, "<NetworkInterceptor> Mock responses cleared.", FromProcedureName
End Sub

' fetch / XHR フック両方を完全に取り除く（ページリフレッシュ前後でも呼べる）
Public Sub DisableInterceptor()
    Const FromProcedureName As String = ThisClassName & ".DisableInterceptor"


    Dim js As String
    js = "(function(){try{"
    js = js & "if(window.__vbaOriginalFetch){window.fetch=window.__vbaOriginalFetch;window.__vbaOriginalFetch=null;}"
    js = js & "if(window.__vbaOriginalXHROpen){XMLHttpRequest.prototype.open=window.__vbaOriginalXHROpen;window.__vbaOriginalXHROpen=null;}"
    js = js & "if(window.__vbaOriginalXHRSend){XMLHttpRequest.prototype.send=window.__vbaOriginalXHRSend;window.__vbaOriginalXHRSend=null;}"
    js = js & "window.__vbaBlockedURLPatterns=[];window.__vbaMockResponses=[];"
    js = js & "window.__vbaNetworkInterceptionInstalled=false;"
    js = js & "return true;}catch(e){return false;}})();"
    ex_CDPBrowser.jsEval js
    printMsg Debug_, "<NetworkInterceptor> Interceptor disabled.", FromProcedureName
End Sub

' =========================================================
' Private: JS フックインストール
' =========================================================

' fetch/XHR のフックコアJSをページに一度だけインストールする
Private Sub installInterceptorIfNeeded_()
    Const FromProcedureName As String = ThisClassName & ".installInterceptorIfNeeded_"


    Dim js As String
    js = "(function(){"
    js = js & "if(window.__vbaNetworkInterceptionInstalled){return true;}"
    js = js & "window.__vbaNetworkInterceptionInstalled=true;"
    js = js & "window.__vbaBlockedURLPatterns=[];"
    js = js & "window.__vbaMockResponses=[];"

    ' ── URL マッチャー ──
    js = js & "window.__vbaMatchUrl=function(u,p){try{return String(u).indexOf(String(p))>=0;}catch(e){return false;}};"
    js = js & "window.__vbaGetMock=function(u){for(var i=0;i<window.__vbaMockResponses.length;i++){var m=window.__vbaMockResponses[i];if(window.__vbaMatchUrl(u,m.urlPattern))return m;}return null;};"
    js = js & "window.__vbaIsBlocked=function(u){for(var i=0;i<window.__vbaBlockedURLPatterns.length;i++){if(window.__vbaMatchUrl(u,window.__vbaBlockedURLPatterns[i]))return true;}return false;};"
    ' fetch フック
    js = js & "if(window.fetch&&!window.__vbaOriginalFetch){"
    js = js & "window.__vbaOriginalFetch=window.fetch.bind(window);"
    js = js & "window.fetch=function(input,init){var url=(typeof input==='string')?input:((input&&input.url)?input.url:String(input));"
    js = js & "if(window.__vbaIsBlocked(url))return Promise.reject(new Error('Blocked by VBA'));"
    js = js & "var mock=window.__vbaGetMock(url);"
    js = js & "if(mock)return Promise.resolve(new Response(mock.body,{status:mock.statusCode,headers:{'Content-Type':mock.contentType}}));"
    js = js & "return window.__vbaOriginalFetch(input,init);};}"
    ' XHR フック
    js = js & "if(window.XMLHttpRequest&&!window.__vbaOriginalXHROpen){"
    js = js & "window.__vbaOriginalXHROpen=XMLHttpRequest.prototype.open;"
    js = js & "window.__vbaOriginalXHRSend=XMLHttpRequest.prototype.send;"
    js = js & "XMLHttpRequest.prototype.open=function(m,u){this.__vbaUrl=u;return window.__vbaOriginalXHROpen.apply(this,arguments);};"
    js = js & "XMLHttpRequest.prototype.send=function(){var u=this.__vbaUrl||'';"
    js = js & "if(window.__vbaIsBlocked(u))throw new Error('Blocked by VBA');"
    js = js & "var mock=window.__vbaGetMock(u);"
    js = js & "if(mock){var x=this;"
    js = js & "Object.defineProperty(x,'readyState',{value:4,configurable:true});"
    js = js & "Object.defineProperty(x,'status',{value:mock.statusCode,configurable:true});"
    js = js & "Object.defineProperty(x,'responseText',{value:String(mock.body),configurable:true});"
    js = js & "Object.defineProperty(x,'response',{value:String(mock.body),configurable:true});"
    js = js & "setTimeout(function(){if(x.onreadystatechange)x.onreadystatechange();if(x.onload)x.onload();},0);return;}"
    js = js & "return window.__vbaOriginalXHRSend.apply(this,arguments);};}"
    js = js & "return true;})();"
    ex_CDPBrowser.jsEval js
    printMsg Debug_, "<NetworkInterceptor> fetch/XHR hook installed.", FromProcedureName
End Sub

' =========================================================
' Private: ユーティリティ
' =========================================================
' 文字列を JS の文字列リテラルとして安全にエスケープする
' ※単純な置換のみ。バイナリや特殊制御文字が含まれる場合は要拡張
Private Function JsStr_(s As String) As String
    Dim r As String
    r = Replace(s, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, Chr(13), "\r")
    r = Replace(r, Chr(10), "\n")
    r = Replace(r, Chr(9), "\t")
    JsStr_ = """" & r & """"
End Function

' rawJson から keyPrefix と keySuffix に挟まれた値を簡易抽出する
Private Function ExtractVal_(src As String, Prefix As String, suffix As String) As String
    Dim p As Long: p = InStr(src, Prefix)
    If p = 0 Then Exit Function
    Dim s As Long: s = p + Len(Prefix)
    Dim e As Long: e = InStr(s, src, suffix)
    If e > s Then ExtractVal_ = Mid(src, s, e - s)
End Function

' CDPBrowser直通ラッパー
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set invokeMethod = ex_CDPBrowser.invokeMethod(methodName, params, alwaysBrowserContext, StopError)
End Function
