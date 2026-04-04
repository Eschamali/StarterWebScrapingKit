VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPexpansion_AdvancedWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================
' CDPexpansion_AdvancedWait
'
' [概要]
' ブラウザの非同期動作（DOM変更、非同期通信等）が完了するまで
' 待機するための拡張機能クラス。
' =========================================================

Private WithEvents CDP_   As CDPBrowser
Attribute CDP_.VB_VarHelpID = -1

' 初期化処理（Friend定義：外から直接Newさせない運用を想定）
Friend Sub Init(cb As CDPBrowser)
    Set CDP_ = cb
End Sub


' =========================================================
' DOM Mutation の待機
' DOMに変更（追加、削除、属性変更）が起きるまで待機する
' =========================================================
Public Function WaitForDomMutation(Optional maxWaitTimeSec As Long = 5) As Boolean
    Dim js As String
    Dim res As String
    Dim startCount As Long, currCount As Long
    Dim startTime As Double
    
    ' Inject MutationObserver
    js = ""
    js = js & "(function(){"
    js = js & " if(window.__vbaDomMutationObserverInstalled) return true;"
    js = js & " window.__vbaDomMutationCount=0;"
    js = js & " var obs=new MutationObserver(function(muts){"
    js = js & "  window.__vbaDomMutationCount += muts.length;"
    js = js & " });"
    js = js & " var root=document.documentElement||document;"
    js = js & " obs.observe(root,{subtree:true,childList:true,attributes:true,characterData:true});"
    js = js & " window.__vbaDomMutationObserverInstalled=true;"
    js = js & " return true;"
    js = js & "})();"
    
    Call CDP_.jsEval(js)
    
    ' Get initial count
    res = CDP_.jsEval("window.__vbaDomMutationCount||0;")
    res = ParseSimpleValue(res) ' プリミティブな戻り値を取得
    If res = "" Then res = "0"
    startCount = CLng(res)
    
    startTime = Timer
    Do
        res = CDP_.jsEval("window.__vbaDomMutationCount||0;")
        res = ParseSimpleValue(res)
        If res = "" Then res = "0"
        currCount = CLng(res)
        
        If currCount > startCount Then
            WaitForDomMutation = True
            Exit Function  ' Mutated!
        End If
        
        If (Timer - startTime) > maxWaitTimeSec Then
            CDP_.printMsg WARN_, "<WaitForDomMutation> " & maxWaitTimeSec & "秒待機しましたがDOM変化が発生しませんでした。", "CDPexpansion_AdvancedWait"
            WaitForDomMutation = False
            Exit Function
        End If
        
        CDP_.sleep 0.1
        DoEvents
    Loop
End Sub


' =========================================================
' Network Idle の待機
' fetch や XMLHttpRequest などの通信が落ち着くまで待機する
' =========================================================
Public Function WaitForNetworkIdle(Optional maxWaitTimeSec As Long = 10, Optional idleThresholdMs As Long = 500) As Boolean
    Dim js As String
    Dim res As String, isIdle As String
    Dim startTime As Double
    
    ' Inject Fetch/XHR hooks
    js = ""
    js = js & "(function(){"
    js = js & " if(window.__vbaNetHookInstalled) return true;"
    js = js & " window.__vbaNetInflight=0; window.__vbaNetLastActivity=Date.now();"
    js = js & " if(window.fetch){"
    js = js & "  var _f=window.fetch;"
    js = js & "  window.fetch=function(){"
    js = js & "   window.__vbaNetInflight++; window.__vbaNetLastActivity=Date.now();"
    js = js & "   return _f.apply(this,arguments).finally(function(){"
    js = js & "    window.__vbaNetInflight--; window.__vbaNetLastActivity=Date.now();"
    js = js & "   });"
    js = js & "  };"
    js = js & " }"
    js = js & " var _open=XMLHttpRequest.prototype.open;"
    js = js & " var _send=XMLHttpRequest.prototype.send;"
    js = js & " XMLHttpRequest.prototype.open=function(){ this.__vbaTracked=true; return _open.apply(this,arguments); };"
    js = js & " XMLHttpRequest.prototype.send=function(){"
    js = js & "  if(this.__vbaTracked){"
    js = js & "   window.__vbaNetInflight++; window.__vbaNetLastActivity=Date.now();"
    js = js & "   this.addEventListener('loadend',function(){"
    js = js & "    window.__vbaNetInflight--; window.__vbaNetLastActivity=Date.now();"
    js = js & "   },{once:true});"
    js = js & "  }"
    js = js & "  return _send.apply(this,arguments);"
    js = js & " };"
    js = js & " window.__vbaNetHookInstalled=true; return true;"
    js = js & "})();"
    
    Call CDP_.jsEval(js)
    
    startTime = Timer
    Do
        ' idleThresholdMs 以上アクティビティがなく、進行中通信(Inflight)が0かを判定するJS
        js = "((window.__vbaNetInflight||0)===0 && (Date.now()-(window.__vbaNetLastActivity||0))>=" & idleThresholdMs & ");"
        res = CDP_.jsEval(js)
        
        isIdle = ParseSimpleValue(res)
        If LCase(isIdle) = "true" Then
            WaitForNetworkIdle = True
            Exit Function
        End If
        
        If (Timer - startTime) > maxWaitTimeSec Then
            CDP_.printMsg WARN_, "<WaitForNetworkIdle> " & maxWaitTimeSec & "秒待機しましたがNetworkIdleになりませんでした。", "CDPexpansion_AdvancedWait"
            WaitForNetworkIdle = False
            Exit Function
        End If
        
        CDP_.sleep 0.1
        DoEvents
    Loop
End Sub




' =========================================================
' 1. 完全静止・非同期処理待ち (Full Idle)
' 用途（例え）: 「完全武装のガードマン」
' SPA等の複雑なページで、画面描画・裏側の通信・値の自動変更など
' 全ての挙動が「シーン」と静けさを取り戻すまで絶対に次へ進ませない堅牢な待機です。
' =========================================================
Public Function WaitForFullIdle(Optional idleTimeoutMs As Long = 1000, Optional maxTimeoutMs As Long = 30000) As Boolean
    Dim js As String
    Dim res As String
    
    js = "new Promise(function(resolve, reject){"
    js = js & "  var maxTimer = setTimeout(function(){ done('timeout'); }, " & maxTimeoutMs & ");"
    js = js & "  var idleTime = " & idleTimeoutMs & ";"
    js = js & "  var s = window.__niw || { req: 0, lastMut: Date.now(), onchange: null };"
    js = js & "  var settled = false, tm = null;"
    js = js & "  function done(v) {"
    js = js & "    if(settled) return; settled = true;"
    js = js & "    if(tm) clearTimeout(tm); clearTimeout(maxTimer);"
    js = js & "    resolve(v);"
    js = js & "  }"
    
    ' Hooks implementation (If not already installed)
    js = js & "  if(!window.__niw){"
    js = js & "    s.of = window.fetch; s.oo = XMLHttpRequest.prototype.open; s.os = XMLHttpRequest.prototype.send;"
    js = js & "    function markChanged() { s.lastMut = Date.now(); if(s.onchange) s.onchange(); }"
    js = js & "    try{window.fetch=function(){ s.req++; markChanged(); return s.of.apply(this,arguments).finally(function(){ s.req--; markChanged(); }); };}catch(e){}"
    js = js & "    try{XMLHttpRequest.prototype.open=function(){ return s.oo.apply(this,arguments); }; XMLHttpRequest.prototype.send=function(){ s.req++; markChanged(); this.addEventListener('loadend',function(){ s.req--; markChanged(); }); return s.os.apply(this,arguments); };}catch(e){}"
    js = js & "    try{ new MutationObserver(function(){ markChanged(); }).observe(document,{childList:true,subtree:true,attributes:true}); }catch(e){}"
    js = js & "    try{ var d1=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value'); if(d1&&d1.configurable){ Object.defineProperty(HTMLInputElement.prototype,'value',{ get:d1.get, set:function(v){ markChanged(); return d1.set.call(this,v); } }); } }catch(e){}"
    js = js & "    window.__niw = s;"
    js = js & "  }"
    
    js = js & "  function tryResolve() {"
    js = js & "    var quietFor = Date.now() - s.lastMut;"
    js = js & "    if(s.req === 0 && quietFor >= idleTime){ done('ok'); } else { schedule(); }"
    js = js & "  }"
    js = js & "  function schedule() {"
    js = js & "    if(tm) clearTimeout(tm);"
    js = js & "    if(s.req === 0){"
    js = js & "      var delay = Math.max(50, idleTime - (Date.now() - s.lastMut));"
    js = js & "      tm = setTimeout(tryResolve, delay);"
    js = js & "    }"
    js = js & "  }"
    js = js & "  s.onchange = schedule; setTimeout(schedule, 0);"
    js = js & "});"

    res = ProcessPromiseResult(CDP_.jsEval(js, awaitPromise:=True))
    
    If LCase(res) = "ok" Then
        WaitForFullIdle = True
    Else
        CDP_.printMsg WARN_, "<WaitForFullIdle> " & maxTimeoutMs & "ms 待機しましたが完全にアイドルになりませんでした。(結果: " & LCase(res) & ")", "CDPexpansion_AdvancedWait"
        WaitForFullIdle = False
    End If
End Function

' =========================================================
' 2. URL遷移完了待ち (URL Redirect)
' 用途（例え）: 「到着ゲートの忠犬」
' ログインボタンを押した後など、どれだけ通信やリダイレクトを挟もうと、
' 最終的に指定のURL文字列が含まれるページに到着するまでずっと待ってくれます。
' =========================================================
Public Function WaitForUrlRedirect(partialUrl As String, Optional maxWaitTimeSec As Long = 30) As Boolean
    Dim js As String
    Dim res As String
    
    js = "new Promise(function(resolve, reject){"
    js = js & "  var maxMs = " & (maxWaitTimeSec * 1000) & ";"
    js = js & "  var maxTimer = setTimeout(function(){ resolve('timeout'); }, maxMs);"
    js = js & "  var tm = setInterval(function(){"
    js = js & "    if(window.location.href.indexOf('" & Replace(partialUrl, "'", "\'") & "') > -1) {"
    js = js & "      clearInterval(tm); clearTimeout(maxTimer); resolve('ok');"
    js = js & "    }"
    js = js & "  }, 200);"
    js = js & "});"
    
    res = ProcessPromiseResult(CDP_.jsEval(js, awaitPromise:=True))
    
    If LCase(res) = "ok" Then
        WaitForUrlRedirect = True
    Else
        CDP_.printMsg WARN_, "<WaitForUrlRedirect> " & maxWaitTimeSec & "秒待機しましたがURLに '" & partialUrl & "' が含まれませんでした。", "CDPexpansion_AdvancedWait"
        WaitForUrlRedirect = False
    End If
End Function

' =========================================================
' 3. 操作起因の波及待機 (Click & Wait for Idle)
' 用途（例え）: 「石を投げて波紋が消えるまで待つ」
' 『絞り込みボタン』等を押した直後に起こる後出しの通信やDOM描画の嵐が
' ドアが閉まるように完全に落ち着くまで、クリックとセットで待つ機能です。
' =========================================================
Public Function ClickAndWaitForIdle(elem As CDPElement, Optional idleTimeoutMs As Long = 1000, Optional maxTimeoutMs As Long = 30000) As Boolean
    ' まず要素をクリック
    elem.click
    
    ' その直後から、波紋が広がり終わるまで待機
    ClickAndWaitForIdle = WaitForFullIdle(idleTimeoutMs, maxTimeoutMs)
End Function


' =========================================================
' ユーティリティ: jsEvalの戻り値(JSON文字列や生値)から結果を抽出する簡易パサー
' =========================================================
Private Function ParseSimpleValue(ByVal val As String) As String
    ' CDPElementやCDPBrowserのjsEvalの戻り値であるプリミティブ値を簡単に抽出する
    If Left(val, 1) = """" Then val = Mid(val, 2)
    If Right(val, 1) = """" Then val = Left(val, Len(val) - 1)
    ParseSimpleValue = val
End Function

Private Function ProcessPromiseResult(ByVal val As String) As String
    ' Promise自体が返す文字列を整形
    ProcessPromiseResult = ParseSimpleValue(val)
End Function
