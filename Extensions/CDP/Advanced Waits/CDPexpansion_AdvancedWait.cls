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
Public Sub WaitForDomMutation(Optional maxWaitTimeSec As Long = 5)
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
            Exit Sub  ' Mutated!
        End If
        
        If (Timer - startTime) > maxWaitTimeSec Then
            MsgBox "<WaitForDomMutation> " & maxWaitTimeSec & "秒待機しましたがDOM変化が発生しませんでした。", vbCritical
            End
        End If
        
        CDP_.sleep 0.1
        DoEvents
    Loop
End Sub


' =========================================================
' Network Idle の待機
' fetch や XMLHttpRequest などの通信が落ち着くまで待機する
' =========================================================
Public Sub WaitForNetworkIdle(Optional maxWaitTimeSec As Long = 10, Optional idleThresholdMs As Long = 500)
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
            Exit Sub
        End If
        
        If (Timer - startTime) > maxWaitTimeSec Then
            MsgBox "<WaitForNetworkIdle> " & maxWaitTimeSec & "秒待機しましたがNetworkIdleになりませんでした。", vbCritical
            End
        End If
        
        CDP_.sleep 0.1
        DoEvents
    Loop
End Sub


' =========================================================
' ユーティリティ: jsEvalの戻り値(JSON文字列や生値)から結果を抽出する簡易パサー
' =========================================================
Private Function ParseSimpleValue(ByVal val As String) As String
    ' CDPElementやCDPBrowserのjsEvalの戻り値であるプリミティブ値を簡単に抽出する
    If Left(val, 1) = """" Then val = Mid(val, 2)
    If Right(val, 1) = """" Then val = Left(val, Len(val) - 1)
    ParseSimpleValue = val
End Function
