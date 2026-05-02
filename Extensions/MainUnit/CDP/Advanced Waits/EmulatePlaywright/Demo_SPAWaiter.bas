Option Explicit

Sub SPAtest()
    '設定シートに基づくブラウザ立ち上げ
    Dim SPApage As CDPBrowser: Set SPApage = 設定シートからのCDP起動
    Dim elem As CDPElement

'    '設定セルから、ユーザ名を取得
'    Dim UserName As String
'    With ShSetting01_StartBrowser
'        UserName = .Range(.UseRangeName(2, "Demo_CDP.demoReattachmentPart2")).value
'    End With
'
'    '1. まずは、既存のTargetIDに接続できるか？
''    If Not c.reattach(UserName, existing_) Then    '前述のSessionIDを引き継ぐ場合
'    If Not SPApage.reattach(UserName) Then
'        '既存のTargetIDが消えちゃったので、別タブへの再接続フェーズへ
'        Debug.Print "既存の`targetID`への再接続に失敗。新しいタブか、今開いている直近のタブに再接続して、そこから処理を再開します。"
'
'        '2. 未接続のタブに接続
'        SPApage.newTab setMain:=True     '新しいタブ生成からでもOK
'    Else
'        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
'    End If


    '↓ここから、あなたのイメージをコードに落とし込む↓

    ' 拡張機能（exCDP_ShadowDOMSearch）のインスタンス化と初期化
    Dim extShadow As New exCDP_ShadowDOMSearch: extShadow.Init SPApage
    
    ' 拡張機能（exCDP_ShadowDOMSearch）のインスタンス化と初期化
    Dim spaWait As New exCDP_SPAWaiter: spaWait.Init SPApage

    '余計なバナーが出た際の予約クリックを仕込む
    ExecuteRegisterAutoClickerByXPath SPApage, "//button[@id='truste-consent-button']"

    'ページ遷移前のSetup
    spaWait.EnableEvents = True

    'SPAページ遷移させる ※内部で「document.readyState」を確認します
    SPApage.navigate "https://developer.servicenow.com/", isLoading

    'DOMやネットワークが落ち着く待機ロジックを仕込む
    ' --- パターン1: SPAの準備完了を待機 (DOMContentLoaded + NetworkIdle(500ms)) ---
    SPApage.printMsg info_, "NetWork監視を開始します....", "Demo"
    Debug.Print "---------------------------------------"
    Debug.Print "Waiting for SPA to be ready..."

    If spaWait.WaitForSPAReady(10, 3) Then        'リダイレクト周りのURLが絡むため、閾値を設ける
        Debug.Print "SPA ページの準備が完了しました (DOMContentLoaded & NetworkIdle)"
    Else
        Debug.Print "タイムアウト: SPA ページの準備完了を待ちきれませんでした"
    End If

    Debug.Assert spaWait.WaitForDOMStable


    '次のページ遷移に備えて、内部状態をリセット
    SPApage.printMsg info_, "NetWork監視ステータスをリセットします", "Demo"
    spaWait.ResetState

    'ボタン押下して、ページ遷移を発動
    Set elem = extShadow.getElementByDeepCss("#utility-sign-in > button")
    Debug.Assert elem.isExist   '※待機に失敗すると、ここで止まります。「document.readyState」だけでは不十分です
    elem.click

    'DOMやネットワークが落ち着く待機ロジックを仕込む
    SPApage.printMsg info_, "次のNetWork監視を開始します....", "Demo"
    Debug.Print "---------------------------------------"
    Debug.Print "Waiting for SPA to be ready..."
    If spaWait.WaitForSPAReady(60, 0) Then         'こっちはそこまで発生しない模様
        Debug.Print "SPA ページの準備が完了しました (DOMContentLoaded & NetworkIdle)"
    Else
        Debug.Print "タイムアウト: SPA ページの準備完了を待ちきれませんでした"
    End If
    SPApage.printMsg info_, "NetWork監視を終了", "Demo"
    spaWait.EnableEvents = False

    '待機が終わったら、入力
    SPApage.getElementByXPath("//input[@id='username']").value = "Insert From VBA!"     '※待機に失敗すると、ここでエラーになります
    spaWait.EnableEvents = False
    MsgBox "適切な待機ロジックが働いてるようです！", vbInformation


    'ブラウザを正常に閉じる
'    SPApage.quit
End Sub



'***************************************************************************************************
'                               ■■■ ヘルパプロシージャ ■■■
'***************************************************************************************************
Private Sub ExecuteRegisterAutoClickerByXPath(UseObject As CDPBrowser, ByVal xpath As String, _
                                             Optional ByVal TimeoutMS As Long = 30000)
    
    Dim sourceJs As String: sourceJs = GetActionJs("autoclicker")
    Dim safeXpath As String: safeXpath = Replace(xpath, "'", "\'")
    sourceJs = Replace(sourceJs, "{{XPATH}}", safeXpath)
    sourceJs = Replace(sourceJs, "{{TIMEOUT}}", CStr(TimeoutMS))
    
    Dim params As New Dictionary: params.Add "source", sourceJs
    
    Dim res As Object
    ' PASS Silent:=True to suppress the massive goog:cdp.sendCommand log
    UseObject.pageEnable
    Set res = UseObject.invokeMethod("Page.addScriptToEvaluateOnNewDocument", params)
    
    If res("identifier") = 1 Then
        Debug.Print "CDP: AutoClicker registered for XPath: " & xpath
    Else
        Debug.Print "CDP Error: AutoClicker registration failed: " & res("Error")
    End If
End Sub

' ========================================================================================
' INTERNAL JS HELPER: GetActionJs
' DESCRIPTION: Provides centralized JavaScript snippets for all browser actions.
'              All actions return a stringified JSON to ensure robust error handling in VBA.
' ========================================================================================
Private Function GetActionJs(ByVal actionType As String) As String
    Dim js As String: js = ""
    Select Case LCase(actionType)
        Case "click"
            ' Scrolls, focuses, and clicks. Standardized JSON response.
            js = js & "try { const e = arguments[0]; e.scrollIntoView({block:'center'}); e.focus(); e.click(); return JSON.stringify({status: 'ok'}); } "
            js = js & "catch(err) { return JSON.stringify({status: 'error', message: err.message}); }"
                 
        Case "input"
            ' SPA-compatible input using execCommand with bubbling event fallback.
            js = js & "try { const e = arguments[0]; const v = arguments[1]; e.scrollIntoView({block:'center',inline:'center'}); e.click(); e.focus(); e.value=''; "
            js = js & "const s = document.execCommand('insertText', false, v); if(!s){ e.value = v; e.dispatchEvent(new Event('input', {bubbles:true})); e.dispatchEvent(new Event('change', {bubbles:true})); } "
            js = js & "e.blur(); return JSON.stringify({status: 'ok'}); } catch(err) { return JSON.stringify({status: 'error', message: err.message}); }"
                 
        Case "select"
            ' Standard Select-Box handling with event dispatching.
            js = js & "try { const s = arguments[0]; const v = arguments[1]; s.scrollIntoView({block:'center',inline:'center'}); s.focus(); s.value = v; "
            js = js & "s.dispatchEvent(new Event('input', {bubbles:true})); s.dispatchEvent(new Event('change', {bubbles:true})); s.blur(); return JSON.stringify({status: 'ok'}); } "
            js = js & "catch(err) { return JSON.stringify({status: 'error', message: err.message}); }"

        Case "select_text"
            ' Normalizes text nodes to handle non-breaking spaces & dynamic content.
            js = js & "try { const s = arguments[0]; const t = arguments[1]; const n = (str) => str.replace(/[\s\u00A0]+/g, ' ').trim(); let f = false; const target = n(t); "
            js = js & "for (let i = 0; i < s.options.length; i++) { if (n(s.options[i].text) === target) { s.value = s.options[i].value; f = true; break; } } "
            js = js & "if (!f) return JSON.stringify({status: 'error', message: 'Option text not found: ' + t}); s.scrollIntoView({block:'center'}); s.focus(); "
            js = js & "s.dispatchEvent(new Event('input', {bubbles:true})); s.dispatchEvent(new Event('change', {bubbles:true})); s.blur(); return JSON.stringify({status: 'ok'}); } "
            js = js & "catch(err) { return JSON.stringify({status: 'error', message: err.message}); }"

        Case "visibility_check"
            ' Visual state validation (Display, Visibility, Opacity, and DOM attachment).
            js = js & "const e = arguments[0]; if (!e || !e.isConnected) return JSON.stringify({status: 'ok', value: false}); const s = window.getComputedStyle(e), r = e.getBoundingClientRect(); "
            js = js & "const v = (s.display !== 'none' && s.visibility !== 'hidden' && parseFloat(s.opacity || '1') > 0 && (r.width > 0 || r.height > 0 || e.getClientRects().length > 0)); "
            js = js & "return JSON.stringify({status: 'ok', value: v});"
   
        Case "ready_state"
            ' Standardized readyState query for the GetCleanResult pipeline.
            js = js & "return JSON.stringify({status: 'ok', value: document.readyState});"

        Case "shadow_click"
            ' ASYNC: Recursively penetrates ShadowRoots with heartbeat and timeout support.
            ' [FIX] Removed 'const hb = arguments[0];'. hb() is now provided via closure.
            js = js & "const selectors = arguments[0]; const timeout = arguments[1]; const find = (arr, t) => { "
            js = js & "return new Promise((res) => { const end = Date.now() + t; const check = () => { hb(); let el = document.querySelector(arr[0]); "
            js = js & "if(el) { for(let i=1; i<arr.length; i++){ if(el.shadowRoot) el = el.shadowRoot.querySelector(arr[i]); else { el=null; break; } if(!el) break; } } "
            js = js & "if(el) res(el); else if(Date.now() < end) setTimeout(check, 100); else res(null); }; check(); }); }; "
            js = js & "const targetEl = await find(selectors, timeout); if(!targetEl) throw new Error('Shadow element not found'); "
            js = js & "targetEl.scrollIntoView({block:'center'}); targetEl.focus(); targetEl.click(); return JSON.stringify({status: 'ok'});"
        
        Case "autoclicker"
            ' IIFE: Background MutationObserver that fires a click as soon as target appears.
            js = js & "(function(x,t){const s=Date.now(),f=()=>{const e=document.evaluate(x,document,null,9,null).singleNodeValue;"
            js = js & "if(e&&(e.offsetWidth>0||e.offsetHeight>0)){e.click();e.dispatchEvent(new MouseEvent('click',{bubbles:true}));"
            js = js & "console.log('BiDi-AutoClicker: Target clicked');return 1}return 0};"
            js = js & "if(f())return;const o=new MutationObserver(()=>{if(f()||(Date.now()-s>t))o.disconnect()});"
            js = js & "o.observe(document.body||document,{childList:1,subtree:1,attributes:1})})('{{XPATH}}',{{TIMEOUT}});"
            
    End Select
    GetActionJs = js
End Function
