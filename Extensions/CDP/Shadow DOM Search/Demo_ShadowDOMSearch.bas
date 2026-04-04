Attribute VB_Name = "Demo_ShadowDOMSearch"
Option Explicit

'***************************************************************************************************
' デモ: Shadow DOM 横断検索機能
'
' 作成した TestHtml.html を開き、通常のCSSセレクタ検索では届かないShadow DOM内部の要素に
' 「CDPexpansion_ShadowDOMSearch」を使って直接アクセス・操作できることを確認します。
'***************************************************************************************************

'ワークスペースパス
'※StarterWebScrapingKitのルートフォルダ を入力してください
Private Const WORKSPACE_PATH As String = ""


Public Sub Demo_ShadowDOMSearch()
    Dim br As CDPBrowser
    Dim extShadow As CDPexpansion_ShadowDOMSearch
    
    Dim htmlPath As String
    htmlPath = "file:///" & WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_shadowRoot\TestHtml.html"
    
    ' ブラウザの起動
    Set br = 設定シートからのCDP起動(htmlPath)
    br.wait
    
    ' 拡張機能（CDPexpansion_ShadowDOMSearch）のインスタンス化と初期化
    Set extShadow = New CDPexpansion_ShadowDOMSearch
    extShadow.Init br
    
    Dim elem As CDPElement
    Dim elems As Collection
    
    br.printMsg info_, "--------------------------------------------------------", "Demo"
    br.printMsg info_, "通常の getElementByQuery による検索実験", "Demo"
    br.printMsg info_, "--------------------------------------------------------", "Demo"
    
    ' Light DOMの要素 (取得可能)
    Set elem = br.getElementByQuery("#light-btn")
    If elem.isExist Then
        br.printMsg info_, "Light DOM Buttonが見つかりました", "Demo"
        elem.click
    Else
        br.printMsg WARN_, "Light DOM Buttonが見つかりません", "Demo"
    End If
    
    ' Deep Shadow DOM内の要素 (通常のCSSセレクタでは取得不能)
    Set elem = br.getElementByQuery("#deep-btn")
    If elem.isExist Then
        br.printMsg info_, "Deep Buttonが見つかりました（！？）", "Demo"
    Else
        br.printMsg WARN_, "Deep Buttonが見つかりませんでした (期待通りの動作です - Shadow DOM内のため)", "Demo"
    End If
    
    br.sleep 1
    
    br.printMsg info_, "--------------------------------------------------------", "Demo"
    br.printMsg info_, "【拡張機能】 getElementByDeepCss による検索実験", "Demo"
    br.printMsg info_, "--------------------------------------------------------", "Demo"
    
    Set elem = extShadow.getElementByDeepCss("#shadow1-btn")
    If elem.isExist Then
         br.printMsg info_, "#shadow1-btn を取得しました。クリックします。", "Demo"
         elem.click
         br.jsEval "function(){ this.style.boxShadow = '0 0 15px #10b981'; }", objectId:=elem.getObjectId
    End If
    br.sleep 1
    
    Set elem = extShadow.getElementByDeepCss("#deep-btn")
    If elem.isExist Then
         br.printMsg info_, "#deep-btn を取得しました。クリックします。", "Demo"
         elem.click
         br.jsEval "function(){ this.style.boxShadow = '0 0 20px #10b981'; }", objectId:=elem.getObjectId
    End If
    br.sleep 1
    
    Set elem = extShadow.getElementByDeepCss("#deep-input")
    If elem.isExist Then
         br.printMsg info_, "#deep-input に文字列をセットします。", "Demo"
         elem.value = "Deep Shadow Input Text!"
         ' 背景色・文字色をJavaScriptでダイナミックに変更して強調表示させる
         br.jsEval "function(){ this.style.backgroundColor = '#064e3b'; this.style.borderColor = '#10b981'; this.style.color = '#fff'; }", objectId:=elem.getObjectId
    End If
    br.sleep 1

    br.printMsg info_, "--------------------------------------------------------", "Demo"
    br.printMsg info_, "【拡張機能】 getElementByDeepText による検索実験", "Demo"
    br.printMsg info_, "--------------------------------------------------------", "Demo"
    
    ' 部分一致で Shadow DOM 内のテキストを検索
    Set elem = extShadow.getElementByDeepText("inside DEEP Shadow", True)
    If elem.isExist Then
         br.printMsg info_, "テキスト探索でDEEP要素を取得しました。文字色を変更します。", "Demo"
         ' 要素の style を直接書き換えて見栄えを変更
         br.jsEval "function(){ this.style.color = '#0ea5e9'; this.style.background = 'rgba(14, 165, 233, 0.2)'; this.innerText = 'テキスト書き換えにも成功しました！'; }", objectId:=elem.getObjectId
    End If
    
    br.sleep 1
    
    br.printMsg info_, "--------------------------------------------------------", "Demo"
    br.printMsg info_, "【拡張機能】 複数取得 (getElementsByDeep...)", "Demo"
    br.printMsg info_, "--------------------------------------------------------", "Demo"
    Set elems = extShadow.getElementsByDeepCss("input[type='text']")
    br.printMsg info_, "ページ全体のテキストボックス数 (ShadowDOM全探索): " & elems.Count, "Demo"
    
    Dim e As CDPElement, i As Long
    i = 1
    For Each e In elems
        br.printMsg info_, i & "個目の値: " & e.value, "Demo"
        ' 見つかった全input要素の枠線を黄色く光らせる演出
        br.jsEval "function(){ this.style.transition = 'all 0.5s'; this.style.borderColor = '#eab308'; this.style.boxShadow = '0 0 10px rgba(234, 179, 8, 0.5)'; }", objectId:=e.getObjectId
        i = i + 1
        br.sleep 0.2
    Next e
    
    br.printMsg info_, "デモ終了しました。ブラウザをご確認ください。", "Demo"
    br.sleep 8
    br.quit
    
End Sub
