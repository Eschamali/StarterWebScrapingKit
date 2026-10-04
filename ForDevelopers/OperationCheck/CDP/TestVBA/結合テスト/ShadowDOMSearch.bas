Attribute VB_Name = "ShadowDOMSearch"
Option Explicit

'***************************************************************************************************
' デモ: Shadow DOM 横断検索機能
'
' 作成した "ForDevelopers\OperationCheck\CDP\TestHtml\Test_CDPElement\ShadowRoot.html" を開き、通常のCSSセレクタ検索では届かないShadow DOM内部の要素に
' 直接アクセス・操作できることを確認します。
'***************************************************************************************************



Sub Test()
    '1. 指定のWebSocketForCDPへ接続
    Const UserNameBrowser As String = "User Data"
    Dim WebSocketCDP As New CDPCoreViaWebSocket
    WebSocketCDP.ReConnectCDP UserNameBrowser

    '2. 繋げたWebSocketオブジェクトを`reattachWebSocket`メソッドに渡す
    Dim b As New CDPBrowser
    b.reattachWebSocket UserNameBrowser, WebSocketCDP

    '3. テストタブに接続
    Dim br As CDPContext
    Set br = b.getTab(setMain:=True, tabName:="Shadow DOM")




    Dim elem As CDPElement
    
    Debug.Print "--------------------------------------------------------"
    Debug.Print "       通常の getElementByQuery による検索実験"
    Debug.Print "--------------------------------------------------------"
    
    ' Light DOMの要素 (取得可能)
    Set elem = br.getElementByQuery("#light-btn")
    If elem.isExist Then
        Debug.Print "Light DOM Buttonが見つかりました"
        Debug.Print "NodeType:" & elem.CurrentNodeType

        elem.click
    Else
        Debug.Print "Light DOM Buttonが見つかりません"
    End If
    
    ' Deep Shadow DOM内の要素 (通常のCSSセレクタでは取得不能)
    Set elem = br.getElementByQuery("#deep-btn")
    If elem.isExist Then
        Debug.Print "Deep Buttonが見つかりました（！？）"
    Else
        Debug.Print "Deep Buttonが見つかりませんでした (期待通りの動作です - Shadow DOM内のため)"
    End If

    Sleep 1
    
    Debug.Print "--------------------------------------------------------"
    Debug.Print "                 ShadowDom 検索実験"
    Debug.Print "--------------------------------------------------------"
    
    Dim tmp As CDPElement
    Set tmp = br.getElementByID("shadow-host-1").GetShadowRoot
    Debug.Print "GetShadowRoot直後の、NodeType:" & tmp.CurrentNodeType

    Set elem = tmp.getElementByXPath("//*[@id='shadow1-btn']") 'あえて、Xpath
    If elem.isExist Then
        Debug.Print "特殊Xpathで、#shadow1-btn を取得しました。クリックします。"
        Debug.Print "NodeType:" & elem.CurrentNodeType
        
        elem.click
        elem.jsEval "function(){ this.style.boxShadow = '0 0 15px #10b981'; }"
    End If
    Sleep 1
    
    Set elem = br.getElementByID("shadow-host-2").GetShadowRoot.getElementByID("inner-host").GetShadowRoot.getElementByID("deep-btn")   '内部ではCSSの「#」検索として機能してます
    If elem.isExist Then
        Debug.Print "#deep-btn を取得しました。クリックします。"
        Debug.Print "NodeType:" & elem.CurrentNodeType
        
        elem.click
        elem.jsEval "function(){ this.style.boxShadow = '0 0 20px #10b981'; }"
    End If
    Sleep 1
    
    Set elem = br.getElementByID("shadow-host-2").GetShadowRoot.getElementByID("inner-host").GetShadowRoot.getElementByID("deep-input")
    If elem.isExist Then
        Debug.Print "#deep-input に文字列をセットします。"
        Debug.Print "NodeType:" & elem.CurrentNodeType

        elem.value = "Deep Shadow Input Text!"
        ' 背景色・文字色をJavaScriptでダイナミックに変更して強調表示させる
        elem.jsEval "function(){ this.style.backgroundColor = '#064e3b'; this.style.borderColor = '#10b981'; this.style.color = '#fff'; }"
    End If

    Set elem = br.getElementByID("shadow-host-2").GetShadowRoot.getElementByID("inner-host").GetShadowRoot.getElementByID("deep-text")
    If elem.isExist Then
        Debug.Print "テキスト探索でDEEP要素を取得しました。文字色を変更します。"
        Debug.Print "NodeType:" & elem.CurrentNodeType

         ' 要素の style を直接書き換えて見栄えを変更
         elem.jsEval "function(){ this.style.color = '#0ea5e9'; this.style.background = 'rgba(14, 165, 233, 0.2)'; this.innerText = 'テキスト書き換えにも成功しました！'; }"
    End If

    Debug.Print "--------------------------------------------------------"
    Debug.Print "                 ShadowDom テキスト要素取り出し"
    Debug.Print "--------------------------------------------------------"
    Dim TextEles As New Collection
    Const SerchCSS As String = "input[type='text']"
    TextEles.Add br.getElementByID("shadow-host-2").GetShadowRoot.getElementByID("inner-host").GetShadowRoot.getElementByQuery(SerchCSS)
    TextEles.Add br.getElementByID("shadow-host-1").GetShadowRoot.getElementByQuery(SerchCSS)
    TextEles.Add br.getElementByQuery(SerchCSS)

    Dim e
    For Each e In TextEles
        ' 見つかった全input要素の枠線を黄色く光らせる演出
        e.jsEval "function(){ this.style.transition = 'all 0.5s'; this.style.borderColor = '#eab308'; this.style.boxShadow = '0 0 10px rgba(234, 179, 8, 0.5)'; }"
        Debug.Print "NodeType:" & e.CurrentNodeType

        Sleep 0.2
    Next

    Sleep 1
    WebSocketCDP.DisconnectCDP
End Sub

