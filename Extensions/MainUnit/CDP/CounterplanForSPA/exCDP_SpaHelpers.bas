Attribute VB_Name = "exCDP_SpaHelpers"
'***************************************************************************************************
'                 SPA向け拡張（exCDP_SpaElement / exCDP_SpaIdle）の共通部品です
'
'   ・ブラウザ側で動かすJavaScriptの共通部品（可視判定、Shadow DOM を含むXPath探索）
'   ・`Runtime.callFunctionOn`の引数組み立て、ページ遷移等による失効エラーの判定
'   ※ JavaScriptの実行（EvalWithArgs）は、各Classが持つ`CDPContext`で行うため各Class側にあります
'
'* 注意事項：・単体では使いません。両クラスと一緒にインポートしてください
'            ・`Option Private Module`のため、Excelのマクロ一覧やワークシート関数には表示されません
'***************************************************************************************************
Option Explicit
Option Private Module



'***************************************************************************************************
'                           ■■■ ブラウザ側で動かすJavaScriptの部品 ■■■
'***************************************************************************************************
'* 機能　　：可視判定関数`vis(e)`を定義するJavaScript片です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：接続済み・サイズあり・visibility:visible・inert祖先なし（slot / Shadow-root も越えて辿る）
'            opacity:0 の要素は、クリック対象として正当な場合が多いため可視扱いです
'* 注意事項：関数本体の先頭に連結して使います（例："function(x){" & VisiblePredicateJs() & "...}"）
'***************************************************************************************************
Public Function VisiblePredicateJs() As String
    Dim js(5) As String
    js(0) = "const vis=e=>{if(!e||!e.isConnected||!e.getClientRects().length)return false;"
    js(1) = "const r=e.getBoundingClientRect();if(!(r.width>0||r.height>0))return false;"
    js(2) = "if(getComputedStyle(e).visibility!=='visible')return false;"
    js(3) = "let q=e,g=0;while(q&&g++<1000){if(q.nodeType===1&&q.hasAttribute('inert'))return false;"
    js(4) = "if(q.assignedSlot){q=q.assignedSlot;continue;}if(q.parentElement){q=q.parentElement;continue;}"
    js(5) = "const rt=q.getRootNode?q.getRootNode():null;q=(rt&&rt.host)?rt.host:null;}return true;};"
    VisiblePredicateJs = Join(js, vbNullString)
End Function

'***************************************************************************************************
'* 機能　　：XPath探索関数`find(x, sh)`と、Shadow-root 列挙関数`roots()`を定義するJavaScript片です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・`roots()`  … document と、全ての open な Shadow-root を深さ優先の文書順で返します
'            ・`find(x, sh)` … XPath`x`に一致する要素の配列を返します。`sh`が真なら`roots()`の全てを探索します
'            ・Shadow-root はXPathの起点にできないため、その最初の子ノードから`..`で親(=Shadow-root)へ戻って評価します
'* 注意事項：・`sh`が真の場合、XPathは`/`始まり（例：`//button[@id='save']`）であること
'            ・関数本体の先頭に連結して使います
'***************************************************************************************************
Public Function FindElementsJs() As String
    Dim js(4) As String
    js(0) = "const roots=()=>{const r=[document];const w=n=>{for(const c of n.querySelectorAll('*')){if(c.shadowRoot){r.push(c.shadowRoot);w(c.shadowRoot);}}};w(document);return r;};"
    js(1) = "const evAll=(x,c)=>{const s=document.evaluate(x,c,null,XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,null),o=[];"
    js(2) = "for(let i=0;i<s.snapshotLength;i++){const n=s.snapshotItem(i);if(n.nodeType===1)o.push(n);}return o;};"
    js(3) = "const find=(x,sh)=>{if(!sh)return evAll(x,document);const o=[];"
    js(4) = "for(const r of roots()){if(r.nodeType===9)o.push(...evAll(x,r));else if(r.firstChild)o.push(...evAll('..'+x,r.firstChild));}return o;};"
    FindElementsJs = Join(js, vbNullString)
End Function



'***************************************************************************************************
'                           ■■■ CDP呼び出しの部品 ■■■
'***************************************************************************************************
'* 機能　　：要素やページの差し替え（ページ遷移・再描画）によって起きたCDPエラーかを判定します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：このエラーは「もう一度取り直せば成功しうる」ものとして扱います
'***************************************************************************************************
Public Function IsStaleError(desc As String) As Boolean
    IsStaleError = (InStr(1, desc, "Could not find object with given id", vbTextCompare) > 0) Or _
                   (InStr(1, desc, "Cannot find context with specified id", vbTextCompare) > 0) Or _
                   (InStr(1, desc, "Cannot find default execution context", vbTextCompare) > 0) Or _
                   (InStr(1, desc, "Execution context was destroyed", vbTextCompare) > 0)
End Function

'***************************************************************************************************
'* 機能　　：`Runtime.callFunctionOn`に渡す引数(CallArgument)の配列を組み立てます
'---------------------------------------------------------------------------------------------------
'* 引数　　：values   JavaScript関数へ順に渡す値（文字列・数値・論理値・Collection 等）
'* 使い方　：ctx.jsEval "function(a,b){...}", , JsArgs("x", True)
'***************************************************************************************************
Public Function JsArgs(ParamArray values() As Variant) As Dictionary()
    Dim a() As Dictionary, i As Long
    ReDim a(0 To UBound(values))

    For i = 0 To UBound(values)
        Set a(i) = New Dictionary
        a(i).Add "value", values(i)
    Next
    JsArgs = a
End Function

'***************************************************************************************************
'* 機能　　：Shadow DOM 探索で使えるXPathの書式か（`/`始まりか）を判定します
'***************************************************************************************************
Public Function IsShadowXPath(xpath As String) As Boolean
    IsShadowXPath = (Left$(LTrim$(xpath), 1) = "/")
End Function
