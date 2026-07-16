Attribute VB_Name = "Demo_WebMCP"
'==============================================================================================================
'           https://github.com/GoogleChromeLabs/webmcp-tools にあるいくつかのWebMCP-Demo を
'                       AIでもなんでもない、ExcelVBAから操作するDemoコードです
'==============================================================================================================
Option Explicit



Sub Make_AI_Original_Penguin_Pizza()
    ' 1. ピザメーカーのサイトを開いたタブ（CDPContext）に接続する
    Dim c As CDPContext: Set c = 設定シートからのCDP起動ForTab("https://googlechromelabs.github.io/webmcp-tools/demos/pizza-maker/")
    
    '-----WebSocketルートの場合は下記を追加----
'    '設定セルから、ユーザ名を取得
'    Dim UserName As String
'    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_CDP.OpenExcelWebView2")
'
'    '指定のWebSocketForCDPへ接続
'    Dim WebSocketCDP As New CDPCoreViaWebSocket
'    Debug.Print WebSocketCDP.AutoConnectBrowserCDP(UserName)
'
'    '繋げたWebSocketオブジェクトを`reattach`メソッドに渡す
'    Dim b As New CDPBrowser
'    If Not b.reattach(UserName, WebSocketCDP) Then MsgBox "「" & UserName & "」に接続できませんでした。WebSocket情報がお亡くなりです。", vbCritical, "Chrome DevTools Protocol": Exit Sub
'
'    '新しいタブに接続
'    Set c = b.newTab(setMain:=True, Url:="https://googlechromelabs.github.io/webmcp-tools/demos/pizza-maker/")
'    c.wait
    '------------------

    ' 2. 自作のWebMCPアシストクラスを初期化
    Dim p As New exCDP_WebMCP
    p.Init c
    p.printWebMCPToolsList
    
    Dim params As Dictionary
    Dim res As String
    
    ' ----------------------------------------------------
    ' 【VBE文字化け対策】
    ' サロゲートペアの絵文字をVBEに直書きすると「?」に化けるため、
    ' Excel標準のUnichar関数を使って、正しいUnicodeコードポイントから動的生成します。
    ' ----------------------------------------------------
    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction
    Dim emojiBacon As String:     emojiBacon = wf.Unichar(129363)      ' (Bacon: U+1F953)
    Dim emojiMushroom As String:  emojiMushroom = wf.Unichar(127812)   ' (Mushroom: U+1F344)
    Dim emojiCorn As String:      emojiCorn = wf.Unichar(127805)       ' (Corn: U+1F33D)
    Dim emojiOlive As String:     emojiOlive = wf.Unichar(129746)      ' (Olive: U+1FAD2)
    Dim emojiPepper As String:    emojiPepper = wf.Unichar(127798)     ' (Hot Pepper: U+1F336)
    Dim emojiPineapple As String: emojiPineapple = wf.Unichar(127821)  ' (Pineapple: U+1F34D)
    Dim emojiHerb As String:      emojiHerb = wf.Unichar(127807)       ' (Herb: U+1F33F)
    ' ----------------------------------------------------
    
    Debug.Print "=================================================="
    Debug.Print " AIオリジナル『シマハイイロ・シーサイド・ピザ』自動調理スタート！"
    Debug.Print "=================================================="
    
    ' --- ① ピザのサイズは、みんなでお腹いっぱい食べられるビッグな「Large」 ---
    Set params = New Dictionary
    params.Add "size", "Large"
    params.Add "number_of_persons", 4
    Debug.Print "1. サイズ設定: " & p.ExecuteWebMCP("set_pizza_size", params).StringKey("output")
    
    ' --- ② ベースは、香り豊かな緑色のバジルソース「Pesto（陸地）」を選択 ---
    Set params = New Dictionary
    params.Add "style", "Pesto"
    Debug.Print "2. ソース設定: " & p.ExecuteWebMCP("set_pizza_style", params).StringKey("output")
    
    ' --- ③ チーズが大好きなので、チーズレイヤー（雪）をダブル（追加）にする ---
    Set params = New Dictionary
    params.Add "layer", "cheese-layer"
    params.Add "action", "add"
    Debug.Print "3. チーズ増量: " & p.ExecuteWebMCP("toggle_layer", params).StringKey("output")
    
    ' --- ④ 【トッピング】ジューシーな旨味「ベーコン」をLargeサイズで7個 ---
    Set params = New Dictionary
    params.Add "topping", emojiBacon
    params.Add "size", "Large"
    params.Add "count", 7
    Debug.Print "4. ベーコン追加: " & p.ExecuteWebMCP("add_topping", params).StringKey("output")
    
    ' --- ⑤ 【トッピング】山の恵みの「マッシュルーム」を10個 ---
    Set params = New Dictionary
    params.Add "topping", emojiMushroom
    params.Add "size", "Medium"
    params.Add "count", 10
    Debug.Print "5. マッシュルーム追加: " & p.ExecuteWebMCP("add_topping", params).StringKey("output")
    
    ' --- ⑥ 【トッピング】ペンギンの黄色いくちばしを表現する「コーン」を12個 ---
    Set params = New Dictionary
    params.Add "topping", emojiCorn
    params.Add "size", "Medium"
    params.Add "count", 12
    Debug.Print "6. コーン追加: " & p.ExecuteWebMCP("add_topping", params).StringKey("output")
    
    ' --- ⑦ 【トッピング】ペンギンの黒い瞳と、地中海の海を表現する「オリーブ」をスモールで8個 ---
    Set params = New Dictionary
    params.Add "topping", emojiOlive
    params.Add "size", "Small"
    params.Add "count", 8
    Debug.Print "7. オリーブ追加: " & p.ExecuteWebMCP("add_topping", params).StringKey("output")
    
    ' --- ⑧ 【トッピング】ピリッとしたエンジニアのスパイス「唐辛子」を3個 ---
    Set params = New Dictionary
    params.Add "topping", emojiPepper
    params.Add "size", "Small"
    params.Add "count", 3
    Debug.Print "8. 唐辛子追加: " & p.ExecuteWebMCP("add_topping", params).StringKey("output")
    
    ' --- ⑨ 【トッピング / デバッグ用】ハワイアン風に「パイナップル」を5個仮置き ---
    ' （※トッピング削除のテストのために、一度あえてパインを載せます）
    Set params = New Dictionary
    params.Add "topping", emojiPineapple
    params.Add "size", "Medium"
    params.Add "count", 5
    Debug.Print "9. パイナップル仮置き: " & p.ExecuteWebMCP("add_topping", params).StringKey("output")
    
    ' --- ⑩ 【デバッグ用】「やっぱりピザにパインは認めない！」と、パインだけを一撃で全消去！ ---
    Set params = New Dictionary
    params.Add "topping", emojiPineapple
    params.Add "all", True
    Debug.Print "10. パイナップル全消去: " & p.ExecuteWebMCP("remove_topping", params).StringKey("output")
    
    ' --- ⑪ 【トッピング】仕上げのフレッシュな緑「バジルハーブ」をスモールで6個散らす ---
    Set params = New Dictionary
    params.Add "topping", emojiHerb
    params.Add "size", "Small"
    params.Add "count", 6
    Debug.Print "11. 仕上げバジル: " & p.ExecuteWebMCP("add_topping", params).StringKey("output")
    
    ' --- ⑫ 【完成 ＆ シェア】完成したオリジナルピザのシェア用URLをCDPで取得！ ---
    Set params = New Dictionary ' 引数なし
    res = p.ExecuteWebMCP("share_pizza", params).StringKey("output")
    
    Debug.Print "=================================================="
    Debug.Print " ！祝・完成！ シマハイイロ・シーサイド・ピザ！"
    Debug.Print "  " & res
    Debug.Print "=================================================="

End Sub

Sub Make_AI_Original_MysteryDoors()

    ' 1. ミステリー・ドアーズ のサイトを開いたタブ（CDPContext）に接続する
    Dim c As CDPContext: Set c = 設定シートからのCDP起動ForTab("https://googlechromelabs.github.io/webmcp-tools/demos/doors/")

    ' 2. 自作のWebMCPアシストクラスを初期化
    Dim p As New exCDP_WebMCP
    p.Init c
    p.printWebMCPToolsList
    
    Dim params As Dictionary
    Dim res As String

    ' --- ① 1つ目の扉を開ける ---
    Debug.Print "1. 1つ目の扉をOpen: " & p.ExecuteWebMCP("openDoor1").StringKey("status")
    p.printWebMCPToolsList

    ' ①-1 動物と会話
    Set params = New Dictionary
    params.Add "choice", "What are you?"
    Debug.Print "1-1. What are you?: " & p.ExecuteWebMCP("talk", params).StringKey("output")

    params("choice") = "Give me a gift"
    Debug.Print "1-2. Give me a gift: " & p.ExecuteWebMCP("talk", params).StringKey("output")
    Debug.Print "1-3. 廊下に戻る: " & p.ExecuteWebMCP("returnToHallway").StringKey("status")

    p.printWebMCPToolsList

    ' --- ② 2つ目の扉を開ける ---
    Debug.Print "2. 2つ目の扉をOpen: " & p.ExecuteWebMCP("openDoor2").StringKey("status")
    p.printWebMCPToolsList
    
    ' ②-1 会話
    Debug.Print "2-1. dance: " & p.ExecuteWebMCP("dance").StringKey("output")
    Debug.Print "2-2. hide: " & p.ExecuteWebMCP("hide").StringKey("output")
    Debug.Print "2-3. 廊下に戻る: " & p.ExecuteWebMCP("returnToHallway").StringKey("status")

    p.printWebMCPToolsList

    ' --- ③ 3つ目の扉を開ける ---
    Debug.Print "3. 3つ目の扉をOpen: " & p.ExecuteWebMCP("openDoor3").StringKey("status")
    p.printWebMCPToolsList

    '③-1. ON(結果はあえて、後で確認)
'    Debug.Print "3-1. castLight: " & p.ExecuteWebMCP("castLight").StringKey("output")
    Dim tmpResult As String
    Debug.Print "3-1. RunMCPAsync - castLight"
    tmpResult = p.ExecuteWebMCPAsync("castLight")
    p.printWebMCPToolsList

    '③-2. 廊下に戻れるまで内部で待機
    Debug.Print "3-2. Wait... returnToHallway"
    p.onExistTool ("returnToHallway")
    p.printWebMCPToolsList

    '簡易的に結果を回収
    Dim castLightResult As BiDiCDPJson
    Dim timerStart As Double: timerStart = Timer
    Do
        Dim tmp As String: tmp = p.TakeResultWebMCP(tmpResult)
        If StrPtr(tmp) Then Set castLightResult = BiDiCDPJson.Parse(tmp).NodeKey("params"): Exit Do
        c.InheritanceCDPBrowser.TakeEvents
    Loop Until Timer - timerStart > 30
    
    Debug.Print "3-3. Result-castLight: " & castLightResult.StringKey("output")
    Debug.Print "3-4. 廊下に戻る: " & p.ExecuteWebMCP("returnToHallway").StringKey("status")

    c.InheritanceCDPBrowser.quit
End Sub

Sub teb()
    Dim c As CDPContext: Set c = 設定シートからのCDP起動ForTab("https://googlechromelabs.github.io/webmcp-tools/demos/smart-home/")

    '自作のWebMCPアシストクラスを初期化
    Dim p As New exCDP_WebMCP
    p.Init c
    p.printWebMCPToolsList

    'VBAからスマートホームのダッシュボードを、思い通りに再配置させるハック！
    Dim params As New Dictionary
    Dim layoutList As New Collection
    
    ' 1. 最前面に配置したいIoT機器のIDを、順番にコレクションに詰める！
    layoutList.Add "camera_front_door"          ' 玄関カメラを最優先
    layoutList.Add "lock_front_door"            ' 玄関の鍵を2番目に優先
    layoutList.Add "smart_lights_living_room"   ' リビング照明を3番目
    
    params.Add "componentIds", layoutList
    
    ' 2. 直接CDPで直撃送信！
    Debug.Print p.ExecuteWebMCP("rearrangeDOMComponents", params).StringKey("output")
End Sub
