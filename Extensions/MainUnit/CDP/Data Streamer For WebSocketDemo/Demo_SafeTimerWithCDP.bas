Attribute VB_Name = "Demo_SafeTimerWithCDP"
'***************************************************************************************************
'           `Chrome DevTools Protocol` と `VBA-SafeTimer`の機能を活かしたDemoです
'       このDemoでは、WebSocket関連の情報をイミディエイトウィンドウ、Excelテーブルに表示させます
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ 監視Demo ■■■
'***************************************************************************************************
'* 機能　　：従来での非同期イベント監視バージョンです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：一生、do~loop から抜け出せないため、コードとして美しくないです
'***************************************************************************************************
Sub StartDoLoopVer()
    '設定シートに基づくブラウザ立ち上げ
    Dim WebSocketDemo_doLoop As CDPBrowser: Set WebSocketDemo_doLoop = 設定シートからのCDP起動

    '拡張機能側へ継承
    Dim d As New exCDP_WebSocketEvents
    d.Init WebSocketDemo_doLoop

    'ネットワーク非同期イベントを有効化
    WebSocketDemo_doLoop.invokeMethod "Network.enable"

    'WebSocketを扱ってるDemoページへ遷移
    WebSocketDemo_doLoop.navigate "https://echo.websocket.org/.ws"

    '非同期イベント監視を開始
    Debug.Print "`Do-Loop`が始動しました。"; "Demo画面にて適当に文字を入力してみてください。"
    Do
        WebSocketDemo_doLoop.sleep 0.05  '50ms間隔で監視
        WebSocketDemo_doLoop.TakeEvents
    Loop


    'ブラウザを閉じればループを強制中断できます
End Sub

'***************************************************************************************************
'* 機能　　：`VBA-SafeTimer`での非同期イベント監視バージョンです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`StartSetTimerVer`プロシージャは、一旦終わってる点に注目です。`VBA-SafeTimer`側で`.TakeEvents`をしています
'***************************************************************************************************
Sub StartSetTimerVer()
    '設定シートに基づくブラウザ立ち上げ
    Dim WebSocketDemo_SafeTimer As CDPBrowser: Set WebSocketDemo_SafeTimer = 設定シートからのCDP起動

    'このプロシージャが終了しても、このクラスオブジェクトは保持するように組む
    Static d As New exCDP_WebSocketEvents

    '拡張機能側へ継承
    d.Init WebSocketDemo_SafeTimer

    'ネットワーク非同期イベントを有効化
    WebSocketDemo_SafeTimer.invokeMethod "Network.enable"

    'WebSocketを扱ってるDemoページへ遷移
    WebSocketDemo_SafeTimer.navigate "https://echo.websocket.org/.ws"

    '50ms間隔で、非同期イベント監視を開始
    d.StartCheckAsyncEvents 50

    Debug.Print "`VBA-SafeTimer`が始動しました。"; "Demo画面にて適当に文字を入力してみてください。"
End Sub



'***************************************************************************************************
'                               ■■■ 各WebSocketイベント情報 ■■■
'***************************************************************************************************
Sub ShowClosed(requestid As String, timestamp As Long)
    Debug.Print "webSocketがCloseされました。　 requestId: " & requestid & " , timestamp: " & timestamp
End Sub

Sub ShowCreated(requestid As String, Url As String)
    Debug.Print "webSocketがCreateされました。  requestId: " & requestid & " , url      : " & Url
End Sub

Sub ShowFrameError(requestid As String, timestamp As Long, ErrorMessage As String)
    Debug.Print "←webSocket受信中にError発生。 requestId: " & requestid & " , timestamp: " & timestamp & " , 原因: " & ErrorMessage
End Sub

Sub ShowFrameReceived(requestid As String, timestamp As Long, payloadData As String)
    Debug.Print "←webSocketが受信。　　　　　　requestId: " & requestid & " , timestamp: " & timestamp & " , RawResponse: " & payloadData
    テーブルへ追加 "受信", requestid, payloadData
End Sub

Sub ShowFrameSent(requestid As String, timestamp As Long, payloadData As String)
    Debug.Print "→webSocketが送信。　　　　　　requestId: " & requestid & " , timestamp: " & timestamp & " , RawSendMes : " & payloadData
    テーブルへ追加 "送信", requestid, payloadData
End Sub

Sub ShowHandshakeResponseReceived(requestid As String, timestamp As Long, response As Dictionary)
    Dim JsonConv As New WebJsonConverter
    Debug.Print "←webSocketが受理されました。　requestId: " & requestid & " , timestamp: " & timestamp & " , RawResponse : " & JsonConv.ConvertToJson(response)
End Sub

Sub ShowWillSendHandshakeRequest(requestid As String, timestamp As Long, response As Dictionary)
    Dim JsonConv As New WebJsonConverter
    Debug.Print "→webSocketリクエスト検知。　　requestId: " & requestid & " , timestamp: " & timestamp & " , RawResponse : " & JsonConv.ConvertToJson(response)
End Sub



'***************************************************************************************************
'                               ■■■ テーブルへ記録版 ■■■
'---------------------------------------------------------------------------------------------------
'* 詳細説明：1. 新規シートを作成してください
'            2. テーブルを「status ,リクエストID,内容」という3列構成で、作成してください
'            3. 必要に応じて、定数を書き換えてください
'***************************************************************************************************
Private Sub テーブルへ追加(Status As String, requestid As String, mes As String)
    Const TableName     As String = "テーブル1"
    Const WorkSheetName As String = "Sheet1"


    Dim SetArray(2)
    SetArray(0) = Status
    SetArray(1) = requestid
    SetArray(2) = mes

    Dim InsertDataRows      As Long: InsertDataRows = 1     '行数
    Dim InsertDataColumns   As Long: InsertDataColumns = 3  '列数

    'テーブルオブジェクトそのものを取得
    Dim TargetTableObj As ListObject: Set TargetTableObj = Sheets(WorkSheetName).ListObjects(TableName)

    'テーブルに格納してるデータ数を取得
    Dim DataRows As Long: DataRows = TargetTableObj.ListRows.Count

    'テーブル自体が確保してる行数を取得
    Dim TableRows As Long: TableRows = TargetTableObj.ListColumns(1).Range.Count

    '存在しないので、テーブル末尾に追加します。
    '※右記のテクニックを参考にした追加方式です→http://officetanaka.net/excel/vba/table/08.htm
    TargetTableObj.ListColumns(1).Range(TableRows + 1 + (DataRows = 0)).Resize(InsertDataRows, InsertDataColumns) = SetArray
    TargetTableObj.ListRows(TargetTableObj.ListRows.Count).Range.Select
End Sub
