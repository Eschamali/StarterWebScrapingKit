Attribute VB_Name = "Demo_WebSocket"
'***************************************************************************************************
'                          WebSocket のデモンストレーションです
'                   これを駆使すれば、FireFox の自動操作も可能です
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'* 機能　　：指定wssプロトコルに新規接続します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・WebsocketのDemoができる「wss://echo.websocket.org」へ接続し、簡単な送受信テストをします
'            ・内部の文字コード変換により、日本語も問題ありません
'* 注意事項：まだ何も受信してない状態で、受信処理をするとフリーズします
'***************************************************************************************************
Sub websocketdemo()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '接続先を設定します
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("echo.websocket.org", "")

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket success"
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode
        Debug.Print WebsocketObj.GetMessage

        '1件分の送受信をしてみる
        '※WorksheetFunction.Unichar　は絵文字を送るときに使えます
        WebsocketObj.SendMessage "うみねこ！みゃ～お！" & WorksheetFunction.Unichar(129418)
        Debug.Print WebsocketObj.GetMessage

        '後始末
        WebsocketObj.CloseWebSocket
    Else
        Debug.Print "Websocket failed"
    End If
End Sub

'***************************************************************************************************
'* 機能　　：指定wsプロトコルに新規接続します。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・Websocket経由によるChrome DevTools Protcol 操作をデモンストレーションします。全てJsonコードでのやり取りとなります
'            ・内部の文字コード変換により、日本語も問題ありません
'            ・FireFox も同じ原理なので、送るJsonコマンドが正しければ自動操作可能です
'* 注意事項：まだ何も受信してない状態で、受信処理をするとフリーズします
'***************************************************************************************************
Sub websocketdemo2()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '接続先のwsプロトコルのURIを指定します
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("127.0.0.1", "devtools/page/EE29EDB8973F1495BBA9AD424144DB3C", 9222, False)

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket success"
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode

        '1件分の送信をしてみる(接続先のブラウザにある全cookie情報抽出)
        Debug.Print WebsocketObj.SendMessage("{""id"":" & 1 & "," & _
                  """method"":""Network.getAllCookies""," & _
                  """params"":{}}")
        
        Debug.Print WebsocketObj.GetMessage
        

        '後始末
        WebsocketObj.CloseWebSocket
    Else
        Debug.Print "Websocket failed"
    End If
End Sub

'***************************************************************************************************
'* 機能　　：既存のWebSocketハンドル値を使って、再接続しやり取りの再開をします
'---------------------------------------------------------------------------------------------------
'* 注意事項：まだ何も受信してない状態で、受信処理をするとフリーズします
'***************************************************************************************************
Sub rewebsocketdemo()
    '前項で得たハンドル値
    Const ReConnectionHandle As LongPtr = 1999387763680^

    'オブジェクトを作成して、再接続用のLETメソッドにセット
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    WebsocketObj.ReConnect = ReConnectionHandle

    '送受信テスト
    WebsocketObj.SendMessage ("{""id"":" & 1 & "," & _
                  """method"":""Browser.getVersion""," & _
                  """params"":{}}")
                
    Debug.Print WebsocketObj.GetMessage()
End Sub
