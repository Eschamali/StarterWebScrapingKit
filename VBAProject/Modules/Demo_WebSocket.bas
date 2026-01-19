Attribute VB_Name = "Demo_WebSocket"
'***************************************************************************************************
'                          WebSocket のデモンストレーションです
'                   これを駆使すれば、FireFox の自動操作も可能です
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ 同期処理 ■■■
'***************************************************************************************************
'* 機能　　：指定wssプロトコルに新規同期接続します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・WebsocketのDemoができる「wss://echo.websocket.org」へ接続し、簡単な送受信テストをします
'            ・内部の文字コード変換により、日本語も問題ありません
'* 注意事項：まだ何も受信してない状態で、受信処理をするとフリーズします
'***************************************************************************************************
Sub WebSocketDemoSync1()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '接続先を設定します
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("echo.websocket.org", "")

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket connect is success. SyncMode."
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode
        Debug.Print WebsocketObj.GetSyncMessage

        '1件分の送信をしてみる
        '※WorksheetFunction.Unichar　は絵文字を送るときに使えます
        Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("うみねこ！みゃ～お！" & WorksheetFunction.Unichar(129418))
        
        '実行結果確認
        Dim ErrorMes As New WinApiError
        If ResultCode Then Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub
        
        '受信メッセージを受け取る
        Debug.Print WebsocketObj.GetSyncMessage(, ResultCode)

        '実行結果確認
        If ResultCode Then Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub

        '後始末
        WebsocketObj.CloseWebSocket
    Else
        Debug.Print "Websocket connect is failed."
    End If
End Sub

'***************************************************************************************************
'* 機能　　：指定wsプロトコルに新規同期接続します。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・Websocket経由によるChrome DevTools Protcol 操作をデモンストレーションします。全てJsonコードでのやり取りとなります
'            ・内部の文字コード変換により、日本語も問題ありません
'            ・FireFox も同じ原理なので、送るJsonコマンドが正しければ自動操作可能です
'* 注意事項：まだ何も受信してない状態で、受信処理をするとフリーズします
'***************************************************************************************************
Sub WebSocketDemoSync2()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '接続先のwsプロトコルのURIを指定します
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("127.0.0.1", "devtools/page/61BE42DB3C9B8C64774F795A2E8E4168", 9222, False)

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket connect is success. SyncMode."
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode

        '1件分の送信をしてみる(接続先のブラウザにある全cookie情報抽出)
        Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("{""id"":" & 1 & "," & _
                  """method"":""Network.getAllCookies""," & _
                  """params"":{}}")

        '実行結果確認
        Dim ErrorMes As New WinApiError
        If ResultCode Then Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub

        '受信メッセージを受け取る
        Debug.Print WebsocketObj.GetSyncMessage(, ResultCode)

        '実行結果確認
        If ResultCode Then Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub

        '後始末
        WebsocketObj.CloseWebSocket
    Else
        Debug.Print "Websocket connect is failed."
    End If
End Sub

'***************************************************************************************************
'* 機能　　：既存のWebSocketハンドル値を使って、再接続しやり取りの再開をします
'---------------------------------------------------------------------------------------------------
'* 注意事項：まだ何も受信してない状態で、受信処理をするとフリーズします
'***************************************************************************************************
Sub ReWebSocketDemoSync()
    '前項で得たハンドル値
    Const ReConnectionHandle As LongPtr = 1510884779936^

    'オブジェクトを作成して、再接続用のLETメソッドにセット
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    WebsocketObj.ReConnect = ReConnectionHandle

    '送信テスト
    Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("{""id"":" & 1 & "," & _
                  """method"":""Browser.getVersion""," & _
                  """params"":{}}")

    '実行結果確認
    Dim ErrorMes As New WinApiError
    If ResultCode Then Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub

    '受信メッセージを受け取る
    Debug.Print WebsocketObj.GetSyncMessage(, ResultCode)

    '実行結果確認
    If ResultCode Then Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub
End Sub



'***************************************************************************************************
'                                   ■■■ 非同期処理 ■■■
'***************************************************************************************************
'* 機能　　：指定wssプロトコルに新規非同期接続します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・WebsocketのDemoができる「wss://echo.websocket.org」へ接続し、簡単な送信テストをします
'            ・内部の文字コード変換により、日本語も問題ありません
'* 注意事項：ここでは、受信はまだしません
'***************************************************************************************************
Sub WebSocketDemoASync1_1_初期化()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '接続先を設定します
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("echo.websocket.org", "", , , AddressOf WebSocketCallback)

    'フラグのリセット
    isDataReady = False

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket connect is success. AsyncMode."
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode

        '1件分の送信をしてみる
        '※WorksheetFunction.Unichar　は絵文字を送るときに使えます
        Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("うみねこ！みゃ～お！" & WorksheetFunction.Unichar(129418))
        
        '実行結果確認
        Dim ErrorMes As New WinApiError
        If ResultCode Then
            Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If

        '受信メッセージを受け取る
        Debug.Print WebsocketObj.GetAsyncMessage(, ResultCode)
            
        '実行結果確認
        If ResultCode Then
            Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If
    Else
        Debug.Print "Websocket connect is failed."
    End If
End Sub

Sub WebSocketDemoASync1_2_受信リクエスト()
    '前項で得たハンドル値
    Const ReConnectionHandle As LongPtr = 2172043420336^

    'オブジェクトを作成して、再接続用のLETメソッドにセット
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    WebsocketObj.ReConnect = ReConnectionHandle

    '受信メッセージを受け取る
    Dim ResultCode As Long
    Debug.Print WebsocketObj.GetAsyncMessage(, ResultCode)
        
    '実行結果確認
    Dim ErrorMes As New WinApiError
    If ResultCode Then
        Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
    End If
End Sub

Sub WebSocketDemoASync1_3_ハンドルから送信()
    '前項で得たハンドル値
    Const ReConnectionHandle As LongPtr = 2172043420336^

    'カウント用
    Static Count As Long
    Count = Count + 1

    'オブジェクトを作成して、再接続用のLETメソッドにセット
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    WebsocketObj.ReConnect = ReConnectionHandle

    '1件分の送信をしてみる
    '※WorksheetFunction.Unichar　は絵文字を送るときに使えます
    Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("うみねこ！みゃ～お！" & Count & WorksheetFunction.Unichar(129418))
        
    '実行結果確認
    Dim ErrorMes As New WinApiError
    If ResultCode Then
        Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
    End If
End Sub

Sub WebSocketDemoASync1_4_後始末()
    '前項で得たハンドル値
    Const ReConnectionHandle As LongPtr = 2519160849248^

    'オブジェクトを作成して、再接続用のLETメソッドにセット
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    WebsocketObj.ReConnect = ReConnectionHandle

    '後始末
    WebsocketObj.CloseWebSocket (True)
End Sub

Sub WebSocketDemoASync2_長文レスポンス()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '接続先を設定します
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("127.0.0.1", "devtools/page/1AAA01F8A73F5568DDF8FF042B62D61C", 9222, False, AddressOf WebSocketCallback)

    'フラグのリセット
    isDataReady = False

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket connect is success. AsyncMode."
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode

        '送信テスト
        Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("{""id"":" & 1 & "," & _
                  """method"":""Network.getAllCookies""," & _
                  """params"":{}}")
        
        '実行結果確認
        Dim ErrorMes As New WinApiError
        If ResultCode Then
            Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If

        '長文受信メッセージを受け取る
        Debug.Print WebsocketObj.GetAsyncMessage(, ResultCode)
            
        '実行結果確認
        If ResultCode Then
            Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If
    Else
        Debug.Print "Websocket connect is failed."
    End If
End Sub
