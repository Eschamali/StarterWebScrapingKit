Attribute VB_Name = "Demo_WebSocket"
'***************************************************************************************************
'                       WebSocket の同期処理用のデモンストレーションです
'                       これを駆使すれば、FireFox の自動操作も可能です
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                     ■■■ ワークシート上での Demo プロシージャ ■■■
'***************************************************************************************************
'* 機能　  ：ワークシート上の設定に基づいて、WebSocket接続を行います
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・WebsocketのDemoができる「wss://echo.websocket.org」へ接続し、簡単な送受信テストをします
'            ・内部の文字コード変換により、日本語も問題ありません
'* 注意事項：まだ何も受信してない状態で、受信処理をするとフリーズします
'***************************************************************************************************
Sub NewConnectSyncModeFromWorksheet()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '設定シートから、接続処理
    With ShSetting02_StartWebSocket
        'WebSocketプロトコルじゃない場合は終了
        If Not (.Range(.UseRangeName(2, "Demo_WebSocket.NewConnectSyncMode")).value) Then MsgBox "WebSocketプロトコルではないため、処理を中断します", vbCritical, "Not WebSocket": Exit Sub
        
        '接続先を設定します
        Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init(.Range(.UseRangeName(3, "Demo_WebSocket.NewConnectSyncModeFromWorksheet")).value, .Range(.UseRangeName(6, "Demo_WebSocket.NewConnectSyncModeFromWorksheet")).value, .Range(.UseRangeName(4, "Demo_WebSocket.NewConnectSyncModeFromWorksheet")).value, .Range(.UseRangeName(5, "Demo_WebSocket.NewConnectSyncModeFromWorksheet")).value)
        
        '成功判定
        If ResultHandleCode Then
            '設定シートに記録
            .Range(.UseRangeName(1, "Demo_WebSocket.NewConnectSyncMode")).value = ResultHandleCode
            
            '成功通知
            MsgBox "WebSocketへの接続を確立しました。", vbInformation, "Sync Success"
        Else
            '失敗通知
            MsgBox "WebSocketへの接続に失敗しました。" & vbCrLf & "VBEから、イミディエイトを御覧ください", vbCritical, "Failure"
        End If
    End With
End Sub

'***************************************************************************************************
'* 機能　  ：ワークシート上の送信内容を送信後、受信処理をします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：簡単な送受信テストをします。受信内容は、テーブルに蓄積されます
'* 注意事項：同期モードの場合、受信データが空だとなにかしらのデータが来るまで一生、フリーズします。
'            そのため、やまびこ方式のWebSocketでのテストを推奨します。
'***************************************************************************************************
Sub SendAndReceiveFromWorksheet()
    'エラーコードの翻訳用
    Dim GetErrorMes As New WinApiError

    '設定シートから、送受信処理
    With ShSetting02_StartWebSocket
        '未設定の場合は即抜け
        If .Range(.UseRangeName(1, "Demo_WebSocket.SendAndReceiveFromWorksheet")).value = "" Then MsgBox "WebSocketへの接続が出来ていません。" & vbCrLf & "送受信の前に、接続処理をして下さい・", vbCritical, "Not Ready": Exit Sub
        
        'オブジェクトを作成して、再接続用のLETメソッドにセット
        Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
        WebsocketObj.ReConnect = .Range(.UseRangeName(1, "Demo_WebSocket.SendAndReceive")).value



        '送信する
        Dim ResultCode As Long
        ResultCode = WebsocketObj.SendMessage(.Range(.UseRangeName(7, "Demo_WebSocket.SendAndReceiveFromWorksheet")).value)
        
        'エラーチェック
        If ResultCode Then
            MsgBox "送信処理にて、エラーが発生しました。" & vbCrLf & vbCrLf & "＜詳細＞" & vbCrLf & GetErrorMes.GetMessage(ResultCode, "winhttp"), vbCritical, "ErrorCode：" & ResultCode
            
            '後始末
            CleanWebsocketHandle .Range(.UseRangeName(1, "Demo_WebSocket.SendAndReceiveFromWorksheet")).value
            .Range(.UseRangeName(1, "Demo_WebSocket.SendAndReceiveFromWorksheet")).ClearContents
            Exit Sub
        End If
        
        '受信する
        Dim ResponseText As String
        ResponseText = WebsocketObj.GetMessageForSync

        'エラーチェック
        If StrPtr(ResponseText) = 0 Then
            MsgBox "受信処理にて、エラーが発生しました。" & vbCrLf & "VBEで、イミディエイトを御覧ください。", vbCritical, "Failure"

            '後始末
            CleanWebsocketHandle .Range(.UseRangeName(1, "Demo_WebSocket.SendAndReceiveFromWorksheet")).value
            .Range(.UseRangeName(1, "Demo_WebSocket.SendAndReceiveFromWorksheet")).ClearContents
            Exit Sub
        End If

        'テーブルに格納
        .AddReceiveBoxTable ResponseText

        'Downloadsフォルダにも保存しておく
        ShSetting02_StartWebSocket.SaveFileUTF8 ResponseText, Environ("USERPROFILE") & "\Downloads", "ResultWebSocket.txt"
    End With

    '終了通知
    MsgBox "送信処理を終え、1件のメッセージを受信しました", vbInformation, "Success"
End Sub

'***************************************************************************************************
'* 機能　  ：ワークシート上にあるWebSocketハンドルで、破棄処理をします
'***************************************************************************************************
Sub CleanWebsocketHandleFromWorksheet()
    With ShSetting02_StartWebSocket
        '未設定の場合は即抜け
        If .Range(.UseRangeName(1, "Demo_WebSocket.SendAndReceiveFromWorksheet")).value = "" Then MsgBox "WebSocketへの接続が出来ていません。" & vbCrLf & "送受信の前に、接続処理をして下さい・", vbCritical, "Not Ready": Exit Sub
        
        CleanWebsocketHandle .Range(.UseRangeName(1, "Demo_WebSocket.SendAndReceiveFromWorksheet")).value
    End With
End Sub


'***************************************************************************************************
'                     ■■■ VBEハードコーディング Demo プロシージャ ■■■
'***************************************************************************************************
'* 機能　  ：指定wssプロトコルに新規接続します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・WebsocketのDemoができる「wss://echo.websocket.org」へ接続し、簡単な送受信テストをします
'            ・内部の文字コード変換により、日本語も問題ありません
'* 注意事項：まだ何も受信してない状態で、受信処理をするとフリーズします
'***************************************************************************************************
Sub WebSocketSyncDemo1()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '接続先を設定します
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("echo.websocket.org", "")

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket connect is success."
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode
        Debug.Print WebsocketObj.GetMessageForSync

        '1件分の送受信をしてみる
        '※WorksheetFunction.Unichar　は絵文字を送るときに使えます
        Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("うみねこ！みゃ～お！" & WorksheetFunction.Unichar(129418))
        
        '実行結果確認
        Dim ErrorMes As New WinApiError
        If ResultCode Then Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub
        
        '受信メッセージを受け取る
        Debug.Print WebsocketObj.GetMessage(ResultCode)

        '実行結果確認
        If ResultCode Then Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub


        '---- 後始末 (ハンドルの再利用の場合は、コメントアウトしてね) ----
        CleanWebsocketHandle ResultHandleCode
    Else
        Debug.Print "Websocket connect is failed."
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
Sub WebSocketSyncDemo2()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '接続先のwsプロトコルのURIを指定します
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("127.0.0.1", "devtools/page/EE29EDB8973F1495BBA9AD424144DB3C", 9222, False)

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket connect is success."
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode

        '1件分の送信をしてみる(接続先のブラウザにある全cookie情報抽出)
        Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("{""id"":" & 1 & "," & _
                  """method"":""Network.getAllCookies""," & _
                  """params"":{}}")

        '実行結果確認
        Dim ErrorMes As New WinApiError
        If ResultCode Then Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub

        '受信メッセージを受け取る
        Debug.Print WebsocketObj.GetMessage(ResultCode)

        '実行結果確認
        If ResultCode Then Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub

        '---- 後始末 (ハンドルの再利用の場合は、コメントアウトしてね) ----
        CleanWebsocketHandle ResultHandleCode
    Else
        Debug.Print "Websocket connect is failed."
    End If
End Sub

'***************************************************************************************************
'* 機能　　：既存のWebSocketハンドル値を使って、再接続しやり取りの再開をします
'---------------------------------------------------------------------------------------------------
'* 注意事項：まだ何も受信してない状態で、受信処理をするとフリーズします
'***************************************************************************************************
Sub ReConnectWebSocketSyncDemo()
    '前項で得たハンドル値
    Const ReConnectionHandle As LongPtr = 585679133072^

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
    Debug.Print WebsocketObj.GetMessage(ResultCode)

    '実行結果確認
    If ResultCode Then Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"): Exit Sub
End Sub
