Attribute VB_Name = "Demo_WebSocketAsync"
'***************************************************************************************************
'                          WebSocket の非同期モードデモンストレーションです
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
Sub NewConnectAsyncModeFromWorksheet()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '設定シートから、接続処理
    With ShSetting02_StartWebSocket
        'WebSocketプロトコルじゃない場合は終了
        If Not (.Range(.UseRangeName(2, "Demo_WebSocket.NewConnectSyncMode")).value) Then MsgBox "WebSocketプロトコルではないため、処理を中断します", vbCritical, "Not WebSocket": Exit Sub
        
        '接続先を設定します
        Dim ResultHandleCode As LongPtr
        ResultHandleCode = WebsocketObj.Init(.Range(.UseRangeName(3, "Demo_WebSocket.NewConnectAsyncMode")).value, .Range(.UseRangeName(6, "Demo_WebSocket.NewConnectAsyncMode")).value, .Range(.UseRangeName(4, "Demo_WebSocket.NewConnectAsyncMode")).value, .Range(.UseRangeName(5, "Demo_WebSocket.NewConnectAsyncMode")).value, AddressOf WebSocketCallback)
        
        '成功判定
        If ResultHandleCode Then
            '設定シートに記録
            .Range(.UseRangeName(1, "Demo_WebSocket.NewConnectSyncMode")).value = ResultHandleCode
            
            '成功通知
            MsgBox "WebSocketへの接続を確立しました。", vbInformation, "Async Success"

            '過去のデータを削除する。メモリ上のゴミをクリーンするため
            With ShSetting02_StartWebSocket
                .CleanReceiveBox
                .CleanReceiveBoxTable
                .InitializeBuffer G_res.Buffer, G_res.CurrentPointer, G_res.BufferLength
            End With
            Set G_res.collect = New Collection

        Else
            '失敗通知
            MsgBox "WebSocketへの接続に失敗しました。" & vbCrLf & "VBEから、イミディエイトを御覧ください", vbCritical, "Failure"
        End If
    End With
    
    '受信予約
    AsyncWebsocketReciveFromWorksheet
End Sub

Sub AsyncWebsocketSendFromWorksheet()
    'エラーコードの翻訳用
    Dim GetErrorMes As New WinApiError

    'ログ把握用クラス
    Dim ViewLog As New Logger

    '設定シートから、送信処理
    With ShSetting02_StartWebSocket
        '未設定の場合は即抜け
        If .Range(.UseRangeName(1, "Demo_WebSocketAsync.AsyncWebsocketSend")).value = "" Then MsgBox "WebSocketへの接続が出来ていません。" & vbCrLf & "送受信の前に、接続処理をして下さい・", vbCritical, "Not Ready": Exit Sub
        
        'オブジェクトを作成して、再接続用のLETメソッドにセット
        Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
        WebsocketObj.ReConnect = .Range(.UseRangeName(1, "Demo_WebSocketAsync.AsyncWebsocketSendFromWorksheet")).value

        '送信リクエストする
        Dim ResultCode As Long
        ResultCode = WebsocketObj.SendMessage(.Range(.UseRangeName(7, "Demo_WebSocketAsync.AsyncWebsocketSendFromWorksheet")).value)

        'エラーチェック
        If ResultCode Then
            ViewLog.LogError GetErrorMes.GetMessage(ResultCode, "winhttp"), "Demo_WebSocketAsync.AsyncWebsocketSendFromWorksheet", ResultCode
        Else
            ViewLog.LogInfo GetErrorMes.GetMessage(ResultCode, "winhttp"), "Demo_WebSocketAsync.AsyncWebsocketSendFromWorksheet"
        End If
    End With
End Sub

Sub AsyncWebsocketReciveFromWorksheet()
    'エラーコードの翻訳用
    Dim GetErrorMes As New WinApiError

    'ログ把握用クラス
    Dim ViewLog As New Logger

    '設定シートから、送信処理
    With ShSetting02_StartWebSocket
        '未設定の場合は即抜け
        If .Range(.UseRangeName(1, "Demo_WebSocketAsync.AsyncWebsocketReciveFromWorksheet")).value = "" Then MsgBox "WebSocketへの接続が出来ていません。" & vbCrLf & "送受信の前に、接続処理をして下さい・", vbCritical, "Not Ready": Exit Sub
        
        'オブジェクトを作成して、再接続用のLETメソッドにセット
        Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
        WebsocketObj.ReConnect = .Range(.UseRangeName(1, "Demo_WebSocketAsync.AsyncWebsocketReciveFromWorksheet")).value

        '受信リクエストする
        Dim ResultCode As Long
        ResultCode = WebsocketObj.GetMessageForAsync

        'エラーチェック
        If ResultCode Then
            ViewLog.LogError GetErrorMes.GetMessage(ResultCode, "winhttp"), "Demo_WebSocketAsync.AsyncWebsocketReciveFromWorksheet", ResultCode
        Else
            ViewLog.LogInfo GetErrorMes.GetMessage(ResultCode, "winhttp"), "Demo_WebSocketAsync.AsyncWebsocketReciveFromWorksheet"
        End If
    End With
End Sub



'***************************************************************************************************
'                     ■■■ VBEハードコーディング Demo プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：指定wssプロトコルに非同期モードとして、新規接続します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・WebsocketのDemoができる「wss://echo.websocket.org」へ接続し、非同期モードを開始します
'            ・内部の文字コード変換により、日本語も問題ありません
'***************************************************************************************************
Sub AsyncWebsocketDemo()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    
    '接続先を設定します
    Dim ResultHandleCode As LongPtr
    'ResultHandleCode = WebsocketObj.Init("echo.websocket.org", "", , , AddressOf WebSocketCallback)
    ResultHandleCode = WebsocketObj.Init("127.0.0.1", "devtools/page/973B6AB28BFBCDD79A2BFFEAC8375589", 9222, False, AddressOf WebSocketCallback)
    
    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket Success"
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode
        
        'ハンドル値をセルに残す
        With ShSetting02_StartWebSocket
            .Range(.UseRangeName(1, "Demo_WebSocketAsync.AsyncWebsocketDemo")) = ResultHandleCode
        End With

        '過去のデータを削除する。メモリ上のゴミをクリーンするため
        With ShSetting02_StartWebSocket
            .CleanReceiveBox
            .CleanReceiveBoxTable
            .InitializeBuffer G_res.Buffer, G_res.CurrentPointer, G_res.BufferLength
        End With
        Set G_res.collect = New Collection

    Else
        Debug.Print "Websocket failed"
    End If
End Sub


Sub AsyncWebsocketRecive()
    '接続済みのハンドル情報を取得する
    Dim ReConnectionHandle As LongPtr
    With ShSetting02_StartWebSocket
        ReConnectionHandle = .Range(.UseRangeName(1, "Demo_WebSocketAsync.AsyncWebsocketDemo"))
    End With

    'オブジェクトを作成して、再接続用のLETメソッドにセット
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    WebsocketObj.ReConnect = ReConnectionHandle

    '受信リクエスト
    If WebsocketObj.GetMessageForAsync = 0 Then Debug.Print "受信予約しました"
End Sub

Sub 受信したのをテーブルに一気に追加()
    ShSetting02_StartWebSocket.AddTableReceiveDatas True
End Sub
