Attribute VB_Name = "WebSocketAsyncHelpers"
'***************************************************************************************************
'             WebSocket の非同期モードを円滑に行うためのヘルパーモジュールです
'                   コールバックを機能するためのモジュールとなります
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                        ■■■ VBA用の変数にコピーするためのWinAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Sub memcpy Lib "msvcrt.dll" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal Count As LongPtr)



'***************************************************************************************************
'                      ■■■ コールバック処理を円滑に行うグローバル定義 ■■■
'***************************************************************************************************
'Websocket蓄積受信状況把握に使用
Public Type WebSocketReceiveManage
    Buffer(4095) As Byte    '第1引数        コールバックで自動で入ってくれる
    BufferLength As Long    '第2引数        ※事前に計算で求める必要あり
    ReceiveBytes As Long    '第3引数        WINHTTP_WEB_SOCKET_STATUS.dwBytesTransferred
    Status As Long          '第4引数        WINHTTP_WEB_SOCKET_STATUS.eBufferType
    CurrentPointer As Long  '第5引数        ※事前に計算で求める必要あり
    result As Long          '戻り値         コールバック内では無意味
    collect As Collection   'チャンク収集   ※バラバラのデータを蓄積させる用
End Type
Global G_res As WebSocketReceiveManage



'***************************************************************************************************
'                                   ■■■ 構造体定義 ■■■
'***************************************************************************************************
Private Type WINHTTP_WEB_SOCKET_STATUS
    dwBytesTransferred As Long
    eBufferType As Long
End Type



'***************************************************************************************************
'                        ■■■ メインとなるコールバックプロシージャ ■■■
'***************************************************************************************************
Public Sub WebSocketCallback(ByVal HINTERNET As LongPtr, ByVal dwContext As LongPtr, ByVal dwInternetStatus As Long, _
                                 ByVal lpvStatusInformation As LongPtr, ByVal dwStatusInformationLength As Long)
    'ログ把握用クラス
    Dim ViewLog As New Logger
    Const ErrorSource As String = "WebSocketAsyncHelpers.WebSocketCallback"

    '万が一、WebSocket 関連以外のコールバックが来ても問題ないように排除する
    Select Case dwInternetStatus
        'WebSocket関連のコールバック値を列挙する
        Case 524288, 1048576, 2097152, 33554432

            'WINHTTP_WEB_SOCKET_STATUS のポインタを基にコピー
            ' memcpy でコピー！
            ' dest: 構造体のアドレス (VarPtr)
            ' src:  ポインタの値 (lpvStatusInformation)
            ' size: 構造体のサイズ (LenB)
            Dim WebSocketStatus As WINHTTP_WEB_SOCKET_STATUS
            memcpy VarPtr(WebSocketStatus), lpvStatusInformation, LenB(WebSocketStatus)
        
        
            '========================= ステータス値　把握用 =========================
            With ShSetting02_StartWebSocket
                ViewLog.LogDebug "------------ WINHTTP_WEB_SOCKET_STATUS ------------", ErrorSource
                ViewLog.LogDebug "Bytes：" & WebSocketStatus.dwBytesTransferred, ErrorSource
                ViewLog.LogDebug "Type ：" & .Name__WINHTTP_WEB_SOCKET_BUFFER_TYPE(WebSocketStatus.eBufferType, ErrorSource) & "(" & WebSocketStatus.eBufferType & ")", ErrorSource
                ViewLog.LogDebug "---------------------------------------------------", ErrorSource
            
                ViewLog.LogDebug "WINHTTP_STATUS_CALLBACK：" & .Name__WINHTTP_STATUS_CALLBACK(dwInternetStatus, ErrorSource) & "(" & dwInternetStatus & ")", ErrorSource
            End With
            '========================================================================
        
        
            'バッファー管理処理に必要なパラメーターを適用する
            G_res.Status = WebSocketStatus.eBufferType
            G_res.ReceiveBytes = WebSocketStatus.dwBytesTransferred

            'WINHTTP_CALLBACK_STATUS に応じた分岐処理
            Dim tmp
            Select Case dwInternetStatus
                'READ_COMPLETE
                Case 524288
                    '1. バッファー情報の更新
                    ShSetting02_StartWebSocket.UpdateBufferInfo G_res.CurrentPointer, G_res.BufferLength, G_res.ReceiveBytes

                    '2. 受信処理を託す
                    tmp = ShSetting02_StartWebSocket.CommonWinHttpWebSocketReceive(G_res)
                    
                    '3. 全ての受信を終えたときの処理
                    '　UTF8_MESSAGE_BUFFER_TYPE  ：Buffer には、UTF-8    メッセージ全体またはその最後の部分が含まれます。
                    If G_res.Status = 2 Then
                        '4. 完成品のプレーンテキストをキューのCollectionに蓄積
                        ShSetting02_StartWebSocket.ReceiveBox = CStr(tmp)
                        
                        '5. 一時蓄積データも初期化
                        Set G_res.collect = New Collection
                        ViewLog.LogInfo "一時Collectionキューをクリーンしました。", ErrorSource
                    
                    '　BINARY_MESSAGE_BUFFER_TYPE：Buffer には、バイナリ メッセージ全体またはその最後の部分が含まれます。
                    ElseIf G_res.Status = 0 Then
                        '4. 完成品のバイナリデータをキューのCollectionに蓄積
                        ShSetting02_StartWebSocket.ReceiveBox = CByte(tmp)
                    
                        '5. 一時蓄積データも初期化
                        Set G_res.collect = New Collection
                        ViewLog.LogInfo "一時Collectionキューをクリーンしました。", ErrorSource

                    End If
                
                'WRITE_COMPLETE
                Case 1048576
                    ViewLog.LogInfo "送信できました！", ErrorSource
                    
                'REQUEST_ERROR
                Case 2097152
                    ViewLog.LogError "WebSocket の処理にて問題が発生しました。", ErrorSource
                
                'CLOSE_COMPLETE
                Case 33554432
                    ViewLog.LogInfo "WebSocket を閉じました。", ErrorSource

            End Select


        '一応、通知しておく
        Case Else
            ViewLog.LogWarn "WebSocket 関連以外のコールバックが来たようです。　WINHTTP_STATUS_CALLBACK：" & dwInternetStatus, ErrorSource
    End Select
End Sub
