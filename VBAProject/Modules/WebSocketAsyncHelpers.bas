Attribute VB_Name = "WebSocketAsyncHelpers"
'***************************************************************************************************
'             WebSocket の非同期モードを円滑に行うためのヘルパーモジュールです
'                   コールバックを機能するためのモジュールとなります
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                        ■■■ VBA用の変数にコピーするためのWinAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)



'***************************************************************************************************
'                                   ■■■ 構造体定義 ■■■
'***************************************************************************************************
'https://learn.microsoft.com/ja-jp/windows/win32/api/winhttp/ns-winhttp-winhttp_web_socket_status
Private Type WINHTTP_WEB_SOCKET_STATUS
    dwBytesTransferred As Long
    eBufferType As Long
End Type



'***************************************************************************************************
'           ■■■ コールバック処理を出来るだけ安定的に、行うためのグローバル定義 ■■■
'***************************************************************************************************
Public Const BufferToAllocate As Long = 4096

'Websocket蓄積受信状況把握に使用
Public Type G_WebSocketReceiveManage
    Buffer(BufferToAllocate - 1) As Byte    '第1引数        コールバックで自動で入ってくれる
    BufferLength As Long                    '第2引数        ※事前に計算で求める必要あり
    ReceiveBytes As Long                    '第3引数        WINHTTP_WEB_SOCKET_STATUS.dwBytesTransferred
    Status As Long                          '第4引数        WINHTTP_WEB_SOCKET_STATUS.eBufferType
    CurrentPointer As Long                  '第5引数        ※事前に計算で求める必要あり
    result As Long                          '戻り値         コールバック内では無意味
    collect As Collection                   'チャンク収集   ※バラバラのデータを蓄積させる用
End Type
Global G_res As G_WebSocketReceiveManage

'フラグ管理
Global isReceiving As Boolean   'メッセージ受信済みフラグ
Global isDataReady As Boolean   '受信予約済みフラグ



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
            CopyMemory WebSocketStatus, ByVal lpvStatusInformation, LenB(WebSocketStatus)
        
        
            '========================= ステータス値　把握用 =========================
            Dim ReceivingProcessing As New WebSocketCommunicator
            With ReceivingProcessing
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


            'WINHTTP_CALLBACK_STATUS に応じたログ処理
            Select Case dwInternetStatus
                'READ_COMPLETE
                Case 524288
                    isReceiving = True
                    isDataReady = False
                    ViewLog.LogInfo "非同期処理により、受信メッセージを格納しました。呼び出し側にて、受信メッセージを処理してください。", ErrorSource

                'WRITE_COMPLETE
                Case 1048576
                    ViewLog.LogInfo "非同期処理により、送信の確認が取れました。必要に応じて、受信予約を行ってください。", ErrorSource
                    
                'REQUEST_ERROR
                Case 2097152
                    ViewLog.LogError "WebSocket の処理にて問題が発生しました。", ErrorSource
                
                'CLOSE_COMPLETE
                Case 33554432
                    ViewLog.LogInfo "WebSocket を閉じました。", ErrorSource
                    
                Case Else
                    ViewLog.LogWarn "`WINHTTP_WEB_SOCKET_STATUS.eBufferType`未定義のコードが来てます：" & dwInternetStatus, ErrorSource
            End Select


        '一応、通知しておく
        Case Else
            ViewLog.LogWarn "WebSocket 関連以外のコールバックが来たようです。　WINHTTP_STATUS_CALLBACK：" & dwInternetStatus, ErrorSource
    End Select
End Sub
