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
