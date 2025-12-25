VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShSetting02_StartWebSocketAsync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'***************************************************************************************************
'                   WebSocket にまつわる共通処理をまとめたものです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                      ■■■ このシートのみ定義されてるセル名称定義 ■■■
'***************************************************************************************************
Property Get UseRangeName(id As Long, ErrorSource As String) As String
    Select Case id
        Case 1: UseRangeName = "WebSocketハンドル"
        Case Else: Err.Raise 1 + vbObjectError, ErrorSource, "渡されたセルIDは、定義されてません"
    End Select
End Property



'***************************************************************************************************
'                 ■■■ 送受信処理をうまくいくためのヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：バッファ初期化
'***************************************************************************************************
Public Sub InitializeBuffer(responseBuffer() As Byte, CurrentPointer As Long, BufferLength As Long)
    Erase responseBuffer
    CurrentPointer = 0
    BufferLength = UBound(responseBuffer) + 1
End Sub

'***************************************************************************************************
'* 機能　　：バッファ管理情報更新
'***************************************************************************************************
Public Sub UpdateBufferInfo(CurrentPointer As Long, BufferLength As Long, ReceiveBytes As Long)
    CurrentPointer = CurrentPointer + ReceiveBytes
    BufferLength = BufferLength - ReceiveBytes
End Sub

'***************************************************************************************************
'* 機能　　：プレーンテキスト「文字列(String)」で送る際、UTF-8バイト配列(Byte())に変換します
'---------------------------------------------------------------------------------------------------
'* 引数    ：str        プレーンテキスト「文字列(String)」
'* 返り値  ：UTF-8バイト配列
'---------------------------------------------------------------------------------------------------
'* 機能説明：ErrorCode：12152　対策
'***************************************************************************************************
Public Function StringToUtf8Bytes(ByVal str As String) As Byte()
    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText str
        
        ' ストリームの先頭に戻ってバイト型に変更
        .Position = 0
        .Type = 1 ' adTypeBinary
        
        ' BOM（先頭3バイト）をスキップして読む
        .Position = 3
        StringToUtf8Bytes = .Read
        
        .Close
    End With
End Function

'***************************************************************************************************
'* 機能　　：バイナリデータをプレーンテキスト「UTF-8文字列(String)」に変換します
'---------------------------------------------------------------------------------------------------
'* 引数    ：bytes  バイナリデータ
'* 返り値  ：プレーンテキスト「UTF-8文字列(String)」
'---------------------------------------------------------------------------------------------------
'* 注意事項：当然ですが、バイナリデータが、UTF-8バイト配列になっていないと、文字化けします
'***************************************************************************************************
Public Function Utf8BytesToString(ByRef Bytes() As Byte) As String
    With CreateObject("ADODB.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .Write Bytes
        
        ' 先頭に戻ってテキストモードに変更
        .Position = 0
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        
        ' 文字列として読み出す
        Utf8BytesToString = .ReadText
        
        .Close
    End With
End Function

'***************************************************************************************************
'* 機能　　：チャンクごとに区切られたコレクションバイト配列を一気にUTF-8変換します
'---------------------------------------------------------------------------------------------------
'* 引数    ：collect  コレクションバイト配列
'* 返り値  ：プレーンテキスト「UTF-8文字列(String)」
'---------------------------------------------------------------------------------------------------
'* 注意事項：当然ですが、各コレクションバイナリデータが、UTF-8バイト配列になっていないと、文字化けします
'***************************************************************************************************
Public Function MakeResponseMessage(collect As Collection) As String
    Dim totalSize As Long
    Dim chunk As Variant
    Dim fullBytes() As Byte
    Dim currentPos As Long
    Dim i As Long
    
    ' 1. まず合計サイズを計算する
    totalSize = 0
    For Each chunk In collect
        ' 配列のサイズを加算（UBound+1）
        totalSize = totalSize + (UBound(chunk) + 1)
    Next chunk
    
    ' 2. 巨大な配列を確保
    ReDim fullBytes(0 To totalSize - 1)
    
    ' 3. バイトデータをコピーして結合       Task：バイナリデータそのものが欲しいならここで取り出す
    currentPos = 0
    For Each chunk In collect
        For i = 0 To UBound(chunk)
            fullBytes(currentPos) = chunk(i)
            currentPos = currentPos + 1
        Next i
    Next chunk
    
    ' 4. ここでUTF-8変換！（さっきの関数を使う）
    ' ※末尾のヌル文字(vbNullChar)除去は、UTF-8変換後にやるか、
    '   変換前にサイズ調整するかですが、ADODBはヌル文字もそのまま変換してくれるので
    '   変換後にReplaceかLeftで切るのが安全です。
    
    Dim resultStr As String
    resultStr = Utf8BytesToString(fullBytes)
    
    ' 5. ヌル文字除去（もしあれば）
    Dim nullPos As Long
    nullPos = InStr(resultStr, vbNullChar)
    If nullPos > 0 Then
        MakeResponseMessage = Left(resultStr, nullPos - 1)
    Else
        MakeResponseMessage = resultStr
    End If
End Function
