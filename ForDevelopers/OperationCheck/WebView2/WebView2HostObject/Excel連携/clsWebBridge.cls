VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWebBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================
'* クラス名：clsWebBridge
'* 役割　　：WebView2 (JS) に公開して、直接メソッドを叩かせるための窓口クラス
'===================================================================================================
Option Explicit

' クラスの先頭（メンバ変数として長生きさせておく）
Private m_Voice As Object

' 【超重要】JSから呼ばれるメソッドの引数は、必ず「ByVal」にしてください！
' （デフォルトの ByRef のままだと、COMマーシャリングで型不一致エラーになりやすいです）

' 1. 動作確認用：JSからメッセージを受け取り、時刻を添えて即座に返事（戻り値）を返す
Public Function Ping(ByVal msg As String) As String
    Ping = "VBA応答 [" & Format$(Now, "hh:nn:ss") & "]: " & msg & " を受け取ったよ！??"
End Function

' 2. 【Webの禁断突破】OSネイティブのフォルダ選択ダイアログを出し、生のフルパスをJSへ返す！
Public Function PickLocalFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "WebView2 (JS) からのリクエストでフォルダを選択"
        If .show = -1 Then
            PickLocalFolder = .SelectedItems(1)
        Else
            PickLocalFolder = "(キャンセルされました)"
        End If
    End With
End Function

' 3. Excelシートとリアルタイム連動！値を書き込んで色を塗り、シート全体の合計値を戻り値で返す！
Public Function UpdateSheetAndCalc(ByVal targetCell As String, ByVal numValue As Long, ByVal colorHex As String) As Double
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 値を書き込み
    ws.Range(targetCell).Value = numValue

    ' HTMLカラーコード (#RRGGBB) を VBA の RGB に変換してセルを塗る
    If Len(colorHex) = 7 And Left$(colorHex, 1) = "#" Then
        Dim r As Long, g As Long, b As Long
        r = CLng("&H" & Mid$(colorHex, 2, 2))
        g = CLng("&H" & Mid$(colorHex, 4, 2))
        b = CLng("&H" & Mid$(colorHex, 6, 2))
        ws.Range(targetCell).Interior.Color = RGB(r, g, b)
    End If

    ' 例として A1:A10 の合計値を計算して、即座に JS 側へ返却する！
    UpdateSheetAndCalc = Application.WorksheetFunction.Sum(ws.Range("A1:A10"))
End Function

' 4. OSの音声合成（非同期版！ 1ミリもフリーズしない！）
Public Sub Speak(ByVal text As String)
    If m_Voice Is Nothing Then Set m_Voice = CreateObject("SAPI.SpVoice")
    
    ' 第2引数に「1 (SVSFlagsAsync)」＋「2 (前の音声をキャンセルして喋り直す)」＝ 3 を渡す！
    m_Voice.Speak text, 3
End Sub
