VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'戻し忘れ防止用ヒューマンエラー
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    #If DEBUG_MODE Then
        ' テストモードがONの場合のみ、終了を物理的に阻止する
        Dim msg As String
        msg = "現在 [DEBUG_MODE] が ON になっています。" & vbCrLf & _
              "この状態でリリースしてはいけません。" & vbCrLf & _
              "プロジェクトプロパティから【条件付きコンパイル引数】を空にしてから閉じてください。"
        
        ' キャンセルを True にして終了を阻止
        Cancel = True
        
        ' 警告メッセージを表示
        MsgBox msg, vbCritical, "リリース前チェック警告"
        
        ' 念のため、編集モードに入っていたら保存ダイアログが出るのを防ぐ
        ThisWorkbook.Saved = True
    #End If
End Sub
