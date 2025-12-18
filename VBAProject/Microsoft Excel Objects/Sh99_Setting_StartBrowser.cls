VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sh99_Setting_StartBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



'***************************************************************************************************
'                           ■■■ イベントプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：入力規則チェックを行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：エラー時は、デフォルト値に書き換えます。前回の値には戻せません。
'***************************************************************************************************
Private Sub Worksheet_Change(ByVal Target As Range)
    'デフォルト値
    Const DefaultValue As String = "Automation Data"


    ' --- チェック対象のセルを指定 ---
    Dim checkCell As Range: Set checkCell = Me.Range(BrowserSetting_RangeID02)

    ' --- 変更されたセル(Target)が、チェック対象セルと重なっているか判定 ---
    ' ※これを入れないと、シート上のどのセルを変更しても処理が走ってしまう
    If Intersect(Target, checkCell) Is Nothing Then Exit Sub

    ' --- 実際の入力チェック処理 ---
    Application.EnableEvents = False
    Select Case checkCell.value
        ' 条件1: セルが空っぽ（空文字列）の場合
        Case ""
            MsgBox "この項目は、入力必須です。デフォルト値に戻します。", vbExclamation, "入力エラー"
            Me.Range(BrowserSetting_RangeID02).value = DefaultValue

        ' 条件2: 特定の禁止文字列が入力された場合
        Case "User Data"
            MsgBox "セキュリティ対策のため、標準のデータディレクトリの指定はできません。" & vbCrLf & "詳細は下記をご覧ください。" & vbCrLf & "https://developer.chrome.com/blog/remote-debugging-port?hl=ja" & vbCrLf & vbCrLf & "デフォルト値に戻します。", vbCritical, "入力エラー"
            Me.Range(BrowserSetting_RangeID02).value = DefaultValue
    
        ' 条件3: 排他設定
    
    End Select
    Application.EnableEvents = True
End Sub
