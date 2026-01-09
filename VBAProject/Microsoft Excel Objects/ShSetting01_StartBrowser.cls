VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShSetting01_StartBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'***************************************************************************************************
'                   CDP経由で、ブラウザ操作する際の準備関連のシートです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                      ■■■ このシートのみ定義されてるセル名称定義 ■■■
'***************************************************************************************************
Property Get UseRangeName(id As Long, ErrorSource As String) As String
    Select Case id
        Case 1: UseRangeName = "ブラウザexeパス"
        Case 2: UseRangeName = "ユーザーデータフォルダ名"
        Case 3: UseRangeName = "追加の起動引数一式"
        Case 4: UseRangeName = "Chromeで使用"
        Case 5: UseRangeName = "既存のデバッグプロセスを使用"
        Case 6: UseRangeName = "常にブラウザを新規で起動"
        Case 7: UseRangeName = "イミディエイトに表示"
        Case 8: UseRangeName = "ブラウザから受信した生のJSONをイミディエイトに表示"
        Case 9: UseRangeName = "ログファイル保存"
        Case Else: Err.Raise 1 + vbObjectError, ErrorSource, "渡されたセルIDは、定義されてません"
    End Select
End Property



'***************************************************************************************************
'                      ■■■ 事実上、このシートしか使用しないテーブル名 ■■■
'***************************************************************************************************
Property Get UseTableName(id As Long, ErrorSource As String) As String
    Select Case id
        Case 1: UseTableName = "PreviousHandleInfo"
        Case 2: UseTableName = "PreviousBrowserInfo"
        Case Else: Err.Raise 2 + vbObjectError, ErrorSource, "渡されたテーブルIDは、定義されてません"
    End Select
End Property



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
    Dim checkCell As Range: Set checkCell = Me.Range(Me.UseRangeName(2, "Me.Worksheet_Change"))

    ' --- 変更されたセル(Target)が、チェック対象セルと重なっているか判定 ---
    ' ※これを入れないと、シート上のどのセルを変更しても処理が走ってしまう
    If Intersect(Target, checkCell) Is Nothing Then Exit Sub

    ' --- 実際の入力チェック処理 ---
    Application.EnableEvents = False
    Select Case checkCell.value
        ' 条件1: セルが空っぽ（空文字列）の場合
        Case ""
            MsgBox "この項目は、入力必須です。デフォルト値に戻します。", vbExclamation, "入力エラー"
            Me.Range(Me.UseRangeName(2, "Me.Worksheet_Change")).value = DefaultValue

        ' 条件2: 特定の禁止文字列が入力された場合
        Case "User Data"
            MsgBox "セキュリティ対策のため、標準のデータディレクトリの指定はできません。" & vbCrLf & "詳細は下記をご覧ください。" & vbCrLf & "https://developer.chrome.com/blog/remote-debugging-port?hl=ja" & vbCrLf & vbCrLf & "デフォルト値に戻します。", vbCritical, "入力エラー"
            Me.Range(Me.UseRangeName(2, "Me.Worksheet_Change")).value = DefaultValue
    
        ' 条件3: 排他設定
    
    End Select
    Application.EnableEvents = True
End Sub
