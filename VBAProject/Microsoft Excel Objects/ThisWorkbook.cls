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



'***************************************************************************************************
'                           ■■■ ここしか使わないWindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function GetHandleInformation Lib "kernel32" (ByVal hObject As LongPtr, lpdwFlags As Long) As Long



'***************************************************************************************************
'                                 ■■■ イベントプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：マクロブックOpen時、ハンドル情報をチェックし、無効ならクリアしておきます
'---------------------------------------------------------------------------------------------------
'* 機能説明：ExcelとCDP連携してる際のハンドル情報は、そのExcelプロセスを完全に落とすと、持ってるハンドル情報は無意味になります。
'            その無意味なハンドル情報を使うとエラーになるため、その対策をします。
'            「Excel自体は落ちてない（プロセスAは生きてる）けど、ブックだけ閉じて開き直した」 パターンも考慮します
'***************************************************************************************************
Private Sub Workbook_Open()
    '1つでもFalseなら、クリアして処理終了
    With Sh99_Setting_StartBrowser
        With .ListObjects(.UseTableName(1, "ThisWorkbook.Workbook_Open")).DataBodyRange
            If Not (IsHandleValid(WorksheetFunction.XLookup("hStdOutRd", .Columns(1), .Columns(2), 0))) Then .Columns(2).Clear: Exit Sub
            If Not (IsHandleValid(WorksheetFunction.XLookup("hStdInWr", .Columns(1), .Columns(2), 0))) Then .Columns(2).Clear: Exit Sub
            If Not (IsHandleValid(WorksheetFunction.XLookup("hCDPOutRd", .Columns(1), .Columns(2), 0))) Then .Columns(2).Clear: Exit Sub
            If Not (IsHandleValid(WorksheetFunction.XLookup("hCDPInWr", .Columns(1), .Columns(2), 0))) Then .Columns(2).Clear: Exit Sub
        End With
    End With

    Debug.Print "デバッグハンドル情報チェック：OK!"
End Sub



'***************************************************************************************************
'                                 ■■■ ヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：自動ブラウザハンドルが既に実行中かどうかを判断します
'---------------------------------------------------------------------------------------------------
'* 返り値　：True       成功すれば非ゼロ
'            False      失敗（無効なハンドル）なら0が返る
'* 引数　　：handleVal  ハンドル番号
'***************************************************************************************************
Private Function IsHandleValid(ByVal handleVal As LongPtr) As Boolean
    ' ハンドルの情報を取得してみる
    Dim dwFlags As Long, result As Long
    result = GetHandleInformation(handleVal, dwFlags)

    '論理値に変換
    ' 成功すれば非ゼロ、失敗（無効なハンドル）なら0が返る
    If result Then
        IsHandleValid = True
    Else
        IsHandleValid = False
        Debug.Print "無効なハンドル情報のため、前回のデバッグハンドル情報を消去しました。"
    End If
End Function
