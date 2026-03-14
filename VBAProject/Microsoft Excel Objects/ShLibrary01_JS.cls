VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShLibrary01_JS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



'***************************************************************************************************
'                             ■■■ プロパティメソッド ■■■
'***************************************************************************************************
'* 機能　　：このシートのみ定義されてるセル名称を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：IDに紐づくセル名称
'* 引数　　：RangeID        セルID
'            ErrorSource    この関数を呼び出したソース
'***************************************************************************************************
Property Get UseRangeName(RangeID As Long, ErrorSource As String) As String
    Select Case RangeID
        Case 1: UseRangeName = "Version_chromiumbidi"
        Case Else: Err.Raise 1 + vbObjectError, ErrorSource, "渡されたセルIDは、定義されてません"
    End Select
End Property

'***************************************************************************************************
'* 機能　　：このシートにあるテーブル名を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：IDに紐づくテーブル名称
'* 引数　　：TableID        テーブルID
'            ErrorSource    この関数を呼び出したソース
'***************************************************************************************************
Property Get UseTableName(TableID As Long, ErrorSource As String) As String
    Select Case TableID
        Case 1: UseTableName = "chromium_bidi"
        Case Else: Err.Raise 2 + vbObjectError, ErrorSource, "渡されたテーブルIDは、定義されてません"
    End Select
End Property
