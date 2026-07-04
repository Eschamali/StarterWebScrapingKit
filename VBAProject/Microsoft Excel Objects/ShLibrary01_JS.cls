VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShLibrary01_JS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'***************************************************************************************************
'      ワークシートから、JavaScriptを簡単に取得出来るようにするヘルパーメソッドを用意してます
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                              ■■■ テーブルの列操作用 ■■■
'***************************************************************************************************
Private Const TableName01 As String = "chromium_bidi"



'***************************************************************************************************
'                             ■■■ WebDriverBiDi関連 ■■■
'***************************************************************************************************
'* 機能　　：Chromiumタブ内で動作するBiDiとCDP間を変換するJavaScriptの取得/更新を行います
'---------------------------------------------------------------------------------------------------
'* GET返値 ：テーブルから結合された`mapperTab.js`
'* LET引数 ：更新させたい`mapperTab.js`　※挙動としては、DELETE-INSERTです
'---------------------------------------------------------------------------------------------------
'* 詳細情報：https://github.com/GoogleChromeLabs/chromium-bidi
'* 注意事項：Excelの制限(https://support.microsoft.com/ja-jp/excel/excel-specifications-and-limits)の都合上、32,767文字ごとに分割挿入、結合を行っております
'***************************************************************************************************
Property Get TableMapperTabJS() As String
    '1. テーブルから配列を取得
    With Me.ListObjects(TableName01)
        'データがない場合は、エラーで止める
        If .ListRows.Count = 0 Then
            Err.Raise vbObjectError, 3 + vbObjectError, "`mapperTab.js`がセットされていません"
            Exit Property
        End If

        '一旦、2次元配列で受け取る
        Dim tmp2, chunkNum As Long
        tmp2 = .DataBodyRange
        chunkNum = UBound(tmp2)
    End With

    '2. 1次元配列に変換
    ReDim tmp1(1 To chunkNum)
    Dim i As Long
    For i = 1 To chunkNum
        tmp1(i) = tmp2(i, 1)
    Next

    '3. 結合して返却
    TableMapperTabJS = Join(tmp1, "")
End Property
Property Let TableMapperTabJS(JavaScriptStr As String)
    '1セルごとの格納文字数上限
    Const LimitCellCharacter As Long = 32767


    '1. 分割数を確認
    Dim NumberOfSplit As Long: NumberOfSplit = Len(JavaScriptStr) / LimitCellCharacter   '小数になったら、切り捨て

    '2. 配列確保
    ReDim JsCodeChunk(NumberOfSplit, 0) '※0始まりなので、「+1」という補正は不要です

    '3. 分割して格納する
    Dim i As Long
    For i = 0 To NumberOfSplit
        JsCodeChunk(i, 0) = Mid(JavaScriptStr, LimitCellCharacter * i + 1, LimitCellCharacter)
    Next

    '4. `mapperTab.js`専用テーブル制御へ
    Application.ScreenUpdating = False      'テーブル追加時のアニメーション動作を止める
    With ShLibrary01_JS
        With .ListObjects(TableName01)

            '4-1. テーブルの中身があったら、全消去
            If .ListRows.Count Then .DataBodyRange.Delete

            '4-2. 挿入範囲の確定
            Dim TargetRange As Range
            Set TargetRange = .ListColumns(1).Range(2).Resize(NumberOfSplit + 1, 1)  '※こっちは、1始まりなので、「+1」という補正が必要

            '4-3. 挿入前に書式を「文字列(@)」に設定して、自動整形を防止
            TargetRange.NumberFormatLocal = "@"

            '4-4. 挿入
            TargetRange.value = JsCodeChunk
        End With
    End With
    Application.ScreenUpdating = True       '元に戻す
End Property

'***************************************************************************************************
'* 機能　　：`mapperTab.js`置き場のExcelテーブルの中身を削除します
'***************************************************************************************************
Public Sub DeleteMapperTabJS()
    'テーブル削除時のアニメーション動作を止める
    Application.ScreenUpdating = False


    'JS置き場のワークシート制御へ
    With ShLibrary01_JS
        With .ListObjects(TableName01)
            'テーブルの中身を全消去
            If .ListRows.Count Then .DataBodyRange.Delete
        End With
    End With


    '元に戻す
    Application.ScreenUpdating = True
End Sub



'***************************************************************************************************
'                             ■■■ プロパティメソッド ■■■
'***************************************************************************************************
'* 機能　　：このシートのみ定義されてるセル名称を設定/取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：セルの中身(`.Value`のみ)
'* LET引数 ：更新させたい値
'* 引数　　：RangeID        セルID
'            ErrorSource    この関数を呼び出したソース。※デバック用
'***************************************************************************************************
Property Let UseRangeID(RangeID As Long, ErrorSource As String, SetValue As Variant)
    Me.Range(RangeMapping(RangeID, ErrorSource)).value = SetValue
End Property
Property Get UseRangeID(RangeID As Long, ErrorSource As String) As Variant
    UseRangeID = Me.Range(RangeMapping(RangeID, ErrorSource)).value
End Property

'***************************************************************************************************
'* 機能　　：数値を基にセル名をマッピングします
'***************************************************************************************************
Private Function RangeMapping(RangeID As Long, ErrorSource) As String
    Select Case RangeID
        Case 1: RangeMapping = "Version_chromiumbidi"
        Case Else: Err.Raise 1 + vbObjectError, ErrorSource, "渡されたセルIDは、定義されてません"
    End Select
End Function
