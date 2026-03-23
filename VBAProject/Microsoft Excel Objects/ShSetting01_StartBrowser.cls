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
'                             ■■■ テーブルの列操作用 ■■■
'***************************************************************************************************
Public Enum Table01ColumnID
    UserName = 1
    PipeName
    HandleValue
End Enum

Public Enum Table02ColumnID
    UserName = 1
    DataName
    DataValue
End Enum

Public Enum Table03ColumnID
    ProtocolName = 1
    階層
    接頭辞
End Enum



'***************************************************************************************************
'                             ■■■ テーブルの行操作用 ■■■
'***************************************************************************************************
Public Enum Table02RowID
    lngLastID = 1
    sessionID
    targetID
End Enum



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
        Case 1: UseRangeName = "ブラウザexeパス"
        Case 2: UseRangeName = "ユーザーデータフォルダ名"
        Case 3: UseRangeName = "追加の起動引数一式"
        Case 4: UseRangeName = "Chromeで使用"
        Case 5: UseRangeName = "LogLevel"
        Case 6: UseRangeName = "常にブラウザを新規で起動"
        Case 7: UseRangeName = "イミディエイトに表示"
        Case 8: UseRangeName = "ブラウザから受信した生のJSONをイミディエイトに表示"
        Case 9: UseRangeName = "ログファイル保存"
        Case 10: UseRangeName = "常にUTF8で送受信"
        Case 11: UseRangeName = "BiDi_CDP_Mapperタブを表示"
        Case 12: UseRangeName = "起動モード"
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
        Case 1: UseTableName = "PreviousHandleInfo"
        Case 2: UseTableName = "PreviousBrowserInfo"
        Case 3: UseTableName = "ProtocolWhiteList"
        Case Else: Err.Raise 2 + vbObjectError, ErrorSource, "渡されたテーブルIDは、定義されてません"
    End Select
End Property

'***************************************************************************************************
'* 機能　　：指定テーブルIDへデータ追加/上書きを一括で行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：TableID        テーブルID
'            ErrorSource    この関数を呼び出したソース
'            SetArray       追加するデータ配列
'---------------------------------------------------------------------------------------------------
'* 機能説明：`lngLastID`,`sessionID`,`targetID`といった順番で、ブラウザセッションテーブルに一括記録します
'* 注意事項：・`SetArray`は、2次元配列にする必要があります
'            ・同名ユーザーの場合は、上書きします
'***************************************************************************************************
Property Let serializeToTable(TableID As Long, ErrorSource As String, SetArray)
    'テーブル追加時のアニメーション動作を止める
    Application.ScreenUpdating = False


    '配列じゃないのはエラーにします
    If Not (IsArray(SetArray)) Then Err.Raise 3 + vbObjectError, ErrorSource, "データ挿入には、配列形式が必要です"

    '入力データ数を取得
    On Error Resume Next
        Dim ErrorExists As Long
        Dim InsertDataColumns   As Long: InsertDataColumns = UBound(SetArray, 2)    '列数
        ErrorExists = Err.Number
    On Error GoTo 0

    '2次元配列じゃないのはエラーにします
    If ErrorExists Then Err.Raise ErrorExists, ErrorSource, "データ挿入には、2次元配列形式が必要です"

    Dim InsertDataRows          As Long: InsertDataRows = UBound(SetArray, 1)       '行数

    '操作するテーブルオブジェクトの基準セルを取得
    Dim TableName As String: TableName = UseTableName(TableID, "ShSetting01_StartBrowser.serializeToTable." & ErrorSource)
    Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName, CStr(SetArray(1, 1)), "ShSetting01_StartBrowser.serializeToTable." & ErrorSource)

    'データ挿入方式判定
    If StartPos Is Nothing Then
        'テーブルオブジェクトそのものを取得
        Dim TargetTableObj As ListObject: Set TargetTableObj = Me.ListObjects(TableName)

        'テーブルに格納してるデータ数を取得
        Dim DataRows As Long: DataRows = TargetTableObj.ListRows.Count

        'テーブル自体が確保してる行数を取得
        Dim TableRows As Long: TableRows = TargetTableObj.ListColumns(1).Range.Count

        '存在しないので、テーブル末尾に追加します。
        '※右記のテクニックを参考にした追加方式です→http://officetanaka.net/excel/vba/table/08.htm
        TargetTableObj.ListColumns(1).Range(TableRows + 1 + (DataRows = 0)).Resize(InsertDataRows, InsertDataColumns) = SetArray
    Else
        '存在するので、先頭セルを基準にリサイズして上書き
        StartPos.Resize(InsertDataRows, InsertDataColumns) = SetArray
    End If


    '元に戻す
    Application.ScreenUpdating = True
End Property

'***************************************************************************************************
'* 機能　　：ブラウザセッション管理テーブルに、データ上書きを単一で行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       対象ユーザー名
'            ReplaceID      置き換えるデータID
'            ErrorSource    この関数を呼び出したソース
'            SetValue       上書きするデータ
'---------------------------------------------------------------------------------------------------
'* 機能説明：UserNameに紐づく、`lngLastID`,`sessionID`,`targetID`のいずれかを指定して、ブラウザセッションテーブルを更新します
'* 注意事項：・新規追加はできません。既存データのみ有効です
'            ・ハンドル情報管理テーブルは非対応です
'***************************************************************************************************
Property Let serializeSingleToHandleInfoTable(UserName As String, ReplaceID As Table02RowID, ErrorSource As String, SetValue As String)
    '操作するテーブルオブジェクトの基準セルを取得
    Dim TableName As String: TableName = UseTableName(2, "ShSetting01_StartBrowser.serializeSingleToHandleInfoTable." & ErrorSource)
    Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName, UserName, "ShSetting01_StartBrowser.serializeSingleToHandleInfoTable." & ErrorSource)

    '存在したら、、先頭セルを基準にオフセットして上書き
    '※オフセットは、0始まりなので補正しておく
    If Not (StartPos Is Nothing) Then StartPos.Offset(ReplaceID - 1, Table02ColumnID.DataValue - 1) = SetValue
End Property

'***************************************************************************************************
'* 機能　　：指定テーブルIDから、対象のユーザ名に紐づくハンドル/セッション情報を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：UserName,TableIDに紐づくテーブル内の情報をDictionary型にしたもの
'* 引数　　：UserName       対象のユーザ名
'            TableID        テーブルID
'            ErrorSource    この関数を呼び出したソース
'---------------------------------------------------------------------------------------------------
'* 注意事項：記録されてない`UserName`を指定すると、`Nothing`が返ります
'***************************************************************************************************
Property Get deserializeFromTable(UserName As String, TableID As Long, ErrorSource As String) As Dictionary
    '操作するテーブルオブジェクトの基準セルを取得
    Dim GetTakeNum As Long
    Dim StartPos As Range: Set StartPos = GetReferenceRange(UseTableName(TableID, "ShSetting01_StartBrowser.serializeToTable." & ErrorSource), UserName, "ShSetting01_StartBrowser.serializeToTable." & ErrorSource, GetTakeNum)

    '基準セルがない場合は、即抜け
    If StartPos Is Nothing Then Exit Property

    '取得するテーブルに応じて、取得行数を変えてデータ取得
    Dim DataSet
    DataSet = StartPos.Offset(0, 1).Resize(GetTakeNum, 2)   '`UserName`情報は不要なので、1列分右へシフトする

    'データ名と値の紐づけを行い、Dictionaryで返却
    Dim i As Long
    Set deserializeFromTable = New Dictionary
    For i = 1 To GetTakeNum
        deserializeFromTable.Add DataSet(i, 1), DataSet(i, 2)
    Next
End Property

'***************************************************************************************************
'* 機能　　：指定テーブルIDから、対象のユーザ名に紐づくハンドル/セッション情報を消去します
'---------------------------------------------------------------------------------------------------
'* 返り値　：UserName,IDに紐づくテーブル内の情報をDictionary型にしたもの
'* 引数　　：UserName       対象のユーザ名
'            TableID        テーブルID
'            ErrorSource    この関数を呼び出したソース
'---------------------------------------------------------------------------------------------------
'* 注意事項：・記録されてない`UserName`を指定しても問題ありません。この場合、テーブルの操作は発生しません。
'            ・`UserName`を`vbnullstring`で指定すると、全データ消去として動作します。
'***************************************************************************************************
Property Let DeleteCDPToTable(TableID As Long, ErrorSource As String, UserName As String)
    'テーブル削除時のアニメーション動作を止める
    Application.ScreenUpdating = False


    '操作対象のテーブル名を取得
    Dim TableName As String: TableName = UseTableName(TableID, "ShSetting01_StartBrowser.serializeToTable." & ErrorSource)

    '文字列が格納されているか？
    If StrPtr(UserName) Then
        '`UserName`をキーに、操作するテーブルオブジェクトの基準セルを取得
        Dim GetTakeNum As Long
        Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName, UserName, "ShSetting01_StartBrowser.serializeToTable." & ErrorSource, GetTakeNum)
        
        '基準セルがない場合は、即抜け
        If StartPos Is Nothing Then Exit Property

        '対象のテーブルに応じて、消去行数をリサイズ
        Dim DataSet As Range
        Set DataSet = StartPos.Resize(GetTakeNum, 3)

        '消去
        DataSet.Delete
    Else
        '1件以上あったら、全消去します
        With Me.ListObjects(TableName)
            If .ListRows.Count Then .DataBodyRange.Delete
        End With
    End If


    '元に戻す
    Application.ScreenUpdating = True
End Property



'***************************************************************************************************
'                           ■■■ イベントプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：入力規則チェックを行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：エラー時は、デフォルト値に書き換えます。前回の値には戻せません。
'***************************************************************************************************
Private Sub Worksheet_Change(ByVal Target As Range)
    ' --- チェック対象のセルとデフォルト値を用意 ---
    Const checkCell1_DefaultValue As String = "Automation Data"
    Dim checkCell1 As Range: Set checkCell1 = Me.Range(Me.UseRangeName(2, "ShSetting01_StartBrowser.Worksheet_Change"))


    ' --- 変更されたセル(Target)が、チェック対象セルと重なっているか判定 ---
    Application.EnableEvents = False
    Select Case True
        '1. ユーザーデータディレクトリ の設定チェック
        Case Not (Intersect(Target, checkCell1) Is Nothing)
            '値を確認します
            Select Case checkCell1.value
                ' 条件1: セルが空っぽ（空文字列）の場合
                Case ""
                    MsgBox "この項目は、入力必須です。デフォルト値に戻します。", vbExclamation, "入力エラー"
                    checkCell1.value = checkCell1_DefaultValue

                ' 条件2: 特定の禁止文字列が入力された場合
                Case "User Data"
                    MsgBox "セキュリティ対策のため、標準のデータディレクトリの指定はできません。" & vbCrLf & "詳細は下記をご覧ください。" & vbCrLf & "https://developer.chrome.com/blog/remote-debugging-port?hl=ja" & vbCrLf & vbCrLf & "デフォルト値に戻します。", vbCritical, "入力エラー"
                    checkCell1.value = checkCell1_DefaultValue

            End Select

        '他のセルをチェックする場合に追加する...
        Case Else

    End Select
    Application.EnableEvents = True
End Sub



'***************************************************************************************************
'                           ■■■ ヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：テーブル操作の基準セルを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：基準セルオブジェクト
'* 引数　　：TableName      このワークシート上のテーブル名
'            SearchUser     検索ユーザー名
'            ErrorSource    この関数を呼び出したソース
'            ResizeRows     有効なデータ行数範囲(※参照渡しによる実質返り値)
'---------------------------------------------------------------------------------------------------
'* 注意事項：基準セルがない場合は、`Nothing`となります
'***************************************************************************************************
Private Function GetReferenceRange(TableName As String, SearchUser As String, ErrorSource As String, Optional ByRef ResizeRows As Long) As Range
    'このワークシート上にあるテーブルオブジェクトを取得
    Dim TargetTableObj As ListObject: Set TargetTableObj = Me.ListObjects(TableName)

    'データ操作の基準となるセルを取得する
    With TargetTableObj.DataBodyRange
        On Error Resume Next
            Select Case TableName
                Case UseTableName(1, ErrorSource)
                    Set GetReferenceRange = WorksheetFunction.XLookup(SearchUser, .Columns(Table01ColumnID.UserName), .Columns(Table01ColumnID.UserName))   '第四引数に`Nothing`を入れればいい感じになると思いきやダメみたいです
                    ResizeRows = 6  '1ユーザーにつき、5つ分のパイプハンドル情報 + プロセスID

                Case UseTableName(2, ErrorSource)
                    Set GetReferenceRange = WorksheetFunction.XLookup(SearchUser, .Columns(Table02ColumnID.UserName), .Columns(Table02ColumnID.UserName))   '第四引数に`Nothing`を入れればいい感じになると思いきやダメみたいです
                    ResizeRows = 3  '1ユーザーにつき、3つ分のブラウザセッション情報

                Case Else
                    On Error GoTo 0
                    Err.Raise 4 + vbObjectError, ErrorSource, "渡されたテーブル名：" & TableName & "　は、このプロパティでは使用できません"
            End Select
        On Error GoTo 0
    End With
End Function
