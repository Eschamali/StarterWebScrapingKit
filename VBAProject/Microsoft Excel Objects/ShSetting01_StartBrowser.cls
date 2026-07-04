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
'               パイプハンドルのセッション、ブラウザセッション管理を担います
'                   これにより、リアタッチ機能が使えるようになります
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                              ■■■ テーブルの列操作用 ■■■
'***************************************************************************************************
Private Const TableName01 As String = "BrowserHandleInfo"

Private Const TableName02 As String = "BrowserContextInfo"
Public Enum Table02ColumnID
    UserName = 1
    DataName
    DataValue
End Enum

Private Const TableName03 As String = "ProtocolWhiteList"
Public Enum Table03ColumnID
    ProtocolName = 1
    階層
    接頭辞
End Enum



'***************************************************************************************************
'                          ■■■ 各ユーザごとの記録データ数設定 ■■■
'***************************************************************************************************
Private Const BrowserHandleNum  As Long = 7 '1ユーザーにつき、6つ分のハンドル情報       + プロセスID
Private Const BrowserContextNum As Long = 3 '1ユーザーにつき、2つ分のブラウザセッション + BiDi-context情報



'***************************************************************************************************
'                             ■■■ 数値からセル名取得 ■■■
'***************************************************************************************************
'* 機能　　：数値で特定のセルの中身を取得します
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
        Case 7: UseRangeName = "イミディエイトに表示"
        Case 8: UseRangeName = "ブラウザから受信した生のJSONをイミディエイトに表示"
        Case 9: UseRangeName = "ログファイル保存"
        Case 10: UseRangeName = "常にUTF8で送受信"
        Case 11: UseRangeName = "BiDi_CDP_Mapperタブを表示"
        Case 12: UseRangeName = "起動モード"
        Case 13: UseRangeName = "BiDi_CDP_Mapperの診断ログ表示"
        Case Else: Err.Raise 1 + vbObjectError, ErrorSource, "渡されたセルIDは、定義されてません"
    End Select
End Property



'***************************************************************************************************
'                       ■■■ 各テーブルへの一括読み書き - INSERT/SELECT ■■■
'***************************************************************************************************
'* 機能　　：各種ハンドル情報をテーブルに記録/読み取りを行います
'---------------------------------------------------------------------------------------------------
'* GET返値 ：`UserName`に紐づくExcelテーブルに載ってる各種パイプ情報を記録したDictionaryオブジェクト
'* LET引数 ：各種パイプ情報を記録したDictionaryオブジェクト
'               ハンドル名 - ハンドル値　というフォーマットとします
'
'* 引数　　：UserName           ハンドル情報に紐づくユーザ名。
'            ErrorSource        この関数を呼び出したソース。※デバック用
'***************************************************************************************************
Property Set TableBrowserHandle(UserName As String, ErrorSource As String, SetDictionary As Scripting.Dictionary)
    '-------------- 1.エラーチェック --------------
    '1-1. `nothing`時は、エラーにします
    If SetDictionary Is Nothing Then Err.Raise 1 + vbObjectError, ErrorSource, "入力データがありません"

    '1-2. データ個数をチェック
    If Not SetDictionary.Count = BrowserHandleNum Then Err.Raise 2 + vbObjectError, ErrorSource, "入力データサイズが合いません"
    '----------------------------------------------

    '2. 記録ヘルパーメソッドを呼び出す
    serializeToTable TableName01, UserName, ErrorSource, SetDictionary
End Property
Property Get TableBrowserHandle(UserName As String, ErrorSource As String) As Scripting.Dictionary
    '1. 取り出しヘルパーメソッドを呼び出す
    Set TableBrowserHandle = deserializeFromTable(TableName01, UserName, ErrorSource, BrowserHandleNum)
End Property

'***************************************************************************************************
'* 機能　　：各種ブラウザコンテキスト(タブ)情報をテーブルに記録/読み取りを行います
'---------------------------------------------------------------------------------------------------
'* GET返値 ：`UserName`に紐づくExcelテーブルに載ってる各種コンテキスト情報を記録したDictionaryオブジェクト
'* LET引数 ：各種コンテキスト情報を記録したDictionaryオブジェクト
'               コンテキスト名 - セッション値　というフォーマットとします
'
'* 引数　　：UserName           コンテキスト情報に紐づくユーザ名。
'            ErrorSource        この関数を呼び出したソース。※デバック用
'***************************************************************************************************
Property Set TableBrowserContext(UserName As String, ErrorSource As String, SetDictionary As Scripting.Dictionary)
    '-------------- 1.エラーチェック --------------
    '1-1. `nothing`時は、エラーにします
    If SetDictionary Is Nothing Then Err.Raise 3 + vbObjectError, ErrorSource, "入力データがありません"

    '1-2. データ個数をチェック
    If Not SetDictionary.Count = BrowserContextNum Then Err.Raise 3 + vbObjectError, ErrorSource, "入力データサイズが合いません"
    '----------------------------------------------

    '2. 記録ヘルパーメソッドを呼び出す
    serializeToTable TableName02, UserName, ErrorSource, SetDictionary
End Property
Property Get TableBrowserContext(UserName As String, ErrorSource As String) As Scripting.Dictionary
    '1. 取り出しヘルパーメソッドを呼び出す
    Set TableBrowserContext = deserializeFromTable(TableName02, UserName, ErrorSource, BrowserContextNum)
End Property

'***************************************************************************************************
'* 機能　　：ホワイトプロトコルリスト設定情報をテーブルから読み取ります
'---------------------------------------------------------------------------------------------------
'* GET返値 ：`UserName`に紐づくExcelテーブルに載ってる各種コンテキスト情報を記録した2次元配列
'---------------------------------------------------------------------------------------------------
'* 注意事項：列単位取得の場合、仕様上、二次元配列で返ります。変換のオーバーヘッドを防ぐためそのまま返却します
'***************************************************************************************************
Property Get TableProtocolWhiteList()
    '1. 指定のデータテーブル、列配列を取得
    Dim WhiteList: WhiteList = Me.ListObjects(TableName03).DataBodyRange.Columns(Table03ColumnID.接頭辞)

    '2. 配列か？
    If IsArray(WhiteList) Then
        'そのまんま返却
        TableProtocolWhiteList = WhiteList
    Else
        '配列じゃない場合は1×1二次元配列にしておく
        ReDim tmp(1 To 1, 1 To 1)
        tmp(1, 1) = WhiteList

        TableProtocolWhiteList = tmp
    End If
End Property



'***************************************************************************************************
'                           ■■■ 各テーブルへの上書き - UPDATE ■■■
'***************************************************************************************************
'* 機能　　：ブラウザコンテキスト(タブ)情報のテーブルに、データ上書きを単一で行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName           対象ユーザー名
'            ReplaceDataName    置き換えるデータ名
'            ErrorSource        この関数を呼び出したソース
'            SetValue           上書きするデータ
'---------------------------------------------------------------------------------------------------
'* 機能説明：UserNameに紐づく、`BiDi-context`,`sessionID`,`targetID`のいずれかを指定して、ブラウザセッションテーブルを更新します
'* 注意事項：・新規追加はできません。既存データのみ有効です
'            ・ハンドル情報管理テーブルは非対応です
'***************************************************************************************************
Property Let UpdateBrowserContext(UserName As String, ReplaceName As String, ErrorSource As String, SetValue As String)
    '1. 操作するテーブルオブジェクトの基準セルを取得
    Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName02, UserName)

    '2. 存在したら、対象ユーザー先頭セルを基準にデータ名を探す
    If Not StartPos Is Nothing Then
        '2-1. match関数で書き換えデータ名位置を特定
        Dim RawPos As Long
        RawPos = WorksheetFunction.Match(ReplaceName, StartPos.Offset(0, Table02ColumnID.DataName - 1).Resize(BrowserContextNum, 1), 0)   '※オフセットは、0始まりなので補正しておく

        '2-2. 実際のデータ値を書き換え
        StartPos.Offset(RawPos - 1, Table02ColumnID.DataValue - 1).value = SetValue       '※オフセットは、0始まりなので補正しておく
    End If
End Property



'***************************************************************************************************
'                           ■■■ テーブルから削除 - DELETE ■■■
'***************************************************************************************************
'* 機能　　：対象のユーザ名に紐づく各種ハンドル情報を消去します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       対象のユーザ名
'            ErrorSource    この関数を呼び出したソース
'---------------------------------------------------------------------------------------------------
'* 注意事項：・記録されてない`UserName`を指定しても問題ありません。この場合、テーブルの操作は発生しません
'            ・`UserName`を`vbnullstring`で指定すると、全データ消去として動作します
'            ・handle解放は別途行う必要があります
'***************************************************************************************************
Property Let DeleteBrowserHandle(ErrorSource As String, UserName As String)
    '1. 削除ヘルパーメソッドを呼び出す
    DeleteTableData TableName01, UserName, ErrorSource, BrowserHandleNum
End Property

'***************************************************************************************************
'* 機能　　：対象のユーザ名に紐づく各種ブラウザコンテキスト(タブ)情報を消去します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       対象のユーザ名
'            ErrorSource    この関数を呼び出したソース
'---------------------------------------------------------------------------------------------------
'* 注意事項：・記録されてない`UserName`を指定しても問題ありません。この場合、テーブルの操作は発生しません
'            ・`UserName`を`vbnullstring`で指定すると、全データ消去として動作します
'            ・session解放は別途行う必要があります
'***************************************************************************************************
Property Let DeleteBrowserContext(ErrorSource As String, UserName As String)
    '1. 削除ヘルパーメソッドを呼び出す
    DeleteTableData TableName02, UserName, ErrorSource, BrowserContextNum
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
'                         ■■■ テーブル操作ヘルパーメソッド ■■■
'***************************************************************************************************
'* 機能　　：指定テーブル名へデータ追加/上書きを一括で行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：TableName      追加/上書き先テーブル名
'            UserName       識別名称
'            ErrorSource    この関数を呼び出したソース　※デバック用
'            InputData      追加するデータ(Dictionary形式)
'---------------------------------------------------------------------------------------------------
'* 機能説明：入力されたDictionary、ユーザ名を基に、指定のテーブルへ一括挿入 or 上書きを行います
'* 注意事項：同名ユーザーの場合は、上書きします
'***************************************************************************************************
Private Sub serializeToTable(TableName As String, UserName As String, ErrorSource As String, InputData As Dictionary)
    'テーブル追加時のアニメーション動作を止める
    Application.ScreenUpdating = False


    '1. Dictionary の`.Count`から、2次元配列の枠を用意
    ReDim SetArray(InputData.Count - 1, 3 - 1)  '※0始まりなので補正

    '2. データ行数と列数を取得
    Dim InsertDataColumns   As Long: InsertDataColumns = UBound(SetArray, 2) + 1    '列数 ※0始まりなので補正
    Dim InsertDataRows      As Long: InsertDataRows = UBound(SetArray, 1) + 1       '行数 ※0始まりなので補正

    '3. Dictionary から、2次元配列へ変換
    Dim i As Long, PipeKeys As Variant
    PipeKeys = InputData.keys
    For i = LBound(SetArray) To UBound(SetArray)
        SetArray(i, 0) = UserName                   '識別名称
        SetArray(i, 1) = PipeKeys(i)                'データ名 or ハンドル名
        SetArray(i, 2) = InputData(PipeKeys(i)) '実際の値
    Next

    '4. 操作するテーブルオブジェクトの基準セルを取得
    Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName, UserName)

    '5. データ挿入方式判定
    If StartPos Is Nothing Then
        'テーブルオブジェクトそのものを取得
        Dim TargetTableObj As ListObject: Set TargetTableObj = Me.ListObjects(TableName)

        'テーブルに格納してるデータ数を取得
        Dim DataRows As Long: DataRows = TargetTableObj.ListRows.Count

        'テーブル自体が確保してる行数を取得
        Dim TableRows As Long: TableRows = TargetTableObj.ListColumns(1).Range.Count

        '存在しないので、テーブル末尾に追加します。
        '※右記のテクニックを参考にした追加方式です → http://officetanaka.net/excel/vba/table/08.htm
        TargetTableObj.ListColumns(1).Range(TableRows + 1 + (DataRows = 0)).Resize(InsertDataRows, InsertDataColumns) = SetArray
    Else
        '存在するので、先頭セルを基準にリサイズして上書き
        StartPos.Resize(InsertDataRows, InsertDataColumns) = SetArray
    End If


    '元に戻す
    Application.ScreenUpdating = True
End Sub

'***************************************************************************************************
'* 機能　　：指定テーブル名から、対象のユーザ名に紐づくハンドル/セッション情報を一括取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：TableName      取り出し先テーブル名
'* 引数　　：UserName       識別名称
'            ErrorSource    この関数を呼び出したソース　※デバック用
'---------------------------------------------------------------------------------------------------
'* 注意事項：記録されてない`UserName`を指定すると、`Nothing`が返ります
'***************************************************************************************************
Private Function deserializeFromTable(TableName As String, UserName As String, ErrorSource As String, TakeRow As Long) As Dictionary
    '操作するテーブルオブジェクトの基準セルを取得
    Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName, UserName)

    '基準セルがない場合は、即抜け
    If StartPos Is Nothing Then Exit Function

    '取得するテーブルに応じて、取得行数を変えてデータ取得
    Dim DataSet
    DataSet = StartPos.Offset(0, 1).Resize(TakeRow, 2)  '`UserName`情報は不要なので、1列分右へシフト後、リサイズ

    'データ名と値の紐づけを行い、Dictionaryで返却
    Dim i As Long
    Set deserializeFromTable = New Dictionary
    For i = 1 To TakeRow
        deserializeFromTable.Add DataSet(i, 1), DataSet(i, 2)
    Next
End Function

'***************************************************************************************************
'* 機能　　：指定テーブル名から、対象のユーザ名に紐づくハンドル/セッション情報を消去します
'---------------------------------------------------------------------------------------------------
'* 引数　　：TableName      削除テーブル名
'            UserName       対象のユーザ名
'            ErrorSource    この関数を呼び出したソース
'---------------------------------------------------------------------------------------------------
'* 注意事項：・記録されてない`UserName`を指定しても問題ありません。この場合、テーブルの操作は発生しません
'            ・`UserName`を`vbnullstring`で指定すると、全データ消去として動作します
'            ・handle解放等は別途行う必要があります
'***************************************************************************************************
Private Sub DeleteTableData(TableName As String, UserName As String, ErrorSource As String, DeleteRow As Long)
    'テーブル削除時のアニメーション動作を止める
    Application.ScreenUpdating = False


    '文字列が格納されているか？
    If StrPtr(UserName) Then
        '`UserName`をキーに、操作するテーブルオブジェクトの基準セルを取得
        Dim GetTakeNum As Long
        Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName, UserName)

        '基準セルがない場合は、即抜け
        If StartPos Is Nothing Then Application.ScreenUpdating = True: Exit Sub

        '対象のテーブルに応じて、消去行数をリサイズ
        Dim DataSet As Range
        Set DataSet = StartPos.Resize(DeleteRow, 3)

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
End Sub

'***************************************************************************************************
'* 機能　　：テーブル操作の基準セルを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：基準セルオブジェクト
'* 引数　　：TableName      このワークシート上のテーブル名
'            SearchUser     検索ユーザー名
'---------------------------------------------------------------------------------------------------
'* 注意事項：・基準セルがない場合は、`Nothing`となります
'            ・存在しないテーブル名はエラーです
'***************************************************************************************************
Private Function GetReferenceRange(TableName As String, SearchUser As String) As Range
    '1. このワークシート上にあるテーブルオブジェクトを取得
    Dim TargetTableObj As ListObject: Set TargetTableObj = Me.ListObjects(TableName)

    '2. データがそもそもない場合は即抜け
    If TargetTableObj.DataBodyRange Is Nothing Then Exit Function

    'データ操作の基準となるセルを取得する
    Dim StartRowPos As Variant  'エラー値に備えて`Variant`にしておく
    With TargetTableObj.DataBodyRange
        '1. 基準となる位置を取得
        StartRowPos = Application.Match(SearchUser, .Columns(1), 0)  '`IsError`で判定させたいので、`Application`を使用。UserName判定なので、列番号は常に1

        '2. 見つけた位置を基に、RangeObjectとして返す
        If Not IsError(StartRowPos) Then Set GetReferenceRange = .Cells(StartRowPos, 1)
    End With
End Function
