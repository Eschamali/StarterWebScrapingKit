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
Private Const LibraryTableName As String = "LibraryJavaScript"  '便利なJavaScriptが詰まってるテーブル名
Private Const mapperTabJSName  As String = "chromium_bidi"



'***************************************************************************************************
'                             ■■■ WebDriverBiDi関連 ■■■
'***************************************************************************************************
'* 機能　　：Chromiumタブ内で動作するBiDiとCDP間を変換するJavaScriptの取得/更新を行います
'---------------------------------------------------------------------------------------------------
'* GET返値 ：テーブルから結合された`mapperTab.js`
'* LET引数 ：更新させたい`mapperTab.js`
'---------------------------------------------------------------------------------------------------
'* 詳細情報：https://github.com/GoogleChromeLabs/chromium-bidi
'* 注意事項：`mapperTab.js`取得/更新専用メソッドです
'***************************************************************************************************
Property Get TableMapperTabJS() As String
    '1. DLookupで、指定したツールのJavaScriptStrを「一発」で取得
    '※長いテキスト型なので、何十万文字あろうが一瞬でそのまま丸ごとVariantで返ってきます
    Dim tmp
    tmp = DLookup("JavaScriptStr", LibraryTableName, "ToolName = '" & mapperTabJSName & "'")

    ' 2. データが存在しない（または空っぽ）の場合は、エラーを投げる（Excel版の0件時エラーを忠実に再現）
    If IsNull(tmp) Or LenB(tmp) = 0 Then
        Err.Raise vbObjectError, 3 + vbObjectError, "`mapperTab.js`がセットされていません"
        Exit Property
    End If

    ' 3. 分割なんてしてないので、結合(Join)も不要！ そのまま文字列に変換して終了
    TableMapperTabJS = CStr(tmp)
End Property
Property Let TableMapperTabJS(JavaScriptStr As String)
    'DB操作用オブジェクトを用意
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim sql As String


    ' 1. 「chromium_bidi」というレコードがすでにテーブルにあるか、SQLで1行だけピンポイント取得
    sql = "SELECT * FROM " & LibraryTableName & " WHERE ToolName = '" & mapperTabJSName & "'"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    ' 2. 存在チェック＆データ挿入方式の判定（UPSERT）
    If rs.EOF Then
        ' 存在しない：新規追加（INSERT）
        rs.AddNew
        rs.Fields("ToolName").value = mapperTabJSName
    Else
        ' 存在する：既存編集（UPDATE）
        rs.Edit
    End If

    '3. 巨大なJavaScriptStrをそのままダイレクトに1行代入！
    rs.Fields("JavaScriptStr").value = JavaScriptStr

    ' 4. データベースに書き込み確定（コミット）
    rs.Update
End Property

'***************************************************************************************************
'* 機能　　：`mapperTab.js`置き場のAccessテーブルの中身を削除します
'***************************************************************************************************
Public Sub DeleteMapperTabJS()
    'DB操作用オブジェクトを用意
    Dim db As DAO.Database: Set db = CurrentDb
    Dim sql As String


    ' 1. SQLで、指定されたToolNameのレコードをピンポイントで削除！
    sql = "DELETE * FROM " & LibraryTableName & " WHERE ToolName = '" & mapperTabJSName & "'"

    ' 2. 実行！（データがすでに消えていても、エラーを出さずに安全に終了します）
    db.Execute sql, dbFailOnError
End Sub

'***************************************************************************************************
'* 機能　　：`mapperTab.js`のバージョン取得/更新を行います
'---------------------------------------------------------------------------------------------------
'* GET返値 ：テーブルに記録中のバージョン
'* LET引数 ：更新させたいバージョン
'***************************************************************************************************
Property Get VersionMapperTabJS() As String
    '1. DLookupで、指定したツールのJavaScriptStrを「一発」で取得
    '※長いテキスト型なので、何十万文字あろうが一瞬でそのまま丸ごとVariantで返ってきます
    Dim tmp
    tmp = DLookup("Version", LibraryTableName, "ToolName = '" & mapperTabJSName & "'")

    ' 2. データが存在しない（または空っぽ）の場合は、`vbnullstring`で返却
    If IsNull(tmp) Or LenB(tmp) = 0 Then Exit Property

    ' 3. 分割なんてしてないので、結合(Join)も不要！ そのまま文字列に変換して終了
    VersionMapperTabJS = CStr(tmp)
End Property
Property Let VersionMapperTabJS(Version As String)
    'DB操作用オブジェクトを用意
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim sql As String


    ' 1. 「chromium_bidi」というレコードがすでにテーブルにあるか、SQLで1行だけピンポイント取得
    sql = "SELECT * FROM " & LibraryTableName & " WHERE ToolName = '" & mapperTabJSName & "'"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    ' 2. 存在チェック＆データ挿入方式の判定（UPSERT）
    If rs.EOF Then
        ' 存在しない：新規追加（INSERT）
        rs.AddNew
        rs.Fields("ToolName").value = mapperTabJSName
    Else
        ' 存在する：既存編集（UPDATE）
        rs.Edit
    End If

    '3. バージョン更新
    rs.Fields("Version").value = Version

    ' 4. データベースに書き込み確定（コミット）
    rs.Update
End Property
