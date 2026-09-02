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
'                   CDP/WebDriverBiDi経由で、ブラウザ操作する際の準備関連のシートです
'                   パイプハンドルのセッション、ブラウザセッション管理を担います
'                       これにより、リアタッチ機能が使えるようになります
'
'                       また、複雑な設定周りを簡略化させるためでもあります
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                              ■■■ テーブルの列操作用 ■■■
'***************************************************************************************************
Private Const TableName01 As String = "BrowserHandleInfo"
Private Const TableName02 As String = "BrowserContextInfo"
Private Const TableName03 As String = "ProtocolWhiteList"



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const Column_SettingUserName    As String = "UserDataDirName"
Private Const PrimaryKeyName            As String = "UserName"



'***************************************************************************************************
'                           ■■■ 簡易起動メソッド - CDP編 ■■■
'***************************************************************************************************
'* 機能　　：この設定シートから、パラメーターを読み込んで、CDPモードでタブ制御ができるところまで準備します
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - CDPContext
'* 引数　　：StartURL   ブラウザ起動時にアクセスしたいURL。指定しない場合は、空ページ(abount:blank)になります。
'                       未指定でも クラスメソッド：navigate で後から、URL遷移も可能です。
'
'            SwtchUser  マルチインスタンス用に別ユーザーを指定するときに使用します。省略時は設定シートに記載されてる名称になります
'---------------------------------------------------------------------------------------------------
'* 詳細説明：VBEによるハードコーディングではなく、設定シートから読み込む方式により、ユーザー側からも手軽に設定変更ができます
'* 注意事項：ここでの起動は、Pipe経路に限ります
'***************************************************************************************************
Friend Function StartCDPModeContext(Optional StartURL As String, Optional SwitchUser As String) As CDPContext
    '設定シートの各セルから設定値を取得し、適用
    With Me
        '起動ブラウザ種類の設定
        '※CDP－Json コマンドによる操作なので、Chromium系統であれば、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim BrowserName As BrowserList: BrowserName = IIf(.UseRangeID(4, "ShSetting01_StartBrowser.StartCDPModeContext"), BrowserList.RunChrome, BrowserList.RunEdge)

        '第2引数が省略ならシート側の設定を適用
        If StrPtr(SwitchUser) = 0 Then SwitchUser = .UseRangeID(2, "ShSetting01_StartBrowser.StartCDPModeContext")

        'ブラウザ起動
        Set StartCDPModeContext = New CDPContext
        StartCDPModeContext.StartAndConnectTab BrowserName, StartURL, SwitchUser, .UseRangeID(3, "ShSetting01_StartBrowser.StartCDPModeContext")
    End With
End Function

'***************************************************************************************************
'* 機能　　：この設定シートから、パラメーターを読み込んで、CDPモードでのパイプ通信確保まで準備します
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - CDPBrowser
'* 引数　　：基本は、前述と一緒なので省略。`WebSocketMode`でWebSocketモードに切り替え可能です。
'---------------------------------------------------------------------------------------------------
'* 注意事項：・このプロシージャだけではまだ、タブ制御するまでの準備が整ってません。
'              `StartCDPModeContext`ではうまくいかず、手動でタブ制御確立するためのものです。
'              この処理後は基本的には、`.getTab` or `.newTab`で、直ぐに確立できると思います。
'            ・Demoプロシージャ：`runTabsAsMany`が分かりやすいと思います
'***************************************************************************************************
Friend Function StartCDPMode(Optional StartURL As String, Optional SwitchUser As String, Optional WebSocketMode As Boolean) As CDPBrowser
    '設定シートの各セルから設定値を取得し、適用
    With Me
        '起動ブラウザ種類の設定
        '※CDP－Json コマンドによる操作なので、Chromium系統であれば、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim BrowserName As BrowserList: BrowserName = IIf(.UseRangeID(4, "ShSetting01_StartBrowser.StartCDPMode"), BrowserList.RunChrome, BrowserList.RunEdge)

        '第2引数が省略ならシート側の設定を適用
        If StrPtr(SwitchUser) = 0 Then SwitchUser = .UseRangeID(2, "ShSetting01_StartBrowser.StartCDPMode")

        'ブラウザ起動
        Set StartCDPMode = New CDPBrowser
        If WebSocketMode Then
            'WebSocket経路で起動
            Dim ReadyWebSocket As New CDPCoreViaWebSocket
            ReadyWebSocket.ConnectCDPWithLocalBrowser BrowserName, StartURL, SwitchUser, .UseRangeID(3, "ShSetting01_StartBrowser.StartCDPMode")
            Set StartCDPMode = ReadyWebSocket.StartCDPMode
        Else
            'Pipe経路で起動
            StartCDPMode.start BrowserName, StartURL, SwitchUser, .UseRangeID(3, "ShSetting01_StartBrowser.StartCDPMode")
        End If
    End With
End Function



'***************************************************************************************************
'                           ■■■ 簡易起動メソッド - WebDriverBiDi編 ■■■
'***************************************************************************************************
'* 機能　　：設定シートから、パラメーターを読み込んで、WebDriverBiDiモードでタブ制御ができるところまで準備します
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - WebDriverBiDiContext
'* 引数　　：StartURL                       ブラウザ起動時にアクセスしたいURL。指定しない場合は、空ページ(abount:blank)になります。
'            SwtchUser                      マルチインスタンス用に別ユーザーを指定するときに使用します。省略時は設定シートに記載されてる名称になります
'            sessionCapabilitiesRequest     `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'---------------------------------------------------------------------------------------------------
'* 詳細説明：VBEによるハードコーディングではなく、設定シートから読み込む方式により、ユーザー側からも手軽に設定変更ができます
'* 注意事項：ここでの起動は、Pipe経路に限ります
'***************************************************************************************************
Friend Function StartBiDiModeContext(Optional StartURL As String, Optional SwitchUser As String, Optional sessionCapabilitiesRequest As Dictionary) As WebDriverBiDiContext
    '設定シートの各セルから設定値を取得し、適用
    With Me
        '起動ブラウザ種類の設定
        '※BiDi-Json コマンドによる操作ですが、Chromium系統に特化した制御のため、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim BrowserName As BrowserList: BrowserName = IIf(.UseRangeID(4, "ShSetting01_StartBrowser.StartBiDiModeContext"), BrowserList.RunChrome, BrowserList.RunEdge)

        '第2引数が省略ならシート側の設定を適用
        If StrPtr(SwitchUser) = 0 Then SwitchUser = .UseRangeID(2, "ShSetting01_StartBrowser.StartBiDiModeContext")

        'ブラウザ起動
        Set StartBiDiModeContext = New WebDriverBiDiContext
        StartBiDiModeContext.StartBiDiModeAndConnectTab BrowserName, StartURL, SwitchUser, .UseRangeID(3, "ShSetting01_StartBrowser.StartBiDiModeContext"), sessionCapabilitiesRequest
    End With
End Function

'***************************************************************************************************
'* 機能　　：設定シートから、パラメーターを読み込んで、BiDiモードでブラウザを起動するヘルパープロシージャです
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - WebDriverBiDiMode
'* 引数　　：基本は、前述と一緒なので省略。`WebSocketMode`でWebSocketモードに切り替え可能です。
'---------------------------------------------------------------------------------------------------
'* 注意事項：このプロシージャだけではまだ、タブ制御するまでの準備が整ってません。
'             `StartBiDiModeContext`ではうまくいかず、手動でタブ制御確立するためのものです。
'             この処理後は基本的には、`.getTab` or `.newTab`で、直ぐに確立できると思います。
'***************************************************************************************************
Friend Function StartBiDiMode(Optional StartURL As String, Optional SwitchUser As String, Optional sessionCapabilitiesRequest As Dictionary, Optional WebSocketMode As Boolean) As WebDriverBiDiMode
    '設定シートの各セルから設定値を取得し、適用
    With Me
        '起動ブラウザ種類の設定
        '※BiDi-Json コマンドによる操作ですが、Chromium系統に特化した制御のため、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim BrowserName As BrowserList: BrowserName = IIf(.UseRangeID(4, "ShSetting01_StartBrowser.StartBiDiMode"), BrowserList.RunChrome, BrowserList.RunEdge)

        '第2引数が省略ならシート側の設定を適用
        If StrPtr(SwitchUser) = 0 Then SwitchUser = .UseRangeID(2, "ShSetting01_StartBrowser.StartBiDiMode")

        'ブラウザ起動
        Set StartBiDiMode = New WebDriverBiDiMode
        If WebSocketMode Then
            Dim ReadyWebSocket As New CDPCoreViaWebSocket
            ReadyWebSocket.ConnectCDPWithLocalBrowser BrowserName, StartURL, SwitchUser, .UseRangeID(3, "ShSetting01_StartBrowser.StartCDPMode")
            Set StartBiDiMode = ReadyWebSocket.StartBiDiMode(sessionCapabilitiesRequest)
        Else
            StartBiDiMode.StartBiDiMode BrowserName, StartURL, SwitchUser, .UseRangeID(3, "ShSetting01_StartBrowser.StartBiDiMode"), sessionCapabilitiesRequest
        End If
    End With
End Function



'***************************************************************************************************
'                             ■■■ 数値でDBの中身取得/反映 ■■■
'***************************************************************************************************
'* 機能　　：このシートのみ定義されてるセル名称を設定/取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：RangeID に紐づく設定値
'* 引数　　：RangeID        セルID
'            ErrorSource    この関数を呼び出したソース。※デバック用
'***************************************************************************************************
Property Get UseRangeID(RangeID As Long, ErrorSource As String) As Variant
    Const SettingNameTableName As String = "BrowserLaunchSettings"


    '1. デフォルトのUserNameを取得
    Dim targetUser As String: targetUser = CurrentUserName
    If LenB(targetUser) = 0 Then Err.Raise 4 + vbObjectError, ErrorSource, "デフォルトの`" & Column_SettingUserName & "`が指定されてません"

    '2. 列名を取得
    Dim colName As String
    colName = RangeMapping(RangeID, ErrorSource)

    '3. DLookupで生データを取得
    UseRangeID = DLookup(RangeMapping(RangeID, ErrorSource), SettingNameTableName, Column_SettingUserName & " = '" & targetUser & "'")

    '4. Nullの場合は、空文字列で代用
    If IsNull(UseRangeID) Then UseRangeID = vbNullString: Exit Property

    '5. 型情報を取得
    Dim colType As Integer
    colType = CurrentDb.TableDefs(SettingNameTableName).Fields(colName).Type

    '6. 列の型に応じて、VBAの最適なデータ型にキャストして返す！
    Select Case colType
        ' ----------------------------------------------------
        ' ■ Boolean（はい/いいえ型）
        ' ----------------------------------------------------
        Case dbBoolean ' 1
            UseRangeID = CBool(UseRangeID)

        ' ----------------------------------------------------
        ' ■ 数値型（整数系：Byte, Integer, Long, BigInt）
        ' ----------------------------------------------------
        Case dbByte, dbInteger, dbLong, dbBigInt ' 2, 3, 4, 16
            UseRangeID = CLng(UseRangeID) ' 整数（Long）に変換

        ' ----------------------------------------------------
        ' ■ 数値型（実数・通貨系：Single, Double, Decimal, Currency）
        ' ----------------------------------------------------
        Case dbSingle, dbDouble, dbDecimal, dbNumeric, dbCurrency ' 5, 6, 7, 19, 20
            UseRangeID = CDbl(UseRangeID) ' 実数（Double）に変換

        ' ----------------------------------------------------
        ' ■ 日付/時刻型
        ' ----------------------------------------------------
        Case dbDate ' 8
            UseRangeID = CDate(UseRangeID) ' 日付型（Date）に変換

        ' ----------------------------------------------------
        ' ■ 文字列型（テキスト型、メモ型、固定幅テキスト型）
        ' ----------------------------------------------------
        Case dbText, dbMemo, dbChar ' 10, 12, 18
            UseRangeID = CStr(UseRangeID) ' 文字列型（String）に変換

        ' ----------------------------------------------------
        ' ■ その他（バイナリ等、基本は安全のため生のまま返す）
        ' ----------------------------------------------------
        Case Else
            'UseRangeID = UseRangeID

    End Select
End Property

'***************************************************************************************************
'* 機能　　：数値を基にセル名をマッピングします
'***************************************************************************************************
Private Function RangeMapping(RangeID As Long, ErrorSource) As String
    Select Case RangeID
        Case 1:  RangeMapping = "BrowserPath"
        Case 2:  RangeMapping = "UserDataDirName"
        Case 3:  RangeMapping = "AddArguments"
        Case 4:  RangeMapping = "UseChrome"
        Case 5:  RangeMapping = "LogLevel"
        Case 7:  RangeMapping = "DisplayImmediate"
        Case 8:  RangeMapping = "ViewBrowserResponse"
        Case 9:  RangeMapping = "LogFileFolderPath"
        Case 10: RangeMapping = "AlwaysSendAndReceiveUTF8"
        Case 11: RangeMapping = "ViewTab_BiDiCDPMapper"
        Case 12: RangeMapping = "StartWindowMode"
        Case 13: RangeMapping = "ViewDebugLogForBiDiCDPMapper"
        Case Else: Err.Raise 1 + vbObjectError, ErrorSource, "渡されたセルIDは、定義されてません"
    End Select
End Function



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
Friend Property Set TableBrowserHandle(UserName As String, ErrorSource As String, SetDictionary As Dictionary)
    '-------------- 1.エラーチェック --------------
    '1-1. `nothing`時は、エラーにします
    If SetDictionary Is Nothing Then Err.Raise 1 + vbObjectError, ErrorSource, "入力データがありません"

    '1-2. データ個数をチェック
    Dim RequiredCount: RequiredCount = CurrentDb.TableDefs(TableName01).Fields.Count - 1 '※主キー分、減算
    If Not SetDictionary.Count = RequiredCount Then Err.Raise 2 + vbObjectError, ErrorSource, "入力データサイズが合いません"
    '----------------------------------------------

    '2. 記録ヘルパーメソッドを呼び出す
    serializeToTable TableName01, UserName, ErrorSource, SetDictionary
End Property
Friend Property Get TableBrowserHandle(UserName As String, ErrorSource As String) As Dictionary
    '1. 取り出しヘルパーメソッドを呼び出す
    Set TableBrowserHandle = deserializeFromTable(TableName01, UserName, ErrorSource)
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
Friend Property Set TableBrowserContext(UserName As String, ErrorSource As String, SetDictionary As Dictionary)
    '-------------- 1.エラーチェック --------------
    '1-1. `nothing`時は、エラーにします
    If SetDictionary Is Nothing Then Err.Raise 3 + vbObjectError, ErrorSource, "入力データがありません"

    '1-2. データ個数をチェック
    Dim RequiredCount: RequiredCount = CurrentDb.TableDefs(TableName02).Fields.Count - 1 '※主キー分、減算
    If Not SetDictionary.Count = RequiredCount Then Err.Raise 3 + vbObjectError, ErrorSource, "入力データサイズが合いません"
    '----------------------------------------------

    '2. 記録ヘルパーメソッドを呼び出す
    serializeToTable TableName02, UserName, ErrorSource, SetDictionary
End Property
Friend Property Get TableBrowserContext(UserName As String, ErrorSource As String) As Dictionary
    '1. 取り出しヘルパーメソッドを呼び出す
    Set TableBrowserContext = deserializeFromTable(TableName02, UserName, ErrorSource)
End Property

'***************************************************************************************************
'* 機能　　：ホワイトプロトコルリスト設定情報をテーブルから読み取ります
'---------------------------------------------------------------------------------------------------
'* GET返値 ：プロトコルリストを記録した1次元配列
'---------------------------------------------------------------------------------------------------
'* 注意事項：・1次元配列への変換コストが入るため、1度のみ呼び出すこと
'            ・最初の要素番号は、`1`始まりです
'            ・1件以上ある前提です
'***************************************************************************************************
Property Get TableProtocolWhiteList()
    'DB操作用オブジェクトを用意
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim rawArray As Variant


    '1. DCountで、テーブル全体のレコード数を動的に、100%正確に数え上げる
    Dim targetCount As Long
    targetCount = DCount("*", TableName03)

    '2. SQLで「CheckInstr」列（またはそれに相当するフィールド名）だけをピンポイントで取得
    Set rs = db.OpenRecordset("SELECT CheckInstr FROM " & TableName03, dbOpenSnapshot)

    '3. GetRowsで一瞬で二次元配列（0始まり・列と行が逆転した状態）として取得
    ' ※rawArrayは (0, 0 To 行数-1) という構造の2次元配列になります
    rawArray = rs.GetRows(targetCount)

    '4. データ数取得
    Dim rowCount As Long
    rowCount = UBound(rawArray, 2) + 1

    '5. 1次元配列に変換させる
    ReDim tmp(1 To rowCount)
    Dim i As Long
    For i = 1 To rowCount
        tmp(i) = rawArray(0, i - 1)
    Next

    '6. 返却
    TableProtocolWhiteList = tmp
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
'            ・ハンドル情報管理テーブルは対象外です
'***************************************************************************************************
Property Let UpdateBrowserContext(UserName As String, ReplaceName As String, ErrorSource As String, SetValue As String)
    'DB操作用オブジェクトを用意
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim sql As String


    ' 1. 対象ユーザーのレコードをSQLでピンポイントに1行取得
    sql = "SELECT * FROM " & TableName02 & " WHERE " & PrimaryKeyName & " = '" & UserName & "'"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    ' 2. レコードが存在したら更新
    If Not rs.EOF Then
        '引数で受け取った「ReplaceName（列名）」をそのまま指定して上書き
        rs.Edit
        rs.Fields(ReplaceName).value = SetValue
        rs.Update
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
    DeleteTableData TableName01, UserName, ErrorSource
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
    DeleteTableData TableName02, UserName, ErrorSource
End Property



'***************************************************************************************************
'                         ■■■ テーブル操作ヘルパーメソッド ■■■
'***************************************************************************************************
'* 機能　　：指定テーブル名へデータ追加/上書きを一括で行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：TableName      追加/上書き先テーブル名
'            UserName       識別名称(主キー)
'            ErrorSource    この関数を呼び出したソース　※デバック用
'            InputData      追加するデータ(Dictionary形式)
'---------------------------------------------------------------------------------------------------
'* 機能説明：入力されたDictionary、ユーザ名を基に、指定のテーブルへ一括挿入 or 上書きを行います
'* 注意事項：同名ユーザーの場合は、上書きします
'***************************************************************************************************
Private Sub serializeToTable(TableName As String, UserName As String, ErrorSource As String, InputData As Dictionary)
    'DB操作用オブジェクトを用意
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim sql As String


    ' 1. 対象のUserNameのレコードを、SQLでピンポイントに取得する
    ' （すでに登録されていれば1行取得でき、未登録なら空のレコードセットが返ります）
    sql = "SELECT * FROM " & TableName & " WHERE " & PrimaryKeyName & " = '" & UserName & "'"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    ' 2. データ挿入方式の判定（新規追加か、編集か）
    If rs.EOF Then
        ' レコードが存在しない? 新規追加モードを起動
        rs.AddNew
        rs.Fields(PrimaryKeyName).value = UserName ' 主キーをセット
    Else
        ' レコードが存在する? 編集（上書き）モードを起動
        rs.Edit
    End If

    ' 3. Dictionaryから、各フィールドへループで一括代入
    ' ★重要：Accessが列名（キー名）を自動で名前解決してくれるため、
    '   Excel版にあった「Match関数による列位置の特定」は完全に不要になります！
    Dim KeyName As Variant
    For Each KeyName In InputData.keys
        rs.Fields(KeyName).value = InputData(KeyName)
    Next

    ' 4. データベースへの書き込みを確定（コミット）
    rs.Update
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
Private Function deserializeFromTable(TableName As String, UserName As String, ErrorSource As String) As Dictionary
    'DB操作用オブジェクトを用意
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim sql As String

    
    '1. 対象のUserNameのレコードをSQLで1行だけピンポイントに取得
    sql = "SELECT * FROM " & TableName & " WHERE " & PrimaryKeyName & " = '" & UserName & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    '2. レコードが存在しない（登録がない）場合は、`Nothing`を返して終了
    If rs.EOF Then Exit Function

    '3. 返却用のDictionaryを作成
    Set deserializeFromTable = New Dictionary

    '4. ★ここが真骨頂：レコードセットの全列（Fields）をオブジェクト指向のループでスキャン
    Dim fld As DAO.Field
    For Each fld In rs.Fields
        ' 「UserName」の情報は不要なので、それ以外の列をDictionaryに格納
        If Not fld.Name = PrimaryKeyName Then deserializeFromTable.Add fld.Name, fld.value
    Next fld
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
Private Sub DeleteTableData(TableName As String, UserName As String, ErrorSource As String)
    'DB操作用オブジェクトを用意
    Dim db As DAO.Database: Set db = CurrentDb
    Dim sql As String


    ' 1. UserNameの有無によって、削除用SQL文を分岐させる
    If StrPtr(UserName) > 0 Then
        ' 個別削除：UserNameが一致するレコードだけをピンポイントで削除
        sql = "DELETE * FROM " & TableName & " WHERE " & PrimaryKeyName & " = '" & UserName & "'"
    Else
        ' 全削除：テーブル内のすべてのレコードを根こそぎ削除
        sql = "DELETE * FROM " & TableName
    End If

    ' 2. SQLを「一撃」で実行！
    ' ※dbFailOnErrorを指定しておくと、もし万が一ロックなどで失敗した際にエラーを検知できます [1.1.2]
    db.Execute sql, dbFailOnError
End Sub



'***************************************************************************************************
'                               ■■■ ヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：現在設定中の識別名称を取得/設定します
'***************************************************************************************************
Property Get CurrentUserName() As String
    '1. 現在の利用ユーザーを取得
    Dim tmp
    tmp = TempVars(Column_SettingUserName)

    '2. Nullの場合は止める
    If IsNull(tmp) Then
        MsgBox "ブラウザの自動操作を開始する前に、設定情報を読み込む必要があります。設定後、再度初めから処理を始めてください。", vbExclamation, "利用ユーザーを確認できません"

        ' ★これだけで、あなたが作った設定画面（フォーム）が一瞬でポップアップします
        DoCmd.OpenForm "SettingRunBrowser", WindowMode:=acDialog
        End
    Else
        CurrentUserName = tmp
    End If
End Property
Property Let CurrentUserName(DefaultUserName As String)
    TempVars(Column_SettingUserName) = DefaultUserName
End Property
