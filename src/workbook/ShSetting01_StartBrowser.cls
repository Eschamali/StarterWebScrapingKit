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
'                              ■■■ WebView2用設定API ■■■
'***************************************************************************************************
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long 'プロセス内環境変数用API



'***************************************************************************************************
'                              ■■■ テーブルの列操作用 ■■■
'***************************************************************************************************
Private Const TableName01 As String = "BrowserHandleInfo"
Private Const TableName02 As String = "BrowserContextInfo"

Private Const TableName03 As String = "ProtocolWhiteList"
Public Enum Table03ColumnID
    ProtocolName = 1
    階層
    接頭辞
End Enum



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
'* 注意事項：Chromeにもキオスクモードはありますが、Edgeほど引数での調整はありません
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
'* 引数　　：前述と一緒なので省略
'---------------------------------------------------------------------------------------------------
'* 注意事項：・このプロシージャだけではまだ、タブ制御するまでの準備が整ってません。
'              `StartCDPModeContext`ではうまくいかず、手動でタブ制御確立するためのものです。
'              この処理後は基本的には、`.getTab` or `.newTab`で、直ぐに確立できると思います。
'            ・Demoプロシージャ：`runTabsAsMany`が分かりやすいと思います
'***************************************************************************************************
Friend Function StartCDPMode(Optional StartURL As String, Optional SwitchUser As String) As CDPBrowser
    '設定シートの各セルから設定値を取得し、適用
    With Me
        '起動ブラウザ種類の設定
        '※CDP－Json コマンドによる操作なので、Chromium系統であれば、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim BrowserName As BrowserList: BrowserName = IIf(.UseRangeID(4, "ShSetting01_StartBrowser.StartCDPMode"), BrowserList.RunChrome, BrowserList.RunEdge)

        '第2引数が省略ならシート側の設定を適用
        If StrPtr(SwitchUser) = 0 Then SwitchUser = .UseRangeID(2, "ShSetting01_StartBrowser.StartCDPMode")

        'ブラウザ起動
        Set StartCDPMode = New CDPBrowser
        StartCDPMode.start BrowserName, StartURL, SwitchUser, .UseRangeID(3, "ShSetting01_StartBrowser.StartCDPMode")
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
'* 引数　　：前述と一緒なので省略
'---------------------------------------------------------------------------------------------------
'* 注意事項：このプロシージャだけではまだ、タブ制御するまでの準備が整ってません。
'             `StartBiDiModeContext`ではうまくいかず、手動でタブ制御確立するためのものです。
'             この処理後は基本的には、`.getTab` or `.newTab`で、直ぐに確立できると思います。
'***************************************************************************************************
Friend Function StartBiDiMode(Optional StartURL As String, Optional SwitchUser As String, Optional sessionCapabilitiesRequest As Dictionary) As WebDriverBiDiMode
    '設定シートの各セルから設定値を取得し、適用
    With Me
        '起動ブラウザ種類の設定
        '※BiDi-Json コマンドによる操作ですが、Chromium系統に特化した制御のため、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim BrowserName As BrowserList: BrowserName = IIf(.UseRangeID(4, "ShSetting01_StartBrowser.StartBiDiMode"), BrowserList.RunChrome, BrowserList.RunEdge)

        '第2引数が省略ならシート側の設定を適用
        If StrPtr(SwitchUser) = 0 Then SwitchUser = .UseRangeID(2, "ShSetting01_StartBrowser.StartBiDiMode")

        'ブラウザ起動
        Set StartBiDiMode = New WebDriverBiDiMode
        StartBiDiMode.StartBiDiMode BrowserName, StartURL, SwitchUser, .UseRangeID(3, "ShSetting01_StartBrowser.StartBiDiMode"), sessionCapabilitiesRequest
    End With
End Function



'***************************************************************************************************
'                             ■■■ 数値でセルの中身取得/反映 ■■■
'***************************************************************************************************
'* 機能　　：このシートのみ定義されてるセル名称を取得します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：セルの中身(`.Value`のみ)
'* 引数　　：RangeID        セルID
'            ErrorSource    この関数を呼び出したソース。※デバック用
'***************************************************************************************************
Property Get UseRangeID(RangeID As Long, ErrorSource As String) As Variant
    UseRangeID = Me.Range(RangeMapping(RangeID, ErrorSource)).value
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
    Dim HeaderColumn: HeaderColumn = Me.ListObjects(TableName01).HeaderRowRange
    Dim RequiredCount: RequiredCount = UBound(HeaderColumn, 2) - 1  '※主キー分、減算

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
    Dim HeaderColumn: HeaderColumn = Me.ListObjects(TableName02).HeaderRowRange
    Dim RequiredCount: RequiredCount = UBound(HeaderColumn, 2) - 1  '※主キー分、減算

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
    '1. 指定のデータテーブル、列配列を取得
    Dim WhiteList: WhiteList = Me.ListObjects(TableName03).DataBodyRange.Columns(Table03ColumnID.接頭辞)

    '2. 配列か？
    If IsArray(WhiteList) Then
        '1次元配列に変換して返却
        ReDim tmp(1 To UBound(WhiteList))
        Dim i As Long
        For i = 1 To UBound(WhiteList)
            tmp(i) = WhiteList(i, 1)
        Next

        TableProtocolWhiteList = tmp
    Else
        '配列じゃない場合は1次元配列にしておく
        ReDim tmp(1 To 1)
        tmp(1) = WhiteList

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
'            ・ハンドル情報管理テーブルは対象外です
'***************************************************************************************************
Property Let UpdateBrowserContext(UserName As String, ReplaceName As String, ErrorSource As String, SetValue As String)
    '1. 操作するテーブルオブジェクトの基準セルを取得
    Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName02, UserName)

    '2. 存在したら、対象ユーザーセルを基準にデータ名を探す
    If Not StartPos Is Nothing Then
        '2-1. ヘッダー情報を取得
        Dim HeaderColumn: HeaderColumn = Me.ListObjects(TableName02).HeaderRowRange

        '2-2. match関数で書き換えデータ列位置を特定
        Dim ColumnPos As Long
        ColumnPos = WorksheetFunction.Match(ReplaceName, HeaderColumn, 0)   '※ここでエラーになる場合は、渡したReplaceName名が間違ってる可能性あり

        '2-3. 実際のデータ値を書き換え
        StartPos.Offset(0, ColumnPos - 1).value = SetValue                  '※オフセットは、0始まりなので補正しておく
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
'                           ■■■ イベントプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：入力規則チェックを行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：エラー時は、デフォルト値に書き換えます。前回の値には戻せません。
'***************************************************************************************************
Private Sub Worksheet_Change(ByVal Target As Range)
    ' --- チェック対象のセルとデフォルト値を用意 ---
    Const checkCell1_DefaultValue As String = "Automation Data"
    Dim checkCell1 As Range: Set checkCell1 = Me.Range(RangeMapping(2, "ShSetting01_StartBrowser.Worksheet_Change"))


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
'            UserName       識別名称(主キー)
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
    ReDim SetArray(InputData.Count)

    '2. データ行数と列数を取得
    Dim InsertDataColumns   As Long: InsertDataColumns = UBound(SetArray) + 1   '列数 ※主キーを含めるので補正
    Const InsertDataRows    As Long = 1                                         '行数 ※1件ずつ追加のため

    '3. UserName(主キー)を指定
    SetArray(0) = UserName

    '4. ヘッダー情報を取得
    Dim TargetTableObj As ListObject: Set TargetTableObj = Me.ListObjects(TableName)
    Dim HeaderColumn: HeaderColumn = TargetTableObj.HeaderRowRange

    '5. Dictionary から、1次元配列へ変換
    Dim i As Long, KeyName, columnIndex As Long
    For Each KeyName In InputData.keys
        '5-1. Keyから列位置を特定
        columnIndex = WorksheetFunction.Match(KeyName, HeaderColumn, 0) '※ここでエラーになる場合は、渡したDictionaryのKey名が間違ってる可能性あり

        '5-2. その位置の配列に適用
        SetArray(columnIndex - 1) = InputData(KeyName) '※0始まりなので補正
    Next

    '6. 操作するテーブルオブジェクトの基準セルを取得
    Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName, UserName)

    '7. データ挿入方式判定
    If StartPos Is Nothing Then
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
Private Function deserializeFromTable(TableName As String, UserName As String, ErrorSource As String) As Dictionary
    '1. 操作するテーブルオブジェクトの基準セルを取得
    Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName, UserName)

    '2. 基準セルがない場合は、即抜け
    If StartPos Is Nothing Then Exit Function

    '3. ヘッダー情報を取得し、格納データ数を取り出す
    Dim HeaderColumn: HeaderColumn = Me.ListObjects(TableName).HeaderRowRange
    Dim TakeColumn As Long: TakeColumn = UBound(HeaderColumn, 2)

    '4. ヘッダー情報分、列から取り出す
    Dim DataSet
    DataSet = StartPos.Resize(1, TakeColumn)

    '5. データ名と値の紐づけを行い、Dictionaryで返却
    '※`UserName`情報は不要なので、2列目から始める
    Dim i As Long
    Set deserializeFromTable = New Dictionary
    For i = 2 To TakeColumn
        deserializeFromTable.Add HeaderColumn(1, i), DataSet(1, i)
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
Private Sub DeleteTableData(TableName As String, UserName As String, ErrorSource As String)
    'テーブル削除時のアニメーション動作を止める
    Application.ScreenUpdating = False


    '文字列が格納されているか？
    If StrPtr(UserName) Then
        '`UserName`をキーに、操作するテーブルオブジェクトの基準セルを取得
        Dim GetTakeNum As Long
        Dim StartPos As Range: Set StartPos = GetReferenceRange(TableName, UserName)

        '基準セルがない場合は、即抜け
        If StartPos Is Nothing Then Application.ScreenUpdating = True: Exit Sub

        'データ数をチェック
        Dim HeaderColumn: HeaderColumn = Me.ListObjects(TableName).HeaderRowRange
        Dim DropColumn As Long: DropColumn = UBound(HeaderColumn, 2)

        '対象のテーブルに応じて、消去行数をリサイズ
        Dim DataSet As Range
        Set DataSet = StartPos.Resize(1, DropColumn)

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



'***************************************************************************************************
'                               ■■■ ヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：実行プロセス内のWebView2デバッグポートを開く際に使います。迷ったら、`9222`に設定して下さい
'---------------------------------------------------------------------------------------------------
'* 注意事項：閉じる場合は、`0`にしてください
'***************************************************************************************************
Property Let EnsureWebView2Debug(port As Long)
    Const EnvironmentName As String = "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"


    If port > 0 Then
        SetEnvironmentVariableW StrPtr(EnvironmentName), StrPtr("--remote-debugging-port=" & port)
        Debug.Print "WebView2のデバッグポートを開けました: " & port
    Else
        SetEnvironmentVariableW StrPtr(EnvironmentName), 0
        Debug.Print "WebView2のデバッグポートを閉じました"
    End If
End Property

'***************************************************************************************************
'* 機能　　：現在設定中の識別名称を取得します
'***************************************************************************************************
Property Get CurrentUserName() As String
    CurrentUserName = Me.UseRangeID(2, "ShSetting01_StartBrowser.CurrentUserName")
End Property
