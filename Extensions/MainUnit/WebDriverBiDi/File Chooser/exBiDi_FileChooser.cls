VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exBiDi_FileChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'       WebDriver BiDi の `input.setFiles` / `input.fileDialogOpened` によるファイル添付を行います
'---------------------------------------------------------------------------------------------------
'   CDP版（`exCDP_FileChooser.cls`）の`Page.fileChooserOpened`に相当するイベントとして、
'   BiDiにも`input.fileDialogOpened`が存在します（Chromium系のBiDi-CDP Mapperが、内部で
'   `Page.fileChooserOpened`をそのままBiDi形式に変換して再送しているだけです）。
'   そのため、このClassはCDP版とほぼ同じ構成（非同期監視・キャンセルモード・添付忘れリトライ）を
'   持たせています。
'
'   ただし、`Page.setInterceptFileChooserDialog`（OSダイアログの実際の抑制）が働くには、
'   `session.new`時点で`unhandledPromptBehavior.file`（or`.default`）を`"ignore"`以外に
'   しておく必要があります（既定は`"ignore"`＝抑制なし）。抑制が効いていない状態だと、
'   イベント自体は届きますが、実際のOSダイアログが開こうとする可能性が残ります。
'   `ShSetting01_StartBrowser.StartBiDiModeContext`の`sessionCapabilitiesRequest`引数で
'   指定してください（Demo_FileChooser.basの`BuildFileDialogCapabilities`を参照）。
'
'   詳細・CDP版との比較、および未検証の点は、このフォルダの ReadMe.md を参照してください。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_BiDiCore  As WebDriverBiDiCore
Attribute ex_BiDiCore.VB_VarHelpID = -1
Private ex_BiDiContext          As WebDriverBiDiContext

'内部使用用
Private Const ThisClassName     As String = "exBiDi_FileChooser"

'ステート管理
Private currentLog             As StateLog  'ログ設定
Private EnableBrowserEvents     As Boolean  'この拡張機能によるブラウザ非同期イベント処理を有効化するかしないかのスイッチング
Private CancelMode              As Boolean  '`input.fileDialogOpened`受信時、添付を行わずスキップするか
Private m_context               As String   '`WebDriverBiDiContext`継承時の`Context`情報
Private WaitLock                 As Boolean  '`AutoWaitFileDialogOpened`の待機ループ用ロック

'ファイル添付系のstatus管理
Private mFilePaths               As Collection   '`input.setFiles - files`に渡すファイルパスのリスト
Private UnprocessedFileDialogList As Collection  '`input.fileDialogOpened`は来たが、`mFilePaths`が未登録で、まだ処理してない`sharedId`のリスト



'***************************************************************************************************
'                            ■■■ スコープ調整プロパティ ■■■
'***************************************************************************************************
'* 機能　　：非同期イベント処理を行うかのスイッチングです
'---------------------------------------------------------------------------------------------------
'* Let引数 ：「ファイル選択ダイアログを丸ごとキャンセルするモード」を発動するかの論理値
'* 引数　　：cancelMode     `True`で、`input.fileDialogOpened`受信時、添付を一切行わずスキップします
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`cancelMode`は、デバック/テストコード用を想定してます
'            ・OSダイアログそのものを抑制したい場合は、`session.new`側の`unhandledPromptBehavior`が
'              別途必要です（このプロパティだけでは、VBA側の後処理をスキップするだけです）
'***************************************************************************************************
Public Property Let EnableEvents(Optional cancel As Boolean, v As Boolean)
    Const FromProcedureName As String = ThisClassName & ".EnableEvents"


    '1. VBA側のフィルター設定
    EnableBrowserEvents = v
    CancelMode = cancel

    '2. `input.fileDialogOpened`イベントの購読切り替え
    '※`session.subscribe`はセッション単位です。他のタブ/インスタンスと共有点である事に留意してください
    Dim TargetEvents As New Collection
    TargetEvents.Add "input.fileDialogOpened"
    Set ex_BiDiContext.ThisWebDriverBiDiMode.sessionSubscribe(v) = TargetEvents

    If Not ex_BiDiContext.ThisWebDriverBiDiMode.LastBiDiJsonError Is Nothing Then
        printMsg WARN_, "`session.subscribe`に失敗した可能性があります。 > " & ex_BiDiContext.ThisWebDriverBiDiMode.LastBiDiJsonError("message"), FromProcedureName
    End If
End Property



'***************************************************************************************************
'                             ■■■ 取得系プロパティ ■■■
'***************************************************************************************************
'* 機能　　：現在登録されているファイルパスの件数を返します
'***************************************************************************************************
Public Property Get FilePathCount() As Long
    FilePathCount = mFilePaths.Count
End Property

'***************************************************************************************************
'* 機能　　：`input.fileDialogOpened`の未処理件数を返します
'***************************************************************************************************
Public Property Get UnprocessedCount() As Long
    UnprocessedCount = UnprocessedFileDialogList.Count
End Property



'***************************************************************************************************
'                               ■■■ 設定系プロパティ ■■■
'***************************************************************************************************
'* 機能　　：添付予定のファイルパスを1件ずつ追加します
'---------------------------------------------------------------------------------------------------
'* 使い方　：
'    fc.AddFilePath = "C:\data\file1.txt"
'    fc.AddFilePath = "C:\data\file2.txt"    ← 複数件は`multiple`属性のinputへ
'***************************************************************************************************
Public Property Let AddFilePath(FilePath As String)
    Const FromProcedureName As String = ThisClassName & ".AddFilePath"


    '1. 空文字判定
    If LenB(Trim(FilePath)) = 0 Then printMsg WARN_, "Empty file path was ignored.", FromProcedureName: Exit Property

    '2. 存在判定
    If LenB(Dir(FilePath)) = 0 Then printMsg WARN_, "File not found, path skipped: " & FilePath, FromProcedureName: Exit Property

    '3. 追加
    mFilePaths.Add FilePath
    printMsg info_, "File path added [" & mFilePaths.Count & "]: " & FilePath, FromProcedureName
End Property

'***************************************************************************************************
'* 機能　　：登録済みファイルパスをすべてクリアします
'***************************************************************************************************
Public Sub ClearFilePaths()
    Set mFilePaths = New Collection
End Sub

'***************************************************************************************************
'* 機能　　：未処理リストをクリアします
'***************************************************************************************************
Public Sub ClearUnprocessed()
    Set UnprocessedFileDialogList = New Collection
End Sub



'***************************************************************************************************
'                      ■■■ メインとなる機能拡張部分 - イベント編 ■■■
'***************************************************************************************************
Private Sub ex_BiDiCore_detectionBiDiEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_BiDiCore_detectionBiDiEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If Not EnableBrowserEvents Then Exit Sub

    '2. 対象イベントのみ処理
    If methodName <> "input.fileDialogOpened" Then Exit Sub

    Dim paramsNode As BiDiCDPJson
    Set paramsNode = BiDiCDPJson.Parse(EventRawJson).NodeKey("params")

    '3. このタブのcontext以外は無視
    If paramsNode.StringKey("context") <> m_context Then Exit Sub

    '4. `element`（script.SharedReference）から`sharedId`を取り出す
    Dim sharedId As String
    If paramsNode.ExistsKey("element") Then sharedId = paramsNode.NodeKey("element").StringKey("sharedId")

    If LenB(sharedId) = 0 Then
        printMsg WARN_, "`input.fileDialogOpened` was received, but `element.sharedId` was not included.", FromProcedureName
        WaitLock = False
        Exit Sub
    End If

    '5. キャンセルモードなら、ログのみに留める
    If CancelMode Then
        printMsg Debug_, "`input.fileDialogOpened` received, but it will be skipped because it is in cancel mode. sharedId: " & sharedId, FromProcedureName
    Else
        '6. 登録済みファイルをチェック
        If Me.FilePathCount Then
            'ファイルリストがあれば添付実行
            AttachFiles sharedId
        Else
            'ファイルリストが無い場合は、退避
            UnprocessedFileDialogList.Add sharedId
            printMsg WARN_, "`input.fileDialogOpened` received, but the file list is not yet registered. Added to pending list. sharedId: " & sharedId, FromProcedureName
        End If
    End If

    '7. `AutoWaitFileDialogOpened`の待機ループを解除
    WaitLock = False
End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************
'* 機能　　：1回分の`input.fileDialogOpened`が来るまで待機します
'---------------------------------------------------------------------------------------------------
'* 返り値　：待機成功論理値
'* 引数　　：TimeOutSecond      タイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 注意事項：`AddFilePath`にて、事前に登録しておくこと
'***************************************************************************************************
Public Function AutoWaitFileDialogOpened(Optional TimeOutSecond As Double = 30) As Boolean
    Const FromProcedureName As String = ThisClassName & ".AutoWaitFileDialogOpened"


    '1. 未登録の場合は、即抜け
    If mFilePaths.Count = 0 Then printMsg WARN_, "No files have been registered.", FromProcedureName: Exit Function

    '2. 現在の経過時間を取得
    Dim timerStart As Double: timerStart = TimerCounter

    '3. イベントが来るまでループ待機
    WaitLock = True
    Do
        If TimerCounter - timerStart > TimeOutSecond * 1000 Then printMsg WARN_, "The `input.fileDialogOpened` event could not be detected within the time limit.", FromProcedureName: Exit Function
        TakeEvents
    Loop While WaitLock

    AutoWaitFileDialogOpened = True
End Function

'***************************************************************************************************
'* 機能　　：`UnprocessedFileDialogList`に退避された先頭から1件処理します
'---------------------------------------------------------------------------------------------------
'* 返り値　：添付成功論理値
'* 注意事項：`AddFilePath`にて、事前に登録しておくこと
'***************************************************************************************************
Public Function RetrySetFileInputFiles() As Boolean
    Const FromProcedureName As String = ThisClassName & ".RetrySetFileInputFiles"


    '1. リストが無い場合は即抜け
    If UnprocessedCount = 0 Then printMsg WARN_, "No Unprocessed have been registered.", FromProcedureName: Exit Function

    '2. ファイルリストが無い場合も即抜け
    If FilePathCount = 0 Then printMsg WARN_, "No files have been registered.", FromProcedureName: Exit Function

    '3. 先頭のsharedIdに対して添付実行
    AttachFiles CStr(UnprocessedFileDialogList(1))

    '4. 先頭を削除
    UnprocessedFileDialogList.Remove 1

    '5. 成功として返す
    RetrySetFileInputFiles = True
    printMsg info_, "ReAttachment complete.", FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：実際に添付処理を行い、ファイル添付リストをクリアします
'---------------------------------------------------------------------------------------------------
'* 引数　　：sharedId    `input.fileDialogOpened`から得た、または`browsingContext.locateNodes`で
'                        解決した`script.SharedReference`の`sharedId`
'***************************************************************************************************
Private Sub AttachFiles(sharedId As String)
    Const FromProcedureName As String = ThisClassName & ".AttachFiles"


    '1. パラメーターを組み立てて送信
    Dim elementRef As New Dictionary
    elementRef.Add "sharedId", sharedId

    Dim params As New Dictionary
    params.Add "element", elementRef
    params.Add "files", mFilePaths
    ExecuteBiDi "input.setFiles", params

    '2. ログ化
    printMsg info_, "Attachment complete via `input.setFiles`. sharedId: " & sharedId, FromProcedureName

    '3. 不整合を防ぐため、アップロードリストを初期化
    ClearFilePaths
End Sub

'***************************************************************************************************
'* 機能　　：CSSセレクタで指定した`<input type="file">`へ、ダイアログを介さず直接ファイルを添付します
'---------------------------------------------------------------------------------------------------
'* 返り値　：添付成功の論理値
'* 引数　　：cssSelector   対象input要素のCSSセレクタ（例: "#singleInput"）
'            files         添付したいファイルの**フルパス**を格納したCollection
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`browsingContext.locateNodes`（CSSロケータ）で対象要素の`sharedId`を取得し、
'            `input.setFiles`へ渡します。`input.fileDialogOpened`イベントの監視は不要です。
'* 注意事項：`input.setFiles`の`element`パラメーターは`script.SharedReference`型で、`sharedId`が
'            必須です（`handle`のみでは、このプロジェクトの検証上、スキーマ的に通りません）。
'            そのため`jsEval(..., Ownership:=True)`で得られる`handle`は使わず、`sharedId`を
'            直接返す`browsingContext.locateNodes`を利用しています。
'* 参考URL ：https://w3c.github.io/webdriver-bidi/#command-input-setFiles
'            https://w3c.github.io/webdriver-bidi/#command-browsingContext-locateNodes
'***************************************************************************************************
Public Function SetFileInputFiles(cssSelector As String, files As Collection) As Boolean
    Const FromProcedureName As String = ThisClassName & ".SetFileInputFiles"


    '1. 未登録の場合は、即抜け
    If files Is Nothing Then printMsg WARN_, "The file list is undefined.", FromProcedureName: Exit Function
    If files.Count = 0 Then printMsg WARN_, "No files have been registered.", FromProcedureName: Exit Function

    '2. `browsingContext.locateNodes`で対象要素の`sharedId`を取得
    Dim locator As New Dictionary
    locator.Add "type", "css"
    locator.Add "value", cssSelector

    Dim locateParams As New Dictionary
    locateParams.Add "locator", locator

    Dim locateResult As BiDiCDPJson
    Set locateResult = ex_BiDiContext.ExecuteBiDi("browsingContext.locateNodes", locateParams)

    Dim nodes As BiDiCDPJson
    Set nodes = locateResult.NodeKey("nodes")

    If nodes Is Nothing Then printMsg WARN_, "The element specified by `" & cssSelector & "` could not be found.", FromProcedureName: Exit Function
    If nodes.Count = 0 Then printMsg WARN_, "The element specified by `" & cssSelector & "` could not be found.", FromProcedureName: Exit Function

    Dim sharedId As String
    sharedId = nodes.NodeIndex(0).StringKey("sharedId")

    If LenB(sharedId) = 0 Then printMsg WARN_, "The element was found, but no `sharedId` was returned.", FromProcedureName: Exit Function

    '3. `input.setFiles`用のパラメーターを構築して実行
    Dim elementRef As New Dictionary
    elementRef.Add "sharedId", sharedId

    Dim params As New Dictionary
    params.Add "element", elementRef
    params.Add "files", files
    ExecuteBiDi "input.setFiles", params

    SetFileInputFiles = True
    printMsg info_, "Attachment complete via `input.setFiles`. selector: " & cssSelector, FromProcedureName
End Function



'***************************************************************************************************
'                           ■■■ `WebDriverBiDiContext`直通メソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension As String = ".bidi.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub

Private Function ExecuteBiDi(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As BiDiCDPJson
    Set ExecuteBiDi = ex_BiDiContext.ExecuteBiDi(methodName, params, StopError)
End Function

Private Sub TakeEvents(Optional destruction As Boolean)
    ex_BiDiCore.TakeEvents , destruction
End Sub

Private Function TimerCounter() As Double
    TimerCounter = CDPHelpers.TimerCounter
End Function



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：この拡張機能Classを動かすのに必要なオブジェクトを継承し、イベント監視を有効化します
'***************************************************************************************************
Friend Sub Init(Inheritance As WebDriverBiDiContext)
    '必要なオブジェクトを継承
    Set ex_BiDiContext = Inheritance
    Set ex_BiDiCore = Inheritance.ThisWebDriverBiDiCore
    CDPHelpers.InheritLogSettings currentLog, Inheritance.ThisWebDriverBiDiCore.DirectControlCDP.ThisCDPCore.LogSettings
    m_context = Inheritance.context

    '既定でイベント監視を有効化（キャンセルモードはOFF）
    Me.EnableEvents = True
End Sub

'***************************************************************************************************
'* 機能　　：初期設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Set mFilePaths = New Collection
    Set UnprocessedFileDialogList = New Collection

    currentLog.LogID = "BFID" & Format(Rnd * 1000, "000")
End Sub
