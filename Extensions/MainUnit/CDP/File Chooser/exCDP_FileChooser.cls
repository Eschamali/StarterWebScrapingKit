VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_FileChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'                               ファイル添付のアシストを行います
'      `Page.fileChooserOpened`の監視を利用して、`DOM.setFileInputFiles`で、添付処理をします
'
'   このアシスト機能により、OSのファイル選択ダイアログ（Win32 APIでウィンドウハンドルを探して……みたいな泥臭い処理）を操作する必要は、不要です！！
'   ブラウザのバージョンによっては、`userGesture:=True`で実行しないと、`Page.fileChooserOpened`が発火しません
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'ブラウザ制御オブジェクト
Private WithEvents ex_CDPCore       As CDPCore      'イベント受け取りオブジェクト
Attribute ex_CDPCore.VB_VarHelpID = -1
Private ex_CDPContext               As CDPContext   '継承元のタブクラス

'内部使用用
Private Const ThisClassName         As String = "exCDP_FileChooser"

'基本ステート管理
Private seeRawSendMsgDbg            As Boolean      'デバッグメッセージ表示用
Private EnableBrowserEvents         As Boolean      'この拡張によるブラウザ非同期イベント処理スイッチ
Private CurrentSessionID            As String       '継承時のSesssionID

'ファイル添付系のstatus管理
Private LastFileChooserOpened       As String       '最後の`Page.fileChooserOpened`情報(Json文字列のまま)
Private CancelMode                  As Boolean      '`Page.setInterceptFileChooserDialog`が`cancel`モードか？
Private mFilePaths                  As Collection   '`DOM.setFileInputFiles - files`に渡すリスト
Private UnprocessedSetFileList      As Collection   '`Page.setInterceptFileChooserDialog`が来たが、`mFilePaths`がセットされておらず、未処理のままのリスト
Private WaitLock                    As Boolean      '`Page.fileChooserOpened`が来るまでの同期ロック



'***************************************************************************************************
'                            ■■■ スコープ調整プロパティ ■■■
'***************************************************************************************************
'* 機能　　：非同期イベント処理を行うかのスイッチングです
'---------------------------------------------------------------------------------------------------
'* Let引数 ：「ファイル選択ダイアログ横取りモード」を発動するかの論理値
'* 引数　　：cancelMode     `True`で、キャンセル扱いとしてブラウザに送信します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・スコープ範囲の調整にどうぞ。名称は、「Application.EnableEvents」から取りました
'            ・`cancelMode`は、デバック/テストコード用を想定してます
'***************************************************************************************************
Public Property Let EnableEvents(Optional cancel As Boolean, v As Boolean)
    '1. VBA側のフィルター設定
    EnableBrowserEvents = v

    '2. Pageの有効化
    ex_CDPContext.pageEnable v

    '3. ブラウザ側からも調整
    '「ファイル選択ダイアログ横取りモード」の切り替え
    Dim params As New Dictionary
    params.Add "enabled", v
    If cancel Then params.Add "cancel", cancel  'キャンセルとして扱う場合
    ExecuteCDP "Page.setInterceptFileChooserDialog", params, False

    '4. クラス内変数に設定を記憶
    CancelMode = cancel
End Property



'***************************************************************************************************
'                             ■■■ 情報取得プロパティ ■■■
'***************************************************************************************************
'* 機能　　：現在登録されているファイルパスの件数を返します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：この件数チェックで、変化の前後をとらえて、アップロード処理済みかのチェックも可能です
'***************************************************************************************************
Public Property Get FilePathCount() As Long
    FilePathCount = mFilePaths.Count
End Property

'***************************************************************************************************
'* 機能　　：`Page.fileChooserOpened`の未処理数を返します
'***************************************************************************************************
Property Get UnprocessedCount() As Long
    UnprocessedCount = UnprocessedSetFileList.Count
End Property



'***************************************************************************************************
'                               ■■■ 設定プロパティ ■■■
'***************************************************************************************************
'* 機能　　：注入するファイルパスを1件ずつ追加します
'---------------------------------------------------------------------------------------------------
'* 使い方　：
'    fc.AddFilePath = "C:\data\file1.txt"
'    fc.AddFilePath = "C:\data\file2.txt"    ← 複数 `input[multiple]` 向け
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - ファイルが存在しない場合は警告ログを出して追加をスキップします
'    - ClearFilePaths でリセットできます
'***************************************************************************************************
Public Property Let AddFilePath(FilePath As String)
    Const FromProcedureName As String = ThisClassName & ".AddFilePath"


    '1. 空文字判定
    If LenB(Trim(FilePath)) = 0 Then
        printMsg WARN_, "Empty file path was ignored.", FromProcedureName
        Exit Property
    End If

    '2. 存在判定
    If LenB(Dir(FilePath)) = 0 Then
        printMsg WARN_, "File not found, path skipped: " & FilePath, FromProcedureName
        Exit Property
    End If

    '3. 追加
    mFilePaths.Add FilePath
    printMsg info_, "File path added [" & mFilePaths.Count & "]: " & FilePath, FromProcedureName

End Property

'***************************************************************************************************
'* 機能　　：登録済みファイルパスをすべてクリアします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：AddFilePath でパスを積み直す前に呼ぶといった使い方が想定されます
'***************************************************************************************************
Public Sub ClearFilePaths()
    Const FromProcedureName As String = ThisClassName & ".ClearFilePaths"


    Dim prev As Long: prev = mFilePaths.Count
    Set mFilePaths = New Collection
    printMsg info_, "File path list cleared. (" & prev & " paths removed)", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：未処理リストをクリアします
'***************************************************************************************************
Public Sub ClearUnprocessed()
    Const FromProcedureName As String = ThisClassName & ".ClearUnprocessed"


    Dim prev As Long: prev = UnprocessedSetFileList.Count
    Set UnprocessedSetFileList = New Collection
    printMsg info_, "Unprocessed list cleared. (" & prev & " removed)", FromProcedureName

End Sub



'***************************************************************************************************
'                      ■■■ メインとなる機能拡張部分 - イベント処理 ■■■
'---------------------------------------------------------------------------------------------------
'* 機能   ：`page.interceptFileChooser`のみ、パース処理したのをただただ、Dictionaryに蓄積していきます
'***************************************************************************************************
Private Sub ex_CDPCore_CDPContextEvent(methodName As String, RawJson As String, sessionID As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPCore_CDPContextEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. このコンテキスト(タブ)クラスで設定中の`SessionId`以外も即抜け
    If CurrentSessionID <> sessionID Then Exit Sub

    '3. 対象のイベントを処理
    Dim FileChooserOpenedParams As BiDiCDPJson
    Select Case methodName
        Case "Page.fileChooserOpened"
            LastFileChooserOpened = RawJson
            Set FileChooserOpenedParams = BiDiCDPJson.Parse(RawJson).NodeKey("params")

            '4. 主要な情報を取り出す
            Dim Target_backendNodeId As Long
            Target_backendNodeId = FileChooserOpenedParams.NumberKey("backendNodeId")

            '5. デバック用に表示
            If seeRawSendMsgDbg Then
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "Page.fileChooserOpened information", FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "frameId       : " & FileChooserOpenedParams.StringKey("frameId"), FromProcedureName
                printMsg Debug_, "mode          : " & FileChooserOpenedParams.StringKey("mode"), FromProcedureName
                printMsg Debug_, "backendNodeId : " & Target_backendNodeId, FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName
            End If

            '5. キャンセルモードの場合は、ログのみに留める
            If CancelMode Then
                printMsg Debug_, "Page.fileChooserOpened - backendNodeId: " & Target_backendNodeId & " was received, but it will be skipped because it is in cancel mode.", FromProcedureName
            Else
                '6. 登録ファイル数チェック
                If Me.FilePathCount Then
                    'ファイルリストがあれば添付処理
                    printMsg Debug_, "Attach the file to Page.fileChooserOpened - backendNodeId: " & Target_backendNodeId, FromProcedureName
                    SetFileInputFiles Target_backendNodeId
                Else
                    'ファイルリストがない場合は、退避
                    UnprocessedSetFileList.Add Target_backendNodeId
                    printMsg WARN_, "Page.fileChooserOpened - backendNodeId: " & Target_backendNodeId & " was received, but the file list is not yet registered. It has been added to the pending list.", FromProcedureName
                End If
            End If

            '7. `AutoWaitFileChooserOpened`による待機ループを解除させておく
            WaitLock = False

        'ページ遷移時の過去の未処理リストの初期化
        Case "Page.frameNavigated"
            Set UnprocessedSetFileList = New Collection
            printMsg info_, "A page transition has occurred, so the Unprocessed list will be reset.", FromProcedureName

    End Select
End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************
'* 機能    ：1件分の`Page.fileChooserOpened`が来るまで待機します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：待機成功可否
'---------------------------------------------------------------------------------------------------
'* 引数　　：TimeOutSecond      タイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 機能説明：`EnableEvents`にて、`cancel`に応じて挙動が変わる待機メソッドです
'            ・`cancel:=True` ：イベントが来たら、予め設定した`mFilePaths`で、添付処理を行いその後、`mFilePaths`をクリアします。その後、待機ループから抜けます。
'            ・`cancel:=False`：イベントが来たら、待機ループを抜けますが、`mFilePaths`はクリアしませんし、添付もしません
'
'* 注意事項：・`AddFilePath`にて、事前に登録してから、実行すること
'            ・処理の仕方によっては、前述の処理(例:`.SimpleClick`)のついでにファイル添付されたといったパターンもあります
'              その場合は、`FilePathCount=0`という判定で「ファイル添付リストに登録したのに確認の時点では`0`ということは、もう添付済みだな」と判定可能です
'***************************************************************************************************
Public Function AutoWaitFileChooserOpened(Optional TimeOutSecond As Double = 30) As Boolean
    Const FromProcedureName As String = ThisClassName & ".AutoWaitFileChooserOpened"


    '1. 未登録の場合は、即抜け
    If mFilePaths.Count = 0 Then printMsg WARN_, "No files have been registered.", FromProcedureName: Exit Function

    '2. 現在の経過時間を取得
    Dim timerStart As Double: timerStart = TimerCounter

    '3. イベントが届くまでループ待機
    WaitLock = True
    Do
        '4. タイムアウトしたら、抜ける
        If TimerCounter - timerStart > TimeOutSecond * 1000 Then printMsg WARN_, "The `Page.fileChooserOpened` event could not be detected within the time limit.", FromProcedureName: Exit Function

        '5. イベント取り出し
        TakeEvents  '内部で`DoEvents`が走ってくれます
    Loop While WaitLock

    '6. 待機成功返却
    AutoWaitFileChooserOpened = True
End Function

'***************************************************************************************************
'* 機能   ：`UnprocessedSetFileList`に退避した物を先頭から1つずつ処理します
'---------------------------------------------------------------------------------------------------
'* 返り値 ：添付成功可否
'---------------------------------------------------------------------------------------------------
'* 注意事項：`AddFilePath`にて、事前に登録してから、実行すること
'***************************************************************************************************
Public Function RetrySetFileInputFiles() As Boolean
    Const FromProcedureName As String = ThisClassName & ".RetrySetFileInputFiles"


    '1. リストにない場合は即抜け
    If UnprocessedCount = 0 Then printMsg WARN_, "No Unprocessed have been registered.", FromProcedureName: Exit Function

    '2. ファイルリストなしも即抜け
    If FilePathCount = 0 Then printMsg WARN_, "No files have been registered.", FromProcedureName: Exit Function

    '3. 漏れのあった先頭の`backendNodeId`に対して、添付処理
    SetFileInputFiles UnprocessedSetFileList(1)

    '4. 先頭を削除
    UnprocessedSetFileList.Remove (1)

    '5. 成功として返す
    RetrySetFileInputFiles = True
    printMsg info_, "ReAttachment complete.", FromProcedureName
End Function

'***************************************************************************************************
'* 機能   ：実際に添付処理を行い、ファイル添付リストをクリアします
'---------------------------------------------------------------------------------------------------
'* 引数　　：InputID    `Page.fileChooserOpened`から得た`backendNodeId`
'---------------------------------------------------------------------------------------------------
'* 注意事項：実行前に、`mFilePaths`が登録されてるか確認すること
'***************************************************************************************************
Private Sub SetFileInputFiles(InputID As Long)
    Const FromProcedureName As String = ThisClassName & ".setFileInputFiles"


    '1. パラメーターを組み立てて送信
    Dim params As New Dictionary
    params.Add "files", mFilePaths
    params.Add "backendNodeId", InputID
    ExecuteCDP "DOM.setFileInputFiles", params

    '2. 別途ログ化
    If seeRawSendMsgDbg Then
        printMsg info_, "Attachment complete: " & WebJsonConverter.serialize(mFilePaths), FromProcedureName
        printMsg info_, "We will initialize the file list in preparation for the next upload.", FromProcedureName
    Else
        printMsg info_, "Attachment complete. We will initialize the file list in preparation for the next upload.", FromProcedureName
    End If

    '3. 不整合を防ぐため、アップロードリストを初期化
    ClearFilePaths
End Sub



'***************************************************************************************************
'                           ■■■ `CDPContext`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.ExecuteCDP`といった記述を防ぐ措置となります。
'            メインの`CDPContext`メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておくこと
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPContext.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function ExecuteCDP(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As BiDiCDPJson
    Set ExecuteCDP = ex_CDPContext.ExecuteCDP(methodName, params, StopError)
End Function

Private Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
    ex_CDPCore.TakeEvents StopPipeError, destruction
End Sub

Private Function TimerCounter() As Double
    TimerCounter = ex_CDPCore.GetTicksMs
End Function



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　  ：この拡張機能Classを動かすのに必要なオブジェクトや情報を渡し、イベント処理機能を有効化させます
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    '必要なCDPオブジェクトを継承
    Set ex_CDPContext = Inheritance
    Set ex_CDPCore = Inheritance.InheritanceCDPCore

    '渡した際のSessionIDをセット
    CurrentSessionID = Inheritance.CurrentSessionID

    '非同期イベント処理を有効化
    EnableEvents = True
End Sub

'***************************************************************************************************
'* 機能　　：初期設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    '各変数の初期化
    Set mFilePaths = New Collection
    Set UnprocessedSetFileList = New Collection

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    seeRawSendMsgDbg = ShSetting01_StartBrowser.UseRangeID(8, FromProcedureName)   'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
End Sub



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
'* 機能　　：拡張機能との連携を切る際に記述します
'***************************************************************************************************
Private Sub Class_Terminate()
    EnableEvents = False
End Sub
