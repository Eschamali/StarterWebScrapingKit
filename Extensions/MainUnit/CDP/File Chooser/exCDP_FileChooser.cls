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
'           `CDPBrowser.cls` に対する機能拡張 - FileChooserインターセプト
'***************************************************************************************************
'* 機能　　：ブラウザのファイル選択ダイアログをCDP経由で横取りし、
'            VBAから指定したファイルパスを注入します。
'            JSで動的に生成されたダイアログも含め、あらゆるファイル選択トリガーに対応します。
'---------------------------------------------------------------------------------------------------
'* 基本的な使い方：
'    Dim fc As New exCDP_FileChooser
'    fc.Init browser
'
'    fc.AddFilePath = "C:\data\file1.txt"    ' ← ファイルを事前登録（複数可）
'    fc.AddFilePath = "C:\data\file2.txt"
'
'    fc.EnableEvents                      ' ← インターセプト有効化/無効化
'    browser.getElementByID("btn").click     ' ← ダイアログをトリガー
'    fc.SetFiles TimeoutSec:=10           ' ← 待機 & 注入（パス不要）
'    fc.DisableEvents                     ' ← 後片付け
'---------------------------------------------------------------------------------------------------
'* 単一ファイルの後方互換 API：
'    fc.SetFile "C:\data\upload.txt"         ' ← パスを直接渡す旧方式（今も使える）
'---------------------------------------------------------------------------------------------------
'* 対応CDPコマンド：
'    - Page.setInterceptFileChooserDialog    (enabled: true/false)
'    - Page.fileChooserOpened               (非同期イベント)
'    - DOM.setFileInputFiles                (backendNodeId 経由でファイル注入)
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - ブラウザが前面（アクティブウィンドウ）にないと fileChooserOpened が発火しません
'    - Init で Page.enable を有効化します（既に有効な場合は副作用なし）
'    - DOM.setFileInputFiles は複数ファイルを同時注入できますが、
'      input 要素に `multiple` 属性がないと最初の1ファイルのみ有効になります
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_CDPCore       As CDPCore      'イベント受け取りオブジェクト
Private ex_CDPBrowser               As CDPBrowser   '継承元のコンテキスト(タブ)
Private ex_CDPContext               As CDPContext   '継承元のブラウザ

'内部使用用
Private Const ThisClassName         As String = "exCDP_FileChooser"  'ログのトレース用
Private Const MaxTimeOut            As Long = 10
Private mJson                       As New WebJsonConverter

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'rawJSONをデバッグログに出すかの調整
Private LogLevel                    As Long         'ログレベル（0:Off, 1:Trace, 2:Debug, 3:Info, 4:Warn, 5:Error）
Private EnableBrowserEvents         As Boolean      'この拡張によるブラウザ非同期イベント処理スイッチ
Private CurrentSessionId            As String       '継承時のSesssionID

'FileChooser イベント受信バッファ
Private Type FileChooserEventBuffer
    backendNodeId       As Long     'fileChooserOpenedで受け取ったbackendNodeId
    mode                As String   'fileChooserOpenedで受け取ったmode (selectSingle / selectMultiple)
    FileChooserDetected As Boolean  'イベント受信済みフラグ
End Type
Private FileChooserEvent As FileChooserEventBuffer

'注入ファイルパスのリスト（AddFilePath で積み上げていく）
Private mFilePaths  As Collection



'***************************************************************************************************
'                               ■■■ 公開プロパティ ■■■
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


    If Len(Trim(FilePath)) = 0 Then
        printMsg WARN_, "Empty file path was ignored.", FromProcedureName
        Exit Property
    End If

    If Dir(FilePath) = "" Then
        printMsg WARN_, "File not found, path skipped: " & FilePath, FromProcedureName
        Exit Property
    End If

    mFilePaths.Add FilePath
    printMsg Debug_, "File path added [" & mFilePaths.Count & "]: " & FilePath, FromProcedureName

End Property

'***************************************************************************************************
'* 機能　　：現在登録されているファイルパスの件数を返します
'***************************************************************************************************
Public Property Get FilePathCount() As Long
    FilePathCount = mFilePaths.Count
End Property

'***************************************************************************************************
'* 機能　　：非同期イベント処理を行うかのスイッチングです
'            この拡張機能の場合はさらに、`Page.setInterceptFileChooserDialog` の有効/無効も合わせて、切り替えます
'---------------------------------------------------------------------------------------------------
'* 機能説明：クラス内のスイッチも切り替えつつ、FileChooserインターセプトを有効化します
'* 注意事項：・登録済みのファイルパス（mFilePaths）はそのまま保持されます
'            ・有効化後はファイル選択トリガーがあっても OSダイアログが出なくなります
'***************************************************************************************************
Public Property Get EnableEvents() As Boolean
    EnableEvents = EnableBrowserEvents
End Property
Public Property Let EnableEvents(ByVal v As Boolean)
    Const FromProcedureName As String = ThisClassName & ".EnableEvents"


    '1. 有効化時は、イベント検出フラグだけリセット（ファイルパスリストは保持）
    If v Then ResetState

    '2. CDPにインターセプトの有効/無効を伝える
    Dim params As New Scripting.Dictionary
    params.Add "enabled", v
    ExecuteCDP "Page.setInterceptFileChooserDialog", params

    '3. クラス内のスイッチも切り替える
    EnableBrowserEvents = v

    '4. ログ出力
    printMsg info_, "FileChooser intercept is " & IIf(v, "enabled.", "disabled."), FromProcedureName

End Property

'***************************************************************************************************
'* 機能　　：最後に検出したFileChooserのbackendNodeIdを返します（デバッグ・低レベル用）
'***************************************************************************************************
Public Property Get LastBackendNodeId() As Long
    LastBackendNodeId = FileChooserEvent.backendNodeId
End Property

'***************************************************************************************************
'* 機能　　：最後に検出したFileChooserのモードを返します（"selectSingle" or "selectMultiple"）
'***************************************************************************************************
Public Property Get LastMode() As String
    LastMode = FileChooserEvent.mode
End Property

'***************************************************************************************************
'* 機能　　：FileChooserイベントが受信済みかどうかを返します
'***************************************************************************************************
Public Property Get IsDetected() As Boolean
    IsDetected = FileChooserEvent.FileChooserDetected
End Property



'***************************************************************************************************
'              ■■■ メインとなる機能拡張部分 非同期イベント編 ■■■
'***************************************************************************************************
'* 機能　　：`Page.fileChooserOpened` イベントを受け取り、backendNodeId を記録します
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
    If CurrentSessionId <> sessionID Then Exit Sub

    '3. 各種非同期イベントの処理
    Select Case methodName
        Case "Page.fileChooserOpened"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & RawJson, FromProcedureName

            '4. デバッグログ用に、Jsonパース処理
            Dim ev As Dictionary: Set ev = mJson.ParseJson(RawJson)("params")
            If seeRawSendMsgDbg Then
                printMsg Debug_, String(80, "="), FromProcedureName
                printMsg Debug_, "Page.fileChooserOpened - detail", FromProcedureName
                printMsg Debug_, String(80, "="), FromProcedureName
                printMsg Debug_, "  frameId       : " & ev("frameId"), FromProcedureName
                printMsg Debug_, "  mode          : " & ev("mode"), FromProcedureName
                printMsg Debug_, "  backendNodeId : " & ev("backendNodeId"), FromProcedureName
                printMsg Debug_, String(80, "="), FromProcedureName
            End If

            '5. JSONをパースしてbackendNodeId と mode を取り出す
            FileChooserEvent.backendNodeId = CLng(ev("backendNodeId"))
            FileChooserEvent.mode = CStr(ev("mode"))

            '6. 「ファイルダイアログが開いたよ！」というフラグを立てる
            FileChooserEvent.FileChooserDetected = True

            printMsg info_, "FileChooser detected. backendNodeId=" & FileChooserEvent.backendNodeId & " mode=" & FileChooserEvent.mode, FromProcedureName

        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName

    End Select

End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************
'* 機能　　：AddFilePath で登録済みのファイルを使ってダイアログが開くのを待ち注入します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功:True / タイムアウト等の失敗:False
'* 引数　　：TimeoutSec     イベント待機タイムアウト（デフォルト: 10秒）
'---------------------------------------------------------------------------------------------------
'* 使い方  ：
'    fc.AddFilePath = "C:\data\upload.txt"   ← 事前にパスを登録
'    fc.EnableIntercept
'    browser.getElementByID("btn").click      ← ダイアログをトリガー
'    fc.SetFileWait                           ← 待機 & 注入（パス不要）
'    fc.DisableIntercept
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - 注入後も mFilePaths は保持されます（次の操作でもそのまま使えます）
'    - 同じパスを再利用しない場合は ClearFilePaths を呼び出してください
'***************************************************************************************************
Public Function SetFiles(Optional TimeoutSec As Long = MaxTimeOut) As Boolean
    Const FromProcedureName As String = ThisClassName & ".SetFiles"


    '--- 事前チェック ---
    If mFilePaths.Count = 0 Then
        printMsg WARN_, "No file paths registered. Call AddFilePath before SetFiles.", FromProcedureName
        SetFiles = False
        Exit Function
    End If

    printMsg info_, "Waiting for FileChooser event... " _
                  & "(files: " & mFilePaths.Count & ", timeout: " & TimeoutSec & "s)", FromProcedureName

    '登録済みパスをデバッグログに出力
    '※Collectionのループは遅いのでログレベル判定をしてから回すこと
    If LogLevelName.Debug_ >= LogLevel And LogLevel > 0 Then
        Dim i As Long
        For i = 1 To mFilePaths.Count
            printMsg Debug_, "  [" & i & "] " & mFilePaths(i), FromProcedureName
        Next i
    End If

    '--- 内部の共通注入処理へ ---
    SetFiles = WaitAndInject(mFilePaths, TimeoutSec)

End Function

'***************************************************************************************************
'* 機能　　：単一ファイルを直接指定する後方互換メソッドです
'---------------------------------------------------------------------------------------------------
'* 引数　　：FilePath       注入するファイルのフルパス
'            TimeoutSec     タイムアウト秒数（デフォルト: 10秒）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：AddFilePath せずに直接呼べる便利メソッドです。
'            内部では単一パスの Collection を作り WaitAndInject を呼びます。
'            本メソッドを呼んでも mFilePaths には影響しません。
'***************************************************************************************************
Public Function SetFile(FilePath As String, Optional TimeoutSec As Long = MaxTimeOut) As Boolean
    Const FromProcedureName As String = ThisClassName & ".SetFile"


    If Dir(FilePath) = "" Then
        printMsg WARN_, "Target file not found: " & FilePath, FromProcedureName
        SetFile = False
        Exit Function
    End If

    printMsg info_, "SetFile (single-shot): " & FilePath, FromProcedureName

    'mFilePaths とは独立した単発 Collection で実行
    Dim tmpCol As New Collection
    tmpCol.Add FilePath
    SetFile = WaitAndInject(tmpCol, TimeoutSec)

End Function

'***************************************************************************************************
'* 機能　　：登録済みファイルパスをすべてクリアします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：AddFilePath でパスを積み直す前に呼ぶといった使い方が想定されます
'***************************************************************************************************
Public Sub ClearFilePaths()
    Const FromProcedureName As String = ThisClassName & ".ClearFilePaths"


    Dim prev As Long: prev = mFilePaths.Count
    Set mFilePaths = New Collection
    printMsg Debug_, "File path list cleared. (" & prev & " paths removed)", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：イベント検出状態をリセットします（EnableEvents=true / SetFileWait の後に自動呼び出し）
'---------------------------------------------------------------------------------------------------
'* 注意事項：ファイルパスリスト（mFilePaths）はクリアしません
'***************************************************************************************************
Public Sub ResetState()
    With FileChooserEvent
        .backendNodeId = -1
        .mode = vbNullString
        .FileChooserDetected = False
    End With
End Sub



'***************************************************************************************************
'                               ■■■ 内部処理 ■■■
'***************************************************************************************************
'* 機能　　：fileChooserOpened を待ってファイルを注入する共通処理
'---------------------------------------------------------------------------------------------------
'* 引数　　：filesCol       注入するファイルパスの Collection
'            TimeoutSec     タイムアウト秒数
'            CallerName     ログのトレース用プロシージャ名
'***************************************************************************************************
Private Function WaitAndInject(filesCol As Collection, TimeoutSec As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".WaitAndInject"


    '--- ① fileChooserOpened イベントが来るまでポーリング ---
    Dim t As Double: t = Timer
    Do
        ex_CDPBrowser.TakeEvents    'パイプから非同期イベントを吸い上げる。これにより上部にある`ex_CDPBrowser_detectionCDPEvent`が呼ばれ、`.FileChooserEvent`に値が入る
        DoEvents

        If FileChooserEvent.FileChooserDetected Then Exit Do

        If (Timer - t) > TimeoutSec Then
            printMsg WARN_, "Timed out waiting for Page.fileChooserOpened. " _
                          & "Make sure the browser is in the foreground.", FromProcedureName
            WaitAndInject = False
            Exit Function
        End If
    Loop
    printMsg Debug_, "fileChooserOpened received in " & Format(Timer - t, "0.0") & "s. " _
                   & "backendNodeId=" & FileChooserEvent.backendNodeId & " mode=" & FileChooserEvent.mode, FromProcedureName

    '--- ② backendNodeId 経由でファイルを注入 ---
    'filesCol を Collection のまま渡すと WebJsonConverter が JSON配列 ["path1","path2",...] にシリアライズ
    Dim injectParams As New Scripting.Dictionary
    injectParams.Add "backendNodeId", FileChooserEvent.backendNodeId
    injectParams.Add "files", filesCol

    Dim injectResult As Scripting.Dictionary
    Set injectResult = ExecuteCDP("DOM.setFileInputFiles", injectParams)
    printMsg Trace_, "DOM.setFileInputFiles sent. files=" & filesCol.Count & " backendNodeId=" & FileChooserEvent.backendNodeId, FromProcedureName

    '--- ③ 後片付け ---
    ResetState
    printMsg info_, "File injection completed. " & filesCol.Count & " file(s) injected.", FromProcedureName
    WaitAndInject = True

End Function



'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_CDPBrowser.sleep seconds
End Sub

Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPContext.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function ExecuteCDP(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set ExecuteCDP = ex_CDPContext.ExecuteCDP(methodName, params, StopError)
End Function



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`CDPBrowser.cls`を継承させます
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    Const FromProcedureName As String = ThisClassName & ".Init"


    '必要なCDPオブジェクトを継承
    Set ex_CDPContext = Inheritance
    Set ex_CDPBrowser = Inheritance.InheritanceCDPBrowser
    Set ex_CDPCore = Inheritance.InheritanceCDPCore

    'SessionIdを取得
    CurrentSessionId = ex_CDPContext.id

    'Page ドメインを有効化（fileChooserOpened イベント受信に必要）
    ex_CDPContext.pageEnable
    printMsg Debug_, "Page domain enabled for FileChooser event reception.", FromProcedureName

    'イベント受信の初期状態は有効にしておく
    EnableEvents = True
End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'パスリストの初期化
    Set mFilePaths = New Collection

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value    'rawJSONをデバッグ出力するか
        LogLevel = .Range(.UseRangeName(5, FromProcedureName)).value            'ログレベルの設定
    End With

End Sub

'***************************************************************************************************
'* 機能　　：終了時にインターセプトを自動無効化して後片付けをします
'***************************************************************************************************
Private Sub Class_Terminate()
    Const FromProcedureName As String = ThisClassName & ".Class_Terminate"


    'インターセプトが残っているまま消えると以降のファイル選択が全部ブロックされるため解除
    On Error Resume Next
    EnableBrowserEvents = False
    ex_CDPContext.pageEnable False
    printMsg Debug_, "FileChooser intercept auto-disabled on class termination.", FromProcedureName

End Sub
