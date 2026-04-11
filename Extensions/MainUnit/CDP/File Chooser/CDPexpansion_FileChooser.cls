VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPexpansion_FileChooser"
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
'* 使い方  ：
'    Dim fc As New CDPexpansion_FileChooser
'    fc.Init browser
'
'    '1. インターセプト有効化
'    fc.EnableIntercept
'
'    '2. ファイル選択をトリガーする操作（クリック等）
'    browser.getElementByID("my-file-input").click
'
'    '3. ダイアログが開くのを待ってファイルを注入
'    fc.SetFile "C:\data\upload.txt"
'
'    '4. 終わったら無効化
'    fc.DisableIntercept
'---------------------------------------------------------------------------------------------------
'* 対応CDPコマンド：
'    - Page.setInterceptFileChooserDialog (enabled: true/false)
'    - Page.fileChooserOpened             (非同期イベント)
'    - DOM.setFileInputFiles              (backendNodeId 経由でファイル注入)
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - ブラウザが前面（アクティブウィンドウ）にないと fileChooserOpened が発火しません
'    - Init で Page.enable を有効化します（既に有効な場合は副作用なし）
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "CDPexpansion_FileChooser"  'ログのトレース用
Private mJson                       As New WebJsonConverter

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'rawJSONをデバッグログに出すかの調整
Private EnableBrowserEvents         As Boolean      'この拡張によるブラウザ非同期イベント処理スイッチ

'FileChooser イベント受信バッファ
Private mBackendNodeId              As Long         'fileChooserOpenedで受け取ったbackendNodeId
Private mMode                       As String       'fileChooserOpenedで受け取ったmode (selectSingle / selectMultiple)
Private mFileChooserDetected        As Boolean      'イベント受信済みフラグ



'***************************************************************************************************
'                               ■■■ 公開プロパティ ■■■
'***************************************************************************************************
'***************************************************************************************************
'* 機能　　：非同期イベント処理を行うかのスイッチングです
'***************************************************************************************************
Public Property Get EnableEvents() As Boolean
    EnableEvents = EnableBrowserEvents
End Property
Public Property Let EnableEvents(ByVal v As Boolean)
    EnableBrowserEvents = v
End Property

'***************************************************************************************************
'* 機能　　：最後に検出したFileChooserのbackendNodeIdを返します（デバッグ・低レベル用）
'***************************************************************************************************
Public Property Get LastBackendNodeId() As Long
    LastBackendNodeId = mBackendNodeId
End Property

'***************************************************************************************************
'* 機能　　：最後に検出したFileChooserのモードを返します（"selectSingle" or "selectMultiple"）
'***************************************************************************************************
Public Property Get LastMode() As String
    LastMode = mMode
End Property

'***************************************************************************************************
'* 機能　　：FileChooserイベントが受信済みかどうかを返します
'***************************************************************************************************
Public Property Get IsDetected() As Boolean
    IsDetected = mFileChooserDetected
End Property



'***************************************************************************************************
'              ■■■ メインとなる機能拡張部分 非同期イベント編 ■■■
'***************************************************************************************************
'* 機能　　：`Page.fileChooserOpened` イベントを受け取り、backendNodeId を記録します
'***************************************************************************************************
Private Sub ex_CDPBrowser_detectionCDPEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPBrowser_detectionCDPEvent"


    '1. 非同期イベント処理自体が無効状態なら、即抜け
    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled in this extended class.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled in this extended class.", FromProcedureName
        Exit Sub
    End If

    '2. 対象のイベント以外は即抜け
    Select Case methodName
        Case "Page.fileChooserOpened"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName

            '3. JSONをパースしてbackendNodeId と mode を取り出す
            If seeRawSendMsgDbg Then
                Dim ev As Dictionary: Set ev = mJson.ParseJson(EventRawJson)("params")
                printMsg Debug_, String(80, "="), FromProcedureName
                printMsg Debug_, "Page.fileChooserOpened ? detail", FromProcedureName
                printMsg Debug_, String(80, "="), FromProcedureName
                printMsg Debug_, "  backendNodeId : " & ev("backendNodeId"), FromProcedureName
                printMsg Debug_, "  mode          : " & ev("mode"), FromProcedureName
                printMsg Debug_, String(80, "="), FromProcedureName
            End If

            Dim params As Dictionary: Set params = mJson.ParseJson(EventRawJson)("params")
            mBackendNodeId = CLng(params("backendNodeId"))
            mMode = CStr(params("mode"))
            mFileChooserDetected = True

            printMsg info_, "FileChooser detected. backendNodeId=" & mBackendNodeId & " mode=" & mMode, FromProcedureName

        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName
            Exit Sub
    End Select

End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：FileChooserインターセプトを有効化します（CDPコマンド送信）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`Page.setInterceptFileChooserDialog` を `enabled:true` で送ります
'            有効化後はファイル選択トリガーがあっても OSダイアログが出なくなります
'***************************************************************************************************
Public Sub EnableIntercept()
    Const FromProcedureName As String = ThisClassName & ".EnableIntercept"


    printMsg info_, "Enabling FileChooser intercept...", FromProcedureName

    '内部フラグをリセットしてから有効化
    ResetState

    Dim params As New Scripting.Dictionary
    params("enabled") = True
    invokeMethod "Page.setInterceptFileChooserDialog", params
    EnableBrowserEvents = True

    printMsg info_, "FileChooser intercept enabled.", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：FileChooserインターセプトを無効化します（CDPコマンド送信）
'***************************************************************************************************
Public Sub DisableIntercept()
    Const FromProcedureName As String = ThisClassName & ".DisableIntercept"


    printMsg info_, "Disabling FileChooser intercept...", FromProcedureName

    Dim params As New Scripting.Dictionary
    params("enabled") = False
    invokeMethod "Page.setInterceptFileChooserDialog", params
    EnableBrowserEvents = False

    printMsg info_, "FileChooser intercept disabled.", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：ファイル選択ダイアログが開くのを待ち、指定ファイルを注入します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功:True / タイムアウト等の失敗:False
'* 引数　　：FilePath       注入するファイルのフルパス（例: "C:\data\upload.txt"）
'            TimeoutSec     イベント待機タイムアウト秒数（デフォルト: 10秒）
'---------------------------------------------------------------------------------------------------
'* 使い方  ：
'    fc.EnableIntercept
'    browser.getElementByID("file-btn").click   ← これがトリガー
'    fc.SetFile "C:\upload.txt"                 ← イベント待ちで自動注入
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - ブラウザが前面（フォーカスされた状態）でないと動作しません
'    - TakeEvents ループ内で自動的に fileChooserOpened を処理します
'***************************************************************************************************
Public Function SetFile(FilePath As String, Optional TimeoutSec As Long = 10) As Boolean
    Const FromProcedureName As String = ThisClassName & ".SetFile"


    printMsg info_, "Waiting for FileChooser event... (timeout: " & TimeoutSec & "s)", FromProcedureName

    'ファイルの存在確認
    If Dir(FilePath) = "" Then
        printMsg WARN_, "Target file not found: " & FilePath, FromProcedureName
        SetFile = False
        Exit Function
    End If
    printMsg Debug_, "Target file confirmed: " & FilePath, FromProcedureName

    '--- ① fileChooserOpened イベントが来るまでポーリング ---
    Dim t As Double: t = Timer
    Do
        ex_CDPBrowser.TakeEvents     'パイプから非同期イベントを吸い上げる
        DoEvents

        If mFileChooserDetected Then Exit Do

        If (Timer - t) > TimeoutSec Then
            printMsg WARN_, "Timed out waiting for Page.fileChooserOpened. " _
                          & "Make sure the browser is in the foreground before calling SetFile.", FromProcedureName
            SetFile = False
            Exit Function
        End If

        ex_CDPBrowser.sleep 0.1     'CPU節約
    Loop
    printMsg Debug_, "fileChooserOpened received in " & Format(Timer - t, "0.0") & "s", FromProcedureName

    '--- ② backendNodeId 経由でファイルを注入 ---
    printMsg Debug_, "Injecting file via DOM.setFileInputFiles. backendNodeId=" & mBackendNodeId, FromProcedureName

    '   files は JSON配列が必要。Collection を渡すと WebJsonConverter が ["path"] にシリアライズしてくれる
    Dim injectParams As New Scripting.Dictionary
    injectParams.Add "backendNodeId", mBackendNodeId

    Dim filesCol As New Collection
    filesCol.Add FilePath                       'ファイルパスを Collection に追加（= JSON配列 ["path"] になる）
    injectParams.Add "files", filesCol

    Dim injectResult As Scripting.Dictionary
    Set injectResult = invokeMethod("DOM.setFileInputFiles", injectParams)
    printMsg Trace_, "DOM.setFileInputFiles sent via invokeMethod. backendNodeId=" & mBackendNodeId, FromProcedureName

    '--- ③ 後片付け ---
    ResetState
    printMsg info_, "File injection completed: " & FilePath, FromProcedureName
    SetFile = True

End Function

'***************************************************************************************************
'* 機能　　：内部状態をリセットします（EnableIntercept や SetFile 後に自動呼び出し）
'***************************************************************************************************
Public Sub ResetState()
    mBackendNodeId = 0
    mMode = ""
    mFileChooserDetected = False
End Sub



'***************************************************************************************************
'                           ■■■ `CDPBrowser`直通メソッド ■■■
'***************************************************************************************************
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_CDPBrowser.sleep seconds
End Sub

Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set invokeMethod = ex_CDPBrowser.invokeMethod(methodName, params, alwaysBrowserContext, StopError)
End Function



'***************************************************************************************************
'                             ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：CDP制御のベースとなる`CDPBrowser.cls`を継承させます
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPBrowser)
    Const FromProcedureName As String = ThisClassName & ".Init"


    '継承
    Set ex_CDPBrowser = Inheritance

    'Page ドメインを有効化（fileChooserOpened イベント受信に必要）
    ex_CDPBrowser.pageEnable
    printMsg Debug_, "Page domain enabled for FileChooser event reception.", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value    'rawJSONをデバッグ出力するか
    End With

End Sub

'***************************************************************************************************
'* 機能　　：終了時にインターセプトを自動無効化して後片付けをします
'***************************************************************************************************
Private Sub Class_Terminate()
    Const FromProcedureName As String = ThisClassName & ".Class_Terminate"


    'インターセプトが残っているまま消えると以降のファイル選択が全部ブロックされるため解除
    On Error Resume Next
    DisableIntercept
    printMsg Debug_, "FileChooser intercept auto-disabled on class termination.", FromProcedureName

End Sub
