VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPexpansion_DownloadWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           `CDPBrowser.cls` に対する機能拡張 - ダウンロード監視・完了待ち
'***************************************************************************************************
'* 機能　　：ブラウザのファイルダウンロードを CDP 経由で監視し、
'            指定フォルダへの保存完了まで VBA 側で待機します。
'---------------------------------------------------------------------------------------------------
'* 基本的な使い方：
'    Dim dw As New CDPexpansion_DownloadWatcher
'    dw.Init browser
'
'    dw.WatchStart "C:\download_dir"   ' ← ダウンロード先フォルダを指定して監視開始
'    browser.getElementByID("dl-btn").click   ' ← ダウンロードをトリガー
'    If dw.WaitCompleted(TimeoutSec:=30) Then
'        Debug.Print "保存先: " & dw.DownloadedFilePath
'    End If
'---------------------------------------------------------------------------------------------------
'* 対応CDPイベント：
'    - Browser.downloadWillBegin   (ダウンロード開始通知)
'    - Browser.downloadProgress    (進捗通知 / 完了・失敗)
'    - Page.downloadWillBegin      (Fallback 用、Chromeバージョンによる)
'    - Page.downloadProgress       (Fallback 用、Chromeバージョンによる)
'---------------------------------------------------------------------------------------------------
'* 対応CDPコマンド：
'    - Browser.setDownloadBehavior  (ダウンロード場所指定 & イベント有効化)
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - WatchStart の前にダウンロードをトリガーすると検出を逃す場合があります
'    - ダウンロード先フォルダは事前に存在している必要があります（自動作成します）
'    - .crdownload / .part 等の一時ファイルは完了とみなしません
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "CDPexpansion_DownloadWatcher"  'ログのトレース用
Private mJson                       As New WebJsonConverter

'ステート管理
Private seeRawSendMsgDbg            As Boolean      'rawJSONをデバッグログに出すかの調整
Private EnableBrowserEvents         As Boolean      'この拡張によるブラウザ非同期イベント処理スイッチ

'ダウンロード情報バッファ（イベントから受け取ったデータを保持）
Private mGuid                       As String       'ダウンロードを一意に識別するGUID
Private mSuggestedFilename          As String       'ブラウザが提案するファイル名
Private mDownloadUrl                As String       'ダウンロード対象のURL
Private mState                      As String       '"" / "inProgress" / "completed" / "canceled"
Private mFilePath                   As String       '完了時に確定する実際の保存先パス
Private mReceivedBytes              As Double       '受信済みバイト数
Private mTotalBytes                 As Double       '合計バイト数

'監視設定
Private mWatchDir                   As String       '監視対象ダウンロードフォルダパス
Private mSnapshotFiles              As Object       'WatchStart 時点でのフォルダスナップショット（既存ファイル把握用）



'***************************************************************************************************
'                               ■■■ 公開プロパティ（読み取り専用） ■■■
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
'* 機能　　：ダウンロードの現在の状態を返します
'---------------------------------------------------------------------------------------------------
'* 返り値　："" / "inProgress" / "completed" / "canceled"
'***************************************************************************************************
Public Property Get state() As String
    state = mState
End Property

'***************************************************************************************************
'* 機能　　：ダウンロードが完了しているか返します
'***************************************************************************************************
Public Property Get IsCompleted() As Boolean
    IsCompleted = (mState = "completed")
End Property

'***************************************************************************************************
'* 機能　　：ダウンロードが開始されたか返します（WillBegin イベント受信済みか）
'***************************************************************************************************
Public Property Get HasStarted() As Boolean
    HasStarted = (mGuid <> "")
End Property

'***************************************************************************************************
'* 機能　　：実際に保存されたファイルの完全パスを返します
'---------------------------------------------------------------------------------------------------
'* 注意事項：完了前は空文字が返ります
'***************************************************************************************************
Public Property Get DownloadedFilePath() As String
    DownloadedFilePath = mFilePath
End Property

'***************************************************************************************************
'* 機能　　：ブラウザが提案するファイル名を返します（拡張子付き）
'***************************************************************************************************
Public Property Get SuggestedFilename() As String
    SuggestedFilename = mSuggestedFilename
End Property

'***************************************************************************************************
'* 機能　　：ダウンロード元のURLを返します
'***************************************************************************************************
Public Property Get DownloadUrl() As String
    DownloadUrl = mDownloadUrl
End Property

'***************************************************************************************************
'* 機能　　：受信済みバイト数を返します（進捗表示用）
'***************************************************************************************************
Public Property Get ReceivedBytes() As Double
    ReceivedBytes = mReceivedBytes
End Property

'***************************************************************************************************
'* 機能　　：合計バイト数を返します（0 の場合は不明）
'***************************************************************************************************
Public Property Get TotalBytes() As Double
    TotalBytes = mTotalBytes
End Property

'***************************************************************************************************
'* 機能　　：進捗率を 0.0 ? 1.0 で返します（合計不明の場合は -1）
'***************************************************************************************************
Public Property Get Progress() As Double
    If mTotalBytes <= 0 Then
        Progress = -1
    Else
        Progress = mReceivedBytes / mTotalBytes
    End If
End Property



'***************************************************************************************************
'              ■■■ メインとなる機能拡張部分 非同期イベント編 ■■■
'***************************************************************************************************
'* 機能　　：ダウンロード関連の CDP イベントを受け取り、内部状態に反映します
'            対応イベント:
'              Browser.downloadWillBegin  → ダウンロード開始情報を記録
'              Browser.downloadProgress   → 進捗・完了・キャンセルを記録
'              Page.downloadWillBegin     → Fallback（古い Chrome 向け）
'              Page.downloadProgress      → Fallback（古い Chrome 向け）
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

        '--- ダウンロード開始（Guid・ファイル名・URLを取得） ---
        Case "Browser.downloadWillBegin", "Page.downloadWillBegin"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName

            Dim beginParams As Dictionary: Set beginParams = mJson.ParseJson(EventRawJson)("params")
            mGuid = CStr(beginParams("guid"))
            mSuggestedFilename = CStr(beginParams("suggestedFilename"))
            mDownloadUrl = CStr(beginParams("url"))
            mState = "inProgress"
            mReceivedBytes = 0
            mTotalBytes = 0

            printMsg info_, "[" & methodName & "] Download started. " _
                          & "guid=" & mGuid & " file=" & mSuggestedFilename & " url=" & mDownloadUrl, FromProcedureName

            If seeRawSendMsgDbg Then
                printMsg Debug_, String(80, "="), FromProcedureName
                printMsg Debug_, "  guid              : " & mGuid, FromProcedureName
                printMsg Debug_, "  suggestedFilename : " & mSuggestedFilename, FromProcedureName
                printMsg Debug_, "  url               : " & mDownloadUrl, FromProcedureName
                printMsg Debug_, String(80, "="), FromProcedureName
            End If

        '--- ダウンロード進捗（進捗・完了・キャンセルを更新） ---
        Case "Browser.downloadProgress", "Page.downloadProgress"
            printMsg Trace_, "`" & methodName & "` was detected. Raw-Json : " & EventRawJson, FromProcedureName

            Dim progParams As Dictionary: Set progParams = mJson.ParseJson(EventRawJson)("params")
            Dim evGuid As String: evGuid = CStr(progParams("guid"))

            '複数ダウンロードが走っている場合でも、監視対象の Guid だけ処理
            If evGuid <> mGuid Then
                printMsg Trace_, "Ignored progress event for different guid: " & evGuid, FromProcedureName
                Exit Sub
            End If

            mState = CStr(progParams("state"))
            mReceivedBytes = CDbl(progParams("receivedBytes"))
            mTotalBytes = CDbl(progParams("totalBytes"))

            Select Case mState
                Case "inProgress"
                    If mTotalBytes > 0 Then
                        printMsg Debug_, "Download progress: " _
                                       & Format(mReceivedBytes / 1024, "0.0") & "KB / " _
                                       & Format(mTotalBytes / 1024, "0.0") & "KB " _
                                       & "(" & Format(Progress * 100, "0.0") & "%)", FromProcedureName
                    Else
                        printMsg Debug_, "Download progress: " _
                                       & Format(mReceivedBytes / 1024, "0.0") & "KB (total unknown)", FromProcedureName
                    End If

                Case "completed"
                    '保存先パスを確定する（Browser.downloadProgress は filePath を含む場合がある）
                    If progParams.Exists("filePath") Then
                        If Len(CStr(progParams("filePath"))) > 0 Then
                            mFilePath = CStr(progParams("filePath"))
                        End If
                    End If

                    '--- filePath が取れなかった場合はフォルダ差分から特定 ---
                    If Len(mFilePath) = 0 Then
                        mFilePath = ResolveFilePathFromDir(mSuggestedFilename)
                    End If

                    printMsg info_, "Download completed! file=" & mFilePath, FromProcedureName

                Case "canceled"
                    printMsg WARN_, "Download was canceled. guid=" & mGuid, FromProcedureName

                Case Else
                    printMsg Debug_, "Unknown download state: " & mState, FromProcedureName
            End Select

        Case Else
            printMsg Trace_, "Asynchronous events received are not processed by this class: " & methodName, FromProcedureName
            Exit Sub

    End Select

End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：ダウンロード監視を開始します
'---------------------------------------------------------------------------------------------------
'* 引数　　：DownloadDir    ダウンロード先フォルダのフルパス（存在しない場合は自動作成）
'            AllowAndName   True: "allowAndName"（ブラウザが自動でユニーク名を付ける）
'                           False: "allow"（SuggestedFilename そのまま）← 通常はこちら
'---------------------------------------------------------------------------------------------------
'* 詳細説明：内部で `Browser.setDownloadBehavior` を送信し、
'            イベント受信を有効化したうえで、フォルダの現状スナップショットを取得します
'***************************************************************************************************
Public Sub WatchStart(DownloadDir As String, Optional AllowAndName As Boolean = False)
    Const FromProcedureName As String = ThisClassName & ".WatchStart"


    printMsg info_, "Starting download watch. dir=" & DownloadDir _
                  & " AllowAndName=" & AllowAndName, FromProcedureName

    '--- フォルダが存在しない場合は作成 ---
    If Dir(DownloadDir, vbDirectory) = "" Then
        MkDir DownloadDir
        printMsg Debug_, "Created download directory: " & DownloadDir, FromProcedureName
    End If

    '--- パス末尾スラッシュを正規化 ---
    mWatchDir = DownloadDir
    If Right$(mWatchDir, 1) = "\" Then mWatchDir = Left$(mWatchDir, Len(mWatchDir) - 1)

    '--- 内部状態をリセット ---
    ResetState

    '--- 現時点のフォルダ内容をスナップショット（既存ファイルを把握） ---
    Set mSnapshotFiles = CreateObject("Scripting.Dictionary")
    SnapshotDirectory mWatchDir, mSnapshotFiles
    printMsg Debug_, "Snapshot taken. Existing files in dir: " & mSnapshotFiles.Count, FromProcedureName

    '--- Browser.setDownloadBehavior でダウンロード先とイベントを設定 ---
    Dim behavior As String: behavior = IIf(AllowAndName, "allowAndName", "allow")
    Dim params As New Scripting.Dictionary
    params.Add "behavior", behavior
    params.Add "downloadPath", mWatchDir
    params.Add "eventsEnabled", True

    invokeMethod "Browser.setDownloadBehavior", params, True   'alwaysBrowserContext=True で Browser ドメイン
    EnableBrowserEvents = True

    printMsg info_, "Download behavior set. behavior=" & behavior & " watchDir=" & mWatchDir, FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：ダウンロードが完了するまで待機します
'---------------------------------------------------------------------------------------------------
'* 返り値　：True=完了 / False=タイムアウト or キャンセル
'* 引数　　：TimeoutSec     最大待機時間（秒）/ デフォルト: 60秒
'---------------------------------------------------------------------------------------------------
'* 確認方法：
'   If dw.WaitCompleted(30) Then
'       MsgBox "保存先: " & dw.DownloadedFilePath
'   End If
'***************************************************************************************************
Public Function WaitCompleted(Optional TimeoutSec As Long = 60) As Boolean
    Const FromProcedureName As String = ThisClassName & ".WaitCompleted"


    printMsg info_, "Waiting for download to complete... (timeout: " & TimeoutSec & "s)", FromProcedureName

    Dim t As Double: t = Timer
    Do
        ex_CDPBrowser.TakeEvents    'パイプから非同期イベントを吸い上げる
        DoEvents

        Select Case mState
            Case "completed"
                printMsg info_, "Download completed in " & Format(Timer - t, "0.0") & "s. " _
                              & "file=" & mFilePath, FromProcedureName
                WaitCompleted = True
                Exit Function

            Case "canceled"
                printMsg WARN_, "Download was canceled by browser.", FromProcedureName
                WaitCompleted = False
                Exit Function
        End Select

        If (Timer - t) > TimeoutSec Then
            printMsg WARN_, "Timed out waiting for download. " _
                          & "State=" & mState & " Received=" & Format(mReceivedBytes / 1024, "0.0") & "KB", FromProcedureName
            WaitCompleted = False
            Exit Function
        End If

        ex_CDPBrowser.sleep 0.2    'CPU節約
    Loop

End Function

'***************************************************************************************************
'* 機能　　：ダウンロードした完了ファイルを別フォルダ・別名で移動保存します
'---------------------------------------------------------------------------------------------------
'* 返り値　：移動後のフル保存先パス / 失敗時は空文字
'* 引数　　：SaveDir        保存先フォルダパス（存在しない場合は自動作成）
'            NewFileName    保存するファイル名（省略時は SuggestedFilename をそのまま使用）
'---------------------------------------------------------------------------------------------------
'* 使い方  ：
'   If dw.WaitCompleted(30) Then
'       Dim saved As String: saved = dw.SaveAs("C:\output", "report_" & Format(Now,"yyyymmdd") & ".pdf")
'   End If
'***************************************************************************************************
Public Function SaveAs(SaveDir As String, Optional NewFileName As String = "") As String
    Const FromProcedureName As String = ThisClassName & ".SaveAs"


    If Len(mFilePath) = 0 Then
        printMsg WARN_, "SaveAs called before download completed. DownloadedFilePath is empty.", FromProcedureName
        SaveAs = ""
        Exit Function
    End If

    If Len(NewFileName) = 0 Then NewFileName = mSuggestedFilename
    If Len(NewFileName) = 0 Then NewFileName = Mid(mFilePath, InStrRev(mFilePath, "\") + 1)

    If Dir(SaveDir, vbDirectory) = "" Then MkDir SaveDir

    Dim destPath As String: destPath = SaveDir & "\" & NewFileName
    printMsg Debug_, "SaveAs: " & mFilePath & " → " & destPath, FromProcedureName

    On Error GoTo ErrHandler
    FileCopy mFilePath, destPath
    Kill mFilePath
    printMsg info_, "File saved as: " & destPath, FromProcedureName
    SaveAs = destPath
    Exit Function

ErrHandler:
    printMsg WARN_, "SaveAs failed. Err=" & Err.Number & " " & Err.Description, FromProcedureName
    SaveAs = ""

End Function

'***************************************************************************************************
'* 機能　　：内部状態をリセットします（WatchStart で自動呼び出し）
'***************************************************************************************************
Public Sub ResetState()
    mGuid = ""
    mSuggestedFilename = ""
    mDownloadUrl = ""
    mState = ""
    mFilePath = ""
    mReceivedBytes = 0
    mTotalBytes = 0
End Sub



'***************************************************************************************************
'                               ■■■ 内部処理 ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：フォルダ内の全ファイルを Dictionary にスナップショットします
'***************************************************************************************************
Private Sub SnapshotDirectory(DirPath As String, ByRef dic As Object)
    Dim f As String: f = Dir(DirPath & "\*.*")
    Do While f <> ""
        Dim fullPath As String: fullPath = DirPath & "\" & f
        If Not dic.Exists(LCase$(f)) Then
            dic.Add LCase$(f), fullPath
        End If
        f = Dir
    Loop
End Sub

'***************************************************************************************************
'* 機能　　：完了後にフォルダを再スキャンして、スナップショット差分から保存パスを特定します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Browser.downloadProgress の filePath がない場合のフォールバックとして使用します
'***************************************************************************************************
Private Function ResolveFilePathFromDir(SuggestedName As String) As String
    Const FromProcedureName As String = ThisClassName & ".ResolveFilePathFromDir"


    If Len(mWatchDir) = 0 Then
        printMsg WARN_, "WatchDir is not set. Cannot resolve file path.", FromProcedureName
        ResolveFilePathFromDir = ""
        Exit Function
    End If

    '--- 1. SuggestedFilename がそのままあればそれを返す ---
    Dim directPath As String: directPath = mWatchDir & "\" & SuggestedName
    If Dir(directPath) <> "" Then
        printMsg Debug_, "Resolved by SuggestedFilename: " & directPath, FromProcedureName
        ResolveFilePathFromDir = directPath
        Exit Function
    End If

    '--- 2. スナップショット差分から新規ファイルを検索 ---
    printMsg Debug_, "SuggestedFilename not found directly, scanning directory for new files...", FromProcedureName
    Dim f As String: f = Dir(mWatchDir & "\*.*")
    Do While f <> ""
        Dim candidate As String: candidate = mWatchDir & "\" & f
        Dim fLower As String: fLower = LCase$(f)

        ' 一時ファイルはスキップ
        If Not IsTemporaryFile(fLower) Then
            ' スナップショット時点に存在しなかった = 新規ファイル
            If Not mSnapshotFiles.Exists(fLower) Then
                printMsg Debug_, "New file found by directory diff: " & candidate, FromProcedureName
                ResolveFilePathFromDir = candidate
                Exit Function
            End If
        End If
        f = Dir
    Loop

    printMsg WARN_, "Could not resolve downloaded file path. SuggestedFilename=" & SuggestedName, FromProcedureName
    ResolveFilePathFromDir = ""

End Function

'***************************************************************************************************
'* 機能　　：ダウンロード中の一時ファイル（.crdownload / .part 等）か判定します
'***************************************************************************************************
Private Function IsTemporaryFile(fileNameLower As String) As Boolean
    IsTemporaryFile = (Right$(fileNameLower, 11) = ".crdownload") _
                   Or (Right$(fileNameLower, 5) = ".part") _
                   Or (Right$(fileNameLower, 4) = ".tmp")
End Function



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
    printMsg Debug_, "CDPexpansion_DownloadWatcher initialized.", FromProcedureName

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

Private Sub Class_Terminate()
    Const FromProcedureName As String = ThisClassName & ".Class_Terminate"
    On Error Resume Next
    EnableBrowserEvents = False
    printMsg Debug_, "CDPexpansion_DownloadWatcher terminated.", FromProcedureName
End Sub
