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
'            複数同時ダウンロードも含めて完了まで VBA 側で待機します。
'---------------------------------------------------------------------------------------------------
'* 基本的な使い方（単一DL）：
'    Dim dw As New CDPexpansion_DownloadWatcher
'    dw.Init browser
'    dw.WatchStart "C:\download_dir"
'    browser.getElementByID("dl-btn").click
'    If dw.WaitCompleted(TimeoutSec:=30) Then
'        Debug.Print "保存先: " & dw.DownloadedFilePath
'    End If
'---------------------------------------------------------------------------------------------------
'* 複数同時DLの使い方（新API）：
'    dw.WatchStart "C:\download_dir"
'    ' ← 複数のダウンロードをトリガー
'    If dw.WaitAllCompleted(TimeoutSec:=120) Then
'        dw.PrintSummary
'    End If
'---------------------------------------------------------------------------------------------------
'* 内部データ構造：
'    downloadInfos (Scripting.Dictionary)
'      Key   : Guid (String)
'      Value : Scripting.Dictionary {
'                KEY_WILL_BEGIN    → Dictionary  ← downloadWillBegin の params を丸ごと格納
'                KEY_PROGRESS      → Dictionary  ← downloadProgress の params を丸ごと格納（最新上書き）
'                KEY_IS_BROWSER    → Boolean     ← True=Browser / False=Page メソッド
'              }
'---------------------------------------------------------------------------------------------------
'* この設計の利点：
'    ① UDT や中間変換が不要 → mJson.ParseJson(EventRawJson)("params") をそのまま格納
'    ② CDP プロトコルが将来フィールドを追加しても自動対応
'    ③ GetWillBeginParams / GetProgressParams で生データへも直接アクセス可能
'---------------------------------------------------------------------------------------------------
'* 対応CDPイベント：
'    - Browser.downloadWillBegin / Page.downloadWillBegin   (ダウンロード開始通知)
'    - Browser.downloadProgress  / Page.downloadProgress    (進捗・完了・失敗通知)
'---------------------------------------------------------------------------------------------------
'* 注意事項：
'    - WatchStart の前にダウンロードをトリガーすると検出を逃す場合があります
'    - ダウンロード先フォルダは WatchStart が自動作成します
'    - .crdownload/.part 等の一時ファイルは完了とみなしません
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                           ■■■ フォルダ一括作成用Windows API ■■■
'***************************************************************************************************
Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As Long



'***************************************************************************************************
'           ■■■ Dictionary キー名 定数（タイプミス防止・補完サポート） ■■■
'***************************************************************************************************

'--- downloadInfos の各エントリ（第1層）のキー ---
Private Const KEY_WILL_BEGIN    As String = "WillBegin"     'downloadWillBegin の生params
Private Const KEY_PROGRESS      As String = "Progress"      'downloadProgress の生params（最新）
Private Const KEY_IS_BROWSER    As String = "isBrowserMethod" 'True=Browser / False=Page

'--- WillBegin params（第2層）のキー ←CDPの生フィールド名に合わせる ---
'https://chromedevtools.github.io/devtools-protocol/tot/Browser/#event-downloadWillBegin
'https://chromedevtools.github.io/devtools-protocol/tot/Page/#event-downloadWillBegin
Private Const WB_GUID           As String = "guid"
Private Const WB_URL            As String = "url"
Private Const WB_SUGGESTED      As String = "suggestedFilename"
Private Const WB_FRAME_ID       As String = "frameId"

'--- Progress params（第2層）のキー ←CDPの生フィールド名に合わせる ---
'https://chromedevtools.github.io/devtools-protocol/tot/Page/#event-downloadProgress
'https://chromedevtools.github.io/devtools-protocol/tot/Browser/#event-downloadProgress
Private Const PROG_GUID         As String = "guid"
Private Const PROG_STATE        As String = "state"
Private Const PROG_RECEIVED     As String = "receivedBytes"
Private Const PROG_TOTAL        As String = "totalBytes"
Private Const PROG_FILE_PATH    As String = "filePath"



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'内部参照用オブジェクト
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

'内部使用用
Private Const ThisClassName         As String = "CDPexpansion_DownloadWatcher"
Private mJson                       As New WebJsonConverter

'ステート管理
Private seeRawSendMsgDbg            As Boolean
Private EnableBrowserEvents         As Boolean
Private BrowserDomain               As Boolean    ' True="Browser"/False="Page" - WillBegin/Progress イベントのどちらを受けるか（両方受けることも可能）



'***************************************************************************************************
'* ★ 多重ダウンロード管理 Dictionary（2層構造）
'---------------------------------------------------------------------------------------------------
'* downloadInfos
'*   └─ Key: Guid (String)
'*   └─ Value: Scripting.Dictionary {
'*        KEY_WILL_BEGIN  → Scripting.Dictionary  ← WillBegin   の params を丸ごと格納
'*        KEY_PROGRESS    → Scripting.Dictionary  ← Progress    の params を丸ごと格納（最新上書き）
'*        KEY_IS_BROWSER  → Boolean
'*      }
'***************************************************************************************************
Private downloadInfos               As Scripting.Dictionary

'最後に開始したダウンロードの GUID（単一DL後方互換 API 向け）
Private mLastGuid                   As String

'監視設定
Private mWatchDir                   As String
Private mSnapshotFiles              As Dictionary   'WatchStart 時点のフォルダスナップショット



'***************************************************************************************************
'* 機能　　：非同期イベント処理を行うかのスイッチングです
'            この拡張機能の場合はさらに、ダウンロード監視の有効/無効も合わせて、切り替えます
'---------------------------------------------------------------------------------------------------
'* 引数　　：DownloadDir    ダウンロード先フォルダのフルパス （省略時は一般的なDownloadフォルダ）
'            AllowAndName   True: "allowAndName" / False: "allow"（通常はFalse）
'            UsePageDomain  True: Pageドメインのイベントを受ける / False: Browserドメインのイベントを受ける（通常はFalse）
'* 機能説明：呼ぶたびに downloadInfos をクリアし、フォルダスナップショットを更新します
'***************************************************************************************************
Public Property Let EnableEvents(Optional DownloadDir As String, Optional AllowAndName As Boolean, Optional UsePageDomain As Boolean, ByVal v As Boolean)
    Const FromProcedureName As String = ThisClassName & ".EnableEvents"


    '1. `DownloadDir`が省略時は、一般的なDownloadフォルダに指定
    If StrPtr(DownloadDir) = 0 Then DownloadDir = Environ("USERPROFILE") & "\Downloads"
    SHCreateDirectoryEx 0, DownloadDir, 0 'フォルダが存在しない場合は作成（存在する場合は何もしない）

    '2. クラスのイベント機能を有効にするか切り替え
    EnableBrowserEvents = v

    '3. 共通のパラメーターを付与
    Dim behavior As String
    '※`behavior`の場合、イベント発火が無効の場合は、`deny`として扱います
    If v Then
        behavior = IIf(AllowAndName, "allowAndName", "allow")
    Else
        behavior = "deny"
    End If

    Dim params As New Scripting.Dictionary
    params.Add "behavior", behavior
    params.Add "downloadPath", DownloadDir

    '4. method名に応じたパラメーター適用
    BrowserDomain = Not UsePageDomain
    Dim UserMethod As String: UserMethod = IIf(BrowserDomain, "Browser.setDownloadBehavior", "Page.setDownloadBehavior")

    'Browserドメインの場合は`eventsEnabled`も付与
    If BrowserDomain Then params.Add "eventsEnabled", v

    '5. CDPに設定を反映（イベント有効/無効の切り替え）
    invokeMethod UserMethod, params, True

    '6. 無効化する場合は、内部状態もクリア
    If v Then
        Set mSnapshotFiles = New Dictionary
        SnapshotDirectory DownloadDir, mSnapshotFiles
        printMsg Debug_, "Snapshot taken. Existing files: " & mSnapshotFiles.Count, FromProcedureName
    Else
        ClearDownloadHistory
        Set mSnapshotFiles = Nothing
    End If

    printMsg info_, "Event handling " & IIf(v, "enabled", "disabled") & ". " _
                  & "Method: " & UserMethod & " | behavior=" & behavior & " | DownloadDir=" & DownloadDir, FromProcedureName

End Property



'***************************************************************************************************
'                       ■■■ 公開プロパティ - 後方互換（単一DL用） ■■■
'***************************************************************************************************
'* 以下のプロパティはすべて「最後に開始したダウンロード」(mLastGuid) の情報を返します
'***************************************************************************************************
'* 機能　　：最後のダウンロードの状態 ("" / "inProgress" / "completed" / "canceled")
'***************************************************************************************************
Public Property Get state() As String
    state = GetState(mLastGuid)
End Property

'***************************************************************************************************
'* 機能　　：最後のダウンロードが完了しているか
'***************************************************************************************************
Public Property Get IsCompleted() As Boolean
    IsCompleted = (GetState(mLastGuid) = "completed")
End Property

'***************************************************************************************************
'* 機能　　：最後のダウンロードが開始済みか（downloadWillBegin 受信済みか）
'***************************************************************************************************
Public Property Get HasStarted() As Boolean
    HasStarted = (Len(mLastGuid) > 0)
End Property

'***************************************************************************************************
'* 機能　　：最後のダウンロードの保存先フルパス
'***************************************************************************************************
Public Property Get DownloadedFilePath() As String
    DownloadedFilePath = GetFilePath(mLastGuid)
End Property

'***************************************************************************************************
'* 機能　　：最後のダウンロードのブラウザ提案ファイル名
'***************************************************************************************************
Public Property Get SuggestedFilename() As String
    SuggestedFilename = GetWillBeginField(mLastGuid, WB_SUGGESTED)
End Property

'***************************************************************************************************
'* 機能　　：最後のダウンロード対象 URL
'***************************************************************************************************
Public Property Get DownloadUrl() As String
    DownloadUrl = GetWillBeginField(mLastGuid, WB_URL)
End Property

'***************************************************************************************************
'* 機能　　：最後のダウンロードの受信済みバイト数
'***************************************************************************************************
Public Property Get ReceivedBytes() As Double
    ReceivedBytes = GetProgressFieldDbl(mLastGuid, PROG_RECEIVED)
End Property

'***************************************************************************************************
'* 機能　　：最後のダウンロードの合計バイト数（0=不明）
'***************************************************************************************************
Public Property Get TotalBytes() As Double
    TotalBytes = GetProgressFieldDbl(mLastGuid, PROG_TOTAL)
End Property

'***************************************************************************************************
'* 機能　　：最後のダウンロードの進捗率 0.0?1.0（合計不明の場合は -1）
'***************************************************************************************************
Public Property Get Progress() As Double
    Progress = GetProgress(mLastGuid)
End Property



'***************************************************************************************************
'                       ■■■ 公開プロパティ - 多重DL管理（新API） ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：現在追跡しているダウンロードの件数
'***************************************************************************************************
Public Property Get DownloadCount() As Long
    DownloadCount = downloadInfos.Count
End Property

'***************************************************************************************************
'* 機能　　：完了したダウンロードの件数
'***************************************************************************************************
Public Function CompletedCount() As Long
    Dim k As Variant
    Dim cnt As Long
    For Each k In downloadInfos.keys
        If GetState(CStr(k)) = "completed" Then cnt = cnt + 1
    Next k
    CompletedCount = cnt
End Function

'***************************************************************************************************
'* 機能　　：進行中のダウンロードの件数
'***************************************************************************************************
Public Function InProgressCount() As Long
    Dim k As Variant
    Dim cnt As Long
    For Each k In downloadInfos.keys
        If GetState(CStr(k)) = "inProgress" Then cnt = cnt + 1
    Next k
    InProgressCount = cnt
End Function



'***************************************************************************************************
'              ■■■ メインとなる機能拡張部分 非同期イベント編 ■■■
'***************************************************************************************************
'* 機能　　：ダウンロード関連の CDP イベントを受け取り、downloadInfos に反映します
'---------------------------------------------------------------------------------------------------
'* 内部処理（WillBegin）：
'   mJson.ParseJson(EventRawJson)("params") を丸ごと KEY_WILL_BEGIN に格納
'* 内部処理（Progress）：
'   mJson.ParseJson(EventRawJson)("params") を丸ごと KEY_PROGRESS に格納（毎回最新で上書き）
'***************************************************************************************************
Private Sub ex_CDPBrowser_detectionCDPEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPBrowser_detectionCDPEvent"


    If EnableBrowserEvents Then
        printMsg Trace_, "Asynchronous event handling is enabled.", FromProcedureName
    Else
        printMsg Trace_, "Asynchronous event handling is disabled.", FromProcedureName
        Exit Sub
    End If

    Select Case methodName

        '--- ダウンロード開始 ---
        Case "Browser.downloadWillBegin", "Page.downloadWillBegin"
            printMsg Trace_, "`" & methodName & "` detected. Raw=" & EventRawJson, FromProcedureName

            '① CDPの生params を丸ごと取得 → UDT/中間変換なし
            Dim wbParams As Scripting.Dictionary
            Set wbParams = mJson.ParseJson(EventRawJson)("params")

            Dim guid As String: guid = CStr(wbParams(WB_GUID))

            '② エントリ作成（KEY_WILL_BEGIN に生params を格納）
            Dim entry As New Scripting.Dictionary
            entry.Add KEY_WILL_BEGIN, wbParams                                         '← 丸ごと格納！
            entry.Add KEY_IS_BROWSER, CBool(methodName = "Browser.downloadWillBegin")
            '   KEY_PROGRESS は最初の Progress イベントで追加

            '③ downloadInfos に登録
            If downloadInfos.Exists(guid) Then downloadInfos.Remove guid
            downloadInfos.Add guid, entry
            mLastGuid = guid

            printMsg info_, "[" & methodName & "] Download started. " _
                          & "guid=" & guid _
                          & " | file=" & CStr(wbParams(WB_SUGGESTED)) _
                          & " | url=" & CStr(wbParams(WB_URL)) _
                          & " | Tracking: " & downloadInfos.Count, FromProcedureName

            If seeRawSendMsgDbg Then
                PrintParamsDebug wbParams, methodName, FromProcedureName
            End If

        '--- ダウンロード進捗 ---
        Case "Browser.downloadProgress", "Page.downloadProgress"
            printMsg Trace_, "`" & methodName & "` detected. Raw=" & EventRawJson, FromProcedureName

            '① CDPの生params を丸ごと取得
            Dim pgParams As Scripting.Dictionary
            Set pgParams = mJson.ParseJson(EventRawJson)("params")

            Dim evGuid As String: evGuid = CStr(pgParams(PROG_GUID))

            '② downloadInfos に対象エントリがなければ新規作成（Progress が先に来た稀なケース）
            If Not downloadInfos.Exists(evGuid) Then
                Dim newEntry As New Scripting.Dictionary
                newEntry.Add KEY_IS_BROWSER, CBool(methodName = "Browser.downloadProgress")
                downloadInfos.Add evGuid, newEntry
                mLastGuid = evGuid
                printMsg WARN_, "Progress received before WillBegin. New entry created. guid=" & evGuid, FromProcedureName
            End If

            Dim inner As Scripting.Dictionary
            Set inner = downloadInfos(evGuid)

            '③ KEY_PROGRESS を最新の params で上書き（毎 Progress イベントで最新状態に更新）
            If inner.Exists(KEY_PROGRESS) Then inner.Remove KEY_PROGRESS
            inner.Add KEY_PROGRESS, pgParams   '← 丸ごと格納！

            '④ 状態別処理
            Dim evState As String: evState = CStr(pgParams(PROG_STATE))

            Select Case evState
                Case "inProgress"
                    Dim recv As Double: recv = CDbl(pgParams(PROG_RECEIVED))
                    Dim total As Double: total = CDbl(pgParams(PROG_TOTAL))
                    If total > 0 Then
                        printMsg Debug_, "[" & Left(evGuid, 8) & "...] " _
                                       & Format(recv / 1024 / 1024, "0.00") & " MB / " _
                                       & Format(total / 1024 / 1024, "0.00") & " MB " _
                                       & "(" & Format(recv / total * 100, "0.0") & "%)" _
                                       & " | In-progress: " & InProgressCount(), FromProcedureName
                    Else
                        printMsg Debug_, "[" & Left(evGuid, 8) & "...] " _
                                       & Format(recv / 1024 / 1024, "0.00") & " MB (total unknown)", FromProcedureName
                    End If

                Case "completed"
                    '--- filePath を確定（Browser メソッドは params に含まれる / Page メソッドは要解決） ---
                    Dim filePath As String: filePath = ""
                    If pgParams.Exists(PROG_FILE_PATH) Then
                        filePath = CStr(pgParams(PROG_FILE_PATH))
                    End If

                    If Len(filePath) = 0 Then
                        '--- Page メソッドなどで filePath がない場合 → フォルダ差分から特定 ---
                        Dim SuggestedName As String: SuggestedName = ""
                        If inner.Exists(KEY_WILL_BEGIN) Then
                            SuggestedName = CStr(inner(KEY_WILL_BEGIN)(WB_SUGGESTED))
                        End If
                        filePath = ResolveFilePathFromDir(SuggestedName)
                        '--- 解決した filePath を Progress params に追記（GetFilePath で取得可能にする） ---
                        pgParams.Add PROG_FILE_PATH, filePath
                    End If

                    printMsg info_, "[" & Left(evGuid, 8) & "...] completed! " _
                                  & "path=" & filePath _
                                  & " | Completed: " & CompletedCount() & "/" & downloadInfos.Count, FromProcedureName

                Case "canceled"
                    printMsg WARN_, "[" & Left(evGuid, 8) & "...] Download canceled.", FromProcedureName

                Case Else
                    printMsg Debug_, "[" & Left(evGuid, 8) & "...] Unknown state: " & evState, FromProcedureName
            End Select

        Case Else
            printMsg Trace_, "Event not handled by this class: " & methodName, FromProcedureName
            Exit Sub

    End Select

End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************
'* 機能　　：指定した GUID のダウンロードが完了するまで待機します（単一DL用）
'---------------------------------------------------------------------------------------------------
'* 返り値　：True=完了 / False=タイムアウト or キャンセル
'* 引数　　：Guid        対象の GUID（省略時は最後に開始したダウンロード）
'            TimeoutSec  最大待機時間（秒）
'---------------------------------------------------------------------------------------------------
'* 2フェーズ設計：
'*   Phase 1 : downloadWillBegin が来るまで待つ（fetch型サイトはサーバー生成に時間がかかるため）
'*   Phase 2 : 受信した GUID の state が "completed" になるまで待つ
'***************************************************************************************************
Public Function WaitCompleted(Optional guid As String, Optional TimeoutSec As Long = 60) As Boolean
    Const FromProcedureName As String = ThisClassName & ".WaitCompleted"


    If Len(guid) = 0 Then guid = mLastGuid

    Dim t As Double: t = Timer

    '--- Phase 1: downloadWillBegin（GUID）が来るまで待機 ---
    '    fetch → Blob → <a click> 型のサイトでは、ボタンクリック後しばらく経って
    '    イベントが発火するため、mLastGuid が空のまま待つ必要がある
    If Len(guid) = 0 Then
        printMsg info_, "Waiting for download to begin... (timeout: " & TimeoutSec & "s)", FromProcedureName
        Do
            ex_CDPBrowser.TakeEvents
            DoEvents

            If Len(mLastGuid) > 0 Then
                guid = mLastGuid
                printMsg info_, "Download began! guid=" & Left(guid, 8) & "... " _
                              & "(elapsed: " & Format(Timer - t, "0.0") & "s)", FromProcedureName
                Exit Do
            End If

            If (Timer - t) > TimeoutSec Then
                printMsg WARN_, "Timed out waiting for download to begin. " _
                              & "No downloadWillBegin event was received.", FromProcedureName
                WaitCompleted = False
                Exit Function
            End If

            ex_CDPBrowser.sleep 0.2
        Loop
    End If

    '--- Phase 2: 受信した GUID の完了を待つ（残り時間内で） ---
    Dim remaining As Double: remaining = TimeoutSec - (Timer - t)
    printMsg info_, "Waiting for completion. guid=" & Left(guid, 8) & _
                  "... (remaining: " & Format(remaining, "0") & "s)", FromProcedureName

    Do
        ex_CDPBrowser.TakeEvents
        DoEvents

        Select Case GetState(guid)
            Case "completed"
                printMsg info_, "Download completed in " & Format(Timer - t, "0.0") & "s. " _
                              & "file=" & GetFilePath(guid), FromProcedureName
                WaitCompleted = True
                Exit Function
            Case "canceled"
                printMsg WARN_, "Download was canceled by browser.", FromProcedureName
                WaitCompleted = False
                Exit Function
        End Select

        If (Timer - t) > TimeoutSec Then
            printMsg WARN_, "Timed out. State=" & GetState(guid), FromProcedureName
            WaitCompleted = False
            Exit Function
        End If

        ex_CDPBrowser.sleep 0.2
    Loop

End Function

'***************************************************************************************************
'* 機能　　：追跡中の全ダウンロードが完了するまで待機します（多重DL用）
'---------------------------------------------------------------------------------------------------
'* 返り値　：True=全件完了 / False=タイムアウト
'* 引数　　：ExpectedCount  完了待ちするダウンロードの期待件数（0=自動判定）
'            TimeoutSec     最大待機時間（秒）
'---------------------------------------------------------------------------------------------------
'* 2フェーズ設計：
'*   Phase 1 : 少なくとも 1件の downloadWillBegin が来るまで待つ
'*   Phase 2 : 全件（ExpectedCount件 or 件数確定後）が完了するまで待つ
'***************************************************************************************************
Public Function WaitAllCompleted(Optional ExpectedCount As Long, Optional TimeoutSec As Long = 120) As Boolean
    Const FromProcedureName As String = ThisClassName & ".WaitAllCompleted"


    Dim t As Double: t = Timer

    '--- Phase 1: 少なくとも1件の downloadWillBegin が来るまで待つ ---
    If downloadInfos.Count = 0 Then
        printMsg info_, "Waiting for downloads to begin... (timeout: " & TimeoutSec & "s)", FromProcedureName
        Do
            ex_CDPBrowser.TakeEvents
            DoEvents

            If downloadInfos.Count > 0 Then
                printMsg info_, "First download detected! " _
                              & "(elapsed: " & Format(Timer - t, "0.0") & "s)", FromProcedureName
                Exit Do
            End If

            If (Timer - t) > TimeoutSec Then
                printMsg WARN_, "Timed out waiting for any download to begin.", FromProcedureName
                WaitAllCompleted = False
                Exit Function
            End If

            ex_CDPBrowser.sleep 0.2
        Loop
    End If

    printMsg info_, "Waiting for ALL downloads. " _
                  & "Tracking: " & downloadInfos.Count _
                  & IIf(ExpectedCount > 0, " (expecting " & ExpectedCount & ")", "") _
                  & " | remaining: " & Format(TimeoutSec - (Timer - t), "0") & "s", FromProcedureName

    '--- Phase 2: 全件完了を待つ ---
    Do
        ex_CDPBrowser.TakeEvents
        DoEvents

        '件数の確定チェック（ExpectedCount 指定時）
        Dim countOk As Boolean
        countOk = (ExpectedCount = 0) Or (downloadInfos.Count >= ExpectedCount)

        If countOk And downloadInfos.Count > 0 Then
            Dim allDone As Boolean: allDone = True
            Dim k As Variant
            For Each k In downloadInfos.keys
                Dim s As String: s = GetState(CStr(k))
                If s <> "completed" And s <> "canceled" Then
                    allDone = False
                    Exit For
                End If
            Next k

            If allDone Then
                printMsg info_, "All " & downloadInfos.Count & " downloads finished in " _
                              & Format(Timer - t, "0.0") & "s. " _
                              & "Completed: " & CompletedCount() & " | Canceled: " _
                              & (downloadInfos.Count - CompletedCount()), FromProcedureName
                WaitAllCompleted = True
                Exit Function
            End If
        End If

        If (Timer - t) > TimeoutSec Then
            printMsg WARN_, "Timed out. Completed: " & CompletedCount() _
                          & "/" & downloadInfos.Count, FromProcedureName
            WaitAllCompleted = False
            Exit Function
        End If

        ex_CDPBrowser.sleep 0.2
    Loop

End Function

'***************************************************************************************************
'* 機能　　：全ダウンロードのサマリーをイミディエイトウィンドウに出力します
'***************************************************************************************************
Public Sub PrintSummary()
    Const FromProcedureName As String = ThisClassName & ".PrintSummary"


    printMsg info_, String(80, "="), FromProcedureName
    printMsg info_, "Download Summary - " _
                  & downloadInfos.Count & " total | " _
                  & CompletedCount() & " completed | " _
                  & InProgressCount() & " in-progress", FromProcedureName
    printMsg info_, String(80, "-"), FromProcedureName

    Dim k As Variant
    Dim i As Long
    For Each k In downloadInfos.keys
        i = i + 1
        Dim guid As String: guid = CStr(k)
        Dim entryD As Scripting.Dictionary: Set entryD = downloadInfos(guid)

        '--- WillBegin 情報 ---
        Dim FileName As String
        If entryD.Exists(KEY_WILL_BEGIN) Then
            FileName = CStr(entryD(KEY_WILL_BEGIN)(WB_SUGGESTED))
        End If

        '--- Progress 情報 ---
        Dim stateStr As String
        Dim recvMB As Double
        Dim totalMB As Double
        Dim pctStr As String
        Dim pathStr As String

        If entryD.Exists(KEY_PROGRESS) Then
            Dim progD As Scripting.Dictionary: Set progD = entryD(KEY_PROGRESS)
            stateStr = CStr(progD(PROG_STATE))
            recvMB = CDbl(progD(PROG_RECEIVED)) / 1024 / 1024
            totalMB = CDbl(progD(PROG_TOTAL)) / 1024 / 1024
            If totalMB > 0 Then pctStr = Format(recvMB / totalMB * 100, "0.0") & "%"
            If progD.Exists(PROG_FILE_PATH) Then pathStr = CStr(progD(PROG_FILE_PATH))
        End If

        printMsg info_, "  [" & i & "] " _
                      & "state=" & Format(stateStr, "@@@@@@@@@@@") & " | " _
                      & "file=" & FileName & " | " _
                      & Format(recvMB, "0.00") & "/" & Format(totalMB, "0.00") & "MB" _
                      & " (" & pctStr & ") | " _
                      & "path=" & pathStr, FromProcedureName
    Next
    printMsg info_, String(80, "="), FromProcedureName

End Sub



'***************************************************************************************************
'                       ■■■ 公開メソッド - GUID 指定の個別取得 ■■■
'***************************************************************************************************
'* 機能　　：指定 GUID のエントリ辞書を返します（WillBegin/Progress/isBrowserMethod を含む）
'---------------------------------------------------------------------------------------------------
'* 使い方  ：
'    Dim info As Scripting.Dictionary: Set info = dw.GetInfo(guid)
'    Debug.Print info("WillBegin")("suggestedFilename")
'    Debug.Print info("Progress")("state")
'***************************************************************************************************
Public Function GetInfo(guid As String) As Scripting.Dictionary
    If downloadInfos.Exists(guid) Then
        Set GetInfo = downloadInfos(guid)
    Else
        Set GetInfo = Nothing
    End If
End Function

'***************************************************************************************************
'* 機能　　：指定 GUID の WillBegin params を返します（CDPの生データ）
'***************************************************************************************************
Public Function GetWillBeginParams(guid As String) As Scripting.Dictionary
    If Not downloadInfos.Exists(guid) Then Set GetWillBeginParams = Nothing: Exit Function
    Dim entryD As Scripting.Dictionary: Set entryD = downloadInfos(guid)
    If Not entryD.Exists(KEY_WILL_BEGIN) Then Set GetWillBeginParams = Nothing: Exit Function
    Set GetWillBeginParams = entryD(KEY_WILL_BEGIN)
End Function

'***************************************************************************************************
'* 機能　　：指定 GUID の Progress params（最新）を返します（CDPの生データ）
'***************************************************************************************************
Public Function GetProgressParams(guid As String) As Scripting.Dictionary
    If Not downloadInfos.Exists(guid) Then Set GetProgressParams = Nothing: Exit Function
    Dim entryD As Scripting.Dictionary: Set entryD = downloadInfos(guid)
    If Not entryD.Exists(KEY_PROGRESS) Then Set GetProgressParams = Nothing: Exit Function
    Set GetProgressParams = entryD(KEY_PROGRESS)
End Function

'***************************************************************************************************
'* 機能　　：全 GUID のコレクションを返します
'***************************************************************************************************
Public Function GetAllGuids() As Collection
    Dim col As New Collection
    Dim k As Variant
    For Each k In downloadInfos.keys
        col.Add k
    Next k
    Set GetAllGuids = col
End Function

'***************************************************************************************************
'* 機能　　：指定 GUID の状態を返します
'***************************************************************************************************
Public Function GetState(guid As String) As String
    If Not downloadInfos.Exists(guid) Then Exit Function
    Dim entryD As Scripting.Dictionary: Set entryD = downloadInfos(guid)
    If Not entryD.Exists(KEY_PROGRESS) Then Exit Function
    GetState = CStr(entryD(KEY_PROGRESS)(PROG_STATE))
End Function

'***************************************************************************************************
'* 機能　　：指定 GUID の保存先ファイルパスを返します
'***************************************************************************************************
Public Function GetFilePath(guid As String) As String
    If Not downloadInfos.Exists(guid) Then Exit Function
    Dim entryD As Scripting.Dictionary: Set entryD = downloadInfos(guid)
    If Not entryD.Exists(KEY_PROGRESS) Then Exit Function
    Dim progD As Scripting.Dictionary: Set progD = entryD(KEY_PROGRESS)
    If Not progD.Exists(PROG_FILE_PATH) Then Exit Function
    GetFilePath = CStr(progD(PROG_FILE_PATH))
End Function

'***************************************************************************************************
'* 機能　　：指定 GUID の進捗率を 0.0~1.0 で返します（合計不明の場合は -1）
'***************************************************************************************************
Public Function GetProgress(guid As String) As Double
    GetProgress = -1
    If Not downloadInfos.Exists(guid) Then Exit Function
    Dim entryD As Scripting.Dictionary: Set entryD = downloadInfos(guid)
    If Not entryD.Exists(KEY_PROGRESS) Then Exit Function
    Dim progD As Scripting.Dictionary: Set progD = entryD(KEY_PROGRESS)
    If Not progD.Exists(PROG_TOTAL) Then Exit Function
    Dim total As Double: total = CDbl(progD(PROG_TOTAL))
    If total <= 0 Then Exit Function
    GetProgress = CDbl(progD(PROG_RECEIVED)) / total
End Function

'***************************************************************************************************
'* 機能　　：ダウンロード完了ファイルを別名・別フォルダへ移動保存します
'---------------------------------------------------------------------------------------------------
'* 引数　　：SaveDir      保存先フォルダ
'            NewFileName  保存するファイル名（省略時は SuggestedFilename を使用）
'            Guid         対象の GUID（省略時は最後に開始したダウンロード）
'***************************************************************************************************
Public Function SaveAs(saveDir As String, Optional NewFileName As String = "", Optional guid As String = "") As String
    Const FromProcedureName As String = ThisClassName & ".SaveAs"


    If Len(guid) = 0 Then guid = mLastGuid

    Dim filePath As String: filePath = GetFilePath(guid)
    If Len(filePath) = 0 Then
        printMsg WARN_, "SaveAs: No completed file. guid=" & guid & " Call WaitCompleted first.", FromProcedureName
        SaveAs = ""
        Exit Function
    End If

    If Len(NewFileName) = 0 Then
        NewFileName = GetWillBeginField(guid, WB_SUGGESTED)
        If Len(NewFileName) = 0 Then NewFileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    End If

    If Dir(saveDir, vbDirectory) = "" Then MkDir saveDir

    Dim destPath As String: destPath = saveDir & "\" & NewFileName
    printMsg Debug_, "SaveAs: " & filePath & " → " & destPath, FromProcedureName

    On Error GoTo ErrHandler
    FileCopy filePath, destPath
    Kill filePath
    printMsg info_, "File saved as: " & destPath, FromProcedureName
    SaveAs = destPath
    Exit Function

ErrHandler:
    printMsg WARN_, "SaveAs failed. Err=" & Err.Number & " " & Err.Description, FromProcedureName
    SaveAs = ""

End Function

'***************************************************************************************************
'* 機能　　：追跡中のダウンロード履歴をすべてクリアします（WatchStart でも自動呼び出し）
'***************************************************************************************************
Public Sub ClearDownloadHistory()
    Const FromProcedureName As String = ThisClassName & ".ClearDownloadHistory"


    Dim prev As Long: prev = downloadInfos.Count
    Set downloadInfos = New Scripting.Dictionary
    mLastGuid = ""
    If prev > 0 Then
        printMsg Debug_, "Download history cleared. (" & prev & " entries removed)", FromProcedureName
    End If

End Sub

'***************************************************************************************************
'* 機能　　：後方互換 - ClearDownloadHistory の旧名称
'***************************************************************************************************
Public Sub ResetState()
    ClearDownloadHistory
End Sub

'***************************************************************************************************
'* 機能　　：ネットワーク速度を意図的に制限します（テスト・デバッグ用）
'---------------------------------------------------------------------------------------------------
'* 引数　　：DownloadKBps  ダウンロード速度の上限（KB/s）デフォルト: 300 KB/s
'---------------------------------------------------------------------------------------------------
'* 速度の目安（ファイルサイズと予想所要時間）：
'*    100 KB/s :  5MB - 50秒 / 50MB - 8.5分
'*    300 KB/s :  5MB - 17秒 / 50MB - 2.8分
'*    500 KB/s :  5MB - 10秒 / 50MB - 1.7分
'*   1000 KB/s :  5MB -  5秒 / 50MB - 50秒
'***************************************************************************************************
Public Sub ThrottleNetwork(Optional DownloadKBps As Long = 300)
    Const FromProcedureName As String = ThisClassName & ".ThrottleNetwork"


    printMsg info_, "Throttling network to " & DownloadKBps & " KB/s...", FromProcedureName

    invokeMethod "Network.enable"

    Dim params As New Scripting.Dictionary
    params.Add "offline", False
    params.Add "latency", CLng(20)
    params.Add "downloadThroughput", CLng(DownloadKBps * 1024)
    params.Add "uploadThroughput", CLng(500& * 1024)

    invokeMethod "Network.emulateNetworkConditions", params
    printMsg info_, "Network throttled to " & DownloadKBps & " KB/s " _
                  & "(" & Format(DownloadKBps / 1024#, "0.0") & " Mbps approx.)", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：ネットワーク速度制限を解除します
'***************************************************************************************************
Public Sub UnthrottleNetwork()
    Const FromProcedureName As String = ThisClassName & ".UnthrottleNetwork"


    Dim params As New Scripting.Dictionary
    params.Add "offline", False
    params.Add "latency", CLng(0)
    params.Add "downloadThroughput", CLng(-1)
    params.Add "uploadThroughput", CLng(-1)

    invokeMethod "Network.emulateNetworkConditions", params
    printMsg info_, "Network throttle removed. Full speed restored.", FromProcedureName

End Sub



'***************************************************************************************************
'                               ■■■ 内部処理 ■■■
'***************************************************************************************************

'***************************************************************************************************
'* 機能　　：WillBegin params の指定フィールド値を文字列で返します
'***************************************************************************************************
Private Function GetWillBeginField(guid As String, FieldKey As String) As String
    If Not downloadInfos.Exists(guid) Then Exit Function
    Dim entryD As Scripting.Dictionary: Set entryD = downloadInfos(guid)
    If Not entryD.Exists(KEY_WILL_BEGIN) Then Exit Function
    Dim wbD As Scripting.Dictionary: Set wbD = entryD(KEY_WILL_BEGIN)
    If Not wbD.Exists(FieldKey) Then Exit Function
    GetWillBeginField = CStr(wbD(FieldKey))
End Function

'***************************************************************************************************
'* 機能　　：Progress params の指定フィールド値を Double で返します（なければ 0）
'***************************************************************************************************
Private Function GetProgressFieldDbl(guid As String, FieldKey As String) As Double
    If Not downloadInfos.Exists(guid) Then Exit Function
    Dim entryD As Scripting.Dictionary: Set entryD = downloadInfos(guid)
    If Not entryD.Exists(KEY_PROGRESS) Then Exit Function
    Dim progD As Scripting.Dictionary: Set progD = entryD(KEY_PROGRESS)
    If Not progD.Exists(FieldKey) Then Exit Function
    GetProgressFieldDbl = CDbl(progD(FieldKey))
End Function

'***************************************************************************************************
'* 機能　　：フォルダ内の全ファイルを Dictionary にスナップショットします
'***************************************************************************************************
Private Sub SnapshotDirectory(DirPath As String, ByRef dic As Object)
    Dim f As String: f = Dir(DirPath & "\*.*")
    Do While f <> ""
        If Not dic.Exists(LCase$(f)) Then
            dic.Add LCase$(f), DirPath & "\" & f
        End If
        f = Dir
    Loop
End Sub

'***************************************************************************************************
'* 機能　　：完了後にフォルダ差分からダウンロードファイルのパスを特定します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Browser.downloadProgress に filePath がない場合（Page メソッド等）のフォールバック
'***************************************************************************************************
Private Function ResolveFilePathFromDir(SuggestedName As String) As String
    Const FromProcedureName As String = ThisClassName & ".ResolveFilePathFromDir"


    If Len(mWatchDir) = 0 Then
        printMsg WARN_, "WatchDir is not set.", FromProcedureName
        Exit Function
    End If

    '--- 1. SuggestedFilename がそのままあればそれを返す ---
    Dim directPath As String: directPath = mWatchDir & "\" & SuggestedName
    If Len(SuggestedName) > 0 And Dir(directPath) <> "" Then
        printMsg Debug_, "Resolved by SuggestedFilename: " & directPath, FromProcedureName
        ResolveFilePathFromDir = directPath
        Exit Function
    End If

    '--- 2. スナップショット差分から新規ファイルを検索 ---
    printMsg Debug_, "Scanning directory for new files (snapshot diff)...", FromProcedureName
    Dim f As String: f = Dir(mWatchDir & "\*.*")
    Do While f <> ""
        Dim fLower As String: fLower = LCase$(f)
        If Not IsTemporaryFile(fLower) Then
            If Not mSnapshotFiles.Exists(fLower) Then
                Dim newPath As String: newPath = mWatchDir & "\" & f
                printMsg Debug_, "New file found by directory diff: " & newPath, FromProcedureName
                ResolveFilePathFromDir = newPath
                Exit Function
            End If
        End If
        f = Dir
    Loop

    printMsg WARN_, "Could not resolve downloaded file path. SuggestedFilename=" & SuggestedName, FromProcedureName

End Function

'***************************************************************************************************
'* 機能　　：ダウンロード中の一時ファイルか判定します
'***************************************************************************************************
Private Function IsTemporaryFile(fileNameLower As String) As Boolean
    IsTemporaryFile = (Right$(fileNameLower, 11) = ".crdownload") _
                   Or (Right$(fileNameLower, 5) = ".part") _
                   Or (Right$(fileNameLower, 4) = ".tmp")
End Function

'***************************************************************************************************
'* 機能　　：params の全キーをデバッグ出力します（seeRawSendMsgDbg=True のとき）
'***************************************************************************************************
Private Sub PrintParamsDebug(params As Scripting.Dictionary, methodName As String, FromProcedureName As String)
    printMsg Debug_, String(60, "-"), FromProcedureName
    printMsg Debug_, "  [" & methodName & "] params:", FromProcedureName
    Dim pk As Variant
    For Each pk In params.keys
        printMsg Debug_, "    " & CStr(pk) & " = " & CStr(params(pk)), FromProcedureName
    Next pk
    printMsg Debug_, String(60, "-"), FromProcedureName
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
    printMsg Debug_, "CDPexpansion_DownloadWatcher initialized.", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：ワークシートから、必要な設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'ブラウザ非同期処理を有効化
    EnableBrowserEvents = True
    Set downloadInfos = New Scripting.Dictionary

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_StartBrowser
        seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value    'rawJSONをデバッグ出力するか
    End With

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    EnableBrowserEvents = False
    Set downloadInfos = Nothing
End Sub
