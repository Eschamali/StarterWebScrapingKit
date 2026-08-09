VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_DownloadWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'                       ファイルのダウンロード監視のアシストを行います
'   `Browser.downloadProgress`/`Browser.downloadWillBegin`の監視を利用して、スマートなDL監視を行います
'
'           このClassでは`Browser`のみ扱います。最新のブラウザではこれが推奨されています。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'ブラウザ制御オブジェクト
Private WithEvents ex_CDPCore       As CDPCore      'イベント受け取りオブジェクト
Attribute ex_CDPCore.VB_VarHelpID = -1
Private ex_CDPBrowser               As CDPBrowser   '継承元のブラウザクラス

'内部使用用
Private Const ThisClassName         As String = "exCDP_DownloadWatcher"

'基本ステート管理
Private defaultBrowserContextId     As String       '継承時のブラウザコンテキスト
Private seeRawSendMsgDbg            As Boolean      'デバッグメッセージ表示用

'ダウンロード系のstatus管理
'※`guid`をKeyに管理します
Private DownloadPathList    As Dictionary   '各`BrowserContextId`ごとの「downloadPath」情報
Private DownloadProgress    As Dictionary   'ダウンロード進捗情報
Private DownloadWillBegin   As Dictionary   'ダウンロード開始情報



'***************************************************************************************************
'                                   ■■■ 各種構造体 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] ダウンロード時の挙動設定一式
'---------------------------------------------------------------------------------------------------'
Public Enum setDownloadBehavior_behavior
    behavior_deny
    behavior_allow
    behavior_allowAndName
    behavior_Default
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] `downloadWillBegin`のparameter取得部分
'---------------------------------------------------------------------------------------------------'
Public Enum downloadWillBegin_parameters
    downloadWillBegin_frameID
    downloadWillBegin_Url
    downloadWillBegin_suggestedFilename
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] `downloadProgress`のparameter取得部分
'---------------------------------------------------------------------------------------------------'
Public Enum downloadProgress_parameters
    downloadProgress_totalBytes
    downloadProgress_receivedBytes
    downloadProgress_state
    downloadProgress_filePath
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] `downloadProgress`の`state`部分
'---------------------------------------------------------------------------------------------------'
Public Enum downloadProgress_params_state
    state_inProgress
    state_completed
    state_canceled
    state_willBegin   '`downloadWillBegin`は来たが、`downloadProgress`が来てないパターン
End Enum



'***************************************************************************************************
'                       ■■■ Enum → 文字列 変換プロシージャ ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] ダウンロード時の挙動設定の種類を文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringSetDownloadBehavior_behavior(param As setDownloadBehavior_behavior) As String
    Select Case param
        Case setDownloadBehavior_behavior.behavior_deny:         EnumToStringSetDownloadBehavior_behavior = "deny"
        Case setDownloadBehavior_behavior.behavior_allow:        EnumToStringSetDownloadBehavior_behavior = "allow"
        Case setDownloadBehavior_behavior.behavior_allowAndName: EnumToStringSetDownloadBehavior_behavior = "allowAndName"
        Case setDownloadBehavior_behavior.behavior_Default:      EnumToStringSetDownloadBehavior_behavior = "default"
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] `downloadWillBegin`のparameter部分を文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringDownloadWillBegin_parameters(param As downloadWillBegin_parameters) As String
    Select Case param
        Case downloadWillBegin_parameters.downloadWillBegin_frameID:            EnumToStringDownloadWillBegin_parameters = "frameId"
        Case downloadWillBegin_parameters.downloadWillBegin_Url:                EnumToStringDownloadWillBegin_parameters = "url"
        Case downloadWillBegin_parameters.downloadWillBegin_suggestedFilename:  EnumToStringDownloadWillBegin_parameters = "suggestedFilename"  '`behavior:=allowAndName`にした場合は、これを参考に元のファイル名に復元可能です
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] `downloadProgress`のparameter部分を文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringDownloadProgress_parameters(param As downloadProgress_parameters) As String
    Select Case param
        Case downloadProgress_parameters.downloadProgress_totalBytes:    EnumToStringDownloadProgress_parameters = "totalBytes"
        Case downloadProgress_parameters.downloadProgress_receivedBytes: EnumToStringDownloadProgress_parameters = "receivedBytes"
        Case downloadProgress_parameters.downloadProgress_state:         EnumToStringDownloadProgress_parameters = "state"
        Case downloadProgress_parameters.downloadProgress_filePath:      EnumToStringDownloadProgress_parameters = "filePath"
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] `downloadProgress - state`部分を文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringDownloadProgressState(param As downloadProgress_params_state) As String
    Select Case param
        Case downloadProgress_params_state.state_inProgress: EnumToStringDownloadProgressState = "inProgress"
        Case downloadProgress_params_state.state_completed:  EnumToStringDownloadProgressState = "completed"
        Case downloadProgress_params_state.state_canceled:   EnumToStringDownloadProgressState = "canceled"
        Case downloadProgress_params_state.state_willBegin:  EnumToStringDownloadProgressState = vbNullString
    End Select
End Function



'***************************************************************************************************
'                               ■■■ 設定プロパティ ■■■
'***************************************************************************************************
'* 機能　　：ファイルをダウンロードする際の挙動を設定/更新をします
'---------------------------------------------------------------------------------------------------
'* Let引数 ：下記の引数設定に紐づく、ダウンロードイベント処理を有効/無効化するか？
'* 引数　　：behavior           ダウンロードの動作
'            downloadPath       ダウンロードしたファイルを保存するフルフォルダパス
'            browserContextId   `Target.createBrowserContext`で得た`browserContextId`を指定します。省略した場合、デフォルトのブラウザコンテキストが使用されます。
'                               複数アカウントで同一サイトに対するDL作業に重宝します。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://chromedevtools.github.io/devtools-protocol/tot/Browser/#method-setDownloadBehavior
'* 注意事項：`browserContextId`に渡せる引数は、`Target.createBrowserContext`で得た`browserContextId`のみです
'***************************************************************************************************
Public Property Let setDownloadBehavior(behavior As setDownloadBehavior_behavior, downloadPath As String, Optional browserContextId As String, eventsEnabled As Boolean)
    Const FromProcedureName As String = ThisClassName & ".setDownloadBehavior - Let"


    '1. CDPパラメーター構築
    Dim CDPparams As New Dictionary
    CDPparams.Add "behavior", EnumToStringSetDownloadBehavior_behavior(behavior)
    If StrPtr(browserContextId) Then CDPparams.Add "browserContextId", browserContextId
    CDPparams.Add "downloadPath", downloadPath
    CDPparams.Add "eventsEnabled", eventsEnabled

    '2. 実行
    ExecuteCDP "Browser.setDownloadBehavior", CDPparams

    '3. 記録用に、`browserContextId`省略時、継承時のブラウザコンテキストを適用
    If StrPtr(browserContextId) = 0 Then browserContextId = defaultBrowserContextId

    '4. ダウンロードパス設定情報の記録/更新
    '4-1.登録済みか？
    If DownloadPathList.Exists(browserContextId) Then
        If eventsEnabled Then
            '設定の更新
            DownloadPathList(browserContextId) = downloadPath
            printMsg info_, "I updated the download monitoring settings. browserContextId: " & browserContextId, FromProcedureName
        Else
            '設定の削除
            DownloadPathList.Remove (browserContextId)
            printMsg info_, "Download monitoring has been stopped. browserContextId: " & browserContextId, FromProcedureName
        End If
    Else
        '新規登録
        If eventsEnabled Then DownloadPathList.Add browserContextId, downloadPath: printMsg info_, "Start download monitoring. browserContextId: " & browserContextId, FromProcedureName
    End If

    '5. 設定情報を表示
    If seeRawSendMsgDbg Then printMsg info_, "setDownloadBehavior - parameters: " & WebJsonConverter.serialize(CDPparams), FromProcedureName
End Property



'***************************************************************************************************
'                  ■■■ ダウンロード情報取得プロパティ - Rawバージョン ■■■
'***************************************************************************************************
'* 機能　　：`guid`を基に、`downloadWillBegin`情報を取り出します
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に応じた`downloadWillBegin`の`parameters`情報。存在しない場合は、`vbnullstring`で返ります
'* 引数　　：guid   `DownloadGuidList`メソッドから得た`GUID`
'            params `downloadWillBegin`のparameters - Enum
'---------------------------------------------------------------------------------------------------
'* 機能説明：https://chromedevtools.github.io/devtools-protocol/tot/Browser/#event-downloadWillBegin
'***************************************************************************************************
Property Get DownloadWillBeginInfo(guid As String, params As downloadWillBegin_parameters) As String
    If DownloadWillBegin.Exists(guid) Then DownloadWillBeginInfo = DownloadWillBegin(guid).StringKey(EnumToStringDownloadWillBegin_parameters(params))
End Property

'***************************************************************************************************
'* 機能　　：`guid`を基に、`downloadProgress`情報を取り出します
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に応じた`downloadProgress`の`parameters`情報
'* 引数　　：guid   `DownloadGuidList`メソッドから得た`GUID`
'            params `downloadProgress`のparameters - Enum
'---------------------------------------------------------------------------------------------------
'* 機能説明：https://chromedevtools.github.io/devtools-protocol/tot/Browser/#event-downloadProgress
'* 注意事項：`downloadWillBegin`は来たが、`downloadProgress`が一度も来てない場合は、`vbNull`で返ります
'***************************************************************************************************
Property Get DownloadProgressInfo(guid As String, params As downloadProgress_parameters)
    '1. `downloadWillBegin`は来たが、`downloadProgress`が一度も来てない場合を確認
    If DownloadProgress.Exists(guid) Then
        '2. parameterに応じて、型指定させて取出し
        Select Case params
            Case downloadProgress_parameters.downloadProgress_totalBytes, downloadProgress_parameters.downloadProgress_receivedBytes
                DownloadProgressInfo = DownloadProgress(guid).NumberKey(EnumToStringDownloadProgress_parameters(params))

            Case downloadProgress_parameters.downloadProgress_state, downloadProgress_parameters.downloadProgress_filePath
                DownloadProgressInfo = DownloadProgress(guid).StringKey(EnumToStringDownloadProgress_parameters(params))
        End Select
    Else
        '2. `downloadProgress`が一度も来てない場合
        DownloadProgressInfo = vbEmpty
    End If
End Property



'***************************************************************************************************
'                 ■■■ ダウンロード情報取得プロパティ - 便利メソッド版 ■■■
'***************************************************************************************************
'* 機能　　：現時点で届いてるダウンロード`guid`リストを返却します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`guid`のある1次元配列
'---------------------------------------------------------------------------------------------------
'* 機能説明：`downloadWillBegin`が届いた順に並んでいます。各種ダウンロード情報を参照するにはまず、これを使います
'* 注意事項：要素は`0`始まりです
'***************************************************************************************************
Property Get DownloadGuidList()
    DownloadGuidList = DownloadWillBegin.keys
End Property

'***************************************************************************************************
'* 機能　　：現時点で届いてるダウンロード`guid`数を返却します
'---------------------------------------------------------------------------------------------------
'* 返り値　：ダウンロード監視中の`guid`数
'***************************************************************************************************
Property Get DownloadGuidCount() As Long
    DownloadGuidCount = DownloadWillBegin.Count
End Property

'***************************************************************************************************
'* 機能　　：`downloadProgress - state`の確認をします
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に応じた`downloadProgress - state`の数
'***************************************************************************************************
Property Get DownloadStateCount(state As downloadProgress_params_state) As Long
    '1. 確認するstatus文字列を取得
    Dim FilterState As String
    FilterState = EnumToStringDownloadProgressState(state)

    '2. 全GUIDを取り出し、ForEachで回す
    Dim DownloadGUID
    For Each DownloadGUID In DownloadWillBegin.keys
        '3. GUIDがあるか？
        If DownloadProgress.Exists(DownloadGUID) Then
            '3-1. 一致するstatusなら、カウント
            If DownloadProgress(DownloadGUID).StringKey("state") = FilterState Then DownloadStateCount = DownloadStateCount + 1
        Else
            '3-1. `downloadWillBegin`しか来てないstatusならカウント
            If state = downloadProgress_params_state.state_willBegin Then DownloadStateCount = DownloadStateCount + 1
        End If
    Next
End Property

'***************************************************************************************************
'* 機能　　：`downloadProgress - state`が、`completed`になってるリストのみ削除します
'***************************************************************************************************
Public Sub ClearDownloadCompleted()
    DeleteStatusGUID downloadProgress_params_state.state_completed
End Sub

'***************************************************************************************************
'* 機能　　：`downloadProgress - state`が、`canceled`になってるリストのみ削除します
'***************************************************************************************************
Public Sub ClearDownloadCanceled()
    DeleteStatusGUID downloadProgress_params_state.state_canceled
End Sub



'***************************************************************************************************
'                      ■■■ メインとなる機能拡張部分 - イベント処理 ■■■
'---------------------------------------------------------------------------------------------------
'* 機能   ：ダウンロード系のイベントのみ、パース処理したのをただただ、Dictionaryに蓄積していきます
'***************************************************************************************************
Private Sub ex_CDPCore_CDPBrowserEvent(methodName As String, RawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPCore_CDPBrowserEvent"


    'この拡張Class上で、1度も設定されてない場合は即抜け
    If DownloadPathList.Count = 0 Then Exit Sub

    '1. 対象のダウンロードイベントを処理
    Select Case methodName
        Dim downloadInfos   As BiDiCDPJson
        Dim DownloadGUID    As String

        'ダウンロード開始情報
        Case "Browser.downloadWillBegin"
            '1-1. Json解析
            Set downloadInfos = BiDiCDPJson.Parse(RawJson).NodeKey("params")

            '1-2. `GUID`を取り出す
            DownloadGUID = downloadInfos.StringKey("guid")

            '1-3. `GUID`をKeyとして登録
            DownloadWillBegin.Add DownloadGUID, downloadInfos

            '1-4. デバッグ用メッセージ表示
            If seeRawSendMsgDbg Then
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "Download detected. information", FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName

                printMsg Debug_, "guid              : " & downloadInfos.StringKey("guid"), FromProcedureName
                printMsg Debug_, "frameId           : " & downloadInfos.StringKey("frameId"), FromProcedureName
                printMsg Debug_, "url               : " & downloadInfos.StringKey("url"), FromProcedureName
                printMsg Debug_, "suggestedFilename : " & downloadInfos.StringKey("suggestedFilename"), FromProcedureName

                printMsg Debug_, String(100, "="), FromProcedureName
            End If

        'ダウンロード進捗情報
        Case "Browser.downloadProgress"
            '1-1. Json解析
            Set downloadInfos = BiDiCDPJson.Parse(RawJson).NodeKey("params")

            '1-2. `GUID`を取り出す
            DownloadGUID = downloadInfos.StringKey("guid")

            '1-3. `GUID`をKeyとして登録/更新
            Set DownloadProgress(DownloadGUID) = downloadInfos

            '1-4. デバッグ用メッセージ表示
            If seeRawSendMsgDbg Then
                printMsg Debug_, String(100, "="), FromProcedureName
                printMsg Debug_, "`" & DownloadGUID & "` Download progress", FromProcedureName
                printMsg Debug_, String(100, "="), FromProcedureName

                printMsg Debug_, "state             : " & downloadInfos.StringKey("state"), FromProcedureName
                printMsg Debug_, "filePath          : " & downloadInfos.StringKey("filePath"), FromProcedureName

                '状況に応じた進捗表示
                Dim totalBytes As Long:     totalBytes = downloadInfos.NumberKey("totalBytes")
                Dim receivedBytes As Long:  receivedBytes = downloadInfos.NumberKey("receivedBytes")
                If totalBytes > 0 Then
                    ' 全体サイズが分かるとき、% 表示
                    Dim progressStr As String
                    progressStr = FormatPercent(receivedBytes / totalBytes, 2)
                    printMsg Debug_, "progress          : " & progressStr & " (" & receivedBytes & " / " & totalBytes & ")", FromProcedureName
                Else
                    ' 全体サイズが分からないとき、受信バイト数だけ表示
                    printMsg Debug_, "progress          : " & Format(receivedBytes / 1024, "#,##0") & " KB received", FromProcedureName
                End If

                printMsg Debug_, String(100, "="), FromProcedureName
            End If

    End Select
End Sub



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************
'* 機能   ：指定のguidが、ダウンロード終るまで待機し続けます
'---------------------------------------------------------------------------------------------------
'* 返り値　：ダウンロード完了論理値
'* 引数　　：guid           監視するguid
'            TimeOutSecond  タイムアウト秒数
'---------------------------------------------------------------------------------------------------
'* 注意事項：・タイムアウトになったら、エラーにはならず、この関数を抜けるだけです
'            ・`TimeOutSecond`は、前回の進捗から変化がない秒数として機能します
'            ・`downloadWillBegin`にないものは即抜けします
'***************************************************************************************************
Public Function AutoWaitSingleCompleted(guid As String, Optional TimeOutSecond As Double = 30) As Boolean
    Const FromProcedureName As String = ThisClassName & ".AutoWaitSingleCompleted"


    '1. `downloadWillBegin`に、その`guid`があるか？
    If Not DownloadWillBegin.Exists(guid) Then printMsg WARN_, "A `guid - " & guid & "` was specified that is not present in `downloadWillBegin`.", ThisClassName: Exit Function

    '2. 現在の経過時間を取得
    Dim timerStart As Double: timerStart = TimerCounter

    '3. メッセージを取り出しながら、進捗をチェックする
    Dim CurrentReceivedBytes As Long
    Do
        'downloadProgress - state に応じた分岐
        Select Case Me.DownloadProgressInfo(guid, downloadProgress_state)
            Case EnumToStringDownloadProgressState(state_inProgress)
                '進捗取り出し
                With Me
                    Dim Progress_totalBytes     As Long: Progress_totalBytes = .DownloadProgressInfo(guid, downloadProgress_totalBytes)
                    Dim Progress_receivedBytes  As Long: Progress_receivedBytes = .DownloadProgressInfo(guid, downloadProgress_receivedBytes)
                End With

                '3-1. 前回値より変化してたらタイマーリセットし、進捗表示
                If CurrentReceivedBytes < Progress_receivedBytes Then
                    'タイマーリセット
                    timerStart = TimerCounter

                    '進捗計算し、表示
                    If Progress_totalBytes > 0 Then
                        ' 全体サイズが分かるとき、% 表示
                        Dim progressStr As String
                        progressStr = FormatPercent(Progress_receivedBytes / Progress_totalBytes, 2)
                        printMsg info_, guid & " - progress : " & progressStr & " (" & Progress_receivedBytes & " / " & Progress_totalBytes & ")", FromProcedureName
                    Else
                        ' 全体サイズが分からないとき、受信バイト数だけ表示
                        printMsg info_, guid & " - progress : " & Format(Progress_receivedBytes / 1024, "#,##0") & " KB received", FromProcedureName
                    End If

                    '現時点での、`receivedBytes`情報を更新
                    CurrentReceivedBytes = Progress_receivedBytes
                End If

            Case EnumToStringDownloadProgressState(state_completed)
                '3-1. 正常終了なら、`True`返却
                AutoWaitSingleCompleted = True
                printMsg info_, "Download completed: " & guid, FromProcedureName
                Exit Do

            Case EnumToStringDownloadProgressState(state_canceled)
                '3-1. ユーザーによるキャンセルあるいは、ウイルス判定による強制キャンセル時は、`False`返却
                AutoWaitSingleCompleted = False
                printMsg WARN_, "Download canceled : " & guid, FromProcedureName
                Exit Do
        End Select

        '3-2. 指定秒数、進捗に変化なかったら、抜ける
        If TimerCounter - timerStart > TimeOutSecond * 1000 Then
            printMsg WARN_, "The process timed out because there was no change in progress for a certain period of time.", FromProcedureName
            Exit Do
        End If

        '3-3. メッセージ取り出し
        TakeEvents  '内部で`DoEvents`が走ってくれます
    Loop
End Function

'***************************************************************************************************
'* 機能   ：指定の数分、ダウンロードが終るまで待機し続けます
'---------------------------------------------------------------------------------------------------
'* 返り値　：ダウンロード完了数
'* 引数　　：CompleteCount  完了とみなす数
'            TimeOutSecond  タイムアウト秒数
'            ClearDLHistory `True`で、`downloadProgress - state`が、`completed`,`canceled`になってる`GUID`を削除してから、処理を開始します
'---------------------------------------------------------------------------------------------------
'* 機能説明：・返り値と`CompleteCount`を比較して、ちゃんとDLが終えたかを確認できます
'            ・この中身のロジックを参考に、bas側で自力で高度な待機も可能です
'            ・基本は複数待ちとして機能しますが、`CompleteCount:=1` × `ClearDLHistory:=True`とすることで、「最新の1件待機」といったことも可能です
'
'* 注意事項：・タイムアウトになったら、エラーにはならず、この関数を抜けるだけです
'            ・`TimeOutSecond`は、前回の進捗から変化がない秒数として機能します
'            ・`ClearDLHistory`は、このメソッドを使うタイミングによっては、不整合になる場合があります。
'              例えば、クリックによる並列DLをして、このメソッドを呼ぶ時点で既に、最初のクリックによるDLが終ってしまった場合です
'              その場合は、bas側で適切なタイミング(DLが来るタイミングの前にクリア等)で手動クリアしてください
'***************************************************************************************************
Public Function AutoWaitMultiCompleted(CompleteCount As Long, Optional TimeOutSecond As Double = 30, Optional ClearDLHistory As Boolean) As Long
    Const FromProcedureName As String = ThisClassName & ".AutoWaitMultiCompleted"


    'エラーガード
    If CompleteCount < 1 Then printMsg ERROR_, "Please set `CompleteCount` to 1 or greater.", FromProcedureName, doRaiseError:=True

    '1. `downloadProgress - state`が、`completed`,`canceled`になってる`GUID`を削除
    If ClearDLHistory Then
        Call ClearDownloadCompleted
        Call ClearDownloadCanceled
    End If

    '2. 現在の経過時間を取得
    Dim timerStart As Double: timerStart = TimerCounter

    '3. `downloadWillBegin`が、指定の数になるまで待機
    Dim CurrentDownloadGuidCount As Long: CurrentDownloadGuidCount = Me.DownloadGuidCount
    Do Until CurrentDownloadGuidCount = CompleteCount
        '3-1. メッセージ取り出し
        TakeEvents  '内部で`DoEvents`が走ってくれます

        '3-2. 前回値に変化があったらタイマーリセットし、更新
        If CurrentDownloadGuidCount < Me.DownloadGuidCount Then timerStart = TimerCounter: CurrentDownloadGuidCount = Me.DownloadGuidCount

        '3-3. 指定数が取り出せたら、タイムアウトチェックの前に抜ける
        If Me.DownloadGuidCount = CompleteCount Then
            printMsg info_, "I found " & CompleteCount & " instances of `downloadWillBegin`.", FromProcedureName
            Exit Do
        End If

        '3-4. 指定秒数、数に変化なかったら、一旦抜けて、今ある数で次のStepへ
        If TimerCounter - timerStart > TimeOutSecond * 1000 Then
            printMsg WARN_, "We were unable to detect " & CompleteCount & " `downloadWillBegin` instances within the time limit. We will continue with " & CurrentDownloadGuidCount & " instances.", FromProcedureName

            '0個なら、続行できないので関数から抜ける
            If CurrentDownloadGuidCount Then
                printMsg WARN_, "Since one or more instances of `downloadWillBegin` were detected, we will proceed to the next step.", FromProcedureName
                Exit Do
            Else
                printMsg ERROR_, "The process will be interrupted because no `downloadWillBegin` requests have been received.", FromProcedureName
                Exit Function
            End If
        End If
    Loop

    '4. 現時点での、GUIDリストを取得しておく
    Dim TargetDownloadGUIDs: TargetDownloadGUIDs = Me.DownloadGuidList

    '5. 各種初期化
    '5-1. 各`receivedBytes`情報を初期化
    Dim DownloadGUID
    Dim DownloadGUIDs_receivedBytes As New Dictionary
    For Each DownloadGUID In TargetDownloadGUIDs
        DownloadGUIDs_receivedBytes.Add DownloadGUID, 0
    Next

    '5-2. 各status情報を初期化
    Dim CanceledGUIDs As New Dictionary
    Dim CompletedGUIDs As New Dictionary

    '6. メッセージを取り出しながら、1個1個、進捗をチェックする
    timerStart = TimerCounter '次のセクションなので、リセット
    Do
        '6-1. 受信サイズとステータスチェック
        For Each DownloadGUID In TargetDownloadGUIDs
            'downloadProgress - state に応じた分岐
            Select Case Me.DownloadProgressInfo(CStr(DownloadGUID), downloadProgress_state)
                Case EnumToStringDownloadProgressState(state_inProgress)
                    '進捗取り出し
                    With Me
                        Dim Progress_totalBytes     As Long:    Progress_totalBytes = .DownloadProgressInfo(CStr(DownloadGUID), downloadProgress_totalBytes)
                        Dim Progress_receivedBytes  As Long:    Progress_receivedBytes = .DownloadProgressInfo(CStr(DownloadGUID), downloadProgress_receivedBytes)
                    End With
    
                    '6-1-1. 前回値より変化してたらタイマーリセットし、進捗表示
                    If DownloadGUIDs_receivedBytes(DownloadGUID) < Progress_receivedBytes Then
                        'タイマーリセット
                        timerStart = TimerCounter
    
                        '進捗計算し、表示
                        If Progress_totalBytes > 0 Then
                            ' 全体サイズが分かるとき、% 表示
                            Dim progressStr As String
                            progressStr = FormatPercent(Progress_receivedBytes / Progress_totalBytes, 2)
                            printMsg info_, DownloadGUID & " - progress : " & progressStr & " (" & Progress_receivedBytes & " / " & Progress_totalBytes & ")", FromProcedureName
                        Else
                            ' 全体サイズが分からないとき、受信バイト数だけ表示
                            printMsg info_, DownloadGUID & " - progress : " & Format(Progress_receivedBytes / 1024, "#,##0") & " KB received", FromProcedureName
                        End If
    
                        '現時点での、`receivedBytes`情報を格納
                        DownloadGUIDs_receivedBytes(DownloadGUID) = Progress_receivedBytes
                    End If

                Case EnumToStringDownloadProgressState(state_completed)
                    '正常終了なら、正常終了リストに追加
                    If Not CompletedGUIDs.Exists(DownloadGUID) Then
                        CompletedGUIDs.Add DownloadGUID, True
                        printMsg info_, CompletedGUIDs.Count + CanceledGUIDs.Count & "/" & CurrentDownloadGuidCount & " Download completed: " & DownloadGUID, FromProcedureName
                    End If

                Case EnumToStringDownloadProgressState(state_canceled)
                    'ユーザーによるキャンセルあるいは、ウイルス判定による強制キャンセル時は、キャンセルリストに追加
                    If Not CanceledGUIDs.Exists(DownloadGUID) Then
                        CanceledGUIDs.Add DownloadGUID, True
                        printMsg WARN_, CanceledGUIDs.Count + CompletedGUIDs.Count & "/" & CurrentDownloadGuidCount & " Download canceled : " & DownloadGUID, FromProcedureName
                    End If
            End Select
        Next

        '6-2. ダウンロード自体が全て終わったら、タイムアウトチェックの前に抜ける
        If (CompletedGUIDs.Count + CanceledGUIDs.Count) = CurrentDownloadGuidCount Then
            printMsg info_, "All downloads are complete.", FromProcedureName
            If CanceledGUIDs.Count Then printMsg WARN_, "There are downloads that will be cancelled.", FromProcedureName
            Exit Do
        End If

        '6-3. 指定秒数、進捗に変化なかったら、抜ける
        If TimerCounter - timerStart > TimeOutSecond * 1000 Then
            printMsg WARN_, "The process timed out because there was no change in progress for a certain period of time.", FromProcedureName
            Exit Do
        End If

        '6-4. メッセージ取り出し
        TakeEvents  '内部で`DoEvents`が走ってくれます
    Loop

    '7. ダウンロードが正常終了した数で、返却
    AutoWaitMultiCompleted = CompletedGUIDs.Count
End Function

'***************************************************************************************************
'* 機能   ：指定のステータスのGUIDを一括削除します
'---------------------------------------------------------------------------------------------------
'* 引数   ：state  Enum - downloadProgress_params_state
'***************************************************************************************************
Private Sub DeleteStatusGUID(state As downloadProgress_params_state)
    Const FromProcedureName As String = ThisClassName & ".DeleteStatusGUID"


    '1. status文字列を取得
    Dim ClearState As String
    ClearState = EnumToStringDownloadProgressState(state)

    '2. 各`guid`情報を`downloadProgress`に照会
    Dim DownloadGUIDs: DownloadGUIDs = DownloadProgress.keys
    Dim CurrentGUID
    For Each CurrentGUID In DownloadGUIDs
        '3. 一致するstatusなら、削除
        If DownloadProgress(CurrentGUID).StringKey("state") = ClearState Then
            DownloadWillBegin.Remove (CurrentGUID)
            DownloadProgress.Remove (CurrentGUID)

            printMsg Debug_, "Removed from the status `" & ClearState & "` : " & CurrentGUID, FromProcedureName
        End If
    Next
End Sub

'***************************************************************************************************
'* 機能   ：`DownloadProgress`,`DownloadWillBegin`の中身をイミディエイトウィンドウに表示させます
'---------------------------------------------------------------------------------------------------
'* 引数   ：guid    見たい`guid`    ※省略すると、全表示します
'***************************************************************************************************
Public Sub PrintDownloadSummary(Optional guid As String)
    Const FromProcedureName As String = ThisClassName & ".PrintDownloadSummary"


    '単一表示判定
    Dim tmp As BiDiCDPJson
    If StrPtr(guid) Then
        'ない場合は即抜け
        If Not DownloadWillBegin.Exists(guid) Then printMsg WARN_, "A `guid - " & guid & "` was specified that is not present in `downloadWillBegin`.", FromProcedureName: Exit Sub

        '1. ヘッダーログ
        printMsg info_, String(100, "="), FromProcedureName
        printMsg info_, "guid - `" & guid & "` information", FromProcedureName
        printMsg info_, String(100, "="), FromProcedureName

        '2. `DownloadWillBegin`情報
        Set tmp = DownloadWillBegin(guid)

        printMsg info_, "guid              : " & tmp.StringKey("guid"), FromProcedureName
        printMsg info_, "frameId           : " & tmp.StringKey("frameId"), FromProcedureName
        printMsg info_, "url               : " & tmp.StringKey("url"), FromProcedureName
        printMsg info_, "suggestedFilename : " & tmp.StringKey("suggestedFilename"), FromProcedureName

        '3. `DownloadProgress`情報
        If DownloadProgress.Exists(guid) Then
            Set tmp = DownloadProgress(guid)

            printMsg info_, "state             : " & tmp.StringKey("state"), FromProcedureName
            printMsg info_, "filePath          : " & tmp.StringKey("filePath"), FromProcedureName
            printMsg info_, "receivedBytes     : " & tmp.NumberKey("receivedBytes"), FromProcedureName
            printMsg info_, "totalBytes        : " & tmp.NumberKey("totalBytes"), FromProcedureName

        Else
            printMsg info_, "`DownloadProgress` is nothing...", FromProcedureName
        End If
    Else
        '1つもない場合は即抜け
        If DownloadWillBegin.Count = 0 Then printMsg WARN_, "There are no instances of `downloadWillBegin`.", FromProcedureName: Exit Sub
        printMsg info_, String(100, "="), FromProcedureName
        printMsg info_, "Download Summary - " _
                      & DownloadWillBegin.Count & " total | " _
                      & DownloadStateCount(state_completed) & " completed | " _
                      & DownloadStateCount(state_inProgress) & " in-progress", FromProcedureName
        printMsg info_, String(100, "-"), FromProcedureName

        'For Eachで回す
        Dim CurrentGUID
        For Each CurrentGUID In DownloadWillBegin.keys
            '1. ヘッダーログ
            printMsg info_, "guid - `" & CurrentGUID & "` information", FromProcedureName
            printMsg info_, String(100, "-"), FromProcedureName

            '2. `DownloadWillBegin`情報
            Set tmp = DownloadWillBegin(CurrentGUID)
    
            printMsg info_, "guid              : " & tmp.StringKey("guid"), FromProcedureName
            printMsg info_, "frameId           : " & tmp.StringKey("frameId"), FromProcedureName
            printMsg info_, "url               : " & tmp.StringKey("url"), FromProcedureName
            printMsg info_, "suggestedFilename : " & tmp.StringKey("suggestedFilename"), FromProcedureName
    
            '3. `DownloadProgress`情報
            If DownloadProgress.Exists(CurrentGUID) Then
                Set tmp = DownloadProgress(CurrentGUID)
    
                printMsg info_, "state             : " & tmp.StringKey("state"), FromProcedureName
                printMsg info_, "filePath          : " & tmp.StringKey("filePath"), FromProcedureName
                printMsg info_, "receivedBytes     : " & tmp.NumberKey("receivedBytes"), FromProcedureName
                printMsg info_, "totalBytes        : " & tmp.NumberKey("totalBytes"), FromProcedureName
    
            Else
                printMsg info_, "`DownloadProgress` is nothing...", FromProcedureName
            End If

            '区切り
            printMsg info_, String(100, "-"), FromProcedureName
        Next
    End If

    '〆
    printMsg info_, String(100, "="), FromProcedureName
End Sub



'***************************************************************************************************
'                           ■■■ `CDPBrowser/CDPCore`直通メソッド ■■■
'***************************************************************************************************
'* 機能　　：いちいち、`○○.ExecuteCDP`といった記述を防ぐ措置となります。
'            メインの`CDPBrowser`メソッドを行う前後の処理等も出来ますが基本は、1行だけに留めておくこと
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function ExecuteCDP(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As BiDiCDPJson
    Set ExecuteCDP = ex_CDPBrowser.ExecuteCDP(methodName, params, StopError)
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
'---------------------------------------------------------------------------------------------------
'* 注意事項：イベント処理自体はブラウザ単位ですがほとんどは`タブ`から始める方が多いことから、`CDPContext`として継承するようにしてます
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPContext)
    '必要なCDPオブジェクトを継承
    Set ex_CDPBrowser = Inheritance.InheritanceCDPBrowser
    Set ex_CDPCore = Inheritance.InheritanceCDPCore

    '渡した際のデフォルトのブラウザコンテキストをセット
    defaultBrowserContextId = Inheritance.CurrentBrowserContextId
End Sub

'***************************************************************************************************
'* 機能　　：初期設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    '各種変数の初期化
    Set DownloadPathList = New Dictionary
    Set DownloadProgress = New Dictionary
    Set DownloadWillBegin = New Dictionary

    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    seeRawSendMsgDbg = ShSetting01_StartBrowser.UseRangeID(8, FromProcedureName)   'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？
End Sub



'***************************************************************************************************
'                                   ■■■ 後始末 ■■■
'***************************************************************************************************
'* 機能　　：拡張機能との連携を切る際に記述します
'***************************************************************************************************
Private Sub Class_Terminate()
    '`setDownloadBehavior`設定を元に戻します
    Dim browserContextId
    For Each browserContextId In DownloadPathList.keys
        '1. CDPパラメーター構築
        Dim CDPparams As New Dictionary
        CDPparams.Add "behavior", EnumToStringSetDownloadBehavior_behavior(Default)
        If browserContextId <> defaultBrowserContextId Then CDPparams.Add "browserContextId", browserContextId
        CDPparams.Add "eventsEnabled", False

        '2. 実行
        ExecuteCDP "Browser.setDownloadBehavior", CDPparams, False
    Next
End Sub
