VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exBiDi_DiscoveryLogRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'       一定時間だけWebDriver BiDiの非同期イベントとDOM変化を記録し、テキストログとして保存します
'---------------------------------------------------------------------------------------------------
'   「準備 → OKを押す → 手動でブラウザを操作 → N秒後に自動停止 → discovery_log.txt保存」という、
'   手動操作の因果関係を後から追いやすくするための調査用レコーダーです。
'
'   記録対象：
'     ・ネットワーク通信（network.beforeRequestSent / responseCompleted / fetchError）
'     ・ページ遷移（browsingContext.navigationStarted / domContentLoaded / load）
'     ・コンソールログ（log.entryAdded）
'     ・DOM要素の変化（MutationObserverをjsEvalで注入し、childList / characterDataを監視）
'
'   詳細・既知の制約は、このフォルダの ReadMe.md を参照してください。
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
Private Const ThisClassName     As String = "exBiDi_DiscoveryLogRecorder"  'ログのトレース用

'ステート管理
Private currentLog              As StateLog  'ログ設定
Private m_context                As String   '`WebDriverBiDiContext`継承時の`Context`情報

'記録状態管理用
Private Type DiscoveryLogState
    IsRecording         As Boolean
    recorderLogs        As Collection   '整形済みログ行（リングバッファ）
    recSuppByKey         As Dictionary   '除外(excludeImagesAndCss)したURLの正規化キー→件数
    subscribedEvents     As Collection   'Stop時に同じ配列でunsubscribeするために保持
    recorderStartTime    As Double       '記録開始時刻(CDPHelpers.TimerCounter基準・ms)
    excludeResources     As Boolean
    captureDom           As Boolean
    droppedHeadCount     As Long
    totalEventCount      As Long
    Limit                As Long
    StartURL             As String
End Type
Private CurrentLogState As DiscoveryLogState



'***************************************************************************************************
'                             ■■■ 取得系プロパティ ■■■
'***************************************************************************************************
'* 機能　　：現在、記録中かどうかを返します
'***************************************************************************************************
Public Property Get IsRecording() As Boolean
    IsRecording = CurrentLogState.IsRecording
End Property



'***************************************************************************************************
'                           ■■■ メインとなる機能拡張部分 ■■■
'***************************************************************************************************
'* 機能　　：記録対象イベントを購読し、DOM監視を仕込んで、記録モードに入ります
'---------------------------------------------------------------------------------------------------
'* 引数　　：excludeImagesAndCss   `True`:画像・CSS等の静的リソース通信をログから除外（既定）
'            captureDom            `True`:MutationObserverによるDOM変化の記録を有効化（既定）
'---------------------------------------------------------------------------------------------------
'* 注意事項：`excludeImagesAndCss`は、あくまでもログへの出力を間引くだけで、実際の通信はブロックしません
'***************************************************************************************************
Public Sub StartDiscoveryLog(Optional excludeImagesAndCss As Boolean = True, Optional captureDom As Boolean = True)
    Const FromProcedureName As String = ThisClassName & ".StartDiscoveryLog"


    '1. エラー情報をクリア
    With ex_BiDiContext.ThisWebDriverBiDiMode
        If Not .LastBiDiJsonError Is Nothing Then .LastBiDiJsonError.RemoveAll
    End With

    '2. 必要なイベントを購読
    Dim targetEvents As New Collection
    targetEvents.Add "network.beforeRequestSent"
    targetEvents.Add "network.responseCompleted"
    targetEvents.Add "network.fetchError"
    targetEvents.Add "browsingContext.navigationStarted"
    targetEvents.Add "browsingContext.domContentLoaded"
    targetEvents.Add "browsingContext.load"
    targetEvents.Add "log.entryAdded"
    Set ex_BiDiContext.ThisWebDriverBiDiMode.sessionSubscribe = targetEvents

    '3. エラーcheck
    With ex_BiDiContext.ThisWebDriverBiDiMode
        If Not .LastBiDiJsonError Is Nothing Then
            If .LastBiDiJsonError.Count Then
                printMsg ERROR_, "`session.subscribe`の実行に失敗しました。 > " & ex_BiDiContext.ThisWebDriverBiDiMode.LastBiDiJsonError("message"), FromProcedureName, doRaiseError:=True
            End If
        End If
    End With

    '4. 開始時点のURLを、失敗しても記録全体を止めない形で取得
    Dim urlResult As Variant
    urlResult = ex_BiDiContext.jsEval("location.href", StopBiDiError:=False)

    '5. 各種初期化
    With CurrentLogState
        Set .recorderLogs = New Collection
        Set .recSuppByKey = New Dictionary
        Set .subscribedEvents = targetEvents
        .recorderStartTime = CDPHelpers.TimerCounter
        .excludeResources = excludeImagesAndCss
        .captureDom = captureDom
        .droppedHeadCount = 0
        .totalEventCount = 0
        .Limit = 5000
        .IsRecording = True
        If VarType(urlResult) = vbString Then .StartURL = urlResult Else .StartURL = "(unknown)"
    End With

    '6. DOM監視を、現在表示中のページへ即座に仕込む
    If captureDom Then InjectMutationObserver

    printMsg info_, "Discovery Log Recorder: 記録を開始しました。(excludeImagesAndCss=" & excludeImagesAndCss & ", captureDom=" & captureDom & ")", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：指定秒数の間、BiDiパイプとDOM変化バッファを定期的にドレインし続けます
'---------------------------------------------------------------------------------------------------
'* 引数　　：seconds               記録し続ける秒数
'            pollIntervalSeconds   ドレイン間隔（既定0.2秒）
'* 注意事項：この間、ユーザーが手動でブラウザを操作することを想定しています
'***************************************************************************************************
Public Sub RecordEventsForSeconds(seconds As Double, Optional pollIntervalSeconds As Double = 0.2)
    Const FromProcedureName As String = ThisClassName & ".RecordEventsForSeconds"


    '1. 記録開始前なら、即抜け
    If Not CurrentLogState.IsRecording Then
        printMsg WARN_, "`StartDiscoveryLog`が実行されていないため、記録できません。", FromProcedureName
        Exit Sub
    End If

    '2. 経過秒数を見ながら、ドレインを繰り返す
    Dim startedAt As Double: startedAt = CDPHelpers.TimerCounter

    Do
        TakeEvents
        If CurrentLogState.captureDom Then DrainDomMutations

        If (CDPHelpers.TimerCounter - startedAt) >= seconds * 1000 Then Exit Do

        CDPHelpers.Sleep pollIntervalSeconds
    Loop
End Sub

'***************************************************************************************************
'* 機能　　：記録を停止し、DOM監視の解除・イベント購読解除を行った上で、ログを整形して返却・保存します
'---------------------------------------------------------------------------------------------------
'* 返り値　：整形済みログ本文（ヘッダー＋時系列本文＋除外サマリー）の全文
'* 引数　　：FolderPath    指定した場合、この保存先フォルダへ`FileName`で保存します（省略時は保存のみスキップ）
'            FileName      保存ファイル名（既定"discovery_log.txt"）
'***************************************************************************************************
Public Function StopAndSaveDiscoveryLog(Optional FolderPath As String, Optional FileName As String = "discovery_log.txt") As String
    Const FromProcedureName As String = ThisClassName & ".StopAndSaveDiscoveryLog"


    '1. 記録開始前なら、即抜け
    If Not CurrentLogState.IsRecording Then
        printMsg WARN_, "記録は開始されていません。", FromProcedureName
        Exit Function
    End If

    '2. 最終ドレイン
    TakeEvents
    If CurrentLogState.captureDom Then DrainDomMutations

    '3. DOM監視の解除
    If CurrentLogState.captureDom Then DetachMutationObserver

    '4. イベント購読の解除
    If Not CurrentLogState.subscribedEvents Is Nothing Then
        Set ex_BiDiContext.ThisWebDriverBiDiMode.sessionSubscribe(False) = CurrentLogState.subscribedEvents
    End If

    CurrentLogState.IsRecording = False

    '5. 本文を時系列でソートして整形
    Dim bodyText As String
    bodyText = Join(SortedLinesArray(), vbCrLf)

    '6. 除外サマリーを整形
    Dim summaryText As String
    summaryText = BuildSuppressedSummary()

    '7. ヘッダーを整形
    Dim headerText As String
    headerText = "===== Discovery Log Recorder =====" & vbCrLf & _
                 "開始URL        : " & CurrentLogState.StartURL & vbCrLf & _
                 "保存時刻        : " & Format(Now, "yyyy-MM-dd HH:mm:ss") & vbCrLf & _
                 "記録行数/検知数 : " & CurrentLogState.recorderLogs.Count & " / " & CurrentLogState.totalEventCount & vbCrLf & _
                 "破棄行数(先頭)  : " & CurrentLogState.droppedHeadCount & vbCrLf & _
                 "==================================="

    Dim finalText As String
    finalText = headerText & vbCrLf & vbCrLf & bodyText & vbCrLf & vbCrLf & summaryText

    '8. 指定があれば保存
    If LenB(FolderPath) > 0 Then
        Dim conv As New CharacterCodeConversion
        conv.BytesToSaveFile conv.BytesFromString(finalText), FolderPath, FileName
        printMsg info_, "Discovery Log を保存しました。 > " & FolderPath & "\" & FileName, FromProcedureName
    End If

    StopAndSaveDiscoveryLog = finalText
End Function



'***************************************************************************************************
'                     ■■■ 非同期イベントの処理 ■■■
'***************************************************************************************************
'* 機能　　：WebDriverBiDiからの非同期イベントのうち、記録対象のみを`FormatAndStoreLog`へ委託します
'***************************************************************************************************
Private Sub ex_BiDiCore_detectionBiDiEvent(methodName As String, RawJson As String)
    '1. 記録モードでなければ、即抜け
    If Not CurrentLogState.IsRecording Then Exit Sub

    '2. 対象イベントのみ委託
    Select Case methodName
        Case "network.beforeRequestSent", "network.responseCompleted", "network.fetchError", _
             "browsingContext.navigationStarted", "browsingContext.domContentLoaded", "browsingContext.load", _
             "log.entryAdded"
            FormatAndStoreLog methodName, RawJson
    End Select
End Sub

'***************************************************************************************************
'* 機能　　：非同期イベント1件分を、ログ1行分の文字列に整形し、リングバッファへ蓄積します
'***************************************************************************************************
Private Sub FormatAndStoreLog(methodName As String, RawJson As String)
    Dim paramsNode As BiDiCDPJson
    Set paramsNode = BiDiCDPJson.Parse(RawJson).NodeKey("params")
    If paramsNode Is Nothing Then Exit Sub

    Dim relSeconds As String
    relSeconds = Format((CDPHelpers.TimerCounter - CurrentLogState.recorderStartTime) / 1000, "0.00")

    Dim logLine As String

    Select Case methodName

        Case "network.beforeRequestSent"
            If paramsNode.StringKey("context") <> m_context Then Exit Sub

            Dim reqNodeA As BiDiCDPJson
            Set reqNodeA = paramsNode.NodeKey("request")
            If reqNodeA Is Nothing Then Exit Sub

            Dim urlA As String: urlA = reqNodeA.StringKey("url")
            Dim methodA As String: methodA = reqNodeA.StringKey("method")

            If CurrentLogState.excludeResources Then
                If IsNoiseUrl(urlA) Then BumpSuppressed NormalizeEndpointKey(urlA): Exit Sub
            End If

            logLine = "[" & relSeconds & "s] [REQ] " & methodA & " " & urlA

        Case "network.responseCompleted"
            If paramsNode.StringKey("context") <> m_context Then Exit Sub

            Dim resNode As BiDiCDPJson
            Set resNode = paramsNode.NodeKey("response")
            If resNode Is Nothing Then Exit Sub

            Dim urlB As String: urlB = resNode.StringKey("url")
            If CurrentLogState.excludeResources Then
                If IsNoiseUrl(urlB) Then Exit Sub
            End If

            Dim statusB As String: statusB = resNode.StringKey("status")
            Dim mimeB As String: mimeB = resNode.StringKey("mimeType")

            logLine = "[" & relSeconds & "s] [RES] " & statusB & " " & mimeB & " " & urlB

        Case "network.fetchError"
            If paramsNode.StringKey("context") <> m_context Then Exit Sub

            Dim reqNodeC As BiDiCDPJson
            Set reqNodeC = paramsNode.NodeKey("request")

            Dim urlC As String
            If Not reqNodeC Is Nothing Then urlC = reqNodeC.StringKey("url")

            If CurrentLogState.excludeResources Then
                If IsNoiseUrl(urlC) Then Exit Sub
            End If

            logLine = "[" & relSeconds & "s] [ERR] " & paramsNode.StringKey("errorText") & " " & urlC

        Case "browsingContext.navigationStarted"
            If paramsNode.StringKey("context") <> m_context Then Exit Sub
            logLine = "[" & relSeconds & "s] [NAV] navigationStarted " & paramsNode.StringKey("url")

        Case "browsingContext.domContentLoaded"
            If paramsNode.StringKey("context") <> m_context Then Exit Sub
            logLine = "[" & relSeconds & "s] [NAV] domContentLoaded " & paramsNode.StringKey("url")

        Case "browsingContext.load"
            If paramsNode.StringKey("context") <> m_context Then Exit Sub
            logLine = "[" & relSeconds & "s] [NAV] load " & paramsNode.StringKey("url")

            '新しいドキュメントに切り替わり、DOM監視(JS側の状態)は失われているため、再度仕込み直す
            If CurrentLogState.captureDom Then InjectMutationObserver

        Case "log.entryAdded"
            Dim sourceNode As BiDiCDPJson
            Set sourceNode = paramsNode.NodeKey("source")
            If sourceNode Is Nothing Then Exit Sub
            If sourceNode.StringKey("context") <> m_context Then Exit Sub

            Dim logText As String: logText = paramsNode.StringKey("text")
            If Len(logText) > 200 Then logText = Left$(logText, 200) & "..."

            logLine = "[" & relSeconds & "s] [LOG] " & paramsNode.StringKey("level") & " " & logText

    End Select

    PushTimelineLine logLine
End Sub



'***************************************************************************************************
'                           ■■■ DOM変化の監視（MutationObserver） ■■■
'***************************************************************************************************
'* 機能　　：現在のドキュメントへ、DOM変化を蓄積するMutationObserverをJavaScriptとして注入します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・`childList`（要素の追加/削除）と`characterData`（テキスト変化）のみを監視します
'              （`attributes`はフレームワークのUI更新等で大量に発火し、ログが埋もれるため既定で対象外）
'            ・同一バッチ内で、既に記録した要素の子孫にあたる追加ノードは、重複記録しません
'            ・相対秒はページ自身の`performance.now()`基準です（`DrainDomMutations`側の
'              ポーリング間隔単位でしかVBA側へ届かないため、他イベントとの厳密な前後関係は保証しません）
'* 注意事項：メインドキュメントのみが対象です。クロスオリジンiframe内のDOMは監視できません
'***************************************************************************************************
Private Sub InjectMutationObserver()
    Dim js As String
    js = "if(!window.__bidiDiscoveryActive){"
    js = js & "window.__bidiDiscoveryActive=true;"
    js = js & "window.__bidiDiscoveryDom=[];"
    js = js & "window.__bidiDiscoveryStart=performance.now();"
    js = js & "window.__bidiDiscoveryDescribe=function(n){"
    js = js & "if(n.nodeType!==1) return '#text';"
    js = js & "var s=n.tagName.toLowerCase();"
    js = js & "if(n.id) s+='#'+n.id;"
    js = js & "else if(n.className && typeof n.className==='string' && n.className) s+='.'+n.className.trim().split(/\s+/).slice(0,2).join('.');"
    js = js & "return s;"
    js = js & "};"
    js = js & "window.__bidiDiscoveryObserver=new MutationObserver(function(records){"
    js = js & "var loggedRoots=[];"
    js = js & "var t=((performance.now()-window.__bidiDiscoveryStart)/1000).toFixed(2);"
    js = js & "records.forEach(function(r){"
    js = js & "if(r.type==='childList'){"
    js = js & "r.addedNodes.forEach(function(n){"
    js = js & "if(n.nodeType!==1) return;"
    js = js & "var dup=false;"
    js = js & "for(var i=0;i<loggedRoots.length;i++){ if(loggedRoots[i]!==n && loggedRoots[i].contains && loggedRoots[i].contains(n)){ dup=true; break; } }"
    js = js & "if(dup) return;"
    js = js & "loggedRoots.push(n);"
    js = js & "window.__bidiDiscoveryDom.push('['+t+'s] [DOM-ADD] '+window.__bidiDiscoveryDescribe(n));"
    js = js & "});"
    js = js & "r.removedNodes.forEach(function(n){"
    js = js & "if(n.nodeType!==1) return;"
    js = js & "window.__bidiDiscoveryDom.push('['+t+'s] [DOM-DEL] '+window.__bidiDiscoveryDescribe(n));"
    js = js & "});"
    js = js & "} else if(r.type==='characterData'){"
    js = js & "var tx=(r.target.textContent||'').slice(0,40);"
    js = js & "window.__bidiDiscoveryDom.push('['+t+'s] [DOM-TXT] '+tx);"
    js = js & "}"
    js = js & "});"
    js = js & "});"
    js = js & "window.__bidiDiscoveryObserver.observe(document.documentElement,{childList:true,subtree:true,characterData:true});"
    js = js & "}"

    '失敗(遷移直後で文脈が一瞬消えている等)しても、記録全体は止めない
    ex_BiDiContext.jsEval js, StopBiDiError:=False
End Sub

'***************************************************************************************************
'* 機能　　：JS側に溜まったDOM変化バッファを回収し、タイムラインへ追記します
'***************************************************************************************************
Private Sub DrainDomMutations()
    Dim js As String
    js = "(function(){"
    js = js & "if(!window.__bidiDiscoveryDom) return '';"
    js = js & "var a=window.__bidiDiscoveryDom;"
    js = js & "window.__bidiDiscoveryDom=[];"
    js = js & "return a.join(String.fromCharCode(1));"
    js = js & "})()"

    Dim raw As Variant
    raw = ex_BiDiContext.jsEval(js, StopBiDiError:=False)

    If VarType(raw) <> vbString Then Exit Sub
    If LenB(raw) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(raw, Chr(1))

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        PushTimelineLine parts(i)
    Next i
End Sub

'***************************************************************************************************
'* 機能　　：MutationObserverを解除します
'***************************************************************************************************
Private Sub DetachMutationObserver()
    Dim js As String
    js = "if(window.__bidiDiscoveryObserver){"
    js = js & "window.__bidiDiscoveryObserver.disconnect();"
    js = js & "window.__bidiDiscoveryObserver=null;"
    js = js & "window.__bidiDiscoveryActive=false;"
    js = js & "}"

    ex_BiDiContext.jsEval js, StopBiDiError:=False
End Sub



'***************************************************************************************************
'                           ■■■ ログ整形・蓄積のヘルパー ■■■
'***************************************************************************************************
'* 機能　　：ログ1行をリングバッファへ追加します。上限を超えたら、最も古い行から捨てます
'***************************************************************************************************
Private Sub PushTimelineLine(s As String)
    If LenB(s) = 0 Then Exit Sub

    With CurrentLogState
        If .recorderLogs.Count >= .Limit Then
            .recorderLogs.Remove 1
            .droppedHeadCount = .droppedHeadCount + 1
        End If
        .recorderLogs.Add s
        .totalEventCount = .totalEventCount + 1
    End With
End Sub

'***************************************************************************************************
'* 機能　　：待機系と直接関係ない、静的リソースのURLかをcheckします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`excludeImagesAndCss`が有効な場合にのみ使用される、簡易的な拡張子ベースの判定です
'***************************************************************************************************
Private Function IsNoiseUrl(Url As String) As Boolean
    If LenB(Url) = 0 Then Exit Function

    '【先頭一致の特殊ガード】data: URI は一瞬で判定して即抜け
    If Left$(Url, 5) = "data:" Then IsNoiseUrl = True: Exit Function

    Dim lUrl As String: lUrl = LCase$(Url)

    '【固定NGリストを用意】配列の生成は最初の1回のみ
    Static patterns As Variant
    If IsEmpty(patterns) Then
        patterns = Array(".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", _
                          ".css", ".woff", ".woff2", ".ttf", ".ico", ".map")
    End If

    Dim p As Variant
    For Each p In patterns
        If InStr(1, lUrl, p, vbTextCompare) > 0 Then
            IsNoiseUrl = True
            Exit Function
        End If
    Next p
End Function

'***************************************************************************************************
'* 機能　　：除外サマリー集計用に、URLからクエリ文字列を除いたキーを作ります
'***************************************************************************************************
Private Function NormalizeEndpointKey(Url As String) As String
    Dim Key As String: Key = Url

    Dim qPos As Long: qPos = InStr(1, Key, "?")
    If qPos > 0 Then Key = Left$(Key, qPos - 1)

    If Len(Key) > 120 Then Key = Left$(Key, 120) & "..."

    NormalizeEndpointKey = Key
End Function

'***************************************************************************************************
'* 機能　　：除外したリクエストのキーごとの件数をカウントアップします
'***************************************************************************************************
Private Sub BumpSuppressed(Key As String)
    If CurrentLogState.recSuppByKey Is Nothing Then Exit Sub

    If CurrentLogState.recSuppByKey.Exists(Key) Then
        CurrentLogState.recSuppByKey(Key) = CurrentLogState.recSuppByKey(Key) + 1
    Else
        CurrentLogState.recSuppByKey.Add Key, 1
    End If
End Sub

'***************************************************************************************************
'* 機能　　：除外サマリーのテキストを組み立てます
'***************************************************************************************************
Private Function BuildSuppressedSummary() As String
    Dim result As String
    result = "----- 除外(excludeImagesAndCss)されたリクエストの集計 -----"

    If CurrentLogState.recSuppByKey Is Nothing Then
        BuildSuppressedSummary = result & vbCrLf & "(なし)"
        Exit Function
    End If

    If CurrentLogState.recSuppByKey.Count = 0 Then
        BuildSuppressedSummary = result & vbCrLf & "(なし)"
        Exit Function
    End If

    Dim k As Variant
    For Each k In CurrentLogState.recSuppByKey.keys
        result = result & vbCrLf & CurrentLogState.recSuppByKey(k) & "件  " & k
    Next k

    BuildSuppressedSummary = result
End Function

'***************************************************************************************************
'* 機能　　：蓄積したログ行を、行頭の相対秒をキーに時系列へ並べ替えます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：挿入ソートを使用します。ネットワーク/ナビゲーション/コンソールのログは発生順に、
'            DOM変化のログはポーリング間隔単位でほぼ発生順に追加されるため、全体としてほぼ整列済み
'            のデータになります。そのため挿入ソートでも実用上十分高速です
'***************************************************************************************************
Private Function SortedLinesArray() As Variant
    Dim N As Long: N = CurrentLogState.recorderLogs.Count

    Dim lines() As String: ReDim lines(1 To N)
    Dim keys() As Double: ReDim keys(1 To N)

    Dim i As Long
    For i = 1 To N
        lines(i) = CurrentLogState.recorderLogs(i)
        keys(i) = ParseLeadingSeconds(lines(i))
    Next i

    Dim j As Long
    Dim tmpLine As String
    Dim tmpKey As Double
    For i = 2 To N
        tmpLine = lines(i)
        tmpKey = keys(i)
        j = i - 1
        Do While j >= 1
            If keys(j) <= tmpKey Then Exit Do
            lines(j + 1) = lines(j)
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        lines(j + 1) = tmpLine
        keys(j + 1) = tmpKey
    Next i

    SortedLinesArray = lines
End Function

'***************************************************************************************************
'* 機能　　："[X.XXs] ..."という行の先頭から、数値部分(X.XX)だけを取り出します
'***************************************************************************************************
Private Function ParseLeadingSeconds(s As String) As Double
    If Left$(s, 1) <> "[" Then Exit Function

    Dim closeBracket As Long: closeBracket = InStr(1, s, "s]")
    If closeBracket = 0 Then Exit Function

    Dim numText As String: numText = Mid$(s, 2, closeBracket - 2)

    On Error Resume Next
    ParseLeadingSeconds = CDbl(numText)
    On Error GoTo 0
End Function



'***************************************************************************************************
'                           ■■■ `WebDriverBiDi`直通メソッド ■■■
'***************************************************************************************************
'ログメッセージを出すときにどうぞ
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension As String = ".bidi.log"


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub

'非同期イベント発火(RaiseEvent)に必要です
Private Sub TakeEvents(Optional destruction As Boolean)
    ex_BiDiCore.TakeEvents , destruction
End Sub



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：この拡張機能Classを動かすのに必要なオブジェクトを継承します
'***************************************************************************************************
Friend Sub Init(Inheritance As WebDriverBiDiContext)
    '必要なオブジェクトを継承
    Set ex_BiDiContext = Inheritance
    Set ex_BiDiCore = Inheritance.ThisWebDriverBiDiCore
    CDPHelpers.InheritLogSettings currentLog, Inheritance.ThisWebDriverBiDiCore.DirectControlCDP.ThisCDPCore.LogSettings

    'コンテキスト情報
    m_context = Inheritance.context
End Sub

'***************************************************************************************************
'* 機能　　：初期設定値を適用します
'***************************************************************************************************
Private Sub Class_Initialize()
    CurrentLogState.Limit = 5000

    currentLog.LogID = "DLID" & Format(Rnd * 1000, "000")
End Sub
