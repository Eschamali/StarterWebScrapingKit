Attribute VB_Name = "Test_BiDiAsyncBenchmark_WaitEvents"
'===================================================================================================
' マルチタブ非同期BiDiラウンド同期ベンチマーク（WaitEvents版）
'---------------------------------------------------------------------------------------------------
' 概要：
'   Test_AsyncBenchmark_WaitEvents.bas（CDP版）のWebDriverBiDi版です。
'   ページ読み込み完了の検知には、`browsingContext`イベント経由のreadyState監視
'   （`browsingContext.navigationStarted` / `domContentLoaded` / `load`）を活用する
'   `WebDriverBiDiContext.WaitEvents`を使用しています：
'     ・BiDiEvents（イベント蓄積用Dictionary）は使用しない
'     ・追加の拡張クラスは不要
'     ・`WaitEvents(WaitMode:=isComplete)`による、タブごとの単純な待機のみ
'   ※上記3イベントの購読自体は、`WebDriverBiDiContext`生成時に内部（`EnableDiscoverContexts`）で
'     自動的に行われるため、CDP版の`Page.enable`同様、本ファイル側で明示的な購読操作は不要です。
'
'   流れは以下のとおり：
'     Step A：全タブへ`navigate(url, till:=Nowait)`で非同期遷移を発行
'             （内部で`ResetWaitState`も行われるため、前ラウンドの状態は自動的にクリアされる。
'             `till:=Nowait`により、CDP版の`WaitMode:=Nowait`と同様、結果を待たず即座に発行だけ行う）
'     Step B：発行済みの全タブに対して、1つずつ`WaitEvents`で読み込み完了を待つ
'             （`WaitEvents`内部の`TakeEvents`は、同一パイプを共有する全タブ分のイベントを
'               まとめて汲み上げるため、1つずつ待っても他タブの読み込みは並行して進行する）
'     Step C・D：バリアを通過したタブへCookie/Screenshotを非同期発行し、Screenshotのみこの
'                ラウンド内で回収する（Cookieは全ラウンド終了後にまとめて回収）
'
'   `UseWebSocket`定数で、Pipe/WebSocketの接続方式を切り替え可能です（WebView2はBiDiの
'   トランスポートとして対応していないため、CDP版のような3値ではなくBoolean切り替えのみです）。
'
'   BiDiコマンド：
'     ・Cookie      → `storage.getCookies`（`partition`に`{type:"context", context:...}`を指定）
'     ・Screenshot  → `browsingContext.captureScreenshot`
'   結果の生JSONは、`ExecuteBiDi`同様`BiDiCDPJson.Parse`で解析し、`"type"="error"`かどうかで
'   成否判定します（CDP版の`{"error":...}/{"result":...}`とはJSON構造が異なる点に注意）。
'===================================================================================================
Option Explicit

Private Const UseWebSocket As Boolean = False   ' False:Pipe / True:WebSocket

Private Const RESULT_SECTION_LINE    As String = "=================================================="
Private Const NUM_TABS               As Long = 30     ' 開くタブ数
Private Const NUM_ROUNDS             As Long = 10      ' 繰り返すラウンド数
Private Const TIMEOUT_LOAD_SEC       As Double = 60   ' 読み込み完了待機のタイムアウト（`TimeOutSecond`に適用）
Private Const TIMEOUT_SCREENSHOT_SEC As Double = 60   ' スクショ回収のタイムアウト
Private Const TIMEOUT_COOKIE_SEC     As Double = 60   ' Cookie回収のタイムアウト
Private Const SAVE_PATH              As String = "Downloads"

Private Const URL_1 As String = "https://www.youtube.com/@islandfox6864/"
Private Const URL_2 As String = "https://www.yahoo.co.jp"
Private Const URL_3 As String = "https://kemono-friends.jp/"
Private Const URL_4 As String = "https://news.yahoo.co.jp"
Private Const URL_5 As String = "https://www.amazon.co.jp"

Private Type TabState
    Index As Long
    TimedOutRounds As Long
End Type

Private Type ScreenshotTicket   ' ラウンドごとに作り直す（このラウンド内でのみ使う）
    TabIndex As Long
    context As WebDriverBiDiContext
    commandID As Long
    Retrieved As Boolean
End Type

Private Type ScreenshotPayload  ' 全ラウンド分を蓄積。デコード・保存は最後にまとめて行う
    TabIndex As Long
    RoundIndex As Long
    Base64Data As String
    FileName As String
    HadError As Boolean
End Type

Private Type CookieTicket       ' 全ラウンド分を蓄積。結果取得も最後まで遅延する
    TabIndex As Long
    RoundIndex As Long
    context As WebDriverBiDiContext
    commandID As Long
    Retrieved As Boolean
    CookieCount As Long
    HadError As Boolean
End Type

'===================================================================================================
' WaitEvents版：ページ読み込み完了の検知を`WebDriverBiDiContext.WaitEvents`に任せる、簡略化パターン
'===================================================================================================
Public Sub Test_BiDiAsyncBenchmark_WaitEvents()
    Dim mode As WebDriverBiDiMode
    Dim tabs() As WebDriverBiDiContext
    Dim tabStates() As TabState
    Dim urls(1 To 5) As String
    Dim t As Long, r As Long
    Dim cookieTickets() As CookieTicket, cookieTicketCount As Long
    Dim screenshotPayloads() As ScreenshotPayload, screenshotPayloadCount As Long
    Dim benchStart As Double

    urls(1) = URL_1: urls(2) = URL_2: urls(3) = URL_3: urls(4) = URL_4: urls(5) = URL_5

    PrintHeader "[BiDi/WaitEventsパターン] マルチタブ非同期ベンチマーク 開始"
    Debug.Print "設定: タブ数=" & NUM_TABS & ", ラウンド数=" & NUM_ROUNDS & ", WebSocket=" & UseWebSocket

    ReDim tabs(1 To NUM_TABS)
    If UseWebSocket Then
        Set mode = ShSetting01_StartBrowser.StartBiDiMode(WebSocketMode:=True)
    Else
        Set mode = ShSetting01_StartBrowser.StartBiDiMode
    End If
    Set tabs(1) = mode.getTab(setMain:=True)

    benchStart = CDPHelpers.TimerCounter

    ReDim tabStates(1 To NUM_TABS)

    For t = 2 To NUM_TABS
        Set tabs(t) = mode.newTab(newWindow:=False)
    Next t

    Randomize
    For t = 1 To NUM_TABS
        tabStates(t).Index = t
    Next t
    mode.TimeOutSecond = TIMEOUT_LOAD_SEC   ' `WaitEvents`のタイムアウトは、Mode単位でこれに依存する

    ReDim cookieTickets(1 To NUM_TABS * NUM_ROUNDS)
    ReDim screenshotPayloads(1 To NUM_TABS * NUM_ROUNDS)

    For r = 1 To NUM_ROUNDS
        Debug.Print RESULT_SECTION_LINE
        Debug.Print "[Round " & r & "/" & NUM_ROUNDS & "] 開始"

        Dim activeTabs() As Boolean
        ReDim activeTabs(1 To NUM_TABS)

        ' --- Step A: 全タブ一斉に非同期遷移 ---
        ' `till:=Nowait`により、内部で`ResetWaitState`（今回分の読み込み状況リセット）＋
        ' `browsingContext.navigate`の非同期発行のみを行い、結果を待たず即座に次のタブへ進む
        For t = 1 To NUM_TABS
            tabs(t).navigate urls(Int(Rnd * 5) + 1), till:=Nowait
        Next t

        ' --- Step B: 各タブの読み込み完了を、`WaitEvents`で1つずつ確認 ---
        For t = 1 To NUM_TABS
            If tabs(t).WaitEvents(WaitMode:=isComplete, WaitError:=False) Then
                activeTabs(t) = True
                Debug.Print "  Tab " & t & " 読み込み完了"
            Else
                tabStates(t).TimedOutRounds = tabStates(t).TimedOutRounds + 1
                activeTabs(t) = False
                Debug.Print "  [WARN] Tab " & t & " 読み込みタイムアウト。このラウンドはスキップします"
            End If
        Next t

        ' --- Step C・D: Cookie/Screenshot発行・このラウンドのScreenshot回収（共通処理） ---
        FireAndDrainRound mode, tabs, activeTabs, r, cookieTickets, cookieTicketCount, screenshotPayloads, screenshotPayloadCount
    Next r

    FinishBenchmark mode, benchStart, tabStates, cookieTickets, cookieTicketCount, screenshotPayloads, screenshotPayloadCount
End Sub

'===================================================================================================
' 共通ヘルパー
'===================================================================================================

'---------------------------------------------------------------------------------------------------
' Step C・D: バリアを通過したタブへCookie/Screenshotを非同期発行し、Screenshotだけこのラウンド内で回収する
' （Cookieは結果を取りに行かず、整理券をcookieTicketsに積むだけ。取得はFinishBenchmarkで最後にまとめて行う）
'---------------------------------------------------------------------------------------------------
Private Sub FireAndDrainRound(mode As WebDriverBiDiMode, tabs() As WebDriverBiDiContext, activeTabs() As Boolean, r As Long, _
                               ByRef cookieTickets() As CookieTicket, ByRef cookieTicketCount As Long, _
                               ByRef screenshotPayloads() As ScreenshotPayload, ByRef screenshotPayloadCount As Long)
    Dim t As Long, i As Long

    ' Step C: Cookie・Screenshot非同期発行
    Dim screenshotTickets() As ScreenshotTicket
    ReDim screenshotTickets(1 To NUM_TABS)
    Dim screenshotTicketCount As Long

    Dim cookieParams As Dictionary, partition As Dictionary
    For t = 1 To NUM_TABS
        If activeTabs(t) Then
            ' `storage.getCookies`は`context`ではなく`partition`引数でタブ範囲を絞る必要があるため、
            ' `context`を自動付与する`WebDriverBiDiContext.ExecuteBiDiAsync`は使わず、Mode経由で組み立てる
            Set cookieParams = New Dictionary
            Set partition = New Dictionary
            partition.Add "type", "context"
            partition.Add "context", tabs(t).context
            cookieParams.Add "partition", partition

            cookieTicketCount = cookieTicketCount + 1
            cookieTickets(cookieTicketCount).TabIndex = t
            cookieTickets(cookieTicketCount).RoundIndex = r
            Set cookieTickets(cookieTicketCount).context = tabs(t)
            cookieTickets(cookieTicketCount).commandID = mode.ExecuteBiDiAsync("storage.getCookies", cookieParams)
            Debug.Print "  Tab " & t & " Cookie非同期要求発行 (整理券:" & cookieTickets(cookieTicketCount).commandID & ")"

            screenshotTicketCount = screenshotTicketCount + 1
            screenshotTickets(screenshotTicketCount).TabIndex = t
            Set screenshotTickets(screenshotTicketCount).context = tabs(t)
            screenshotTickets(screenshotTicketCount).commandID = tabs(t).ExecuteBiDiAsync("browsingContext.captureScreenshot", Nothing)
            Debug.Print "  Tab " & t & " Screenshot非同期要求発行 (整理券:" & screenshotTickets(screenshotTicketCount).commandID & ")"
        End If
    Next t

    ' Step D: このラウンドのScreenshot結果をまとめて取り出す（デコード・保存はまだしない）
    Dim drainStart As Double: drainStart = CDPHelpers.TimerCounter
    Dim remaining As Long: remaining = screenshotTicketCount

    Do While remaining > 0
        mode.TakeEvents

        For i = 1 To screenshotTicketCount
            If Not screenshotTickets(i).Retrieved Then
                Dim resJson As String
                resJson = mode.TakeResultBiDi(screenshotTickets(i).commandID)

                If Len(resJson) > 0 Then
                    screenshotTickets(i).Retrieved = True
                    remaining = remaining - 1

                    screenshotPayloadCount = screenshotPayloadCount + 1
                    screenshotPayloads(screenshotPayloadCount).TabIndex = screenshotTickets(i).TabIndex
                    screenshotPayloads(screenshotPayloadCount).RoundIndex = r
                    screenshotPayloads(screenshotPayloadCount).FileName = "bidi_bench_tab" & screenshotTickets(i).TabIndex & "_round" & r & ".png"

                    Dim resNode As BiDiCDPJson
                    Set resNode = BiDiCDPJson.Parse(resJson)
                    If resNode.StringKey("type") = "error" Then
                        screenshotPayloads(screenshotPayloadCount).HadError = True
                        Debug.Print "  [WARN] Tab " & screenshotTickets(i).TabIndex & " Round " & r & " スクショ取得エラー: " & resNode.StringKey("message")
                    Else
                        screenshotPayloads(screenshotPayloadCount).Base64Data = resNode.NodeKey("result").StringKey("data")
                        Debug.Print "  Tab " & screenshotTickets(i).TabIndex & " Round " & r & " スクショ取得完了"
                    End If

                ElseIf CDPHelpers.TimerCounter - drainStart > TIMEOUT_SCREENSHOT_SEC * 1000 Then
                    screenshotTickets(i).Retrieved = True
                    remaining = remaining - 1

                    screenshotPayloadCount = screenshotPayloadCount + 1
                    screenshotPayloads(screenshotPayloadCount).TabIndex = screenshotTickets(i).TabIndex
                    screenshotPayloads(screenshotPayloadCount).RoundIndex = r
                    screenshotPayloads(screenshotPayloadCount).HadError = True
                    Debug.Print "  [WARN] Tab " & screenshotTickets(i).TabIndex & " Round " & r & " スクショ取得タイムアウト"
                End If
            End If
        Next i
    Loop
End Sub

'---------------------------------------------------------------------------------------------------
' 全ラウンド終了後: Cookie一括取得 → Screenshot一括保存 → サマリー出力 → ブラウザ終了
'---------------------------------------------------------------------------------------------------
Private Sub FinishBenchmark(mode As WebDriverBiDiMode, benchStart As Double, ByRef tabStates() As TabState, _
                             ByRef cookieTickets() As CookieTicket, cookieTicketCount As Long, _
                             ByRef screenshotPayloads() As ScreenshotPayload, screenshotPayloadCount As Long)
    Debug.Print RESULT_SECTION_LINE
    Debug.Print "全ラウンドの遷移・要求が完了しました。Cookie一括取得フェーズに移ります..."
    DrainAllCookies mode, cookieTickets, cookieTicketCount

    Debug.Print RESULT_SECTION_LINE
    Debug.Print "Screenshot一括保存フェーズに移ります..."
    Dim saveDir As String: saveDir = Environ("UserProfile") & "\" & SAVE_PATH
    Dim savedCount As Long: savedCount = SaveAllScreenshots(screenshotPayloads, screenshotPayloadCount, saveDir)

    Dim t As Long, i As Long, cookieSum As Long
    For i = 1 To cookieTicketCount
        cookieSum = cookieSum + cookieTickets(i).CookieCount
    Next i

    PrintHeader "[BiDi/WaitEventsパターン] ベンチマーク結果"
    Debug.Print "  タブ数               : " & NUM_TABS
    Debug.Print "  ラウンド数            : " & NUM_ROUNDS
    Debug.Print "  接続方式             : " & IIf(UseWebSocket, "WebSocket", "Pipe")
    Debug.Print "  経過時間             : " & Format((CDPHelpers.TimerCounter - benchStart) / 1000, "0.0") & " 秒"
    For t = 1 To NUM_TABS
        Debug.Print "  Tab " & t & " タイムアウト回数    : " & tabStates(t).TimedOutRounds & " / " & NUM_ROUNDS & " ラウンド"
    Next t
    Debug.Print "  Cookie取得チケット数  : " & cookieTicketCount & " (Cookie総数: " & cookieSum & ")"
    Debug.Print "  Screenshot保存数     : " & savedCount & " / " & screenshotPayloadCount
    Debug.Print "  Screenshot保存先     : " & saveDir
    Debug.Print RESULT_SECTION_LINE

    mode.quit
End Sub

'---------------------------------------------------------------------------------------------------
' 全ラウンド分のCookie整理券を、まとめて取得する（結果が来ていない間は待つ、タイムアウトで諦める）
'---------------------------------------------------------------------------------------------------
Private Sub DrainAllCookies(mode As WebDriverBiDiMode, ByRef cookieTickets() As CookieTicket, cookieTicketCount As Long)
    Dim drainStart As Double: drainStart = CDPHelpers.TimerCounter
    Dim remaining As Long: remaining = cookieTicketCount
    Dim i As Long

    Do While remaining > 0
        mode.TakeEvents

        For i = 1 To cookieTicketCount
            If Not cookieTickets(i).Retrieved Then
                Dim resJson As String
                resJson = mode.TakeResultBiDi(cookieTickets(i).commandID)

                If Len(resJson) > 0 Then
                    cookieTickets(i).Retrieved = True
                    remaining = remaining - 1

                    Dim resNode As BiDiCDPJson
                    Set resNode = BiDiCDPJson.Parse(resJson)
                    If resNode.StringKey("type") = "error" Then
                        cookieTickets(i).HadError = True
                    Else
                        cookieTickets(i).CookieCount = resNode.NodeKey("result").NodeKey("cookies").Count
                    End If

                ElseIf CDPHelpers.TimerCounter - drainStart > TIMEOUT_COOKIE_SEC * 1000 Then
                    cookieTickets(i).Retrieved = True
                    cookieTickets(i).HadError = True
                    remaining = remaining - 1
                    Debug.Print "  [WARN] Tab " & cookieTickets(i).TabIndex & " Round " & cookieTickets(i).RoundIndex & " Cookie取得タイムアウト"
                End If
            End If
        Next i
    Loop
End Sub

'---------------------------------------------------------------------------------------------------
' 蓄積済みのScreenshot(Base64)を、まとめてデコード・Downloadsフォルダへ保存する
'---------------------------------------------------------------------------------------------------
Private Function SaveAllScreenshots(ByRef screenshotPayloads() As ScreenshotPayload, screenshotPayloadCount As Long, saveDir As String) As Long
    Dim DataConv As New WebCrypto
    Dim CharConv As New CharacterCodeConversion
    Dim i As Long, savedCount As Long

    For i = 1 To screenshotPayloadCount
        If Not screenshotPayloads(i).HadError And Len(screenshotPayloads(i).Base64Data) > 0 Then
            Dim Bytes() As Byte
            Bytes = DataConv.Decode(screenshotPayloads(i).Base64Data, edfBase64)
            CharConv.BytesToSaveFile Bytes, saveDir, screenshotPayloads(i).FileName
            savedCount = savedCount + 1
        End If
    Next i

    SaveAllScreenshots = savedCount
End Function

Private Sub PrintHeader(msg As String)
    Debug.Print ""
    Debug.Print RESULT_SECTION_LINE
    Debug.Print "  " & msg
    Debug.Print RESULT_SECTION_LINE
End Sub
