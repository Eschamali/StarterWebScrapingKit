Attribute VB_Name = "Test_AsyncBenchmark_WaitEvents"
'===================================================================================================
' マルチタブ非同期CDPラウンド同期ベンチマーク（WaitEvents版）
'---------------------------------------------------------------------------------------------------
' 概要：
'   Test_AsyncBenchmark.bas（Inline / ClassBased の2パターン）の亜種です。
'   ページ読み込み完了の検知を、新しく追加された`Page`イベント経由のreadyState監視
'   （`Page.frameStartedLoading` / `Page.domContentEventFired` / `Page.loadEventFired`）を
'   活用する`CDPContext.WaitEvents`に置き換え、以下のとおり簡略化しています：
'     ・BrowserEvents（イベント蓄積用Dictionary）は使用しない
'     ・追加の拡張クラス（exCDP_PageLoadWatcher等）は不要
'     ・`WaitEvents(WaitMode:=isComplete)`による、タブごとの単純な待機のみ
'
'   流れは以下のとおり：
'     Step A：全タブへ`navigate(url, WaitMode:=Nowait)`で非同期遷移を発行
'             （内部で`ResetWaitState`も行われるため、前ラウンドの状態は自動的にクリアされる）
'     Step B：発行済みの全タブに対して、1つずつ`WaitEvents`で読み込み完了を待つ
'             （`WaitEvents`内部の`TakeEvents`は、同一パイプを共有する全タブ分のイベントを
'               まとめて汲み上げるため、1つずつ待っても他タブの読み込みは並行して進行する）
'     Step C・D：バリアを通過したタブへCookie/Screenshotを非同期発行し、Screenshotのみこの
'                ラウンド内で回収する（Cookieは全ラウンド終了後にまとめて回収）
'
'   ※旧版にあった`Network.requestWillBeSent`の件数集計は、BrowserEvents依存の機能だったため、
'     このバージョンでは行いません。
'
'   `TestType`定数で、Pipe/WebSocket/WebView2の接続方式を切り替え可能です（Test_AsyncBenchmark.bas
'   と同じ切り替え方）。WebView2はイベント名ごとの個別購読が必要なモデルのため、`WaitEvents`が
'   見ている3イベント（Page.frameStartedLoading/domContentEventFired/loadEventFired）を
'   `SubscribeCdpEvent`で明示的に購読しています。
'===================================================================================================
Option Explicit

Private Const WebView2Mode As Boolean = False

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
    context As CDPContext
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
    context As CDPContext
    commandID As Long
    Retrieved As Boolean
    CookieCount As Long
    HadError As Boolean
End Type

'===================================================================================================
' WaitEvents版：ページ読み込み完了の検知を`CDPContext.WaitEvents`に任せる、簡略化パターン
'===================================================================================================
Public Sub Test_AsyncBenchmark_WaitEvents()
    Dim chrome As CDPBrowser
    Dim tabs() As CDPContext
    Dim tabStates() As TabState
    Dim urls(1 To 5) As String
    Dim t As Long, r As Long
    Dim cookieTickets() As CookieTicket, cookieTicketCount As Long
    Dim screenshotPayloads() As ScreenshotPayload, screenshotPayloadCount As Long
    Dim benchStart As Double

    urls(1) = URL_1: urls(2) = URL_2: urls(3) = URL_3: urls(4) = URL_4: urls(5) = URL_5

    PrintHeader "[WaitEventsパターン] マルチタブ非同期ベンチマーク 開始"
    Debug.Print "設定: タブ数=" & NUM_TABS & ", ラウンド数=" & NUM_ROUNDS

    ReDim tabs(1 To NUM_TABS)
    Set chrome = New CDPBrowser
    If WebView2Mode Then
        With WebView2Form
            If Not .StartCDPModeWebView2 Then Debug.Print "WebView2起動失敗": Exit Sub
            Set tabs(1) = .ThisCDPContext
            Set chrome = tabs(1).ThisCDPBrowser

            '1つ目のフォームを表示
            .show False
        End With
    Else
        Set chrome = ShSetting01_StartBrowser.StartCDPMode
        Set tabs(1) = chrome.getTab(setMain:=True)
    End If

    benchStart = CDPHelpers.TimerCounter

    ReDim tabStates(1 To NUM_TABS)

    For t = 2 To NUM_TABS
        Set tabs(t) = chrome.newTab(newWindow:=False)
    Next t

    Randomize
    For t = 1 To NUM_TABS
        tabStates(t).Index = t
        tabs(t).TimeOutSecond = TIMEOUT_LOAD_SEC   ' `WaitEvents`のタイムアウトは、これに依存する
    Next t

    ReDim cookieTickets(1 To NUM_TABS * NUM_ROUNDS)
    ReDim screenshotPayloads(1 To NUM_TABS * NUM_ROUNDS)

    For r = 1 To NUM_ROUNDS
        Debug.Print RESULT_SECTION_LINE
        Debug.Print "[Round " & r & "/" & NUM_ROUNDS & "] 開始"

        Dim activeTabs() As Boolean
        ReDim activeTabs(1 To NUM_TABS)

        ' --- Step A: 全タブ一斉に非同期遷移 ---
        ' `WaitMode:=Nowait`により、内部で`ResetWaitState`（今回分の読み込み状況リセット）＋
        ' `Page.navigate`の非同期発行のみを行い、結果を待たず即座に次のタブへ進む
        For t = 1 To NUM_TABS
            tabs(t).navigate urls(Int(Rnd * 5) + 1), WaitMode:=Nowait
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
        FireAndDrainRound chrome, tabs, activeTabs, r, cookieTickets, cookieTicketCount, screenshotPayloads, screenshotPayloadCount
    Next r

    FinishBenchmark chrome, benchStart, tabStates, cookieTickets, cookieTicketCount, screenshotPayloads, screenshotPayloadCount
End Sub

'===================================================================================================
' 共通ヘルパー
'===================================================================================================

'---------------------------------------------------------------------------------------------------
' Step C・D: バリアを通過したタブへCookie/Screenshotを非同期発行し、Screenshotだけこのラウンド内で回収する
' （Cookieは結果を取りに行かず、整理券をcookieTicketsに積むだけ。取得はFinishBenchmarkで最後にまとめて行う）
'---------------------------------------------------------------------------------------------------
Private Sub FireAndDrainRound(chrome As CDPBrowser, tabs() As CDPContext, activeTabs() As Boolean, r As Long, _
                               ByRef cookieTickets() As CookieTicket, ByRef cookieTicketCount As Long, _
                               ByRef screenshotPayloads() As ScreenshotPayload, ByRef screenshotPayloadCount As Long)
    Dim t As Long, i As Long

    ' Step C: Cookie・Screenshot非同期発行
    Dim screenshotTickets() As ScreenshotTicket
    ReDim screenshotTickets(1 To NUM_TABS)
    Dim screenshotTicketCount As Long

    For t = 1 To NUM_TABS
        If activeTabs(t) Then
            cookieTicketCount = cookieTicketCount + 1
            cookieTickets(cookieTicketCount).TabIndex = t
            cookieTickets(cookieTicketCount).RoundIndex = r
            Set cookieTickets(cookieTicketCount).context = tabs(t)
            cookieTickets(cookieTicketCount).commandID = tabs(t).ExecuteCDPAsync("Network.getAllCookies", Nothing)
            Debug.Print "  Tab " & t & " Cookie非同期要求発行 (整理券:" & cookieTickets(cookieTicketCount).commandID & ")"

            screenshotTicketCount = screenshotTicketCount + 1
            screenshotTickets(screenshotTicketCount).TabIndex = t
            Set screenshotTickets(screenshotTicketCount).context = tabs(t)
            screenshotTickets(screenshotTicketCount).commandID = tabs(t).ExecuteCDPAsync("Page.captureScreenshot", Nothing)
            Debug.Print "  Tab " & t & " Screenshot非同期要求発行 (整理券:" & screenshotTickets(screenshotTicketCount).commandID & ")"
        End If
    Next t

    ' Step D: このラウンドのScreenshot結果をまとめて取り出す（デコード・保存はまだしない）
    Dim drainStart As Double: drainStart = CDPHelpers.TimerCounter
    Dim remaining As Long: remaining = screenshotTicketCount

    Do While remaining > 0
        chrome.TakeEvents

        For i = 1 To screenshotTicketCount
            If Not screenshotTickets(i).Retrieved Then
                Dim resJson As String
                resJson = screenshotTickets(i).context.TakeResultCDP(screenshotTickets(i).commandID)

                If Len(resJson) > 0 Then
                    screenshotTickets(i).Retrieved = True
                    remaining = remaining - 1

                    screenshotPayloadCount = screenshotPayloadCount + 1
                    screenshotPayloads(screenshotPayloadCount).TabIndex = screenshotTickets(i).TabIndex
                    screenshotPayloads(screenshotPayloadCount).RoundIndex = r
                    screenshotPayloads(screenshotPayloadCount).FileName = "bench_tab" & screenshotTickets(i).TabIndex & "_round" & r & ".png"

                    Dim resDic As Dictionary
                    Set resDic = WebJsonConverter.Parse(resJson).value
                    If resDic.Exists("error") Then
                        screenshotPayloads(screenshotPayloadCount).HadError = True
                        Debug.Print "  [WARN] Tab " & screenshotTickets(i).TabIndex & " Round " & r & " スクショ取得エラー: " & resDic("error")("message")
                    ElseIf resDic.Exists("result") Then
                        If resDic("result").Exists("data") Then
                            screenshotPayloads(screenshotPayloadCount).Base64Data = resDic("result")("data")
                            Debug.Print "  Tab " & screenshotTickets(i).TabIndex & " Round " & r & " スクショ取得完了"
                        End If
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
Private Sub FinishBenchmark(chrome As CDPBrowser, benchStart As Double, ByRef tabStates() As TabState, _
                             ByRef cookieTickets() As CookieTicket, cookieTicketCount As Long, _
                             ByRef screenshotPayloads() As ScreenshotPayload, screenshotPayloadCount As Long)
    Debug.Print RESULT_SECTION_LINE
    Debug.Print "全ラウンドの遷移・要求が完了しました。Cookie一括取得フェーズに移ります..."
    DrainAllCookies chrome, cookieTickets, cookieTicketCount

    Debug.Print RESULT_SECTION_LINE
    Debug.Print "Screenshot一括保存フェーズに移ります..."
    Dim saveDir As String: saveDir = Environ("UserProfile") & "\" & SAVE_PATH
    Dim savedCount As Long: savedCount = SaveAllScreenshots(screenshotPayloads, screenshotPayloadCount, saveDir)

    Dim t As Long, i As Long, cookieSum As Long
    For i = 1 To cookieTicketCount
        cookieSum = cookieSum + cookieTickets(i).CookieCount
    Next i

    PrintHeader "[WaitEventsパターン] ベンチマーク結果"
    Debug.Print "  タブ数               : " & NUM_TABS
    Debug.Print "  ラウンド数            : " & NUM_ROUNDS
    Debug.Print "  経過時間             : " & Format((CDPHelpers.TimerCounter - benchStart) / 1000, "0.0") & " 秒"
    For t = 1 To NUM_TABS
        Debug.Print "  Tab " & t & " タイムアウト回数    : " & tabStates(t).TimedOutRounds & " / " & NUM_ROUNDS & " ラウンド"
    Next t
    Debug.Print "  Cookie取得チケット数  : " & cookieTicketCount & " (Cookie総数: " & cookieSum & ")"
    Debug.Print "  Screenshot保存数     : " & savedCount & " / " & screenshotPayloadCount
    Debug.Print "  Screenshot保存先     : " & saveDir
    Debug.Print RESULT_SECTION_LINE

    If WebView2Mode Then WebView2Form.hide
    chrome.quit
    If WebView2Mode Then Unload WebView2Form
End Sub

'---------------------------------------------------------------------------------------------------
' 全ラウンド分のCookie整理券を、まとめて取得する（結果が来ていない間は待つ、タイムアウトで諦める）
'---------------------------------------------------------------------------------------------------
Private Sub DrainAllCookies(chrome As CDPBrowser, ByRef cookieTickets() As CookieTicket, cookieTicketCount As Long)
    Dim drainStart As Double: drainStart = CDPHelpers.TimerCounter
    Dim remaining As Long: remaining = cookieTicketCount
    Dim i As Long

    Do While remaining > 0
        chrome.TakeEvents

        For i = 1 To cookieTicketCount
            If Not cookieTickets(i).Retrieved Then
                Dim resJson As String
                resJson = cookieTickets(i).context.TakeResultCDP(cookieTickets(i).commandID)

                If Len(resJson) > 0 Then
                    cookieTickets(i).Retrieved = True
                    remaining = remaining - 1

                    Dim resDic As Dictionary
                    Set resDic = WebJsonConverter.Parse(resJson).value
                    If resDic.Exists("error") Then
                        cookieTickets(i).HadError = True
                    ElseIf resDic.Exists("result") Then
                        If resDic("result").Exists("cookies") Then cookieTickets(i).CookieCount = resDic("result")("cookies").Count
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
