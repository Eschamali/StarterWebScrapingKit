'===================================================================================================
' マルチタブ非同期CDPラウンド同期ベンチマーク
'---------------------------------------------------------------------------------------------------
' 概要：
'   指定タブ数を開き、ラウンドごとに「全タブ一斉ランダム遷移 → 全タブ読み込み完了バリア →
'   Cookie/Screenshot非同期要求」を繰り返す。Screenshotはラウンドごとに結果を回収して保持し、
'   Cookieは全ラウンド終了後にまとめて回収する。最後にScreenshotを一括でDownloadsへ保存する。
'
'   ページ読み込み完了検知（バリア判定, Step B）は、以下の2パターンで用意している：
'     ・Test_AsyncBenchmark_RoundSync_Inline     : CDPContext.BrowserEventsを直接ポーリング（追加クラス不要）
'     ・Test_AsyncBenchmark_RoundSync_ClassBased : exCDP_PageLoadWatcher拡張クラスを利用（要インポート）
'   両者はStep A/Bの判定部分のみが異なり、それ以外（Cookie/Screenshot発行・回収・保存・サマリー）は
'   共通の非公開ヘルパーを呼び出している。
'===================================================================================================
Option Explicit

Private Const RESULT_SECTION_LINE   As String = "=================================================="
Private Const NUM_TABS              As Long = 30      ' 開くタブ数
Private Const NUM_ROUNDS            As Long = 10       ' 繰り返すラウンド数
Private Const TIMEOUT_LOAD_SEC      As Double = 30    ' 読み込み完了バリアのタイムアウト
Private Const TIMEOUT_SCREENSHOT_SEC As Double = 20   ' スクショ回収のタイムアウト
Private Const TIMEOUT_COOKIE_SEC    As Double = 15    ' Cookie回収のタイムアウト
Private Const SAVE_PATH             As String = "Downloads"

Private Const URL_1 As String = "https://www.youtube.com/@islandfox6864/"
Private Const URL_2 As String = "https://www.yahoo.co.jp"
Private Const URL_3 As String = "https://kemono-friends.jp/"
Private Const URL_4 As String = "https://news.yahoo.co.jp"
Private Const URL_5 As String = "https://www.amazon.co.jp"

Private Type TabState
    Index As Long
    LoadedThisRound As Boolean
    TimedOutRounds As Long
    NetworkEventCountThisRound As Long   ' バリア成立時点でのNetwork.requestWillBeSent件数
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
' Pattern A: CDPContext.BrowserEvents を直接ポーリングする、bas完結パターン
'===================================================================================================
Public Sub Test_AsyncBenchmark_RoundSync_Inline()
    Dim chrome As CDPBrowser
    Dim tabs() As CDPContext
    Dim tabStates() As TabState
    Dim urls(1 To 5) As String
    Dim t As Long, r As Long
    Dim cookieTickets() As CookieTicket, cookieTicketCount As Long
    Dim screenshotPayloads() As ScreenshotPayload, screenshotPayloadCount As Long
    Dim benchStart As Double

    urls(1) = URL_1: urls(2) = URL_2: urls(3) = URL_3: urls(4) = URL_4: urls(5) = URL_5

    PrintHeader "[Inlineパターン] マルチタブ非同期ベンチマーク 開始"
    Debug.Print "設定: タブ数=" & NUM_TABS & ", ラウンド数=" & NUM_ROUNDS

    ReDim tabs(1 To NUM_TABS)
    '---- Pipe版 ----
    Set chrome = ShSetting01_StartBrowser.StartCDPMode
    Set tabs(1) = chrome.getTab(setMain:=True)
    '----------------
    
    '---- WebSocket版 ----
'    '3. 設定セルから、ユーザ名を取得
'    Dim UserName As String
'    UserName = ShSetting01_StartBrowser.CurrentUserName
'
'    '4. 指定のWebSocketForCDPへ接続
'    Dim WebSocketCDP As New CDPCoreViaWebSocket
'    Debug.Print WebSocketCDP.AutoConnectBrowserCDP(UserName)
'
'    '5. 繋げたWebSocketオブジェクトを`reattach`メソッドに渡す
'    Set chrome = New CDPBrowser
'    If Not chrome.reattach(UserName, WebSocketCDP) Then MsgBox "「" & UserName & "」に接続できませんでした。WebSocket情報がお亡くなりです。", vbCritical, "Chrome DevTools Protocol": Exit Sub
'    Set tabs(1) = chrome.newTab(setMain:=True)
    '---------------------
    benchStart = chrome.TimerCounter



    ReDim tabStates(1 To NUM_TABS)

    For t = 2 To NUM_TABS
        Set tabs(t) = chrome.newTab(newWindow:=False)
    Next t

    Randomize
    For t = 1 To NUM_TABS
        tabStates(t).Index = t

        tabs(t).ExecuteCDP "Page.enable"
        tabs(t).ExecuteCDP "Network.enable"
        tabs(t).SetFilterEvents = "Page.loadEventFired"
        tabs(t).SetFilterEvents = "Network.requestWillBeSent"
    Next t

    ReDim cookieTickets(1 To NUM_TABS * NUM_ROUNDS)
    ReDim screenshotPayloads(1 To NUM_TABS * NUM_ROUNDS)

    For r = 1 To NUM_ROUNDS
        Debug.Print RESULT_SECTION_LINE
        Debug.Print "[Round " & r & "/" & NUM_ROUNDS & "] 開始"

        Dim activeTabs() As Boolean
        ReDim activeTabs(1 To NUM_TABS)

        ' --- Step A: 全タブ一斉に非同期遷移 ---
        For t = 1 To NUM_TABS
            Set tabs(t).BrowserEvents = New Dictionary   ' 前ラウンドの残留イベントをクリアしてから発行
            tabStates(t).LoadedThisRound = False

            Dim navParams As Scripting.Dictionary
            Set navParams = New Scripting.Dictionary
            navParams.Add "url", urls(Int(Rnd * 5) + 1)
            tabs(t).ExecuteCDPAsync "Page.navigate", navParams
        Next t

        ' --- Step B: 全タブ読み込み完了バリア ---
        Dim barrierStart As Double: barrierStart = chrome.TimerCounter
        Dim pendingCount As Long: pendingCount = NUM_TABS
        Do While pendingCount > 0
            chrome.TakeEvents

            For t = 1 To NUM_TABS
                If Not tabStates(t).LoadedThisRound Then
                    If tabs(t).BrowserEvents("EventMethods").Exists("Page.loadEventFired") Then
                        tabStates(t).LoadedThisRound = True
                        If tabs(t).BrowserEvents("EventMethods").Exists("Network.requestWillBeSent") Then
                            tabStates(t).NetworkEventCountThisRound = tabs(t).BrowserEvents("EventMethods")("Network.requestWillBeSent").Count
                        Else
                            tabStates(t).NetworkEventCountThisRound = 0
                        End If
                        activeTabs(t) = True
                        pendingCount = pendingCount - 1
                        Debug.Print "  Tab " & t & " 読み込み完了 (NetworkEvents=" & tabStates(t).NetworkEventCountThisRound & ")"

                    ElseIf chrome.TimerCounter - barrierStart > TIMEOUT_LOAD_SEC * 1000 Then
                        tabStates(t).LoadedThisRound = True
                        tabStates(t).TimedOutRounds = tabStates(t).TimedOutRounds + 1
                        activeTabs(t) = False
                        pendingCount = pendingCount - 1
                        Debug.Print "  [WARN] Tab " & t & " 読み込みタイムアウト。このラウンドはスキップします"
                    End If
                End If
            Next t

            If pendingCount > 0 Then chrome.sleep 0.05
        Loop

        ' --- Step C・D: Cookie/Screenshot発行・このラウンドのScreenshot回収（共通処理） ---
        FireAndDrainRound chrome, tabs, activeTabs, r, cookieTickets, cookieTicketCount, screenshotPayloads, screenshotPayloadCount
    Next r

    FinishBenchmark chrome, benchStart, tabStates, cookieTickets, cookieTicketCount, screenshotPayloads, screenshotPayloadCount, "Inline"
End Sub

'===================================================================================================
' Pattern B: exCDP_PageLoadWatcher 拡張クラスを利用するパターン
' ※事前に `VBAProject/Class/exCDP_PageLoadWatcher.cls` をインポートしておくこと
'===================================================================================================
Public Sub Test_AsyncBenchmark_RoundSync_ClassBased()
    Dim chrome As CDPBrowser
    Dim tabs() As CDPContext
    Dim watchers() As exCDP_PageLoadWatcher
    Dim tabStates() As TabState
    Dim urls(1 To 5) As String
    Dim t As Long, r As Long
    Dim cookieTickets() As CookieTicket, cookieTicketCount As Long
    Dim screenshotPayloads() As ScreenshotPayload, screenshotPayloadCount As Long
    Dim benchStart As Double

    urls(1) = URL_1: urls(2) = URL_2: urls(3) = URL_3: urls(4) = URL_4: urls(5) = URL_5

    PrintHeader "[ClassBasedパターン] マルチタブ非同期ベンチマーク 開始"
    Debug.Print "設定: タブ数=" & NUM_TABS & ", ラウンド数=" & NUM_ROUNDS

    ReDim tabs(1 To NUM_TABS)
    '---- Pipe版 ----
    Set chrome = ShSetting01_StartBrowser.StartCDPMode
    Set tabs(1) = chrome.getTab(setMain:=True)
    '----------------
    
    '---- WebSocket版 ----
'    '3. 設定セルから、ユーザ名を取得
'    Dim UserName As String
'    UserName = ShSetting01_StartBrowser.CurrentUserName
'
'    '4. 指定のWebSocketForCDPへ接続
'    Dim WebSocketCDP As New CDPCoreViaWebSocket
'    Debug.Print WebSocketCDP.AutoConnectBrowserCDP(UserName)
'
'    '5. 繋げたWebSocketオブジェクトを`reattach`メソッドに渡す
'    Set chrome = New CDPBrowser
'    If Not chrome.reattach(UserName, WebSocketCDP) Then MsgBox "「" & UserName & "」に接続できませんでした。WebSocket情報がお亡くなりです。", vbCritical, "Chrome DevTools Protocol": Exit Sub
'    Set tabs(1) = chrome.newTab(setMain:=True)
    '---------------------

    benchStart = chrome.TimerCounter

    ReDim watchers(1 To NUM_TABS)
    ReDim tabStates(1 To NUM_TABS)

    For t = 2 To NUM_TABS
        Set tabs(t) = chrome.newTab(newWindow:=False)
    Next t

    Randomize
    For t = 1 To NUM_TABS
        tabStates(t).Index = t

        Set watchers(t) = New exCDP_PageLoadWatcher
        watchers(t).Init tabs(t)   ' Page/Network有効化はクラス内部で完了する
    Next t

    ReDim cookieTickets(1 To NUM_TABS * NUM_ROUNDS)
    ReDim screenshotPayloads(1 To NUM_TABS * NUM_ROUNDS)

    For r = 1 To NUM_ROUNDS
        Debug.Print RESULT_SECTION_LINE
        Debug.Print "[Round " & r & "/" & NUM_ROUNDS & "] 開始"

        Dim activeTabs() As Boolean
        ReDim activeTabs(1 To NUM_TABS)

        ' --- Step A: 全タブ一斉に非同期遷移 ---
        For t = 1 To NUM_TABS
            tabStates(t).LoadedThisRound = False
            watchers(t).NavigateAsync urls(Int(Rnd * 5) + 1)   ' 内部でHasLoaded/NetworkEventCountをリセットしてから発行
        Next t

        ' --- Step B: 全タブ読み込み完了バリア ---
        Dim barrierStart As Double: barrierStart = chrome.TimerCounter
        Dim pendingCount As Long: pendingCount = NUM_TABS
        Do While pendingCount > 0
            chrome.TakeEvents   ' 1回のポンプで全タブのwatcherへイベントが配布される

            For t = 1 To NUM_TABS
                If Not tabStates(t).LoadedThisRound Then
                    If watchers(t).HasLoaded Then
                        tabStates(t).LoadedThisRound = True
                        tabStates(t).NetworkEventCountThisRound = watchers(t).NetworkEventCount
                        activeTabs(t) = True
                        pendingCount = pendingCount - 1
                        Debug.Print "  Tab " & t & " 読み込み完了 (NetworkEvents=" & tabStates(t).NetworkEventCountThisRound & ")"

                    ElseIf chrome.TimerCounter - barrierStart > TIMEOUT_LOAD_SEC * 1000 Then
                        tabStates(t).LoadedThisRound = True
                        tabStates(t).TimedOutRounds = tabStates(t).TimedOutRounds + 1
                        activeTabs(t) = False
                        pendingCount = pendingCount - 1
                        Debug.Print "  [WARN] Tab " & t & " 読み込みタイムアウト。このラウンドはスキップします"
                    End If
                End If
            Next t

            If pendingCount > 0 Then chrome.sleep 0.05
        Loop

        ' --- Step C・D: Cookie/Screenshot発行・このラウンドのScreenshot回収（共通処理） ---
        FireAndDrainRound chrome, tabs, activeTabs, r, cookieTickets, cookieTicketCount, screenshotPayloads, screenshotPayloadCount
    Next r

    FinishBenchmark chrome, benchStart, tabStates, cookieTickets, cookieTicketCount, screenshotPayloads, screenshotPayloadCount, "ClassBased"
End Sub

'===================================================================================================
' 共通ヘルパー（Pattern A/B どちらからも呼ばれる）
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
    Dim drainStart As Double: drainStart = chrome.TimerCounter
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

                ElseIf chrome.TimerCounter - drainStart > TIMEOUT_SCREENSHOT_SEC * 1000 Then
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

        If remaining > 0 Then chrome.sleep 0.05
    Loop
End Sub

'---------------------------------------------------------------------------------------------------
' 全ラウンド終了後: Cookie一括取得 → Screenshot一括保存 → サマリー出力 → ブラウザ終了
'---------------------------------------------------------------------------------------------------
Private Sub FinishBenchmark(chrome As CDPBrowser, benchStart As Double, ByRef tabStates() As TabState, _
                             ByRef cookieTickets() As CookieTicket, cookieTicketCount As Long, _
                             ByRef screenshotPayloads() As ScreenshotPayload, screenshotPayloadCount As Long, _
                             patternName As String)
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

    PrintHeader "[" & patternName & "パターン] ベンチマーク結果"
    Debug.Print "  タブ数               : " & NUM_TABS
    Debug.Print "  ラウンド数            : " & NUM_ROUNDS
    Debug.Print "  経過時間             : " & Format((chrome.TimerCounter - benchStart) / 1000, "0.0") & " 秒"
    For t = 1 To NUM_TABS
        Debug.Print "  Tab " & t & " タイムアウト回数    : " & tabStates(t).TimedOutRounds & " / " & NUM_ROUNDS & " ラウンド"
    Next t
    Debug.Print "  Cookie取得チケット数  : " & cookieTicketCount & " (Cookie総数: " & cookieSum & ")"
    Debug.Print "  Screenshot保存数     : " & savedCount & " / " & screenshotPayloadCount
    Debug.Print "  Screenshot保存先     : " & saveDir
    Debug.Print RESULT_SECTION_LINE

    chrome.quit
End Sub

'---------------------------------------------------------------------------------------------------
' 全ラウンド分のCookie整理券を、まとめて取得する（結果が来ていない間は待つ、タイムアウトで諦める）
'---------------------------------------------------------------------------------------------------
Private Sub DrainAllCookies(chrome As CDPBrowser, ByRef cookieTickets() As CookieTicket, cookieTicketCount As Long)
    Dim drainStart As Double: drainStart = chrome.TimerCounter
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

                ElseIf chrome.TimerCounter - drainStart > TIMEOUT_COOKIE_SEC * 1000 Then
                    cookieTickets(i).Retrieved = True
                    cookieTickets(i).HadError = True
                    remaining = remaining - 1
                    Debug.Print "  [WARN] Tab " & cookieTickets(i).TabIndex & " Round " & cookieTickets(i).RoundIndex & " Cookie取得タイムアウト"
                End If
            End If
        Next i

        If remaining > 0 Then chrome.sleep 0.1
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
