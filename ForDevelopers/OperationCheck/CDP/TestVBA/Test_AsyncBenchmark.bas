Attribute VB_Name = "Test_AsyncBenchmark"
'===================================================================================================
' CDP 複数タブ非同期遷移＆非同期スクショ一括保存ベンチマークテスト
'---------------------------------------------------------------------------------------------------
' 目的：
'   複数タブを用いて、非同期遷移（Page.navigate）と非同期スクリーンショット（Page.captureScreenshot）
'   を組み合わせた並列実行の動作を検証する。
'   また、取得した非同期コマンドの整理券（CommandID）をもとに、最終フェーズで一括保存を行う。
'===================================================================================================
Option Explicit

'---------------------------------------------------------------------------------------------------
' 定数定義
'---------------------------------------------------------------------------------------------------
Private Const RESULT_SECTION_LINE   As String = "=================================================="
Private Const NUM_TABS              As Long = 3      ' テストするタブ数
Private Const NUM_LAPS              As Long = 3      ' 各タブが繰り返す遷移・スクショのラップ数
Private Const SAVE_PATH             As String = "Downloads" ' ダウンロードフォルダ配下に保存 (Environ("UserProfile") & "\Downloads")

' ネットサーフィン対象URL（安定していて、かつネットワークイベントが発生しやすいサイト）
Private Const URL_1 As String = "https://www.google.com"
Private Const URL_2 As String = "https://www.yahoo.co.jp"
Private Const URL_3 As String = "https://example.com"
Private Const URL_4 As String = "https://news.yahoo.co.jp"
Private Const URL_5 As String = "https://www.amazon.co.jp"

' 各タブの状態を管理する構造体
Private Type TabState
    Index As Long           ' タブのインデックス (1 to NUM_TABS)
    Context As CDPContext   ' タブオブジェクト
    CurrentLap As Long      ' 現在のラップ数 (1 to NUM_LAPS)
    Status As String        ' "NAVIGATING", "COMPLETED"
    TargetUrl As String     ' 現在遷移中のURL
End Type

' 非同期スクショの整理券情報を記録する構造体
Private Type ScreenshotTicket
    TabIndex As Long        ' タブのインデックス
    Context As CDPContext   ' タブオブジェクト
    CommandID As Long       ' 非同期コマンドID (整理券番号)
    Lap As Long             ' ラップ数
    FileName As String      ' 保存ファイル名
    Saved As Boolean        ' 保存完了フラグ
End Type

'===================================================================================================
' メインテストプロシージャ（ここを実行する）
'===================================================================================================
Sub Test_AsyncBenchmark_Main()
    Dim chrome As New CDPBrowser
    Dim tabs() As CDPContext
    Dim tabStates() As TabState
    Dim tickets() As ScreenshotTicket
    Dim ticketCount As Long
    Dim urls(1 To 5) As String
    Dim i As Long, t As Long
    Dim allFinished As Boolean
    Dim saveDir As String
    
    saveDir = Environ("UserProfile") & "\" & SAVE_PATH
    ticketCount = 0
    
    ' URL配列の初期化
    urls(1) = URL_1
    urls(2) = URL_2
    urls(3) = URL_3
    urls(4) = URL_4
    urls(5) = URL_5

    Debug.Print RESULT_SECTION_LINE
    Debug.Print "[非同期ベンチマークテスト] 開始"
    Debug.Print "実行時刻: " & Format(Now, "yyyy/mm/dd hh:mm:ss")
    Debug.Print "設定: タブ数=" & NUM_TABS & ", ラップ数=" & NUM_LAPS
    Debug.Print RESULT_SECTION_LINE

    ' 1. ブラウザの起動とタブの用意
    Debug.Print "ブラウザを起動しています..."
    Set chrome = 設定シートからのCDP起動ForBrowser
    
    ReDim tabs(1 To NUM_TABS)
    ReDim tabStates(1 To NUM_TABS)
    
    ' 最初のタブを取得 (runTabsAsMany に準じる)
    Set tabs(1) = chrome.getTab(setMain:=True)
    
    ' 2番目以降のタブを作成 (newWindow:=False)
    For t = 2 To NUM_TABS
        Set tabs(t) = chrome.newTab(newWindow:=False)
    Next t

    ' 各タブのイベントの有効化と初期遷移
    Randomize
    For t = 1 To NUM_TABS
        Set tabStates(t).Context = tabs(t)
        tabStates(t).Index = t
        tabStates(t).CurrentLap = 1
        tabStates(t).Status = "NAVIGATING"
        
        ' イベント監視の有効化 (Page.loadEventFired をキャッチするため)
        tabs(t).ExecuteCDP "Page.enable"
        tabs(t).SetFilterEvents = "Page.loadEventFired"
        Set tabs(t).BrowserEvents = New Dictionary
        
        ' 初期遷移先をランダムに決定
        Dim rndIdx As Long
        rndIdx = Int(Rnd * 5) + 1
        tabStates(t).TargetUrl = urls(rndIdx)
        
        ' 非同期遷移の依頼
        Dim navParams As New Scripting.Dictionary
        navParams.Add "url", tabStates(t).TargetUrl
        
        Call tabs(t).ExecuteCDPAsync("Page.navigate", navParams)
        Debug.Print "Tab " & t & " Lap 1: 非同期遷移開始 -> " & tabStates(t).TargetUrl
    Next t

    ' 2. イベントループによる遷移・スクショ要求の並列制御
    Debug.Print RESULT_SECTION_LINE
    Debug.Print "非同期処理のイベントループを開始します..."
    
    Do
        ' ブラウザ全体のイベントポーリング（すべてのタブのイベントが処理される）
        chrome.TakeEvents
        
        allFinished = True
        For t = 1 To NUM_TABS
            If tabStates(t).Status = "NAVIGATING" Then
                allFinished = False
                
                ' Page.loadEventFired が発生したか確認
                If tabStates(t).Context.BrowserEvents("EventMethods").Exists("Page.loadEventFired") Then
                    Debug.Print "Tab " & t & " Lap " & tabStates(t).CurrentLap & " 読み込み完了！"
                    
                    ' (a) スクショ非同期依頼
                    Dim snapParams As New Scripting.Dictionary
                    ' 高速化のため getFullPage:=False 相当 (paramsは空でビューポートのみキャプチャ)
                    
                    Dim snapCmdID As Long
                    snapCmdID = tabStates(t).Context.ExecuteCDPAsync("Page.captureScreenshot", snapParams)
                    
                    ' 整理券を記録
                    ticketCount = ticketCount + 1
                    ReDim Preserve tickets(1 To ticketCount)
                    
                    tickets(ticketCount).TabIndex = t
                    Set tickets(ticketCount).Context = tabStates(t).Context
                    tickets(ticketCount).CommandID = snapCmdID
                    tickets(ticketCount).Lap = tabStates(t).CurrentLap
                    tickets(ticketCount).FileName = "bench_tab" & t & "_lap" & tabStates(t).CurrentLap & ".png"
                    tickets(ticketCount).Saved = False
                    
                    Debug.Print "  -> Tab " & t & " Lap " & tabStates(t).CurrentLap & " スクショ非同期依頼完了 (整理券番号: " & snapCmdID & ")"
                    
                    ' (b) 次の遷移を依頼するか、完了とするか
                    If tabStates(t).CurrentLap < NUM_LAPS Then
                        tabStates(t).CurrentLap = tabStates(t).CurrentLap + 1
                        tabStates(t).Status = "NAVIGATING"
                        
                        ' 次のランダムURLを選択
                        rndIdx = Int(Rnd * 5) + 1
                        tabStates(t).TargetUrl = urls(rndIdx)
                        
                        ' イベントバッファをクリアして、次のロードに備える
                        Set tabStates(t).Context.BrowserEvents = New Dictionary
                        
                        ' 非同期遷移の依頼
                        Dim nextNavParams As New Scripting.Dictionary
                        nextNavParams.Add "url", tabStates(t).TargetUrl
                        Call tabStates(t).Context.ExecuteCDPAsync("Page.navigate", nextNavParams)
                        Debug.Print "  -> Tab " & t & " Lap " & tabStates(t).CurrentLap & " 非同期遷移開始 -> " & tabStates(t).TargetUrl
                    Else
                        tabStates(t).Status = "COMPLETED"
                        Debug.Print "  -> Tab " & t & " 全ラップの依頼が完了しました。"
                    End If
                End If
            End If
        Next t
        
        ' CPU負荷削減のためのスリープ
        chrome.sleep 0.05
    Loop Until allFinished

    ' 3. 整理券を基に画像を一括保存するフェーズ
    Debug.Print RESULT_SECTION_LINE
    Debug.Print "全リクエストの送信が完了しました。画像の一括保存フェーズに移ります..."
    
    Dim DataConv As New WebCrypto
    Dim CharConv As New CharacterCodeConversion
    Dim allSaved As Boolean
    Dim savedCount As Long
    
    savedCount = 0
    
    Do
        chrome.TakeEvents
        allSaved = True
        
        For i = 1 To ticketCount
            If Not tickets(i).Saved Then
                allSaved = False
                
                ' ResultCDPFromWithEvents で結果が戻っているか確認
                Dim resJson As String
                resJson = tickets(i).Context.ResultCDPFromWithEvents(tickets(i).CommandID)
                
                If Len(resJson) > 0 Then
                    ' パース処理
                    Dim resDic As Dictionary
                    Set resDic = tickets(i).Context.InheritanceCDPBrowser.jsConverter.ParseJson(resJson)
                    
                    If Not resDic Is Nothing Then
                        If resDic.Exists("error") Then
                            Dim errMsg As String
                            errMsg = resDic("error")("message")
                            Debug.Print "  [保存失敗] Tab " & tickets(i).TabIndex & " Lap " & tickets(i).Lap & " : " & errMsg
                            tickets(i).Saved = True ' エラー終了として扱う
                            savedCount = savedCount + 1
                        ElseIf resDic.Exists("result") Then
                            Dim resultData As Dictionary
                            Set resultData = resDic("result")
                            
                            If resultData.Exists("data") Then
                                Dim b64 As String
                                b64 = resultData("data")
                                
                                ' Base64データをデコードしてファイル保存
                                Dim bytes() As Byte
                                bytes = DataConv.Decode(b64, edfBase64)
                                CharConv.BytesToSaveFile bytes, saveDir, tickets(i).FileName
                                
                                Debug.Print "  [保存成功] Tab " & tickets(i).TabIndex & " Lap " & tickets(i).Lap & " -> " & tickets(i).FileName
                                tickets(i).Saved = True
                                savedCount = savedCount + 1
                            End If
                        End If
                    End If
                End If
            End If
        Next i
        
        If Not allSaved Then chrome.sleep 0.1
    Loop Until allSaved

    Debug.Print RESULT_SECTION_LINE
    Debug.Print "[テスト終了] ベンチマーク結果"
    Debug.Print "  総リクエスト整理券数 : " & ticketCount
    Debug.Print "  保存処理完了数       : " & savedCount
    Debug.Print "  保存先フォルダ       : " & saveDir
    Debug.Print RESULT_SECTION_LINE

    ' ブラウザを閉じる
    chrome.quit
End Sub
