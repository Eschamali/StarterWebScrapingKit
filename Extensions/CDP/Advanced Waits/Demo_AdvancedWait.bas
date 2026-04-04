Attribute VB_Name = "Demo_AdvancedWait"
Option Explicit

' 動作確認用のデモマクロです。
' このモジュールと `CDPexpansion_AdvancedWait` をインポートして実行してください。

'ワークスペースパス
'※StarterWebScrapingKitのルートフォルダ を入力してください
Private Const WORKSPACE_PATH As String = ""


Public Sub Demo_AdvancedWait()
    Dim br As CDPBrowser
    Dim extWait As CDPexpansion_AdvancedWait
    Dim elem As CDPElement
    Dim startTime As Double
    Dim htmlPath As String
    
    htmlPath = "file:///" & WORKSPACE_PATH & "\ForDevelopers\OperationCheck\CDP\TestHtml\Test_AdvancedWait\TestHtml.html"
    
    ' ブラウザの起動
    Set br = 設定シートからのCDP起動(htmlPath)
    
    ' 拡張機能（高度な待機）のインスタンス化と初期設定
    Set extWait = New CDPexpansion_AdvancedWait
    extWait.Init br
    
    br.printMsg info_, "Test環境にアクセスし、Advanced Waitの初期化が完了しました。デモを開始します。", "Demo"
    br.sleep 2
    
    ' ========================================================
    ' 1. DOM Mutation 待機のテスト
    ' ========================================================
    br.printMsg info_, WorksheetFunction.Unichar(9654) & " [1] DOM Mutation テストを開始します。", "Demo"
    
    Set elem = br.getElementByQuery("#btn-dom-mutation")
    If elem.isExist Then
        br.printMsg info_, "  - DOM追加トリガーボタンをクリックします。(2秒後にDOM変化発生)", "Demo"
        elem.click
        
        br.printMsg info_, "  - WaitForDomMutation（タイムアウト5秒）を開始します...", "Demo"
        startTime = Timer
        
        ' 拡張機能の機能呼び出し：DOM要素が変動するまで（最大5秒）待機
        extWait.WaitForDomMutation 5
        
        br.printMsg info_, "  - " & WorksheetFunction.Unichar(10004) & " 検知成功！ " & Format(Timer - startTime, "0.00") & "秒で変化を捉えました。", "Demo"
    End If
    
    br.sleep 2
    
    ' ========================================================
    ' 2. Network Idle 待機のテスト
    ' ========================================================
    br.printMsg info_, WorksheetFunction.Unichar(9654) & " [2] Network Idle テストを開始します。", "Demo"
    
    Set elem = br.getElementByQuery("#btn-network-idle")
    If elem.isExist Then
        br.printMsg info_, "  - 非同期通信トリガーボタンをクリックします。(時間差で通信発生)", "Demo"
        elem.click
        
        br.printMsg info_, "  - WaitForNetworkIdle（タイムアウト10秒）を開始します...", "Demo"
        startTime = Timer
        
        ' 拡張機能の機能呼び出し：非同期通信が完全に落ち着くまで待機
        ' 引数2: 500ms(デフォルト)の無通信時間があれば完了とみなす
        extWait.WaitForNetworkIdle 10
        
        br.printMsg info_, "  - " & WorksheetFunction.Unichar(10004) & " 検知成功！ " & Format(Timer - startTime, "0.00") & "秒で通信の終了(Idle)を確認しました。", "Demo"
    End If
    
    br.sleep 3
    br.printMsg info_, "デモがすべて完了しました。5秒後にブラウザを閉じます。", "Demo"
    br.sleep 5
    br.quit
    
End Sub

' =========================================================
' 追加実装された高度な待機機構のデモ
' =========================================================
Public Sub Demo_WaitOptions()
    Dim br As CDPBrowser
    Dim extWait As CDPexpansion_AdvancedWait
    Dim elem As CDPElement
    Dim startTime As Double
    Dim htmlPath As String
    
    ' ログインデモサイトのログイン画面を開く
    Dim loginUrl As String: loginUrl = "https://hotel-example-site.takeyaqa.dev/ja/login.html"
    Set br = 設定シートからのCDP起動(loginUrl)
    Set extWait = New CDPexpansion_AdvancedWait
    extWait.Init br
    
    br.printMsg info_, "[Demo_WaitOptions] アドバンスド待機機能のデモを開始します。", "Demo"
    br.printMsg info_, "  - テストサイト: " & loginUrl, "Demo"
    br.sleep 1
    
    ' --- 1. ClickAndWaitForIdle のテスト ---
    br.printMsg info_, WorksheetFunction.Unichar(9654) & " [A] ClickAndWaitForIdle テストを開始します。", "Demo"
    Set elem = br.getElementByQuery("#btn-network-idle")
    If elem.isExist Then
        br.printMsg info_, "  - 要素をクリックし、発生した通信の波紋が完全に消えるまで待ちます...", "Demo"
        startTime = Timer
        
        ' クリックとFullIdle待機を1アクションで実行 (タイムアウト10秒)
        Dim isSuccessIdle As Boolean
        isSuccessIdle = extWait.ClickAndWaitForIdle(elem, 500, 10000)
        
        If isSuccessIdle Then
            br.printMsg info_, "  - " & WorksheetFunction.Unichar(10004) & " 波及イベントの終息を確認！ " & Format(Timer - startTime, "0.00") & "秒", "Demo"
        Else
            br.printMsg WARN_, "  - " & WorksheetFunction.Unichar(10008) & " タイムアウトしました。", "Demo"
            MsgBox "ClickAndWaitForIdle がタイムアウトしました。", vbExclamation, "Demo Error"
        End If
    End If
    br.sleep 2
    
    ' --- 2. WaitForUrlRedirect のテスト（CDPNetworkイベント版）---
    ' BiDiの ExecuteIsUrlContains に相当する正攻法の実装です。
    ' Main04に沿ったログイン遷移シナリオで実証します。
    br.printMsg info_, WorksheetFunction.Unichar(9654) & " [B] WaitForUrlRedirect (ログイン遷移检知)テストを開始します。", "Demo"
    br.printMsg info_, "  - ログイン情報: ichiro@example.com / password", "Demo"
    br.printMsg info_, "  - メールアドレスとパスワードを入力して「ログイン」ボタンを押してください。", "Demo"
    br.printMsg info_, "  - mypage.html への遷移をCDPNetworkイベント(ネイティブ)で監視開始...", "Demo"
    
    startTime = Timer
    Dim isSuccessUrl As Boolean
    isSuccessUrl = extWait.WaitForUrlRedirect("mypage.html", 30)
    
    If isSuccessUrl Then
        br.printMsg info_, "  - " & WorksheetFunction.Unichar(10004) & " ログイン遷移を検知！ " & Format(Timer - startTime, "0.00") & "秒 / URL: " & br.url, "Demo"
        MsgBox "デモ成功：全ての高度な待機が想定通り機能しました。", vbInformation, "Demo Success"
    Else
        br.printMsg WARN_, "  - " & WorksheetFunction.Unichar(10008) & " ログイン遷移を検知できませんでした。", "Demo"
        MsgBox "WaitForUrlRedirect がタイムアウトしました。", vbExclamation, "Demo Error"
    End If
    
    br.sleep 3
    br.quit
End Sub
