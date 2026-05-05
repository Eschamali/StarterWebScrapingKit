Attribute VB_Name = "Demo_AdvancedWait"
Option Explicit

' 動作確認用のデモマクロです。
' このモジュールと `exCDP_AdvancedWait` をインポートして実行してください。

'ワークスペースパス
'※StarterWebScrapingKitのルートフォルダ を入力してください
Private Const WORKSPACE_PATH As String = ""


Public Sub Demo_AdvancedWait()
    Dim br As CDPBrowser
    Dim extWait As exCDP_AdvancedWait
    Dim elem As CDPElement
    Dim startTime As Double
    Dim htmlPath As String
    
    htmlPath = "file:///" & WORKSPACE_PATH & "\Extensions\OperationCheck\TestHtml\Test_AdvancedWait\TestHtml.html"
    
    ' ブラウザの起動
    Set br = 設定シートからのCDP起動(htmlPath)
    
    ' 拡張機能（高度な待機）のインスタンス化と初期設定
    Set extWait = New exCDP_AdvancedWait
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
