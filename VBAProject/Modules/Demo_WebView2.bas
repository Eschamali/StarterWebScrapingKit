Attribute VB_Name = "Demo_WebView2"
Option Explicit
'***
' Demo_WebView2.bas
' WebView2Core.cls の動作確認デモ
'
' 【実行前の準備】
' 1. VBEプロジェクトに WebView2Core.cls と WebView2Callbacks.bas をインポート
' 2. UserForm に Frame または PictureBox を配置し、そのhWndを使うか、
'    下記のように UserForm 自体の hWnd を親ウィンドウとして指定する
' 3. ExcelウィンドウのhWndを使う最小デモ → Sub TestWebView2Simple を実行
'***

'----------------------------------------------------------------------
' TestWebView2Simple
'   最もシンプルなデモ。Excelウィンドウに WebView2 を重ねて表示。
'   ※実際のアプリでは UserForm の Frame の hWnd を使う
'----------------------------------------------------------------------
Public Sub TestWebView2Simple()
    Dim wv2 As New WebView2Core

    ' Excel のメインウィンドウハンドルを取得
    Dim hWnd As LongPtr
    hWnd = Application.hWnd

    ' 画面左上 300x400px に WebView2 を表示
    Dim ok As Boolean
    ok = wv2.Initialize(hWnd, 0, 0, 800, 600, "https://eschamali.github.io/StarterWebScrapingKit/")

    If Not ok Then
        MsgBox "初期化コマンド送信失敗: " & wv2.LastErrorDescription, vbCritical
        Exit Sub
    End If

    ' ---- 初期化完了待機ループ ----
    ' DoEvents だけでは COM STA コールバックが層かない場合があるため、
    ' ProcessMessages（PeekMessage/DispatchMessage）でメッセージキューを層かせる
    Debug.Print "[WV2] Waiting for Ready... (check Immediate Window for callback logs)"
    Dim t As Single: t = Timer
    Do While Not wv2.IsReady And Timer - t < 15
        wv2.ProcessMessages
    Loop

    If wv2.IsReady Then
        MsgBox "WebView2 の初期化に成功しました！" & vbCrLf & "OKを押すと閉じます。", vbInformation
    Else
        MsgBox "タイムアウト。LastError: 0x" & Hex(wv2.LastErrorCode) & vbCrLf & wv2.LastErrorDescription, vbCritical
    End If

    wv2.Quit
    Set wv2 = Nothing
End Sub

'----------------------------------------------------------------------
' TestWebView2Form  ★推奨★
'   UserForm の Frame hWnd を親として WebView2 を埋め込むデモ。
'   vbModeless なので Excel の操作を維持しながら使える。
'
'   ★ Application.hWnd を親にするとクラッシュする問題の解決版 ★
'----------------------------------------------------------------------
Public Sub TestWebView2Form()
    WebView2Frame.Show vbModeless
End Sub

'----------------------------------------------------------------------
' TestWebView2FormModal
'   モーダル版（Excel 操作をブロックして WebView2 を表示）
'----------------------------------------------------------------------
Public Sub TestWebView2FormModal()
    WebView2Frame.Show
End Sub
