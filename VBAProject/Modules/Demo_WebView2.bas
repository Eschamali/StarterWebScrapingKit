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



'***************************************************************************************************
'                               ■■■ 設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：設定シートから、パラメーターを読み込んで、WebView2を起動するヘルパープロシージャです
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - WebView2Browser
'* 引数　　：StartURL                       ブラウザ起動時にアクセスしたいURL。指定しない場合は、空ページ(abount:blank)になります。
'            SwtchUser                      マルチインスタンス用に別ユーザーを指定するときに使用します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：VBEによるハードコーディングではなく、設定シートから読み込む方式により、ユーザー側からも手軽に設定変更ができます
'* 注意事項：Demoモジュールにあるコードですが、他の部分で共用してるため、消さずにどこかにカット&ペーストしておくとよいでしょう
'***************************************************************************************************
Public Function 設定シートからのWebView2起動(Optional StartURL As String, Optional SwitchUser As String) As WebView2Browser
    '設定シートの各セルから設定値を取得し、適用
    With ShSetting01_StartBrowser
        '第2引数が省略ならシート側の設定を適用
        Dim UseDataDir As String: UseDataDir = IIf(StrPtr(SwitchUser) = 0, .Range(.UseRangeName(2, "Demo_WebView2.設定シートからのWebView2起動")).value, SwitchUser)

        'ブラウザ起動
        Set 設定シートからのWebView2起動 = New WebView2Browser
        設定シートからのWebView2起動.start UseDataDir, .Range(.UseRangeName(12, "Demo_WebView2.設定シートからのWebView2起動")).value, StartURL, .Range(.UseRangeName(3, "Demo_WebView2.設定シートからのWebView2起動")).value
    End With
End Function





'----------------------------------------------------------------------
' TestWebView2Simple
'   最もシンプルなデモ。Excelウィンドウに WebView2 を重ねて表示。
'   ※実際のアプリでは UserForm の Frame の hWnd を使う
'----------------------------------------------------------------------
Public Sub TestWebView2Simple()
    Dim wv2 As New WebView2Core
Dim hwndParent
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

    wv2.quit
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
    WebView2InExcelForm.show vbModeless
End Sub

'----------------------------------------------------------------------
' TestWebView2FormModal
'   モーダル版（Excel 操作をブロックして WebView2 を表示）
'----------------------------------------------------------------------
Public Sub TestWebView2FormModal()
    WebView2InExcelForm.show
End Sub

'----------------------------------------------------------------------
' TestWebView2CDP ? CDP (CallDevToolsProtocolMethod) の動作確認
' Immediate に [CDP] Invoke が出ればコールバックは呼ばれている。出なければ OFF_WV2_CallDevToolsProtocolMethod の index を 22,36,37,39 などに変更して再試行
'----------------------------------------------------------------------
Public Sub TestWebView2CDP()
    Dim wv2 As New WebView2Core
    Dim hWnd As LongPtr: hWnd = Application.hWnd
    If Not wv2.Initialize(hWnd, 0, 0, 600, 400, "about:blank", "Automation Data") Then
        MsgBox "Initialize 失敗: " & wv2.LastErrorDescription, vbCritical
        Exit Sub
    End If
    Dim t As Single: t = Timer
    Do While Not wv2.IsReady And Timer - t < 15
        wv2.ProcessMessages
    Loop
    If Not wv2.IsReady Then
        MsgBox "Ready タイムアウト", vbCritical
        wv2.quit
        Exit Sub
    End If
    Dim params As String: params = "{""expression"":""1+1""}"
    Dim result As String: result = wv2.CallDevToolsProtocolMethod("Runtime.evaluate", params)
    Debug.Print "[CDP] Runtime.evaluate result: " & result
    MsgBox "CDP 結果（Immediate を確認）: " & Left$(result, 200) & IIf(Len(result) > 200, "...", ""), vbInformation
    wv2.quit
    Set wv2 = Nothing
End Sub


Sub WebView2Browserクラスで起動するDemo()
    Dim test As WebView2Browser: Set test = 設定シートからのWebView2起動("https://discord.com/app")
    Dim paramCDP As New Dictionary, ResultCDP As New Dictionary
    paramCDP.Add "expression", "1+1"
    Set ResultCDP = test.invokeMethod("Runtime.evaluate", paramCDP)
    Dim JsonConv As New WebJsonConverter

    MsgBox "CDP 結果: " & JsonConv.ConvertToJson(ResultCDP), vbInformation

    test.quit
    
End Sub


