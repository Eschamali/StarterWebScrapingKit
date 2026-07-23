Attribute VB_Name = "Test_strBuffer"
'===================================================================================================
' strBuffer 断片化検証テスト
'---------------------------------------------------------------------------------------------------
' 目的：
'   CDPContext 内の `strBuffer` が正しく機能しているかを検証する。
'   パイプから受信したJSONが途中で切れた（断片化した）際に、`strBuffer` に蓄積して
'   次回受信分と結合し、完全なJSONに復元できるかを確認する。
'
' 検証のアプローチ：
'   以下を組み合わせることで、意図的にパイプバッファの断片化を誘発する。
'     ① Page.captureScreenshot   → base64画像データで大量レスポンス（最有力）
'     ② Network.getAllCookies     → Cookieが多いサイトで大量レスポンス
'     ③ Networkイベントの同時発生 → ページ遷移中に複数イベントがまとめて到達
'
' 確認方法：
'   「ブラウザ起動設定」シートのログレベルを「1: Trace」に設定してから実行してください。
'   イミディエイト ウィンドウに以下が出れば、strBuffer が実際に使われた証拠です：
'     → "The Split(N) data is missing JSON data. Accumulate until completion."
'     → "Before-strBuffer = ..."
'     → "After-strBuffer = ..."
'     → "The merge completes the JSON data."
'
' 注意：
'   ・このテストは Test_strBuffer_Main() を実行してください
'   ・ログレベルは「ブラウザ起動設定」シートから「Trace(1)」に変更してください
'   ・実行後、イミディエイト ウィンドウのログを確認してください
'===================================================================================================
Option Explicit


'---------------------------------------------------------------------------------------------------
' 定数定義
'---------------------------------------------------------------------------------------------------
Private Const RESULT_SECTION_LINE   As String = "==============================="

' キャプチャ対象とするNetworkイベント名
Private Const EV_REQUEST_SENT      As String = "Network.requestWillBeSent"
Private Const EV_RESPONSE_RECEIVED As String = "Network.responseReceived"
Private Const EV_LOADING_FINISHED  As String = "Network.loadingFinished"

' ネットサーフィン対象URL（Networkイベントが多く発生しそうなサイトを選択）
Private Const URL_1 As String = "https://www.google.com"
Private Const URL_2 As String = "https://www.yahoo.co.jp"
Private Const URL_3 As String = "https://www.amazon.co.jp"
Private Const URL_4 As String = "https://news.yahoo.co.jp"
Private Const URL_5 As String = "https://twitter.com"


'===================================================================================================
' メインテストプロシージャ（ここを実行する）
'===================================================================================================
'***************************************************************************************************
'* 機能　　：strBuffer 断片化検証テストのエントリーポイント
'---------------------------------------------------------------------------------------------------
'* 詳細説明：スクリーンショット・getAllCookies・Networkイベントをランダムな順序で大量実行し、
'            パイプ受信データの断片化（strBuffer蓄積）が発生するかを確認するテストです。
'---------------------------------------------------------------------------------------------------
'* 実行前に：「ブラウザ起動設定」シートのログレベルを Trace(=1) に設定してください
'***************************************************************************************************
Sub Test_strBuffer_Main()

    ' ① ブラウザ起動（設定シート準拠）
    Dim br As CDPContext: Set br = ShSetting01_StartBrowser.StartCDPModeContext

    ' ② 統計カウンタの初期化
    Dim countSnap        As Long    ' snapPage 実行回数
    Dim countCookies     As Long    ' getAllCookies 実行回数
    Dim countNavigation  As Long    ' ナビゲーション回数

    Debug.Print RESULT_SECTION_LINE
    Debug.Print "[strBuffer 断片化検証テスト] 開始"
    Debug.Print "実行時刻: " & Format(Now, "yyyy/mm/dd hh:mm:ss")
    Debug.Print RESULT_SECTION_LINE

    ' ③ Networkイベントを有効化（パイプに大量イベントを流す）
    br.SetFilterEvents = EV_REQUEST_SENT
    br.SetFilterEvents = EV_RESPONSE_RECEIVED
    br.SetFilterEvents = EV_LOADING_FINISHED
    Set br.BrowserEvents = New Dictionary       'イベントキャプチャを有効化
    br.ExecuteCDP "Network.enable"


    '===================================================
    ' フェーズ1：ランダム順のCDPコマンド + ネットサーフィン
    '===================================================
    Debug.Print "[フェーズ1] ランダムな順序でCDPコマンドを実行しながらネットサーフィン"

    Dim urls(1 To 5) As String
    urls(1) = URL_1
    urls(2) = URL_2
    urls(3) = URL_3
    urls(4) = URL_4
    urls(5) = URL_5

    ' ランダムシードを設定
    Randomize

    Dim i As Long
    For i = 1 To 5
        ' ランダムに操作順を決める (1=先にスナップ, 2=先にCookie, 3=スナップ後Cookie)
        Dim actionOrder As Long: actionOrder = Int(Rnd * 3) + 1

        ' URL遷移（Networkイベントが大量発生）
        Dim targetUrl As String: targetUrl = urls(i)
        Debug.Print " [" & i & "/" & 5 & "] ナビゲーション → " & targetUrl
        br.navigate targetUrl
        countNavigation = countNavigation + 1

        ' TakeEventsでパイプをフラッシュ（イベントデータを受信）
        br.InheritanceCDPBrowser.TakeEvents

        ' actionOrder に応じてCDPコマンドをランダム実行
        Select Case actionOrder
            Case 1
                ' パターンA：スクショ → Cookie → スクショ
                Debug.Print "   Order A: スクショ → Cookie → スクショ"
                ExecSnapPage br, i, countSnap
                ExecGetAllCookies br, i, countCookies
                ExecSnapPage br, i, countSnap       '間髪入れずに2枚目（断片化しやすい）

            Case 2
                ' パターンB：Cookie → スクショ → Cookie
                Debug.Print "   Order B: Cookie → スクショ → Cookie"
                ExecGetAllCookies br, i, countCookies
                ExecSnapPage br, i, countSnap
                ExecGetAllCookies br, i, countCookies

            Case 3
                ' パターンC：スクショ(フルページ) → Cookie
                Debug.Print "   Order C: スクショ(フルページ) → Cookie"
                ExecSnapPageFull br, i, countSnap   'フルページのほうがデータ量が多い
                ExecGetAllCookies br, i, countCookies
        End Select

        ' 再度TakeEventsで追加イベント収集（イベントの連鎖を起こしやすくする）
        br.InheritanceCDPBrowser.TakeEvents
    Next i


    '===================================================
    ' フェーズ2：連打モード（間隔なしで大量実行）
    '===================================================
    Debug.Print RESULT_SECTION_LINE
    Debug.Print "[フェーズ2] 連打モード（間隔なしで大量CDPコマンド実行）"

    ' もう一度ネットサーフィン中に連続実行
    Dim j As Long
    For j = 1 To 3
        ' ランダムにURLを選んでナビゲート
        Dim rndIdx As Long: rndIdx = Int(Rnd * 5) + 1
        Debug.Print " [連打 " & j & "/3] → " & urls(rndIdx)
        br.navigate urls(rndIdx)
        countNavigation = countNavigation + 1

        ' TakeEventsは呼ばずに（バッファをあえて溜める）連続実行
        ExecSnapPage br, j, countSnap
        ExecGetAllCookies br, i, countCookies
        ExecSnapPageFull br, j, countSnap       '連続で全ページキャプチャ（最もデータが多い）
        ExecGetAllCookies br, j, countCookies   '再度Cookie（連打）

        br.InheritanceCDPBrowser.TakeEvents
    Next j


    '===================================================
    ' テスト結果サマリー出力
    '===================================================
    Debug.Print RESULT_SECTION_LINE
    Debug.Print "[テスト完了] 実行サマリー"
    Debug.Print "  ナビゲーション実行回数      : " & countNavigation
    Debug.Print "  snapPage 実行回数            : " & countSnap
    Debug.Print "  getAllCookies 実行回数       : " & countCookies
    Debug.Print ""
    Debug.Print "【確認ポイント】"
    Debug.Print "  ログ中に以下が出れば strBuffer が実際に機能した証拠です："
    Debug.Print "  → 'The Split(N) data is missing JSON data. Accumulate until completion.'"
    Debug.Print "  → 'Before-strBuffer = ...'"
    Debug.Print "  → 'The merge completes the JSON data.'"
    Debug.Print RESULT_SECTION_LINE

    ' ブラウザを閉じる
    br.InheritanceCDPBrowser.quit
End Sub


'===================================================================================================
' ヘルパープロシージャ群
'===================================================================================================

'***************************************************************************************************
'* 機能　　：スクリーンショット（通常ビュー）を実行し、カウントアップします
'---------------------------------------------------------------------------------------------------
'* 引数　　：br          CDPContextオブジェクト
'            Step_       ステップ番号（ログ表示用）
'            Count       カウンタ変数（参照渡しでインクリメント）
'***************************************************************************************************
Private Sub ExecSnapPage(br As CDPContext, Step_ As Long, ByRef Count As Long)
    Dim savePath As String: savePath = Environ("UserProfile") & "\Downloads"
    Dim fileName As String: fileName = "strBuffer_test_step" & Step_ & "_" & Count & ".png"

    Debug.Print "    [snapPage] 実行中... → " & fileName
    br.snapPage savePath, fileName, False   '通常ビュー（現在表示領域のみ）
    Count = Count + 1
End Sub


'***************************************************************************************************
'* 機能　　：スクリーンショット（フルページ）を実行し、カウントアップします
'---------------------------------------------------------------------------------------------------
'* 引数　　：br          CDPContextオブジェクト
'            Step_       ステップ番号（ログ表示用）
'            Count       カウンタ変数（参照渡しでインクリメント）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：getFullPage = True にするとページ全体をキャプチャするため、
'            base64データが膨大になり、断片化を誘発しやすい
'***************************************************************************************************
Private Sub ExecSnapPageFull(br As CDPContext, Step_ As Long, ByRef Count As Long)
    Dim savePath As String: savePath = Environ("UserProfile") & "\Downloads"
    Dim fileName As String: fileName = "strBuffer_test_full_step" & Step_ & "_" & Count & ".png"

    Debug.Print "    [snapPage Full] 実行中... → " & fileName
    br.snapPage savePath, fileName, True    'フルページ（ページ全体）← データ量が最大
    Count = Count + 1
End Sub


'***************************************************************************************************
'* 機能　　：Network.getAllCookies を実行し、カウントアップします
'---------------------------------------------------------------------------------------------------
'* 引数　　：br          CDPContextオブジェクト
'            Step_       ステップ番号（ログ表示用）
'            Count       カウンタ変数（参照渡しでインクリメント）
'***************************************************************************************************
Private Sub ExecGetAllCookies(br As CDPContext, Step_ As Long, ByRef Count As Long)
    Debug.Print "    [getAllCookies] 実行中..."
    Dim result As Dictionary: Set result = br.ExecuteCDP("Network.getAllCookies")

    If Not result Is Nothing Then
        If result.Exists("cookies") Then
            Debug.Print "    [getAllCookies] 取得件数: " & result("cookies").Count & " 件"
        End If
    End If

    Count = Count + 1
End Sub
