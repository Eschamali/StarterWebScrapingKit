Attribute VB_Name = "Demo_OfficeScript"
'***************************************************************************************************
'       exCDP_OfficeScript 拡張 - デモ & 動作確認 モジュール
'***************************************************************************************************
'* 機能　　：`exCDP_OfficeScript.cls` を使った、VBA → CDP → WebView2 → Officeスクリプト実行のサンプルです
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\OfficeScriptViaVBA\exCDP_OfficeScript.cls
'---------------------------------------------------------------------------------------------------
'* 注意事項：・実行前に「exCDP_OfficeScript.cls」をVBAプロジェクトに取り込んでください
'            ・実行前に、Excelの[自動化]タブからスクリプトを開き「コード エディター」を表示してください
'            ・エディターで開いているスクリプトの中身は、デモのスクリプトで上書きされます
'            ・デモのスクリプトは、アクティブシートのA1/A2を書き換えます。空のシートで試してください
'            ・`CDPBrowser`の用意と切断はbas側(このモジュール)で行い、`exCDP_OfficeScript.Init`に渡します。
'              切断(`DisconnectCDP`)の前に、必ず`exCDP_OfficeScript.Dispose`(または破棄)してください
'***************************************************************************************************
Option Explicit

'Demo 02(OnTimeで結果回収)用：マクロ終了後も接続と整理券番号を保持する
Private m_Browser       As CDPBrowser
Private m_OfficeScript  As exCDP_OfficeScript
Private m_Ticket        As Long
Private m_AsyncStart    As Double



'***************************************************************************************************
'                            ■■■ 初めに ■■■
'***************************************************************************************************
'* 機能　　：デバッグ経路を確保します
'***************************************************************************************************
Sub FirstFunction()
    Dim tmp As New exCDP_OfficeScript
    tmp.OpenDebugPortWebView2
End Sub



'***************************************************************************************************
'                         ■■■ CDPBrowser の用意(bas側の責務) ■■■
'***************************************************************************************************
'* 機能　　：デスクトップ版Excel：OfficeアドインのWebView2へ接続した`CDPBrowser`を返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：接続済みの`CDPBrowser`。接続できなければ Nothing
'---------------------------------------------------------------------------------------------------
'* 注意事項：WebView2のデバッグポートは、環境変数`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`で開けるため、
'            Office側のWebView2が起動する「前」に`EnsureWebView2DebugPort`が呼ばれている必要があります。
'            接続できない場合は、一度実行した後にコードエディターを開き直して再実行してください
'***************************************************************************************************
Public Function ConnectDesktopExcel() As CDPBrowser
    Const USERDATA_SUBPATH As String = "\Microsoft\Office\16.0\Wef\webview2\2\1\EBWebView"   '%LOCALAPPDATA%配下


    On Error GoTo Failed

    '1. OfficeアドインのWebView2のユーザーデータフォルダ
    Dim UserDataDir As String
    UserDataDir = Environ("LOCALAPPDATA") & USERDATA_SUBPATH

    '2. WebView2のデバッグポートを開け、`DevToolsActivePort`から接続
    BrowserCapabilities.Set_UserDataDirName = UserDataDir
    Dim WebSocketCDP As CDPCoreViaWebSocket: Set WebSocketCDP = New CDPCoreViaWebSocket
    If Not WebSocketCDP.AutoConnectDevToolsActivePort(BrowserCapabilities) Then
        Debug.Print "[ConnectDesktopExcel] `DevToolsActivePort`から接続できませんでした。コードエディターを開き直して再実行してください"
        Exit Function
    End If

    '3. WebSocketモードとして設定
    ReattachOptions.Set_WebSocketMode(WebSocketCDP) = UserDataDir

    '4. 再接続情報を`reattach`メソッドに渡す
    Set ConnectDesktopExcel = New CDPBrowser
    ConnectDesktopExcel.reattach ReattachOptions
    Exit Function

Failed:
    Debug.Print "[ConnectDesktopExcel] 接続に失敗しました。Error " & Err.Number & ": " & Err.Description
End Function

'***************************************************************************************************
'* 機能　　：Web版Excel：Excelを開いているEdgeへ接続した`CDPBrowser`を返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：接続済みの`CDPBrowser`。接続できなければ Nothing
'---------------------------------------------------------------------------------------------------
'* 注意事項：・事前にEdgeのリモートデバッグを有効にしてください(`edge://inspect/#remote-debugging`で許可、等)
'            ・`DevToolsActivePort`に書かれたブラウザ用WebSocketで繋ぐため、`/json`系が無効でも繋がります
'            ・別のユーザーデータフォルダで起動したEdgeの場合は、UserDataDirを書き換えてください
'***************************************************************************************************
Public Function ConnectWebExcel() As CDPBrowser
    Const MainUserData As String = "User Data"


    On Error GoTo Failed

    '1. ワークシートの設定に委ねた、今目の前のブラウザへ接続させる
    Dim WebSocketCDP As CDPCoreViaWebSocket: Set WebSocketCDP = New CDPCoreViaWebSocket
    BrowserCapabilities.Set_UserDataDirName = MainUserData
    If Not WebSocketCDP.AutoConnectDevToolsActivePort(BrowserCapabilities) Then
        Debug.Print "[ConnectWebExcel] `DevToolsActivePort`から接続できませんでした。Edgeのリモートデバッグを有効にしてください"
        Exit Function
    End If

    '2. WebSocketモードとして設定
    ReattachOptions.Set_WebSocketMode(WebSocketCDP) = MainUserData

    '3. 再接続情報を`reattach`メソッドに渡す
    Set ConnectWebExcel = New CDPBrowser
    ConnectWebExcel.reattach ReattachOptions
    Exit Function

Failed:
    Debug.Print "[ConnectWebExcel] 接続に失敗しました。Error " & Err.Number & ": " & Err.Description
End Function



'***************************************************************************************************
'                            ■■■ Demo 01：同期実行 ■■■
'***************************************************************************************************
'* 機能　　：スクリプトを書き込んで実行し、完了までVBA側で待ちます
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - 戻り値が "Succeeded" であること
'   - アクティブシートのA1/A2に値が入ること
'   - `LastOutputs`に console.log の内容("sheet: ...")が入ること
'***************************************************************************************************
Sub Demo_OfficeScript_01_同期実行()

    '1. bas側でCDPBrowserを用意
    Dim Browser As CDPBrowser: Set Browser = ConnectDesktopExcel()
    If Browser Is Nothing Then MsgBox "接続できませんでした。イミディエイトウィンドウを確認してください。", vbCritical: Exit Sub

    '2. 継承して、コードエディターを確保
    Dim os As New exCDP_OfficeScript
    If os.Init(Browser) Then

        '3. 同期実行
        Dim Status As String
        Status = os.RunScript(SampleScript_HelloWorld())

        '4. 結果確認
        Debug.Print "[Demo01] status = " & Status
        Debug.Print os.LastLogs

        Dim Line As Variant
        For Each Line In os.LastOutputs
            Debug.Print "[Demo01] console: " & Line
        Next
    Else
        MsgBox "コード エディターが見つかりませんでした。イミディエイトウィンドウを確認してください。", vbCritical
    End If

    '5. 後片付け(先にClassを手放してから切断)
    os.Dispose
    Browser.DisconnectCDP

End Sub



'***************************************************************************************************
'                  ■■■ Demo 02：非同期実行(マクロ終了後にOnTimeで結果回収) ■■■
'***************************************************************************************************
'* 機能　　：実行を開始したらマクロをいったん終了し、`Application.OnTime`で結果を取りに行きます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Excelを操作可能な状態に戻したまま、裏でOfficeスクリプトを走らせたい場合の例です
'* 確認ポイント：
'   - 実行開始直後にマクロが終わり、少し後にイミディエイトへ結果が出ること
'***************************************************************************************************
Sub Demo_OfficeScript_02_非同期実行()

    '1. 接続(結果回収まで保持するため、モジュールレベル変数に入れる)
    Set m_Browser = ConnectDesktopExcel()
    If m_Browser Is Nothing Then MsgBox "接続できませんでした。イミディエイトウィンドウを確認してください。", vbCritical: Exit Sub

    Set m_OfficeScript = New exCDP_OfficeScript
    If Not m_OfficeScript.Init(m_Browser) Then
        Demo_OfficeScript_02_Cleanup
        MsgBox "コード エディターが見つかりませんでした。イミディエイトウィンドウを確認してください。", vbCritical
        Exit Sub
    End If

    '2. 実行開始 → 整理券番号を受け取る
    m_Ticket = m_OfficeScript.RunScriptAsync(SampleScript_HelloWorld())
    If m_Ticket = 0 Then
        Debug.Print "[Demo02] 開始できませんでした: " & m_OfficeScript.LastLogs
        Demo_OfficeScript_02_Cleanup
        Exit Sub
    End If
    Debug.Print "[Demo02] 実行開始。整理券番号: " & m_Ticket

    '3. 結果回収を予約して、マクロは終了
    m_AsyncStart = Timer
    Application.OnTime Now + TimeSerial(0, 0, 1), "Demo_OfficeScript.Demo_OfficeScript_02_Poll"

End Sub

'***************************************************************************************************
'* 機能　　：Demo 02 の結果回収(`Application.OnTime`から呼ばれる)
'***************************************************************************************************
Sub Demo_OfficeScript_02_Poll()
    Const TimeoutSec As Double = 120


    If m_OfficeScript Is Nothing Then Exit Sub

    '1. 整理券番号で、1回だけ確認
    Dim Status As String
    If Not m_OfficeScript.TryGetResult(m_Ticket, Status) Then
        '未完了：タイムアウトしていなければ再予約
        If Timer - m_AsyncStart < TimeoutSec Then
            Application.OnTime Now + TimeSerial(0, 0, 1), "Demo_OfficeScript.Demo_OfficeScript_02_Poll"
            Exit Sub
        End If
        Status = "Timeout"
    End If

    '2. 結果表示
    Debug.Print "[Demo02] ticket " & m_Ticket & " status = " & Status
    Debug.Print m_OfficeScript.LastLogs

    '3. 後片付け
    Demo_OfficeScript_02_Cleanup
End Sub

'***************************************************************************************************
'* 機能　　：Demo 02 の後片付け(先にClassを手放してから切断)
'***************************************************************************************************
Private Sub Demo_OfficeScript_02_Cleanup()
    Set m_OfficeScript = Nothing
    If Not m_Browser Is Nothing Then m_Browser.DisconnectCDP
    Set m_Browser = Nothing
End Sub



'***************************************************************************************************
'               ■■■ Demo 03：非同期実行(実行中にVBA側で別作業 → 整理券で待つ) ■■■
'***************************************************************************************************
'* 機能　　：実行開始後にVBA側で別の処理をして、あとから整理券番号で結果を待ちます
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'   - 実行中は`IsRunning`がTrueになること
'   - 実行中にもう1本`RunScriptAsync`しても、0(開始不可)が返ること
'   - `WaitResult`で、最初の整理券番号の結果が取れること
'***************************************************************************************************
Sub Demo_OfficeScript_03_整理券で待つ()

    Dim Browser As CDPBrowser: Set Browser = ConnectDesktopExcel()
    If Browser Is Nothing Then Exit Sub

    Dim os As New exCDP_OfficeScript
    If os.Init(Browser) Then

        '1. 実行開始
        Dim Ticket As Long
        Ticket = os.RunScriptAsync(SampleScript_HelloWorld())
        If Ticket = 0 Then
            Debug.Print "[Demo03] 開始できませんでした: " & os.LastLogs
        Else
            Debug.Print "[Demo03] 整理券番号: " & Ticket

            '2. 実行中にVBA側で別作業(ここでは状態確認と、2本目の投入を試す)
            Debug.Print "[Demo03] IsRunning = " & os.IsRunning
            Debug.Print "[Demo03] 2本目の整理券番号(0なら開始不可) = " & os.RunScriptAsync(SampleScript_HelloWorld())

            '3. 1本目の結果を待つ
            Debug.Print "[Demo03] status = " & os.WaitResult(Ticket, 60)
            Debug.Print os.LastLogs
        End If
    End If

    '4. 後片付け
    os.Dispose
    Browser.DisconnectCDP

End Sub



'***************************************************************************************************
'                         ■■■ Demo 04：エラーになるスクリプト ■■■
'***************************************************************************************************
'* 機能　　：スクリプト内で例外が発生した場合に、Failed とエラーメッセージが取れるかを確認します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：エディター側の出力ログには「予期しないエラーが発生しました。」という汎用メッセージしか来ないため、
'            例外の詳細は「Office JS Sandbox」タブ側から回収した`LastExceptionDetails`で確認します
'* 既知の制約：Sandboxタブは、Excel起動後の最初の実行時に作られます。
'              Excel起動直後にこのデモを最初に動かすと、例外の詳細はNothingになります(Demo01等を先に1回実行してください)
'* 確認ポイント：
'   - 戻り値が "Failed" であること
'   - `LastExceptionMessage`が "intentional error from VBA" であること
'   - `LastExceptionDetails`の"stack"に、main関数のスタックが入ること
'   - `LastLogs`に[exception]行が追記されること
'***************************************************************************************************
Sub Demo_OfficeScript_04_エラー確認()

    Dim Browser As CDPBrowser: Set Browser = ConnectDesktopExcel()
    If Browser Is Nothing Then Exit Sub

    Dim os As New exCDP_OfficeScript
    If os.Init(Browser) Then

        '1. 実行
        Debug.Print "[Demo04] status = " & os.RunScript(SampleScript_Error())
        Debug.Print os.LastLogs

        '2. 例外の詳細
        If os.LastExceptionDetails Is Nothing Then
            Debug.Print "[Demo04] 例外の詳細は取得できませんでした(Sandboxタブ未接続、または例外なし)"
        Else
            Debug.Print "[Demo04] name    = " & os.LastExceptionDetails.StringKey("name")
            Debug.Print "[Demo04] message = " & os.LastExceptionMessage
            Debug.Print "[Demo04] stack   = " & os.LastExceptionDetails.StringKey("stack")
            Debug.Print "[Demo04] json    = " & os.LastExceptionDetails.Stringify(True)
        End If
    End If

    '3. 後片付け
    os.Dispose
    Browser.DisconnectCDP

End Sub



'***************************************************************************************************
'                  ■■■ Demo 05：Web版(Edgeで開いたExcel)で実行 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザで開いたExcel(Web版)のコードエディターを確保し、同期実行 → 例外の確認まで行います
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Web版のシートはCanvas描画ですが、Officeスクリプト経由ならセルの値を直接読み書きできます
'* 注意事項：・事前にEdgeのリモートデバッグを有効にしてください(`edge://inspect/#remote-debugging`で許可、等)
'            ・事前に、Web版Excelでスクリプトを開き「コード エディター」を表示してください
'            ・エディターで開いているスクリプトの中身は、デモのスクリプトで上書きされます
'            ・`DisconnectCDP`はCDPの接続を閉じるだけで、Edge自体は閉じません
'* 確認ポイント：
'   - 1本目が "Succeeded" で、アクティブシートのA1/A2に値が入ること
'   - 2本目が "Failed" で、`LastExceptionMessage`が "intentional error from VBA" であること
'     (Web版は、Excelを開いた直後の最初の実行でも取れる。"name"/"stack"は空文字)
'***************************************************************************************************
Sub Demo_OfficeScript_05_Web版()

    '1. bas側でCDPBrowserを用意(Excelを開いているEdgeへ接続)
    Dim Browser As CDPBrowser: Set Browser = ConnectWebExcel()
    If Browser Is Nothing Then MsgBox "Edgeに接続できませんでした。イミディエイトウィンドウを確認してください。", vbCritical: Exit Sub

    '2. 継承して、Web版のコードエディター(iframe)を確保
    Dim os As New exCDP_OfficeScript
    If os.Init(Browser, WebMode:=True) Then
        Debug.Print "[Demo05] IsWeb = " & os.IsWeb

        '3. 同期実行
        Debug.Print "[Demo05] status = " & os.RunScript(SampleScript_HelloWorld())
        Debug.Print os.LastLogs

        '4. 例外になるスクリプト
        Debug.Print "[Demo05] status = " & os.RunScript(SampleScript_Error())
        Debug.Print os.LastLogs
        If os.LastExceptionDetails Is Nothing Then
            Debug.Print "[Demo05] 例外の詳細は取得できませんでした"
        Else
            Debug.Print "[Demo05] message = " & os.LastExceptionMessage
            Debug.Print "[Demo05] json    = " & os.LastExceptionDetails.Stringify(True)
        End If
    Else
        MsgBox "Web版のコード エディターが見つかりませんでした。イミディエイトウィンドウを確認してください。", vbCritical
    End If

    '5. 後片付け(先にClassを手放してから切断。Edgeは閉じない)
    os.Dispose
    Browser.DisconnectCDP

End Sub



'***************************************************************************************************
'                             ■■■ デモ用スクリプト ■■■
'***************************************************************************************************
'* 機能　　：アクティブシートのA1/A2に書き込み、シート名をconsole.logに出します
'***************************************************************************************************
Private Function SampleScript_HelloWorld() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    const sheet = workbook.getActiveWorksheet();" & vbLf
    s = s & "    sheet.getRange(""A1"").setValue(""Hello from VBA -> CDP -> OfficeScript"");" & vbLf
    s = s & "    sheet.getRange(""A2"").setValue(new Date().toLocaleString());" & vbLf
    s = s & "    console.log(""sheet: "" + sheet.getName());" & vbLf
    s = s & "}" & vbLf
    SampleScript_HelloWorld = s
End Function

'***************************************************************************************************
'* 機能　　：わざと例外を投げます
'***************************************************************************************************
Private Function SampleScript_Error() As String
    Dim s As String
    s = s & "function main(workbook: ExcelScript.Workbook) {" & vbLf
    s = s & "    console.log(""before error"");" & vbLf
    s = s & "    throw new Error(""intentional error from VBA"");" & vbLf
    s = s & "}" & vbLf
    SampleScript_Error = s
End Function
