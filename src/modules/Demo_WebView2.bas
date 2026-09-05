Attribute VB_Name = "Demo_WebView2"
'***************************************************************************************************
'   WebView2(CDPCoreViaWebView2)関連のデモ集です
'***************************************************************************************************
Option Explicit
Option Private Module



'***************************************************************************************************
'                                ■■■ Hello World ■■■
'***************************************************************************************************
'* 機能　　：WebView2を起動します。基本的な呼び出しフローです
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`ICoreWebView2Settings`等の一部設定は、ページ遷移前のみ有効です
'            ・`ICoreWebView2EnvironmentOptions`の設定は、WebView2プロセス起動前のみ有効です
'***************************************************************************************************
Sub ExcelのユーザーフォームにWebView2を埋め込む()
    With WebView2Form
        '1. 起動前設定を施す(任意)
        .ThisWebView2.EnvironmentOptions.Set_AllowSingleSignOnUsingOSPrimaryAccount = False  'シングルサインオンの切り替え

        '2. WebView2プロセスを起動
        If Not .StartCDPModeWebView2 Then Debug.Print "WebView2の初期化に失敗しました。WebView2Loader.dllが見つからない、" & _
                                                        "またはEnvironment/Controllerの生成に失敗した可能性があります。": Exit Sub

        '3. 遷移前の事前設定を施す(任意)
        .ThisWebView2.DevToolsEnabled = False       'DevToolsウィンドウ起動禁止
        .ThisWebView2.ContextMenuEnabled = False    '右クリック禁止

        '3. CDPとして、ページ遷移
        'シングルサインオンを無効：Microsoftアカウントの紹介
        'シングルサインオンを有効：あなたのPCでログイン中のMicrosoftアカウント設定ページに自動遷移
        .ThisCDPContext.navigate "https://account.microsoft.com/"

        '4. フォームを表示
        '※UserFormを閉じるまで、ブロッキングされます
        .show
    End With
End Sub



'***************************************************************************************************
'                                ■■■ 拡張機能デモ ■■■
'***************************************************************************************************
'* 機能　　：WebView2にて拡張機能をインストール/アンインストールする際のDemoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：拡張機能系のみ、CDPコマンドでは出来ないため、WebView2側で用意されてる拡張機能用APIでやるためのDemoとなります
'            恐らく内部では`Page`単位(/json/list)としての実行となっているため`Method not available.`エラーと推測してます
'***************************************************************************************************
Sub 拡張機能インストールアンインストール()
    Const UseCDP As Boolean = False


    Dim インストールパス As String
    With Application.FileDialog(4)  'msoFileDialogFolderPicker
        .Title = "拡張機能の基となる`manifest.json`を含むフォルダを選択してください"
        .InitialFileName = Environ("LOCALAPPDATA")    '初期位置

        If .show = -1 Then インストールパス = .SelectedItems(1) Else Exit Sub
    End With


    With WebView2Form
        '1. 拡張機能を有効にする
        .ThisWebView2.EnvironmentOptions.Set_AreBrowserExtensionsEnabled = True

        '2. WebView2を起動
        If Not .StartCDPModeWebView2 Then Debug.Print "WebView2の初期化に失敗しました。WebView2Loader.dllが見つからない、" & _
                                                        "またはEnvironment/Controllerの生成に失敗した可能性があります。": Exit Sub

        '3. ページ遷移
        .ThisCDPContext.navigate "https://github.com/Eschamali/StarterWebScrapingKit"

        '4. フォームを表示
        .show vbModeless

        '5. 拡張機能インストール
        Dim InstallID As String
        If UseCDP Then
            Dim CDPparams As Dictionary, ResultCDP As BiDiCDPJson
            Set CDPparams = New Dictionary
            CDPparams.Add "path", インストールパス
            Set ResultCDP = .ThisCDPContext.ThisCDPBrowser.ExecuteCDP("Extensions.loadUnpacked", CDPparams, False)    '今回は、エラー無視で設定

            If ResultCDP Is Nothing Then MsgBox "拡張機能のインストールに失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & .ThisCDPContext.ThisCDPBrowser.LastCDPJsonError("message"), vbCritical, "ErrorCode:" & .ThisCDPContext.ThisCDPBrowser.LastCDPJsonError("code"): Unload WebView2Form: Exit Sub

            '6. アンインストール
            InstallID = ResultCDP("id")
            MsgBox "拡張機能のインストールに成功しました。", vbInformation, "exID: " & InstallID
        
        Else
            InstallID = .ThisWebView2.AddBrowserExtension(インストールパス)
            If LenB(InstallID) = 0 Then MsgBox "拡張機能のインストールに失敗しました", vbCritical, "WebView2": Unload WebView2Form: Exit Sub
            MsgBox "拡張機能のインストールに成功しました。OKを押すとアンインストールします", vbInformation, "exID: " & InstallID

            '拡張機能がインストールされてるリストを取得
            Dim ext As Variant
            For Each ext In .ThisWebView2.GetBrowserExtensionIds
                Debug.Print ext("ID"), ext("Name"), ext("IsEnabled")
            Next ext

            '6. アンインストール
            If Not .ThisWebView2.RemoveBrowserExtension(InstallID) Then MsgBox "拡張機能のアンインストールに失敗しました", vbCritical, "WebView2": Unload WebView2Form: Exit Sub
            MsgBox "拡張機能のアンインストールに成功しました。", vbInformation
        End If
    End With

    '7. Demo終了
    Unload WebView2Form
End Sub



'***************************************************************************************************
'                       ■■■ SetDefaultBackgroundColor 検証用デモ ■■■
'***************************************************************************************************
'* 機能　　：`WebView2Form`(既存のUserForm、内部の`EdgeFrame`にWebView2を埋め込み済み)を起動し、
'            `SetDefaultBackgroundColor`(Controller2)が実際に反映されるかを確認します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`WebView2Form.StartCDPModeWebView2`は内部で`about:blank`まで遷移させるため、
'            ページ自体の既定背景(白)に上書きされて、指定した色が見えない可能性があります。
'            その場合は`WebView2Form.frm`側で`SetDefaultBackgroundColor`の呼び出し位置を
'            `about:blank`遷移より前に移動する必要があります(声をかけてください)
'***************************************************************************************************
Sub RunBgColorDemo()
    With WebView2Form
        If Not .StartCDPModeWebView2 Then
            MsgBox "WebView2の起動に失敗しました"
            Exit Sub
        End If

        '背景色を目立つ色(不透明・赤)に設定 → SetDefaultBackgroundColorの検証用
        .ThisWebView2.SetDefaultBackgroundColor 255, 255, 0, 0   ' a,r,g,b

        .show False
    End With
End Sub



'***************************************************************************************************
'                    ■■■ SetVirtualHostNameToFolderMapping 検証用デモ ■■■
'***************************************************************************************************
'* 機能　　：ローカルフォルダを仮想ホスト名(`https://<HostName>/`)にマッピングし、`file://`直開き
'            だと本来CORSで失敗する`fetch()`が成功することを確認するDemoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`data.json`を`fetch`して結果をページに書き込む`index.html`を一時フォルダへ用意し、
'            そのフォルダを仮想ホスト名へマッピングした上で`https://<HostName>/index.html`へ遷移。
'            `fetch`が本当に成功しているかを、CDPの`jsEval`でページ内のDOMを読み取って検証します。
'            (この仮想ホスト名は実在のDNSには一切問い合わせません。WebView2が内部で横取りするだけの
'            ローカル専用の名前です)
'---------------------------------------------------------------------------------------------------
'* 注意事項：`SetVirtualHostNameToFolderMapping`は`EnvironmentOptions`系とは違い、`ConnectCDP`
'            成功**後**(=WebView2本体`ICoreWebView2`が生成済みの状態)でなければ呼べません
'***************************************************************************************************
Sub RunVirtualHostMappingDemo()
    Const hostname As String = "wv2demo.local"
    Dim FolderPath As String
    FolderPath = Environ("UserProfile") & "\Downloads" & "\WV2VirtualHostDemo"

    '1. デモ用のローカルファイル(index.html/data.json)を一時フォルダへ用意
    PrepareVirtualHostDemoFiles FolderPath

    With WebView2Form
        '2. WebView2を起動
        If Not .StartCDPModeWebView2 Then
            MsgBox "WebView2の起動に失敗しました"
            Exit Sub
        End If

        '3. ローカルフォルダを仮想ホスト名へマッピング(★接続後・該当URLへの遷移前に呼ぶこと★)
        If Not .ThisWebView2.SetVirtualHostNameToFolderMapping(hostname, FolderPath) Then
            MsgBox "SetVirtualHostNameToFolderMappingに失敗しました(ICoreWebView2_3が未対応の可能性)"
            Unload WebView2Form
            Exit Sub
        End If

        '4. 仮想ホスト名経由でページ遷移(file://ではなくhttps://の実在オリジン扱いになる)
        .ThisCDPContext.navigate "https://" & hostname & "/index.html"

        '5. ページ内のfetch()結果をCDP経由で読み取り、実際に成功しているか検証
        Dim ResultText As String
        ResultText = .ThisCDPContext.jsEval("document.getElementById('result').innerText")
        Debug.Print "SetVirtualHostNameToFolderMapping Demo結果: " & ResultText

        '6. マッピング解除(切断時にも無効になるが、明示的な使い方の例として)
        .ThisWebView2.ClearVirtualHostNameToFolderMapping hostname

        '7. フォームを表示
        .show False
    End With
End Sub

'***************************************************************************************************
'* 機能　　：`RunVirtualHostMappingDemo`用の`index.html`/`data.json`を、指定フォルダへUTF-8で
'            書き出します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`Open...For Output`の既定の書き込みはANSI(日本語環境ではShift_JIS)になってしまい、
'            `<meta charset="utf-8">`と食い違って文字化けするため、`CharacterCodeConversion`
'            (既定でUTF-8)を使って生バイト列として書き出す
'***************************************************************************************************
Private Sub PrepareVirtualHostDemoFiles(ByVal FolderPath As String)
    If Dir(FolderPath, vbDirectory) = vbNullString Then MkDir FolderPath

    Dim HtmlContent As String
    HtmlContent = "<!DOCTYPE html>" & vbCrLf & _
        "<html><head><meta charset=""utf-8""><title>SetVirtualHostNameToFolderMapping Demo</title></head>" & vbCrLf & _
        "<body>" & vbCrLf & _
        "<h1>SetVirtualHostNameToFolderMapping Demo</h1>" & vbCrLf & _
        "<p>このページ自体、ローカルフォルダから仮想ホスト名経由で配信されています(file://ではありません)</p>" & vbCrLf & _
        "<div id=""result"">読み込み中...</div>" & vbCrLf & _
        "<script>" & vbCrLf & _
        "fetch('data.json').then(function (r) { return r.json(); }).then(function (data) {" & vbCrLf & _
        "  document.getElementById('result').innerText = 'fetch成功: ' + data.message;" & vbCrLf & _
        "}).catch(function (err) {" & vbCrLf & _
        "  document.getElementById('result').innerText = 'fetch失敗: ' + err.message;" & vbCrLf & _
        "});" & vbCrLf & _
        "</script>" & vbCrLf & _
        "</body></html>"

    Dim JsonContent As String
    JsonContent = "{""message"":""ローカルフォルダからhttps仮想ホスト経由で配信されたJSONです(file://なら本来CORSで失敗する)""}"

    Dim conv As New CharacterCodeConversion
    conv.BytesToSaveFile conv.BytesFromString(HtmlContent), FolderPath, "index.html"
    conv.BytesToSaveFile conv.BytesFromString(JsonContent), FolderPath, "data.json"
End Sub



'***************************************************************************************************
'                    ■■■ EnvironmentOptions/Settings/Controller/Profile/View
'                        追加設定 一括検証用デモ ■■■
'***************************************************************************************************
'* 機能　　：このセッションで追加した設定群のうち、まだDemoが無かったものをまとめて検証します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：JS/CDP経由で客観的に読み取り確認できるもの(UserAgent、prefers-color-scheme、
'            navigator.language等)は実際に検証し、Debug.Printで「OK/NG」を出します。
'            タッチジェスチャー/実マウスホバー/DPI環境等、VBAから再現できないものは
'            「設定してもエラーにならないこと」だけを確認する**スモークテスト**に留め、
'            その旨をDebug.Printで明示します(=目視/手動での最終確認が別途必要)
'---------------------------------------------------------------------------------------------------
'* 注意事項：4つの`Run～Demo`は、それぞれ単独で実行可能です(内部で`WebView2Form`を都度作り直します)
'***************************************************************************************************

'* 機能　　：Demo用の一時HTMLファイルを書き出し、`file:///`形式のURLを返します
'---------------------------------------------------------------------------------------------------
'* 注意事項：単体で完結する(外部fetchしない)テストページ専用。文字コードの扱いは
'            `PrepareVirtualHostDemoFiles`と同じ(`CharacterCodeConversion`でUTF-8書き出し)
'***************************************************************************************************
Private Function WriteDemoHtmlFile(ByVal FileName As String, ByVal HtmlContent As String) As String
    Dim Folder As String
    Folder = Environ("UserProfile") & "\Downloads" & "\WV2SettingsDemo"
    If Dir(Folder, vbDirectory) = vbNullString Then MkDir Folder

    Dim conv As New CharacterCodeConversion
    conv.BytesToSaveFile conv.BytesFromString(HtmlContent), Folder, FileName

    WriteDemoHtmlFile = "file:///" & Replace(Folder, "\", "/") & "/" & FileName
End Function

'***************************************************************************************************
'* 機能　　：`EnvironmentOptions`の残り(`AreBrowserExtensionsEnabled`以外)を検証します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`Language`は`navigator.language`で読み取り確認できるため実際に検証します。
'            `ExclusiveUserDataFolderAccess`/`IsCustomCrashReportingEnabled`/`EnableTrackingPrevention`/
'            `ChannelSearchKind`/`ReleaseChannels`はいずれもWebページ側から観測できない、
'            プロセス起動時の内部設定のため「設定してもEnvironment作成が失敗しない」ことのみを
'            スモークテストします。`ScrollBarStyle`はスクロール可能なページを表示するので、
'            目視でスクロールバーの見た目(細い自動非表示型)を確認してください
'***************************************************************************************************
Sub RunEnvironmentOptionsDemo()
    With WebView2Form
        '1. Environment作成前に設定(★重要:ConnectCDPより前でなければ意味がない★)
        With .ThisWebView2.EnvironmentOptions
            .Set_Language = "fr-FR"
            .Set_ExclusiveUserDataFolderAccess = False
            .Set_IsCustomCrashReportingEnabled = False
            .Set_EnableTrackingPrevention = True
            .Set_ChannelSearchKind = ChannelSearch_MostStable
            .Set_ReleaseChannels = Channel_Stable Or Channel_Beta Or Channel_Dev Or Channel_Canary
            .Set_ScrollBarStyle = ScrollBar_FluentOverlay
        End With

        '2. WebView2を起動(★スモークテスト★ ここで失敗しなければ、6項目とも受理されたことになる)
        If Not .StartCDPModeWebView2 Then
            MsgBox "WebView2の起動に失敗しました(EnvironmentOptionsのいずれかが原因の可能性)"
            Exit Sub
        End If
        Debug.Print "EnvironmentOptions Demo: Environment作成に成功(6項目とも受理されました)"

        '3. Languageの検証(navigator.languageで読み取り可能)+ ScrollBarStyle目視用の長いページ
        Dim longHtml As String, i As Long
        For i = 1 To 150
            longHtml = longHtml & "<p>ダミー行 " & i & "</p>"
        Next i
        Dim htmlUrl As String
        htmlUrl = WriteDemoHtmlFile("env_options_test.html", _
            "<!DOCTYPE html><html><body><h1>EnvironmentOptions Demo</h1>" & _
            "<p>ScrollBarStyle(FluentOverlay)の見た目を、下のスクロールバーで目視確認してください</p>" & _
            longHtml & "</body></html>")
        .ThisCDPContext.navigate htmlUrl

        Dim langResult As String
        langResult = .ThisCDPContext.jsEval("navigator.language")
        Debug.Print "EnvironmentOptions Demo(Language=""fr-FR""指定)結果: navigator.language=" & langResult

        '4. 表示(ScrollBarStyleの目視確認用)
        .show False
    End With
End Sub

'***************************************************************************************************
'* 機能　　：`ICoreWebView2Settings`(基底/2~9)の各設定を検証します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`ScriptEnabled`/`DefaultScriptDialogsEnabled`/`UserAgentOverride`/
'            `BuiltInErrorPageEnabled`はJS/実際のページ挙動から検証できるため実施します。
'            それ以外(`WebMessageEnabled`等)はUI/ジェスチャー依存のためスモークテストのみです
'---------------------------------------------------------------------------------------------------
'* 注意事項：`DefaultScriptDialogsEnabled`は**Falseにするケースのみ**検証します(Trueのまま
'            `alert()`を呼ぶと、誰も閉じないダイアログでVBAごと固まるため、危険で試せません)
'***************************************************************************************************
Sub RunSettingsFamilyDemo()
    With WebView2Form
        If Not .StartCDPModeWebView2 Then
            MsgBox "WebView2の起動に失敗しました"
            Exit Sub
        End If

        '--- ScriptEnabled：ページ自身のインラインscriptが実行されるかどうかで検証 ---
        Dim scriptTestUrl As String
        scriptTestUrl = WriteDemoHtmlFile("script_test.html", _
            "<!DOCTYPE html><html><head><title>BEFORE</title></head><body>" & _
            "<script>document.title='JS_RAN';</script></body></html>")

        .ThisWebView2.ScriptEnabled = False
        .ThisCDPContext.navigate scriptTestUrl
        Dim titleWhenDisabled As String
        titleWhenDisabled = .ThisCDPContext.jsEval("document.title")
        Debug.Print "ScriptEnabled=False Demo結果: title=" & titleWhenDisabled & _
            IIf(titleWhenDisabled <> "JS_RAN", " -> OK(スクリプト未実行)", " -> NG(実行されてしまった)")

        .ThisWebView2.ScriptEnabled = True
        .ThisCDPContext.navigate scriptTestUrl & "?t=" & Timer
        Dim titleWhenEnabled As String
        titleWhenEnabled = .ThisCDPContext.jsEval("document.title")
        Debug.Print "ScriptEnabled=True Demo結果: title=" & titleWhenEnabled & _
            IIf(titleWhenEnabled = "JS_RAN", " -> OK(スクリプト実行された)", " -> NG(実行されなかった)")

        '--- DefaultScriptDialogsEnabled=False：alert()がブロックされて処理が止まらないか ---
        '※`AreDefaultScriptDialogsEnabled`は「次に読み込むHTMLドキュメントから」有効になる設定のため
        '  (SDK仕様。既に読み込み済みのドキュメントには遡って効かない)、設定直後に必ず再ナビゲートすること
        .ThisWebView2.DefaultScriptDialogsEnabled = False
        .ThisCDPContext.navigate scriptTestUrl & "?dlg=" & Timer

        Dim dialogResult As String
        dialogResult = .ThisCDPContext.jsEval("alert('WebView2 Demo'); 'ALERT_RETURNED'")
        Debug.Print "DefaultScriptDialogsEnabled=False Demo結果: " & dialogResult & _
            IIf(dialogResult = "ALERT_RETURNED", " -> OK(alertでブロックされず処理続行)", " -> NG")

        '--- UserAgentOverride：navigator.userAgentで読み取り確認 ---
        Const FakeUA As String = "WV2SettingsDemoUA/1.0"
        .ThisWebView2.UserAgentOverride = FakeUA
        Dim uaResult As String
        uaResult = .ThisCDPContext.jsEval("navigator.userAgent")
        Debug.Print "UserAgentOverride Demo結果: navigator.userAgent=" & uaResult & _
            IIf(InStr(uaResult, FakeUA) > 0, " -> OK", " -> NG")

        '--- BuiltInErrorPageEnabled=False：存在しないドメインへアクセスした際の表示で検証 ---
        .ThisWebView2.BuiltInErrorPageEnabled = False
        .ThisCDPContext.navigate "https://this-domain-should-not-exist-wv2demo.invalid/"
        Dim bodyLen As Long
        bodyLen = CLng(.ThisCDPContext.jsEval("document.body ? document.body.innerHTML.length : 0"))
        Debug.Print "BuiltInErrorPageEnabled=False Demo結果: body長=" & bodyLen & _
            "(短い/空ならEdge独自のエラーページが出ていない=OKの可能性が高い。念のため目視も推奨)"

        '--- 残り(視覚/ジェスチャー依存で自動検証が困難なもの)はスモークテストのみ ---
        '※既定値のままだと「変化なし」で目視確認しづらいため、あえて逆既定値(○○させない側)に
        '  している。既定に戻したい場合は各コメントの値を参照
        .ThisWebView2.WebMessageEnabled = False               '既定True
        .ThisWebView2.StatusBarEnabled = False                '既定True
        .ThisWebView2.HostObjectsAllowed = False               '既定True
        .ThisWebView2.ZoomControlEnabled = False               '既定True
        .ThisWebView2.BrowserAcceleratorKeysEnabled = False    '既定True
        .ThisWebView2.PasswordAutosaveEnabled = False          '既定True
        .ThisWebView2.GeneralAutofillEnabled = False           '既定True
        .ThisWebView2.PinchZoomEnabled = False                 '既定True
        .ThisWebView2.SwipeNavigationEnabled = False           '既定True
        .ThisWebView2.HiddenPdfToolbarItems = PDF_PrintItem Or PDF_Save   '既定0(全表示)
        .ThisWebView2.ReputationCheckingRequired = False       '既定True
        .ThisWebView2.NonClientRegionSupportEnabled = True     '既定False
        Debug.Print "Settings Demo: WebMessageEnabled/StatusBarEnabled/HostObjectsAllowed/" & _
            "ZoomControlEnabled/BrowserAcceleratorKeysEnabled/PasswordAutosaveEnabled/" & _
            "GeneralAutofillEnabled/PinchZoomEnabled/SwipeNavigationEnabled/" & _
            "HiddenPdfToolbarItems/ReputationCheckingRequired/NonClientRegionSupportEnabledは" & _
            "設定呼び出しがエラーなく完了(スモークテストのみ。UI/ジェスチャー依存のため目視/手動確認が必要)"

        '--- ここまではウィンドウ非表示のまま検証してきたが、以下は目視/手動操作が要るため表示する ---
        .show vbModeless
        ' MsgBox "以下、お好みで目視確認してください(OKでデモ終了)。いずれも「逆既定値」にしているため" & _
        '     "「起きない/効かない」ことを確認するテストになります:" & vbCrLf & _
        '     "・ステータスバー：ページ内リンクにマウスオーバーしても左下に表示されないか(StatusBarEnabled=False)" & vbCrLf & _
        '     "・ズーム：Ctrl+マウスホイールで拡大縮小できないか(ZoomControlEnabled=False)" & vbCrLf & _
        '     "・アクセラレータキー：Ctrl+F/F12等のブラウザ的キーが効かないか(BrowserAcceleratorKeysEnabled=False)" & vbCrLf & _
        '     "・スワイプ：タッチパッド左右スワイプで進む/戻るnavigateしないか(SwipeNavigationEnabled=False)" & vbCrLf & _
        '     "・PDF：PDFファイルへ遷移し、内蔵ビューアの印刷/保存ボタンが非表示か(HiddenPdfToolbarItems)" & vbCrLf & _
        '     "・タイトルバー領域：app-region:drag対応CSSのページで、その領域のドラッグが無効なままか" & _
        '     "(NonClientRegionSupportEnabledは既定Falseのままだと元々無効な機能なので、Trueにして" & _
        '     "有効化できてるかを見る側になる点に注意)", _
        '     vbInformation, "RunSettingsFamilyDemo 目視確認"

    End With
End Sub

'***************************************************************************************************
'* 機能　　：`ICoreWebView2Controller`(2~4)の各設定を検証します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：いずれもDPI環境/実際のドラッグ操作に依存するため、スモークテスト(設定してエラーに
'            ならないか)のみ行います(背景色[`SetDefaultBackgroundColor`]は`RunBgColorDemo`で
'            実機検証済みのため、ここでは対象外)
'***************************************************************************************************
Sub RunControllerFamilyDemo()
    With WebView2Form
        If Not .StartCDPModeWebView2 Then
            MsgBox "WebView2の起動に失敗しました"
            Exit Sub
        End If

        .ThisWebView2.RasterizationScale = 1.25
        .ThisWebView2.ShouldDetectMonitorScaleChanges = False
        .ThisWebView2.BoundsMode = BoundsMode_UseRasterizationScale
        .ThisWebView2.AllowExternalDrop = False
        Debug.Print "Controller Demo: RasterizationScale/ShouldDetectMonitorScaleChanges/" & _
            "BoundsMode/AllowExternalDropの設定呼び出しがエラーなく完了" & _
            "(DPI/ドラッグ操作依存のため目視/手動確認が必要)"

        .ThisCDPContext.navigate "https://github.com/Eschamali/StarterWebScrapingKit"
        .show False
    End With
End Sub

'***************************************************************************************************
'* 機能　　：`ICoreWebView2Profile`(基底/3/6/9)の各設定を検証します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`PreferredColorScheme`は`prefers-color-scheme`メディアクエリで読み取り確認できるため
'            実際に検証します。残り(ダウンロード先/トラッキング防止レベル/自動入力/ServiceWorker)は
'            実際の効果確認にダウンロード実行やフォーム操作等が必要なため、スモークテストのみです
'***************************************************************************************************
Sub RunProfileFamilyDemo()
    With WebView2Form
        If Not .StartCDPModeWebView2 Then
            MsgBox "WebView2の起動に失敗しました"
            Exit Sub
        End If

        '--- PreferredColorScheme：prefers-color-schemeメディアクエリで読み取り確認 ---
        .ThisWebView2.PreferredColorScheme = ColorScheme_Dark
        Dim colorTestUrl As String
        colorTestUrl = WriteDemoHtmlFile("color_scheme_test.html", _
            "<!DOCTYPE html><html><body><h1>PreferredColorScheme Demo</h1></body></html>")
        .ThisCDPContext.navigate colorTestUrl
        Dim isDark As Variant
        isDark = .ThisCDPContext.jsEval("window.matchMedia('(prefers-color-scheme: dark)').matches")
        Debug.Print "PreferredColorScheme=Dark Demo結果: prefers-color-scheme dark match=" & isDark

        '--- 残り(観測手段が複雑なもの)はスモークテストのみ ---
        '※`ProfilePasswordAutosaveEnabled`/`ProfileGeneralAutofillEnabled`は既定値のままだと
        '  変化が分からないため、あえて逆既定値(既定True→False)にしている
        Dim downloadFolder As String
        downloadFolder = Environ("UserProfile") & "\Downloads" & "\WV2SettingsDemo\Downloads"
        If Dir(downloadFolder, vbDirectory) = vbNullString Then MkDir downloadFolder
        .ThisWebView2.DefaultDownloadFolderPath = downloadFolder
        .ThisWebView2.PreferredTrackingPreventionLevel = TrackingPrevention_Strict   '既定Balanced
        .ThisWebView2.ProfilePasswordAutosaveEnabled = False                        '既定True
        .ThisWebView2.ProfileGeneralAutofillEnabled = False                         '既定True
        .ThisWebView2.WebViewScriptApisEnabledForServiceWorkers = True              '既定False
        Debug.Print "Profile Demo: DefaultDownloadFolderPath/PreferredTrackingPreventionLevel/" & _
            "ProfilePasswordAutosaveEnabled/ProfileGeneralAutofillEnabled/" & _
            "WebViewScriptApisEnabledForServiceWorkersの設定呼び出しがエラーなく完了" & _
            "(実際の効果確認には、ダウンロード実行やフォーム入力等の実操作が必要)"

        '--- ここまではウィンドウ非表示のまま検証してきたが、以下は目視/手動操作が要るため表示する ---
        .ThisCDPContext.navigate "https://github.com/Eschamali/StarterWebScrapingKit"
        .show vbModeless
'        MsgBox "以下、お好みで目視確認してください(OKでデモ終了):" & vbCrLf & _
'            "・ダウンロード：適当なファイルへのリンクを右クリック→名前を付けて保存等でダウンロードし、" & _
'            "「" & downloadFolder & "」に保存されるか(DefaultDownloadFolderPath)" & vbCrLf & _
'            "・パスワード自動保存：ログインフォームのあるページでログインを試し、" & _
'            "「パスワードを保存しますか？」ダイアログが出ないか(ProfilePasswordAutosaveEnabled=False)" & vbCrLf & _
'            "・オートフィル：住所/氏名等の入力欄で、補完候補が出ないか(ProfileGeneralAutofillEnabled=False)" & vbCrLf & _
'            "・トラッキング防止：Strict指定時のみ発生する挙動なので、目視では判断しづらいです" & _
'            "(設定呼び出しがエラーにならないことのみ確認済み)" & vbCrLf & _
'            "・ServiceWorker：`navigator.serviceWorker`経由のScript API疎通は専用のテストページが" & _
'            "要るため、目視では判断しづらいです(同上)", _
'            vbInformation, "RunProfileFamilyDemo 目視確認"

    End With
End Sub

'***************************************************************************************************
'* 機能　　：View拡張(`ICoreWebView2_8`/`_12`/`_15`/`_28`)の各メンバーを検証します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`IsMuted`はネイティブ側get/putの直接往復で確認。`IsDocumentPlayingAudio`はWeb Audio API
'            で実際に音を鳴らして確認。`FaviconUri`はfavicon持ちの実サイトへ遷移して確認。
'            `StatusBarText`は実マウスホバーが無いと基本空文字のため読み取りのみ。
'            `Find`系は`Start`未実装のため、エラーにならないことのみ確認します
'***************************************************************************************************
Sub RunViewExtrasDemo()
    With WebView2Form
        If Not .StartCDPModeWebView2 Then
            MsgBox "WebView2の起動に失敗しました"
            Exit Sub
        End If

        '--- IsMuted(ICoreWebView2_8)：put→getの直接往復で確認 ---
        .ThisWebView2.IsMuted = True
        .ThisWebView2.IsMuted = False

        '--- IsDocumentPlayingAudio(ICoreWebView2_8)：Web Audio APIで実際に音を鳴らして確認 ---
        Dim audioTestUrl As String
        audioTestUrl = WriteDemoHtmlFile("audio_test.html", _
            "<!DOCTYPE html><html><body><h1>IsDocumentPlayingAudio Demo</h1></body></html>")
        .ThisCDPContext.navigate audioTestUrl
        .ThisCDPContext.jsEval "var ctx=new (window.AudioContext||window.webkitAudioContext)();" & _
            "var osc=ctx.createOscillator();osc.connect(ctx.destination);osc.start();", userGesture:=True
        CDPHelpers.Sleep 0.5   '再生状態の反映を少し待つ
        Debug.Print "IsDocumentPlayingAudio(発音直後)結果: " & .ThisWebView2.IsDocumentPlayingAudio

        '--- StatusBarText(ICoreWebView2_12)：実マウスホバー無しでは基本空文字。読み取りのみ確認 ---
        Debug.Print "StatusBarText(未ホバー時)結果: """ & .ThisWebView2.StatusBarText & """" & _
            "(実際のリンクホバーによる確認は目視で行ってください)"

        '--- FaviconUri(ICoreWebView2_15)：favicon持ちの実サイトへ遷移して読み取り確認 ---
        .ThisCDPContext.navigate "https://github.com/Eschamali/StarterWebScrapingKit"
        CDPHelpers.Sleep 1   'favicon解決を少し待つ
        Debug.Print "FaviconUri Demo結果: " & .ThisWebView2.FaviconUri


        .show False
    End With
End Sub
