Attribute VB_Name = "WebDriverBiDiOnlineUpdater"
'***************************************************************************************************
'       ブラウザのバージョンに合わせて、chromium-bidi の mapperTab.js を取得・更新します
'---------------------------------------------------------------------------------------------------
'   npm の `latest` タグではなく、インストールしてるブラウザのバージョンを基準に mapperTab.js の版を選びます。
'   `chromium-bidi`のリリース仕様変更に伴い、少し複雑なロジックを求められることから、メインのリリースには含めず、
'   拡張機能ポジションとして、切り出させていただきました。exeのメタ情報を利用して起動不要で、バージョン確認します。
'
'   目的：実行端末にインストールしているEdge or Chrome のバージョンを基に、`jsdelivr`から適切な`mapperTab.js`の版をツールに保存します
'
'   照合モード：
'     ・Chrome のプリセット       ：完全一致（`157.0.8103.20` → `157.0.8103-20`）
'     ・Edge のプリセット/カスタム：2番目まで一致（`157.0.` で始まる版のうち最新）
'     ・どちらも該当なし          ：下限の `157.0.8090-0`
'
'   どのブラウザ・チャンネルを使うかは、ワークシート「ブラウザ起動設定」で設定します。
'
'   詳細・既知の制約は、このフォルダの ReadMe.md を参照してください。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
'npm（jsDelivr）関連
Private Const NPM_NAME          As String = "chromium-bidi"
Private Const MAPPER_JS_NAME    As String = "mapperTab.js"
Private Const JSDELIVR_DATA_URL As String = "https://data.jsdelivr.com/"
Private Const JSDELIVR_CDN_URL  As String = "https://cdn.jsdelivr.net/"
Private Const FLOOR_VERSION     As String = "157.0.8090-0"  '該当する版がないときに使う下限。`puppeteer`が採用した初めての版(https://github.com/puppeteer/puppeteer/blob/browsers-v3.2.4/package-lock.json#L33028)

'httpリクエストのタイムアウト値(ms)
Private Const HTTP_TIMEOUT_MS   As Long = 30000

'どの照合モードで版が決まったか
Private Enum OnlineUpdaterMatchMode
    omNone = 0          'まだ判定していない、または判定に失敗した
    omExact = 1         '完全一致
    omSecond = 2        '2番目まで一致
    omFloor = 3         '該当なしのため、下限を使った
End Enum

'最後の判定結果
Private Type OnlineUpdaterState
    browserVersion      As String
    MatchMode           As OnlineUpdaterMatchMode
    ResolvedVersion     As String
    Skipped             As Boolean
End Type
Private LastState As OnlineUpdaterState



'***************************************************************************************************
'                               ■■■ メイン処理 ■■■
'***************************************************************************************************
Sub RunUpdateWebDriverBiDi()
    '1. ワークシート設定に基づいたbrowser設定を渡す
    Dim BrowserHost As CDPHost: Set BrowserHost = New CDPHost
    BrowserHost.Init BrowserCapabilities, False

    '2. 情報を確保し、`Nothing`化
    Dim CurrentBrowserVersion  As String
    CurrentBrowserVersion = BrowserHost.getBrowserVersion: Debug.Print "Installed browser version: " & CurrentBrowserVersion
    Set BrowserHost = Nothing
    Set BrowserCapabilities = Nothing

    '3. まずは、バージョン一覧を取得する
    '※エラーの場合は、抜ける
    Dim VersionList As Collection
    Set VersionList = FetchNpmVersions
    If VersionList Is Nothing Then ResetState: Exit Sub

    '4. 得たバージョンリストを基に、最も近いバージョン番号を取得する
    Dim UsemapperTabjsVer As String
    UsemapperTabjsVer = SelectVersion(CurrentBrowserVersion, True, VersionList)
    Debug.Print "`mapperTab.js` version that is close to your installed browser version: " & UsemapperTabjsVer

    '5. 得たバージョンの`mapperTab.js`を取得
    '※エラーの場合は、抜ける
    Dim mapperTabjsStr As String
    mapperTabjsStr = DownloadMapperTab(UsemapperTabjsVer)
    If StrPtr(mapperTabjsStr) = 0 Then ResetState: Exit Sub

    '6. 更新処理
    Dim UpdateBiDi As New WebDriverBiDiCore
    If UpdateBiDi.UpdateWebDriverBiDi(mapperTabjsStr, UsemapperTabjsVer) Then MsgBox "インストールブラウザに近い、`mapperTab.js`によるアップデートに成功しました。" & vbCrLf & CurrentBrowserVersion, vbInformation, "Success - " & UsemapperTabjsVer Else MsgBox "アップデートに失敗しました。Excelテーブルに埋め込んだJavaScript文字列とアップロードファイルとの一致が確認できなかったか、適切なライセンス情報がありませんでした。" & vbCrLf & CurrentBrowserVersion, vbCritical, "failure"

    '7. 後始末
    ResetState
End Sub



'***************************************************************************************************
'                                   ■■■ http処理 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 各種Request
'---------------------------------------------------------------------------------------------------'
'***************************************************************************************************
'* 機能　　：jsDelivr から、chromium-bidi の版の一覧を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：版の文字列のCollection。失敗した場合は`Nothing`
'***************************************************************************************************
Private Function FetchNpmVersions() As Collection
    '1. リクエスト情報を作成し、結果を受け取る
    Dim res As WebResponse
    Set res = RequestTo(JSDELIVR_DATA_URL, "v1/package/npm/" & NPM_NAME, vbWebFormatJson)

    '2. httpコード確認
    If res.StatusCode <> vbWebStatusOk Then MsgBox NPM_NAME & " の版の一覧の取得に失敗しました。", vbCritical, "HTTP: " & res.StatusCode: Exit Function

    '3. バージョン一覧として返却
    Set FetchNpmVersions = res.Data("versions")
End Function

'***************************************************************************************************
'* 機能　　：jsDelivr から mapperTab.js をダウンロードし、ワークシートに保存します
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数の`mapperTab.js`の中身、そのもの。失敗時は、`vbnullstring`とします。
'* 引数　　：ver   mapperTab.js の版
'---------------------------------------------------------------------------------------------------
'* 注意事項：・改行はLFにそろえます。CRLFのままだと、テーブルへ分割して置くときに`CR`と`LF`の間で
'              分かれて、読み戻したときに一致しなくなることがあるためです
'            ・保存後に読み戻して一致を確かめ、一致しなければテーブルの中身を消します
'***************************************************************************************************
Private Function DownloadMapperTab(Ver As String) As String
    '1. リクエスト情報を作成し、結果を受け取る
    Dim res As WebResponse
    Set res = RequestTo(JSDELIVR_CDN_URL, "npm/" & NPM_NAME & "@" & Ver & "/out/Default/gen/src/" & MAPPER_JS_NAME, vbWebFormatPlainText)
    
    '2. httpコード確認
    If res.StatusCode <> vbWebStatusOk Then MsgBox MAPPER_JS_NAME & " " & Ver & " のダウンロードに失敗しました。", vbCritical, "HTTP: " & res.StatusCode: Exit Function

    '3. 中身として、返却
    DownloadMapperTab = res.Content
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 共通ヘルパー
'---------------------------------------------------------------------------------------------------'
'***************************************************************************************************
'* 機能　　：GETリクエストを送ります
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスオブジェクト - WebResponse
'* 引数　　：BaseUrl        ホスト名のURL（末尾は`/`）
'            Resource       ホスト名に続くパス
'***************************************************************************************************
Private Function RequestTo(BaseUrl As String, Resource As String, ResponseFormat As VbWebFormat) As WebResponse
    Dim req As WebRequest: Set req = New WebRequest
    With req
        .Method = vbWebHttpGet
        .ResponseFormat = ResponseFormat
        .Resource = Resource
    End With

    Dim Client As WebClient: Set Client = New WebClient
    With Client
        .BaseUrl = BaseUrl
        .TimeoutMS = HTTP_TIMEOUT_MS
    End With

    Set RequestTo = Client.Execute(req)
End Function



'***************************************************************************************************
'                             ■■■ 探索処理 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] メイン探索ロジック
'---------------------------------------------------------------------------------------------------'
'***************************************************************************************************
'* 機能　　：候補の版の中から、ブラウザのバージョンに合うものを選びます
'---------------------------------------------------------------------------------------------------
'* 返り値　：選んだ版。該当なしの場合は下限の版
'* 引数　　：BrowserVersion   ブラウザのバージョン（例：`157.0.8103.20`）
'            ExactMatch       `True`:完全一致を先に試す  `False`:2番目まで一致だけ
'            Candidates       npm の版の一覧（文字列のCollection）
'---------------------------------------------------------------------------------------------------
'* 注意事項：・判定の内訳は`LastMatchMode`で確認できます
'            ・`17.0.2`のような旧形式の版は、候補から除きます
'            ・完全一致がない場合は、2番目まで一致で探し直します
'***************************************************************************************************
Private Function SelectVersion(browserVersion As String, ExactMatch As Boolean, Candidates As Collection) As String
    Dim browserParts() As Long
    LastState.MatchMode = omNone

    '1. ブラウザのバージョンを分解（`A.B.C.D`）
    If Not ParseDottedVersion(browserVersion, browserParts) Then
        LastState.MatchMode = omFloor
        SelectVersion = FLOOR_VERSION
        Exit Function
    End If

    '2. 完全一致：`A.B.C.D` → `A.B.C-D`
    Dim cand As Variant, candParts() As Long
    If ExactMatch Then
        Dim exactName As String
        exactName = browserParts(0) & "." & browserParts(1) & "." & browserParts(2) & "-" & browserParts(3)
        For Each cand In Candidates
            If CStr(cand) = exactName Then
                LastState.MatchMode = omExact
                SelectVersion = exactName
                Exit Function
            End If
        Next
    End If

    '3. 2番目まで一致：`A.B.`が同じ版のうち、数値として最大のもの
    Dim best As String
    For Each cand In Candidates
        If ParseBidiVersion(CStr(cand), candParts) Then
            If candParts(0) = browserParts(0) And candParts(1) = browserParts(1) Then
                If LenB(best) = 0 Then
                    best = CStr(cand)
                ElseIf CompareBidiVersion(CStr(cand), best) > 0 Then
                    best = CStr(cand)
                End If
            End If
        End If
    Next

    If LenB(best) Then
        LastState.MatchMode = omSecond
        SelectVersion = best
    Else
        '4. 該当なし
        LastState.MatchMode = omFloor
        SelectVersion = FLOOR_VERSION
    End If
End Function

'***************************************************************************************************
'* 機能　　：chromium-bidi の版（`A.B.C-D`）を、数値として比べます
'---------------------------------------------------------------------------------------------------
'* 返り値　：`a`が大きければ正、同じなら0、小さければ負。どちらかが解釈できない場合は0
'---------------------------------------------------------------------------------------------------
'* 注意事項：文字列のまま比べると`157.0.8103-8`が`157.0.8103-15`より大きくなるため、各要素を数値で比べます
'***************************************************************************************************
Private Function CompareBidiVersion(a As String, b As String) As Long
    Dim pa() As Long, pb() As Long
    If Not ParseBidiVersion(a, pa) Then Exit Function
    If Not ParseBidiVersion(b, pb) Then Exit Function

    Dim i As Long
    For i = 0 To 3
        If pa(i) <> pb(i) Then
            CompareBidiVersion = IIf(pa(i) > pb(i), 1, -1)
            Exit Function
        End If
    Next
End Function

'***************************************************************************************************
'* 機能　　：最後の判定結果をクリアします
'***************************************************************************************************
Private Sub ResetState()
    Dim emptyState As OnlineUpdaterState
    LastState = emptyState
End Sub

'---------------------------------------------------------------------------------------------------
' [ SECTION ] バージョン分解
'---------------------------------------------------------------------------------------------------'
'***************************************************************************************************
'* 機能　　：ブラウザのバージョン（`A.B.C.D`）を、4つの数値に分けます
'***************************************************************************************************
Private Function ParseDottedVersion(Ver As String, ByRef parts() As Long) As Boolean
    Dim Items() As String: Items = Split(Ver, ".")
    If UBound(Items) <> 3 Then Exit Function

    ParseDottedVersion = ToLongParts(Items, parts)
End Function

'***************************************************************************************************
'* 機能　　：chromium-bidi の版（`A.B.C-D`）を、4つの数値に分けます
'---------------------------------------------------------------------------------------------------
'* 注意事項：`17.0.2`のような旧形式は`False`を返します
'***************************************************************************************************
Private Function ParseBidiVersion(Ver As String, ByRef parts() As Long) As Boolean
    Dim halves() As String: halves = Split(Ver, "-")
    If UBound(halves) <> 1 Then Exit Function

    Dim Items() As String: Items = Split(halves(0), ".")
    If UBound(Items) <> 2 Then Exit Function

    ReDim Preserve Items(3)
    Items(3) = halves(1)
    ParseBidiVersion = ToLongParts(Items, parts)
End Function

'***************************************************************************************************
'* 機能　　：4つの数字の文字列を、Longの配列にします（数字以外が含まれていれば`False`）
'***************************************************************************************************
Private Function ToLongParts(Items() As String, ByRef parts() As Long) As Boolean
    ReDim parts(3)

    Dim i As Long
    For i = 0 To 3
        If LenB(Items(i)) = 0 Or Items(i) Like "*[!0-9]*" Then Exit Function
        If Len(Items(i)) > 9 Then Exit Function
        parts(i) = CLng(Items(i))
    Next

    ToLongParts = True
End Function
