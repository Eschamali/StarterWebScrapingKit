Attribute VB_Name = "Demo_WebSocketViaNamedPipe"
'***************************************************************************************************
'               名前付きパイプの仕組みを利用したWebSocket連携機能のDemoコードです
'       Excel ←NamedPipe→ PowerShell ←WebSocket→ Chromium といった連携を前提とします
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ Demo用に定義 ■■■
'***************************************************************************************************
Private Const DefaultName   As String = "ChromiumWebSocket" 'デフォルト識別名称

Private ErrorDetail         As New WinApiError  'エラーコードから、詳細を取得するやつ



'***************************************************************************************************
'                        ■■■ 名前付きパイプ関連の処理プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：名前付きパイプを作成し、PowerShellが繋いでくるまで待機します
'---------------------------------------------------------------------------------------------------
'* 注意事項：ExcelはPowerShellが繋いでくるまで「フリーズ（待機状態）」になります
'***************************************************************************************************
Sub FirstStep()
    '識別名称を設定する
    Dim UseName As String
    With ShSetting01_StartBrowser
        'UseName = .Range(.UseRangeName(2, "Demo_WebSocketViaNamedPipe.FirstStep")).value  '設定セルから、ユーザ名を取得する場合
        UseName = DefaultName   'こちらで用意された`PowerShell`の名称で
    End With

    '名前付きパイプを作成し、接続処理
    Dim WebSocketMode As New WebSocketViaNamedPipe
    Dim ResultCode As Long: ResultCode = WebSocketMode.OpenAndConnectNamePipe(UseName)

    'エラーチェック
    If ResultCode Then
        MsgBox "名前付きパイプの作成に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & ErrorDetail.GetMessage(ResultCode, "kernel32"), vbCritical, "ErrorCode: " & ResultCode
    Else
        MsgBox "名前付きパイプの作成に成功し、接続が完了しました。", vbInformation, "Success"
    End If
End Sub

'***************************************************************************************************
'* 機能　　：作成済みの名前付きパイプハンドルを基に、再接続を行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：ExcelはPowerShellが繋いでくるまで「フリーズ（待機状態）」になります
'***************************************************************************************************
Sub ReConnect()
    '識別名称を設定する
    Dim UseName As String
    With ShSetting01_StartBrowser
        'UseName = .Range(.UseRangeName(2, "Demo_WebSocketViaNamedPipe.FirstStep")).value  '設定セルから、ユーザ名を取得する場合
        UseName = DefaultName   'こちらで用意された`PowerShell`の名称で
    End With

    'Excelテーブルから、既存の名前付きパイプを読み込み、再接続する
    Dim WebSocketMode As New WebSocketViaNamedPipe
    Dim ResultCode As Long: ResultCode = WebSocketMode.ReConnectNamedPipe(DefaultName)

    'エラーチェック
    If ResultCode Then
        MsgBox "既存の名前付きパイプへの再接続に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & ErrorDetail.GetMessage(ResultCode, "kernel32"), vbCritical, "ErrorCode: " & ResultCode
    Else
        MsgBox "既存の名前付きパイプへの再接続に成功しました。", vbInformation, "Success"
    End If
End Sub

'***************************************************************************************************
'* 機能　　：作成済みの名前付きパイプハンドルを基に、破棄処理を行います
'---------------------------------------------------------------------------------------------------
'* 注意事項：Excelに記録されてない作成済みの名前付きパイプハンドルは破棄できません。
''           破棄したにもかかわらず接続等でエラーが出る場合は、Excelプロセスの再起動が必要です。
'***************************************************************************************************
Sub cleanNamedPipe()
    '識別名称を設定する
    Dim UseName As String
    With ShSetting01_StartBrowser
        'UseName = .Range(.UseRangeName(2, "Demo_WebSocketViaNamedPipe.FirstStep")).value  '設定セルから、ユーザ名を取得する場合
        UseName = DefaultName   'こちらで用意された`PowerShell`の名称で
    End With

    'Excelテーブルから、既存の名前付きパイプを読み込み、clean処理しておく
    Dim WebSocket As New WebSocketViaNamedPipe
    WebSocket.ClosePipeCDP UseName
    Debug.Print "クリーン処理、完了"
End Sub



'***************************************************************************************************
'                           ■■■テンプレートプロシージャ ■■■
'***************************************************************************************************
'* 注意事項：・事前に、`ConnectNamedPipe`を済ませること
'            ・専用の`PowerShell`が起動中であること
'            ・WebSocket経由の場合は常に`.reattach`始まりとなります
'***************************************************************************************************
Sub WebSocketによる冒険の始まり()
    Dim WebSocketCDP As New CDPBrowser

    '識別名称を設定する
    Dim UseName As String
    With ShSetting01_StartBrowser
        'UseName = .Range(.UseRangeName(2, "Demo_WebSocketViaNamedPipe.FirstStep")).value  '設定セルから、ユーザ名を取得する場合
        UseName = DefaultName   'こちらで用意された`PowerShell`の名称で
    End With

    '1. まずは、既存のTargetIDに接続できるか？
    If Not WebSocketCDP.reattach(UseName) Then
        '既存のTargetIDが消えちゃったので、別タブへの再接続フェーズへ
        Debug.Print "既存の`targetID`への再接続に失敗。新しいタブか、今開いている直近のタブに再接続して、そこから処理を再開します。"

        '2. 未接続のタブに接続
        WebSocketCDP.getTab setMain:=True
        'WebSocketCDP.newTab setMain:=True     '新しいタブでもOK
    Else
        Debug.Print "既存の`targetID`への再接続に成功。このタブで処理を再開できます。"
    End If

    '↓ 3．再接続できたので、ここから、あなたのイメージをコードに落とし込む ↓
    '例）ページ遷移してみる




    'ブラウザを正常に閉じる
    WebSocketCDP.quit
End Sub
