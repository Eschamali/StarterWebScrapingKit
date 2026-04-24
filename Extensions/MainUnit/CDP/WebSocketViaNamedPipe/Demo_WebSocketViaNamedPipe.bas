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



'***************************************************************************************************
'                               ■■■ 初期設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：PowerShell が待受けしている名前付きパイプに Excel から接続します
'---------------------------------------------------------------------------------------------------
'* 注意事項：先に`StartWebSocket.ps1`を実行して待受けさせてください。待受けが無いと エラー番号:2 を返します
'***************************************************************************************************
Sub FirstStep()
    '識別名称を設定する
    Dim UseName As String
    With ShSetting01_StartBrowser
        'UseName = .Range(.UseRangeName(2, "Demo_WebSocketViaNamedPipe.FirstStep")).value  '設定セルから、ユーザ名を取得する場合
        UseName = DefaultName   'こちらで用意された`PowerShell`の名称で
    End With

    '名前付きパイプへクライアント接続
    Dim WebSocketMode As New WebSocketViaNamedPipe
    Dim ResultCode As Long: ResultCode = WebSocketMode.ConnectNamePipe(UseName)

    'エラーチェック
    Dim ErrorDetail As New WinApiError  'エラーコードから、詳細を取得するやつ
    If ResultCode Then
        MsgBox "名前付きパイプへの接続に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & ErrorDetail.GetMessage(ResultCode, "kernel32"), vbCritical, "ErrorCode: " & ResultCode
    Else
        MsgBox "名前付きパイプへの接続が完了しました。", vbInformation, "Success"
    End If
End Sub



'***************************************************************************************************
'                           ■■■テンプレートプロシージャ ■■■
'***************************************************************************************************
'* 注意事項：・事前に、パイプへ接続（ConnectNamePipe）を済ませること
'            ・専用の`PowerShell`（StartWebSocket.ps1）が起動中であること
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
    WebSocketCDP.quit   '実行と共に、名前付きパイプのハンドルもクリーンします
End Sub
