VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebSocketViaNamedPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'               名前付きパイプの仕組みを利用したWebSocket連携機能を提供します。
'       Excel ←NamedPipe→ PowerShell ←WebSocket→ Chromium といった連携確保まで担います。
' 本来、このツールは`remote-debugging-pipe`特化で構築してる都合上、`remote-debugging-port`の直接統合は堅牢性を落としてしまう問題がありました。
' しかし、`PowerShell × NamedPipe`という仕組みにより、拡張機能として提供することで、コア部分をいじらず、実質対応可能となりました。
' 予め、所定の`PowerShell`を起動しておくことで、下記のようなニッチな目的にも対応できます
' ・AndroidスマートフォンのChromiumブラウザの自動制御
' ・WebView2向けの環境変数`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`に、`--remote-debugging-port=9222`を付与した自動制御
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, ByVal lpSecurityAttributes As LongPtr) As LongPtr
Private Declare PtrSafe Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long



'***************************************************************************************************
'                               ■■■ 変数/定数宣言 ■■■
'***************************************************************************************************
'名前付きパイプ名である接頭辞
Private Const PIPE_NAME     As String = "\\.\pipe\"

'名前付きパイプのセッティング
Private Const PIPE_ACCESS_DUPLEX    As Long = &H3   '送受信可能
Private Const PIPE_TYPE_BYTE        As Long = &H0   'バイトモード
Private Const PIPE_WAIT             As Long = &H0   '同期モード
Private Const INVALID_HANDLE_VALUE  As LongPtr = -1 '失敗サイン

'このクラスで使用するハンドル値
Private hNamePipe As LongPtr



'***************************************************************************************************
'                               ■■■ メイン処理 ■■■
'***************************************************************************************************
'* 機能　　：名前付きパイプを作成し、接続処理し、接続情報等を記録します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：エラーコード
'* 引数　　：UserName   接続名称
'---------------------------------------------------------------------------------------------------
'* 注意事項：ExcelはPowerShellが繋いでくるまで「フリーズ（待機状態）」になります
'***************************************************************************************************
Function OpenAndConnectNamePipe(UserName As String) As Long
    '1. 先着1つ分の、名前付き接続パイプを作成
    Dim UsePipeName As String: UsePipeName = PIPE_NAME & UserName
    hNamePipe = CreateNamedPipe(UsePipeName, PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE Or PIPE_WAIT, 1, 2 ^ 20, 2 ^ 20, 0, 0)  'パイプ容量は、`CDPCore`に合わせる

    '失敗時は、エラーコードを返却して、抜ける
    If hNamePipe = INVALID_HANDLE_VALUE Then OpenAndConnectNamePipe = Err.LastDllError: Exit Function

    '2. 接続
    '★注意★ ここでExcelはPowerShellが繋いでくるまで「フリーズ（待機状態）」になります！
    Application.StatusBar = "PowerShell からのパイプ接続 `" & UsePipeName & "` を待機してます...　※接続が来ない限り、フリーズは解消しません"
    ConnectNamedPipe hNamePipe, 0
    Application.StatusBar = False

    '3. PowerShellが繋ぐとフリーズが解けてここに進む
    serialize UserName      'パイプ情報などをExcelテーブルにメモする
End Function

'***************************************************************************************************
'* 機能　　：作成済みの名前付きパイプハンドルを基に、再接続を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   接続識別名称
'* 返り値  ：エラーコード
'---------------------------------------------------------------------------------------------------
'* 詳細説明：リアタッチする際にこれを呼び出します。一旦切断後、再接続する形です。
'* 注意事項：・事前に、`deserialize` or `OpenAndConnectNamePipe` をした状態にすること
'            ・ExcelはPowerShellが繋いでくるまで「フリーズ（待機状態）」になります
'***************************************************************************************************
Public Function ReConnectNamedPipe(Optional UserName As String) As Long
    '識別名称が指定されてる場合は、それを利用して、`hNamePipe`を上書き
    If StrPtr(UserName) Then deserialize UserName

    '一旦切断して、再接続する
    Application.StatusBar = "PowerShell からのパイプ再接続を待機してます...　※接続が来ない限り、フリーズは解消しません"
    DisconnectNamedPipe hNamePipe   'ErrorCode: 232 を回避するため
    ConnectNamedPipe hNamePipe, 0
    Application.StatusBar = False

    'エラーコードを返却　※正常時は、0
    ReConnectNamedPipe = Err.LastDllError
End Function

'***************************************************************************************************
'* 機能　　：正式にパイプハンドル情報をクリーニングします
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   接続識別名称
'            OnlyDisconnect     True ：ハンドル自体を使いまわす場合に
'                               False：ハンドル自体もクリーンしたい場合に
'---------------------------------------------------------------------------------------------------
'* 注意事項：・事前に、`deserialize` or `OpenAndConnectNamePipe` をした状態にすること
'            ・これをしないと、`OpenAndConnectNamePipe`や`ReConnectNamedPipe`にて、エラーになります
'***************************************************************************************************
Public Sub ClosePipeCDP(Optional UserName As String, Optional OnlyDisconnect As Boolean)
    '識別名称が指定されてる場合は、それを利用して、`hNamePipe`を上書き
    If StrPtr(UserName) Then deserialize UserName

    DisconnectNamedPipe hNamePipe
    If Not OnlyDisconnect Then CloseHandle hNamePipe
End Sub



'***************************************************************************************************
'                            ■■■ 設定情報読み書き処理 ■■■
'***************************************************************************************************
'* 機能　　：接続情報をExcelテーブルに残します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   接続名称
'---------------------------------------------------------------------------------------------------
'* 注意事項：`remote-debugging-pipe`と共存する都合上、一部は仮値にしてます
'***************************************************************************************************
Private Sub serialize(UserName As String)
    '------------------ 1. パイプ情報の記録準備 ------------------
    '※主要となる情報以外は一旦、一律0
    Dim tmp As String
    With WorksheetFunction
        tmp = "{""" & .TextJoin(""";""", False, _
                            .TextJoin(""",""", False, UserName, "hStdOutRd", 0), _
                            .TextJoin(""",""", False, UserName, "hStderrOutRd", 0), _
                            .TextJoin(""",""", False, UserName, "hStdInWr", 0), _
                            .TextJoin(""",""", False, UserName, "hCDPOutRd", hNamePipe), _
                            .TextJoin(""",""", False, UserName, "hCDPInWr", hNamePipe), _
                            .TextJoin(""",""", False, UserName, "dwProcessId", 0) _
        ) & """}"
    End With

    'Excelテーブルに、パイプハンドル情報を記録する
    'Evaluate関数を利用して、上記で作った記述方式で簡易的な2次元配列を作成します
    ShSetting01_StartBrowser.serializeToTable(1, "WebSocketViaNamedPipe.serialize") = Evaluate(tmp)


    '------------------ 2. タブ情報の記録準備 ------------------
    '※主要となる情報以外は一旦、一律空欄
    With WorksheetFunction
        tmp = "{""" & .TextJoin(""";""", False, _
                                    .TextJoin(""",""", False, UserName, "lngLastID", 1), _
                                    .TextJoin(""",""", False, UserName, "sessionID", ""), _
                                    .TextJoin(""",""", False, UserName, "targetID", "") _
        ) & """}"
    End With

    'Excelテーブルに、タブ情報欄を確保する
    'Evaluate関数を利用して、上記で作った記述方式で簡易的な2次元配列を作成します
    ShSetting01_StartBrowser.serializeToTable(2, "WebSocketViaNamedPipe.serialize") = Evaluate(tmp)

End Sub

'***************************************************************************************************
'* 機能　　：接続情報をExcelテーブルから取得し、クラス内変数に適用します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   接続名称
'---------------------------------------------------------------------------------------------------
'* 詳細説明：リアタッチする際にこれを呼び出します
'* 注意事項：`remote-debugging-pipe`と共存する都合上、一部しか取得しません
'***************************************************************************************************
Private Sub deserialize(UserName As String)
    'Excelテーブルから、`UserName`に紐づくパイプハンドル情報を読み込む
    Dim PipeInfo As Dictionary
    Set PipeInfo = ShSetting01_StartBrowser.deserializeFromTable(UserName, 1, "WebSocketViaNamedPipe.deserialize")

    '取得に失敗した場合は、接続情報が失ってるので、抜ける
    If PipeInfo Is Nothing Then Exit Sub

    'クラス内変数に適用
    hNamePipe = PipeInfo("hCDPOutRd")
    'hNamePipe = PipeInfo("hCDPInWr")   '`hCDPOutRd`と同じなので省略
End Sub
