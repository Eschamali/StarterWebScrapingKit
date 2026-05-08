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
' ・`inspect/#remote-debugging`による今、目の前のブラウザから、自動制御
'
' 【パイプ役割】PowerShell が NamedPipe サーバー（待受け）、Excel がクライアント（接続側）です。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long



'***************************************************************************************************
'                               ■■■ 変数/定数宣言 ■■■
'***************************************************************************************************
'名前付きパイプ名である接頭辞
Private Const PIPE_Landmark As String = "\\.\pipe\"

'名前付きパイプ関連の設定定数
Private Const GENERIC_READ            As Long = &H80000000  '名前付きパイプの読み取り
Private Const GENERIC_WRITE           As Long = &H40000000  '名前付きパイプへ書き込み
Private Const OPEN_EXISTING           As Long = 3           '開設済み名前付きパイプへの接続
Private Const FILE_ATTRIBUTE_NORMAL   As Long = &H80

Private Const INVALID_HANDLE_VALUE  As LongPtr = -1 '失敗サイン

'このクラスで使用するハンドル値
Private hNamePipe As LongPtr



'***************************************************************************************************
'                               ■■■ メイン処理 ■■■
'***************************************************************************************************
'* 機能　　：PowerShell が作成済みの名前付きパイプにクライアントとして接続し、接続情報等を記録します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：エラーコード　 ※0＝成功
'* 引数　　：UserName       接続名称
'---------------------------------------------------------------------------------------------------
'* 注意事項：・PowerShell 側でまず、名前付きパイプを生成しないと接続に失敗します
'            ・ツール側の仕様上、この`UserName`は設定シートの`ユーザーデータフォルダ名`と共存します。
'            　スクレイピングする際は、この設定値に紐づけられたパイプ値としてやり取りが行われます
'***************************************************************************************************
Public Function ConnectNamePipe(UserName As String) As Long
    '0. Excelに記録中のパイプがあったらCloseする
    If deserialize(UserName) Then CloseHandle hNamePipe

    '1. 名前付きパイプのフルパス作成
    Dim UsePipeName As String: UsePipeName = PIPE_Landmark & UserName

    '2. そのフルパスの名前付きパイプへ接続
    hNamePipe = CreateFile(UsePipeName, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    '3. 成功判定
    If hNamePipe = INVALID_HANDLE_VALUE Then ConnectNamePipe = Err.LastDllError: Exit Function

    '4. パイプハンドルを記録するプロシージャへ
    serialize UserName
End Function



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
    '※主要となる情報以外は一旦、一律0とし、必要なデータを`Dictionary`に詰める
    Dim tmp As New Dictionary
    tmp.Add "hStdOutRd", 0
    tmp.Add "hStderrOutRd", 0
    tmp.Add "hStdInWr", 0
    tmp.Add "hCDPOutRd", hNamePipe
    tmp.Add "hCDPInWr", hNamePipe
    tmp.Add "dwProcessId", 0

    'Excelテーブルに、名前付きパイプハンドル情報を記録する
    Set ShSetting01_StartBrowser.serializeToTable(UserName, 1, "WebSocketViaNamedPipe.serialize") = tmp


    '------------------ 2. タブ情報の記録準備 ------------------
    '※主要となる情報以外は一旦、一律空欄とし、必要なデータを`Dictionary`に詰める
    tmp.RemoveAll
    tmp.Add "lngLastID", 1
    tmp.Add "sessionID", ""
    tmp.Add "targetID", ""

    'Excelテーブルに、タブ情報欄を確保する
    Set ShSetting01_StartBrowser.serializeToTable(UserName, 2, "WebSocketViaNamedPipe.serialize") = tmp

End Sub

'***************************************************************************************************
'* 機能　　：接続情報をExcelテーブルから取得し、クラス内変数に適用します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName   接続名称
'---------------------------------------------------------------------------------------------------
'* 詳細説明：リアタッチする際にこれを呼び出します
'* 注意事項：`remote-debugging-pipe`と共存する都合上、一部しか取得しません
'***************************************************************************************************
Private Function deserialize(UserName As String) As Boolean
    'Excelテーブルから、`UserName`に紐づくパイプハンドル情報を読み込む
    Dim PipeInfo As Dictionary
    Set PipeInfo = ShSetting01_StartBrowser.deserializeFromTable(UserName, 1, "WebSocketViaNamedPipe.deserialize")

    '取得に失敗した場合は、接続情報が失ってるので、抜ける
    If PipeInfo Is Nothing Then Exit Function

    'クラス内変数に適用
    hNamePipe = PipeInfo("hCDPOutRd")
    'hNamePipe = PipeInfo("hCDPInWr")   '`hCDPOutRd`と同じなので省略

    '成功サインを返す
    deserialize = True
End Function
