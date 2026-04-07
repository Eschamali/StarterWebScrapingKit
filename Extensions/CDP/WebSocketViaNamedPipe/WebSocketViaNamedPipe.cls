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
'              PowerShellを経由したWebSocket方式のブラウザ操作を担います
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
'共有のパイプ名である接頭辞
Private Const PIPE_NAME     As String = "\\.\pipe\"
Private Const DefaultName   As String = "ChromiumWebSocket"

'名前付きパイプのセッティング
Private Const PIPE_ACCESS_DUPLEX    As Long = &H3   '送受信可能
Private Const PIPE_TYPE_BYTE        As Long = &H0   'バイトモード
Private Const PIPE_WAIT             As Long = &H0   '同期モード
Private Const INVALID_HANDLE_VALUE  As LongPtr = -1 '失敗サイン

Private hNamePipe As LongPtr



'***************************************************************************************************
'                               ■■■ メイン処理 ■■■
'***************************************************************************************************
Sub OpenAndConnectNamePipe(Optional UserName As String = DefaultName)
    '1. パイプを作成
    Dim UsePipeName As String: UsePipeName = PIPE_NAME & UserName
    hNamePipe = CreateNamedPipe(UsePipeName, PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE Or PIPE_WAIT, 1, 2 ^ 20, 2 ^ 20, 0, 0)  'パイプ容量は、`CDPCore`に合わせる
    If hNamePipe = INVALID_HANDLE_VALUE Then
        MsgBox "パイプ: " & UsePipeName & "  の作成に失敗しました", vbCritical
        Exit Sub
    End If

    '2. 接続
    '★注意★ ここでExcelはPowerShellが繋いでくるまで「フリーズ（待機状態）」になります！
    Debug.Print "パイプを開設しました。PowerShellからの接続を待っています..."
    ConnectNamedPipe hNamePipe, 0

    '3. PowerShellが繋ぐとフリーズが解けてここに進む
    serialize UserName
End Sub

Private Function serialize(UserName As String) As String
'-------------------------------------------------------------------
' Create serialized Json string from the current browser instance.
'-------------------------------------------------------------------

    'パイプ情報の記録準備
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

    '作成情報の返却
    serialize = tmp


    'タブ情報の記録準備
    With WorksheetFunction
        tmp = "{""" & .TextJoin(""";""", False, _
                                    .TextJoin(""",""", False, UserName, "lngLastID", 1), _
                                    .TextJoin(""",""", False, UserName, "sessionID", ""), _
                                    .TextJoin(""",""", False, UserName, "targetID", "") _
        ) & """}"
    End With
    ShSetting01_StartBrowser.serializeToTable(2, "WebSocketViaNamedPipe.serialize") = Evaluate(tmp)

End Function

Public Sub deserialize(Optional UserName As String = DefaultName)
'--------------------------------------------------------------------
' Decipher serialized Json string and assign values to CDP variables.
'--------------------------------------------------------------------

    'Excelテーブルから、`UserName`に紐づくパイプハンドル情報を読み込む
    Dim PipeInfo As Dictionary
    Set PipeInfo = ShSetting01_StartBrowser.deserializeFromTable(UserName, 1, "WebSocketViaNamedPipe.deserialize")

    '取得に失敗した場合は、接続情報が失ってるので、抜ける
    If PipeInfo Is Nothing Then Exit Sub

    'クラス内変数に適用
    hNamePipe = PipeInfo("hCDPOutRd")
    'hNamePipe = PipeInfo("hCDPInWr")

End Sub

'***************************************************************************************************
'* 機能　　：正式にパイプハンドル情報をクリーニングします
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に、`deserialize` or `OpenAndConnectNamePipe` をした状態にすること
'***************************************************************************************************
Public Sub ClosePipeCDP()
    DisconnectNamedPipe hNamePipe
    CloseHandle hNamePipe
End Sub
