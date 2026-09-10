VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WebMCPMazeForm 
   Caption         =   "WebMCP Maze Demo"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18540
   OleObjectBlob   =   "WebMCPMazeForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "WebMCPMazeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'           https://googlechromelabs.github.io/webmcp-tools/demos/webmcp-maze/ をWebView2で
'                embedし、迷路の移動操作をボタン経由でWebMCPツール実行に変換するDemoです
'---------------------------------------------------------------------------------------------------
'   本来はAIエージェントが`move`/`look`等のWebMCPツールを叩いて自動攻略するデモですが、
'   ここではその起動役をユーザー自身のボタン操作に置き換え、内部的にWebMCPを発動させています
'***************************************************************************************************



'***************************************************************************************************
'                               ■■■ 必要なWindowsAPI定義 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr



'***************************************************************************************************
'                                 ■■■ 迷路デモの固定設定 ■■■
'***************************************************************************************************
Private Const MazeURL As String = "https://googlechromelabs.github.io/webmcp-tools/demos/webmcp-maze/"
Private Const MaxStatusChars As Long = 900



'***************************************************************************************************
'                               ■■■ 制御に必要な変数定義 ■■■
'***************************************************************************************************
'各種オブジェクト
Private fWebView2               As CDPCoreViaWebView2
Private WithEvents fCDPEvent    As CDPCore      '非同期イベント処理用(参照保持を兼ねる)
Attribute fCDPEvent.VB_VarHelpID = -1
Private fCDPContext             As CDPContext   'タブ情報
Private fMCP                    As exCDP_WebMCP 'WebMCP制御

'自身の各ハンドルを保存する変数
Private myFormHwnd      As LongPtr
Private myEdgeFrameHwnd As LongPtr

'Frameの下マージン、および右側パネル追従用の初期値
Private BottomMargin       As Long
Private InitialInsideWidth As Long
Private PanelControlRefs(1 To 10)         As MSForms.Control
Private PanelControlOriginalLeft(1 To 10) As Long

'ゲーム状態
Private GameStarted     As Boolean  '`start_game`実行済みかどうか
Private FacingDirection As String   '`使う`ボタンが対象とする、直近の移動方向
Private StatusLog       As String   'StatusLabelに表示するログ本文



'***************************************************************************************************
'                              ■■■ ウィンドウスタイルの定数 ■■■
'***************************************************************************************************
Private Const GWL_STYLE         As Long = -16
Private Const WS_THICKFRAME     As Long = &H40000 'サイズ変更枠
Private Const WS_MAXIMIZEBOX    As Long = &H10000 '最大化ボタン
Private Const WS_MINIMIZEBOX    As Long = &H20000 '最小化ボタン



'***************************************************************************************************
'                            ■■■ WebView2 / WebMCP 接続 ■■■
'***************************************************************************************************
'* 機能　　：WebView2を起動し、迷路デモページに接続、WebMCP制御の準備までを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：成功可否論理値
'* 注意事項：追加の起動引数は、既存の「ブラウザ起動設定」シートの設定をそのまま流用します
'***************************************************************************************************
Private Function ConnectToMaze() As Boolean
    '1. WebView2の追加起動引数を、既存の設定シートから準備
    fWebView2.EnvironmentOptions.Set_AdditionalBrowserArguments = ShSetting01_StartBrowser.UseRangeID(3, "WebMCPMazeForm.ConnectToMaze")

    '2. WebView2を起動
    If Not fWebView2.ConnectCDP(ShSetting01_StartBrowser.CurrentUserName, myEdgeFrameHwnd) Then Exit Function

    '3. サイズをセットし、表示
    AdjustEdgeSize
    fWebView2.Visible = True

    '4. タブに接続(★WebView2は起動直後、既に"about:blank"を開いているタブしか無いため、
    '   ここで実URLを渡すと`Target.getTargets`に一致するタブが無く`Nothing`が返ってしまう。
    '   一旦"about:blank"で取得し、遷移は後段の`navigate`で別途行う★)
    Dim t As New CDPBrowser: t.reattachWebView2 ShSetting01_StartBrowser.CurrentUserName, fWebView2
    Set fCDPContext = t.getTab(setMain:=True, Url:="about:blank")

    '5. 非同期イベント処理に備える(CDPCoreの参照保持も兼ねる)
    Set fCDPEvent = t.ThisCDPCore

    '6. WebView2は「ドメインをenableすれば以後全イベントが自動で流れる」CDP-over-WebSocketとは違い、
    '   イベント名ごとに個別の購読(`SubscribeCdpEvent`)が必要(`CDPCoreViaWebView2.cls`の仕様)。
    '   これをしないと、`WebMCP.enable`しても`toolsAdded`等が一切届かず、使えるコマンドがいつまで
    '   経っても出てこない(=`onExistTool`が延々タイムアウトする)。実際の遷移より先に購読しておく
    '   ★VBEデバッグ時の注意★ 1つでもイベント購読した状態で`Stop`/ブレークポイント中に
    '   「リセット」を行うとExcelごとクラッシュします(想定内の既知動作)。`Stop`/ブレークポイント中の
    '   ローカルウィンドウ閲覧等の通常デバッグ操作自体は問題ありません。どうしてもリセットが必要な
    '   場合は、イミディエイトウィンドウで`fWebView2.UnsubscribeAllCdpEvents`を実行してから行うこと
    fWebView2.SubscribeCdpEvent "Page.frameNavigated"
    fWebView2.SubscribeCdpEvent "WebMCP.toolsAdded"
    fWebView2.SubscribeCdpEvent "WebMCP.toolsRemoved"
    fWebView2.SubscribeCdpEvent "WebMCP.toolInvoked"
    fWebView2.SubscribeCdpEvent "WebMCP.toolResponded"

    '7. WebMCP拡張クラスを初期化(内部で`WebMCP.enable`/`Page.enable`相当を行う)
    Set fMCP = New exCDP_WebMCP
    fMCP.Init fCDPContext

    '8. ここで初めて、迷路デモページへ実際に遷移する
    fCDPContext.navigate MazeURL

    '9. 成功
    ConnectToMaze = True
End Function



'***************************************************************************************************
'                                       ■■■ サイズ変更 ■■■
'***************************************************************************************************
'* 機能　　：WebView2のサイズをFrame内ににピッタリはめ込み、右側の操作パネルを常に右端に追従させます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・係数 1.333 は ポイント(VBA) → ピクセル(API) の標準的な変換レートにより、変換してリサイズします
'            ・WebView2はネイティブウィンドウのため、Frame/ボタン等より常に手前に描画されます
'              (いわゆる「airspace」問題)。そのため、Frameを単純に右へ伸ばすだけの実装だと、
'              最大化時にボタン一式がWebView2の裏へ完全に隠れてしまいます。ここでは右側パネルの
'              各コントロール自体をウィンドウ右端に追従させて横移動させ、Frameの幅は
'              「パネルの一番左端の手前まで」に制限することで、重なりが発生しないようにしています
'* 注意事項：画面のDPI設定によってはズレる場合があるので、微調整してください
'***************************************************************************************************
Private Sub AdjustEdgeSize()
    Const PointToPixel As Double = 1.3333
    Const PanelGap As Long = 12  'EdgeFrameと右側パネルの隙間

    '1. 右側パネル(ボタン群/StatusLabel)を、ウィンドウの拡大縮小分だけ水平移動して右端に追従させる
    Dim shiftX As Long: shiftX = Me.InsideWidth - InitialInsideWidth
    If shiftX < 0 Then shiftX = 0

    Dim i As Long
    For i = 1 To UBound(PanelControlRefs)
        PanelControlRefs(i).Left = PanelControlOriginalLeft(i) + shiftX
    Next i

    '2. パネルの一番左端の座標を求める
    Dim panelLeftMost As Long: panelLeftMost = PanelControlRefs(1).Left
    For i = 2 To UBound(PanelControlRefs)
        If PanelControlRefs(i).Left < panelLeftMost Then panelLeftMost = PanelControlRefs(i).Left
    Next i

    '3. EdgeFrameは、パネルへ重ならない範囲までしか広げない
    Dim tmp As Long
    tmp = panelLeftMost - PanelGap - Me.EdgeFrame.Left
    If tmp >= 0 Then Me.EdgeFrame.Width = tmp

    tmp = Me.InsideHeight - BottomMargin - Me.EdgeFrame.Top
    If tmp >= 0 Then Me.EdgeFrame.height = tmp

    If fWebView2 Is Nothing Then Exit Sub

    Dim pxWidth As Long, pxHeight As Long
    pxWidth = Me.EdgeFrame.InsideWidth * PointToPixel
    pxHeight = Me.EdgeFrame.InsideHeight * PointToPixel
    fWebView2.Resize pxWidth, pxHeight
End Sub

Private Sub UserForm_Resize()
    AdjustEdgeSize
End Sub



'***************************************************************************************************
'                                  ■■■ ゲーム進行ボタン ■■■
'***************************************************************************************************
'* 機能　　：初回押下時はブラウザ接続からゲーム開始まで、2回目以降は`start_game`による再スタートを行います
'***************************************************************************************************
Private Sub StartButton_Click()
    Me.StartButton.Enabled = False
    DoEvents

    If fCDPContext Is Nothing Then
        If Not ConnectToMaze() Then
            AppendStatus "ブラウザの起動に失敗しました。「ブラウザ起動設定」シートの起動引数や、Edgeのインストール状況を確認してください。"
            Me.StartButton.Enabled = True
            Exit Sub
        End If
    End If

    On Error GoTo ErrHandler
    fMCP.onExistTool "start_game", timeOutInSeconds:=20

    Dim res As BiDiCDPJson
    Set res = ToolResult(fMCP.ExecuteWebMCP("start_game"))
    If res Is Nothing Then Err.Raise vbObjectError + 1, , "start_gameの応答を解釈できませんでした。"

    GameStarted = True
    FacingDirection = "north"
    Me.UseButton.Caption = "使う (" & DirectionLabel(FacingDirection) & "向き)"
    Me.StartButton.Caption = "ゲーム再開"
    SetGameControlsEnabled True

    AppendStatus "ゲーム開始: " & res.StringKey("message")
    RefreshLook

    Me.StartButton.Enabled = True
    Exit Sub

ErrHandler:
    AppendStatus "ゲーム開始に失敗しました: " & Err.Description & vbCrLf & _
        "(WebMCPは実験的機能のため、Edgeのバージョン/フラグ設定によっては`start_game`ツールが現れません)"
    Me.StartButton.Enabled = True
End Sub

Private Sub LookButton_Click()
    RefreshLook
End Sub

Private Sub MoveNorthButton_Click()
    DoMove "north"
End Sub

Private Sub MoveSouthButton_Click()
    DoMove "south"
End Sub

Private Sub MoveEastButton_Click()
    DoMove "east"
End Sub

Private Sub MoveWestButton_Click()
    DoMove "west"
End Sub

Private Sub PickupButton_Click()
    If Not fMCP.isExistTool("pickup") Then HandleGameplayToolsGone: Exit Sub

    Dim res As BiDiCDPJson: Set res = ToolResult(fMCP.ExecuteWebMCP("pickup"))
    If res Is Nothing Then Exit Sub
    If res.BoolKey("success") Then
        AppendStatus "拾った: " & res.StringKey("message")
    Else
        AppendStatus "拾えませんでした: " & res.StringKey("reason")
    End If
    RefreshLook
End Sub

Private Sub DropButton_Click()
    If Not fMCP.isExistTool("drop") Then HandleGameplayToolsGone: Exit Sub

    Dim res As BiDiCDPJson: Set res = ToolResult(fMCP.ExecuteWebMCP("drop"))
    If res Is Nothing Then Exit Sub
    If res.BoolKey("success") Then
        AppendStatus "置いた: " & res.StringKey("message")
    Else
        AppendStatus "置けませんでした: " & res.StringKey("reason")
    End If
    RefreshLook
End Sub

Private Sub UseButton_Click()
    If Not fMCP.isExistTool("use") Then HandleGameplayToolsGone: Exit Sub

    Dim params As New Dictionary
    params.Add "direction", FacingDirection

    Dim res As BiDiCDPJson: Set res = ToolResult(fMCP.ExecuteWebMCP("use", params))
    If res Is Nothing Then Exit Sub
    If res.BoolKey("success") Then
        AppendStatus "使った(" & DirectionLabel(FacingDirection) & "向き): " & res.StringKey("message")
    Else
        AppendStatus "使えませんでした(" & DirectionLabel(FacingDirection) & "向き): " & res.StringKey("reason")
    End If
    RefreshLook
End Sub



'***************************************************************************************************
'                                  ■■■ ゲーム操作の内部処理 ■■■
'***************************************************************************************************
'* 機能　　：`ExecuteWebMCP`の戻り値から、ツール本体の戻り値(success/position/reason等)を取り出します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`WebMCP.toolResponded`の`params`は、`{invocationId, status, output:{...}}`という形で、
'            ツール自身の戻り値は`output`キーの中にネストされたオブジェクトとして入っています
'            (実機確認済み。稀に`output`がJSON文字列として来るケースにも一応保険で対応しています)
'***************************************************************************************************
Private Function ToolResult(ByVal RawResponse As BiDiCDPJson) As BiDiCDPJson
    If RawResponse Is Nothing Then Exit Function

    If RawResponse.ExistsKey("output") Then
        Dim outputNode As BiDiCDPJson: Set outputNode = RawResponse.NodeKey("output")
        If Not outputNode Is Nothing Then
            Set ToolResult = outputNode
            Exit Function
        End If

        Dim outputVal As Variant: outputVal = RawResponse.Item("output")
        If VarType(outputVal) = vbString Then
            Set ToolResult = BiDiCDPJson.Parse(CStr(outputVal))
            Exit Function
        End If
    End If

    Set ToolResult = RawResponse
End Function

'***************************************************************************************************
'* 機能　　：`move`WebMCPツールを実行し、成否をStatusLabelへ反映します
'---------------------------------------------------------------------------------------------------
'* 注意事項：ゴール到達(`atExit`)と同時に、迷路ページ側は`move`/`look`等のツール一式を丸ごと削除し
'            `start_game`だけを再登録します。この削除イベントは`move`自身の応答が返るより先に届く
'            ことがあり、そのまま`RefreshLook`(=`look`実行)すると、既に消えたツール名で
'            `exCDP_WebMCP`内部の`toolsMap`にアクセスして`Err.Raise 9`(VBA-FastDictionary)になる
'            ため、`atExit`成立時は`RefreshLook`を呼ばず`HandleGameplayToolsGone`で締める
'***************************************************************************************************
Private Sub DoMove(ByVal Direction As String)
    If Not fMCP.isExistTool("move") Then HandleGameplayToolsGone: Exit Sub

    FacingDirection = Direction
    Me.UseButton.Caption = "使う (" & DirectionLabel(Direction) & "向き)"

    Dim params As New Dictionary
    params.Add "direction", Direction

    Dim res As BiDiCDPJson: Set res = ToolResult(fMCP.ExecuteWebMCP("move", params))
    If res Is Nothing Then Exit Sub
    If res.BoolKey("success") Then
        Dim posNode As BiDiCDPJson: Set posNode = res.NodeKey("position")
        AppendStatus "移動(" & DirectionLabel(Direction) & "): 成功 -> (行" & CLng(posNode.NumberKey("row")) & _
            ", 列" & CLng(posNode.NumberKey("col")) & ")"

        If res.BoolKey("atExit") Then
            AppendStatus "★ゴールに到達しました！★"
            HandleGameplayToolsGone
            Exit Sub
        End If
    Else
        AppendStatus "移動(" & DirectionLabel(Direction) & "): 失敗 -> " & res.StringKey("reason")
    End If

    RefreshLook
End Sub

'***************************************************************************************************
'* 機能　　：`look`WebMCPツールを実行し、現在地/所持品/進める方向などをStatusLabelへ反映します
'***************************************************************************************************
Private Sub RefreshLook()
    If Not fMCP.isExistTool("look") Then HandleGameplayToolsGone: Exit Sub

    Dim res As BiDiCDPJson: Set res = ToolResult(fMCP.ExecuteWebMCP("look"))
    If res Is Nothing Then Exit Sub

    Dim posNode As BiDiCDPJson: Set posNode = res.NodeKey("position")
    Dim msg As String
    msg = "現在地: (行" & CLng(posNode.NumberKey("row")) & ", 列" & CLng(posNode.NumberKey("col")) & ")" & vbCrLf

    Dim invVal As Variant: invVal = res.Item("inventory")
    msg = msg & "所持品: " & IIf(IsNull(invVal), "なし", invVal) & vbCrLf

    Dim openDirs As BiDiCDPJson: Set openDirs = res.NodeKey("openDirections")
    Dim i As Long, openList As String
    For i = 0 To openDirs.Count - 1
        openList = openList & DirectionLabel(openDirs.StringAt(i)) & " "
    Next i
    msg = msg & "進める方向: " & IIf(LenB(openList) = 0, "なし", openList) & vbCrLf

    If res.ExistsKey("blockedDirections") Then
        Dim blocked As BiDiCDPJson: Set blocked = res.NodeKey("blockedDirections")
        Dim k As Long
        For k = 0 To blocked.Count - 1
            msg = msg & "・" & DirectionLabel(blocked.KeyAt(k)) & "側に障害物: " & blocked.ValueAt(k) & vbCrLf
        Next k
    End If

    Dim collVal As Variant: collVal = res.Item("collectibleHere")
    If Not IsNull(collVal) Then msg = msg & "足元のアイテム: " & collVal & vbCrLf

    If res.BoolKey("atExit") Then msg = msg & "★ここが出口です★" & vbCrLf

    msg = msg & "移動回数: " & CLng(res.NumberKey("moveCount"))

    AppendStatus msg
End Sub

'***************************************************************************************************
'* 機能　　：ゲームクリア等でゲームプレイ系ツール(look/move/pickup/drop/use)が丸ごと消えた際、
'            UIを「ゲーム開始待ち」の状態に戻します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：迷路ページはゴール到達時、これらのツールを削除して`start_game`だけを再登録し直す
'            ため、消えたツールへ引き続きアクセスしようとするとエラーになる。各ボタンの冒頭で
'            `isExistTool`チェックに失敗した場合、ここでUIを安全な状態に戻す
'***************************************************************************************************
Private Sub HandleGameplayToolsGone()
    GameStarted = False
    SetGameControlsEnabled False
    Me.StartButton.Caption = "① ゲーム開始"
    AppendStatus "移動/確認系ツールが利用できなくなりました(ゴール到達、またはページ遷移の可能性)。" & _
        "「ゲーム開始」で再スタートできます。"
End Sub

'***************************************************************************************************
'* 機能　　：StatusLabelへログを追記し、直近の内容が見えるように文字数で切り詰めます
'***************************************************************************************************
Private Sub AppendStatus(ByVal Text As String)
    StatusLog = Text & vbCrLf & "――――――――――――――――" & vbCrLf & StatusLog
    If Len(StatusLog) > MaxStatusChars Then StatusLog = Left$(StatusLog, MaxStatusChars)
    Me.StatusLabel.Caption = StatusLog
End Sub

'***************************************************************************************************
'* 機能　　：方向コード(north/south/east/west)を日本語表示に変換します
'***************************************************************************************************
Private Function DirectionLabel(ByVal Direction As String) As String
    Select Case LCase$(Direction)
        Case "north": DirectionLabel = "北"
        Case "south": DirectionLabel = "南"
        Case "east":  DirectionLabel = "東"
        Case "west":  DirectionLabel = "西"
        Case Else:    DirectionLabel = Direction
    End Select
End Function

'***************************************************************************************************
'* 機能　　：ゲーム開始前後で、移動/周囲確認/所持品ボタンの活性状態を切り替えます
'***************************************************************************************************
Private Sub SetGameControlsEnabled(ByVal Enabled As Boolean)
    Me.LookButton.Enabled = Enabled
    Me.MoveNorthButton.Enabled = Enabled
    Me.MoveSouthButton.Enabled = Enabled
    Me.MoveEastButton.Enabled = Enabled
    Me.MoveWestButton.Enabled = Enabled
    Me.PickupButton.Enabled = Enabled
    Me.DropButton.Enabled = Enabled
    Me.UseButton.Enabled = Enabled
End Sub



'***************************************************************************************************
'                        ■■■ このUserForm用のCDP非同期イベント処理 ■■■
'***************************************************************************************************
'* 機能　　：ページ遷移(再読み込み等)を検知し、ゲーム状態を安全にリセットします
'***************************************************************************************************
Private Sub fCDPEvent_CDPContextEvent(methodName As String, RawJson As String, sessionID As String)
    If fCDPContext Is Nothing Then Exit Sub
    If fCDPContext.CurrentSessionID <> sessionID Then Exit Sub

    Select Case methodName
        Case "Page.frameNavigated"
            If GameStarted Then
                GameStarted = False
                SetGameControlsEnabled False
                Me.StartButton.Caption = "① ゲーム開始"
                AppendStatus "ページ遷移を検知したため状態をリセットしました。再度「ゲーム開始」を押してください。"
            End If

        Case Else
            Exit Sub
    End Select
End Sub



'***************************************************************************************************
'                                 ■■■ 初期化/後始末 ■■■
'***************************************************************************************************
'* 機能　　：操作に必要なハンドル情報を取得し、初期状態を整えます
'***************************************************************************************************
Private Sub UserForm_Initialize()
    '1. このUserForm自体のウィンドウハンドルを取得する
    ' （"ThunderDFrame" はExcel UserFormのクラス名です）
    myFormHwnd = FindWindow("ThunderDFrame", Me.Caption)

    '2. 現在のスタイルを取得し、このUserFormにリサイズ機能を追加
    Dim currentStyle As LongPtr
    currentStyle = GetWindowLongPtr(myFormHwnd, GWL_STYLE)
    SetWindowLongPtr myFormHwnd, GWL_STYLE, currentStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX

    '3. 埋め込み先のEdgeフレームのハンドル情報を取得
    myEdgeFrameHwnd = Me.EdgeFrame.[_GethWnd]

    '4. フレームの下マージン計算
    BottomMargin = Me.InsideHeight - Me.EdgeFrame.height - Me.EdgeFrame.Top

    '4-2. 右側パネル追従用に、現在の横幅と各コントロールの初期Leftを記録しておく
    InitialInsideWidth = Me.InsideWidth
    Set PanelControlRefs(1) = Me.StartButton:      PanelControlOriginalLeft(1) = Me.StartButton.Left
    Set PanelControlRefs(2) = Me.LookButton:       PanelControlOriginalLeft(2) = Me.LookButton.Left
    Set PanelControlRefs(3) = Me.MoveNorthButton:  PanelControlOriginalLeft(3) = Me.MoveNorthButton.Left
    Set PanelControlRefs(4) = Me.MoveWestButton:   PanelControlOriginalLeft(4) = Me.MoveWestButton.Left
    Set PanelControlRefs(5) = Me.MoveEastButton:   PanelControlOriginalLeft(5) = Me.MoveEastButton.Left
    Set PanelControlRefs(6) = Me.MoveSouthButton:  PanelControlOriginalLeft(6) = Me.MoveSouthButton.Left
    Set PanelControlRefs(7) = Me.PickupButton:     PanelControlOriginalLeft(7) = Me.PickupButton.Left
    Set PanelControlRefs(8) = Me.DropButton:       PanelControlOriginalLeft(8) = Me.DropButton.Left
    Set PanelControlRefs(9) = Me.UseButton:        PanelControlOriginalLeft(9) = Me.UseButton.Left
    Set PanelControlRefs(10) = Me.StatusLabel:     PanelControlOriginalLeft(10) = Me.StatusLabel.Left

    '5. WebView2のコアオブジェクトを初期化
    Set fWebView2 = New CDPCoreViaWebView2

    '6. ゲーム状態の初期化
    FacingDirection = "north"
    SetGameControlsEnabled False
End Sub

Private Sub UserForm_Terminate()
    Set fMCP = Nothing
    If Not fWebView2 Is Nothing Then fWebView2.DisconnectCDP
    Set fWebView2 = Nothing
    Set fCDPEvent = Nothing
    Set fCDPContext = Nothing
End Sub
