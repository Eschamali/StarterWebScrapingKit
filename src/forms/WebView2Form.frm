VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WebView2Form 
   Caption         =   "WebView2Demo"
   ClientHeight    =   8295.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15000
   OleObjectBlob   =   "WebView2Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "WebView2Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'                         ユーザーフォームに本物のWebView2埋め込みます
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 必要なWindowsAPI定義 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long 'WebView2用追加起動引数設定用



'***************************************************************************************************
'                               ■■■ 制御に必要な変数定義 ■■■
'***************************************************************************************************
'制御のコアとなるオブジェクト
Private fWebView2           As CDPCoreViaWebView2
Private WithEvents CDPEvent As CDPCore      '非同期イベント処理用
Attribute CDPEvent.VB_VarHelpID = -1
Private fCDPContext         As CDPContext   'タブ情報
Attribute fCDPContext.VB_VarHelpID = -1

'自身の各ハンドルを保存する変数
Private myFormHwnd      As LongPtr
Private myEdgeFrameHwnd As LongPtr

'Frameのマージン
Private RightMargin As Long
Private BottomMargin As Long



'***************************************************************************************************
'                              ■■■ ウィンドウスタイルの定数 ■■■
'***************************************************************************************************
Private Const GWL_STYLE         As Long = -16
Private Const WS_THICKFRAME     As Long = &H40000 'サイズ変更枠
Private Const WS_MAXIMIZEBOX    As Long = &H10000 '最大化ボタン
Private Const WS_MINIMIZEBOX    As Long = &H20000 '最小化ボタン



'***************************************************************************************************
'                                   ■■■ 新規起動 ■■■
'***************************************************************************************************
Friend Function StartCDPModeWebView2(Optional SwitchUser As String) As CDPContext
    '1. WebView2の追加起動引数準備
    Const EnvironmentName As String = "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"
    SetEnvironmentVariableW StrPtr(EnvironmentName), StrPtr(ShSetting01_StartBrowser.UseRangeID(3, ".StartCDPModeWebView2"))

    '2. 引数が省略されてる場合は、ワークシートの設定を適用
    If StrPtr(SwitchUser) = 0 Then SwitchUser = ShSetting01_StartBrowser.CurrentUserName

    '3. WebView2を起動
    Dim isActive As Boolean
    Set fWebView2 = New CDPCoreViaWebView2
    isActive = fWebView2.ConnectCDP(SwitchUser, myEdgeFrameHwnd)

    '4. クリア
    SetEnvironmentVariableW StrPtr(EnvironmentName), 0

    '5. 起動失敗したら、抜ける
    If Not isActive Then Set fWebView2 = Nothing: Exit Function

    '6. サイズをセット
    AdjustEdgeSize

    '7. 可視化
    SwitchVisible.value = True
    fWebView2.Visible = True

    '8. タブ接続まで行う
    Dim t As New CDPBrowser: t.reattachWebView2 SwitchUser, fWebView2
    Set fCDPContext = t.getTab(setMain:=True)

    '9. 非同期イベント処理に備える
    Set CDPEvent = t.InheritanceCDPCore

    '10. 返却
    Set StartCDPModeWebView2 = fCDPContext
End Function



'***************************************************************************************************
'                                       ■■■ サイズ変更 ■■■
'***************************************************************************************************
'* 機能　　：WebView2のサイズをFrame内ににピッタリはめ込む処理をします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：係数 1.333 は ポイント(VBA) → ピクセル(API) の標準的な変換レートにより、変換してリサイズします
'* 注意事項：画面のDPI設定によってはズレる場合があるので、微調整してください
'***************************************************************************************************
Private Sub AdjustEdgeSize()
    ' 【設定】 堀（外周の余白）のサイズをポイント単位で指定します
    Const PointToPixel  As Double = 1.3333

    ' Frameの幅と高さを、UserFormの内部サイズから余白を引いた値にする
    Dim tmp As Long

    tmp = Me.InsideWidth - RightMargin - Me.EdgeFrame.Left
    If tmp >= 0 Then Me.EdgeFrame.Width = tmp

    tmp = Me.InsideHeight - BottomMargin - Me.EdgeFrame.Top
    If tmp >= 0 Then Me.EdgeFrame.height = tmp


    ' --- 第2段階：APIの世界（EdgeをFrameに追従させる） ---
    ' Frameのサイズ（ポイント）を、API用のピクセルに変換（係数1.333）
    ' ※DPI設定によっては 1.333 以外（例：1.25 等）になる場合がありますが、標準はこれです。
    Dim pxWidth As Long
    Dim pxHeight As Long
    pxWidth = Me.EdgeFrame.InsideWidth * PointToPixel
    pxHeight = Me.EdgeFrame.InsideHeight * PointToPixel

    ' APIを使って、EdgeのウィンドウをFrameの左上(0,0)にピッタリはめ込む！
    ' (Frameの中にSetParentされているので、0,0はFrameの左上を意味します)
    fWebView2.Resize pxWidth, pxHeight
End Sub

Private Sub UserForm_Resize()
    AdjustEdgeSize
End Sub



'***************************************************************************************************
'                                       ■■■ 機能面 ■■■
'***************************************************************************************************
'* 機能　　：ボタン押下時、テキストボックスに入力したURLにページ遷移します
'***************************************************************************************************
Private Sub navigateButton_Click()
    fCDPContext.navigate Me.TextURLBox.Text
End Sub

'***************************************************************************************************
'* 機能　　：ブラウザを描画するかの切り替えが発生します
'***************************************************************************************************
Private Sub SwitchVisible_Click()
    fWebView2.Visible = SwitchVisible.value
End Sub

'***************************************************************************************************
'* 機能　　：WebView2のコアプロパティを提供します
'***************************************************************************************************
Property Get controlWebView2() As CDPCoreViaWebView2
    controlWebView2 = fWebView2
End Property



'***************************************************************************************************
'                                       ■■■ 初期化 ■■■
'***************************************************************************************************
'* 機能　　：操作に必要なハンドル情報を取得します
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

    '4. フレームの右下マージン計算
    RightMargin = Me.InsideWidth - Me.EdgeFrame.Width - Me.EdgeFrame.Left
    BottomMargin = Me.InsideHeight - Me.EdgeFrame.height - Me.EdgeFrame.Top
End Sub
