VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WebView2InExcelForm 
   Caption         =   "WebView2InExcelForm"
   ClientHeight    =   8295.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15000
   OleObjectBlob   =   "WebView2InExcelForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "WebView2InExcelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'            別スクリプトで起動したWebView2をユーザーフォームに埋め込んで
'           名前付きパイプを駆使した3者間でのやり取りによるCDP通信を実現させます
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 必要なWindowsAPI定義 ■■■
'***************************************************************************************************
' --- ウィンドウ操作用 WinAPI 宣言 ---
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

' --- 名前付きパイプ用 WinAPI 宣言 ---
Private Declare PtrSafe Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, ByVal lpSecurityAttributes As LongPtr) As LongPtr
Private Declare PtrSafe Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long



'***************************************************************************************************
'                               ■■■ 制御に必要な変数定義 ■■■
'***************************************************************************************************
'外部から渡されるWebView2のウィンドウハンドル
Private WebView2Hwnd            As LongPtr
Private Const WebView2FormTitle As String = "ExcelWebView2_Host"
Private WithEvents targetCDP    As CDPBrowser
Attribute targetCDP.VB_VarHelpID = -1

'自身のUserFormのハンドルを保存する変数
Private myFormHwnd          As LongPtr
Private myWebView2FrameHwnd As LongPtr

'Frameのマージン
Private RightMargin As Long
Private BottomMargin As Long

'パイプ操作関連の定数
Private Const PIPE_ACCESS_DUPLEX As Long = &H3 ' 送受信可能
Private Const PIPE_TYPE_BYTE As Long = &H0     ' バイトモード
Private Const PIPE_WAIT As Long = &H0          ' ブロッキングモード（手動実行の要！）
Private Const INVALID_HANDLE_VALUE As LongPtr = -1

' 共有のパイプ名
Private Const PIPE_NAME As String = "\\.\pipe\LOCAL\mojo.ExcelWebView2Pipe"
Private hPipe As LongPtr



'***************************************************************************************************
'                              ■■■ ウィンドウスタイルの定数 ■■■
'***************************************************************************************************
Private Const GWL_STYLE     As Long = -16
Private Const WS_THICKFRAME As Long = &H40000       ' サイズ変更枠
Private Const WS_CHILD      As Long = &H40000000    ' 子ウィンドウ
Private Const WS_VISIBLE    As Long = &H10000000    ' headlessモードでも強制表示させるやつ



'***************************************************************************************************
'                               ■■■ メインプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：WebView2を誘拐してUserFormに埋め込むメソッドです
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功可否論理値
'* 引数　　：TargetCDPBrowser   CDPモードでStartした後のオブジェクト変数
'***************************************************************************************************
Public Function AttachWebView2(TargetCDPBrowser As CDPBrowser) As Boolean
    '1. 必要な変数を適用
    Set targetCDP = TargetCDPBrowser
    WebView2Hwnd = FindWindow(vbNullString, WebView2FormTitle)

    '受け取り失敗時は抜ける
    If WebView2Hwnd = 0 Then Exit Function

    '2. Edgeをユーザーフォームに埋め込むための準備
    SetWindowLongPtr WebView2Hwnd, GWL_STYLE, WS_CHILD Or WS_VISIBLE

    '3. EdgeをこのUserForm内のFrame内に誘拐（SetParent）する！
    SetParent WebView2Hwnd, myWebView2FrameHwnd

    '4. Frame内サイズを合わせる
    AdjustWebView2Size

    '5. 成功で送る
    AttachWebView2 = True
End Function

'***************************************************************************************************
'* 機能　　：WebView2のサイズをFrame内ににピッタリはめ込む処理をします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：係数 1.333 は ポイント(VBA) → ピクセル(API) の標準的な変換レートにより、変換してリサイズします
'* 注意事項：画面のDPI設定によってはズレる場合があるので、微調整してください
'***************************************************************************************************
Private Sub AdjustWebView2Size()
    ' 【設定】 堀（外周の余白）のサイズをポイント単位で指定します
    Const MARGIN_PT     As Long = 40 ' 上下左右に15ポイントの余白を作る
    Const PointToPixel  As Double = 1.3333

    ' Frameの幅と高さを、UserFormの内部サイズから余白を引いた値にする
    Dim tmp As Long

    tmp = Me.InsideWidth - RightMargin - Me.WebView2Frame.Left
    If tmp >= 0 Then Me.WebView2Frame.Width = tmp

    tmp = Me.InsideHeight - BottomMargin - Me.WebView2Frame.Top
    If tmp >= 0 Then Me.WebView2Frame.height = tmp


    ' --- 第2段階：APIの世界（EdgeをFrameに追従させる） ---
    ' Frameのサイズ（ポイント）を、API用のピクセルに変換（係数1.333）
    ' ※DPI設定によっては 1.333 以外（例：1.25 等）になる場合がありますが、標準はこれです。
    Dim pxWidth As Long
    Dim pxHeight As Long
    pxWidth = Me.WebView2Frame.InsideWidth * PointToPixel
    pxHeight = Me.WebView2Frame.InsideHeight * PointToPixel

    ' APIを使って、EdgeのウィンドウをFrameの左上(0,0)にピッタリはめ込む！
    ' (Frameの中にSetParentされているので、0,0はFrameの左上を意味します)
    MoveWindow WebView2Hwnd, 0, 0, pxWidth, pxHeight, 1
End Sub

Public Sub AdjustEdgeSizeMemo()
    ' 係数 1.333 は ポイント(VBA) → ピクセル(API) の標準的な変換レートです
    ' 画面のDPI設定によってはズレる場合があるので、微調整してください
    Dim pxWidth As Long
    Dim pxHeight As Long
    
    pxWidth = Me.InsideWidth * 1.333
    pxHeight = Me.InsideHeight * 1.333

    ' Edgeを左上(0,0)に配置し、フォームの大きさに合わせる
    MoveWindow WebView2Hwnd, -7, -31, pxWidth + 14, pxHeight + 31 + 7, 1
End Sub



'***************************************************************************************************
'                                       ■■■ 機能面 ■■■
'***************************************************************************************************
'* 機能　　：ボタン押下時、テキストボックスに入力したURLにページ遷移します
'***************************************************************************************************
Private Sub CommandButton1_Click()
    targetCDP.navigate Me.TextURLBox.Text
End Sub

Public Function StartCreateNamedPipe() As Boolean
    ' パイプを作成
    Dim tmp As LongPtr
    tmp = CreateNamedPipe(PIPE_NAME, PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE Or PIPE_WAIT, 1, 1024, 2 ^ 20, 0, 0)
    If tmp = INVALID_HANDLE_VALUE Then
        MsgBox "パイプの作成に失敗しました??", vbCritical
        Exit Function
    End If

    Debug.Print "パイプを開設しました。WebView2からの接続を待っています..."
    Debug.Print "パイプハンドル：" & tmp

    ' ★注意★ ここでExcelはWebView2が繋いでくるまで「フリーズ（待機状態）」になります！
    hPipe = tmp
    ConnectNamedPipe tmp, 0

    ' WebView2が繋ぐとフリーズが解けてここに進む
    Debug.Print "WebView2が接続してきました！"
    StartCreateNamedPipe = True
End Function

Public Sub QuitCloseNamedPipe()
    DisconnectNamedPipe hPipe
    CloseHandle hPipe
End Sub



'***************************************************************************************************
'* 機能　　：UserFormのサイズが変更されたら、中のWebView2のサイズも追従させます
'***************************************************************************************************
Private Sub UserForm_Resize()
    Call AdjustWebView2Size
End Sub



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
    SetWindowLongPtr myFormHwnd, GWL_STYLE, currentStyle Or WS_THICKFRAME

    '3. 埋め込み先のWebView2フレームのハンドル情報を取得
    myWebView2FrameHwnd = Me.WebView2Frame.[_GethWnd]

    '4. フレームの右下マージン計算
    RightMargin = Me.InsideWidth - Me.WebView2Frame.Width - Me.WebView2Frame.Left
    BottomMargin = Me.InsideHeight - Me.WebView2Frame.height - Me.WebView2Frame.Top
End Sub
