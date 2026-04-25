VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EdgeInExcelForm 
   Caption         =   "EdgeInUserForm"
   ClientHeight    =   8295.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15000
   OleObjectBlob   =   "EdgeInExcelForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "EdgeInExcelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'           Edgeブラウザをユーザーフォームに埋め込んで、WebView2っぽい雰囲気にします
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 必要なWindowsAPI定義 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetActiveWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr



'***************************************************************************************************
'                               ■■■ 制御に必要な変数定義 ■■■
'***************************************************************************************************
'外部から渡されるEdgeのウィンドウハンドル
Private EdgeHwnd                As LongPtr
Private WithEvents targetCDP    As CDPBrowser
Attribute targetCDP.VB_VarHelpID = -1

'自身のUserFormのハンドルを保存する変数
Private myFormHwnd      As LongPtr
Private myEdgeFrameHwnd As LongPtr

'Frameのマージン
Private RightMargin As Long
Private BottomMargin As Long



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
'* 機能　　：Edgeを誘拐してUserFormに埋め込むメソッドです
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功可否論理値
'* 引数　　：TargetCDPBrowser   CDPモードでStartした後のオブジェクト変数
'***************************************************************************************************
Public Function AttachEdge(TargetCDPBrowser As CDPBrowser) As Boolean
    '1. 必要な変数を適用
    Set targetCDP = TargetCDPBrowser
    EdgeHwnd = targetCDP.BrowserWindowHandle(True)

    '受け取り失敗時は抜ける
    If EdgeHwnd = 0 Then Exit Function

    '2. Edgeをユーザーフォームに埋め込むための準備
    SetWindowLongPtr EdgeHwnd, GWL_STYLE, WS_CHILD Or WS_VISIBLE

    '3. EdgeをこのUserForm内のFrame内に誘拐（SetParent）する！
    SetParent EdgeHwnd, myEdgeFrameHwnd

    '4. Frame内サイズを合わせる
    AdjustEdgeSize

    '5. 成功で送る
    AttachEdge = True
End Function

'***************************************************************************************************
'* 機能　　：EdgeのサイズをFrame内ににピッタリはめ込む処理をします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：係数 1.333 は ポイント(VBA) → ピクセル(API) の標準的な変換レートにより、変換してリサイズします
'* 注意事項：画面のDPI設定によってはズレる場合があるので、微調整してください
'***************************************************************************************************
Private Sub AdjustEdgeSize()
    ' 【設定】 堀（外周の余白）のサイズをポイント単位で指定します
    Const MARGIN_PT     As Long = 40 ' 上下左右に15ポイントの余白を作る
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
    MoveWindow EdgeHwnd, 0, 0, pxWidth, pxHeight, 1
End Sub



'***************************************************************************************************
'                                       ■■■ 機能面 ■■■
'***************************************************************************************************
'* 機能　　：ボタン押下時、テキストボックスに入力したURLにページ遷移します
'***************************************************************************************************
Private Sub navigateButton_Click()
    targetCDP.navigate Me.TextURLBox.Text
End Sub



'***************************************************************************************************
'                        ■■■ Edgeにフォーカスするためのイベント関連 ■■■
'***************************************************************************************************
'* 機能　　：UserformがShowされたのと同時に、Edgeにフォーカスします
'---------------------------------------------------------------------------------------------------
'* 注意事項：・最初の`Show`しか効果ありません
'            ・Userform同士のウィンドウアクティブ/非アクティブしか効果ありません
'***************************************************************************************************
Private Sub UserForm_Activate()
    Call attachToEdgeFocus
End Sub

'***************************************************************************************************
'* 機能　　：Frameの外周にあたるUserformに、クリックすると、Edgeにフォーカスがあたります
'---------------------------------------------------------------------------------------------------
'* 注意事項：ウィンドウのタイトルバークリックでは反応しません
'***************************************************************************************************
Private Sub UserForm_Click()
    Call attachToEdgeFocus
End Sub

'***************************************************************************************************
'* 機能　　：Frameの外周にあたるUserformに、マウスが触れると、Edgeにフォーカスがあたります
'---------------------------------------------------------------------------------------------------
'* 詳細説明：前者は、Frame枠外周全域、後者は特定のラベルエリアのみとなります
'* 注意事項：・利用ユーザーが「電流イライラ棒裏技wazappu」のプロフェッショナルで、光速で堀(UserForm領域イベント)をすり抜けて城(Edge)をクリックした場合、フォーカス移動検知が間に合わず「文字が入力できない」という現象が発生することを懸念してください
'            ・イベント検知領域が、別ウィンドウと重なって、直でEdge領域に行っても失敗します
'***************************************************************************************************
'Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'    Call attachToEdgeFocus
'End Sub

Private Sub Notice_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call attachToEdgeFocus
    Call FocusNotify        'ついでに通知もしておく
End Sub

'***************************************************************************************************
'* 機能　　：UserFormのサイズが変更されたら、中のEdgeのサイズも追従させ、ついでに、フォーカスもしておきます
'***************************************************************************************************
Private Sub UserForm_Resize()
    Call attachToEdgeFocus
    Call AdjustEdgeSize
End Sub

'***************************************************************************************************
'* 機能　　：Userform内にあるEdgeにフォーカスを当て、キーボード入力ができるようにします
'---------------------------------------------------------------------------------------------------
'* 詳細説明：埋め込まれたEdgeは、アクティブ化反映ができない仕様により、プロシージャ呼び出しによるフォーカスでなんとか実現しました
'* 注意事項：このプロシージャは、Userformのイベントで実行するように仕向けてください。
'***************************************************************************************************
Private Sub attachToEdgeFocus()
    '別プロセスウィンドウがアクティブになっても、ExcelUserformをアクティブ化させる処理
    bringToForeground myFormHwnd

    'Edgeハンドルにフォーカスさせる
    FocusEdge EdgeHwnd
End Sub

'***************************************************************************************************
'* 機能　　：Userform内にあるEdgeにフォーカスを当てる処理をします
'***************************************************************************************************
Private Sub FocusEdge(ByVal EdgeHwnd As LongPtr)
    Dim pid As Long
    Dim tidTarget As Long, tidMe As Long

    tidMe = GetCurrentThreadId()
    tidTarget = GetWindowThreadProcessId(EdgeHwnd, pid)

    ' SetFocus は「同じ入力キューにアタッチされている必要」[2](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setfocus)[3](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-attachthreadinput)
    If tidTarget <> tidMe Then AttachThreadInput tidMe, tidTarget, 1
    SetFocus EdgeHwnd
    If tidTarget <> tidMe Then AttachThreadInput tidMe, tidTarget, 0
End Sub

'***************************************************************************************************
'* 機能　　：他ウィンドウがアクティブ状態であっても、UserFormにアクティブ化させる処理です
'***************************************************************************************************
Private Sub bringToForeground(ByVal hWndTop As LongPtr)
    Dim fg As LongPtr, pid As Long
    Dim tidFG As Long, tidMe As Long

    fg = GetForegroundWindow()
    tidFG = GetWindowThreadProcessId(fg, pid)
    tidMe = GetCurrentThreadId()

    ' いまのフォアグラウンドスレッドと入力キュー共有（成功すると前面化が通りやすい）
    If tidFG <> tidMe Then AttachThreadInput tidMe, tidFG, 1

    ' SetForegroundWindow は制限がある（失敗することがあるのは仕様）[1](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow)
    SetForegroundWindow (hWndTop)
    SetActiveWindow hWndTop  ' アクティブ化（前面でないと効かないことがある）

    If tidFG <> tidMe Then AttachThreadInput tidMe, tidFG, 0
End Sub

'***************************************************************************************************
'* 機能　　：フォーカスした旨のステータスを表示させます
'***************************************************************************************************
Private Sub FocusNotify()
    Me.focusNotice.Visible = True
    DoEvents

    Dim endTime As Single
    endTime = Timer + 1
    Do While Timer < endTime
        DoEvents ' これを入れないとExcelがフリーズしてイベントが拾えない！
    Loop

    Me.focusNotice.Visible = False
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

    '3. 埋め込み先のEdgeフレームのハンドル情報を取得
    myEdgeFrameHwnd = Me.EdgeFrame.[_GethWnd]

    '4. フレームの右下マージン計算
    RightMargin = Me.InsideWidth - Me.EdgeFrame.Width - Me.EdgeFrame.Left
    BottomMargin = Me.InsideHeight - Me.EdgeFrame.height - Me.EdgeFrame.Top
End Sub
