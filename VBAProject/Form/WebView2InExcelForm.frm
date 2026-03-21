VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WebView2InExcelForm 
   Caption         =   "EdgeInUserForm"
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
Option Explicit
'***
' WebView2Frame.frm
' WebView2 を UserForm の Frame（Win32 HWND）に埋め込むコンテナ。
'
' 構造:
'   ┌─────────────────────────────────────────┐
'   │ [txtUrl ____________________] [btnGo]   │
'   ├─────────────────────────────────────────┤
'   │                                         │
'   │       wv2Container (Frame)              │
'   │       ← _GethWnd で hWnd 取得 →        │
'   │                                         │
'   └─────────────────────────────────────────┘
'
' インポート方法:
'   VBE メニュー → ファイル → ファイルのインポート → WebView2Frame.frm を選択
'
' 実行方法:
'   Demo_WebView2.bas の TestWebView2Form() を実行
'***

' Win32 API
#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowW" ( _
    ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

Private Const GWL_STYLE     As Long = -16
Private Const WS_THICKFRAME As Long = &H40000   ' ユーザーリサイズ許可
' SetWindowPos フラグ
Private Const SWP_NOSIZE       As Long = &H1
Private Const SWP_NOMOVE       As Long = &H2
Private Const SWP_NOZORDER     As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20   ' ノンクライアント領域を強制再描画

' VBAポイント → ピクセル変換係数（96DPI標準）
Private Const PT2PX As Double = 1.3333

' WebView2Core インスタンス（WithEvents でイベントを受け取る）
Private WithEvents m_wv2 As WebView2Core
Attribute m_wv2.VB_VarHelpID = -1

' 状態フラグ
Private m_Ready       As Boolean
Private m_Initialized As Boolean  ' StartWebView2 の二度呼び防止
Private m_InResize    As Boolean  ' UserForm_Resize の再入防止

'Frameのマージン
Private fRightMargin As Long
Private fBottomMargin As Long

' 初期 URL / 追加引数（外部から WebView2Core を渡された場合に使用）
Private m_InitialUrl    As String
Private m_AdditionalArg As String
Private m_UserDataName  As String


' =====================================================================
' UserForm ライフサイクル
' =====================================================================

Private Sub UserForm_Initialize()
    'フレームの右下マージン計算
    fRightMargin = Me.InsideWidth - Me.wv2Container.Width - Me.wv2Container.Left
    fBottomMargin = Me.InsideHeight - Me.wv2Container.height - Me.wv2Container.Top

    ' WebView2Core インスタンス生成（hWnd取得は Activate で行う）
    ' 既に外部から AttachCore で渡されている場合は、新しく作らない
    If m_wv2 Is Nothing Then
        Set m_wv2 = New WebView2Core
    End If
End Sub

Private Sub UserForm_Activate()
    ' フォームが画面に出てから初期化する（コントロールが描画済みになる）
    If m_Initialized Then Exit Sub
    m_Initialized = True

    ' ① WS_THICKFRAME を付与してリサイズ可能にする
    ApplyThickFrame

    ' ② WebView2 初期化（ループ内で Ready まで待機する）
    StartWebView2

    ' ③ WebView2 初期化完了後に WS_THICKFRAME を再適用する
    '    （DoEvents ループ中に Excel がスタイルをリセットするため、完了後に再適用）
    '    StartWebView2 が返った後なので再入なしで安全
    ApplyThickFrame
End Sub

Private Sub UserForm_Resize()
    ' 再入防止（put_Bounds が WM_SIZE を発火させてもループしない）
    If m_InResize Then Exit Sub
    m_InResize = True
    On Error GoTo ResizeCleanup

    Const TOOLBAR_H As Single = 21  ' ツールバー行の高さ（ポイント）
    Const BTN_W     As Single = 48
    Const MARGIN    As Single = 50

    ' --- wv2Container（Frame）のリサイズ追従 ---
    Dim tmp As Long

    tmp = Me.InsideWidth - fRightMargin - Me.wv2Container.Left
    If tmp >= 0 Then Me.wv2Container.Width = tmp

    tmp = Me.InsideHeight - fBottomMargin - Me.wv2Container.Top
    If tmp >= 0 Then Me.wv2Container.height = tmp

    ' --- WebView2 リサイズ（ハイブリッド方式：Win32 API で即リサイズ ＋ タイマーで COM 同期） ---
    '     EdgeInExcelForm と同様の方式。ResizeDirect で Win32 レベルの追従性を確保し、
    '     ScheduleResize (Timer) で COM 側の内部状態を安全に更新する。
    If m_Ready And Not m_wv2 Is Nothing Then
        Dim hFrame As LongPtr
        hFrame = Me.wv2Container.[_GethWnd]
        
        ' 1. Win32 API (MoveWindow) で子ウィンドウを即座にリサイズ（爆速追従）
        Dim pxW As Long, pxH As Long
        pxW = CLng(Me.wv2Container.InsideWidth * PT2PX)
        pxH = CLng(Me.wv2Container.InsideHeight * PT2PX)
        m_wv2.ResizeDirect pxW, pxH
        
        ' 2. タイマー経由で COM (put_Bounds) を実行（内部状態の確定）
        m_wv2.ScheduleResize hFrame
    End If

ResizeCleanup:
    m_InResize = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not m_wv2 Is Nothing Then
        m_wv2.quit
        Set m_wv2 = Nothing
    End If
    m_Ready = False
End Sub

' =====================================================================
' コントロール イベント
' =====================================================================

Private Sub btnGo_Click()
    If Not m_Ready Then
        MsgBox "WebView2 がまだ準備中です。しばらくお待ちください。", vbInformation
        Exit Sub
    End If
    Dim Url As String: Url = Trim(Me.txtUrl.Text)
    If Len(Url) = 0 Then Exit Sub
    ' プロトコルが無ければ https:// を補完
    If InStr(Url, "://") = 0 Then Url = "https://" & Url
    Me.txtUrl.Text = Url
    m_wv2.navigate Url
End Sub

Private Sub txtUrl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Enter キーでもナビゲート
    If KeyCode = 13 Then btnGo_Click
End Sub

' =====================================================================
' WebView2Core イベント
' =====================================================================

Private Sub m_wv2_Ready()
    m_Ready = True
    Me.Caption = "WebView2 Browser - 準備完了"
    If WV2logView Then Debug.Print "[WebView2Frame] m_wv2_Ready: WebView2 の初期化が完了しました"
    ' ? ここは ProcessMessages ループの中なので ApplyThickFrame を呼び出さない。
    '    WS_THICKFRAME 再適用は UserForm_Activate の StartWebView2 完了後で行う。
End Sub

Private Sub m_wv2_NavigationCompleted(ByVal isSuccess As Boolean, ByVal webErrorStatus As Long)
    If isSuccess Then
        Me.Caption = "WebView2 Browser - 完了"
    Else
        Me.Caption = "WebView2 Browser - ナビゲーションエラー (status=" & webErrorStatus & ")"
    End If
End Sub

Private Sub m_wv2_InitializeFailed(ByVal ErrorCode As Long, ByVal Description As String)
    MsgBox "WebView2 初期化失敗:" & vbCrLf & Description & vbCrLf & _
           "(HRESULT: 0x" & Hex(ErrorCode) & ")", vbCritical, "WebView2Frame"
    Unload Me
End Sub

' =====================================================================
' Private ヘルパー
' =====================================================================

'----------------------------------------------------------------------
' StartWebView2
'   wv2Container Frame の hWnd を親として WebView2 を初期化する。
'   ★ ここが核心：Application.hWnd ではなく Frame の hWnd を渡す ★
'----------------------------------------------------------------------
Private Sub StartWebView2()
    ' Frame コントロールの Win32 HWND を取得
    Dim hFrame As LongPtr
    hFrame = Me.wv2Container.[_GethWnd]

    If hFrame = 0 Then
        MsgBox "wv2Container の hWnd が取得できませんでした。" & vbCrLf & _
               "フォームが表示されてから呼び出してください。", vbCritical
        Exit Sub
    End If

    ' Frame のピクセルサイズを計算
    Dim pxW As Long, pxH As Long
    pxW = Me.wv2Container.InsideWidth * PT2PX
    pxH = Me.wv2Container.InsideHeight * PT2PX

    If WV2logView Then Debug.Print "[WebView2Frame] StartWebView2"
    If WV2logView Then Debug.Print "  hFrame = 0x" & Hex(hFrame)
    If WV2logView Then Debug.Print "  Size   = " & pxW & " x " & pxH & " px"

    Me.Caption = "WebView2 Browser - 初期化中..."

    ' WebView2Core を初期化
    '   ★ 親hWnd = hFrame（UserForm上のFrameコントロール）★
    Dim ok As Boolean
    Dim initUrl As String
    Dim addArgs As String

    ' 外部から指定された URL / 引数があればそれを使い、無ければ従来の既定値を使う
    If Len(m_InitialUrl) > 0 Then initUrl = m_InitialUrl

    addArgs = m_AdditionalArg

    ok = m_wv2.Initialize(hFrame, 0, 0, pxW, pxH, initUrl, m_UserDataName, addArgs)

    If Not ok Then
        ' エラーは InitializeFailed で既に表示済み。Unload 後は m_wv2 が無効になるため参照しない
        Exit Sub
    End If

    ' Ready になるまでメッセージポンプを回す
    Dim t As Single: t = Timer
    Do While Not m_wv2.IsReady And Timer - t < 20
        m_wv2.ProcessMessages
        DoEvents
    Loop

    If Not m_wv2.IsReady Then
        On Error Resume Next
        Dim errCode As Long, errDesc As String
        errCode = m_wv2.LastErrorCode: errDesc = m_wv2.LastErrorDescription
        On Error GoTo 0
        MsgBox "初期化タイムアウト" & vbCrLf & _
               "LastError: 0x" & Hex(errCode) & vbCrLf & _
               errDesc, vbCritical
    End If
End Sub

'----------------------------------------------------------------------
' AttachCore
'   外部（例: WebView2Browser クラス）で生成済みの WebView2Core を注入する。
'   必要に応じて初期 URL / 追加引数も一緒に渡す。
'----------------------------------------------------------------------
Public Sub AttachCore(ByVal core As WebView2Core, Optional ByVal initialUrl As String, Optional ByVal additionalArgs As String, Optional ByVal UserDataName As String)
    Set m_wv2 = core
    m_InitialUrl = initialUrl
    m_AdditionalArg = additionalArgs
    m_UserDataName = UserDataName
End Sub

'----------------------------------------------------------------------
' FindFormHwnd
'   UserForm の Win32 hWnd を ThunderDFrame クラス名で検索する。
'----------------------------------------------------------------------
Private Function FindFormHwnd(ByVal Caption As String) As LongPtr
    FindFormHwnd = FindWindow(StrPtr("ThunderDFrame"), StrPtr(Caption))
End Function

'----------------------------------------------------------------------
' ApplyThickFrame
'   WS_THICKFRAME を付与してリサイズ可能にする。
'   ★ SetWindowPos(SWP_FRAMECHANGED) は使わない ★
'     SWP_FRAMECHANGED → WM_NCCALCSIZE → WM_SIZE → UserForm_Resize → m_wv2.Resize
'     の連鎖で WebView2 が枠外に飛ぶ原因になる。
'     WS_THICKFRAME ビットだけで WM_NCHITTEST が正しく返るため、
'     スタイル設定だけでドラッグリサイズは機能する。（EdgeInExcelForm.frm と同じ方法）
'----------------------------------------------------------------------
Private Sub ApplyThickFrame()
    Dim hForm As LongPtr
    hForm = FindFormHwnd(Me.Caption)
    If hForm = 0 Then Exit Sub

    Dim sty As LongPtr
    sty = GetWindowLongPtr(hForm, GWL_STYLE)
    SetWindowLongPtr hForm, GWL_STYLE, sty Or WS_THICKFRAME
    ' ← SetWindowPos / SWP_FRAMECHANGED は呼ばない
End Sub
