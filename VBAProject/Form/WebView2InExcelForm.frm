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
Private Const WS_THICKFRAME As Long = &H40000  ' ユーザーリサイズ許可

' VBAポイント → ピクセル変換係数（96DPI標準）
Private Const PT2PX As Double = 1.3333

' WebView2Core インスタンス（WithEvents でイベントを受け取る）
Private WithEvents m_wv2 As WebView2Core
Attribute m_wv2.VB_VarHelpID = -1

' 状態フラグ
Private m_Ready       As Boolean
Private m_Initialized As Boolean  ' StartWebView2 の二度呼び防止

' =====================================================================
' UserForm ライフサイクル
' =====================================================================

Private Sub UserForm_Initialize()
    ' WebView2Core インスタンス生成（hWnd取得は Activate で行う）
    Set m_wv2 = New WebView2Core
End Sub

Private Sub UserForm_Activate()
    ' フォームが画面に出てから初期化する（コントロールが描画済みになる）
    If m_Initialized Then Exit Sub
    m_Initialized = True

    ' ① WS_THICKFRAME を付与してリサイズ可能にする
    '    Activate 内で呼ぶことで、hWnd を確実に取得できる
    Dim hForm As LongPtr
    hForm = FindFormHwnd(Me.Caption)
    If hForm <> 0 Then
        Dim sty As LongPtr
        sty = GetWindowLongPtr(hForm, GWL_STYLE)
        SetWindowLongPtr hForm, GWL_STYLE, sty Or WS_THICKFRAME
    End If

    ' ② WebView2 初期化
    StartWebView2
End Sub

Private Sub UserForm_Resize()
    Const TOOLBAR_H As Single = 21  ' ツールバー行の高さ（ポイント）
    Const BTN_W     As Single = 48
    Const MARGIN    As Single = 3
    On Error Resume Next

    ' --- ツールバーのリサイズ追従 ---
    Me.btnGo.Left  = Me.InsideWidth - BTN_W - MARGIN
    Me.btnGo.Width = BTN_W
    Me.txtUrl.Width = Me.btnGo.Left - MARGIN * 2

    ' --- wv2Container（Frame）のリサイズ追従 ---
    Me.wv2Container.Width  = Me.InsideWidth
    Me.wv2Container.Height = Me.InsideHeight - TOOLBAR_H

    ' --- WebView2 のバウンドを更新し、位置変化を通知する ---
    If m_Ready And Not m_wv2 Is Nothing Then
        Dim pxW As Long, pxH As Long
        pxW = Me.wv2Container.InsideWidth  * PT2PX
        pxH = Me.wv2Container.InsideHeight * PT2PX
        m_wv2.Resize 0, 0, pxW, pxH
        m_wv2.NotifyPositionChanged  ' IME位置・ポップアップ位置を再計算させる
    End If
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
    Debug.Print "[WebView2Frame] m_wv2_Ready: WebView2 の初期化が完了しました"
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

    Debug.Print "[WebView2Frame] StartWebView2"
    Debug.Print "  hFrame = 0x" & Hex(hFrame)
    Debug.Print "  Size   = " & pxW & " x " & pxH & " px"

    Me.Caption = "WebView2 Browser - 初期化中..."

    ' WebView2Core を初期化
    '   ★ 親hWnd = hFrame（UserForm上のFrameコントロール）★
    Dim ok As Boolean
    ok = m_wv2.Initialize(hFrame, 0, 0, pxW, pxH, "https://eschamali.github.io/StarterWebScrapingKit/")

    If Not ok Then
        MsgBox "WebView2 Initialize 失敗:" & vbCrLf & m_wv2.LastErrorDescription, vbCritical
        Exit Sub
    End If

    ' Ready になるまでメッセージポンプを回す
    Dim t As Single: t = Timer
    Do While Not m_wv2.IsReady And Timer - t < 20
        m_wv2.ProcessMessages
        DoEvents
    Loop

    If Not m_wv2.IsReady Then
        MsgBox "初期化タイムアウト" & vbCrLf & _
               "LastError: 0x" & Hex(m_wv2.LastErrorCode) & vbCrLf & _
               m_wv2.LastErrorDescription, vbCritical
    End If
End Sub

'----------------------------------------------------------------------
' FindFormHwnd
'   UserForm の Win32 hWnd を ThunderDFrame クラス名で検索する。
'----------------------------------------------------------------------
Private Function FindFormHwnd(ByVal Caption As String) As LongPtr
    FindFormHwnd = FindWindow(StrPtr("ThunderDFrame"), StrPtr(Caption))
End Function
