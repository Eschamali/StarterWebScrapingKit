VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriverBiDiContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'           WebDriverBiDi制御のうち、Context(タブ等)操作としてのメソッドやパラメータ運用を行います
'               必要最低限のメソッドを提供しますが実際の送信は、`WebDriverBiDiMode`側で行います
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const ThisClassName As String = "WebDriverBiDiContext" 'ログのトレース用



'***************************************************************************************************
'                                 ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'WebDriverBiDi全体制御用クラスオブジェクト
Private m_BiDiMode  As WebDriverBiDiMode
Attribute m_BiDiMode.VB_VarHelpID = -1

'現在、制御中のContext(タブ)パラメーター
Private m_contextId As String



'***************************************************************************************************
'                              ■■■ BiDicommand操作関連 ■■■
'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "browsingContext.navigate")
'*           params        引数を持ったDictionary (省略時、`context`の付与)
'*           StopBiDiError 実行失敗時に`Err.Raise`で停止し、以降の処理を止めるか
'* 返値    ：ブラウザから返却されたBiDi結果 (パースされたDictionary)
'---------------------------------------------------------------------------------------------------
'* 機能説明：`params`に、`context`を付け足した上で実行されます
'***************************************************************************************************
Public Function ExecuteBiDi(ByVal methodName As String, Optional params As Dictionary, Optional StopBiDiError As Boolean = True) As Dictionary
    '`context`情報の追加
    If params Is Nothing Then Set params = New Dictionary
    params.Add "context", m_contextId

    '実行
    Set ExecuteBiDi = m_BiDiMode.ExecuteBiDi(methodName, params, StopBiDiError)
End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (非同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "browsingContext.navigate")
'*           params        引数を持ったDictionary (省略時、`context`の付与)
'*           StopError     実行失敗時に`Err.Raise`で停止し、以降の処理を止めるか
'* 返値    ：送信したメッセージの ID
'---------------------------------------------------------------------------------------------------
'* 機能説明：`params`に、`context`を付け足した上で実行し、結果の待機は行いません。
'***************************************************************************************************
Public Function ExecuteBiDiAsync(ByVal methodName As String, Optional params As Dictionary, Optional StopError As Boolean = True) As Long
    '`context`情報の追加
    If params Is Nothing Then Set params = New Dictionary
    params.Add "context", m_contextId

    '実行
    ExecuteBiDiAsync = m_BiDiMode.ExecuteBiDiAsync(methodName, params, StopError)
End Function



'***************************************************************************************************
'                              ■■■ JavaScript コア機能 ■■■
'***************************************************************************************************
'                                           工事中...




'***************************************************************************************************
'                           ■■■ 各種BiDiヘルパーコマンド ■■■
'***************************************************************************************************
'* 機能　　：URL遷移を行います
'***************************************************************************************************
Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
    '1. 待機モードの設定
    Dim targetState As String
    Select Case till
        Case isLoading: targetState = "loading"
        Case isInteractive: targetState = "interactive"
        Case isComplete: targetState = "complete"
    End Select

    '2. パラメーター組み立て
    Dim BiDiparams As New Dictionary
    BiDiparams.Add "url", strURL
    BiDiparams.Add "wait", targetState

    '3. 実行
    ExecuteBiDi "browsingContext.navigate", BiDiparams
End Sub



'***************************************************************************************************
'                                   ■■■ 黒魔術 ■■■
'***************************************************************************************************
'* 機能　　：現在のContext(タブ等)に対して、`CDPContext`オブジェクトに変換します
'---------------------------------------------------------------------------------------------------
'* 返値    ：このClassで設定中の`context`に紐づく、`CDPContext`オブジェクト　※変換失敗時は`Nothing`で返ります
'---------------------------------------------------------------------------------------------------
'* 詳細情報：`browsingContext.getTree`等で得れる`context`は、`targetId`と一緒である仕様を活かした変換術です
'* 注意事項：`goog:cdp.getSession`による`sessionId`共有は不可能みたいなので、`Target.attachToTarget`で新しいSessionIdとして制御します
'***************************************************************************************************
Public Function ConvertToCDPContext() As CDPContext
    Const FromProcedureName As String = ThisClassName & ".ConvertToCDPContext"


    '1. context に紐づく、新規`SessionId`の発行を試みる
    Dim newSessionID As String: newSessionID = m_BiDiMode.DirectControlCDP.attachToTab(m_contextId)

    '2. 発行失敗時はエラーで即抜け
    If newSessionID = "" Then m_BiDiMode.printMsg WARN_, "Failed to issue a SessionID for CDP from the WebDriver BiDi context. > " & m_BiDiMode.DirectControlCDP.LastCDPJsonError("message"), FromProcedureName: Exit Function

    '3. 無事発行出来たら、変換手続き
    '------ CDPContext.clsとして返却する準備 ------
    Set ConvertToCDPContext = New CDPContext
    With ConvertToCDPContext
        '4. BiDiブラウザのパイプハンドル情報を渡す
        Set .InheritanceCDPCore = m_BiDiMode.DirectControlCDP.InheritanceCDPCore

        '5. BiDiブラウザ情報を渡す
        Set .InheritanceCDPBrowser = m_BiDiMode.DirectControlCDP

        '6. 先ほど発行したSessionIDを渡す
        .id = newSessionID

        '7. ステータスの更新
        .getSessionInfo
    End With
End Function



'***************************************************************************************************
'                              ■■■ 継承用プロパティ ■■■
'***************************************************************************************************
'* 機能　　：このClass内で定義してる`WebDriverBiDiMode.cls`を受け取ります
'***************************************************************************************************
Property Set InheritanceWebDriverBiDiMode(Target As WebDriverBiDiMode)
    Set m_BiDiMode = Target
End Property

'***************************************************************************************************
'* 機能　　：`browsingContext.getTree`等で得た`context`を適用させます
'***************************************************************************************************
Property Let context(param As String)
    m_contextId = param
End Property
