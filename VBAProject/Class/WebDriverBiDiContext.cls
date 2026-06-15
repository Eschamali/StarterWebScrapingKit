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
Private WithEvents m_BiDiCore   As WebDriverBiDiCore    'イベント受け取りオブジェクト
Attribute m_BiDiCore.VB_VarHelpID = -1
Private m_BiDiMode              As WebDriverBiDiMode    'BiDi全体
Attribute m_BiDiMode.VB_VarHelpID = -1

'各種ステータス
Private Type browsingContextInfo
    context          As String   '現在、制御中のContext(タブ)パラメーター
    Url              As String   '現時点でのURL
End Type
Private InfoList As browsingContextInfo
Private contextDestroyed As Boolean       'お亡くなりサイン

Private LogID  As String



'***************************************************************************************************
'                              ■■■ ブラウザ起動関連処理 ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに対する所定の起動引数をつけて、WebDriverBiDiとして、起動させ、非同期イベント：`browsingContext.contextCreated`を利用した初回タブ接続処理を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：Name                           使用ブラウザ名※現時点では、Edgeとchromeのみです
'            appUrl                         最初の起動時に遷移したいURL
'            cleanActive                    `True` :デバッグブラウザを(あれば)Kill後、ブラウザを起動する
'                                           `False`:起動チェックをせずに、ブラウザを起動します
'
'            userProfile                    UserDataDir名を指定します
'            addArgs                        追加の起動引数を設けます
'            KioskMode                      0(省略)：通常モード(キオスクモードは使いません)
'                                           1      ：キオスクモード デジタル/対話型サイネージ
'                                           2      ：キオスクモード パブリック ブラウジング
'
'            sessionCapabilitiesRequest     `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'***************************************************************************************************
Public Sub StartBiDiModeAndConnectTab(Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional cleanActive As Boolean = True, _
    Optional userProfile As String, Optional addArgs As String, Optional KioskMode As edgeKioskType, _
    Optional sessionCapabilitiesRequest As Dictionary)


    '1. 初期化処理
    Set m_BiDiMode = New WebDriverBiDiMode

    '2. WebDriverBiDi としてブラウザを起動させ、`Session`を確立
    m_BiDiMode.StartBiDiMode Name, appUrl, cleanActive, userProfile, addArgs, KioskMode, sessionCapabilitiesRequest

    '3. イベントを継承
    Set m_BiDiCore = m_BiDiMode.InheritanceBiDiCore

    '4. タブの接続へ
    attachToContext
End Sub



'***************************************************************************************************
'                              ■■■ BiDi非同期イベント関連 ■■■
'***************************************************************************************************
'* 機能　　：BiDiの非同期イベント
'---------------------------------------------------------------------------------------------------
'* 機能説明：`Context`の生存確認に使います
'***************************************************************************************************
Private Sub m_BiDiCore_detectionBiDiEvent(methodName As String, RawJson As String)
    Select Case methodName
        '`context`が消えた
        Case "browsingContext.contextDestroyed"
            '1. 自分の`context`か確認し、そうであれば、フラグ更新

    End Select
End Sub



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
    params.Add "context", InfoList.context

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
    params.Add "context", InfoList.context

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
'* 機能　　：初回の`browsingContext.contextCreated`にある`context`を取り出します
'***************************************************************************************************
Private Sub attachToContext()
    Const FromProcedureName As String = ThisClassName & ".attachToContext"


    '1. 箱を用意
    Dim FirstTabs As Collection: Set FirstTabs = New Collection

    '2. その箱を渡す
    Set m_BiDiMode.RunCollect_InfoList = FirstTabs

    '3. 購読処理
    m_BiDiMode.EnableDiscoverContexts = True

    '4. 初回購読は直ぐに、`browsingContext.contextCreated`が来るので、それを拝見
    Dim tmp
    Dim FirstTab As Dictionary
    For Each tmp In FirstTabs
        '1. Jsonパース処理
        Set FirstTab = m_BiDiMode.jsConverter.ParseJson(tmp)("params")

        '2. 純粋な「新しいタブ」または「新しいウィンドウ」か判定し、そうであれば、`context`を適用
        If IsNull(FirstTab("parent")) Then
            With InfoList
                .context = FirstTab("context")
                .Url = FirstTab("url")
            End With

            printMsg info_, "I was able to obtain a valid `context` from `browsingContext.contextCreated`. → " & InfoList.context, FromProcedureName
            Exit For
        End If
    Next

    '5. 後始末
    Set m_BiDiMode.RunCollect_InfoList = Nothing

    '6. `context`が取れてない場合は停止
    If StrPtr(InfoList.context) = 0 Then printMsg ERROR_, "A valid `context` could not be obtained from `browsingContext.contextCreated`.", FromProcedureName, doRaiseError:=True
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
    Dim newSessionID As String: newSessionID = m_BiDiCore.DirectControlCDP.attachToTab(InfoList.context)

    '2. 発行失敗時はエラーで即抜け
    If newSessionID = "" Then m_BiDiMode.printMsg WARN_, "Failed to issue a SessionID for CDP from the WebDriver BiDi context. > " & m_BiDiCore.DirectControlCDP.LastCDPJsonError("message"), FromProcedureName: Exit Function

    '3. 無事発行出来たら、変換手続き
    '------ CDPContext.clsとして返却する準備 ------
    m_BiDiCore.DirectControlCDP.EnableDiscoverTargets = True    '`CDPContext.cls`としての制御なので、カウントしておく
    Set ConvertToCDPContext = New CDPContext
    With ConvertToCDPContext
        '4. BiDiブラウザのパイプハンドル情報を渡す
        Set .InheritanceCDPCore = m_BiDiCore.DirectControlCDP.InheritanceCDPCore

        '5. BiDiブラウザ情報を渡す
        Set .InheritanceCDPBrowser = m_BiDiCore.DirectControlCDP

        '6. 先ほど発行したSessionIDを渡す
        .id = newSessionID

        '7. ステータスの更新
        .getSessionInfo
    End With
End Function

'***************************************************************************************************
'                              ■■■ PUBLIC PROPERTIES ■■■
'***************************************************************************************************
'* 機能　　：`context`(タブ等)が閉鎖されたか確認します
'---------------------------------------------------------------------------------------------------
'* 返値    ：論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：このClassの`targetId`に対する`browsingContext.contextDestroyed`を受け取ると、`True`になります
'***************************************************************************************************
Property Get isContextClosed() As Boolean
    isContextClosed = contextDestroyed
End Property



'***************************************************************************************************
'                        ■■■ WebDriverBiDiCore.cls 直通メソッド ■■■
'***************************************************************************************************
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    m_BiDiCore.printMsg LogLevel_, LogID, strMsg, From, isHeader, doRaiseError
End Sub



'***************************************************************************************************
'                              ■■■ 複製/継承用プロパティ ■■■
'***************************************************************************************************
'* 機能　　：このClass内で定義してる`WebDriverBiDiMode.cls`を受け取ったり継承させます
'***************************************************************************************************
Property Get InheritanceWebDriverBiDiMode()
    Set InheritanceWebDriverBiDiMode = m_BiDiMode
End Property
Property Set InheritanceWebDriverBiDiMode(Target As WebDriverBiDiMode)
    Set m_BiDiMode = Target
End Property

'***************************************************************************************************
'* 機能　　：このClass内で定義してる`WebDriverBiDiCore.cls`を受け取ったり継承させます
'***************************************************************************************************
Property Get InheritanceWebDriverBiDiCore()
    Set InheritanceWebDriverBiDiCore = m_BiDiCore
End Property
Property Set InheritanceWebDriverBiDiCore(Target As WebDriverBiDiCore)
    Set m_BiDiCore = Target
End Property

'***************************************************************************************************
'* 機能　　：`browsingContext.getTree`等で得た`context`を適用させます
'***************************************************************************************************
Property Let context(param As String)
    InfoList.context = param
End Property



'***************************************************************************************************
'                       ■■■ このClassオブジェクト生成/破棄時のアクション ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    'Assign an objID for easy debug identification
    LogID = "BTID" & Format(Rnd * 1000, "000")
End Sub

Private Sub Class_Terminate()
    m_BiDiMode.EnableDiscoverContexts = False
End Sub
