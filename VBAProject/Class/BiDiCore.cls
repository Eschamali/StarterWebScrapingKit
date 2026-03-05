VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BiDiCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
' Class       : BiDiCore
' Author      : 
' Created     : 2026/03/05
' Description : WebDriver BiDiの低レベル通信を担うクラス。
'               CDPBrowser経由でBiDiマッパーを実行し、BiDiコマンドの送受信・イベント取得を行います。
'==============================================================================================================

Option Explicit

'内部参照用オブジェクト
Private m_CDPBrowser As CDPBrowser
Private m_TargetID As String
Private m_CharConv As CharacterCodeConversion
Private m_JsonConverter As WebJsonConverter

'ステート管理
Private m_MessageID As Long
Private m_Events As Dictionary

'***************************************************************************************************
'* 機能　　：初期化処理
'---------------------------------------------------------------------------------------------------
'* 引数　　：CDPInst        CDPBrowser インスタンス
'*           targetID       紐づける対象タブのtargetID
'*           mapperJSPath   mapperTab.js のフルパス
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiマッパーのスクリプトを読み込み、対象のコンテキストで実行可能にします。
'***************************************************************************************************
Public Sub Init(CDPInst As CDPBrowser, ByVal targetID As String, ByVal mapperJSPath As String)
    Set m_CDPBrowser = CDPInst
    m_TargetID = targetID
    Set m_CharConv = New CharacterCodeConversion
    Set m_JsonConverter = New WebJsonConverter
    Set m_Events = New Dictionary
    
    m_MessageID = 0
    
    '1. BiDi用バインディングの追加
    Dim paramsCDP As New Dictionary
    paramsCDP.Add "bindingName", "cdp"
    paramsCDP.Add "targetId", m_TargetID
    paramsCDP.Add "inheritPermissions", True
    m_CDPBrowser.invokeMethod "Target.exposeDevToolsProtocol", paramsCDP, True
    
    'JSからVBAへメッセージを送るための関数名を "sendBidiResponse" に設定
    paramsCDP.RemoveAll
    paramsCDP.Add "name", "sendBidiResponse"
    m_CDPBrowser.invokeMethod "Runtime.addBinding", paramsCDP
    
    '2. mapperTab.js の読み込みと評価
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(mapperJSPath) Then
        Err.Raise vbObjectError + 1, "BiDiCore.Init", "指定された mapperTab.js が見つかりません: " & mapperJSPath
    End If
    
    Dim jsCode As String
    ' 文字コード(UTF-8)で読み込むためにADODB.Stream等を使うか、CharacterCodeConversionクラスの機能を使用
    ' ここでは CharConv の機能を利用（引数仕様が FolderPath, FileName と想定）
    Dim folderPath As String, fileName As String
    folderPath = fso.GetParentFolderName(mapperJSPath)
    fileName = fso.GetFileName(mapperJSPath)
    Dim jsBytes() As Byte
    jsBytes = m_CharConv.BytesFromSavedFile(folderPath, fileName)
    jsCode = m_CharConv.BytesToString(jsBytes)
    
    ' パースしたJavaScriptをブラウザ上で実行
    m_CDPBrowser.jsEval jsCode
    
    '3. ブートストラップ（初期化）JSの実行
    m_CDPBrowser.jsEval "window.runMapperInstance('" & m_TargetID & "')"
    
    Set fso = Nothing
End Sub

'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (非同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "session.new")
'*           params        引数を持ったDictionary (省略可能)
'* 返値    ：送信したメッセージの ID
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを構成し、JSON化してブラウザへ投げます。結果の待機は行いません。
'***************************************************************************************************
Public Function invokeMethodAsync(ByVal methodName As String, Optional params As Dictionary = Nothing) As Long
    m_MessageID = m_MessageID + 1
    invokeMethodAsync = m_MessageID
    
    Dim bidiCmdObj As New Dictionary
    bidiCmdObj.Add "id", invokeMethodAsync
    bidiCmdObj.Add "method", methodName
    
    If params Is Nothing Then
        bidiCmdObj.Add "params", New Dictionary
    Else
        bidiCmdObj.Add "params", params
    End If
    
    Dim bidiCmdJson As String
    bidiCmdJson = m_JsonConverter.ConvertToJson(bidiCmdObj)
    
    'JSONのダブルクォートやエスケープ処理による構文乱れを防ぐため、単純な置換で対処するか、
    'もしくは シングルクォート内で安全にエスケープするか、CDPのパラメータとして渡します。
    'ここではPoCに倣い、シングルクォートで文字列化しJSに渡すため、シングルクォートとバックスラッシュをエスケープします。
    Dim safeJson As String
    safeJson = Replace(bidiCmdJson, "\", "\\")
    safeJson = Replace(safeJson, "'", "\'")
    
    Dim paramsCDP As New Dictionary
    paramsCDP.Add "expression", "window.onBidiMessage('" & safeJson & "');"
    
    ' コマンド送信
    m_CDPBrowser.invokeMethod "Runtime.evaluate", paramsCDP
End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "session.new")
'*           params        引数を持ったDictionary (省略可能)
'* 返値    ：ブラウザから返却されたBiDi結果 (パースされたDictionary)
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを送信後、返却されるまでイベントを収集・待機します。
'***************************************************************************************************
Public Function invokeMethod(ByVal methodName As String, Optional params As Dictionary = Nothing) As Dictionary
    Dim targetID As Long
    targetID = invokeMethodAsync(methodName, params)
    
    Do
        TakeEvents
        
        If m_Events.Exists(CStr(targetID)) Then
            Set invokeMethod = m_Events(CStr(targetID))
            m_Events.Remove CStr(targetID) '取得後はキューから消す
            Exit Do
        End If
        
        m_CDPBrowser.sleep
        DoEvents
    Loop While True
End Function

'***************************************************************************************************
'* 機能　　：イベントの取得・パース
'---------------------------------------------------------------------------------------------------
'* 機能説明：CDP側のイベントを取り込み、BiDi（`Runtime.bindingCalled`）からのコールバックデータを
'*           JSONパースして内部キューに仕分けます。
'***************************************************************************************************
Public Sub TakeEvents()
    m_CDPBrowser.TakeEvents
    
    Const SearchEventName As String = "Runtime.bindingCalled"
    
    If m_CDPBrowser.BrowserEvents("EventMethods").Exists(SearchEventName) Then
        Dim tmp As Variant
        Dim eventList As Collection
        Set eventList = m_CDPBrowser.BrowserEvents("EventMethods")(SearchEventName)
        
        For Each tmp In eventList
            If tmp("params")("name") = "sendBidiResponse" Then
                Dim payloadStr As String
                payloadStr = tmp("params")("payload")
                
                'BiDiレスポンスのパース
                Dim bidiRes As Dictionary
                Set bidiRes = m_JsonConverter.ParseJson(payloadStr)
                
                '結果が id 付きのレスポンス（コマンド結果）か、イベントコールバックか判定
                If bidiRes.Exists("id") Then
                    ' コマンドの応答
                    Dim resId As String
                    resId = CStr(bidiRes("id"))
                    ' 念のため Null判定
                    If Not IsNull(bidiRes("id")) Then
                        If Not m_Events.Exists(resId) Then
                            m_Events.Add resId, bidiRes
                        End If
                    End If
                ElseIf bidiRes.Exists("method") Then
                    ' BiDiの非同期イベント (例: log.entryAdded など)
                    Dim evName As String
                    evName = bidiRes("method")
                    
                    ' イベント用のキューに保存（必要に応じてCollectionのリストにする）
                    If Not m_Events.Exists(evName) Then
                        m_Events.Add evName, New Collection
                    End If
                    m_Events(evName).Add bidiRes
                End If
            End If
        Next
        
        ' 処理が終わったら、CDP側の蓄積済みイベントリストをクリアする(重複処理を防ぐため)
        m_CDPBrowser.BrowserEvents("EventMethods").Remove SearchEventName
    End If
End Sub

'***************************************************************************************************
'* プロパティ：Events
'---------------------------------------------------------------------------------------------------
'* 機能説明：内部で蓄積したBiDi側のイベントキュー(Dictionary等)を取得します。
'***************************************************************************************************
Public Property Get Events() As Dictionary
    Set Events = m_Events
End Property
