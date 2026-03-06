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
Private m_CharConv As New CharacterCodeConversion
Private m_JsonConverter As New WebJsonConverter

'ステート管理
Private m_MessageID As Long
Private m_Events As Dictionary



'***************************************************************************************************
'* 機能　　：初期化処理
'---------------------------------------------------------------------------------------------------
'* 引数　　：CDPInst        CDPBrowser インスタンス
'*           targetID       紐づける対象タブのtargetID
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiマッパーのスクリプトを読み込み、対象のコンテキストで実行可能にします。
'***************************************************************************************************
Public Sub Init(Optional StartURL As String, Optional SwitchUser As String)
    'セルから設定を適用
    Set m_CDPBrowser = 設定シートからの起動(StartURL, SwitchUser)

    '新規タブに対して、BiDiコンソールタブを準備し、メイン化
    m_CDPBrowser.newTab , setMain:=True
    m_TargetID = ShSetting01_StartBrowser.deserializeFromTable("Automation Data", 2, "TestBiDi")("targetID")    '犠牲先のTargetIDを取得

    'イベント初期化
    Set m_CDPBrowser.BiDiCapture = New Collection
    Set m_Events = New Dictionary

    'コマンドIDリセット
    m_MessageID = 0

    '1. BiDi用バインディングの追加
    Dim paramsCDP As New Dictionary
    paramsCDP.Add "bindingName", "cdp"
    paramsCDP.Add "targetId", m_TargetID
    paramsCDP.Add "inheritPermissions", True
    m_CDPBrowser.invokeMethod "Target.exposeDevToolsProtocol", paramsCDP, True

    '2. JSからVBAへメッセージを送るための関数名を "sendBidiResponse" に設定
    paramsCDP.RemoveAll
    paramsCDP.Add "name", "sendBidiResponse"
    m_CDPBrowser.invokeMethod "Runtime.addBinding", paramsCDP

    '3. パースしたJavaScriptをブラウザ上で実行
    m_CDPBrowser.jsEval Take_mapperTabJSstring

    '4. ブートストラップ（初期化）JSの実行
    m_CDPBrowser.jsEval "window.runMapperInstance('" & m_TargetID & "')"
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
    
    If m_CDPBrowser.BiDiCapture.Count Then
        Dim tmp As Variant
        Dim tmpJ As Object
        Dim eventList As Collection
        Set eventList = m_CDPBrowser.BiDiCapture
        
        For Each tmp In eventList
            Set tmpJ = m_JsonConverter.ParseJson(tmp)
            If tmpJ("params")("name") = "sendBidiResponse" Then
                Dim payloadStr As String
                payloadStr = tmpJ("params")("payload")
                
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
        Set m_CDPBrowser.BiDiCapture = New Collection
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



'***************************************************************************************************
'                           ■■■ ワークシートからJS取得 ■■■
'***************************************************************************************************
'* 機能　　：指定のテーブルから、mapperTab.jsコードを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：mapperTab.jsコード
'---------------------------------------------------------------------------------------------------
'* 機能説明：テーブルに格納した分割文字列を結合して返却します
'***************************************************************************************************
Function Take_mapperTabJSstring() As String
    Dim tmp2
    Dim chunkNum As Long

    'JS置き場のワークシート制御へ
    With ShLibrary01_JS
        With .ListObjects(.UseTableName(1, "BiDiCore.InsertMapperTab"))
            'データがそもそもない場合はここで抜ける
            If .ListRows.Count = 0 Then Exit Function

            '一旦、2次元配列で受け取る
            tmp2 = .DataBodyRange
            chunkNum = UBound(tmp2)
        End With
    End With

    '1次元配列に変換
    ReDim tmp1(1 To chunkNum)
    Dim i As Long
    For i = 1 To chunkNum
        tmp1(i) = tmp2(i, 1)
    Next

    '返却
    Take_mapperTabJSstring = Join(tmp1, "")
End Function



'***************************************************************************************************
'                               ■■■ 更新関連 ■■■
'***************************************************************************************************
'* 機能　　：受け取った`mapperTab.js`のコードを、32,767文字ごとに分割し、テーブルに格納します
'***************************************************************************************************
Private Sub InsertMapperTab(Code As String)
    '1セルごとの格納文字数上限
    Const LimitCellCharacter As Long = 32767

    'テーブル追加時のアニメーション動作を止める
    Application.ScreenUpdating = False


    '1. 分割数を確認
    Dim NumberOfSplit As Long: NumberOfSplit = Len(Code) / LimitCellCharacter

    '2. 配列確保
    ReDim JsCodeChunk(NumberOfSplit, 0)

    '3. 分割して格納する
    Dim i As Long
    For i = 0 To NumberOfSplit
        JsCodeChunk(i, 0) = Mid(Code, LimitCellCharacter * i + 1, LimitCellCharacter)
    Next

    'JS置き場のワークシート制御へ
    With ShLibrary01_JS
        With .ListObjects(.UseTableName(1, "BiDiCore.InsertMapperTab"))

            '4. テーブルの中身を全消去
            If .ListRows.Count Then .DataBodyRange.Delete

            '5. 挿入範囲の確定
            Dim TargetRange As Range
            Set TargetRange = .ListColumns(1).Range(2).Resize(NumberOfSplit + 1, 1)

            '6. 挿入前に書式を「文字列(@)」に設定
            TargetRange.NumberFormatLocal = "@"

            '7. 挿入
            TargetRange.value = JsCodeChunk
        End With
    End With


    '元に戻す
    Application.ScreenUpdating = True
End Sub

'***************************************************************************************************
'* 機能　　：ローカルファイルから、JSコードを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功論理値
'* 引数　　：FolderPath     フォルダパス
'            FileName       JSファイル名
'---------------------------------------------------------------------------------------------------
'* 機能説明：読み取ったファイルと、Excelテーブルに挿入した文字列との一致確認も行っていきます
'***************************************************************************************************
Public Function FromLocalFile(FolderPath As String, Optional FileName As String = "mapperTab.js") As Boolean
    'ファイルから、Jsコードを取得
    Dim JsCode As String
    With m_CharConv
        JsCode = .BytesToString(.BytesFromSavedFile(FolderPath, FileName))
    End With

    'テーブルに置く作業
    InsertMapperTab JsCode

    '正しく置けてるか確認
    FromLocalFile = (Take_mapperTabJSstring = JsCode)
End Function
