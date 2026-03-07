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
Private m_CharConv As New CharacterCodeConversion
Private m_JsonConverter As New WebJsonConverter

'ステート管理
Private m_MessageID As Long
Private m_Events As Dictionary
Private m_AccumulatedCommandResults As Dictionary
Private m_MaxCommandResults As Long
Private m_ErrorDetails As Dictionary

Private Const ThisClassName As String = "BiDiCore"                'ログのトレース用


'***************************************************************************************************
'* 機能　　：初期化処理
'---------------------------------------------------------------------------------------------------
'* 引数　　：StartURL                   最初の起動時に遷移したいURL
'*           SwitchUser                 設定シートとは別のユーザディレクトリを指定するときに使います
'            sessionCapabilitiesRequest `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'---------------------------------------------------------------------------------------------------
'* 機能説明：・BiDiマッパーのスクリプトをExcelテーブルから読み込み、対象のコンテキストで実行可能にします。
'            ・新規立ち上げ用に使い、セッション確立まで担います。
'***************************************************************************************************
Public Sub start(Optional StartURL As String, Optional SwitchUser As String, Optional sessionCapabilitiesRequest As Dictionary)
    Const FromProcedureName As String = ThisClassName & ".Init"


    'セルから設定を適用
    Set m_CDPBrowser = 設定シートからの起動(StartURL, SwitchUser)
    With m_CDPBrowser
        '新規タブに対して、BiDiコンソールタブを準備し、メイン化
        .newTab , setMain:=True
        m_CDPBrowser.printMsg Trace_, "Starting `BiDi-CDP Mapper`...", FromProcedureName

        'イベント初期化
        Set .BiDiCapture = New Collection
        Set m_Events = New Dictionary
        Set m_AccumulatedCommandResults = New Dictionary
        Set m_ErrorDetails = New Dictionary
        m_MaxCommandResults = 100

        'コマンドIDリセット
        m_MessageID = 0

        '1. BiDi用バインディングの追加
        Dim paramsCDP As New Dictionary
        paramsCDP.Add "bindingName", "cdp"
        paramsCDP.Add "targetId", m_CDPBrowser.CurrentTargetID
        paramsCDP.Add "inheritPermissions", True
        m_CDPBrowser.invokeMethod "Target.exposeDevToolsProtocol", paramsCDP, True

        '2. JSからVBAへメッセージを送るための関数名を "sendBidiResponse" に設定
        paramsCDP.RemoveAll
        paramsCDP.Add "name", "sendBidiResponse"
        m_CDPBrowser.invokeMethod "Runtime.addBinding", paramsCDP

        '3. パースしたJavaScriptをブラウザ上で実行
        m_CDPBrowser.jsEval Take_mapperTabJSstring

        '4. ブートストラップ（初期化）JSの実行
        m_CDPBrowser.jsEval "window.runMapperInstance('" & m_CDPBrowser.CurrentTargetID & "')"
        m_CDPBrowser.printMsg Debug_, "`BiDi-CDP Mapper` has been launched", FromProcedureName

        '5. BiDi とのセッションを確立する
        '※ここは、BiDi用のコマンドとして使用
        Dim resultBiDi As Dictionary

        '`session.new`の`Parameters`が省略されてる場合は、規定値を作成する
        If sessionCapabilitiesRequest Is Nothing Then
            Set sessionCapabilitiesRequest = New Dictionary
            sessionCapabilitiesRequest.Add "capabilities", New Dictionary
        End If

        '確立処理！
        Set resultBiDi = invokeMethod("session.new", sessionCapabilitiesRequest)
    End With
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
'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (非同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "session.new")
'*           params        引数を持ったDictionary (省略可能)
'*           StopError     実行失敗時に`Err.Raise`で停止し、以降の処理を止めるか
'* 返値    ：送信したメッセージの ID
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを構成し、JSON化してブラウザへ投げます。結果の待機は行いません。
'***************************************************************************************************
Public Function invokeMethodAsync(ByVal methodName As String, Optional params As Dictionary = Nothing, Optional StopError As Boolean = True) As Long
    Const FromProcedureName As String = ThisClassName & ".invokeMethodAsync"

    Dim paramStr As String
    Dim results As Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = m_JsonConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}", Async:=True)

    '非同期の場合は、`WriteFile`等のエラーをチェック
    If results.Items(0) Then
        If StopError Then
            m_CDPBrowser.printMsg ERROR_, "Failed to invokeAsync method """ & methodName & """", FromProcedureName, doRaiseError:=True
        Else
            m_CDPBrowser.printMsg WARN_, "Failed to invokeAsync method """ & methodName & """", FromProcedureName
        End If
    End If

    '実行したIDを返却
    invokeMethodAsync = results.keys(0)
    m_CDPBrowser.printMsg Debug_, "AsyncInvoked """ & methodName & """ with ID: " & invokeMethodAsync, FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：BiDiコマンドの送信 (同期)
'---------------------------------------------------------------------------------------------------
'* 引数　　：methodName    メソッド名 (例: "session.new")
'*           params        引数を持ったDictionary (省略可能)
'*           StopError     実行失敗時に`Err.Raise`で停止し、以降の処理を止めるか
'* 返値    ：ブラウザから返却されたBiDi結果 (パースされたDictionary)
'---------------------------------------------------------------------------------------------------
'* 機能説明：BiDiコマンドを送信後、返却されるまでイベントを収集・待機します。
'***************************************************************************************************
Public Function invokeMethod(ByVal methodName As String, Optional params As Dictionary = Nothing, Optional StopError As Boolean = True) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".invokeMethod"

    Dim paramStr As String
    Dim results As Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = m_JsonConverter.ConvertToJson(params)
    Set results = SendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}", Async:=False)

    'Print error if there is
    If invokeError(results) Then
        If StopError Then
            m_CDPBrowser.printMsg ERROR_, "Failed to invoke method """ & methodName & """" & vbCrLf & vbCrLf & "=== RawErrorResponse ===" & vbCrLf & m_JsonConverter.ConvertToJson(results), FromProcedureName, doRaiseError:=True
        Else
            m_CDPBrowser.printMsg WARN_, "Failed to invoke method """ & methodName & """", FromProcedureName
        End If
    End If

    'Return result object (BiDiの場合はトップレベルか、resultの中身を返すかですが、そのまま返す設計とします)
    Set invokeMethod = results
    m_CDPBrowser.printMsg Debug_, "Invoked """ & methodName & """ with result: " & m_JsonConverter.ConvertToJson(results), FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：コマンドの結果にエラーが含まれているかチェックします
'---------------------------------------------------------------------------------------------------
'* 返り値　：エラー有無
'* 引数　　：ResultsObj    `invokeMethod`の生結果のDictionary
'***************************************************************************************************
Public Function invokeError(ResultsObj As Dictionary) As Boolean
    'BiDiコマンドエラーは "type":"error" 等で返る
    If ResultsObj.Exists("type") Then
        If ResultsObj("type") = "error" Then
            Set m_ErrorDetails = ResultsObj
            invokeError = True
            Exit Function
        End If
    ElseIf ResultsObj.Exists("error") Then
        Set m_ErrorDetails = ResultsObj
        invokeError = True
        Exit Function
    End If
    invokeError = False
End Function

'***************************************************************************************************
'* 機能　　：`BiDiコマンド送信のコア処理
'---------------------------------------------------------------------------------------------------
'* 返値　　：Dictionary形式の結果。Async=Trueの場合は ID:StatusCode等の簡易結果
'***************************************************************************************************
Private Function SendMessage(strMessage As String, Optional Async As Boolean) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".SendMessage"

    m_MessageID = m_MessageID + 1
    Dim lngCurrentId As Long
    lngCurrentId = m_MessageID

    '末尾の「}」を削って id を付与する
    strMessage = Left(strMessage, Len(strMessage) - 1)
    strMessage = strMessage & ", ""id"":" & lngCurrentId & "}"

    m_CDPBrowser.printMsg Debug_, "sendBiDi-JsonCommand = " & strMessage, FromProcedureName

    'JSONのダブルクォートやエスケープ処理による構文乱れを防ぐためエスケープ
    Dim safeJson As String
    safeJson = Replace(strMessage, "\", "\\")
    safeJson = Replace(safeJson, "'", "\'")

    Dim paramsCDP As New Dictionary
    paramsCDP.Add "expression", "window.onBidiMessage('" & safeJson & "');"

    'CDP経由でBiDiコマンド送信 (CDPのinvokeMethodを使う。同期実行して送信エラーがないか確認)
    Dim cdpResult As Dictionary
    Set cdpResult = m_CDPBrowser.invokeMethod("Runtime.evaluate", paramsCDP, StopError:=False)

    If Async Then
        Set SendMessage = New Dictionary
        '簡易的な送信成否（本来はWriteFileの結果など）を格納して抜ける
        If m_CDPBrowser.invokeError(cdpResult) Then
            SendMessage.Add lngCurrentId, 1 'Error code 1
        Else
            SendMessage.Add lngCurrentId, 0 'Success code 0
        End If
        Exit Function
    End If

    '同期の場合は結果を待つ
    Dim timerStart As Double: timerStart = Timer
    Dim timerOut As Long: timerOut = 60 '60秒待機
    Dim ResultSendCommand As Boolean

    m_CDPBrowser.printMsg Trace_, "Starting Processes incoming messages from the browser...", FromProcedureName
    Do
        'イベントを取り込みながら目的のIDを探す
        TakeEvents checkCommandID:=lngCurrentId

        'm_EventsやAccumulatedの中で目的のIDが見つかったかのチェックを行う
        If m_AccumulatedCommandResults.Exists(lngCurrentId) Then
            'パースして結果オブジェクトとして返す
            Set SendMessage = m_JsonConverter.ParseJson(m_AccumulatedCommandResults(lngCurrentId))
            m_AccumulatedCommandResults.Remove lngCurrentId
            ResultSendCommand = True
        End If

        If Timer - timerStart > timerOut Then Err.Raise vbObjectError + 1, FromProcedureName, "Timeout waiting for BiDi to return message."
        
        If Not ResultSendCommand Then
            m_CDPBrowser.sleep
            DoEvents
        End If
    Loop Until ResultSendCommand

    m_CDPBrowser.printMsg Trace_, "Command result ID:" & lngCurrentId & " was returned from BiDi-Json, so the receiving loop will end.", FromProcedureName
End Function

'***************************************************************************************************
'* プロパティ：LastBiDiJsonError
'---------------------------------------------------------------------------------------------------
'* 機能説明：invokeErrorを通った際に保管した最後に起きたJSON形式の生エラーデータです
'***************************************************************************************************
Public Property Get LastBiDiJsonError() As Dictionary
    Set LastBiDiJsonError = m_ErrorDetails
End Property

'***************************************************************************************************
'* プロパティ：ResultBiDiForAsync
'---------------------------------------------------------------------------------------------------
'* 機能説明：指定した引数コマンドIDの非同期実行したコマンド結果を得ます
'***************************************************************************************************
Public Property Get ResultBiDiForAsync(CommandID As Long) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".ResultBiDiForAsync"

    If m_AccumulatedCommandResults.Exists(CommandID) Then
        Set ResultBiDiForAsync = m_JsonConverter.ParseJson(m_AccumulatedCommandResults(CommandID))
        m_CDPBrowser.printMsg info_, "Acquired result log from execution command ID. - ID: " & CommandID, FromProcedureName

        m_AccumulatedCommandResults.Remove CommandID
        m_CDPBrowser.printMsg Trace_, "Deleted the resulting log corresponding to the execution command ID. - ID: " & CommandID, FromProcedureName
    Else
        Set ResultBiDiForAsync = Nothing
        m_CDPBrowser.printMsg WARN_, "There is no target command ID result log. - ID: " & CommandID, FromProcedureName
    End If
End Property

'***************************************************************************************************
'* 機能　　：イベントの取得・パース
'---------------------------------------------------------------------------------------------------
'* 引数　　：checkCommandID  確認するコマンドID (同期送信で結果を待つ際に指定)
'---------------------------------------------------------------------------------------------------
'* 機能説明：CDP側のイベントを取り込み、BiDi（`Runtime.bindingCalled`）からのコールバックデータを
'*           高速判定し、内部キューに仕分けます。
'***************************************************************************************************
Public Sub TakeEvents(Optional checkCommandID As Long = -1)
    Const FromProcedureName As String = ThisClassName & ".TakeEvents"

    m_CDPBrowser.TakeEvents
    
    If m_CDPBrowser.BiDiCapture.Count Then
        Dim tmp As Variant
        Dim payloadStr As String
        Dim tmpJ As Dictionary
        Dim eventList As Collection
        Set eventList = m_CDPBrowser.BiDiCapture
        
        For Each tmp In eventList
            'CDP側の形式 {"method":"Runtime.bindingCalled","params":{"name":"sendBidiResponse","payload":"{...}"}}
            'まずはペイロードを取り出す
            Set tmpJ = m_JsonConverter.ParseJson(tmp)
            If tmpJ("params")("name") = "sendBidiResponse" Then
                payloadStr = tmpJ("params")("payload")
                
                'BiDiレスポンスを高速判定で振り分け
                BrowserReceivedDataCheck payloadStr, checkCommandID
            End If
        Next
        
        ' 処理が終わったら、CDP側の蓄積済みイベントリストをクリアする(重複処理を防ぐため)
        Set m_CDPBrowser.BiDiCapture = New Collection
    End If
End Sub

'***************************************************************************************************
'* 機能　　：BiDiコマンドの受信データを高速にチェックし、仕分けます
'---------------------------------------------------------------------------------------------------
'* 引数　　：BiDiJsonString    BiDiの生JSON文字列
'*           checkCommandID    待ち受け中のコマンド結果ID
'***************************************************************************************************
Private Function BrowserReceivedDataCheck(BiDiJsonString As String, checkCommandID As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".BrowserReceivedDataCheck"

    Dim ResultCommandID As Long
    Dim AsynchronousEventName As String
    Dim bidiRes As Dictionary

    '1. {"type":"success" または {"type":"error" ならコマンド結果
    If Left(BiDiJsonString, 16) = "{""type"":""success" Or Left(BiDiJsonString, 14) = "{""type"":""error" Then
        'IDを取得する
        If GetContentsID(BiDiJsonString, ResultCommandID) Then
            '現在待ち受け中のIDと一致するか？
            If ResultCommandID = checkCommandID Then
                '待ち受け中のIDだったので、この場ではパースせずに文字列のまま `m_AccumulatedCommandResults` に入れて終了（SendMessage側で取り出す）
                m_AccumulatedCommandResults.Add ResultCommandID, BiDiJsonString
                m_CDPBrowser.printMsg Trace_, "Matched expected Command ID: " & checkCommandID, FromProcedureName
            Else
                '非同期で実行したコマンドや、過去のタイムアウト済みの結果
                m_AccumulatedCommandResults.Add ResultCommandID, BiDiJsonString
                m_CDPBrowser.printMsg info_, "Received async Command ID: " & ResultCommandID, FromProcedureName

                '蓄積上限を超えたら古いものを削除
                If m_AccumulatedCommandResults.Count > m_MaxCommandResults Then
                    m_AccumulatedCommandResults.Remove m_AccumulatedCommandResults.keys(0)
                End If
            End If
        Else
            m_CDPBrowser.printMsg WARN_, "Type is success/error but no ID found: " & BiDiJsonString, FromProcedureName
        End If

    '2. {"type":"event" なら非同期イベント
    ElseIf Left(BiDiJsonString, 15) = "{""type"":""event""" Then
        If GetContentsMethod(BiDiJsonString, AsynchronousEventName) Then
            Set bidiRes = m_JsonConverter.ParseJson(BiDiJsonString)

            If Not m_Events.Exists(AsynchronousEventName) Then
                m_Events.Add AsynchronousEventName, New Collection
            End If
            m_Events(AsynchronousEventName).Add bidiRes
            m_CDPBrowser.printMsg Trace_, "Stored BiDi Event: " & AsynchronousEventName, FromProcedureName
        Else
            m_CDPBrowser.printMsg WARN_, "Type is event but no method found: " & BiDiJsonString, FromProcedureName
        End If

    '未知のパターン
    Else
        m_CDPBrowser.printMsg WARN_, "Unknown BiDi JSON format: " & BiDiJsonString, FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックし、`id`を高速抽出します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：jsonString         BiDiのレスポンス(JSON文字列そのまま)
'*            idValueContents   コマンド結果のID値を格納する変数(参照渡し)
'***************************************************************************************************
Private Function GetContentsID(jsonString As String, ByRef idValueContents As Long) As Boolean
    Const SearchKeyName As String = """id"":"
    Dim StartPos As Long
    Dim EndPos As Long
    
    StartPos = InStr(1, jsonString, SearchKeyName)
    If StartPos > 0 Then
        GetContentsID = True
        StartPos = StartPos + Len(SearchKeyName)
        EndPos = InStr(StartPos, jsonString, ",")
        If EndPos = 0 Then EndPos = InStr(StartPos, jsonString, "}") ' 末尾にある場合
        
        idValueContents = CLng(Mid(jsonString, StartPos, EndPos - StartPos))
    Else
        GetContentsID = False
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックし、`method`を高速抽出します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有无论理値
'* 引数　　：jsonString             BiDiのレスポンス(JSON文字列そのまま)
'*            methodValueContents  非同期イベントのメソッド名(参照渡し)
'***************************************************************************************************
Private Function GetContentsMethod(jsonString As String, ByRef methodValueContents As String) As Boolean
    Const SearchKeyName As String = """method"":"""
    Dim StartPos As Long
    Dim EndPos As Long
    
    StartPos = InStr(1, jsonString, SearchKeyName)
    If StartPos > 0 Then
        GetContentsMethod = True
        StartPos = StartPos + Len(SearchKeyName)
        EndPos = InStr(StartPos, jsonString, """")
        
        methodValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
    Else
        GetContentsMethod = False
    End If
End Function

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
            'データがそもそもない場合は、エラーで止める
            If .ListRows.Count = 0 Then
                Err.Raise vbObjectError, , "mapperTab.js がセットされていません"
                Exit Function
            End If

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
