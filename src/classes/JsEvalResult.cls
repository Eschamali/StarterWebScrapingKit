VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsEvalResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================================
'               JavaScriptの実行結果を詳細に記録するための専用記録オブジェクトとなります
'   CDPなら、Runtime.evaluate/Runtime.callFunctionOn の Runtime.RemoteObject,exceptionDetails           https://chromedevtools.github.io/devtools-protocol/#/Runtime.RemoteObject
'   BiDiなら、script.evaluate/script.callFunction    の script.RemoteValue  ,script.ExceptionDetails    https://www.w3.org/TR/webdriver-bidi/#cddl-type-scriptevaluationresult
'
'                                            を記録します。
'
'   処理の仕組み上、プロパティメソッドが呼ばれるまで、`Get-value`以外の取り出しは行いません。
'   そのため毎回、同じプロパティメソッドで呼び出すとそのたびにJson分析が始まるため、使いまわす場合は、呼び出し側の変数に格納するようにしてください
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const ThisClassName As String = "JsEvalResult"    'ログのトレース用



'***************************************************************************************************
'                                ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 基本情報
'---------------------------------------------------------------------------------------------------'
Private m_isCDP             As Boolean      '制御中のプロトコル。`False`なら、BiDi
Private m_ResultObject      As BiDiCDPJConv  '実行結果部分のオブジェクト(`Runtime.RemoteObject` or `script.RemoteValue`)
Private m_ExceptionDetails  As BiDiCDPJConv  '実行例外部分のオブジェクト(`exceptionDetails` or `script.ExceptionDetails`)
Private m_isException       As Boolean      '実行時例外が発生したか？
Private m_VarTypeJS         As String       '結果型(JS向け)
Private m_VarTypeVBA        As VbVarType    '結果型(VBA向け)
Private m_RunJS             As String       '実際に実行したJavaScript※ログ用
Private m_AutoValue         As Variant      '優先順位に基づく最適な、`value`

'---------------------------------------------------------------------------------------------------
' [ SECTION ] CDP特有
'---------------------------------------------------------------------------------------------------'
Private m_VarSubType        As String       'オブジェクトタイプ(`type:=Object`時に入る)

'---------------------------------------------------------------------------------------------------
' [ SECTION ] その他
'---------------------------------------------------------------------------------------------------
Private m_LogID             As String       'ログID



'***************************************************************************************************
'                                 ■■■ 分析処理の入口 ■■■
'***************************************************************************************************
'* 機能　　：JavaScriptの実行結果を`VarType`で判別できる仕組みまで分析させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：isCDP          `True` :CDP
'                           `False`:BiDi
'            ReturnObject   実行結果`result`のノード情報。※例外情報も格納するため、結果のまんまが必要です
'            IFEXCEPTION    例外発生時の代替値
'            RunJS          実行したJavaScript内容。※指定すると、実行結果と実行内容の照らし合わせがし易くなります
'---------------------------------------------------------------------------------------------------
'* 注意事項：利用前に、CDP/BiDi自体のコマンド実行成功可否確認が必要です
'***************************************************************************************************
Public Sub InitForReturnObject(isCDP As Boolean, ReturnObject As BiDiCDPJConv, Optional IFEXCEPTION As Variant, Optional RunJS As String)
    Const FromProcedureName As String = ThisClassName & ".InitForReturnObject"
    Const exceptionSignName As String = "exception"


    '1. 各種初期化
    m_isCDP = isCDP
    m_RunJS = RunJS

    '2. プロトコルごとのオブジェクト分析
    If m_isCDP Then
        '1. 結果を格納
        Set m_ResultObject = ReturnObject.NodeKey("result")
        If CDPHelpers.currentLog.seeRawSendMsgDbg Then printMsg Debug_, "jsEval - RemoteObject - JsonString = " & m_ResultObject.Stringify, FromProcedureName

        '2. 例外情報に応じた処理
        If ReturnObject.ExistsKey("exceptionDetails") Then
            '2-1. 例外情報の方へ格納
            m_isException = True
            Set m_ExceptionDetails = ReturnObject.NodeKey("exceptionDetails")

            '2-2. 結果本体自体もあるが、わざわざ解析する必要性はないので、手動で型情報を設定して、抜ける
            '---VBA側---
            '`IFERROR`的な設定があったらそっちを優先に適用させる
            If IsMissing(IFEXCEPTION) Then
                m_VarTypeVBA = vbError
                m_AutoValue = CVErr(xlErrValue)
            Else
                If IsObject(IFEXCEPTION) Then
                    m_VarTypeVBA = vbObject
                    Set m_AutoValue = IFEXCEPTION
                Else
                    m_VarTypeVBA = VarType(IFEXCEPTION)
                    m_AutoValue = IFEXCEPTION
                End If
            End If
            '---JS側---
            m_VarTypeJS = exceptionSignName
            m_VarSubType = vbNullString
            printMsg Debug_, "JsResultType: " & m_VarTypeJS, FromProcedureName
            Exit Sub
        Else
            '2-1. 例外情報はない旨として更新
            m_isException = False
            Set m_ExceptionDetails = Nothing
        End If

        '3. JS結果の型情報を取り出す
        m_VarTypeJS = m_ResultObject.StringKey("type")
        m_VarSubType = m_ResultObject.StringKey("subtype")
        printMsg Debug_, "JsResultType: " & m_VarTypeJS & " , subtype: " & m_VarSubType, FromProcedureName

        '4. `value`の分析へ
        AnalysisRuntimeRemoteObject
    Else
        '1. 例外情報に応じた処理
        If ReturnObject.StringKey("type") = "success" Then
            '1-1. 実行結果の方へ格納
            m_isException = False
            Set m_ResultObject = ReturnObject.NodeKey("result")
            Set m_ExceptionDetails = Nothing
            If CDPHelpers.currentLog.seeRawSendMsgDbg Then printMsg Debug_, "jsEval - script.EvaluationResultSuccess - JsonString = " & m_ResultObject.Stringify, FromProcedureName

        Else
            '1-1. 例外情報の方へ格納
            m_isException = True
            Set m_ResultObject = Nothing
            Set m_ExceptionDetails = ReturnObject.NodeKey("exceptionDetails")
            If CDPHelpers.currentLog.seeRawSendMsgDbg Then printMsg Debug_, "jsEval - script.EvaluationResultException - JsonString = " & m_ExceptionDetails.Stringify, FromProcedureName

            '1-2. 結果本体はないので、手動で型情報を設定して、抜ける
            '---VBA側---
            '`IFERROR`的な設定があったらそっちを優先に適用させる
            If IsMissing(IFEXCEPTION) Then
                m_VarTypeVBA = vbError
                m_AutoValue = CVErr(xlErrValue)
            Else
                If IsObject(IFEXCEPTION) Then
                    m_VarTypeVBA = vbObject
                    Set m_AutoValue = IFEXCEPTION
                Else
                    m_VarTypeVBA = VarType(IFEXCEPTION)
                    m_AutoValue = IFEXCEPTION
                End If
            End If
            '---JS側---
            m_VarTypeJS = exceptionSignName
            m_VarSubType = vbNullString '※BiDiはないので、ヌル文字で
            printMsg Debug_, "JsResultType: " & m_VarTypeJS, FromProcedureName
            Exit Sub
        End If

        '2. JS結果の型情報を取り出す
        m_VarTypeJS = m_ResultObject.StringKey("type")
        m_VarSubType = vbNullString '※BiDiはないので、ヌル文字で
        printMsg Debug_, "JsResultType: " & m_VarTypeJS, FromProcedureName

        '3. `value`の分析へ
        AnalysisScriptRemoteValue
    End If
End Sub

'***************************************************************************************************
'* 機能　　：JavaScriptの実行結果を`VarType`で判別できる仕組みまで分析させます
'---------------------------------------------------------------------------------------------------
'* 引数　　：isCDP          `True` :CDP
'                           `False`:BiDi
'            ResultObject   CDPなら`Runtime.RemoteObject`、BiDiなら`script.RemoteValue`のノードを指定します
'            RunJS          実行したJavaScript内容。※指定すると、実行結果と実行内容の照らし合わせがし易くなります
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`ExceptionResultObject`から呼び出す想定です
'***************************************************************************************************
Public Sub InitForResultObject(isCDP As Boolean, ResultObject As BiDiCDPJConv, Optional RunJS As String)
    Const FromProcedureName As String = ThisClassName & ".InitForResultObject"


    '1. 各種初期化
    m_isCDP = isCDP
    m_RunJS = RunJS
    Set m_ResultObject = ResultObject

    '2. プロトコルごとのオブジェクト分析
    If m_isCDP Then
        '2-1. JS結果の型情報を取り出す
        m_VarTypeJS = m_ResultObject.StringKey("type")
        m_VarSubType = m_ResultObject.StringKey("subtype")
        printMsg Debug_, "JsResultType: " & m_VarTypeJS & " , subtype: " & m_VarSubType, FromProcedureName

        '2-2. `value`の分析へ
        AnalysisRuntimeRemoteObject
    Else
        '2-1. JS結果の型情報を取り出す
        m_VarTypeJS = m_ResultObject.StringKey("type")
        m_VarSubType = vbNullString '※BiDiはないので、ヌル文字で
        printMsg Debug_, "JsResultType: " & m_VarTypeJS, FromProcedureName

        '2-2. `value`の分析へ
        AnalysisScriptRemoteValue
    End If
End Sub



'***************************************************************************************************
'                            ■■■ 各プロトコルごとのValue分析 ■■■
'***************************************************************************************************
'* 機能　　：CDPとしての、JavaScript実行で返った`Runtime.RemoteObject`から、最適な結果値を取り出します
'---------------------------------------------------------------------------------------------------
'* 注意事項：返す優先順位は、`deepSerializedValue` > `objectId` > `value` > `unserializableValue` として、`m_AutoValue`に格納します
'***************************************************************************************************
Private Sub AnalysisRuntimeRemoteObject()
    Const FromProcedureName As String = ThisClassName & ".AnalysisRuntimeRemoteObject"


    '------------------------ 1. 特定項目有無チェック ------------------------
    '`deepSerializedValue`があったらそれを最優先に返します
    If m_ResultObject.ExistsKey("deepSerializedValue") Then
        '見つかったので、`deepSerializedValue`を返して、抜ける
        Set m_AutoValue = m_ResultObject.NodeKey("deepSerializedValue")

        '`Object`であるサインを返し、抜ける
        m_VarTypeVBA = vbObject
        printMsg info_, "JS-Return type: object-deepSerializedValue", FromProcedureName
        Exit Sub

    '`objectId`があったら、こっちを返します
    '※type=objectで、`returnByValue:=False`なら`null`以外、ここを通ります
    ElseIf m_ResultObject.ExistsKey("objectId") Then
        '見つかったので、`objectId`を返す
        m_AutoValue = m_ResultObject.StringKey("objectId")

        '`ObjectID`であるサインを返し、抜ける
        m_VarTypeVBA = vbDataObject '見た目は文字列、実態はObjectという特殊性質を表すものとして付与
        printMsg info_, "JS-Return type: objectId", FromProcedureName
        Exit Sub
    End If


    '------------------------ 2. primitive values OR JsonNodeの取得を試みる ------------------------
    '以降は基本、`returnByValue:=True`でないと通りません
    '`primitive values` の場合も、ここを通り、値の取得を行います
    '例外として、`object-null`の場合は`returnByValue:=False`でもここを通って、null返却します

    '型に応じた分岐処理
    Select Case m_VarTypeJS
        Case "object"
            ' object型の場合はsubtypeで更に細かく分岐
            Select Case m_VarSubType
                Case "null"
                    m_AutoValue = Null
                    printMsg info_, "JS-Return type: object-null", FromProcedureName

                Case Else
                    '`returnByValue:=True`の場合は`null`以外全て、Json文字列として生成されます。
                    'つまり、`type:="Object"`ではあるが、`subtype`は空欄、でも値自体はあるということです。
                    'よって、Node形式として返却します。
                    Set m_AutoValue = m_ResultObject.NodeKey("value")
                    printMsg info_, "JS-Return type: object(returnByValue:=True)", FromProcedureName

            End Select

        Case "function"
            ' 関数が返ってきた場合は、「{"type":"function","value":{}}」なので、`SET`
            Set m_AutoValue = m_ResultObject.NodeKey("value")
            printMsg info_, "JS-Return type: function", FromProcedureName

        Case "undefined"
            ' JSの undefined は VBAでは Empty が一番近い
            m_AutoValue = Empty
            printMsg info_, "JS-Return type: undefined", FromProcedureName

        Case "string"
            ' 文字列として認識するように返す
            m_AutoValue = m_ResultObject.StringKey("value")
            printMsg info_, "JS-Return type: string", FromProcedureName

        Case "number"
            'VBAで扱える型か？
            If m_ResultObject.ExistsKey("unserializableValue") Then
                'VBAで扱えない特殊な"number"の場合は、警告を出し、「#NUM!」で返します。
                printMsg WARN_, "This `number` can't be processed in VBA. It returns `#NUM!` instead. Actual JS result: " & m_ResultObject.StringKey("unserializableValue") & " , description: " & m_ResultObject.StringKey("description"), FromProcedureName
                m_AutoValue = CVErr(xlErrNum)
            Else
                ' Double型として認識するように返す
                ' ※JavaScriptは、整数専用等の型がないため
                m_AutoValue = CDbl(m_ResultObject.NumberKey("value"))
            End If

            printMsg info_, "JS-Return type: number", FromProcedureName

        Case "boolean"
            ' boolean型として認識するように返す
            m_AutoValue = m_ResultObject.BoolKey("value")
            printMsg info_, "JS-Return type: boolean", FromProcedureName

        Case "symbol"
            '※`symbol`は、`returnByValue:=True`にすると、「Object couldn't be returned by value.」というCDPエラーで返るので通らないと思うが念のため。
            m_AutoValue = m_ResultObject.Item("value")
            printMsg info_, "JS-Return type: symbol", FromProcedureName

        Case "bigint"
            ' BigInt は JSON に変換できない巨大整数のため、unserializableValue の中にある値をそのまんま文字列として、返却します
            m_AutoValue = m_ResultObject.StringKey("unserializableValue")
            printMsg WARN_, "Unserializable BigInt detected. Returning as String: " & m_AutoValue & " , description: " & m_ResultObject.StringKey("description"), FromProcedureName
            printMsg info_, "JS-Return type: bigint", FromProcedureName

        Case Else
            'ドキュメント化されてない型があったら、安全のため。止めておく
            printMsg ERROR_, "This data type cannot be handled in VBA. type: " & m_VarTypeJS, FromProcedureName, doRaiseError:=True

    End Select

    '`Set`or`Let`判定用に確認
    m_VarTypeVBA = VarType(m_AutoValue)
End Sub

'***************************************************************************************************
'* 機能　　：BiDiとしての、JavaScript実行で返ったJsonDictionaryから、最適な結果値を取り出します
'---------------------------------------------------------------------------------------------------
'* 注意事項：返す優先順位は、`script.SharedId` > `script.Handle` > `value` として、`m_AutoValue`に格納します
'***************************************************************************************************
Private Sub AnalysisScriptRemoteValue()
    Const FromProcedureName As String = ThisClassName & ".AnalysisScriptRemoteValue"


    '------------------------ 1. 特定項目有無チェック ------------------------
    '※type=node時、ここを通ります
    '`script.sharedId`があったらそれを最優先に返します。
    If m_ResultObject.ExistsKey("sharedId") Then
        '見つかったので、`script.sharedId`を返す
        m_AutoValue = m_ResultObject.StringKey("sharedId")

        '`sharedId`であるサインを返し、抜ける
        m_VarTypeVBA = vbDataObject '見た目は文字列、実態はObjectという特殊性質を表すものとして付与
        printMsg info_, "JS-Return type: script.sharedId", FromProcedureName
        Exit Sub

    '※type=objectで、`resultOwnership`:="root"時、ここを通ります
    '`script.Handle`があったら、こっちを返します
    ElseIf m_ResultObject.ExistsKey("handle") Then
        '見つかったので、`script.Handle`を返す
        m_AutoValue = m_ResultObject.StringKey("handle")

        '`handle`であるサインを返し、抜ける
        m_VarTypeVBA = vbDataObject '見た目は文字列、実態はObjectという特殊性質を表すものとして付与
        printMsg info_, "JS-Return type: handle", FromProcedureName
        Exit Sub
    End If


    '------------------------ 2. primitive values OR JsonNodeの取得を試みる ------------------------
    '参考：https://w3c.github.io/webdriver-bidi/#type-script-RemoteValue
    '以降は基本、`resultOwnership:="none"`でないと通りません
    '`primitive values` の場合も、ここを通り、値の取得を行います
    printMsg info_, "JS-Return type: " & m_VarTypeJS, FromProcedureName

    Select Case m_VarTypeJS
        Case "symbol", "function", "weakmap", "weakset", "generator", "error", "proxy", "promise", "typedarray", "arraybuffer"
            '`value`そのものがないので、`Nothing`として返却
            Set m_AutoValue = Nothing
            printMsg WARN_, "`" & m_VarTypeJS & "` cannot be returned as a `value`.", FromProcedureName

        Case "array", "object", "map", "set", "nodelist", "htmlcollection", "window", "regexp", "node"
            'Json文字列として生成されるので、Node形式のObjectとして返却
            Set m_AutoValue = m_ResultObject.NodeKey("value")

        Case "undefined"
            ' JSの undefined は VBAでは Empty が一番近い
            m_AutoValue = Empty

        Case "null"
            m_AutoValue = Null

        Case "string"
            ' 文字列として認識するように返す
            m_AutoValue = m_ResultObject.StringKey("value")

        Case "number"
            '1. 一旦おまかせで取る
            m_AutoValue = m_ResultObject.Item("value")

            '2. VBAで扱えない特殊な"number"(その場合は大抵、文字列として格納されている)であるかを判定
            If VarType(m_AutoValue) = vbString Then
                '警告を出し、「#NUM!」で上書き
                printMsg WARN_, "This `number` can't be processed in VBA. It returns `#NUM!` instead. Actual JS result: " & m_AutoValue, FromProcedureName
                m_AutoValue = CVErr(xlErrNum)
            Else
                '`Double`として置き換え
                m_AutoValue = CDbl(m_AutoValue)
            End If

        Case "boolean"
            ' boolean型として認識するように返す
            m_AutoValue = m_ResultObject.BoolKey("value")

        Case "bigint"
            ' BigInt は JSON に変換できない巨大整数のため、値を文字列としてそのまんま返却します
            m_AutoValue = m_ResultObject.StringKey("value")
            printMsg WARN_, "Unserializable BigInt detected. Returning as String: " & m_AutoValue, FromProcedureName

        Case "date"
            'BiDiの場合、`date`は文字列として返却されるため、VBAの`vbDate`として返してあげる
            m_AutoValue = UtcConverter.ParseIso(m_ResultObject.StringKey("value"))

        Case Else
            'ドキュメント化されてない型があったら、安全のため。止めておく
            printMsg ERROR_, "This data type cannot be handled in VBA. type: " & m_VarTypeJS, FromProcedureName, doRaiseError:=True

    End Select

    '`Set`or`Let`判定用に確認
    m_VarTypeVBA = VarType(m_AutoValue)
End Sub



'***************************************************************************************************
'                              ■■■ 各種プロパティメソッド ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 例外情報(Runtime.ExceptionDetails/script.ExceptionDetails)
'             ※一旦、BiDiとの共通分のみ用意します
'---------------------------------------------------------------------------------------------------
Property Get rawException() As BiDiCDPJConv
    Set rawException = m_ExceptionDetails
End Property

Property Get isException() As Boolean
    isException = m_isException
End Property

Property Get ExceptionColumnNumber() As Long
    '例外情報がない時は、即抜け
    If Not m_isException Then Exit Property

    '抽出
    ExceptionColumnNumber = m_ExceptionDetails.NumberKey("columnNumber")
End Property

Property Get ExceptionResultObject() As JsEvalResult
    '1. 例外情報がない時は、即抜け
    If Not m_isException Then Exit Property

    '2. 新規生成
    Set ExceptionResultObject = New JsEvalResult

    '3. 分析依頼
    ExceptionResultObject.InitForResultObject m_isCDP, m_ExceptionDetails.NodeKey("exception"), m_RunJS
End Property

Property Get ExceptionLineNumber() As Long
    '例外情報がない時は、即抜け
    If Not m_isException Then Exit Property

    '抽出
    ExceptionLineNumber = m_ExceptionDetails.NumberKey("lineNumber")
End Property

Property Get ExceptionStackTrace() As BiDiCDPJConv
    '例外情報がない時は、即抜け
    If Not m_isException Then Exit Property

    '抽出
    Set ExceptionStackTrace = m_ExceptionDetails.NodeKey("stackTrace")
End Property

Property Get ExceptionText() As String
    '例外情報がない時は、即抜け
    If Not m_isException Then Exit Property

    '抽出
    ExceptionText = m_ExceptionDetails.StringKey("text")
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 実行結果情報 - BiDi/CDP共通
'---------------------------------------------------------------------------------------------------
Property Get ResultType() As String
    ResultType = m_VarTypeJS
End Property

Property Get ResultTypeForVBA() As VbVarType
    ResultTypeForVBA = m_VarTypeVBA
End Property

Property Get value() As Variant
Attribute value.VB_UserMemId = 0
    If m_VarTypeVBA = vbObject Then Set value = m_AutoValue Else value = m_AutoValue
End Property

Property Get rawValue() As Variant
    CDPHelpers.SetVar(rawValue) = m_ResultObject.Item("value")
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 実行結果情報 - CDPのみ
'---------------------------------------------------------------------------------------------------
Property Get ClassName() As String
    ClassName = m_ResultObject.StringKey("className")
End Property

Property Get Description() As String
    Description = m_ResultObject.StringKey("description")
End Property

Property Get objectId() As String
    objectId = m_ResultObject.StringKey("objectId")
End Property

Property Get ResultSubtype() As String
    ResultSubtype = m_VarSubType
End Property

Property Get UnserializableValue() As String
    UnserializableValue = m_ResultObject.StringKey("unserializableValue")
End Property

Property Get CustomPreview() As BiDiCDPJConv
    Set CustomPreview = m_ResultObject.NodeKey("customPreview")
End Property

Property Get DeepSerializedValue() As BiDiCDPJConv
    Set DeepSerializedValue = m_ResultObject.NodeKey("deepSerializedValue")
End Property

Property Get Preview() As BiDiCDPJConv
    Set Preview = m_ResultObject.NodeKey("preview")
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 実行結果情報 - BiDiのみ
'---------------------------------------------------------------------------------------------------
Property Get handle() As String
    handle = m_ResultObject.StringKey("handle")
End Property

Property Get internalId() As String
    internalId = m_ResultObject.StringKey("internalId")
End Property

Property Get sharedId() As String
    sharedId = m_ResultObject.StringKey("sharedId")
End Property



'***************************************************************************************************
'                                ■■■ 各種ラッパーメソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    'ログファイル拡張子は、制御中のプロトコルに合わせる
    Dim LogFileExtension As String
    LogFileExtension = IIf(m_isCDP, ".cdp.log", ".bidi.log")


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, CDPHelpers.currentLog.logPath
End Sub

Private Sub Class_Initialize()
    'ログID設定
    m_LogID = "JSID" & Format(Rnd * 1000, "000")
End Sub
