VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromeDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWebDriver

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const CLASS_NAME As String = "Chrome_WidgetWin_1"
Private Const CAPTION_NAME As String = "新しいタブ - Google Chrome" '"data:, - Google Chrome"

''Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
'        ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
'

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "USER32.DLL" ( _
        ByVal hwnd As LongPtr, ByRef ProcessId As Long) As Long

Private Declare PtrSafe Sub SetForeGroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal ms As LongPtr)

Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr

Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As LongPtr


Private webSocket_ As a1_WebSocketCommunicator
Private json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos
Private CDP_ As a4_ExecuteCDP
Private events_ As a5_CDPEventHandler
Private Helper_ As a6_ExecuteHelperFunction

Private processId_ As Long
Private chromeHwnd_ As LongPtr
'Private userDataDir_ As String

Private Sub Class_Initialize()
'    userDataDir_ = GetTempFolderPath() & "BrowserDriver\"
     
    processId_ = StartChrome
    
    New_InnerObjects

    CDP_.SetDiscoverTargets
    CDP_.AttachToTarget basicInfo_.targetId
    
    WaitForActivation

    chromeHwnd_ = GetForegroundWindow
    
End Sub

Private Function GetTempFolderPath() As String
    Dim i As Long: i = 1
    Do
        If left(UCase(Environ(i)), 5) = "TEMP=" Then
            GetTempFolderPath = Mid(Environ(i), 6) & "\"
            Exit Do
        End If
        i = i + 1
    Loop Until Environ(i) = ""
End Function

Private Function StartChrome() As Long
    Const CHROME_EXE_REG_KEY As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\"
    
    Dim userDataDir As String
    userDataDir = CreateObject("Shell.Application").Namespace("shell:Downloads").Self.Path & "\マクロ実行専用UserDataフォルダForChrome\"
    Dim chromeExePath As String
    Dim objExec As Object
    With CreateObject("WScript.Shell")
        chromeExePath = """" & .RegRead(CHROME_EXE_REG_KEY) & """"
        Set objExec = .Exec(chromeExePath & "--remote-debugging-Port=9222 --disable-popup-blocking --no-first-run --disable-fre " _
                            & " --user-data-dir=" & "" & userDataDir & "")
                            ' --enable-automation"　’ OAuthでこれが原因で弾かれるケースがあるため不使用
                            ' --guest　’ ゲストモードでもユーザーデータディレクトリ必須
                            ' --bwsi  ' ゲストモードと同じ
                            
'        テスト用
'        Set objExec = .Exec(chromeExePath & "--remote-debugging-Port=9222 --disable-popup-blocking --no-first-run --disable-fre --bwsi")
                            
        StartChrome = objExec.ProcessId
        Set objExec = Nothing
    End With
End Function

'ブラウザ起動待ち＆SetForeGroundWindowにも使用
Private Sub WaitForActivation()
On Error Resume Next
    
    Dim counter As Long
recheck:
    AppActivate processId_
    If Err.Number <> 0 Then
        counter = counter + 1
        If counter > 20 Then
            MsgBox "10秒経過しましたがブラウザが有効になりませんでした。" & vbCrLf & _
                        "処理を終了します"
            End
        End If
        Debug.Print Err.Number & ":" & Err.description
        Helper_.SleepByWinAPI 500
        Err.Number = 0
        GoTo recheck
    End If
        
On Error GoTo 0
End Sub

'内部オブジェクトの生成
Private Sub New_InnerObjects()
    Set json_ = New a2_JSONHandler
    Set basicInfo_ = New a3_BasicInfos: basicInfo_.Init json_
    Set webSocket_ = New a1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.webSocketURLPath) = False Then End
    
    Set events_ = New a5_CDPEventHandler
    events_.Init json_, basicInfo_
    
    Dim msgGenerator As a4x1_CDPMessageGenerator
    Set msgGenerator = New a4x1_CDPMessageGenerator
    Dim Handler As a4x2_CDPMessageHandler
    Set Handler = New a4x2_CDPMessageHandler
    Handler.Init webSocket_, json_, events_
    Set CDP_ = New a4_ExecuteCDP
    CDP_.Init msgGenerator, Handler, json_
    
    Set Helper_ = New a6_ExecuteHelperFunction
End Sub




'Private Sub Class_Terminate()
'    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
'    If fso.FolderExists(userDataDir_) Then fso.DeleteFolder userDataDir_, True
'
'End Sub


































'以下､インターフェイス公開メソッド
'************************************************
'Public
'現在のページのURL
Private Property Get IWebDriver_URL() As String
    CDP_.SetDiscoverTargets  'これを使ってPage情報最新化
    Dim i As Long
    For i = 1 To events_.PageInfos.count
        If events_.PageInfos(i)("targetId") = basicInfo_.targetId Then
            IWebDriver_URL = events_.PageInfos(i)("url")
            Exit Property
        End If
    Next
End Property

'Public
'現在のページのURLのクエリ文字列を辞書で返す
Private Property Get IWebDriver_QueryString() As Object
    Dim currentURL As String: currentURL = IWebDriver_URL
    Dim QueryString As String: QueryString = CDP_.GetQueryStringFrom(currentURL)
    If QueryString = "" Then
        IWebDriver_QueryString = Nothing
    End If
    
    ' クエリ文字列内のkey,valueペアを辞書に変換
    If left(QueryString, 1) = "?" Then QueryString = Mid(QueryString, 2)
    Dim key_values As Variant: key_values = Split(QueryString, "&")
    Dim queryStringDic As Object: Set queryStringDic = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim key_value As Variant
    For i = LBound(key_values) To UBound(key_values)
        key_value = Split(key_values(i), "=")
        queryStringDic.Add key_value(0), key_value(1)
    Next
    Set IWebDriver_QueryString = queryStringDic
End Property

' Public
' クエリ文字列のキー指定による値取得
Private Function IWebDriver_GetValueFromQueryString(KeyName As String) As String
    Dim queryStringDic As Object
    Set queryStringDic = IWebDriver_QueryString()
    If queryStringDic Is Nothing Then
        IWebDriver_GetValueFromQueryString = ""
    Else
        If queryStringDic.exists(KeyName) Then
            IWebDriver_GetValueFromQueryString = queryStringDic(KeyName)
        Else
            IWebDriver_GetValueFromQueryString = ""
        End If
    End If
End Function

'Public
'現在のページのHTMLソースを取得
Private Function IWebDriver_PageSource(target As TargetIs) As String
    Dim res As String
    Select Case target
        Case TargetIs.html_outerhtml
            res = CDP_.ExecuteJavaScript("document.all[0].outerHTML;")
        Case TargetIs.html_innerhtml
            res = CDP_.ExecuteJavaScript("document.all[0].innerHTML;")
        Case TargetIs.body_outerhtml
            res = CDP_.ExecuteJavaScript("document.body.outerHTML;")
        Case TargetIs.body_innerhtml
            res = CDP_.ExecuteJavaScript("document.body.innerHTML;")
        Case TargetIs.body_innertext
            res = CDP_.ExecuteJavaScript("document.body.innerText;")
    End Select
    
    Dim htmlSource As String
    htmlSource = json_.GetValue(res, "result", "result", "value")
    IWebDriver_PageSource = htmlSource
End Function


'Public
'第1引数で指定のURLを開く
Private Sub IWebDriver_OpenURL(targetURL As String)
    Dim res As String
    res = CDP_.OpenURL(targetURL)
    If InStr(res, "invalid URL") > 0 Then
        Helper_.DisplayErrorAndEnd _
            "＜OpenURL＞" & vbCrLf & _
            "指定された以下URLは無効なURLです。" & vbCrLf & _
            targetURL & vbCrLf
    End If
    
    CDP_.WaitingForWebPageDisplay
    
    CDP_.SetDiscoverTargets
    CDP_.PageEnable
'    CDP_.DOMEnable
    basicInfo_.DocumentObjectId = CDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = CDP_.GetDocumentNodeID
End Sub

'Public
'第1引数のJavaScriptコードの実行
Private Function IWebDriver_ExecuteScript(script As String) As String
    Dim responseMsg As String
    responseMsg = CDP_.ExecuteJavaScript(script)
    IWebDriver_ExecuteScript = responseMsg
End Function

'Public
'第1引数のCDPメソッドのコードの実行
Private Function IWebDriver_ExecuteCDP(CDPMethodJson As String) As String
    IWebDriver_ExecuteCDP = CDP_.ExecuteDirect(CDPMethodJson)
End Function

'Public
'第1引数で指定されるXpathの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementByXpath(xPath As String) As IWebElement
    
    '属性指定時に日本語が含まれるとエラーになるため要エンコード
    xPath = json_.EncodeURL(xPath)
    
    Dim Codes As Variant: Codes = Helper_.GetXpathCodes(xPath)
    Dim objectId As String: objectId = CDP_.GetObjectIdOfElementBy(xPath, basicInfo_, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("xpath", xPath)
    Set IWebDriver_FindElementByXpath = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & xPath
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("XPath", xPath)
End Function

'Public
'第1引数で指定されるIdの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementById(Id As String) As IWebElement
    ' id指定に日本語が含まれるとエラーになるため要エンコード
    Id = json_.EncodeURL(Id)
    
    Dim Codes As Variant: Codes = Helper_.GetIdCodes(Id)
    Dim objectId As String: objectId = CDP_.GetObjectIdOfElementBy(Id, basicInfo_, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("id", Id)
    Set IWebDriver_FindElementById = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & Id
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Id", Id)
End Function

'Public
'第1引数で指定されるNameの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementByName(name As String) As IWebElement
    'name指定に日本語が含まれるとエラーになるため要エンコード
    name = json_.EncodeURL(name)
    
    Dim Codes As Variant: Codes = Helper_.GetNameCodes(name)
    Dim objectId As String: objectId = CDP_.GetObjectIdOfElementBy(name, basicInfo_, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("name", name)
    Set IWebDriver_FindElementByName = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & name
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Name", name)
End Function


'Public
'tag名で要素コレクションを取得（要素グループのLengthを取得し、Length分要素作成ループ）
Private Function IWebDriver_FindElementsByTag(tagName As String) As IWebElements
    Dim countCodes As Variant: countCodes = Helper_.GetTagCountCodes(tagName)
    Dim elemsCount As Long: elemsCount = CDP_.GetElementsCountBy(tagName, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    
    Dim objectId As String
    Dim Codes As Variant
    Dim index As Long
    For index = 0 To elemsCount - 1
        Codes = Helper_.GetTagCodes(tagName, index)
        objectId = CDP_.GetObjectIdOfElementBy(tagName, basicInfo_, Codes)
        Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("tag", tagName, index)
        elements.Add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByTag = elems
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & tagName
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Tag", tagName, "○○")
End Function

'Public
'class名で要素コレクションを取得（要素グループのLengthを取得し、Length分要素作成ループ）
Private Function IWebDriver_FindElementsByClassName(className As String) As IWebElements
    Dim countCodes As Variant: countCodes = Helper_.GetClassNameCountCodes(className)
    Dim elemsCount As Long: elemsCount = CDP_.GetElementsCountBy(className, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    
    Dim objectId As String
    Dim Codes As Variant
    Dim index As Long
    For index = 0 To elemsCount - 1
        Codes = Helper_.GetClassNameCodes(className, index)
        objectId = CDP_.GetObjectIdOfElementBy(className, basicInfo_, Codes)
        Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("className", className, index)
        elements.Add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByClassName = elems
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & className
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("className", className, "○○")
End Function

Private Function IWebDriver_FindElementsByXPath(xPath As String) As IWebElements
    '属性指定時に日本語が含まれるとエラーになるため要エンコード
    xPath = json_.EncodeURL(xPath)
    
    Dim countCodes As Variant: countCodes = Helper_.GetXPathCountCodes(xPath)
    Dim elemsCount As Long: elemsCount = CDP_.GetElementsCountBy(xPath, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    
    Dim objectId As String
    Dim Codes As Variant
    Dim index As Long
    For index = 0 To elemsCount - 1
        Codes = Helper_.GetXPathResultCodes(xPath, index)
        objectId = CDP_.GetObjectIdOfElementBy(xPath, basicInfo_, Codes)
        Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("xpath", xPath, index)
        elements.Add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, events_, Helper_)
    Next
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByXPath = elems
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & xPath
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("XPath", xPath, "○○")
End Function


'Public
'第1引数で指定される値のフレームにフレームを切替
Private Sub IWebDriver_SwitchFrameByNameOrUrlOrIndex(value As Variant)
    Dim frameInfos As Collection
    Dim FrameId As String
    Dim count As Long
    
    Do 'frame関連イベント発生タイミングによって情報取得に失敗するケースあり。念のため3回までループ。
        Set frameInfos = GetFrameInfos
        If IsNumeric(value) Then
            If frameInfos.count < value Then GoTo quit
            FrameId = json_.GetValue(frameInfos(value), "id")
        Else
            FrameId = GetTargetFrameId(frameInfos, value)
        End If
        count = count + 1
    Loop While FrameId = "" And count < 3
    
    If FrameId = "" Then GoTo quit
    
    Dim ExecutionContextId As Long
    ExecutionContextId = CDP_.CreateIsolatedFrameWorld(FrameId)
    basicInfo_.ExecutionContextId = ExecutionContextId
    basicInfo_.IsolatedFrameId = FrameId
    
Exit Sub

quit:
Helper_.DisplayErrorAndEnd "＜SwitchFrameByNameOrUrlOrIndex＞" & vbCrLf & _
                                "指定された以下のframeは存在しません。" & vbCrLf & _
                                value & vbCrLf & vbCrLf & _
                                "※iframeにName属性があれば、Nameでの指定を推奨します。" & vbCrLf & vbCrLf & _
                                "※フレームのIndex番号、Name、URLを確認したい場合、" & vbCrLf & _
                                "　 DebugPrintAccessibleFrameInfoを使用してください。" & vbCrLf & _
                                "　 クロームから通知されるフレーム情報がイミディエイトに出力されます。" & vbCrLf & vbCrLf & _
                                "※このメソッドでうまくいかない場合、" & vbCrLf & _
                                "　 SwitchFrameByIframeElementの" & vbCrLf & _
                                "　 使用を検討してください。" & vbCrLf
End Sub

'frameツリーを取得し、各フレームの情報（JSONオブジェクト文字列）をコレクションとして集約
'frameツリーはresult→frameTree
'→frame(Pageの親フレーム)とChildFrames(配列の中にPageの子frame情報(オブジェクト)）の構成
Private Function GetFrameInfos() As Collection
    Dim collect As Collection: Set collect = New Collection
    Dim json As String: json = CDP_.GetFrameTree
'    collect.add json_.GetObject(json, "result", "frameTree", "frame"))
    
    Dim frmInfo As String, frminfo2 As String
    Dim count As Long
    Do
        'まずchildFramesから一つのFrame情報（オブジェクト）取得
        frmInfo = json_.GetObject(json, "result", "frameTree", "childFrames", count)
        If frmInfo = "" Then Exit Do
        'そのオブジェクトの中のframeプロパティ(オブジェクト)を取得し、コレクションに加える
        frminfo2 = json_.GetObject(frmInfo, "frame")
        collect.Add frminfo2
        'childFramesプロパティがある場合、再帰構造なので再帰関数で同様の処理を回す
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        count = count + 1
    Loop
    Set GetFrameInfos = collect
End Function
'childFrameがある場合、再帰処理
Private Sub GetChildFrameInfo(collect As Collection, json As String)
    Dim frmInfo As String, frminfo2 As String
    Dim count As Long
    Do
        frmInfo = json_.GetObject(json, "childFrames", count)
        If frmInfo = "" Then Exit Do
        frminfo2 = json_.GetObject(frmInfo, "frame")
        collect.Add frminfo2
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        count = count + 1
    Loop
End Sub
'各フレームのname,urlと指定された値を比較。一致するフレームのIdを返す。
Private Function GetTargetFrameId(frameInfos As Collection, value As Variant) As String
    Dim frameInfo As Variant
    Dim name As String, URL As String
    For Each frameInfo In frameInfos
        name = json_.GetValue(frameInfo, "name")
        URL = json_.GetValue(frameInfo, "url")
        If name = value Or URL = value Then
            GetTargetFrameId = json_.GetValue(frameInfo, "id")
            Exit Function
        End If
    Next
End Function

'Public
'切り替え可能なフレームの情報をイミディエイトウィンドウに表示
Private Sub IWebDriver_DebugPrintAccessibleFrameInfo()
    CDP_.PageEnable
    CDP_.RuntimeEnable
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim i As Long
    For i = 1 To frameInfos.count
        Debug.Print "index：" & i & _
                           "　name：" & json_.GetValue(frameInfos(i), "name") & _
                           "　url：" & json_.GetValue(frameInfos(i), "url")
    Next
End Sub

'Public
Private Sub IWebDriver_SwitchFrameByIframeElement(IframeElement As IWebElement)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim elem As b0x0_WebElement: Set elem = IframeElement
    If elem.FrameId = "" Then
        Helper_.DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
                                      "指定されたElementがIframeではありません" & vbCrLf & vbCrLf & _
                                      "Iframe要素を引数に渡して下さい" & vbCrLf
    End If
    
    CDP_.GetTargetInfo
    CDP_.SetDiscoverTargets

    'frameElementのframeIdが正当なframeIdならIsolatedWorldで対応
    Dim ExecutionContextId As Long
    ExecutionContextId = CDP_.CreateIsolatedFrameWorld(elem.FrameId)
    If ExecutionContextId <> 0 Then
        basicInfo_.ExecutionContextId = ExecutionContextId
        basicInfo_.IsolatedFrameId = elem.FrameId
        Exit Sub
    Else
        Helper_.DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
                                       "このIframeはこのメソッドでは切り替えができません。" & vbCrLf & vbCrLf & _
                                      "※SwitchFrameByNameOrUrlOrIndexの" & vbCrLf & _
                                      "　 使用を検討してください。" & vbCrLf
    End If

End Sub

'Public
'デフォルトのフレームに戻す
Private Sub IWebDriver_SwitchFrameToDefault()
    If basicInfo_.IsolatedFrameId <> "" Then
        basicInfo_.IsolatedFrameId = ""
        basicInfo_.ExecutionContextId = 0
    End If
    
'    If basicInfo_.IsTargetPageDefault = False Then
'        SwitchTargetToPageDefault
'    End If
End Sub


'Public
'第1引数で指定されるXpathの要素有無の確認
Private Function IWebDriver_IsExistsElementByXpath(xPath As String) As Boolean
    '属性指定時に日本語が含まれるとエラーになるため要エンコード
    xPath = json_.EncodeURL(xPath)
    IWebDriver_IsExistsElementByXpath = _
        CDP_.IsExistsElementBy(xPath, basicInfo_, Helper_.GetXpathCodes(xPath))
End Function
'Public
'第1引数で指定されるIdの要素有無の確認
Private Function IWebDriver_IsExistsElementById(Id As String) As Boolean
    IWebDriver_IsExistsElementById = _
        CDP_.IsExistsElementBy(Id, basicInfo_, Helper_.GetIdCodes(Id))
End Function
'Public
'第1引数で指定されるNameの要素有無の確認
Private Function IWebDriver_IsExistsElementByName(name As String) As Boolean
    IWebDriver_IsExistsElementByName = _
        CDP_.IsExistsElementBy(name, basicInfo_, Helper_.GetNameCodes(name))
End Function




'Public
'Xpath(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)
    '属性指定時に日本語が含まれるとエラーになるため要エンコード
    xPath = json_.EncodeURL(xPath)
    
    If CDP_.WaitForElementToAppearBy(xPath, _
        maxWaitTime_sec, basicInfo_, Helper_.GetXpathCodes(xPath)) = False Then
        
        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByXPath>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             xPath
    End If
End Sub
'Public
'Id(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearById(Id As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(Id, _
        maxWaitTime_sec, basicInfo_, Helper_.GetIdCodes(Id)) = False Then
        
        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearById>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             Id
    End If
End Sub
'Public
'Name(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(name, _
        maxWaitTime_sec, basicInfo_, Helper_.GetNameCodes(name)) = False Then
    
        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByName>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             name
    End If
End Sub


'Public
'AlertダイアログのOKボタンを押して閉じる
Private Sub IWebDriver_HandleAlertDialog()
    Dim res As String
    Dim fileSafeCounter As Long
    
    Do
        Helper_.SleepByWinAPI 500
        res = CDP_.HandleAlertDialog
        If InStr(res, "No dialog is showing") = 0 Then
            Helper_.SleepByWinAPI 500
            Exit Sub
        End If
        
        fileSafeCounter = fileSafeCounter + 1
    Loop While fileSafeCounter < 6

quit:
    Helper_.DisplayErrorAndEnd "<HandleAlertDialog>" & vbCrLf & _
                                   "ダイアログが表示されていません。" & vbCrLf & _
                                   "※このメソッドはAlertが表示される場合に使用します。" & vbCrLf
End Sub

'Public
Private Sub IWebDriver_SleepByWinAPI(milliSecond As Long)
    Helper_.SleepByWinAPI milliSecond
End Sub

'Public
'タブの移動（Index指定：一番左のタブを１とする）
Private Sub IWebDriver_SwitchTabByIndex(index As Long)
    Dim res As String
    Dim client As Object: Set client = CreateObject("MSXML2.ServerXMLHTTP")
    client.Open "GET", "http://127.0.0.1:9222/json"
    client.setRequestHeader "Content-Type", "application/json"
    client.send
    res = client.responseText
    Dim count As Long: count = json_.GetJsonLength(res)
    
    Dim tmpCollection As Collection: Set tmpCollection = New Collection
    Dim jsObjStr As String, typeName As String
    Dim tempDic As Object
    Dim i As Long
    For i = 0 To count - 1
        jsObjStr = json_.GetObject(res, i)
        typeName = json_.GetValue(jsObjStr, "type")
        If typeName = "page" Then
            Set tempDic = CreateObject("Scripting.Dictionary")
            tempDic.Add "title", json_.GetValue(jsObjStr, "title")
            tempDic.Add "targetId", json_.GetValue(jsObjStr, "id")
            tempDic.Add "webSocketDebuggerUrl", json_.GetValue(jsObjStr, "webSocketDebuggerUrl")
            tmpCollection.Add tempDic
        End If
    Next
    
    If index > tmpCollection.count Then Err.Raise Number:=10000, description:="指定されたIndexがタブの数を超えています。"
    
    'jsonの情報がタブの順と逆順のため並べ替え
    Dim tabCollection As Collection: Set tabCollection = New Collection
    For i = tmpCollection.count() To 1 Step -1
        tabCollection.Add tmpCollection(i)
    Next
    
     Dim webSocketURLPath As String
     webSocketURLPath = Replace(tabCollection(index)("webSocketDebuggerUrl"), "ws://127.0.0.1:9222", "")
    
    webSocket_.CloseWebSocket
    Set webSocket_ = Nothing
    
    Set webSocket_ = New a1_WebSocketCommunicator
    If webSocket_.Init(webSocketURLPath) = False Then End
    Dim msgGenerator As a4x1_CDPMessageGenerator
    Set msgGenerator = New a4x1_CDPMessageGenerator
    Dim Handler As a4x2_CDPMessageHandler
    Set Handler = New a4x2_CDPMessageHandler
    Handler.Init webSocket_, json_, events_
    Set CDP_ = New a4_ExecuteCDP
    CDP_.Init msgGenerator, Handler, json_
    
    Dim targetId As String
    targetId = tabCollection(index)("targetId")
    
    basicInfo_.Update targetId, json_, True
    
    CDP_.AttachToTarget targetId
    CDP_.ActivateTarget targetId

    basicInfo_.DocumentObjectId = CDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = CDP_.GetDocumentNodeID
End Sub

'Public
Private Sub IWebDriver_CloseWindow()
    CDP_.SessionClose
End Sub

'Public
Private Sub IWebDriver_ScrollWindow(yPixel As Long, Optional xPixel As Long = 0&)
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, xPixel, yPixel
End Sub

'Public
Private Sub IWebDriver_SetForeGroundWindow()
    WaitForActivation
End Sub

'Private Sub IWebDriver_AllowPopUpWindowForURL(targetURL As String)
'    Dim chromeExe As String
'    Dim objExec As Object
'    With CreateObject("WScript.Shell")
'        .RegWrite name:="HKEY_LOCAL_MACHINE\SOFTWARE\Policies\{Google}", _
'                        value:=targetURL, _
'                        Type:="REG_SZ"
'    End With
'End Sub

'Public
Private Sub IWebDriver_Back()
    Dim result As String
    result = CDP_.DispatchMouseEvent("mousePressed", 0, 0, "back", 1)
    result = CDP_.DispatchMouseEvent("mouseReleased", 0, 0, "back", 1)
End Sub

'Public
Private Sub IWebDriver_Forward()
    Dim result As String
    result = CDP_.DispatchMouseEvent("mousePressed", 0, 0, "forward", 1)
    result = CDP_.DispatchMouseEvent("mouseReleased", 0, 0, "forward", 1)
End Sub

'Private Sub IWebDriver_PointerMove(viewportX As Long, viewportY As Long)
'    Dim result As String
'    result = CDP_.DispatchMouseEvent("mouseMoved", viewportX, viewportY, "none", 0)
'End Sub

'Public
' Screenshotを取り、画像ファイルとして保存し、保存先パスを返す。
Private Function IWebDriver_ScreenShotSaveAsFile( _
    Optional folderPath As String = "", _
    Optional fileName As String = "", _
    Optional imageFormat As Image = Image.png) As String
    
    ' 保存フォルダ指定ありの場合、存在チェック
    If folderPath <> "" Then
        If Dir(folderPath, vbDirectory) = "" Then Helper_.DisplayErrorAndEnd "指定されたフォルダが存在しません"
    End If
    ' ファイル名指定ありの場合、NG文字が含まれていないかチェック
    If fileName <> "" Then
        If Helper_.IncludeNGCharInFileName(fileName) Then Helper_.DisplayErrorAndEnd "指定されたファイル名に使用不可の文字が含まれています"
    End If
    
    ' Screenshot取得し、Base64デコード
    Dim formatName As String: formatName = Helper_.GetFormatName(imageFormat)
    Dim base64EncodedImageData As String, decodedImageData As Variant
    base64EncodedImageData = CDP_.CaptureScreenshot(formatName)
    decodedImageData = Helper_.DecodeBase64(base64EncodedImageData)
    
    ' 画像ファイルとして保存
    If folderPath = "" Then
        folderPath = ThisWorkbook.Path
    Else
        If Right(folderPath, 1) = "\" Then folderPath = left(folderPath, Len(folderPath) - 1)
    End If
    If fileName = "" Then fileName = "screenshot" & Helper_.GetFileNumber(folderPath, formatName)
    Dim fullPath As String: fullPath = folderPath & "\" & fileName & "." & formatName
    With CreateObject("ADODB.Stream")
        .Type = 1  ' 1 = adTypeBinary
        .Open
        .Write decodedImageData
        .SaveToFile fullPath, 2  ' 2= adSaveCreateOverWrite
        .Close
    End With
    
    IWebDriver_ScreenShotSaveAsFile = fullPath
End Function




' Public
' Screenshotを取り、ワークシートに貼り付け。
Public Function IWebDriver_ScreenShotPasteToSheet( _
    left As Single, top As Single, _
    Optional width As Single = 0, Optional height As Single = 0, _
    Optional sheetName As String = "", _
    Optional imageFormat As Image = Image.png) As Shape

    If sheetName <> "" And Helper_.ExistsSheetName(sheetName) = False Then
        Helper_.DisplayErrorAndEnd "指定されたシート名のワークシートが存在しません"
    End If
    
    
    Dim imageFilePath As String
    imageFilePath = IWebDriver_ScreenShotSaveAsFile(, , imageFormat)
    Dim shp As Shape
    If sheetName <> "" Then
        Set shp = ThisWorkbook.Worksheets(sheetName).Shapes.AddPicture(imageFilePath, False, True, left, top, width, height)
    Else
        Set shp = ThisWorkbook.ActiveSheet.Shapes.AddPicture(imageFilePath, False, True, left, top, width, height)
    End If
    
    If width = 0 Then shp.ScaleWidth 1#, msoTrue
    If height = 0 Then shp.ScaleHeight 1#, msoTrue
 
    Kill imageFilePath
    
    Set IWebDriver_ScreenShotPasteToSheet = shp
End Function

' Public
' Windowサイズ変更
Private Sub IWebDriver_ShowWindow(size As WindowSize)
    Select Case size
        Case WindowSize.normal
            ShowWindow chromeHwnd_, WindowSize.normal
        Case WindowSize.minimize
            ShowWindow chromeHwnd_, WindowSize.minimize
        Case WindowSize.maximize
            ShowWindow chromeHwnd_, WindowSize.maximize
    End Select
End Sub

' Public
' 暗号論的擬似乱数を取得
Private Function IWebDriver_GetSecureRandomNumber(count As Long) As Variant
    If count = 0 Then
        IWebDriver_GetSecureRandomNumber = Nothing
        Exit Function
    End If

    Dim randomNums As String: randomNums = CDP_.GetSecureRondomNumber(count)
    
    Dim result As Collection: Set result = New Collection
    If count = 1 Then
        result.Add CLngLng(randomNums)
    Else
        Dim tmpArray As Variant: tmpArray = Split(randomNums, ",")
        Dim i As Long
        For i = LBound(tmpArray) To UBound(tmpArray)
            result.Add CLngLng(tmpArray(i))
        Next
    End If
    
    Set IWebDriver_GetSecureRandomNumber = result
End Function

' Public
' 暗号論的擬似乱数を元にランダム文字列を取得
Private Function IWebDriver_GetURLSafeRundomString(charNum As Long) As String
    Dim chars As String: chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"
    Dim length As Long: length = Len(chars)
    Dim charsArray() As Variant: ReDim charsArray(length - 1)
    
    Dim i As Long
    For i = 1 To length
        charsArray(i - 1) = Mid(chars, i, 1)
    Next
    
    Dim rondomNums As Collection: Set rondomNums = IWebDriver_GetSecureRandomNumber(charNum)
    Dim rondomString As String
    Dim rondomNum As Variant
    Dim index As Long
    For Each rondomNum In rondomNums
        index = rondomNum Mod (length)
        rondomString = rondomString & charsArray(index)
    Next
    IWebDriver_GetURLSafeRundomString = rondomString
End Function

' Public
' Base64URLエンコードされたSHA256の値を取得
Private Function IWebDriver_GetBase64URLEncodedSHA256(str As String) As String
    Dim base64EncodedSHA256 As String: base64EncodedSHA256 = CDP_.GetBase64EncodedSHA256(str)
    
    Dim URLEncoded As String
    URLEncoded = Replace(base64EncodedSHA256, "=", "")
    URLEncoded = Replace(URLEncoded, "+", "-")
    URLEncoded = Replace(URLEncoded, "/", "_")
    IWebDriver_GetBase64URLEncodedSHA256 = URLEncoded
End Function

' Public
' URIコンポーネントエンコードした値を取得
Private Function IWebDriver_EncodeURIConpornent(str As String) As String
    IWebDriver_EncodeURIConpornent = json_.EncodeURLCompornent(str)
End Function



