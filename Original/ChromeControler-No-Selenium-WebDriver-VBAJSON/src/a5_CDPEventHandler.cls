VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a5_CDPEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos

''ダイアログ情報
Public DialogInfoDic As Object
Public PageInfos As Collection
Public IframeInfos As Collection

Public IsCurrentTargetScheduledNavigation As Boolean
Public IsCurrentTargetStartedLoading As Boolean
Public IsCurrentTargetClearedScheduledNavigation As Boolean
Public IsCurrentTargetNavigated As Boolean

Public IsIsolatedFrameScheduledNavigation As Boolean
Public IsIsolatedFrameStartedLoading As Boolean
Public IsIsolatedFrameClearedScheduledNavigation As Boolean
Public IsIsolatedFrameNavigated As Boolean

Public IsCurrentContextIdDestroyed As Boolean



Public Sub Init(json As a2_JSONHandler, basicInfo As a3_BasicInfos)
    Set json_ = json
    Set basicInfo_ = basicInfo
    
    Set DialogInfoDic = CreateObject("Scripting.Dictionary")
    InitializeDialogInfo
    Set PageInfos = New Collection
    Set IframeInfos = New Collection
End Sub

'各種イベント発生時に情報を取得する
Public Sub GetInfo(EventInfo As String)
    Dim eventName As String
    eventName = json_.GetValue(EventInfo, "method")
    
    If InStr(eventName, "frame") > 0 Then
        Dim Id As String
        If InStr(eventName, "Navigated") = 0 Then
            Id = json_.GetValue(EventInfo, "params", "frameId")
        Else
            Id = json_.GetValue(EventInfo, "params", "frame", "id")
        End If
        Dim isCurrentTarget As Boolean
        isCurrentTarget = (Id = basicInfo_.targetId)
        Dim isIsolatedFrameId
        isIsolatedFrameId = (Id = basicInfo_.IsolatedFrameId)
    End If
    
    Select Case eventName
        Case "Page.javascriptDialogOpening"
            SetDialogInfo EventInfo
            Debug.Print "★★DialogOpening"
        Case "Page.javascriptDialogClosed"
            InitializeDialogInfo
            Debug.Print "★★DialogClosed"
            
        Case "Target.targetCreated"
            AddTargetInfo EventInfo
        Case "Target.targetInfoChanged"
            ChangeTargetInfo EventInfo
        Case "Target.targetDestroyed"
            RemoveTargetInfo EventInfo
            
        Case "Page.frameScheduledNavigation" '非推奨となっているがこれが返ってくる。
            If InStr(EventInfo, "anchorClick") > 0 Then
                If isCurrentTarget Then IsCurrentTargetScheduledNavigation = True
                If isIsolatedFrameId Then
                    IsIsolatedFrameScheduledNavigation = True
                    Debug.Print "★★IsolatedFrameScheduledNavigation"
                End If
            End If
        Case "Page.frameStartedLoading"
            If isCurrentTarget Then IsCurrentTargetStartedLoading = True
            If isIsolatedFrameId Then
                IsIsolatedFrameStartedLoading = True
                Sleep 1000
            End If
'        Case "Page.frameClearedScheduledNavigation" '非推奨となっているがこれが返ってくる。
'            If isCurrentTarget Then IsCurrentTargetClearedScheduledNavigation = True
'            If isIsolatedFrameId Then IsIsolatedFrameClearedScheduledNavigation = True
        Case "Page.frameNavigated"
            If isCurrentTarget Then IsCurrentTargetNavigated = True
            If isIsolatedFrameId Then
                IsIsolatedFrameNavigated = True
                Debug.Print "★★IsolatedFrameNavigated"
            End If
        Case "Runtime.executionContextDestroyed"
            Dim contextId As Long
            contextId = json_.GetValue(EventInfo, "params", "executionContextId")
            If contextId = basicInfo_.ExecutionContextId Then
                IsCurrentContextIdDestroyed = True
                Debug.Print "★★IsCurrentContextIdDestroyed"
            End If
    End Select
End Sub
'アラートダイアログ出現時に情報を取得
Private Sub SetDialogInfo(ByVal EventInfo As String)
    DialogInfoDic("IsExistDialog") = True
    DialogInfoDic("DialogMessage") = json_.GetValue(EventInfo, "params", "message")
    DialogInfoDic("DialogType") = json_.GetValue(EventInfo, "params", "type")
    If json_.GetValue(EventInfo, "params", "hasBrowserHandler") = "True" Then
        DialogInfoDic("HasBrowserHandler") = True
    Else
        DialogInfoDic("HasBrowserHandler") = False
    End If
    DialogInfoDic("DefaultPrompt") = json_.GetValue(EventInfo, "params", "defaultPrompt")
End Sub
'アラートダイアログ出現時に情報を格納する辞書の初期化
Private Sub InitializeDialogInfo()
    DialogInfoDic.RemoveAll
    DialogInfoDic.add "IsExistDialog", False
    DialogInfoDic.add "DialogMessage", ""
    DialogInfoDic.add "DialogType", ""
    DialogInfoDic.add "HasBrowserHandler", False
    DialogInfoDic.add "DefaultPrompt", ""
End Sub

'TargetのTypeにより､それぞれコレクションに情報を取得する｡
Private Sub AddTargetInfo(ByVal EventInfo As String)
    Dim tempDic As Object
    Set tempDic = CreateObject("Scripting.Dictionary")
    tempDic.add "targetId", json_.GetValue(EventInfo, "params", "targetInfo", "targetId")
    tempDic.add "title", json_.GetValue(EventInfo, "params", "targetInfo", "title")
    tempDic.add "url", json_.GetValue(EventInfo, "params", "targetInfo", "url")
    tempDic.add "browserContextId", json_.GetValue(EventInfo, "params", "targetInfo", "browserContextId")
    
    Dim targetType As String
    targetType = json_.GetValue(EventInfo, "params", "targetInfo", "type")
    Select Case targetType
      Case "page"
        AddInfoOf PageInfos, tempDic, EventInfo
      Case "iframe"
        AddInfoOf IframeInfos, tempDic, EventInfo
    End Select
End Sub
'Chrome上でTargetが追加されたら、各TargetTypeの情報を管理するCollectionにその情報をDictionaryとしてまとめて追加する。
Private Sub AddInfoOf(infos As Collection, tempDic As Object, _
                                    ByVal EventInfo As String)
                                        
    '本来targetInfoChangedで来るべき内容がtargetCreatedで来るケースあり。
    'SetDiscoverTarget使用すると既に開かれているPageもCreateされる模様
    Dim i As Long
    For i = 1 To infos.count
        If infos(i)("targetId") = tempDic("targetId") Then
            infos(i)("title") = tempDic("targetId")
            infos(i)("url") = tempDic("url")
            infos(i)("browserContextId") = tempDic("browserContextId")
            Exit Sub
        End If
    Next
    infos.add tempDic
End Sub

'Chrome上でPage(タブ)の情報が変更されたら、Page情報を管理するCollection内の該当Pageの情報を変更する。
Private Sub ChangeTargetInfo(ByVal EventInfo As String)
    Dim targetType As String
    targetType = json_.GetValue(EventInfo, "params", "targetInfo", "type")
    
    Select Case targetType
      Case "page"
        ChangeInfoOf PageInfos, EventInfo
      Case "iframe"
        ChangeInfoOf IframeInfos, EventInfo
    End Select
End Sub
Private Sub ChangeInfoOf(infos As Collection, _
                                        ByVal EventInfo As String)
    Dim targetId As String
    targetId = json_.GetValue(EventInfo, "params", "targetInfo", "targetId")
    Dim info As Variant
    For Each info In infos
        If info("targetId") = targetId Then
            info("title") = json_.GetValue(EventInfo, "params", "targetInfo", "title")
            info("url") = json_.GetValue(EventInfo, "params", "targetInfo", "url")
            info("browserContextId") = json_.GetValue(EventInfo, "params", "targetInfo", "browserContextId")
        End If
    Next
End Sub

'Chrome上でPage(タブ)が削除されたら、Page情報を管理するCollectionからその情報を削除する。
Private Sub RemoveTargetInfo(ByVal EventInfo As String)
    Dim targetType As String
    targetType = json_.GetValue(EventInfo, "params", "targetInfo", "type")
    
    Select Case targetType
      Case "page"
        RemoveInfoOf PageInfos, EventInfo
      Case "iframe"
        RemoveInfoOf IframeInfos, EventInfo
    End Select
End Sub
Private Sub RemoveInfoOf(infos As Collection, _
                                        ByVal EventInfo As String)
    Dim targetId As String
    targetId = json_.GetValue(EventInfo, "params", "targetInfo", "targetId")
    
    Dim i As Long
    For i = 1 To infos.count
        If infos(i)("targetId") = targetId Then
            infos.Remove i
        End If
    Next
End Sub

Public Sub ForDebug_ShowtargetInfo()
    Dim i  As Long
    Debug.Print "PageInfos↓==============="
    For i = 1 To PageInfos.count
        Debug.Print i & "：" & PageInfos(i)("targetId")
    Next
    Debug.Print "IframeInfos↓==============="
    For i = 1 To IframeInfos.count
        Debug.Print i & "：" & IframeInfos(i)("targetId")
    Next
    Debug.Print "========================"
End Sub
