Attribute VB_Name = "Kit_CustomsReplay"
Option Explicit

' StarterWebScrapingKit / CDP replay for the 2026-10-18 Discovery Log.
' Entry point: RunCustomsDiscoveryWithKit
' This module is imported into the ORIGINAL StarterWebScrapingKit workbook.
' It does not alter the Kit's classes or its settings sheet.
' Browser intentionally stays open so the user can check the result.
' Warning: This is a source-based replay; it has NOT been run against a live browser.

Private Const START_URL As String = "https://www.customs.go.jp/toukei/srch/index.htm?M=01&P=0"
Private Const POPUP_SECONDS As Double = 12#
Private Const ELEMENT_SECONDS As Double = 12#
Private Const RESULT_SECONDS As Double = 20#

Public Sub RunCustomsDiscoveryWithKit()
    Dim mainContext As CDPContext
    Dim popup As CDPContext
    Dim root As CDPBrowser
    Dim radio As CDPElement
    Dim browseButton As CDPElement
    Dim groupList As CDPElement
    Dim confirmButton As CDPElement
    Dim searchButton As CDPElement
    Dim seen As Collection
    Dim popupId As String
    Dim sessionId As String
    Dim optionValue As String
    Dim selectedLabel As String
    Dim stage As String

    On Error GoTo ReplayError
    stage = "Browser startup"
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " | START | "; START_URL

    ' Uses the browser profile/options on the Kit's settings sheet.
    Set mainContext = ShSetting01_StartBrowser.StartCDPModeContext(START_URL)
    Set root = mainContext.ThisCDPBrowser

    stage = "Select Import radio"
    Set radio = WaitFrameElement(mainContext, "input#la_imp[name='OptExport']", _
                                  ELEMENT_SECONDS)
    ClickByCDP mainContext, radio
    If Not CBool(radio.jsEval("function(){return this.checked}")) Then _
        Fail "Import radio did not become checked."
    Debug.Print "[OK] Import radio selected: #la_imp"

    stage = "Open HS Group popup"
    ' Take the page-target baseline before clicking.
    Set seen = SnapshotPageIds(root)

    ' sibling=3/3 is recorded in the log (3rd INPUT among 3 INPUT siblings).
    ' Never click an arbitrary 'Browse' button if its position no longer matches.
    Set browseButton = WaitFrameElement(mainContext, "input[type='button'][value='éQè∆']", _
                                       ELEMENT_SECONDS, 3, 3)
    ClickByCDP mainContext, browseButton

    popupId = WaitNewPopupTarget(root, mainContext.CurrentTargetID, seen, POPUP_SECONDS)
    If LenB(popupId) = 0 Then Fail "No newly opened popup target was found."
    sessionId = root.attachToTab(popupId)
    If LenB(sessionId) = 0 Then Fail "Target.attachToTarget failed for the popup."

    ' Same attach flow as CDPBrowser.getTab, but with an exact targetId.
    root.EnableDiscoverTargets = True
    Set popup = New CDPContext
    popup.Init root, root.ExEventCore
    popup.CurrentSessionID = sessionId
    Debug.Print "[OK] Popup attached, targetId="; popupId

    stage = "Choose HS Group"
    Set groupList = WaitFrameElement(popup, "select[name='LstHSGroup']", _
                                     ELEMENT_SECONDS)
    ' The Discovery Log exposes OPTION TEXT but redacts the option VALUE.
    ' Obtain the *current* value from the live option whose text matches.
    optionValue = GetGroupValueByVisibleText(groupList, "ÇOÇUóﬁÅ`ÇPÇTóﬁ")
    If LenB(optionValue) = 0 Then _
        Fail "HS Group 06-15 option has no matching live value."
    groupList.setSelection optionValue
    selectedLabel = CStr(groupList.jsEval( _
                    "function(){return this.selectedOptions.length?this.selectedOptions[0].text:''}"))
    If InStr(selectedLabel, "ÇOÇUóﬁÅ`ÇPÇTóﬁ") = 0 Then _
        Fail "Selected HS Group does not match the Discovery Log: " & selectedLabel
    Debug.Print "[OK] HS Group: "; selectedLabel

    stage = "Confirm HS Group"
    Set confirmButton = WaitFrameElement(popup, _
        "form[name='form01c'] input[type='button'][value='åàíË']", ELEMENT_SECONDS, 1, 2)
    ClickByCDP popup, confirmButton
    Debug.Print "[OK] Confirm clicked"

    stage = "Submit Search"
    ' The original frame can be re-written after closing the popup.
    ' Re-find the button rather than reusing an old CDP objectId.
    Set searchButton = WaitFrameElement(mainContext, "input[type='submit'][value*='åü']", _
                                        ELEMENT_SECONDS, 1, 2)
    ' The log's value contains extra spaces: ' åü  çı '.
    If CStr(searchButton.jsEval( _
        "function(){return this.value.replace(/[\s\u00a0]/g,'')}")) <> "åüçı" Then _
        Fail "The intended Search button could not be verified."
    ClickByCDP mainContext, searchButton
    Debug.Print "[OK] Search clicked"

    stage = "Validate result navigation"
    If Not WaitForResultFrame(mainContext, RESULT_SECONDS) Then _
        Fail "Result frame /JCWSV01/servlet/JCWSV01 was not observed."
    Debug.Print "[OK] Result frame reached: /JCWSV01/servlet/JCWSV01"
    Debug.Print "[DONE] Replay reached the observed result URL. Verify result contents manually."
    Exit Sub

ReplayError:
    Debug.Print "[ERROR] "; stage; " | "; Err.Number; " | "; Err.Description
    Debug.Print "The browser is left open for inspection."
    MsgBox "Kitçƒé¿çsÇÕÅu" & stage & "ÅvÇ≈í‚é~ÇµÇ‹ÇµÇΩÅB" & vbCrLf & _
           Err.Description & vbCrLf & "è⁄ç◊ÇÕVBEÇÃÉCÉ~ÉfÉBÉGÉCÉgÉEÉBÉìÉhÉE(Ctrl+G)ÇämîFÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", _
           vbExclamation, "Kit Customs Replay"
End Sub

' Search *same-origin* frame and iframe documents, including old FRAMESET trees.
' A CDPElement is constructed from the returned Runtime.RemoteObject objectId.
' siblingOrdinal/siblingTotal are optional safety guards from the Discovery Log.
Private Function FindFrameElement(ByVal mainContext As CDPContext, ByVal selector As String, _
                                  Optional ByVal siblingOrdinal As Long = 0, _
                                  Optional ByVal siblingTotal As Long = 0) As CDPElement
    Dim js As String
    Dim obj As Variant
    Dim item As CDPElement
    Dim validRule As String

    If siblingOrdinal > 0 Then
        validRule = "var a=Array.prototype.filter.call(e.parentElement.children," & _
            "function(n){return n.tagName===e.tagName});" & _
            "if(a.length!==" & CStr(siblingTotal) & "||a.indexOf(e)!==" & _
            CStr(siblingOrdinal - 1) & ")continue;"
    Else
        validRule = vbNullString
    End If

    js = "(function(){" & _
         "function scan(d){" & _
         "var a=d.querySelectorAll(" & JsLiteral(selector) & ");" & _
         "for(var i=0;i<a.length;i++){var e=a[i];" & validRule & "return e;}" & _
         "var f=d.querySelectorAll('frame,iframe');" & _
         "for(var j=0;j<f.length;j++){try{var q=f[j].contentDocument;" & _
         "if(q){var x=scan(q);if(x)return x;}}catch(ignore){}}" & _
         "return null;}return scan(document);})()"

    obj = mainContext.jsEval(js, IFEXCEPTION:=Null)
    If VarType(obj) = vbString Then
        Set item = New CDPElement
        item.Init mainContext, vbString, CStr(obj), js
        Set FindFrameElement = item
    End If
End Function

Private Function WaitFrameElement(ByVal mainContext As CDPContext, ByVal selector As String, _
                                  ByVal seconds As Double, _
                                  Optional ByVal siblingOrdinal As Long = 0, _
                                  Optional ByVal siblingTotal As Long = 0) As CDPElement
    Dim startMs As Double
    Dim item As CDPElement
    startMs = CDPHelpers.TimerCounter

    Do
        On Error Resume Next
        Set item = FindFrameElement(mainContext, selector, siblingOrdinal, siblingTotal)
        If Err.Number <> 0 Then
            Err.Clear
            Set item = Nothing
        End If
        On Error GoTo 0
        If Not item Is Nothing Then
            If item.isExist Then
                Set WaitFrameElement = item
                Exit Function
            End If
        End If
        CDPHelpers.Sleep 0.15
    Loop While CDPHelpers.TimerCounter - startMs < seconds * 1000#
    Fail "Timed out locating: " & selector & _
         " (sibling " & siblingOrdinal & "/" & siblingTotal & ")"
End Function

' Send actual CDP mouse input, not the Kit's synthetic MouseEvent-dispatch click.
' Unlike CDPElement.sendClick, account for nested frame origins when computing XY.
' Only works for same-origin frames without CSS transforms/zoom complications.
Private Sub ClickByCDP(ByVal mainContext As CDPContext, ByVal item As CDPElement)
    Dim x As Double, y As Double
    Dim p As Dictionary
    Dim baseJs As String

    item.jsEval "function(){this.scrollIntoView({block:'center',inline:'center'});" & _
                "var w=this.ownerDocument.defaultView;" & _
                "while(w!==w.top){var f=w.frameElement;if(!f)break;" & _
                "f.scrollIntoView({block:'center',inline:'center'});w=w.parent;}}"

    baseJs = "function(){var r=this.getBoundingClientRect()," & _
             "w=this.ownerDocument.defaultView,v=r.%AXIS%+r.%SIZE%/2;" & _
             "while(w!==w.top){var f=w.frameElement;if(!f)return -1;" & _
             "var b=f.getBoundingClientRect();v+=b.%AXIS%+f.%BORDER%;w=w.parent;}return v;}"
    x = CDbl(item.jsEval(Replace(Replace(Replace(baseJs, _
          "%AXIS%", "left"), "%SIZE%", "width"), "%BORDER%", "clientLeft")))
    y = CDbl(item.jsEval(Replace(Replace(Replace(baseJs, _
          "%AXIS%", "top"), "%SIZE%", "height"), "%BORDER%", "clientTop")))
    If x < 0 Or y < 0 Then Fail "Could not compute page-space mouse coordinates."

    Set p = New Dictionary
    p("type") = "mousePressed"
    p("x") = x
    p("y") = y
    p("button") = "left"
    p("clickCount") = 1
    mainContext.ExecuteCDP "Input.dispatchMouseEvent", p
    p("type") = "mouseReleased"
    mainContext.ExecuteCDP "Input.dispatchMouseEvent", p
End Sub

Private Function SnapshotPageIds(ByVal br As CDPBrowser) As Collection
    Dim result As BiDiCDPJson, infos As BiDiCDPJson
    Dim token As Long
    Dim ids As New Collection
    Set result = br.ExecuteCDP("Target.getTargets")
    Set infos = result.NodeKey("targetInfos")
    token = infos.FirstChildToken
    Do While token > 0
        If infos.TokenString(token, "type") = "page" Then _
            ids.Add infos.TokenString(token, "targetId")
        token = infos.NextToken(token)
    Loop
    Set SnapshotPageIds = ids
End Function

Private Function CollectionHasId(ByVal ids As Collection, ByVal targetId As String) As Boolean
    Dim obj As Variant
    For Each obj In ids
        If CStr(obj) = targetId Then
            CollectionHasId = True
            Exit Function
        End If
    Next obj
End Function

Private Function WaitNewPopupTarget(ByVal br As CDPBrowser, _
                                    ByVal openerId As String, _
                                    ByVal seen As Collection, _
                                    ByVal seconds As Double) As String
    Dim t0 As Double
    Dim result As BiDiCDPJson, infos As BiDiCDPJson
    Dim token As Long
    Dim id As String, opener As String
    Dim candidates As Long, candidate As String
    Dim direct As String
    t0 = CDPHelpers.TimerCounter
    Do
        candidates = 0
        candidate = vbNullString
        direct = vbNullString
        Set result = br.ExecuteCDP("Target.getTargets")
        Set infos = result.NodeKey("targetInfos")
        token = infos.FirstChildToken
        Do While token > 0
            If infos.TokenString(token, "type") = "page" Then
                id = infos.TokenString(token, "targetId")
                If Not CollectionHasId(seen, id) Then
                    candidates = candidates + 1
                    candidate = id
                    opener = infos.TokenString(token, "openerId")
                    If opener = openerId Then direct = id
                End If
            End If
            token = infos.NextToken(token)
        Loop
        If LenB(direct) Then
            WaitNewPopupTarget = direct
            Exit Function
        End If
        If candidates = 1 Then
            WaitNewPopupTarget = candidate
            Exit Function
        End If
        If candidates > 1 Then Fail "Ambiguous popup: multiple new page targets."
        CDPHelpers.Sleep 0.15
    Loop While CDPHelpers.TimerCounter - t0 < seconds * 1000#
End Function

' Use the option's observed DISPLAY TEXT. Never invent a redacted value.
Private Function GetGroupValueByVisibleText(ByVal listElement As CDPElement, _
                                              ByVal partialLabel As String) As String
    Dim js As String
    Dim v As Variant
    js = "function(){var opts=Array.prototype.filter.call(this.options," & _
         "function(o){return o.text.indexOf(" & JsLiteral(partialLabel) & ")>=0});" & _
         "return opts.length===1?opts[0].value:'';}"
    v = listElement.jsEval(js)
    If Not IsNull(v) And Not IsError(v) Then GetGroupValueByVisibleText = CStr(v)
End Function

Private Function WaitForResultFrame(ByVal mainContext As CDPContext, _
                                    ByVal seconds As Double) As Boolean
    Dim js As String
    Dim startMs As Double
    Dim v As Variant
    js = "(function(){function find(d){" & _
         "try{if(d.location.pathname.toLowerCase().indexOf(" & _
         "'/jcwsv01/servlet/jcwsv01')>=0)return true;}catch(ignore){}" & _
         "var f=d.querySelectorAll('frame,iframe');" & _
         "for(var i=0;i<f.length;i++){try{if(f[i].contentDocument&&" & _
         "find(f[i].contentDocument))return true;}catch(ignore){}}" & _
         "return false;}return find(document);})()"
    startMs = CDPHelpers.TimerCounter
    Do
        On Error Resume Next
        v = mainContext.jsEval(js, IFEXCEPTION:=False)
        If Err.Number <> 0 Then
            Err.Clear
            v = False
        End If
        On Error GoTo 0
        If VarType(v) = vbBoolean Then
            If CBool(v) Then
                WaitForResultFrame = True
                Exit Function
            End If
        End If
        CDPHelpers.Sleep 0.25
    Loop While CDPHelpers.TimerCounter - startMs < seconds * 1000#
End Function

Private Function JsLiteral(ByVal value As String) As String
    Dim out As String
    out = Replace(value, "\", "\\")
    out = Replace(out, "'", "\'")
    out = Replace(out, vbCr, "\r")
    out = Replace(out, vbLf, "\n")
    JsLiteral = "'" & out & "'"
End Function

Private Sub Fail(ByVal why As String)
    Err.Raise vbObjectError + 2026, "Kit_CustomsReplay", why
End Sub
