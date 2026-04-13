VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "exCDP_CookieAndDOMStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           `CDPBrowser.cls` に対する機能拡張 - Cookie / DOMStorage（Local / Session）
'***************************************************************************************************
'* [概要]
'*   CDP の `Storage` ドメイン（Cookie）および `DOMStorage` ドメイン（localStorage / sessionStorage）
'*   を、マクロ側から `invokeMethod` / `jsEval` を直接書かずに操作するためのラッパーです。
'*
'* [できること]
'*   - Cookie: 全削除、1件設定、一覧取得（`Storage.getCookies` の `cookies` に相当する Object）
'*   - localStorage / sessionStorage: 全削除、キー設定・取得・1件削除、全件取得（`entries` の Object）
'*
'* [返り値の形]
'*   - `GetCookies` / `LocalStorageGetAll` / `SessionStorageGetAll` は **JSON 文字列にせず**
'*     CDP 応答をパースしたままの Object（通常は `Collection`）を返します。
'*     ログ出力やシリアライズは呼び出し側で `WebJsonConverter.ConvertToJson` 等を使ってください。
'*
'* [使い方]
'*    Dim st As New exCDP_CookieAndDOMStorage
'*    st.Init browser
'*    Dim cookies As Object: Set cookies = st.GetCookies
'*
'* [注意事項]
'*   - 対象は現在の `CDPBrowser` セッションがアタッチしているフレーム／タブのオリジンです。
'*   - `securityOrigin` は `location.origin`（`jsEval`）で取得します（内部実装のみの利用）。
'*   - `Storage.getStorageKey` が失敗する環境では `storageKey` を空文字として `DOMStorage` に渡します。
'*   - 本クラスは非同期イベントを利用しないため、既定で `EnableEvents=False` です。
'***************************************************************************************************
Option Explicit


'***************************************************************************************************
'                               ■■■ 共通定義変数/定数 ■■■
'***************************************************************************************************
Private WithEvents ex_CDPBrowser    As CDPBrowser
Attribute ex_CDPBrowser.VB_VarHelpID = -1

Private Const ThisClassName         As String = "exCDP_CookieAndDOMStorage"

Private seeRawSendMsgDbg            As Boolean
Private EnableBrowserEvents         As Boolean

Private mDomStorageEnabled          As Boolean



'***************************************************************************************************
'                               ■■■ 公開プロパティ ■■■
'***************************************************************************************************
'* 機能　　：本拡張が `detectionCDPEvent` を処理するか（現状は未使用のため既定 False）
'***************************************************************************************************
Public Property Get EnableEvents() As Boolean
    EnableEvents = EnableBrowserEvents
End Property

Public Property Let EnableEvents(ByVal v As Boolean)
    EnableBrowserEvents = v
End Property



'***************************************************************************************************
'                     ■■■ 非同期イベント（テンプレート互換・現状未使用） ■■■
'***************************************************************************************************
Private Sub ex_CDPBrowser_detectionCDPEvent(methodName As String, EventRawJson As String)
    Const FromProcedureName As String = ThisClassName & ".ex_CDPBrowser_detectionCDPEvent"


    If Not EnableBrowserEvents Then
        Exit Sub
    End If

    printMsg Trace_, "CookieAndDOMStorage: event ignored (no handlers): " & methodName, FromProcedureName
End Sub



'***************************************************************************************************
'                           ■■■ Cookie（Storage ドメイン） ■■■
'***************************************************************************************************
'* 機能　　：ブラウザに保持されている Cookie をすべて削除します（`Storage.clearCookies`）
'***************************************************************************************************
Public Sub ClearCookies()
    Const FromProcedureName As String = ThisClassName & ".ClearCookies"


    printMsg Debug_, "Invoking Storage.clearCookies ...", FromProcedureName
    cdpInvoke "Storage.clearCookies"
    printMsg info_, "Storage.clearCookies completed.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：Cookie を1件設定します（`Storage.setCookies`）
'---------------------------------------------------------------------------------------------------
'* 引数　　：`url` または `domain` のいずれかを指定するとスコープが明確になります。
'*            `expires` は 0 のとき省略します（CDP の CookieParam に含めません）。
'* 返り値　：コマンドが発行できれば True（CDP エラー時は `invokeMethod` 側で例外）
'***************************************************************************************************
Public Function SetCookie( _
        ByVal Name As String, _
        ByVal value As String, _
        Optional ByVal Url As String, _
        Optional ByVal domain As String, _
        Optional ByVal Path As String = "/", _
        Optional ByVal Secure As Boolean, _
        Optional ByVal httpOnly As Boolean, _
        Optional ByVal sameSite As String, _
        Optional ByVal expires As Double _
    ) As Boolean

    Const FromProcedureName As String = ThisClassName & ".SetCookie"


    Dim one As New Scripting.Dictionary
    one.Add "name", Name
    one.Add "value", value
    If Len(Url) > 0 Then one.Add "url", Url
    If Len(domain) > 0 Then one.Add "domain", domain
    If Len(Path) > 0 Then one.Add "path", Path
    one.Add "secure", Secure
    one.Add "httpOnly", httpOnly
    If Len(sameSite) > 0 Then one.Add "sameSite", sameSite
    If expires <> 0# Then one.Add "expires", expires

    Dim arr As New Collection
    arr.Add one

    Dim params As New Scripting.Dictionary
    params.Add "cookies", arr

    printMsg Debug_, "Storage.setCookies: name=" & Name & " url=" & Url, FromProcedureName
    cdpInvoke "Storage.setCookies", params
    printMsg info_, "Storage.setCookies completed for cookie name=" & Name, FromProcedureName
    SetCookie = True
End Function

'***************************************************************************************************
'* 機能　　：現在のブラウジングコンテキストの Cookie 一覧を Object で返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`Storage.getCookies` の結果ディクショナリ内 `cookies` キーに対応するオブジェクト
'*            （通常は `Collection`。取得失敗時やキー欠落時は `Nothing`）
'---------------------------------------------------------------------------------------------------
'* 注意事項：JSON 文字列化は行いません。表示用は呼び出し側で `WebJsonConverter.ConvertToJson` を使用してください。
'***************************************************************************************************
Public Function GetCookies() As Object
    Const FromProcedureName As String = ThisClassName & ".GetCookies"


    Dim res As Scripting.Dictionary
    Set res = cdpInvoke("Storage.getCookies")

    If res Is Nothing Then
        printMsg WARN_, "Storage.getCookies returned no result dictionary.", FromProcedureName
        Exit Function
    End If

    If Not res.Exists("cookies") Then
        printMsg WARN_, "Storage.getCookies: result has no 'cookies' key.", FromProcedureName
        Exit Function
    End If

    Set GetCookies = res("cookies")
    printMsg Debug_, "GetCookies -> TypeName=" & TypeName(GetCookies), FromProcedureName
End Function



'***************************************************************************************************
'                           ■■■ localStorage（DOMStorage） ■■■
'***************************************************************************************************
'* 機能　　：localStorage を空にします（`DOMStorage.clear`）
'***************************************************************************************************
Public Sub LocalStorageClear()
    Const FromProcedureName As String = ThisClassName & ".LocalStorageClear"


    EnsureDOMStorageEnabled FromProcedureName
    cdpInvoke "DOMStorage.clear", BuildStorageParams(True)
    printMsg info_, "localStorage cleared.", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：localStorage にキーと値を設定します（`DOMStorage.setDOMStorageItem`）
'***************************************************************************************************
Public Sub LocalStorageSetItem(ByVal Key As String, ByVal value As String)
    Const FromProcedureName As String = ThisClassName & ".LocalStorageSetItem"


    EnsureDOMStorageEnabled FromProcedureName
    Dim p As Scripting.Dictionary
    Set p = BuildStorageParams(True)
    p.Add "key", Key
    p.Add "value", value
    cdpInvoke "DOMStorage.setDOMStorageItem", p
    printMsg Debug_, "localStorage set: key=" & Key, FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：localStorage からキーに対応する値を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：見つからない場合は空文字
'***************************************************************************************************
Public Function LocalStorageGetItem(ByVal Key As String) As String
    Const FromProcedureName As String = ThisClassName & ".LocalStorageGetItem"


    LocalStorageGetItem = DomStorageGetItemCore(True, Key, FromProcedureName)
End Function

'***************************************************************************************************
'* 機能　　：localStorage の全エントリを Object で返します（`DOMStorage.getDOMStorageItems` の `entries`）
'---------------------------------------------------------------------------------------------------
'* 返り値　：通常は `[["k","v"],...]` をパースした `Collection`。取得失敗時は `Nothing`。
'***************************************************************************************************
Public Function LocalStorageGetAll() As Object
    Const FromProcedureName As String = ThisClassName & ".LocalStorageGetAll"


    Set LocalStorageGetAll = DomStorageGetAllCore(True, FromProcedureName)
End Function

'***************************************************************************************************
'* 機能　　：localStorage から指定キーを削除します（`DOMStorage.removeDOMStorageItem`）
'***************************************************************************************************
Public Sub LocalStorageRemoveItem(ByVal Key As String)
    Const FromProcedureName As String = ThisClassName & ".LocalStorageRemoveItem"


    EnsureDOMStorageEnabled FromProcedureName
    Dim p As Scripting.Dictionary
    Set p = BuildStorageParams(True)
    p.Add "key", Key
    cdpInvoke "DOMStorage.removeDOMStorageItem", p
    printMsg Debug_, "localStorage remove key=" & Key, FromProcedureName
End Sub



'***************************************************************************************************
'                           ■■■ sessionStorage（DOMStorage） ■■■
'***************************************************************************************************
'* 機能　　：sessionStorage を空にします（`DOMStorage.clear`）
'***************************************************************************************************
Public Sub SessionStorageClear()
    Const FromProcedureName As String = ThisClassName & ".SessionStorageClear"


    EnsureDOMStorageEnabled FromProcedureName
    cdpInvoke "DOMStorage.clear", BuildStorageParams(False)
    printMsg info_, "sessionStorage cleared.", FromProcedureName
End Sub


'***************************************************************************************************
'* 機能　　：sessionStorage にキーと値を設定します（`DOMStorage.setDOMStorageItem`）
'***************************************************************************************************
Public Sub SessionStorageSetItem(ByVal Key As String, ByVal value As String)
    Const FromProcedureName As String = ThisClassName & ".SessionStorageSetItem"


    EnsureDOMStorageEnabled FromProcedureName
    Dim p As Scripting.Dictionary
    Set p = BuildStorageParams(False)
    p.Add "key", Key
    p.Add "value", value
    cdpInvoke "DOMStorage.setDOMStorageItem", p
    printMsg Debug_, "sessionStorage set: key=" & Key, FromProcedureName
End Sub


'***************************************************************************************************
'* 機能　　：sessionStorage からキーに対応する値を取得します
'***************************************************************************************************
Public Function SessionStorageGetItem(ByVal Key As String) As String
    Const FromProcedureName As String = ThisClassName & ".SessionStorageGetItem"


    SessionStorageGetItem = DomStorageGetItemCore(False, Key, FromProcedureName)
End Function


'***************************************************************************************************
'* 機能　　：sessionStorage の全エントリを Object で返します（`entries`）
'***************************************************************************************************
Public Function SessionStorageGetAll() As Object
    Const FromProcedureName As String = ThisClassName & ".SessionStorageGetAll"


    Set SessionStorageGetAll = DomStorageGetAllCore(False, FromProcedureName)
End Function


'***************************************************************************************************
'* 機能　　：sessionStorage から指定キーを削除します（`DOMStorage.removeDOMStorageItem`）
'***************************************************************************************************
Public Sub SessionStorageRemoveItem(ByVal Key As String)
    Const FromProcedureName As String = ThisClassName & ".SessionStorageRemoveItem"


    EnsureDOMStorageEnabled FromProcedureName
    Dim p As Scripting.Dictionary
    Set p = BuildStorageParams(False)
    p.Add "key", Key
    cdpInvoke "DOMStorage.removeDOMStorageItem", p
    printMsg Debug_, "sessionStorage remove key=" & Key, FromProcedureName
End Sub



'***************************************************************************************************
'                           ■■■ 内部実装（DOMStorage 共通） ■■■
'***************************************************************************************************
Private Sub EnsureDOMStorageEnabled(ByVal FromProcedureName As String)
    If mDomStorageEnabled Then Exit Sub

    printMsg Debug_, "DOMStorage.enable (first use)", FromProcedureName
    cdpInvoke "DOMStorage.enable"
    mDomStorageEnabled = True
End Sub


Private Function CurrentSecurityOrigin() As String
    Const FromProcedureName As String = ThisClassName & ".CurrentSecurityOrigin"


    On Error Resume Next
    CurrentSecurityOrigin = ex_CDPBrowser.jsEval("location.origin", dbgMsg:=False)
    If Err.Number <> 0 Then
        Err.Clear
        CurrentSecurityOrigin = vbNullString
        printMsg WARN_, "location.origin could not be read via jsEval.", FromProcedureName
    End If
    On Error GoTo 0
End Function


Private Function CurrentStorageKey() As String
    Const FromProcedureName As String = ThisClassName & ".CurrentStorageKey"


    CurrentStorageKey = vbNullString

    On Error Resume Next
    Dim res As Scripting.Dictionary
    Dim emptyParams As New Scripting.Dictionary
    Set res = cdpInvoke("Storage.getStorageKey", emptyParams, , False)
    If res Is Nothing Then GoTo EH

    If res.Exists("storageKey") Then
        CurrentStorageKey = CStr(res("storageKey"))
    End If

EH:
    If Err.Number <> 0 Then
        Err.Clear
        CurrentStorageKey = vbNullString
    End If
    On Error GoTo 0

    printMsg Trace_, "Storage.getStorageKey -> """ & CurrentStorageKey & """", FromProcedureName
End Function


Private Function BuildStorageIdDict(ByVal isLocal As Boolean) As Scripting.Dictionary
    Const FromProcedureName As String = ThisClassName & ".BuildStorageIdDict"


    Dim origin As String
    Dim sk As String

    origin = CurrentSecurityOrigin()
    sk = CurrentStorageKey()

    If Len(origin) = 0 Then
        printMsg ERROR_, "securityOrigin is empty. Navigate to a valid http(s) page first.", FromProcedureName, doRaiseError:=True
    End If

    Dim d As New Scripting.Dictionary
    d.Add "securityOrigin", origin
    d.Add "storageKey", sk
    d.Add "isLocalStorage", isLocal
    Set BuildStorageIdDict = d
End Function


Private Function BuildStorageParams(ByVal isLocal As Boolean) As Scripting.Dictionary
    Dim p As New Scripting.Dictionary
    p.Add "storageId", BuildStorageIdDict(isLocal)
    Set BuildStorageParams = p
End Function


Private Function DomStorageGetAllCore(ByVal isLocal As Boolean, ByVal FromProcedureName As String) As Object
    EnsureDOMStorageEnabled FromProcedureName

    Dim res As Scripting.Dictionary
    Set res = cdpInvoke("DOMStorage.getDOMStorageItems", BuildStorageParams(isLocal))

    If res Is Nothing Or Not res.Exists("entries") Then
        printMsg WARN_, "getDOMStorageItems: no entries in result.", FromProcedureName
        Exit Function
    End If

    Set DomStorageGetAllCore = res("entries")
    printMsg Debug_, "DomStorageGetAllCore -> TypeName=" & TypeName(DomStorageGetAllCore), FromProcedureName
End Function


Private Function DomStorageGetItemCore(ByVal isLocal As Boolean, ByVal Key As String, ByVal FromProcedureName As String) As String
    EnsureDOMStorageEnabled FromProcedureName

    Dim res As Scripting.Dictionary
    Set res = cdpInvoke("DOMStorage.getDOMStorageItems", BuildStorageParams(isLocal))

    If res Is Nothing Or Not res.Exists("entries") Then
        DomStorageGetItemCore = vbNullString
        printMsg Debug_, "DOMStorage get miss (no entries) key=" & Key, FromProcedureName
        Exit Function
    End If

    Dim coll As Object
    Set coll = res("entries")
    If coll Is Nothing Then
        DomStorageGetItemCore = vbNullString
        Exit Function
    End If

    Dim i As Long
    For i = 1 To coll.Count
        Dim pairObj As Object
        Set pairObj = coll(i)
        If TypeName(pairObj) = "Collection" Then
            Dim c As Collection
            Set c = pairObj
            If c.Count >= 2 Then
                If CStr(c.Item(1)) = Key Then
                    DomStorageGetItemCore = CStr(c.Item(2))
                    printMsg Trace_, "DOMStorage get hit key=" & Key, FromProcedureName
                    Exit Function
                End If
            End If
        End If
    Next i

    DomStorageGetItemCore = vbNullString
    printMsg Debug_, "DOMStorage get miss key=" & Key, FromProcedureName
End Function



'***************************************************************************************************
'                           ■■■ `CDPBrowser` 直通ヘルパ ■■■
'***************************************************************************************************
Private Sub sleep(Optional seconds As Double = 0.5)
    ex_CDPBrowser.sleep seconds
End Sub

Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
    ex_CDPBrowser.printMsg LogLevel_, strMsg, From, isHeader, doRaiseError
End Sub

Private Function cdpInvoke(methodName As String, Optional params As Scripting.Dictionary, Optional alwaysBrowserContext As Boolean = False, Optional StopError As Boolean = True) As Scripting.Dictionary
    Set cdpInvoke = ex_CDPBrowser.invokeMethod(methodName, params, alwaysBrowserContext, StopError)
End Function



'***************************************************************************************************
'                             ■■■ 初期設定 ■■■
'***************************************************************************************************
Friend Sub Init(Inheritance As CDPBrowser)
    Const FromProcedureName As String = ThisClassName & ".Init"


    Set ex_CDPBrowser = Inheritance
    printMsg info_, "CookieAndDOMStorage: Init bound to CDPBrowser session.", FromProcedureName
End Sub


Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    EnableBrowserEvents = False
    mDomStorageEnabled = False

    With ShSetting01_StartBrowser
        seeRawSendMsgDbg = .Range(.UseRangeName(8, FromProcedureName)).value
    End With
    On Error GoTo 0
End Sub
