Attribute VB_Name = "Demo_CookieAndDOMStorage"
'***************************************************************************************************
'       exCDP_CookieAndDOMStorage 拡張 - デモ & 動作確認
'***************************************************************************************************
'* 機能　　：Cookie / localStorage / sessionStorage を高レベル API から操作するサンプルです。
'---------------------------------------------------------------------------------------------------
'* 対応拡張：Extensions\MainUnit\CDP\Cookie and DOMStorage\exCDP_CookieAndDOMStorage.cls
'---------------------------------------------------------------------------------------------------
'* 前提　　：`Demo_CDP.bas` の `設定シートからのCDP起動` が利用可能なブックで実行してください。
'*            テスト先は `https://example.com`（オリジンが明確な https ページ）を推奨します。
'***************************************************************************************************
Option Explicit


'***************************************************************************************************
'* 内部ヘルパ：`GetCookies` / `*GetAll` が返す Object を JSON 文字列にします（表示用）
'*   - `Nothing` のときは `[]` を返します（Cookie / entries とも空配列想定）
'***************************************************************************************************
Private Function ObjToJsonForDebug(ByVal o As Object, ByVal jc As WebJsonConverter) As String
    If o Is Nothing Then
        ObjToJsonForDebug = "[]"
        Exit Function
    End If
    On Error Resume Next
    ObjToJsonForDebug = jc.ConvertToJson(o)
    If Err.Number <> 0 Then
        Err.Clear
        ObjToJsonForDebug = "[]"
    End If
    On Error GoTo 0
End Function


'***************************************************************************************************
'* 機能　　：Cookie と Storage 関連の一連操作を実行します。
'---------------------------------------------------------------------------------------------------
'* 確認ポイント：
'*   - ClearCookies / SetCookie / GetCookies がエラーなく完走すること
'*   - LocalStorage / SessionStorage の Set / Get / GetAll / Clear が期待どおりであること
'*   - イミディエイトウィンドウに JSON 文字列が出力されること
'***************************************************************************************************
Sub Demo_CookieAndDOMStorage()

    Const FromProcedureName As String = "Demo_CookieAndDOMStorage.Demo_CookieAndDOMStorage"


    Dim br As CDPContext
    Set br = 設定シートからのCDP起動ForTab("https://example.com")
    br.show
    br.wait

    Dim st As New exCDP_CookieAndDOMStorage
    st.Init br

    Dim jc As New WebJsonConverter

    ' -----------------------------
    ' Cookie
    ' -----------------------------
    st.ClearCookies

    Call st.SetCookie("vba_cookie", "abc123", "https://example.com", "", "/", False, False, "", 0)

    Debug.Print "Cookies=" & ObjToJsonForDebug(st.GetCookies, jc)

    st.ClearCookies

    ' -----------------------------
    ' LocalStorage
    ' -----------------------------
    st.LocalStorageClear
    st.LocalStorageSetItem "k1", "v1"

    Debug.Print "LocalStorage(k1)=" & st.LocalStorageGetItem("k1")
    Debug.Print "LocalStorageAll=" & ObjToJsonForDebug(st.LocalStorageGetAll, jc)

    ' -----------------------------
    ' SessionStorage
    ' -----------------------------
    st.SessionStorageClear
    st.SessionStorageSetItem "s1", "sv1"

    Debug.Print "SessionStorage(s1)=" & st.SessionStorageGetItem("s1")
    Debug.Print "SessionStorageAll=" & ObjToJsonForDebug(st.SessionStorageGetAll, jc)

    Debug.Print "[" & FromProcedureName & "] 完了。ブラウザを閉じます。"
    br.InheritanceCDPBrowser.quit
End Sub


'***************************************************************************************************
'* 機能　　：RemoveItem API の簡易確認
'***************************************************************************************************
Sub Demo_CookieAndDOMStorage_RemoveItems()

    Const FromProcedureName As String = "Demo_CookieAndDOMStorage.Demo_CookieAndDOMStorage_RemoveItems"


    Dim br As CDPContext
    Set br = 設定シートからのCDP起動ForTab("https://example.com")
    br.show
    br.wait

    Dim st As New exCDP_CookieAndDOMStorage
    st.Init br

    Dim jc As New WebJsonConverter

    st.LocalStorageClear
    st.LocalStorageSetItem "a", "1"
    st.LocalStorageSetItem "b", "2"
    st.LocalStorageRemoveItem "a"
    Debug.Print "LocalStorageAll(a削除後)=" & ObjToJsonForDebug(st.LocalStorageGetAll, jc)

    st.SessionStorageClear
    st.SessionStorageSetItem "x", "9"
    st.SessionStorageRemoveItem "x"
    Debug.Print "SessionStorageAll(x削除後)=" & ObjToJsonForDebug(st.SessionStorageGetAll, jc)

    Debug.Print "[" & FromProcedureName & "] 完了。"
    br.InheritanceCDPBrowser.quit
End Sub
