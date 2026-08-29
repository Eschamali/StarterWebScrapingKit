VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPCoreHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'                             ローカルPC上のブラウザに関する機能群です
'                            Chromium系のみ取り扱います。※WebView2を除く
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ WindowsAPI宣言 ■■■
'***************************************************************************************************
'ブラウザパス、ポリシー確認用API
Private Declare PtrSafe Function RegGetValueW Lib "Advapi32" ( _
    ByVal hKey As LongPtr, _
    ByVal lpSubKey As LongPtr, _
    ByVal lpValue As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef pdwType As Long, _
    ByVal pvData As LongPtr, _
    ByRef pcbData As Long _
) As Long

'DJP 20241105 to solve problems with terminating sessions on terminal server
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" ( _
    ) As Long

Private Declare PtrSafe Function ProcessIdToSessionId Lib "kernel32" ( _
    ByVal dwProcessId As Long, _
    ByRef pSessionId As Long) As Long



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
'===============レジストリ関連の定数===============
' ルートキー
Private Const HKEY_LOCAL_MACHINE    As Long = &H80000002    'このPCの、"すべて"のユーザー
Private Const HKEY_CURRENT_USER     As Long = &H80000001    '"今、ログインしているユーザー"だけ
' 戻り値の型（RegGetValueW の dwFlags に指定する制限）
Private Const RRF_RT_REG_SZ         As Long = &H2           ' 文字列 (REG_SZ) に限定
Private Const RRF_RT_REG_DWORD      As Long = &H10          ' 32ビット数値 (REG_DWORD) に限定
' 型サイズ
Private Const SIZE_REG_DWORD        As Long = 4             ' DWORD は 4バイト (32bit) と決まっている
' APIの戻り値
Private Const RRF_SUCCESS           As Long = 0             '成功サイン
'==================================================

'その他
Private Const ThisClassName             As String = "CDPCoreHost"   'ログのトレース用



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'---------------- ポリシー状況一覧 ----------------
Private Type StatePolicy
    Exists      As Boolean  '存在有無
    Contents    As Variant  '中身の値
End Type

Private Type CheckPolicy
    UserDataDir As StatePolicy  'https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/userdatadir
End Type
Private currentPolicy As CheckPolicy
Private UseChrome     As Boolean
'--------------------------------------------------

'取得するブラウザ設定
Private UserBrowserMode As BrowserList
Private DebugUserName   As String
Private cRemotePort     As Long

'ログ設定
Private currentLog              As StateLog



'***************************************************************************************************
'                                   ■■■ 起動前準備 ■■■
'***************************************************************************************************
'* 機能　　：このClassが機能する必要な情報を渡します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       記録する識別名称
'            BrowserType    ブラウザタイプ
'            RemotePort     Port経由で制御する場合は指定
'***************************************************************************************************
Friend Sub InitInfo(UserName As String, BrowserType As BrowserList, Optional RemotePort As Long)
    DebugUserName = UserName
    UserBrowserMode = BrowserType
    cRemotePort = RemotePort
End Sub

'***************************************************************************************************
'* 機能　　：新しいデバッグブラウザセッションを開始する前にクリーンアップを行います
'---------------------------------------------------------------------------------------------------
'* 機能説明：・基本は対象ブラウザの`remote-debugging-pipe/port`プロセスのみを終了します
'            ・コマンドライン引数に加えて、ターミナルサーバー上で (このユーザーに属さない) セッションを閉じる際の問題も解決しています
'***************************************************************************************************
Public Sub cleanUpSessions()
    Const FromProcedureName As String = ThisClassName & ".cleanUpSessions"


    '1. 今、Excelが動いてるプロセスIDとそのログインユーザーのセッションIDを取得します
    Dim lProcessId As Long, lSessionId As Long
    lProcessId = GetCurrentProcessId()          'Get the current(Excel) process ID
    ProcessIdToSessionId lProcessId, lSessionId 'Get the Session ID for the current(Excel) process

    ' 2. exe名を取得
    Dim exeName As String, tmp
    tmp = Split(getBrowserPath, "\")
    exeName = tmp(UBound(tmp))

    ' 3. 起動中のプロセスを巡る
    Dim proc As Object
    On Error Resume Next
    For Each proc In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process Where SessionId = " & lSessionId)
        '4. 起動中のブラウザ名であるか確認
        If LCase$(proc.Name) = LCase$(exeName) Then
            '5. Killタイプ判定
            If currentPolicy.UserDataDir.Exists Then
                '6. デバッグブラウザ(CdpSessions)であるか問わず、通常ブラウザもKillします
                printMsg Debug_, "The `UserDataDir` policy is in effect, so all running browser instances will be terminated..." & proc.CommandLine, FromProcedureName
                proc.Terminate

            ElseIf InStr(1, proc.CommandLine, GenerateArgRemoteDebugging) > 0 And InStr(1, proc.CommandLine, GenerateArgUserDataDir) > 0 Then
                '6. デバッグブラウザ(CdpSessions)であるかつ、対象ユーザー名なら、Killします
                printMsg Debug_, "Detect debug browser: " & proc.CommandLine, FromProcedureName
                printMsg Debug_, "I'm going to kill that browser.", FromProcedureName
                proc.Terminate
            End If
        End If
        DoEvents
    Next
    On Error GoTo 0
    Err.Clear

    printMsg info_, "cleanUpSessions completed and previous CdpSessionID removed", FromProcedureName

End Sub

Public Sub setCrashStateNormal()
'------------------------------------------------------------------
' Set Crashed state to Normal to prevent the "Restore Sessions?"
' popup from appearing every time the browser is started after
' being closed by browser.quit method.
' Requires user data dir profile stored in private var br.User and
' the browser type stored in private variable br.Type.
' Update: 22/01/26 - Tweaked to work well with the new multi-instance
'                    structure.
'------------------------------------------------------------------
    Const FromProcedureName As String = ThisClassName & ".setCrashStateNormal"


    Dim filePath_Preferences As String
    Dim strContents As String, strNewContents As String
    Dim intF As Integer

    '状況に応じたクラッシュパス情報を指定する
    '`UserDataDir`のポリシーがあったらそっちを優先
    If currentPolicy.UserDataDir.Exists Then
        'Use this switch instead if your organization disables UserDataDir with policy
        filePath_Preferences = currentPolicy.UserDataDir.Contents & "\" & DebugUserName & "\Preferences"
    Else
        'user-data-dir
        Select Case UserBrowserMode
            Case BrowserList.RunChrome: filePath_Preferences = Environ("LOCALAPPDATA") & "\Google\Chrome\" & DebugUserName & "\Default\Preferences"
            Case BrowserList.RunEdge:   filePath_Preferences = Environ("LOCALAPPDATA") & "\Microsoft\Edge\" & DebugUserName & "\Default\Preferences"
        End Select
    End If


    'クラッシュ情報がない場合は、ここで抜ける
    If LenB(Dir(filePath_Preferences)) = 0 Then Exit Sub

    intF = FreeFile
    Open filePath_Preferences For Binary As #intF
        strContents = Input$(LOF(intF), #intF)
    Close #intF

   'Precise replacement: only replace exit_type value
    strNewContents = Replace(strContents, """exit_type"":""Crashed""", """exit_type"":""Normal""", 1, -1, vbBinaryCompare)

    If strNewContents <> strContents Then
        intF = FreeFile
        Open filePath_Preferences For Binary As #intF
            Put #intF, 1, strNewContents
        Close #intF
    End If

    printMsg info_, "setCrashStateNormal completed", FromProcedureName

End Sub

'***************************************************************************************************
'* 機能　　：利用するブラウザパスを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：ブラウザパス　 ※失敗時は、`vbnullstring`で返します
'***************************************************************************************************
Public Function getBrowserPath() As String
    Const FromProcedureName As String = ThisClassName & ".getBrowserPath"


    Dim AppPathName As String
    AppPathName = EnumToStringBrowserList_exeName(UserBrowserMode)

    '1. カスタムパスの指定があったらそっちを優先
    With ShSetting01_StartBrowser
        If LenB(.UseRangeID(1, FromProcedureName)) Then
            '1-1. カスタムパスを指定(ポータブル版ブラウザなど)
            getBrowserPath = .UseRangeID(1, FromProcedureName)

        Else
            '1-1. 必要な変数を用意
            Dim strKeyPath      As String   'ベースパス
            Dim ValueSize       As Long     '値の大きさ

            '1-2. ブラウザのインストールパスが記録されてるレジストリ場所を指定
            strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & AppPathName

            '1-3. レジストリから、ブラウザインストールパスを特定
            '※文字列で取るので仕組み上まず、文字数チェックを行います
            If RegGetValueW(HKEY_LOCAL_MACHINE, StrPtr(strKeyPath), 0, RRF_RT_REG_SZ, 0, 0, ValueSize) = RRF_SUCCESS Then
                '1-4. 得た文字数を基に改めて、値の中身を取り出しつつ、末尾の`vbNullChar`も落とすように調整する
                getBrowserPath = String(ValueSize / 2 - 2, vbNullChar)
                RegGetValueW HKEY_LOCAL_MACHINE, StrPtr(strKeyPath), 0, RRF_RT_REG_SZ, 0, StrPtr(getBrowserPath), ValueSize
            End If
        End If
    End With

    '2. パスが空の場合は、次の捜索へ
    If LenB(getBrowserPath) = 0 Then
        '2-1. デフォルトインストールパスチェック
        '※Chrome,Edgeのみ確認します
        Select Case UserBrowserMode
            Case BrowserList.RunChrome
                If LenB(Dir(Environ("ProgramFiles") & "\Google\Chrome\Application\" & AppPathName)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles") & "\Google\Chrome\Application\" & AppPathName
                ElseIf LenB(Dir(Environ("ProgramFiles(x86)") & "\Google\Chrome\Application\" & AppPathName)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles(x86)") & "\Google\Chrome\Application\" & AppPathName
                End If
            Case BrowserList.RunEdge
                If LenB(Dir(Environ("ProgramFiles") & "\Microsoft\Edge\Application\" & AppPathName)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles") & "\Microsoft\Edge\Application\" & AppPathName
                ElseIf LenB(Dir(Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\" & AppPathName)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\" & AppPathName
                End If
        End Select

        '2-2. 存在するか？
        If LenB(getBrowserPath) > 0 Then
            'しなかったら、`vbnullstring`にしておく
            If LenB(Dir(getBrowserPath)) = 0 Then getBrowserPath = vbNullString
        End If
    End If
End Function



'***************************************************************************************************
'                                 ■■■ 起動引数生成アシスト ■■■
'***************************************************************************************************
'* 機能　　：`user-data-dir`引数生成を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       生成フォルダ名
'            BrowserType    ブラウザタイプ
'
'* 返り値　：`--user-data-dir="..."`として返却
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ブラウザに合わせたパスと同じ場所に配置します
'* 注意事項：デバッグ時はこれが必須です。https://developer.chrome.com/blog/remote-debugging-port
'***************************************************************************************************
Private Function GenerateArgUserDataDir() As String
    Const ArgName As String = "user-data-dir"


    Select Case UserBrowserMode
        Case BrowserList.RunEdge:    GenerateArgUserDataDir = "--" & ArgName & "=""" & Environ("LOCALAPPDATA") & "\Microsoft\Edge\" & DebugUserName & """"
        Case BrowserList.RunChrome:  GenerateArgUserDataDir = "--" & ArgName & "=""" & Environ("LOCALAPPDATA") & "\Google\Chrome\" & DebugUserName & """"
    End Select
End Function

'***************************************************************************************************
'* 機能　　：`profile-directory`引数生成を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       生成フォルダ名
'            BrowserType    ブラウザタイプ
'
'* 返り値　：`--profile-directory="..."`として返却
'---------------------------------------------------------------------------------------------------
'* 注意事項：`UserDataDir`ポリシーに引っかかる場合は、これを代わりに指定します
'***************************************************************************************************
Private Function GenerateArgProfileDirectory() As String
    Const ArgName As String = "profile-directory"


    Select Case UserBrowserMode
        Case BrowserList.RunEdge:    GenerateArgProfileDirectory = "--" & ArgName & "=""" & DebugUserName & """"
        Case BrowserList.RunChrome:  GenerateArgProfileDirectory = "--" & ArgName & "=""" & DebugUserName & """"
    End Select
End Function

'***************************************************************************************************
'* 機能　　：`UserDataDir`ポリシーに応じて、最適なユーザーデータ保存先引数を提供します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`GenerateArgUserDataDir` or `GenerateArgUserDataDir`のいずれか
'---------------------------------------------------------------------------------------------------
'* 注意事項：`CheckRemoteDebuggingBlockedByPolicy`を事前に処理させておくこと
'***************************************************************************************************
Public Function GenerateArgUseUserData() As String
    GenerateArgUseUserData = IIf(currentPolicy.UserDataDir.Exists, GenerateArgProfileDirectory, GenerateArgUserDataDir)
End Function

'***************************************************************************************************
'* 機能　　：初期起動時の遷移URL引数生成を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：appUrl     初期遷移URL
'* 返り値　：`--app="..."`として返却。
'---------------------------------------------------------------------------------------------------
'* 注意事項：引数を空で渡した場合は、「--homepage ""about:blank"""」で返却します
'***************************************************************************************************
Public Function GenerateArgApp(appUrl As String) As String
    Const ArgName As String = "app"


    If LenB(appUrl) Then
        GenerateArgApp = "--" & ArgName & "=""" & appUrl & """"
    Else
        GenerateArgApp = "--homepage ""about:blank"""
    End If
End Function

'***************************************************************************************************
'* 機能　　：デバッグ引数を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：`--remote-debugging-...`として返却
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`InitInfo`の第三引数に応じて変わります
'***************************************************************************************************
Public Function GenerateArgRemoteDebugging() As String
    GenerateArgRemoteDebugging = IIf(cRemotePort = 0, "--remote-debugging-pipe", "--remote-debugging-port=" & cRemotePort)
End Function



'***************************************************************************************************
'                                   ■■■ ポリシー情報 ■■■
'***************************************************************************************************
'* 機能　　：リモート デバッグに関するポリシーチェックを行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：BrowserMode    ブラウザタイプ
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・企業によっては、ポリシー設定でリモートデバッグできないようにしています。
'　　　　　　・設定されていると、起動はできたのに、なんか操作できないといった混乱が発生するため、このチェックを設けています
'            ・現在は下記の確認に対応してます
'               ・RemoteDebuggingAllowed
'               ・UserDataDir
'***************************************************************************************************
Public Sub CheckRemoteDebuggingBlockedByPolicy()
    Const FromProcedureName As String = ThisClassName & ".IsRemoteDebuggingBlockedByPolicy"


    'ブラウザ名に応じて、ベースとなるキーパスを設定
    Dim RegPathBrowser As String
    RegPathBrowser = EnumToStringBrowserList_RegPath(UserBrowserMode)

    '"HKLM(0)"と"HKCU(1)"を確認する
    '最初は、PC単位、次に、現在のログインユーザー単位で確認します
    Dim i As Long
    Dim ErrorMes As String
    For i = 0 To 1
        ' ------------ 1. RemoteDebuggingAllowed ------------
        '※これに引っかかると残念ですが、CDP制御はできません
        If CheckBrowserPolicyRemoteDebuggingAllowed(i) Then
            '通達
            ErrorMes = IIf(i, "現在のログインPCアカウントでは、ブラウザのリモートデバッグが、ブロックされています。" & vbCrLf & vbCrLf & "CDP機能は、利用できません。" & vbCrLf & "これが個人のPCアカウントの場合は、レジストリ操作にて`1`にして下さい。", _
                                "会社のセキュリティポリシーにより、ブラウザのリモートデバッグが、ブロックされています。" & vbCrLf & vbCrLf & "CDP機能は、利用できません。" & vbCrLf & "情報システム部門に、お問い合わせください。")

            printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True
            Exit For
        End If


        '------------ 2. UserDataDir ------------
        '※これに引っかかるとCDP制御はできますが、デバッグブラウザとは関係ない普通のブラウザも毎回、落とさないといけない制限となります
        currentPolicy.UserDataDir.Contents = CheckBrowserPolicyUserDataDir(i)

        'ポリシーとして存在してるなら中身のチェック
        If StrPtr(currentPolicy.UserDataDir.Contents) Then
            'フラグ立て
            currentPolicy.UserDataDir.Exists = True

            '空文字チェック
            If LenB(currentPolicy.UserDataDir.Contents) = 0 Then
                '空文字でも固定化されるため、そのパスを提示しておく
                currentPolicy.UserDataDir.Contents = Environ("LOCALAPPDATA") & "\" & RegPathBrowser & "\User Data"
            End If

            '通達
            ErrorMes = IIf(i, "現在のログインPCアカウントでは、ブラウザデータの配置場所が`" & currentPolicy.UserDataDir.Contents & "`に固定されています。起動中のブラウザを全て落とします。これが個人のPCアカウントであり、この仕様を回避する場合は、レジストリ操作にて削除して下さい。", _
                              "会社のセキュリティポリシーにより、ブラウザデータの配置場所が`" & currentPolicy.UserDataDir.Contents & "`に固定されています。起動中のブラウザを全て落とします。")

            printMsg WARN_, ErrorMes, FromProcedureName
            Exit For
        End If
    Next
End Sub

'***************************************************************************************************
'* 機能　　：`RemoteDebuggingAllowed`の確認を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：useHKCU            True ：HKEY_CURRENT_USER　側を確認
'                               False：HKEY_LOCAL_MACHINE 側を確認
'
'* 返り値　：`RemoteDebuggingAllowed:=0`で、`True`となります
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/remotedebuggingallowed
'***************************************************************************************************
Private Function CheckBrowserPolicyRemoteDebuggingAllowed(ByVal useHKCU As Boolean) As Boolean
    Const checkValueName As String = "RemoteDebuggingAllowed"


    '1. 必要な変数を用意
    Dim KeyExists       As Long     '存在有無結果
    Dim ValueContents   As Long     '値の中身
    Dim ValueSize       As Long     '値の大きさ
    Dim TargetBrowser   As String: TargetBrowser = "SOFTWARE\Policies" & EnumToStringBrowserList_RegPath(UserBrowserMode)

    '2. ルートキーをセット
    Dim RootTree As Long
    RootTree = IIf(useHKCU, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE)

    '3. 32bit値としてチェックする
    ValueSize = SIZE_REG_DWORD
    KeyExists = RegGetValueW(RootTree, StrPtr(TargetBrowser), StrPtr(checkValueName), RRF_RT_REG_DWORD, 0, VarPtr(ValueContents), ValueSize)

    '4. そのキー名があるか？
    If KeyExists = RRF_SUCCESS Then
        '0:リモート デバッグを使用できません
        '1:リモート デバッグを使用できる
        CheckBrowserPolicyRemoteDebuggingAllowed = IIf(ValueContents = 0, True, False)
    End If
End Function

'***************************************************************************************************
'* 機能　　：`UserDataDir`の確認を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：useHKCU            True ：HKEY_CURRENT_USER　側を確認
'                               False：HKEY_LOCAL_MACHINE 側を確認
'
'* 返り値　：パスが返ったらそれが固定の保存場所です。`vbnullstring`の場合は、制限なしです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/userdatadir
'***************************************************************************************************
Private Function CheckBrowserPolicyUserDataDir(ByVal useHKCU As Boolean) As String
    Const checkValueName As String = "UserDataDir"


    '1. 必要な変数を用意
    Dim KeyExists       As Long     '存在有無結果
    Dim ValueSize       As Long     '値の大きさ
    Dim TargetBrowser   As String: TargetBrowser = "SOFTWARE\Policies" & EnumToStringBrowserList_RegPath(UserBrowserMode)

    '2. ルートキーをセット
    Dim RootTree As Long
    RootTree = IIf(useHKCU, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE)

    '3. 文字列なのでまずは文字数を取得
    KeyExists = RegGetValueW(RootTree, StrPtr(TargetBrowser), StrPtr(checkValueName), RRF_RT_REG_SZ, 0, 0, ValueSize)

    '4. そのキー名があるか？
    If KeyExists = RRF_SUCCESS Then
        '5. 得た文字数を基に、バッファーを確保しつつ、末尾の`vbNullChar`も落とすように調整する
        CheckBrowserPolicyUserDataDir = String(ValueSize / 2 - 2, vbNullChar)

        '6. 返却
        RegGetValueW RootTree, StrPtr(TargetBrowser), StrPtr(checkValueName), RRF_RT_REG_SZ, 0, StrPtr(CheckBrowserPolicyUserDataDir), ValueSize
    End If
End Function



'***************************************************************************************************
'                                     ■■■ ログヘルパー ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, currentLog.LogID, isHeader, doRaiseError, currentLog.logPath
End Sub



'***************************************************************************************************
'                                     ■■■ 初期化 ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    Const FromProcedureName As String = ThisClassName & ".Class_Initialize"


    'Assign an objID for easy debug identification
    currentLog.LogID = "CHID" & Format(Rnd * 1000, "000")

    'ログ設定
    currentLog.seeRawSendMsgDbg = ShSetting01_StartBrowser.UseRangeID(8, FromProcedureName)
    currentLog.logPath = ShSetting01_StartBrowser.UseRangeID(9, FromProcedureName)
End Sub
