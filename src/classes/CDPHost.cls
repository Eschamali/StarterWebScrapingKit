VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPHost"
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

Private Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As LongPtr



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

'OpenProcess 用アクセス権（TerminateProcess / 終了待ち / 生存確認に必要な最小セット）
Private Const PROCESS_TERMINATE                 As Long = &H1&
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
Private Const SYNCHRONIZE                       As Long = &H100000

'その他
Private Const ThisClassName             As String = "CDPCoreHost"           'ログのトレース用
Private Const WebSocketInfoFileName     As String = "DevToolsActivePort"



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
Private PipeMode        As Boolean

'ログ設定
Private currentLog      As StateLog



'***************************************************************************************************
'                                   ■■■ 起動前準備 ■■■
'***************************************************************************************************
'* 機能　　：このClassが機能する必要な情報を渡します
'---------------------------------------------------------------------------------------------------
'* 引数　　：UserName       記録する識別名称
'            BrowserType    ブラウザタイプ
'            UsePipe        Port経由で制御する場合は`False`
'***************************************************************************************************
Friend Sub InitInfo(UserName As String, BrowserType As BrowserList, UsePipe As Boolean)
    '各種情報反映
    DebugUserName = UserName
    UserBrowserMode = BrowserType
    PipeMode = UsePipe
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
        filePath_Preferences = getUserDataDirPath & "\Default\Preferences"
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
Property Get getBrowserPath() As String
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
End Property

'***************************************************************************************************
'* 機能　　：ブラウザが使用するユーザーパスを返します
'---------------------------------------------------------------------------------------------------
'* 返り値　："user-data-dir"のパス部分
'---------------------------------------------------------------------------------------------------
'* 詳細説明：以下の用途で役立ちます
'            ・`Preferences`内のクラッシュ情報をクリアする
'            ・`--remote-debugging-port=0`で起動した際の、`DevToolsActivePort`を拝見して、自動で空いてるポートに繋げる仕組み
'***************************************************************************************************
Property Get getUserDataDirPath() As String
    If currentPolicy.UserDataDir.Exists Then
        '`UserDataDir`ポリシーがある場合はそっちを優先に返却
        getUserDataDirPath = currentPolicy.UserDataDir.Contents
    Else
        Select Case UserBrowserMode
            Case BrowserList.RunEdge:    getUserDataDirPath = Environ("LOCALAPPDATA") & "\Microsoft\Edge\" & DebugUserName
            Case BrowserList.RunChrome:  getUserDataDirPath = Environ("LOCALAPPDATA") & "\Google\Chrome\" & DebugUserName
        End Select
    End If
End Property

'***************************************************************************************************
'* 機能　　：プロセスIDから、操作可能なプロセスハンドル(`hProcess`)を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功時は `hProcess`。失敗時は `0`
'* 引数　　：ProcessId      `Shell` 関数や `CreateProcess` が返すプロセスID
'---------------------------------------------------------------------------------------------------
'* 詳細説明：VBA の `Shell` は PID しか返さないため、`TerminateProcess` 用のハンドルはこちらで開きます
'* 注意事項：取得したハンドルは呼び出し側が `CloseHandle` すること
'***************************************************************************************************
Public Function OpenProcessHandle(ProcessId As LongPtr) As LongPtr
    Const FromProcedureName As String = ThisClassName & ".OpenProcessHandle"


    If ProcessId = 0 Then Exit Function

    OpenProcessHandle = OpenProcess(PROCESS_TERMINATE Or PROCESS_QUERY_LIMITED_INFORMATION Or SYNCHRONIZE, 0, CLng(ProcessId))
    If OpenProcessHandle = 0 Then
        Dim errNo As Long: errNo = Err.LastDllError
        printMsg WARN_, "OpenProcess failed. PID:" & ProcessId & " ErrorCode:" & errNo & " - " & WinApiError.GetMessage(errNo, "kernel32"), FromProcedureName
    Else
        printMsg Debug_, "OpenProcess succeeded. PID:" & ProcessId & " hProcess:" & OpenProcessHandle, FromProcedureName
    End If
End Function



'***************************************************************************************************
'                            ■■■ DevToolsActivePort 関連 ■■■
'***************************************************************************************************
'* 機能　　：ユーザーデータフォルダ直下にある「DevToolsActivePort」ファイルを削除します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：古いWebSocket情報を見ないようにする対策です
'***************************************************************************************************
Public Sub DeleteFileDevToolsActivePort()
    Dim DevToolsActivePortFilePath As String: DevToolsActivePortFilePath = getUserDataDirPath & "\" & WebSocketInfoFileName
    If LenB(Dir(DevToolsActivePortFilePath)) Then Kill DevToolsActivePortFilePath
End Sub

'***************************************************************************************************
'* 機能　　：ユーザーデータフォルダ直下にある「DevToolsActivePort」ファイルが出現するまで待機します
'---------------------------------------------------------------------------------------------------
'* 返り値　：10s後、あったら存在したら`True`となります
'---------------------------------------------------------------------------------------------------
'* 詳細説明：10s間待機します
'* 注意事項：必要に応じて、前述のプロシージャで削除を行います
'***************************************************************************************************
Public Function AutoWaitExistsFileDevToolsActivePort() As Boolean
    Const FromProcedureName As String = ThisClassName & ".AutoWaitExistsFileDevToolsActivePort"
    Const MaxTick As Long = 100 '無難に10.0s待機


    Dim TickCount As Long
    Do
        CDPHelpers.Sleep 0.1
        TickCount = TickCount + 1

        If LenB(Dir(getUserDataDirPath & "\" & WebSocketInfoFileName)) Then AutoWaitExistsFileDevToolsActivePort = True: Exit Do
        printMsg Debug_, "Waiting for the DevToolsActivePort file to appear...", FromProcedureName
    Loop Until TickCount = MaxTick
End Function

'***************************************************************************************************
'* 機能　　：`DevToolsActivePort`ファイルから、WebSocket接続情報を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：下記の構成となる1次元配列。※ない場合は、`vbEmpty`
'            要素0：ポート番号
'            要素1：WebSocketパス
'***************************************************************************************************
Public Function AnalysisDevToolsActivePort() As Variant
    'ファイル存在チェック
    Dim FilePath_DevToolsActivePort As String: FilePath_DevToolsActivePort = getUserDataDirPath & "\" & WebSocketInfoFileName
    If LenB(Dir(FilePath_DevToolsActivePort)) Then
        'そのパスを読み込んで、接続情報を取得する
        '1. ファイルOpen
        Dim intFree As Integer
        intFree = FreeFile
        Open FilePath_DevToolsActivePort For Input As #intFree

        '2. 中身を全部読む
        Dim tmp
        tmp = Input(LOF(intFree), #intFree)

        '3. 閉じる
        Close #intFree

        '4. 分割
        AnalysisDevToolsActivePort = Split(tmp, vbLf)
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
        Case BrowserList.RunEdge:    GenerateArgUserDataDir = "--" & ArgName & "=""" & getUserDataDirPath & """"
        Case BrowserList.RunChrome:  GenerateArgUserDataDir = "--" & ArgName & "=""" & getUserDataDirPath & """"
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
    GenerateArgRemoteDebugging = IIf(PipeMode, "--remote-debugging-pipe", "--remote-debugging-port=0")
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
                '空文字だとデフォルトパスで使用されて結果的にリモートデバッグ不可能になるので止める
                ErrorMes = IIf(i, "現在のログインPCアカウントでは、ブラウザデータの配置場所がデフォルトパスに固定されています。" & vbCrLf & vbCrLf & "Chromiumの仕様上、デフォルトパスのデバッグはできません。" & vbCrLf & "これが個人のPCアカウントの場合は、レジストリ操作にて削除するか、何かパスを指定して下さい。", _
                                  "会社のセキュリティポリシーにより、ブラウザデータの配置場所がデフォルトパスに固定されています。" & vbCrLf & vbCrLf & "Chromiumの仕様上、デフォルトパスのデバッグはできません。" & vbCrLf & "情報システム部門に、お問い合わせください。")

                printMsg ERROR_, ErrorMes, FromProcedureName, doRaiseError:=True
                Exit For
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
