Attribute VB_Name = "CDPHelpers"
'==============================================================================================================
'                              汎用的に使うCDP関連の便利プロシージャです
'                        特定のClassでしか使わない物はここに配置してはいけません
'==============================================================================================================
Option Explicit
Option Private Module



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

'----- 待機関連 -----
Private Declare PtrSafe Sub sleep2 Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long    'タイマー用
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long         '周波数取得用



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
Private Const ThisModuleName   As String = "CDPHelpers" 'トレース用



'***************************************************************************************************
'                                   ■■■ 各種変数 ■■■
'***************************************************************************************************
Private m_Frequency As Currency '実行マシンでの周波数記録用



'***************************************************************************************************
'                       ■■■ Enum → 文字列 変換プロシージャ ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] ブラウザ種別をexe名で返します
'---------------------------------------------------------------------------------------------------
Public Function EnumToStringBrowserList(param As BrowserList) As String
    Select Case param
        Case BrowserList.RunChrome: EnumToStringBrowserList = "chrome.exe"
        Case BrowserList.RunEdge:   EnumToStringBrowserList = "msedge.exe"
    End Select
End Function



'***************************************************************************************************
'                                   ■■■ パス情報 ■■■
'***************************************************************************************************
'* 機能　　：利用するブラウザパスを取得します
'---------------------------------------------------------------------------------------------------
'* 引数　　：BrowserType    起動するブラウザ種別（`BrowserList`）
'* 返り値　：ブラウザパス　 ※失敗時は、`vbnullstring`で返します
'***************************************************************************************************
Public Function getBrowserPath(ByVal BrowserType As BrowserList) As String
    Const FromProcedureName As String = ThisModuleName & ".getBrowserPath"

    Dim AppPathName As String
    AppPathName = EnumToStringBrowserList(BrowserType)

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
        Select Case BrowserType
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
'                                   ■■■ ポリシー情報 ■■■
'***************************************************************************************************
'* 機能　　：`RemoteDebuggingAllowed`の確認を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：useHKCU            True ：HKEY_CURRENT_USER　側を確認
'                               False：HKEY_LOCAL_MACHINE 側を確認
'
'            TargetBrowser      確認先のブラウザポリシーパス("SOFTWARE\Policies"以降を、"\"始まりから入力)
'
'* 返り値　：`RemoteDebuggingAllowed:=0`で、`True`となります
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/remotedebuggingallowed
'***************************************************************************************************
Public Function CheckBrowserPolicyRemoteDebuggingAllowed(ByVal useHKCU As Boolean, ByVal TargetBrowser As String) As Boolean
    Const checkValueName As String = "RemoteDebuggingAllowed"


    '1. 必要な変数を用意
    Dim KeyExists       As Long     '存在有無結果
    Dim ValueContents   As Long     '値の中身
    Dim ValueSize       As Long     '値の大きさ
    TargetBrowser = "SOFTWARE\Policies" & TargetBrowser

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
'            TargetBrowser      確認先のブラウザポリシーパス("SOFTWARE\Policies"以降を"、\"始まりから入力)
'
'* 返り値　：パスが返ったらそれが固定の保存場所です。`vbnullstring`の場合は、制限なしです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-browser-policies/userdatadir
'***************************************************************************************************
Public Function CheckBrowserPolicyUserDataDir(ByVal useHKCU As Boolean, ByVal TargetBrowser As String) As String
    Const checkValueName As String = "UserDataDir"


    '1. 必要な変数を用意
    Dim KeyExists       As Long     '存在有無結果
    Dim ValueSize       As Long     '値の大きさ
    TargetBrowser = "SOFTWARE\Policies" & TargetBrowser

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
'                                     ■■■ 待機系 ■■■
'***************************************************************************************************
'* 機能　　：シンプルな待機です
'---------------------------------------------------------------------------------------------------
'* 引数　　：seconds    何秒間止めるか？
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Custom sleep function. Sleep by 0.5s by default.
'* 注意事項：Useful for a quick necessary pause when needed.
'***************************************************************************************************
Public Sub Sleep(Optional seconds As Double = 0.5)
    Const baseUnit As Long = 1000    'ie. millisecs


    sleep2 seconds * baseUnit
    DoEvents
End Sub

'***************************************************************************************************
'* 機能　　：PC起動時からの「経過時間（ミリ秒）」を正確に返します（Timer関数の神互換）
'---------------------------------------------------------------------------------------------------
'* 返り値　：PC起動時からの「経過時間（ミリ秒）」
'---------------------------------------------------------------------------------------------------
'* 詳細説明：これを使う前に、`InitQueryPerformanceFrequency`を呼び出すこと
'***************************************************************************************************
Public Function TimerCounter() As Double
    '1. 今の「振動回数」を取得する
    Dim currentCount As Currency
    Call QueryPerformanceCounter(currentCount)

    '2. 割り算して「秒」にし、1000倍して「ミリ秒（Double型）」にして返す！
    TimerCounter = (CDbl(currentCount) / CDbl(m_Frequency)) * 1000#
End Function

'***************************************************************************************************
'* 機能　　：実行マシンの「1秒間の振動数（周波数）」を取得して記憶します
'***************************************************************************************************
Public Sub InitQueryPerformanceFrequency()
    Call QueryPerformanceFrequency(m_Frequency)
End Sub
