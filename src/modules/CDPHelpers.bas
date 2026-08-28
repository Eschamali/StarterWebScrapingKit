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
Private Const ThisModuleName   As String = "CDPHelpers"    'トレース用



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
