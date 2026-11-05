VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'                  VBAから、WindowsAPI使用時のエラーコードから詳細情報を出力します
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                             ■■■ WindowsAPIと定数宣言 ■■■
'***************************************************************************************************
'エラーメッセージを受け取る核API
Private Declare PtrSafe Function FormatMessageW Lib "kernel32" ( _
    ByVal dwFlags As Long, ByVal lpSource As LongPtr, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByVal Arguments As LongPtr) As Long

'特定のDLL向けのエラーメッセージを使用する場合に
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr

'文字列バッファー管理用
Private Declare PtrSafe Function SysReAllocString Lib "oleaut32" (ByVal pbstr As LongPtr, ByVal psz As LongPtr) As Long
Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
Private Const FORMAT_MESSAGE_FROM_HMODULE    As Long = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM     As Long = &H1000



'***************************************************************************************************
'                               ■■■ メインプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：エラー番号からメッセージを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に沿ったエラーメッセージ
'* 引数　　：ErrorCode      通常は、Err.LastDllError を指定します
'            DllName        特定のWinAPI(○○.dll)のエラーと分かってるなら、それを指定します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・返る言語は、Windowsの設定言語に依存します
'            ・`DllName`に渡す名前は既に、VBAにロード済みである前提として扱います
'***************************************************************************************************
Public Function GetMessage(ByVal ErrorCode As Long, Optional ByVal DllName As String) As String
    Dim msg As String
    Dim hModule As LongPtr

    '1. DLL指定があり、読み込み済みなら、そのDLLのメッセージテーブルから引く
    If LenB(DllName) Then hModule = GetModuleHandle(DllName)
    If hModule Then
        If TryFormat(FORMAT_MESSAGE_FROM_HMODULE Or FORMAT_MESSAGE_FROM_SYSTEM, hModule, ErrorCode, msg) Then GetMessage = msg: Exit Function
    End If

    '2. 引けなかったら、システムから引く
    If TryFormat(FORMAT_MESSAGE_FROM_SYSTEM, 0, ErrorCode, msg) Then GetMessage = msg: Exit Function

    '3. それでも無理なら、諦める
    GetMessage = "定義されていないエラーです (Code: " & ErrorCode & ")"
End Function



'***************************************************************************************************
'                                   ■■■ ヘルパーメソッド ■■■
'***************************************************************************************************
'* 機能　　：FormatMessage にメモリを確保させ、Stringに写してから解放する
'---------------------------------------------------------------------------------------------------
'* 返り値　：実行成功可否
'* 引数　　：Flags      エラーメッセージの取り出し手法フラグ
'            hModule    特定のDLLハンドル
'            ErrorCode  通常は、Err.LastDllError を指定します
'            Msg        `ErrorCode × hModule`に紐づく、エラーメッセージを`ByRef`として、返します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・返る言語は、Windowsの設定言語に依存します
'            ・`DllName`に渡す名前は既に、VBAにロード済みである前提として扱います
'***************************************************************************************************
Private Function TryFormat(ByVal Flags As Long, ByVal hModule As LongPtr, ByVal ErrorCode As Long, ByRef msg As String) As Boolean
    Dim pBuf As LongPtr

    '1. lpBuffer には「ポインターを書き込んでもらう変数のアドレス」を渡す
    If FormatMessageW(Flags Or FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                      hModule, ErrorCode, 0, VarPtr(pBuf), 0, 0) = 0 Then Exit Function

    '2. Stringに写して、FormatMessage が確保したメモリを返す
    SysReAllocString VarPtr(msg), pBuf
    LocalFree pBuf

    '3. 末尾の改行を落とす
    msg = Replace(msg, vbCrLf, "")
    TryFormat = True
End Function
