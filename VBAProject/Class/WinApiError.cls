VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'                    VBAから、WindowsAPI使用時のエラーコードを詳細情報を出力します
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                             ■■■ WindowsAPIと定数宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, ByVal lpSource As LongPtr, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long

Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200



'***************************************************************************************************
'                               ■■■ 公開プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：エラー番号からメッセージを取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に沿ったエラーメッセージ
'* 引数　　：ErrorCode      通常は、Err.LastDllError を指定します
'            DllName        特定のWinAPI(○○.dll)のエラーと分かってるなら、それを指定します
'---------------------------------------------------------------------------------------------------
'* 注意事項：返る言語は、Windowsの設定言語に依存します
'***************************************************************************************************
Public Function GetMessage(ByVal ErrorCode As Long, Optional ByVal DllName As String) As String
    Dim Buffer As String
    Dim length As Long
    Dim hModule As LongPtr
    Dim Flags As Long
    Dim isLoaded As Boolean
    
    Buffer = Space(1024)
    Flags = FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS
    
    ' DLL指定がある場合
    If DllName <> "" Then
        ' まずロード済みか探す
        hModule = GetModuleHandle(DllName)
        
        ' なければロードする
        If hModule = 0 Then
            hModule = LoadLibrary(DllName)
            isLoaded = True
        End If
        
        If hModule <> 0 Then
            Flags = Flags Or FORMAT_MESSAGE_FROM_HMODULE
        End If
    End If
    
    ' メッセージ取得
    length = FormatMessage(Flags, hModule, ErrorCode, 0, Buffer, Len(Buffer), 0)
    
    If length > 0 Then
        ' 末尾の改行とかヌル文字を消して返す
        GetMessage = Replace(Left(Buffer, length), vbCrLf, "")
    Else
        GetMessage = "定義されていないエラーです (Code: " & ErrorCode & ")"
    End If
End Function
