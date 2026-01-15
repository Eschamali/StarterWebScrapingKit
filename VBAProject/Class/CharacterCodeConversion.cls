VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharacterCodeConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'       VBAの「Unicode」(UTF-16)文字列を別の文字コードに変換したり、その逆を行ったりします。
'           WebAPIのやり取りとか、日本語を含んだCDP-Jsonコマンドを行う際に利用します。
'            おまけとして、ファイルの読み書きも対応します。ADODB の代わりにどうぞ!
'***************************************************************************************************
' basUtf8FromString

' Written by David Ireland DI Management Services Pty Limited 2015
' <http://di-mgt.com.au> <http://cryptosys.net>
' @license MIT license <http://opensource.org/licenses/MIT>
' [2015-06-30] First published.
' [2018-07-27] Updated to cope with empty or null input string.
' [2018-08-15] Added Utf8BytesToString and BytesLength functions.
' [2018-08-17] Added declaration for 64-bit Office.
' [2026-01-15] 任意の変換元/変換先のコードページの設定ができるように改良。
'              `BytesFromString`メソッドにて、終端null文字を保持するか除去するかのスイッチングを実装。
'              シンプルなファイルの読み書き機能を追加(バイナリーモードのみ)



Option Explicit
'***************************************************************************************************
'                               ■■■ 必要なAPI宣言 ■■■
'***************************************************************************************************
''' WinApi function that maps a UTF-16 (wide character) string to a new character string
#If VBA7 Then
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long
#Else
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long
#End If
    
''' Maps a character string to a UTF-16 (wide character) string
#If VBA7 Then
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long _
    ) As Long
#Else
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long _
    ) As Long
#End If


'***************************************************************************************************
'                            ■■■ 必要なクラス変数/定数宣言 ■■■
'***************************************************************************************************
'------ 定数 ------
Private Const CP_UTF8 As Long = 65001   ' CodePage constant for UTF-8

'------ 変数 ------
Private UseCharacterCode As Long



'***************************************************************************************************
'                                ■■■ プロパティメソッド ■■■
'***************************************************************************************************
'* 機能　　：変換元/変換先の文字コードを指定します
'---------------------------------------------------------------------------------------------------
'* GET返値 ：現在のコードページ
'* LET引数 ：セットしたいコードページ
'---------------------------------------------------------------------------------------------------
'* 機能説明：VBAの「Unicode」(UTF-16)文字列を、どんな文字コードに変換するか、あるいは外部から得たどんな文字コードをVBAの「Unicode」(UTF-16)文字列に戻すかを設定できます
'* 注意事項：デフォルトは、65001(UTF-8)にしてるので基本的には変更する必要はないと思います
'***************************************************************************************************
Property Let CodePage(Code As Long)
    UseCharacterCode = Code
End Property

Property Get CodePage() As Long
    CodePage = UseCharacterCode
End Property



'***************************************************************************************************
'                            ■■■ プライベートメソッド ■■■
'***************************************************************************************************
''' Return length of byte array or zero if uninitialized
Private Function BytesLength(abBytes() As Byte) As Long
    ' Trap error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
End Function



'***************************************************************************************************
'                            ■■■ メインの文字コード変換メソッド ■■■
'***************************************************************************************************
''' Return byte array with VBA "Unicode" string encoded in Specified character code
Public Function BytesFromString(strInput As String, Optional TerminatingNullHold As Boolean) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    
    ' Catch empty or null input string
    BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(UseCharacterCode, 0&, StrPtr(strInput), -1, 0&, 0&, 0&, 0&)

    ' Boolean引数`TerminatingNullRemove`の仕様を利用して、終端nullを除去するかの補正を行う
    ReDim abBuffer(nBytes - 2 - TerminatingNullHold) ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(UseCharacterCode, 0&, StrPtr(strInput), -1, VarPtr(abBuffer(0)), nBytes - 1 - TerminatingNullHold, 0&, 0&)
    BytesFromString = abBuffer
End Function

''' Return VBA "Unicode" string from byte array encoded in Specified character code
Public Function BytesToString(abArray() As Byte) As String
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
    BytesToString = ""
    ' Catch uninitialized input array
    nBytes = BytesLength(abArray)
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(UseCharacterCode, 0&, VarPtr(abArray(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(UseCharacterCode, 0&, VarPtr(abArray(0)), nBytes, StrPtr(strOut), nChars)
    BytesToString = Left$(strOut, nChars)
End Function



'***************************************************************************************************
'                           ■■■ ファイルやり取りプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：バイト配列を指定パスに保存します
'---------------------------------------------------------------------------------------------------
'* 引数　　：abArray    指定の文字コードのルールに従ったバイナリ配列
'            FolderPath 保存先フォルダパス
'            FileName   ファイル名
'---------------------------------------------------------------------------------------------------
'* 機能説明：テキストも所詮、バイナリデータです。「Open ステートメント」を使ったバイナリ書き込みを行います。
'* 注意事項：・事前に、`BytesFromString`で、変換後のバイト配列を得る必要があります。
'            ・同名パスが、既にある場合は、上書きされます。呼び出し側で対処してください。
'            ・競合対策で、他のプロセスからの操作の制限を設けます。
'***************************************************************************************************
Sub BytesToSaveFile(abArray() As Byte, FolderPath As String, FileName As String)
    Dim fNum As Integer: fNum = FreeFile
    Open FolderPath & "\" & FileName For Binary Access Write Lock Read Write As #fNum
        Put #fNum, , abArray
    Close #fNum
End Sub

'***************************************************************************************************
'* 機能　　：ファイルの中身をバイナリデータとして読み込みます
'---------------------------------------------------------------------------------------------------
'* 返り値　：バイト配列
'* 引数　　：FolderPath 保存先フォルダパス
'            FileName   ファイル名
'---------------------------------------------------------------------------------------------------
'* 機能説明：テキストも所詮、バイナリデータです。「Open ステートメント」でバイナリデータで読み込んだ後、`BytesToString`でデコードしてください。
'* 注意事項：競合対策で、他のプロセスからの操作の制限を設けます
'***************************************************************************************************
Function BytesFromSavedFile(FolderPath As String, FileName As String) As Byte()
    Dim fNum As Integer: fNum = FreeFile
    Open FolderPath & "\" & FileName For Binary Access Read Lock Read Write As #fNum
        'ファイルサイズがあるか？
        If LOF(fNum) > 0 Then
            ' ファイルサイズ分の配列を確保（0ベースなので -1）
            ReDim BytesFromSavedFile(0 To LOF(fNum) - 1)
            
            ' 一括読み込み！
            Get #fNum, , BytesFromSavedFile
        Else
            ' 0バイトファイルの場合
            BytesFromSavedFile = vbNullString
        End If
    Close #fNum
End Function



'***************************************************************************************************
'                           ■■■ クラス構築時プロシージャ ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    '変換先/変換元 のデフォルトコードページは、UTF-8にします
    UseCharacterCode = CP_UTF8
End Sub
