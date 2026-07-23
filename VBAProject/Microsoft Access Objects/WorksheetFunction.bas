Attribute VB_Name = "WorksheetFunction"
'==============================================================================================================
'                      Excel依存の一部をAccessで再現するためのHelperパッチモジュールです
'==============================================================================================================
Option Compare Database
Option Explicit

'エラー番号用
Public Const xlErrValue As Long = 2015
Public Const xlErrNA    As Long = 2042
Public Const xlErrNum   As Long = 2036



'***************************************************************************************************
'* 機能　　：Unicodeコードポイントから文字（絵文字含む）を生成します（WorksheetFunction.Unicharの完全代替）
'* 引数　　：CodePoint   10進数（例: 128512）または 16進数（例: &H1F600）のUnicodeコードポイント
'***************************************************************************************************
Public Function Unichar(CodePoint As Long) As String
    ' 1. 16ビット以内（U+FFFF以下）の通常の文字なら、VBA標準のChrWをそのまま使用
    If CodePoint <= &HFFFF Then
        Unichar = ChrW(CodePoint)
        Exit Function
    End If
    
    ' 2. 16ビットを超える絵文字など（U+10000以上）は、サロゲートペアの計算で分解して結合！
    Dim cp As Long: cp = CodePoint - &H10000
    Dim highSurrogate As Long: highSurrogate = &HD800 + (cp \ &H400)
    Dim lowSurrogate  As Long: lowSurrogate = &HDC00 + (cp Mod &H400)
    
    ' 2つのサロゲート文字を連結して、1つの絵文字として完成させる
    Unichar = ChrW(highSurrogate) & ChrW(lowSurrogate)
End Function
