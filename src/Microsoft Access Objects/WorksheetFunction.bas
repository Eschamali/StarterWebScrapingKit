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
Public Function Unichar(Code As Long) As String
    ' 無効なUnicodeコードポイントのチェック
    If Code < 0 Or Code > 1114111 Then ' 1114111 = U+10FFFF
        Err.Raise 5, "Unichar", "無効なUnicodeコードポイントです。"
        Exit Function
    End If

    ' 1. 基本多言語面（BMP: U+0000 ～ U+FFFF）
    If Code < 2 ^ 16 Then
        If Code < 2 ^ 15 Then
            Unichar = ChrW(Code)
        Else
            Unichar = ChrW(Code - 2 ^ 16)
        End If

    ' 2. 追加多言語面（サロゲートペア領域: U+10000 ～ U+10FFFF）
    Else
        Dim uDash As Long
        Dim highSurrogate As Long
        Dim lowSurrogate As Long

        uDash = Code - 2 ^ 16

        ' 55296 = U+D800 (ハイサロゲート開始)
        ' 56320 = U+DC00 (ローサロゲート開始)
        highSurrogate = 55296 + (uDash \ 2 ^ 10)
        lowSurrogate = 56320 + (uDash Mod 2 ^ 10)

        ' ChrW用にInteger範囲(-32768～32767)へ変換して結合
        Unichar = ChrW(highSurrogate - 2 ^ 16) & ChrW(lowSurrogate - 2 ^ 16)
    End If
End Function
