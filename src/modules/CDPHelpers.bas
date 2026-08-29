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
'----- 待機関連 -----
Private Declare PtrSafe Sub sleep2 Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long    'タイマー用
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long         '周波数取得用



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const ThisModuleName   As String = "CDPHelpers"         'トレース用



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
Public Function EnumToStringBrowserList_exeName(param As BrowserList) As String
    Select Case param
        Case BrowserList.RunChrome: EnumToStringBrowserList_exeName = "chrome.exe"
        Case BrowserList.RunEdge:   EnumToStringBrowserList_exeName = "msedge.exe"
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ブラウザ種別を相対パス名で返します
'---------------------------------------------------------------------------------------------------
Public Function EnumToStringBrowserList_RegPath(param As BrowserList) As String
    Select Case param
        Case BrowserList.RunChrome: EnumToStringBrowserList_RegPath = "\Google\Chrome"
        Case BrowserList.RunEdge:   EnumToStringBrowserList_RegPath = "\Microsoft\Edge"
    End Select
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
