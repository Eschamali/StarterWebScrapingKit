Attribute VB_Name = "CDP01_ExcelName"
'***************************************************************************************************
'                          CDPに関する名前定義のみをまとめたモジュールです
'                   シートオブジェクト：Sh99_Setting_StartBrowser　に定義しています
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                 ■■■ 特定シート向けセル名 ■■■
'***************************************************************************************************
Public Const BrowserSetting_RangeID01 As String = "ブラウザexeパス"
Public Const BrowserSetting_RangeID02 As String = "ユーザーデータフォルダ名"
Public Const BrowserSetting_RangeID03 As String = "追加の起動引数一式"
Public Const BrowserSetting_RangeID04 As String = "Chromeで使用"
Public Const BrowserSetting_RangeID05 As String = "既存のデバッグプロセスを使用"
Public Const BrowserSetting_RangeID06 As String = "常にブラウザを新規で起動"



'***************************************************************************************************
'                                   ■■■ テーブル名 ■■■
'***************************************************************************************************
Public Const BrowserSetting_TableID01 As String = "PreviousHandleInfo"
Public Const BrowserSetting_TableID02 As String = "PreviousBrowserInfo"
