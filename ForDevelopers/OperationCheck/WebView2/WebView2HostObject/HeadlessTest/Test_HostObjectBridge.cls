VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_HostObjectBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   `Test_WV2HostObject`用のブリッジClassです(AddHostObjectToScriptでJSへ公開する側)
'
'   Public  : JSから見える想定(メソッド / Property / 変数)
'   Friend  : JSからは見えない想定(VBAの中だけで使う口)
'***************************************************************************************************
Option Explicit

'Public変数(JSからはプロパティとして見える想定)
Public Count As Long

Private m_tag          As String
Private m_value        As Variant
Private m_lastArgTypes As String



'* VBAの中だけで使う口(JSからは見えない想定)
Friend Sub Init(ByVal tag As String)
    m_tag = tag
End Sub

Friend Property Get Tag() As String
    Tag = m_tag
End Property

'* 文字列を受けて、印を付けて返す
Public Function Echo(ByVal s As Variant) As Variant
    Count = Count + 1
    Echo = CStr(s) & "(" & m_tag & ")"
End Function

'* 数値の受け渡し(JSの数値がどの型で届くかも控える)
Public Function Add(ByVal a As Variant, ByVal b As Variant) As Variant
    m_lastArgTypes = TypeName(a) & "," & TypeName(b)
    Add = a + b
End Function

Public Property Get ArgTypes() As Variant
    ArgTypes = m_lastArgTypes
End Property

'* 読み書きできるProperty
Public Property Get Value() As Variant
    Value = m_value
End Property

Public Property Let Value(ByVal v As Variant)
    m_value = v
End Property

'* 解放されたことをテスト側へ知らせる
Private Sub Class_Terminate()
    Test_WV2HostObject.NotifyTerminated m_tag
End Sub
