VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SettingRunBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub 設定を保存_Click()
    '1. もし画面上のデータが編集されていたら（Dirty ＝ 汚れている状態なら）
    'それを強制的にテーブルに書き込んで保存（Dirty ＝ False）します
    If Me.Dirty Then Me.Dirty = False

    '2. 利用ユーザーを設定
    ShSetting01_StartBrowser.CurrentUserName = Me.UserDataDirName.value

    '3. 完了メッセージを表示して安心感を与える
    MsgBox "現在のユーザー：" & ShSetting01_StartBrowser.CurrentUserName & " に設定されました。", vbInformation, "保存完了"
End Sub
