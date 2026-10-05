VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SettingRunBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'             CDP/WebDriverBiDi経由で、ブラウザ操作する際の準備関連のAcees向けフォームです
'***************************************************************************************************
Option Compare Database
Option Explicit



'***************************************************************************************************
'* 機能　　：フォームを開いた時点での、利用ユーザー名をセットします
'***************************************************************************************************
Private Sub Form_Load()
    '利用ユーザーを設定
    ShSetting01_StartBrowser.CurrentUserName = Me.UserDataDirName.value
End Sub

'***************************************************************************************************
'* 機能　　：保存ボタン押下時の、利用ユーザー名をセットし、設定値も保存します
'***************************************************************************************************
Private Sub ButtonSave_Click()
    '1. もし画面上のデータが編集されていたら（Dirty ＝ 汚れている状態なら）
    'それを強制的にテーブルに書き込んで保存（Dirty ＝ False）します
    If Me.Dirty Then Me.Dirty = False

    '2. 利用ユーザーを設定
    ShSetting01_StartBrowser.CurrentUserName = Me.UserDataDirName.value

    '3. 完了メッセージを表示して安心感を与える
    MsgBox "現在のユーザー：" & ShSetting01_StartBrowser.CurrentUserName & " に設定されました。", vbInformation, "保存完了"
End Sub
