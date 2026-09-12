Attribute VB_Name = "Demo_WebDriverBiDiPlus"
Option Explicit



Sub BiDiPlusによる冒険の始まり()
    '設定シートに基づくブラウザ立ち上げ
    Dim HelloWorldAutomationBrowser As WebDriverBiDiContext
    Set HelloWorldAutomationBrowser = ShSetting01_StartBrowser.StartBiDiModeContext

    'BiDi+へUpdate
    Dim bidiPlus As New WebDriverBiDiPlusContext
    bidiPlus.UpgradeBiDiPlus HelloWorldAutomationBrowser
    Set HelloWorldAutomationBrowser = Nothing

    bidiPlus.ExecuteCDP "Network.enable"
    bidiPlus.SubscribeCdpEvent = Array("Network.requestWillBeSent", "Network.loadingFinished")
    bidiPlus.ThisWebDriverBiDiContext.navigate "https://github.com/GoogleChromeLabs/chromium-bidi"
    

    'ブラウザを正常に閉じる
    bidiPlus.ThisWebDriverBiDiContext.ThisWebDriverBiDiMode.quit
End Sub

