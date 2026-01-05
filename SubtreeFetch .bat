@echo OFF
chcp 65001

cd /d %~dp0

git fetch https://github.com/Eschamali/StarterWebScrapingKit.git OtherThanWebSocket
git subtree pull --prefix=Original/Chromium-Automation-with-CDP-for-VBA https://github.com/GCuser99/Chromium-Automation-with-CDP-for-VBA.git main
git subtree pull --prefix=Original/ChromeControler-No-Selenium-WebDriver-VBAJSON https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON.git master
git subtree pull --prefix=Original/VBA-WEB https://github.com/VBA-Tools-v2/VBA-Web.git master
git subtree pull --prefix=Original/SeleniumVBA https://github.com/GCuser99/SeleniumVBA.git main
