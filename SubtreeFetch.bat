@echo OFF

cd /d %~dp0

REM Git更新
git fetch https://github.com/Eschamali/StarterWebScrapingKit.git main
git subtree pull --prefix=Original/Chromium-Automation-with-CDP-for-VBA https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA.git main
git subtree pull --prefix=Original/ChromeControler-No-Selenium-WebDriver-VBAJSON https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON.git master
git subtree pull --prefix=Original/VBA-WEB https://github.com/VBA-Tools-v2/VBA-Web.git master
git subtree pull --prefix=Original/VBA-FastJSON https://github.com/cristianbuse/VBA-FastJSON.git master
git subtree pull --prefix=Original/vbacollective-json https://github.com/vbacollective/json.git main

REM asset内の更新
curl https://cdn.jsdelivr.net/npm/chromium-bidi@latest/lib/iife/mapperTab.js > assset\mapperTab.js
curl https://data.jsdelivr.com/v1/package/npm/chromium-bidi
