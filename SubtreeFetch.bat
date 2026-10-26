@echo OFF

cd /d %~dp0

REM GitçXêV
git fetch https://github.com/Eschamali/StarterWebScrapingKit.git main
git subtree pull --prefix=Original/Chromium-Automation-with-CDP-for-VBA https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA.git main
git subtree pull --prefix=Original/VBA-WEB https://github.com/VBA-Tools-v2/VBA-Web.git master
git subtree pull --prefix=Original/VBA-FastJSON https://github.com/cristianbuse/VBA-FastJSON.git master
git subtree pull --prefix=Original/vbacollective-json https://github.com/vbacollective/json.git main
git subtree pull --prefix=Original/VBA-FastDictionary https://github.com/cristianbuse/VBA-FastDictionary.git master
