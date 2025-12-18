@echo OFF
chcp 65001

cd /d %~dp0

SET FromPath=Original\ChromeControler-No-Selenium-WebDriver-VBAJSON\src\
SET ToPATH=VBAProject\Class\
copy /y %FromPath%a1x1_HTTPCommunicator.cls %ToPATH%WebSocketHTTPCommunicator.cls
copy /y %FromPath%a1_WebSocketCommunicator.cls %ToPATH%WebSocketCommunicator.cls


SET FromPath=Original\Chromium-Automation-with-CDP-for-VBA\src\
copy /y %FromPath%*.cls %ToPATH%


SET FromPath=Original\SeleniumVBA\src\VBA\
copy /y %FromPath%WebJsonConverter.cls %ToPATH%CDPJConv.cls


SET FromPath=Original\VBA-WEB\src\vbProject\VBA-Web\
copy /y %FromPath%*.cls %ToPATH%
copy /y %FromPath%Helpers\*.cls %ToPATH%
SET ToPATH=VBAProject\Modules\
copy /y %FromPath%Helpers\*.bas %ToPATH%
