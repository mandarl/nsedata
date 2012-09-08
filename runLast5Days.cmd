@ECHO OFF
CLS

SET Choice=
SET /P Choice=Enter the number of days (default 5): 
IF '%Choice%'=='' SET Choice=5

SET Ami=
SET /P Ami=Launch Amibroker when done? (default Y): 
IF /I '%Ami%'=='' SET Ami=Y


cscript Equity.vbs -dDelta %Choice%




IF /I '%Ami%'=='N' GOTO EndLabel

start C:\progra~1\AmiBro~1\Broker.exe

:EndLabel

pause