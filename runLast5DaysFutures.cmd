@ECHO OFF
CLS

SET Choice=
SET /P Choice=Enter the number of days (default 5): 
IF '%Choice%'=='' SET Choice=5

SET Ami=
SET /P Ami=Launch Amibroker when done? (default Y): 
IF /I '%Ami%'=='' SET Ami=Y


cscript Futures.vbs -dDelta %Choice% -DBPath D:\PROGRA~1\AMIBRO~2\DataF




IF /I '%Ami%'=='N' GOTO EndLabel

start D:\progra~1\AmiBro~2\Broker.exe

:EndLabel
pause