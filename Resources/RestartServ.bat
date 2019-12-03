@ECHO OFF
:BEGIN
ECHO Hey there! This process to Restart sql servises ... 
:END
@ECHO OFF
:BEGIN
ECHO  Powered By Ahmed Rashwan-Ocean Soft technical support!
:END
color 1b
net stop MSSQLSERVER /y
net stop SQLSERVERAGENT /y
net stop SQLBrowser /y
net start MSSQLSERVER /y
net Start SQLSERVERAGENT /y
net Start SQLBrowser /y