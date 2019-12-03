@ECHO OFF
:BEGIN
ECHO Hey there! This process to stop sql servises ... 
:END
@ECHO OFF
:BEGIN
ECHO  Powered By Ahmed Rashwan-Ocean Soft technical support!
:END
color 1b
net stop MSSQLSERVER /y
net stop SQLSERVERAGENT /y
net stop SQLBrowser /y