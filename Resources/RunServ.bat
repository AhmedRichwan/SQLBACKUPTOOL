@ECHO OFF
:BEGIN
ECHO Hey there! This process to activate sql servises upen stopping... 
:END
@ECHO OFF
:BEGIN
ECHO  Powered By Ahmed Rashwan-Ocean Soft technical support!
:END
color 1b
net start MSSQLSERVER /y
net Start SQLSERVERAGENT /y
net Start SQLBrowser /y