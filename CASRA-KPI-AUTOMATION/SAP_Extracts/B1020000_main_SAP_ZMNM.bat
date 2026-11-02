:: Changing directory to where the .bat is located
::cd /C "%~dp0"
 
 
@echo off
echo BAT location: "%~dp0"
setlocal
 
:: Specify the path to your configuration file
set "configFile=config.txt"
 
:: Read the configuration file and set environment variables
for /f "usebackq tokens=1,2 delims==" %%A in ("%configFile%") do (
    set "%%A=%%B"
)
 
%python% Main_SAP_ZMNM.py %username% %password%
 
@pause

endlocal