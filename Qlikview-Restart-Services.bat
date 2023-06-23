REM @echo off

REM +---------------------------------------------+
REM |Created by @ale-novoa                        |
REM |2022 - 11 - 23                               |
REM |This batch file uses multiple loops to verify|
REM |that the processes were correctly killed     |
REM |before attempting to stop services           |
REM |And then it verifies if the services were    |
REM |correctly stops, otherwise is kills them     |
REM |                                             |
REM | Important to note:                          |
REM |This script does not stop QVS                |
REM +---------------------------------------------+


setlocal enabledelayedexpansion

set qvbactive=0
set qdsactive=0
set qdscactive=0




REM -- MAIN CODE RUNS HERE - calls to subroutines
:start
CALL :killqvb
CALL :stopqdsservice
CALL :stopqdscservice
timeout 10
CALL :startservices
GOTO exit





REM ##############   DEFINE SUBROUTINES   ##############


REM -- Detect if there are any qvb.exe processes running.
:checkqvb
tasklist /fi "ImageName eq qvb.exe" /fo csv 2>NUL | find /I "qvb.exe">NUL
if "%ERRORLEVEL%"=="0" (
	echo There are qvb processes running
	set qvbactive=1
) else (
	echo there are no qvb.exe running
	set qvbactive=0
)
GOTO :EOF


REM -- Detect if there are any qvb.exe processes running.
:checkqds
tasklist /fi "ImageName eq QVDistributionService.exe" /fo csv 2>NUL | find /I "QVDistributionService.exe">NUL
if "%ERRORLEVEL%"=="0" (
	echo QDS is running
	set qdsactive=1
) else (
	echo QDS is not running
	set qdsactive=0
)
GOTO :EOF


REM -- Detect if there are any qvb.exe processes running.
:checkqdsc
tasklist /fi "ImageName eq QVDirectoryServiceConnector.exe" /fo csv 2>NUL | find /I "QVDirectoryServiceConnector.exe">NUL
if "%ERRORLEVEL%"=="0" (
	echo QDSC is running
	set qdscactive=1
) else (
	echo QDSC is not running
	set qdscactive=0
)
GOTO :EOF


REM -- checks if there are any processes running and kills them on loop untill all dead
:killqvb
CALL :checkqvb
if "%qvbactive%"=="1" (
	TASKKILL /F /IM qvb.exe /T
	timeout 2
	goto killqvb
) 
GOTO :EOF


REM -- Stop the service, if not successfull go to kill process
:stopqdsservice
NET STOP "QlikView Distribution Service"
IF NOT "%ERRORLEVEL%"=="0" (
	CALL :killqdsprocess
)
GOTO :EOF


REM -- kill the QDS process
:killqdsprocess
CALL :checkqds
IF "%qdsactive%"=="1" (
	TASKKILL /F /IM QVDistributionService.exe
	timeout 2
	goto killqdsprocess
)
goto :EOF



REM -- Stop the service, if not successfull go to kill process
:stopqdscservice
NET STOP "QlikView Directory Service Connector"
IF NOT "%ERRORLEVEL%"=="0" (
	CALL :killqdscprocess
)
GOTO :EOF


REM -- kil the QDSC process
:killqdscprocess
CALL :checkqdsc
IF "%qdscactive%"=="1" (
	TASKKILL /F /IM QVDirectoryServiceConnector.exe
	timeout 2
	goto killqdscprocess
)
goto :EOF



REM -- Start the services
:startservices
NET START "QlikView Directory Service Connector"
NET START "QlikView Distribution Service"
GOTO :EOF



REM -- Exit
:exit
echo the services have been successfully restarted
exit










