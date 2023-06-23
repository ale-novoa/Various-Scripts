@echo off

REM ┌──────────────────────────────────────────────────┐
REM │This batch will check if I am connected           │
REM │to the Network I need (by pinging a specific IP)  │
REM │If so, trigger actions                            │
REM │If not, wait some seconds (only x number of times)│
REM └──────────────────────────────────────────────────┘

setlocal enabledelayedexpansion

set connected=0
set looptimes=0
set servertoping=hostname


REM -- MAIN CODE RUNS HERE - calls to subroutines
:start
CALL :startup
GOTO exit




REM ##############   DEFINE SUBROUTINES   ##############


REM -- Detect if we are in the Network I need.
:checkconnection
ping -n 1 %servertoping% | find "TTL=" >nul
if "%ERRORLEVEL%"=="0" (
	echo We are able to ping, therefore within the network
	set connected=1
) else (
	echo unable to ping, therefore not in the network
	set connected=0
)
GOTO :EOF


REM -- list of websites to open on startup
:launchwebsites
start http://hostname
timeout 10
start http://Second-Intra-Website
timeout 10
start https://jira/issues/
timeout 10
start https://www.worldtimebuddy.com/
GOTO :EOF

REM -- list of apps to open on startup
:launchapps
start "" "C:\Program Files\Notepad++\notepad++.exe"
start "" "C:\Program Files\OtherApp\OtherApp.exe"
GOTO :EOF


REM -- main subroutine to run
:startup
CALL :checkconnection
if "%connected%"=="1" (
	CALL :launchwebsites
	CALL :launchapps
) else (
	set /A looptimes=looptimes+1
	if %looptimes% LSS 7 (
		timeout 10
		goto startup
	)
)
GOTO :EOF

REM -- Exit
:exit
if "%connected%"=="1" (
	echo startup has been performed, exiting batch
) else (
	echo limit reached, no apps have been launched, exiting batch
)	
exit
	





