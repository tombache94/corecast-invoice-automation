@echo off
REM Daily dashboard refresh -- invoked by Windows Task Scheduler.
REM Output goes to data\refresh.log so failures can be inspected after the fact.
cd /d "%~dp0\.."
node scripts\refresh-dashboard.js >> data\refresh.log 2>&1
