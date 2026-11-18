@echo off
chcp 65001 >nul
title DataScrub ^& Variance Studio (Port 4195)
echo Starting DataScrub ^& Variance Studio on http://localhost:4195 ...
node --experimental-sqlite --no-warnings server.js
pause
