@echo off
cd /d "%~dp0"
git add MySchoolChecksPlus/app.ico
git commit -m "ui: new icon - white plus on blue"
git push origin main
echo.
echo Done!
pause
