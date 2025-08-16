@echo off
REM Installer script for Khaled's Attendance Tracker
echo ================================================================
echo Installing Khaled's Attendance Tracker
echo ================================================================

REM Get current directory
set "SOURCE_DIR=%~dp0"

REM Create application directory in user's local folder
set "INSTALL_DIR=%USERPROFILE%\AppData\Local\Khaleds_Attendance_Tracker"

echo Creating installation directory...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo Copying application files...
copy "%SOURCE_DIR%Khaleds_Attendance_Tracker.exe" "%INSTALL_DIR%\" /Y

REM Copy web templates if they exist
if exist "%SOURCE_DIR%web_templates" (
    echo Copying web templates...
    xcopy "%SOURCE_DIR%web_templates" "%INSTALL_DIR%\web_templates" /E /I /Y >nul
)

REM Copy assets if they exist
if exist "%SOURCE_DIR%assets" (
    echo Copying assets...
    xcopy "%SOURCE_DIR%assets" "%INSTALL_DIR%\assets" /E /I /Y >nul
)

echo Creating desktop shortcut...
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%USERPROFILE%\Desktop\Khaleds Attendance Tracker.lnk'); $s.TargetPath='%INSTALL_DIR%\Khaleds_Attendance_Tracker.exe'; $s.WorkingDirectory='%INSTALL_DIR%'; $s.IconLocation='%INSTALL_DIR%\Khaleds_Attendance_Tracker.exe'; $s.Description='Professional Time Management Solution'; $s.Save()"

echo Creating start menu shortcut...
if not exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Khaled's Software" mkdir "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Khaled's Software"
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%APPDATA%\Microsoft\Windows\Start Menu\Programs\Khaleds Software\Khaleds Attendance Tracker.lnk'); $s.TargetPath='%INSTALL_DIR%\Khaleds_Attendance_Tracker.exe'; $s.WorkingDirectory='%INSTALL_DIR%'; $s.IconLocation='%INSTALL_DIR%\Khaleds_Attendance_Tracker.exe'; $s.Description='Professional Time Management Solution'; $s.Save()"

echo.
echo ================================================================
echo Installation Complete!
echo ================================================================
echo.
echo Khaled's Attendance Tracker has been installed to:
echo %INSTALL_DIR%
echo.
echo Shortcuts created:
echo - Desktop: Khaleds Attendance Tracker
echo - Start Menu: Khaleds Software ^> Khaleds Attendance Tracker
echo.
echo Features:
echo - Modern desktop interface with PyQt5
echo - Mobile PWA access for iPhone/Android
echo - Word and PDF document processing
echo - Automatic system tray integration
echo - Professional branding and splash screen
echo.
echo You can now run the application from your desktop or start menu!
echo.
pause
