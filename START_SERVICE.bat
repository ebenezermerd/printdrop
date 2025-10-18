@echo off
echo ============================================
echo     AUTO-PRINT BACKGROUND SERVICE
echo ============================================
echo.
echo Starting enhanced auto-print service...
echo.
echo Features:
echo   - PDF printing (SumatraPDF)
echo   - Word document printing (COM automation)  
echo   - Text/CSV printing (PowerShell)
echo   - Full print management support
echo.
echo Network access: \\DA-SERVER\PrintDrop
echo Target printer: Canon iR2004/2204 UFRII LT
echo.

cd /d "C:\PrintDrop"

echo Service starting in background...
start /min "Auto-Print Service" python auto_print.py

echo.
echo SUCCESS: Auto-print service started in background!
echo.
echo The service is now running. You can:
echo   - Drop files into \\DA-SERVER\PrintDrop from any network device
echo   - Files will automatically print within 5 seconds
echo   - Check the minimized window for service logs
echo.
echo To stop the service, close the minimized Python window.
echo.
echo Press any key to exit this setup window...
pause >nul