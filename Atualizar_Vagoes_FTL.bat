@echo off
chcp 65001 >nul
echo ============================================================================
echo   AUTOMACAO DE CERTIFICADOS DE VAGOES-TANQUE FTL (PORTAL LOCARGA CSN)
echo ============================================================================
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0Atualizar_Vagoes_FTL.ps1"
echo.
pause

