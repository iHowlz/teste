@echo off
title Microsoft Activation Scripts (Windows HWID + Office KMS)
color 0a
mode con: cols=75 lines=25

:: Verifica se está sendo executado como administrador
NET FILE >nul 2>&1
if '%errorlevel%' NEQ '0' (
    echo.
    echo ERROR: Este script requer privil•gios de administrador.
    echo Execute como Administrador.
    pause >nul
    exit /b
)

:: Verifica se o sistema ‚ de 64-bit
if exist "%ProgramFiles(x86)%" (
    set "arch=64-bit"
) else (
    set "arch=32-bit"
)

:: Menu principal
:menu
cls
echo.
echo  =============================================
echo     Microsoft Activation Scripts (MAS)       
echo  =============================================
echo.
echo  [1] Ativar Windows com HWID (Licen‡a Digital)
echo  [2] Ativar Office com KMS (180 dias)
echo.
echo  [0] Sair
echo.
set /p choice="Escolha uma op‡Æo [0-2]: "

if "%choice%"=="1" goto HWID
if "%choice%"=="2" goto OfficeKMS
if "%choice%"=="0" exit
goto menu

:: =============================================
:: Ativa‡Æo do Windows via HWID (Digital License)
:: =============================================
:HWID
cls
echo.
echo  =============================================
echo       ATIVAR WINDOWS (HWID - LICEN•A DIGITAL)
echo  =============================================
echo.
echo  Este m‚todo ativa o Windows permanentemente
echo  vinculando uma licen‡a ao hardware do seu PC.
echo.
echo  [*] Processando ativa‡Æo...
echo.
cscript //nologo %SystemRoot%\System32\slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99 >nul
cscript //nologo %SystemRoot%\System32\slmgr.vbs /ipk VK7JG-NPHTM-C97JM-9MPGT-3V66T >nul
cscript //nologo %SystemRoot%\System32\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX >nul
cscript //nologo %SystemRoot%\System32\slmgr.vbs /skms s8.uk.to >nul
cscript //nologo %SystemRoot%\System32\slmgr.vbs /ato >nul
echo  [✓] Windows ativado com sucesso (Licen‡a Digital).
echo.
pause
goto menu

:: =============================================
:: Ativa‡Æo do Office via KMS (180 dias)
:: =============================================
:OfficeKMS
cls
echo.
echo  =============================================
echo       ATIVAR OFFICE (KMS - 180 DIAS)
echo  =============================================
echo.
echo  Este m‚todo ativa o Office usando um servidor KMS.
echo  A ativa‡Æo precisa ser renovada a cada 180 dias.
echo.
echo  [*] Procurando instala‡Æo do Office...
echo.
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    set OfficePath="%ProgramFiles%\Microsoft Office\Office16"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
    set OfficePath="%ProgramFiles(x86)%\Microsoft Office\Office16"
) else (
    echo  [✗] Microsoft Office nÆo encontrado.
    echo.
    pause
    goto menu
)

echo  [*] Ativando Microsoft Office...
echo.
cscript //nologo %OfficePath%\ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul
cscript //nologo %OfficePath%\ospp.vbs /sethst:s8.uk.to >nul
cscript //nologo %OfficePath%\ospp.vbs /act >nul
echo  [✓] Office ativado via KMS (vÆlido por 180 dias).
echo.
pause
goto menu