@ECHO off
setlocal EnableDelayedExpansion
title SETRIX&cls&echo =====================================================================================&echo -=SETRIX=-&echo =====================================================================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2019&echo - Microsoft Office Professional Plus 2019&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Activating Office...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&set i=1&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul||goto notsupported
:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=s8.now.im
if %i% EQU 3 set KMS=s9.now.im
if %i% GTR 3 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul
:ato
echo ============================================================================&echo.&echo.&cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo - MSGuides website: MSGuides.com&echo.&echo - Consider supporting the devs: donate.msguides.com&echo.&echo ============================================================================&echo &if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... &echo Please wait...&echo. & set /a i+=1 & goto skms)
goto halt
:notsupported
echo ============================================================================&echo.&echo Your version of Office is not supported.&echo.&goto halt
:busy
echo ============================================================================&echo.&echo The server is busy at the moment.&echo.
:halt
pause >nul