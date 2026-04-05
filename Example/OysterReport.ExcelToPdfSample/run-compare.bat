@echo off
setlocal

set "ROOT=%~dp0\..\.."
pushd "%ROOT%" >nul

set "DOTNET_CLI_HOME=%ROOT%\.dotnet-home"
if not exist "%DOTNET_CLI_HOME%" mkdir "%DOTNET_CLI_HOME%" >nul 2>&1
set "DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1"
set "DOTNET_CLI_TELEMETRY_OPTOUT=1"

set "INPUT=%ROOT%\Example\OysterReport.ExcelToPdfSample\seikyusyo.xlsx"
set "MAIN_OUT=%ROOT%\Example\OysterReport.ExcelToPdfSample\seikyusyo.pdf"
set "WORK_OUT=%ROOT%\Example\OysterReport.ExcelToPdfSample\seikyusyo.WorkExcelToPdf.pdf"
set "WORK_PROJECT=%ROOT%\__Sandbox\ExcelToPdf\WorkExcelToPdf\WorkExcelToPdf.csproj"
set "TEMP_DIR=%ROOT%\Example\OysterReport.ExcelToPdfSample\.compare-temp"
set "MAIN_INPUT=%TEMP_DIR%\seikyusyo.oyster.xlsx"
set "WORK_INPUT=%TEMP_DIR%\seikyusyo.work.xlsx"

if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%" >nul 2>&1
copy /y "%INPUT%" "%MAIN_INPUT%" >nul
if errorlevel 1 goto :error
copy /y "%INPUT%" "%WORK_INPUT%" >nul
if errorlevel 1 goto :error

dotnet build OysterReport.slnx -v minimal /p:UseSharedCompilation=false /m:1 /nr:false /p:NuGetAudit=false
if errorlevel 1 goto :error

dotnet build "%WORK_PROJECT%" -v minimal /p:ImportDirectoryBuildProps=false /p:ImportDirectoryBuildTargets=false /p:GeneratePackageOnBuild=false /p:NuGetAudit=false
if errorlevel 1 goto :error

if exist "%MAIN_OUT%" del /q "%MAIN_OUT%"
if exist "%WORK_OUT%" del /q "%WORK_OUT%"

dotnet run --project Example\OysterReport.ExcelToPdfSample\OysterReport.ExcelToPdfSample.csproj --no-build --no-restore -- "%MAIN_INPUT%" "%MAIN_OUT%"
if errorlevel 1 goto :error
if not exist "%MAIN_OUT%" goto :error

dotnet run --project "%WORK_PROJECT%" --no-build --no-restore -- "%WORK_INPUT%" "%WORK_OUT%"
if errorlevel 1 goto :error
if not exist "%WORK_OUT%" goto :error

del /q "%MAIN_INPUT%" >nul 2>&1
del /q "%WORK_INPUT%" >nul 2>&1

echo.
echo OysterReport   : %MAIN_OUT%
echo WorkExcelToPdf : %WORK_OUT%
popd >nul
exit /b 0

:error
popd >nul
exit /b %errorlevel%
