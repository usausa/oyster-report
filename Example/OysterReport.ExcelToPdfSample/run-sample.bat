@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..\..") do set "REPO_ROOT=%%~fI"

set "DOTNET_CLI_HOME=%REPO_ROOT%\.dotnet-home"
set "DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1"
set "DOTNET_NOLOGO=1"

pushd "%REPO_ROOT%"
dotnet build OysterReport.slnx -v minimal /p:UseSharedCompilation=false /m:1 /nr:false /p:NuGetAudit=false || goto :fail
dotnet run --project "%SCRIPT_DIR%OysterReport.ExcelToPdfSample.csproj" --no-build --no-restore -- "%SCRIPT_DIR%seikyusyo.xlsx" "%SCRIPT_DIR%seikyusyo.pdf" || goto :fail
popd
exit /b 0

:fail
set "ERR=%errorlevel%"
popd
exit /b %ERR%
