@echo off
echo Installing GoLang...

:: Set version and installation directory
set GO_VERSION=1.21.1
set INSTALL_DIR=C:\Go

:: Download Go installer
echo Downloading Go %GO_VERSION%...
curl -LO https://dl.google.com/go/go%GO_VERSION%.windows-amd64.msi

:: Install Go
echo Installing Go...
msiexec /i go%GO_VERSION%.windows-amd64.msi /quiet /norestart

:: Set up environment variables
setx PATH "%INSTALL_DIR%\bin;%PATH%"

:: Clean up
echo Cleaning up...
del go%GO_VERSION%.windows-amd64.msi

echo GoLang installation completed.
pause
