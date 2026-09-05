@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "SCRIPT_DIR=%~dp0"
set "SOURCE_DIR="
set "ADDON_NAME="
set "ADDON_VERSION="
set "ADDON_TYPE=wps"

if exist "%SCRIPT_DIR%manifest.xml" if exist "%SCRIPT_DIR%index.html" (
  set "SOURCE_DIR=%SCRIPT_DIR%"
)

if not defined SOURCE_DIR (
  for /d %%D in ("%SCRIPT_DIR%..\release\Banyan_*") do (
    if exist "%%~fD\manifest.xml" if exist "%%~fD\index.html" if not defined SOURCE_DIR (
      set "SOURCE_DIR=%%~fD"
    )
  )
)

if not defined SOURCE_DIR (
  for /d %%D in ("%SCRIPT_DIR%..\release\*") do (
    if exist "%%~fD\manifest.xml" if exist "%%~fD\index.html" if not defined SOURCE_DIR (
      set "SOURCE_DIR=%%~fD"
    )
  )
)

if not defined SOURCE_DIR (
  echo [ERROR] 找不到可安装的发布目录。
  echo 请确认本脚本与 release\Banyan_版本号 目录一起分发，或者将脚本放在发布目录中。
  pause
  exit /b 1
)

set "MANIFEST_PATH=%SOURCE_DIR%\manifest.xml"
if not exist "%MANIFEST_PATH%" (
  echo [ERROR] 找不到 manifest.xml。
  pause
  exit /b 1
)

for /f "tokens=2 delims=>" %%A in ('findstr /i /c:"<Name>" "%MANIFEST_PATH%"') do (
  if not defined ADDON_NAME set "ADDON_NAME=%%A"
)
for /f "tokens=2 delims=>" %%A in ('findstr /i /c:"<ApiVersion>" "%MANIFEST_PATH%"') do (
  if not defined ADDON_VERSION set "ADDON_VERSION=%%A"
)
if defined ADDON_NAME set "ADDON_NAME=!ADDON_NAME:</Name=!"
if defined ADDON_VERSION set "ADDON_VERSION=!ADDON_VERSION:</ApiVersion=!"

if not defined ADDON_NAME (
  echo [ERROR] 无法从 manifest.xml 读取插件名称。
  pause
  exit /b 1
)

if not defined ADDON_VERSION (
  echo [ERROR] 无法从 manifest.xml 读取插件版本。
  pause
  exit /b 1
)

if not defined APPDATA (
  echo [ERROR] 找不到 APPDATA 环境变量。
  pause
  exit /b 1
)

set "ADDON_PATH=%APPDATA%\kingsoft\wps\jsaddons"
set "TARGET_DIR=%ADDON_PATH%\%ADDON_NAME%_%ADDON_VERSION%"
set "PUBLISH_XML=%ADDON_PATH%\publish.xml"
set "TEMP_PUBLISH=%ADDON_PATH%\publish.xml.tmp"

echo 安装目录：%TARGET_DIR%
echo 发布目录：%SOURCE_DIR%

if not exist "%ADDON_PATH%" mkdir "%ADDON_PATH%" >nul 2>nul

if exist "%TARGET_DIR%" (
  rmdir /s /q "%TARGET_DIR%"
)

robocopy "%SOURCE_DIR%" "%TARGET_DIR%" /E /NFL /NDL /NJH /NJS /NP /R:1 /W:1 >nul
if errorlevel 8 (
  echo [ERROR] 复制发布目录失败。
  if exist "%TEMP_PUBLISH%" del /f /q "%TEMP_PUBLISH%" >nul 2>nul
  pause
  exit /b 1
)

if exist "%PUBLISH_XML%" (
  findstr /i /c:"<jsplugins>" "%PUBLISH_XML%" >nul
  if errorlevel 1 (
    echo [ERROR] publish.xml 根节点必须是 ^<jsplugins^>。
    if exist "%TEMP_PUBLISH%" del /f /q "%TEMP_PUBLISH%" >nul 2>nul
    pause
    exit /b 1
  )

  >"%TEMP_PUBLISH%" (
    echo ^<?xml version="1.0" encoding="UTF-8"?^>
    echo ^<jsplugins^>
    findstr /v /i /c:"<?xml" /c:"<jsplugins>" /c:"</jsplugins>" /c:"name=\"%ADDON_NAME%\"" "%PUBLISH_XML%"
    echo(  ^<jsplugin name="%ADDON_NAME%" type="%ADDON_TYPE%" url="%ADDON_NAME%_%ADDON_VERSION%" version="%ADDON_VERSION%" enable="enable_dev" install="null" customDomain="" /^>
    echo ^</jsplugins^>
  )
  move /y "%TEMP_PUBLISH%" "%PUBLISH_XML%" >nul
) else (
  >"%PUBLISH_XML%" (
    echo ^<?xml version="1.0" encoding="UTF-8"?^>
    echo ^<jsplugins^>
    echo(  ^<jsplugin name="%ADDON_NAME%" type="%ADDON_TYPE%" url="%ADDON_NAME%_%ADDON_VERSION%" version="%ADDON_VERSION%" enable="enable_dev" install="null" customDomain="" /^>
    echo ^</jsplugins^>
  )
)

echo.
echo 安装完成，请重启 WPS Office 后启用插件。
pause
exit /b 0