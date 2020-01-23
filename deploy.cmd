@if "%SCM_TRACE_LEVEL%" NEQ "4" @echo off

IF "%SITE_ROLE%" == "bot" (
  deploy.bot.cmd
) ELSE (
    echo You have to set SITE_ROLE setting to "bot"
    exit /b 1 
)