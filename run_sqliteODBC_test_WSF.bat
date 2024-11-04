@echo off
REM force unicode support
REM mode con codepage select=65001
setlocal

setlocal EnableDelayedExpansion
set "arg1="
call set "arg1=%%1"

if defined arg1 goto :arg_exists

echo please provide either 32 or 64 as command line argument.
exit /b

:arg_exists

if not exist "C:\temp" ( mkdir "C:\temp" )

:switch-case-example
  :: Call and mask out invalid call targets
  goto :switch-case-N-%arg1% 2>nul || (
    echo please provide either 32 or 64 as command line argument.
  )
  goto :switch-case-end
  
  :switch-case-N-32

    echo echo Running tests as 32 bit process...
    c:\Windows\syswow64\cscript.exe //nologo sqliteODBC_tests.wsf 32
    goto :switch-case-end
    
  :switch-case-N-64
  
    echo echo Running tests as 64 bit process...
    c:\Windows\system32\cscript.exe //nologo sqliteODBC_tests.wsf 64
    goto :switch-case-end
    
:switch-case-end


