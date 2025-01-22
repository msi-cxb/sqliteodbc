@echo off
setlocal EnableDelayedExpansion

SET __DEBUGECHO=ECHO
IF NOT DEFINED __DEBUGECHO (SET __DEBUGECHO=REM)

REM bit is either 32 or 64
set "bit="
call set "bit=%%1"


REM install is either true or false
set "install="
call set "install=%%2"

set "ARCH="
set "CLARG="
if defined bit (
    if defined install (
        goto :arg_exists
    )
)
echo usage: buildandinstall.bat [bit (32 or 64) required] [install (true or false) required]
echo example: build 64 bit driver and install as admin
echo     buildandinstall.bat 64 true
goto :exit

:arg_exists

if %bit%==32 (
    set ARCH=x86
) else if %bit%==64 (
    set ARCH=amd64
) else (
    echo please provide either 32 or 64 as command line argument.
    goto :EOF
)

REM setup Visual Studio
CALL :fn_ConfigVisualStudio

REM report the compiler architecture as a check
CALL :fn_GetCompilerArch

REM now that environemnt is set up, go build the code

echo cleaning...

REM clean is reqiured if you are switching between 32 and 64 bit builds
nmake -f sqlite3odbc.mak clean
REM if errorlevel 1 (echo clean error & goto :exit)

echo building...

REM build driver, extension, and commandline utilities
nmake -f sqlite3odbc.mak all 

REM build just the driver
REM nmake -f sqlite3odbc.mak driver
REM nmake -f sqlite3odbc.mak extensions

REM CXB=1 compiles in the tracing
REM nmake -f sqlite3odbc.mak driver CXB=1

if errorlevel 1 (echo build error & goto :exit)

set installDir=.\install\%bit%bit

if %install%==true (
  CALL :fn_CopyFilesToInstall
  CALL :fn_InstallDriver
)

echo complete.

GOTO :exit

REM ***************************************************
:fn_ConfigVisualStudio
    REM
    REM Visual Studio 2017 / 2019 / 2022 / future versions (hopefully)...
    REM
    CALL :fn_TryUseVsWhereExe
    IF NOT DEFINED VSWHEREINSTALLDIR GOTO skip_detectVisualStudio2017
    SET VSVARS32=%VSWHEREINSTALLDIR%\Common7\Tools\VsDevCmd.bat
    IF EXIST "%VSVARS32%" (
            ECHO Using Visual Studio 2017 / 2019 / 2022...
            set CLARG=-arch=%ARCH%
            %__DEBUGECHO% VSVARS32="%VSVARS32%" %CLARG%
            GOTO skip_detectVisualStudio
    )
    :skip_detectVisualStudio2017

    REM
    REM Visual Studio 2015
    REM
    IF NOT DEFINED VS140COMNTOOLS GOTO skip_detectVisualStudio2015
    SET VSVARS32=%VS140COMNTOOLS%..\..\VC\vcvarsall.bat
    IF EXIST "%VSVARS32%" (
        ECHO Using Visual Studio 2015...
        if %bit%==64 (
            set CLARG=x86_%ARCH% 
        ) else (
            SET CLARG=%ARCH% 
        )
        %__DEBUGECHO% VSVARS32="%VSVARS32%" %ARCH%
        GOTO skip_detectVisualStudio
    )
    :skip_detectVisualStudio2015

    REM
    REM Visual Studio 2013
    REM
    IF NOT DEFINED VS120COMNTOOLS GOTO skip_detectVisualStudio2013
    SET VSVARS32=%VS120COMNTOOLS%..\..\VC\vcvarsall.bat
    IF EXIST "%VSVARS32%" (
        ECHO Using Visual Studio 2013...
        if %bit%==64 (
            set CLARG=x86_%ARCH% 
        ) else (
            SET CLARG=%ARCH% 
        )
        %__DEBUGECHO% VSVARS32="%VSVARS32%" %ARCH%
        GOTO skip_detectVisualStudio
    )
    :skip_detectVisualStudio2013

    REM
    REM Visual Studio 2012
    REM
    IF NOT DEFINED VS110COMNTOOLS GOTO skip_detectVisualStudio2012
    SET VSVARS32=%VS140COMNTOOLS%..\..\VC\vcvarsall.bat
    IF EXIST "%VSVARS32%" (
        ECHO Using Visual Studio 2012...
        if %bit%==64 (
            set CLARG=x86_%ARCH% 
        ) else (
            SET CLARG=%ARCH% 
        )
        %__DEBUGECHO% VSVARS32="%VSVARS32%" %ARCH%
        GOTO skip_detectVisualStudio
    )
    :skip_detectVisualStudio2012

    REM
    REM Visual Studio 2010
    REM
    IF NOT DEFINED VS100COMNTOOLS GOTO skip_detectVisualStudio2010
    SET VSVARS32=%VS100COMNTOOLS%..\..\VC\vcvarsall.bat
    IF EXIST "%VSVARS32%" (
        ECHO Using Visual Studio 2010...
        if %bit%==64 (
            set CLARG=x86_%ARCH% 
        ) else (
            SET CLARG=%ARCH% 
        )
        %__DEBUGECHO% VSVARS32="%VSVARS32%" %ARCH%
        GOTO skip_detectVisualStudio
    )
    :skip_detectVisualStudio2010

    REM
    REM NOTE: At this point, the appropriate Visual Studio version should be
    REM       selected.
    REM
    :skip_detectVisualStudio

    SET VSVARS32=%VSVARS32:\\=\%
    %__DEBUGECHO% "%VSVARS32%" %CLARG%
    CALL "%VSVARS32%" %CLARG% 1>nul
    GOTO :EOF

REM ***************************************************
:fn_GetCompilerArch
    set "cl_arch="
    SET _cmd=cl /? 
    FOR /F "delims=" %%G IN ('%_cmd% 2^>^&1 ^| findstr /C:"Version"') DO (
        for %%A in (%%G) do (
            set cl_arch=%%A
        )
    )
    echo cl.exe compiler architectue is %cl_arch%
    GOTO :EOF

REM ***************************************************
:fn_TryUseVsWhereExe
    IF DEFINED VSWHERE_EXE GOTO skip_setVsWhereExe
    SET VSWHERE_EXE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe
    IF NOT EXIST "%VSWHERE_EXE%" SET VSWHERE_EXE=%ProgramFiles%\Microsoft Visual Studio\Installer\vswhere.exe
    :skip_setVsWhereExe

    IF NOT EXIST "%VSWHERE_EXE%" (
        ECHO The "VsWhere" tool does not appear to be installed.
        GOTO :EOF
    ) ELSE (
        %__DEBUGECHO% VSWHERE_EXE="%VSWHERE_EXE%"
    )
    SET VS_WHEREIS_CMD="%VSWHERE_EXE%" -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath -latest
    %__DEBUGECHO% VS_WHEREIS_CMD=%VS_WHEREIS_CMD%

    FOR /F "delims=" %%D IN ('%VS_WHEREIS_CMD%') DO (SET VSWHEREINSTALLDIR=%%D)

    IF NOT DEFINED VSWHEREINSTALLDIR (
        ECHO Visual Studio 2017 / 2019 / 2022 is not installed.
    GOTO :EOF
    )
    %__DEBUGECHO% Visual Studio 2017 / 2019 / 2022 is installed.
    %__DEBUGECHO% VsWhereInstallDir = '%VSWHEREINSTALLDIR%'
    GOTO :EOF
    
REM ***************************************************
:fn_CopyFilesToInstall
    echo copying SQLite3 files to %installDir%...
    if not exist %installDir% ( mkdir  %installDir% )
    xcopy /Q /Y bfsvtab.dll %installDir% 1>nul
    xcopy /Q /Y checkfreelist.dll %installDir% 1>nul
    xcopy /Q /Y crypto.dll %installDir% 1>nul
    xcopy /Q /Y csv.dll %installDir% 1>nul
    xcopy /Q /Y decimal.dll %installDir% 1>nul
    xcopy /Q /Y extension-functions.dll %installDir% 1>nul
    xcopy /Q /Y fileio.dll %installDir% 1>nul
    xcopy /Q /Y ieee754.dll %installDir% 1>nul
    xcopy /Q /Y inst.exe %installDir% 1>nul
    xcopy /Q /Y regexp.dll %installDir% 1>nul
    xcopy /Q /Y series.dll %installDir% 1>nul
    xcopy /Q /Y sha1.dll %installDir% 1>nul
    xcopy /Q /Y shathree.dll %installDir% 1>nul
    xcopy /Q /Y sqlfcmp.dll %installDir% 1>nul
    xcopy /Q /Y totype.dll %installDir% 1>nul
    xcopy /Q /Y uuid.dll %installDir% 1>nul
    xcopy /Q /Y vfsstat.dll %installDir% 1>nul
    xcopy /Q /Y wholenumber.dll %installDir% 1>nul
    xcopy /Q /Y sqldiff.exe %installDir% 1>nul
    xcopy /Q /Y sqlite3.exe %installDir% 1>nul
    xcopy /Q /Y uninst.exe %installDir% 1>nul
    xcopy /Q /Y sqlite3odbc.dll %installDir% 1>nul
    xcopy /Q /Y SQLiteODBCInstaller.exe %installDir% 1>nul
    
    if %bit%==32 (
    SET VC_REDIST=%VCINSTALLDIR%Redist\MSVC\v143\vc_redist.x86.exe
    ) else if %bit%==64 (
    SET VC_REDIST=%VCINSTALLDIR%Redist\MSVC\v143\vc_redist.x64.exe
    )
    echo adding VC_REDIsT installer %VC_REDIST%...
    xcopy /Q /Y "%VC_REDIST%" %installDir% 1>nul
    GOTO :EOF

REM ***************************************************
:fn_InstallDriver
    set REL_PATH=%installDir%
    set ABS_PATH=
    pushd %REL_PATH%
    set ABS_PATH=%CD%
    popd    
    echo installing as admin from %ABS_PATH% to %appdata%\sqlite\%bit%bit
    if not exist %appdata%\sqlite (mkdir %appdata%\sqlite)
    if exist %appdata%\sqlite\%bit%bit (rmdir /S /Q %appdata%\sqlite\%bit%bit)
    echo %appdata%\sqlite\%bit%bit
    mkdir %appdata%\sqlite\%bit%bit
    echo copy %ABS_PATH%\*.dll to %appdata%\sqlite\%bit%bit
    call COPY  %ABS_PATH%\*.dll %appdata%\sqlite\%bit%bit > NUL
    echo copy %ABS_PATH%\*.exe to %appdata%\sqlite\%bit%bit
    call COPY  %ABS_PATH%\*.exe %appdata%\sqlite\%bit%bit > NUL
    REM timeout /t 1 > NUL
    ping -n 1 -w 0.2 10.0.0.1 > NUL
    Powershell Start cmd.exe -ArgumentList "/c","cd",%appdata%\sqlite\%bit%bit,"'&'","SQLiteODBCInstaller.exe","-u","-a","-q","'&'","SQLiteODBCInstaller.exe","-i","-d=sql3","-q" -Verb Runas
    if errorlevel 1 (echo install error & goto exit)
    GOTO :EOF
    
REM ***************************************************
:exit
