# VC++ 2015/2017/2019/2022 Makefile
# uses the SQLite3 amalgamation source which must
# be unpacked below in the same folder as this makefile

# --- VARIABLES ---------------------------------------------------------------
CC=		cl
LN=		link
RC=		rc
LTLIB = lib

!IF "$(DEBUG)" == "1"
LDEBUG=		/DEBUG
CDEBUG=		-Zi
!ELSE
LDEBUG=		/RELEASE
!ENDIF

!IF "$(CXB)" == "1"
__CXB= -D__CXB
!ENDIF

# -O2 maybe 5% faster...
# MAXPERF=    -O2 -Ot
# A subset of /O2 that doesn't include /GF or /Gy.
MAXPERF=    -Ox -Ot

# this ensures that .dll and .exe files built multiple times without
# changes to the code will result in exact same binary output 
DETERMINISTIC= /Brepro

# https://dev.to/yumetodo/list-of-mscver-and-mscfullver-8nd
# this will report compiler version in SQLite using 'pragma compile_options;' query
!IFNDEF MSVC_VER
MSVC_VER=_MSC_VER
!ENDIF

# these flags are specific to the ODBC dll
CFLAGS= -I. \
         $(MAXPERF)  \
        -EHsc  \
        $(__CXB)  \
        -D_DLL  \
        -nologo  \
        $(CDEBUG)  \
        $(DETERMINISTIC) \
        -DSQLITE_THREADSAFE=0 \
        -DSQLITE_OS_WIN=1 \
        -DHAVE_SQLITE3COLUMNTABLENAME=1 \
        -DHAVE_SQLITE3PREPAREV2=1 \
        -DHAVE_SQLITE3VFS=1 \
        -DHAVE_SQLITE3LOADEXTENSION=1 \
        -DSQLITE_ENABLE_COLUMN_METADATA=1 \
        -DSQLITE_SOUNDEX=1 \
        -DSQLITE_TEMP_STORE=1 \
        -DHAVE_SQLSETPOSIROW \
        -DHAVE_SQLLEN \
        -DHAVE_SQLULEN \
        -DHAVE_SQLROWCOUNT \
        -DSQLITE_WIN32_UTF8_CONVERSION \
        -DSQLITE_WIN32_MALLOC \
        -DSQLITE_ODBC_WIN32_MALLOC \
        -DSQLITE_ENABLE_EXPLAIN_COMMENTS \
        -DSQLITE_INTROSPECTION_PRAGMAS \
        -DSQLITE_ENABLE_UNKNOWN_SQL_FUNCTION \
        -DSQLITE_ENABLE_STMTVTAB \
        -DSQLITE_ENABLE_DBPAGE_VTAB \
        -DSQLITE_ENABLE_DBSTAT_VTAB \
        -DSQLITE_ENABLE_JSON1 \
        -DSQLITE_ENABLE_RTREE \
        -DSQLITE_ENABLE_FTS4 \
        -DSQLITE_ENABLE_FTS5 \
        -DSQLITE_ENABLE_MATH_FUNCTIONS \
        -DSQLITE_ENABLE_GEOPOLY \
        -DSQLITE_ENABLE_STAT4 \
        -DHAVE_SQLITETRACE=1 \
        -DBYTE_ORDER=LITTLE_ENDIAN \
        -DWITHOUT_SHELL=1
 
        
CFLAGSEXE= \
        -I. \
        $(MAXPERF) \
        -EHsc \
        -nologo \
        $(CDEBUG) \
        $(DETERMINISTIC)

DLLLFLAGS=$(LDEBUG) \
        $(DETERMINISTIC) \
        /NOLOGO \
        /SUBSYSTEM:WINDOWS \
        /DLL \
        /NODEFAULTLIB

DLLLIBS=msvcrt.lib \
        vcruntime.lib \
        ucrt.lib \
        odbc32.lib \
        odbccp32.lib \
        legacy_stdio_definitions.lib \
        legacy_stdio_wide_specifiers.lib \
        kernel32.lib \
        user32.lib \
        comdlg32.lib

!IF "$(VSCMD_ARG_TGT_ARCH)" == "x64"
PLATFORM = x64
!ELSE
PLATFORM = x86
!ENDIF
LTLIBOPTS=/NOLOGO /MACHINE:$(PLATFORM)

# --- EMBEDDED EXTENSIONS -----------------------------------------------------
# Space-separated list of .c filenames to compile into sqlite3odbc.dll.
# Each file must reside in this directory. Override on the command line:
#   nmake -f sqlite3odbc.mak driver EMBED_EXTENSIONS="eval.c series.c"
# To build with no embedded extensions:
#   nmake -f sqlite3odbc.mak driver EMBED_EXTENSIONS=
EMBED_EXTENSIONS = bfsvtab.c \
        checkfreelist.c \
        crypto.c \
        csv.c \
        decimal.c \
        eval.c \
        extension-functions.c \
        fileio.c \
        ieee754.c \
        regexp.c \
        series.c \
        sha1.c \
        shathree.c \
        sqlfcmp.c \
        sqlite-path.c \
        totype.c \
        uuid.c \
        uint.c \
        wholenumber.c
        # vfsstat.c must NOT be embedded: its init returns SQLITE_OK_LOAD_PERMANENTLY
        # (256), which the auto_extension processing loop treats as a non-zero return
        # and aborts -- any extension registered after vfsstat would never be called.
# Companion .c files required by entries in EMBED_EXTENSIONS.
# These are compiled with CFLAGS_EMBEDDED and linked into sqlite3odbc.dll
# but do NOT register an init function -- they just supply symbols that the
# extension objects reference.  Example:
#   sqlfcmp.c needs cwalk.c
#   crypto.c  needs md5.c shaone.c shatwo.c
# Add entries here whenever an embedded extension fails to link due to
# unresolved symbols that live in a separate source file.
EMBED_EXT_DEPS = cwalk.c \
        md5.c \
        shaone.c \
        shatwo.c
# Same flags as CFLAGS plus -DSQLITE_CORE.  The -DSQLITE_CORE flag causes
# SQLITE_EXTENSION_INIT1 and SQLITE_EXTENSION_INIT2 to expand to nothing so
# embedded extensions call SQLite functions directly instead of through the
# sqlite3_api_routines pointer table.
CFLAGS_EMBEDDED = $(CFLAGS) -DSQLITE_CORE
# Derive embedded object file list via NMAKE suffix substitution.
# csv.c -> csv_emb.obj, eval.c -> eval_emb.obj, etc.
# The _emb suffix distinguishes these from standalone extension objects
# (which compile without -DSQLITE_CORE and cannot share .obj names).
EMBED_EXT_OBJS = $(EMBED_EXTENSIONS:.c=_emb.obj)
EMBED_EXT_DEP_OBJS = $(EMBED_EXT_DEPS:.c=_dep.obj)
# Compile rules for each _emb.obj and _dep.obj file are in ext_embed_rules.mak.
# That file is committed with defaults matching EMBED_EXTENSIONS above.
# After changing EMBED_EXTENSIONS, run: nmake -f sqlite3odbc.mak setup-embed
# to regenerate ext_embed_rules.mak before rebuilding the driver.
!INCLUDE ext_embed_rules.mak

# --- TOP-LEVEL TARGETS -------------------------------------------------------
all:    driver \
        exe \
        extensions \
        sqlite3.dll
        
driver: sqlite3odbc.dll \
        inst.exe \
        uninst.exe \
        adddsn.exe \
        remdsn.exe \
        addsysdsn.exe \
        remsysdsn.exe \
        SQLiteODBCInstaller.exe

exe:    sqlite3.exe \
        sqldiff.exe

extensions: bfsvtab.dll \
        checkfreelist.dll \
        crypto.dll \
        csv.dll \
        decimal.dll \
        eval.dll \
        extension-functions.dll \
        fileio.dll \
        ieee754.dll \
        path.dll \
        regexp.dll \
        series.dll \
        sha1.dll \
        shathree.dll \
        sqlfcmp.dll \
        totype.dll \
        uuid.dll \
        uint.dll \
        vfsstat.dll \
        wholenumber.dll

# copy driver to %APPDATA%\sqlite\64bit and install via elevated UAC prompt
install: SQLiteODBCInstaller.exe sqlite3odbc.dll extensions exe
	if not exist %APPDATA%\sqlite\64bit mkdir %APPDATA%\sqlite\64bit
	copy /Y sqlite3odbc.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y SQLiteODBCInstaller.exe %APPDATA%\sqlite\64bit 1>nul
	copy /Y bfsvtab.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y checkfreelist.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y crypto.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y csv.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y decimal.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y eval.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y extension-functions.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y fileio.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y ieee754.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y path.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y regexp.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y series.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y sha1.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y shathree.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y sqlfcmp.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y totype.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y uuid.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y uint.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y vfsstat.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y wholenumber.dll %APPDATA%\sqlite\64bit 1>nul
	copy /Y sqlite3.exe %APPDATA%\sqlite\64bit 1>nul
	copy /Y sqldiff.exe %APPDATA%\sqlite\64bit 1>nul
	Powershell Start cmd.exe -ArgumentList "/c","cd",%APPDATA%\sqlite\64bit,"'&'","SQLiteODBCInstaller.exe","-u","-a","-q","'&'","SQLiteODBCInstaller.exe","-i","-d=sql3","-q" -Verb Runas

uninstall: SQLiteODBCInstaller.exe
	if not exist %APPDATA%\sqlite\64bit mkdir %APPDATA%\sqlite\64bit
	copy /Y SQLiteODBCInstaller.exe %APPDATA%\sqlite\64bit 1>nul
	Powershell Start cmd.exe -ArgumentList "/c","cd",%APPDATA%\sqlite\64bit,"'&'","SQLiteODBCInstaller.exe","-u","-a","-q" -Verb Runas

clean:
        del *.obj
        del *.res
        del *.exp
        del *.ilk
        del *.pdb
        del *.res
        del resource3.h
        del *.dll
        del *.lib
        del *.exe
        if exist ext_init.c del ext_init.c

# --- RESOURCE PIPELINE -------------------------------------------------------

.c.obj:
		$(CC) $(CFLAGS) /c $<

uninst.exe:	inst.exe
		copy inst.exe uninst.exe

inst.exe:	inst.c
		$(CC) $(CFLAGSEXE) inst.c $(DLLLIBS)

remdsn.exe:	adddsn.exe
		copy adddsn.exe remdsn.exe

adddsn.exe:	adddsn.c
		$(CC) $(CFLAGSEXE) adddsn.c $(DLLLIBS)

remsysdsn.exe:	adddsn.exe
		copy adddsn.exe remsysdsn.exe

addsysdsn.exe:	adddsn.exe
		copy adddsn.exe addsysdsn.exe

fixup.exe:	fixup.c
		$(CC) $(CFLAGSEXE) fixup.c

mkopc3.exe:	mkopc3.c
		$(CC) $(CFLAGSEXE) mkopc3.c

SQLiteODBCInstaller.exe:	SQLiteODBCInstaller.c
        $(CC) $(CFLAGSEXE) \
        SQLiteODBCInstaller.c \
        kernel32.lib \
        user32.lib

sqlite3odbc.c:	resource3.h sqlite3odbc.h sqliteodbc_sqltypes.h

sqlite3odbc.res:	sqlite3odbc.rc resource3.h
		$(RC) -I. -fo sqlite3odbc.res -r sqlite3odbc.rc

OBJECTS=	sqlite3odbc.obj sqlite3.obj $(EMBED_EXT_OBJS) $(EMBED_EXT_DEP_OBJS) ext_init.obj

# Generate both ext_init.c and ext_embed_rules.mak from EMBED_EXTENSIONS.
# Run this target after changing EMBED_EXTENSIONS to update the rules file
# before rebuilding the driver: nmake -f sqlite3odbc.mak setup-embed
setup-embed: gen_embed_init.ps1
	powershell -NoProfile -ExecutionPolicy Bypass -File gen_embed_init.ps1 \
	-ExtensionFiles "$(EMBED_EXTENSIONS)" \
	-DepFiles "$(EMBED_EXT_DEPS)" \
	-OutputFile ext_init.c \
	-RulesFile ext_embed_rules.mak
# Generate extension registration source from EMBED_EXTENSIONS list.
# Also regenerates ext_embed_rules.mak so compile rules stay in sync.
ext_init.c: gen_embed_init.ps1 sqlite3odbc.mak
	powershell -NoProfile -ExecutionPolicy Bypass -File gen_embed_init.ps1 \
	-ExtensionFiles "$(EMBED_EXTENSIONS)" \
	-DepFiles "$(EMBED_EXT_DEPS)" \
	-OutputFile ext_init.c \
	-RulesFile ext_embed_rules.mak
ext_init.obj: ext_init.c
	$(CC) $(CFLAGS) /c ext_init.c
sqlite3odbc.dll:	$(OBJECTS) sqlite3odbc.res
		$(LN) $(DLLLFLAGS) $(OBJECTS) sqlite3odbc.res \
		-def:sqlite3odbc.def -out:$@ $(DLLLIBS)

sqlite3.dll: sqlite3.c
		$(CC) $(CFLAGS) -DSQLITE_API=__declspec(dllexport) sqlite3.c -link -dll -out:sqlite3.dll $(DLLLIBS)
VERSION_C:	fixup.exe VERSION
		.\fixup < VERSION > VERSION_C . ,

resource3.h:	resource.h.in VERSION_C fixup.exe
        .\fixup < resource.h.in > resource3.h \
        --VERS-- @VERSION \
        --VERS_C-- @VERSION_C

# --- EXTENSION DLLS ----------------------------------------------------------

extension-functions.dll:	extension-functions.obj
		$(LN) $(DLLLFLAGS) extension-functions.obj -out:$@ $(DLLLIBS)

csv.dll:	csv.obj
		$(LN) $(DLLLFLAGS) csv.obj -out:$@ $(DLLLIBS)

eval.dll:	eval.obj
		$(LN) $(DLLLFLAGS) eval.obj -out:$@ $(DLLLIBS)

regexp.dll:	regexp.obj
		$(LN) $(DLLLFLAGS) regexp.obj -out:$@ $(DLLLIBS)

checkfreelist.dll:	checkfreelist.obj
		$(LN) $(DLLLFLAGS) checkfreelist.obj -out:$@ $(DLLLIBS) 

shathree.dll:	shathree.obj
		$(LN) $(DLLLFLAGS) shathree.obj -out:$@ $(DLLLIBS) 

fileio.dll:	fileio.obj
		$(LN) $(DLLLFLAGS) fileio.obj -out:$@ $(DLLLIBS) 

series.dll:	series.obj
		$(LN) $(DLLLFLAGS) series.obj -out:$@ $(DLLLIBS) 

sha1.dll:	sha1.obj
		$(LN) $(DLLLFLAGS) sha1.obj -out:$@ $(DLLLIBS) 
        
sqlfcmp.dll:	sqlfcmp.obj
		$(LN) $(DLLLFLAGS) sqlfcmp.obj -out:$@ $(DLLLIBS) 

totype.dll:	totype.obj
		$(LN) $(DLLLFLAGS) totype.obj -out:$@ $(DLLLIBS) 

wholenumber.dll:	wholenumber.obj
		$(LN) $(DLLLFLAGS) wholenumber.obj -out:$@ $(DLLLIBS) 

decimal.dll:	decimal.obj
		$(LN) $(DLLLFLAGS) decimal.obj -out:$@ $(DLLLIBS) 

ieee754.dll:	ieee754.obj
		$(LN) $(DLLLFLAGS) ieee754.obj -out:$@ $(DLLLIBS) 

vfsstat.dll:	vfsstat.obj
		$(LN) $(DLLLFLAGS) vfsstat.obj -out:$@ $(DLLLIBS) 

uuid.dll:	uuid.obj
		$(LN) $(DLLLFLAGS) uuid.obj -out:$@ $(DLLLIBS) 

crypto.dll:	crypto.obj md5.obj shaone.obj shatwo.obj
		$(LN) $(DLLLFLAGS) $** -out:$@ $(DLLLIBS)

bfsvtab.dll:	bfsvtab.obj
		$(LN) $(DLLLFLAGS) bfsvtab.obj -out:$@ $(DLLLIBS) 
        
uint.dll:	uint.obj
		$(LN) $(DLLLFLAGS) uint.obj -out:$@ $(DLLLIBS) 

path.dll: sqlite-path.obj cwalk.obj
		$(LN) $(DLLLFLAGS) sqlite-path.obj cwalk.obj -out:$@ $(DLLLIBS) 

# --- STANDALONE EXECUTABLES ---------------------------------------------------
sqlite3.exe: shell.c sqlite3.c
    $(CC) $(CFLAGSEXE) -Fo$(TEMP)\ shell.c sqlite3.c $(DLLLIBS) -Fesqlite3.exe \
    -DSQLITE_THREADSAFE=0 \
    -DSQLITE_ENABLE_MATH_FUNCTIONS \
    -DSQLITE_ENABLE_EXPLAIN_COMMENTS \
    -DSQLITE_INTROSPECTION_PRAGMAS \
    -DSQLITE_ENABLE_UNKNOWN_SQL_FUNCTION \
    -DSQLITE_ENABLE_STMTVTAB \
    -DSQLITE_ENABLE_DBPAGE_VTAB \
    -DSQLITE_ENABLE_DBSTAT_VTAB \
    -DSQLITE_ENABLE_OFFSET_SQL_FUNC \
    -DSQLITE_ENABLE_JSON1 \
    -DSQLITE_ENABLE_RTREE \
    -DSQLITE_ENABLE_FTS4 \
    -DSQLITE_ENABLE_FTS5 \
    -DSQLITE_ENABLE_GEOPOLY \
    -DHAVE_SQLITETRACE=1
    
sqldiff.exe: sqldiff.c
    $(CC) $(CFLAGSEXE) -Fo$(TEMP)\ sqlite3.c sqldiff.c sqlite3_stdio.c $(DLLLIBS) -Fesqldiff.exe \
    -DSQLITE_THREADSAFE=0 \
    -DHAVE_SQLITETRACE=1

