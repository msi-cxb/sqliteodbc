# VC++ 2015/2017/2019/2022 Makefile
# uses the SQLite3 amalgamation source which must
# be unpacked below in the same folder as this makefile

CC=		cl
LN=		link
RC=		rc

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

# exclude warnings that clutter the screen
EXCLUDETHESEWARNINGS=/wd4477

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
        $(EXCLUDETHESEWARNINGS) \
        -DSQLITE_THREADSAFE=0 \
        -DSQLITE_OS_WIN=1 \
        -DHAVE_SQLITE3COLUMNTABLENAME=1 \
        -DHAVE_SQLITE3PREPAREV2=1 \
        -DHAVE_SQLITE3VFS=1 \
        -DHAVE_SQLITE3LOADEXTENSION=1 \
        -DSQLITE_ENABLE_COLUMN_METADATA=1 \
        -DSQLITE_SOUNDEX=1 \
        -DSQLITE_TEMP_STORE=1 \
        -DFILEIO_WIN32_DLL \
        -DSQLITE_WIN32_MALLOC \
        -DSQLITE_ODBC_WIN32_MALLOC \
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
        $(EXCLUDETHESEWARNINGS) \
        $(DETERMINISTIC)

DLLLFLAGS=$(LDEBUG) \
        $(DETERMINISTIC) \
        /NOLOGO \
        /SUBSYSTEM:WINDOWS \
        /DLL \
        /NODEFAULTLIB \
        # /MACHINE:IX86

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

all:    driver \
        exe \
        extensions
        
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
        extension-functions.dll \
        fileio.dll \
        ieee754.dll \
        regexp.dll \
        shathree.dll \
        series.dll \
        sha1.dll \
        sqlfcmp.dll \
        totype.dll \
        uuid.dll \
        wholenumber.dll \
        vfsstat.dll

# needs to be run as administrator
install: clean SQLiteODBCInstaller.exe sqlite3odbc.dll
    SQLiteODBCInstaller.exe -i -d=sql3 -q

# needs to be run as administrator
# use quickinstall to rebuild without a clean
quickinstall: SQLiteODBCInstaller.exe sqlite3odbc.dll
    SQLiteODBCInstaller.exe -i -d=sql3 -q

# needs to be run as administrator
uninstall: clean SQLiteODBCInstaller.exe
    SQLiteODBCInstaller.exe -u -a -q

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

sqlite3odbc.c:	resource3.h sqlite3odbc.h sqltypes.h

sqlite3odbc.res:	sqlite3odbc.rc resource3.h
		$(RC) -I. -fo sqlite3odbc.res -r sqlite3odbc.rc

OBJECTS=	sqlite3odbc.obj sqlite3.obj

sqlite3odbc.dll:	$(OBJECTS) sqlite3odbc.res
		$(LN) $(DLLLFLAGS) $(OBJECTS) sqlite3odbc.res \
		-def:sqlite3odbc.def -out:$@ $(DLLLIBS)

VERSION_C:	fixup.exe VERSION
		.\fixup < VERSION > VERSION_C . ,

resource3.h:	resource.h.in VERSION_C fixup.exe
        .\fixup < resource.h.in > resource3.h \
        --VERS-- @VERSION \
        --VERS_C-- @VERSION_C

extension-functions.dll:	extension-functions.obj
		$(LN) $(DLLLFLAGS) extension-functions.obj -out:$@ $(DLLLIBS)

csv.dll:	csv.obj
		$(LN) $(DLLLFLAGS) csv.obj -out:$@ $(DLLLIBS)

regexp.dll:	regexp.obj
		$(LN) $(DLLLFLAGS) regexp.obj -out:$@ $(DLLLIBS)

checkfreelist.dll:	checkfreelist.obj
		$(LN) $(DLLLFLAGS) checkfreelist.obj -out:$@ $(DLLLIBS) 

shathree.dll:	shathree.obj
		$(LN) $(DLLLFLAGS) shathree.obj -out:$@ $(DLLLIBS) 

fileio.dll:	fileio.obj sqlite3.obj
		$(LN) $(DLLLFLAGS) -DFILEIO_WIN32_DLL fileio.obj sqlite3.obj -out:$@ $(DLLLIBS) 

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
        
sqlite3.exe: shell.c sqlite3.c
    $(CC) $(CFLAGSEXE) shell.c sqlite3.c $(DLLLIBS) -Fesqlite3.exe \
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
    $(CC) $(CFLAGSEXE) sqlite3.c sqldiff.c sqlite3_stdio.c $(DLLLIBS) -Fesqldiff.exe \
    -DSQLITE_THREADSAFE=0 \
    -DHAVE_SQLITETRACE=1

