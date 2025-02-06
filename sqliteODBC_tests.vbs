option explicit

class classTimer
    Private fStartTime
    Private fStopTime
    Private fCurrentTime
    Private lCounter
    
    '**********************************************************************
    Private Sub Class_Initialize
        lCounter = 0
    end sub

    '**********************************************************************
    Private Sub Class_Terminate
    end sub

    '**********************************************************************
    Public Property Let StartTime(f) 
        fStartTime = f 
    End Property
    
    Public Property Get StartTime 
        StartTime = fStartTime 
    End Property 
    
    '**********************************************************************
    Public Property Let StopTime(f) 
        fStopTime = f 
    End Property
    
    Public Property Get StopTime 
        StopTime = fStopTime 
    End Property 
    
    '**********************************************************************
    Public Property Get CurrentTime
        fCurrentTime = Timer
        CurrentTime = fCurrentTime 
    End Property 
    
    '**********************************************************************
    Public Property Get ElapsedTime
        fCurrentTime = Timer
        ElapsedTime = fCurrentTime - fStartTime
    End Property 
    
    '**********************************************************************
    Public Property Get Counter 
        Counter = lCounter 
    End Property 
    
    '**********************************************************************
    public function StartTimer()
        StartTime = Timer
        StartTimer = true
    end function
    
    '**********************************************************************
    public function StopTimer()
        StopTime = Timer
        StopTimer = true
    end function
    
    '**********************************************************************
    public function IncrementCounter()
        lCounter = lCounter + 1
        IncrementCounter = lCounter
    end function
        
    '**********************************************************************
    public function ResultString()
    
        dim sString
        sString = "Timer Results: Total time: " & _
            Fix((fStopTime - fStartTime)/60) & " min "  & _
            ((round(fStopTime - fStartTime)) mod 60) & " sec"
            
        if lCounter > 0 then
            sString = sString & " Average time per run: " & _
                ((round(fStopTime - fStartTime))/lCounter) & " sec"
        end if
    
        ResultString = sString
        
    end function
    
    '**********************************************************************
    private Function Pad(strText, nLen, strChar, bFront) 
        Dim nStartLen 
        If strChar = "" Then 
            strChar = "0" 
        End If 
        nStartLen = Len(strText) 
        If Len(strText) >= nLen Then 
            Pad = strText 
        Else 
            If bFront Then 
                Pad = String(nLen - Len(strText), strChar) & strText 
            Else 
                Pad = strText & String(nLen - Len(strText), strChar) 
            End If 
        End If 
    End Function
    
end class

' so that javascript can create new object instance of VBScript class
public function NewclassSqliteOdbcTests()
    Set NewclassSqliteOdbcTests = new classSqliteOdbcTests
End Function

Const adUseClient = 3
Const adOpenStatic = 3
Const adOpenDynamic = 2

Const adLockBatchOptimistic = 4
Const adLockOptimistic = 3
Const adLockPessimistic = 2
Const adLockReadOnly = 1

Const ForReading = 1 
Const ForWriting = 2 
Const ForAppending = 8


' CommandTypeEnum
Const adCmdUnspecified = -1	        ' Does not specify the command type argument.
Const adCmdText = 1	                ' Evaluates CommandText as a textual definition of a command or stored procedure call.
Const adCmdTable = 2                ' Evaluates CommandText as a table name whose columns are 
                                    ' all returned by an internally generated SQL query.
Const adCmdStoredProc = 4           ' Evaluates CommandText as a stored procedure name.
Const adCmdUnknown  = 8             ' Indicates that the type of command in the CommandText property is not known.
Const adCmdFile = 256               ' Evaluates CommandText as the file name of a persistently stored Recordset. 
                                    ' Used with Recordset.Open or Requery only.
Const adCmdTableDirect = 512        ' Evaluates CommandText as a table name whose columns are all returned. 
                                    ' Used with Recordset.Open or Requery only. To use the Seek method, 
                                    ' the Recordset must be opened with adCmdTableDirect. This value cannot 
                                    ' be combined with the ExecuteOptionEnum value adAsyncExecute.

' ExecuteOptionEnum
Const adAsyncExecute = 16	        ' Indicates that the command should execute asynchronously.
                                    ' This value cannot be combined with the CommandTypeEnum value 
                                    ' adCmdTableDirect.
Const adAsyncFetch = 32	            ' Indicates that the remaining rows after the initial quantity 
                                    ' specified in the CacheSize property should be retrieved asynchronously.
Const adAsyncFetchNonBlocking = 64  ' Indicates that the main thread never blocks while retrieving. 
                                    ' If the requested row has not been retrieved, the current row 
                                    ' automatically moves to the end of the file.
                                    ' If you open a Recordset from a Stream containing a persistently 
                                    ' stored Recordset, adAsyncFetchNonBlocking will not have an effect; 
                                    ' the operation will be synchronous and blocking.
                                    ' adAsynchFetchNonBlocking has no effect when the adCmdTableDirect option 
                                    ' is used to open the Recordset.
Const adExecuteNoRecords = 128	    ' Indicates that the command text is a command or stored procedure 
                                    ' that does not return rows (for example, a command that only inserts data). 
                                    ' If any rows are retrieved, they are discarded and not returned.
                                    ' adExecuteNoRecords can only be passed as an optional parameter 
                                    ' to the Command or Connection Execute method.
Const adExecuteStream = 1024        ' Indicates that the results of a command execution should be returned 
                                    ' as a stream.adExecuteStream can only be passed as an optional parameter 
                                    ' to the Command Execute method.
Const adExecuteRecord = 2048        ' Indicates that the CommandText is a command or stored procedure that 
                                    ' returns a single row which should be returned as a Record object.
Const adOptionUnspecified = -1	    ' Indicates that the command is unspecified.

class classSqliteOdbcTests
    private objConn 
    private objRS 
    private strQuery 
    private iBitness
    private sBitPath
    private objFSO
    private strFolder
    private oWShell
    private separator
    private aQueryResults
    private dDataTypes
    
    private bOpenFirstTime
    private bHasAccess
    private sAppdataPath
    private bVerboseOutput
    private dbSqlite3
    private oTimer
    
    '*************************************************************************
    sub class_initialize()
        set oTimer = new classTimer

        bOpenFirstTime = true
        
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set oWShell = CreateObject("WScript.Shell")
        
        separator = ","
        
        aQueryResults = Array()
        redim aQueryResults(4)
        aQueryResults(0) = ""
        aQueryResults(1) = ""
        set aQueryResults(2) = CreateObject("Scripting.Dictionary")
        aQueryResults(3) = ""
        
        on error resume next
        'try wscript first...if it works then move on
        strFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
        'if variable not defined (error = 500), then we are running in an hta
        if err.number = 500 then
            strFolder = objFSO.GetParentFolderName(Replace(location.pathname,"%20"," "))
        end if
        on error goto 0
        
        log "strFolder [" & strFolder & "]"
        
        Set dDataTypes = CreateObject("Scripting.Dictionary")
        dDataTypes.add 20,"adBigInt" 'Indicates an eight-byte signed integer (DBTYPE_I8).
        dDataTypes.add 128,"adBinary" 'Indicates a binary value (DBTYPE_BYTES).
        dDataTypes.add 11,"adBoolean" 'Indicates a Boolean value (DBTYPE_BOOL).
        dDataTypes.add 8,"adBSTR" 'Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR).
        dDataTypes.add 136,"adChapter" 'Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER).
        dDataTypes.add 129,"adChar" 'Indicates a string value (DBTYPE_STR).
        dDataTypes.add 6,"adCurrency" 'Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000.
        dDataTypes.add 7,"adDate" 'Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
        dDataTypes.add 133,"adDBDate" 'Indicates a date value (yyyymmdd) (DBTYPE_DBDATE).
        dDataTypes.add 134,"adDBTime" 'Indicates a time value (hhmmss) (DBTYPE_DBTIME).
        dDataTypes.add 135,"adDBTimeStamp" 'Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
        dDataTypes.add 14,"adDecimal" 'Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
        dDataTypes.add 5,"adDouble" 'Indicates a double-precision floating-point value (DBTYPE_R8).
        dDataTypes.add 0,"adEmpty" 'Specifies no value (DBTYPE_EMPTY).
        dDataTypes.add 10,"adError" 'Indicates a 32-bit error code (DBTYPE_ERROR).
        dDataTypes.add 64,"adFileTime" 'Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME).
        dDataTypes.add 72,"adGUID" 'Indicates a globally unique identifier (GUID) (DBTYPE_GUID).
        dDataTypes.add 3,"adInteger" 'Indicates a four-byte signed integer (DBTYPE_I4).
        dDataTypes.add 205,"adLongVarBinary" 'Indicates a long binary value.
        dDataTypes.add 201,"adLongVarChar" 'Indicates a long string value.
        dDataTypes.add 203,"adLongVarWChar" 'Indicates a long null-terminated Unicode string value.
        dDataTypes.add 131,"adNumeric" 'Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
        dDataTypes.add 138,"adPropVariant" 'Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT).
        dDataTypes.add 4,"adSingle" 'Indicates a single-precision floating-point value (DBTYPE_R4).
        dDataTypes.add 2,"adSmallInt" 'Indicates a two-byte signed integer (DBTYPE_I2).
        dDataTypes.add 16,"adTinyInt" 'Indicates a one-byte signed integer (DBTYPE_I1).
        dDataTypes.add 21,"adUnsignedBigInt" 'Indicates an eight-byte unsigned integer (DBTYPE_UI8).
        dDataTypes.add 19,"adUnsignedInt" 'Indicates a four-byte unsigned integer (DBTYPE_UI4).
        dDataTypes.add 18,"adUnsignedSmallInt" 'Indicates a two-byte unsigned integer (DBTYPE_UI2).
        dDataTypes.add 17,"adUnsignedTinyInt" 'Indicates a one-byte unsigned integer (DBTYPE_UI1).
        dDataTypes.add 132,"adUserDefined" 'Indicates a user-defined variable (DBTYPE_UDT).
        dDataTypes.add 204,"adVarBinary" 'Indicates a binary value.
        dDataTypes.add 200,"adVarChar" 'Indicates a string value.
        dDataTypes.add 12,"adVariant" 'Indicates an Automation Variant (DBTYPE_VARIANT).
        dDataTypes.add 139,"adVarNumeric" 'Indicates a numeric value.
        dDataTypes.add 202,"adVarWChar" 'Indicates a null-terminated Unicode character string.
        dDataTypes.add 130,"adWChar" 'Indicates a null-terminated Unicode character string (DBTYPE_WSTR).
        
        log "class_initialize"
    end sub

    '*************************************************************************
    sub class_terminate()
        set objFSO = nothing
        set oWShell = nothing
        set oTimer = nothing
        log "class_terminate"
    end sub
    
    '********************************************
    public function executeTests(i)
        select case i
            case 32
                iBitness = 32
                sBitPath = "32bit"
            case 64
                iBitness = 64
                sBitPath = "64bit"
            case else
                log "script requires you provide 32 or 64"
                exit function
        end select
        
        log "classSqliteOdbcTests running " & sBitPath
        
        main
        
    end function

    '********************************************
    public function main
        on error resume next

        log "****************************************************************************"
        log "main"
        log ""

        bVerboseOutput = false
        dim runTests: runTests = true
        
        if objFSO.FolderExists(strFolder & "\testDBs") = false then
            objFSO.CreateFolder(strFolder & "\testDBs")
        end if

        log "dbSqlite3 [" & dbSqlite3 & "]"

        sqlite_version
        If Err.Number <> 0 Then wscript.quit -1
        
        ' sqlite features
        if runTests then
            longSqlStringReturn
            If Err.Number <> 0 Then wscript.quit -1

            log dumpPragma
            If Err.Number <> 0 Then wscript.quit -1

            dbSqlite3 = strFolder & "\testDBs\testfile.sqlite3"
            sqlite3_BuiltIn_Tests
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_basic_cluster
            If Err.Number <> 0 Then wscript.quit -1

            sqlite3_rtree_tests
            If Err.Number <> 0 Then wscript.quit -1

            sqlite3_fts_tests
            If Err.Number <> 0 Then wscript.quit -1

            sqlite3_page_size_tests
            If Err.Number <> 0 Then wscript.quit -1

            sqlite3_feature_tests
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_big_numbers
            If Err.Number <> 0 Then wscript.quit -1

            recursiveCTE
            If Err.Number <> 0 Then wscript.quit -1

            generate_series
            If Err.Number <> 0 Then wscript.quit -1

            calendarExamples
            If Err.Number <> 0 Then wscript.quit -1

            graphExampleOne
            If Err.Number <> 0 Then wscript.quit -1

            graphExampleTwo
            If Err.Number <> 0 Then wscript.quit -1
            
            isValidIntOrFloat
            If Err.Number <> 0 Then wscript.quit -1

            varCharStringIssue
            If Err.Number <> 0 Then wscript.quit -1

            unicodeCharacterTest
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_geopoly
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_right_join
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_full_outer_join
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_isDistinctFrom
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_havingWithoutGroupBy
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_localtimeModifierMaintainsFractSecs
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_unixepochFunction
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_autoModifier
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_juliandayModifier
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_strictTable
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_double_quoted_strings
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_strftime
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_json
            If Err.Number <> 0 Then wscript.quit -1

            getTableInfoSimple
            If Err.Number <> 0 Then wscript.quit -1

            getTableInfoDetails
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_json_virtual_columns
            If Err.Number <> 0 Then wscript.quit -1
            
            sqlite_simple_virtual_columns
            If Err.Number <> 0 Then wscript.quit -1
        end if

        ' extension tests
        if runTests then
            dbSqlite3 = strFolder & "\testDBs\csv.sqlite3"
            sqlite_extension_functions_csv
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_functions_tests
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_functions_checkfreelist
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_functions_ieee754
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_functions_regex
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_functions_series
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_functions_sha
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_functions_totype
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_functions_wholenumber
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_uuid
            If Err.Number <> 0 Then wscript.quit -1

            dbSqlite3 = strFolder & "\testDBs\testfile.sqlite3"
            sqlite_extension_bfsvtab
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_decimal
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_extension_fileio
            If Err.Number <> 0 Then wscript.quit -1

            sqlite_sqlfcmp
            If Err.Number <> 0 Then wscript.quit -1
        
            dbSqlite3 = strFolder & "\testDBs\crypto.sqlite3"
            sqlite_extension_crypto
            If Err.Number <> 0 Then wscript.quit -1

            ' vfsstat doesn't work in ODBC (can't load extension before opening db)
            ' sqlite_extension_vfsstat
            ' If Err.Number <> 0 Then wscript.quit -1

        end if
        
        ' insert tests
        if runTests then
            insertTests
            If Err.Number <> 0 Then wscript.quit -1
        end if
        
        ' inventory the testDb folder then clean up
        if runTests then
            testDbInventory
            If Err.Number <> 0 Then wscript.quit -1
            if objFSO.FolderExists(strFolder & "\testDBs") then objFSO.DeleteFolder(strFolder & "\testDBs")
        end if

        ' can be used to create some test tables
        ' r = number of rows to insert
        ' c = number of columns per row (max is 998)
        ' t = column data type (INTEGER, REAL, TEXT)
        ' p = which driver to use (SQL3 or 
        ' ipt = number of records per transaction
        ' pk = use primary key (true/false)
        ' pg = pipe "|" separated pragma string
        ' function test(r,c,t,p,ipt,pk,pg)
        ' test 100,10,"INTEGER","SQL3 ",100,true,"journal=OFF",true,false
        ' test 10,1000,"TEXT","SQL3 ",100,true,"journal=OFF",false,false

        log "FINISHED!"
        on error goto 0
    end function

    '********************************************
    public function sqlite_version
        dim oShell: Set oShell = WScript.CreateObject("WScript.Shell")
        dim proc_arch: proc_arch = oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
        set oShell = nothing
        log "****************************************************************************"
        log "sqlite_version arch = " & proc_arch
        opendb "MEM  "
        log query("SELECT sqlite_version() as vers, sqlite_source_id() as srcId;")
        closedb
    end function

    '********************************************
    public function sqlite_basic_cluster
        ' https://sqlite.org/forum/forumpost/3be3abdbff
        ' I would like to automatically cluster dates together. 
        ' If the dates are further than 5 days apart, I'd like them to form a new group.
        dim q
        dim retValue: retValue = 0
        dim result
        opendb "MEM  "
        
        result = query2csv("create temp table some_data(action DATE);")
        if result = -1 then retValue = retValue + 1

        ' First set
        result = query2csv("insert into temp.some_data VALUES ('2000-01-01');")
        if result = -1 then retValue = retValue + 1
        result = query2csv("insert into temp.some_data VALUES ('2000-01-01');")
        if result = -1 then retValue = retValue + 1
        result = query2csv("insert into temp.some_data VALUES ('2000-01-02');")
        if result = -1 then retValue = retValue + 1
        result = query2csv("insert into temp.some_data VALUES ('2000-01-04');")
        if result = -1 then retValue = retValue + 1

        ' second set
        result = query2csv("insert into temp.some_data VALUES ('2000-01-25');")
        if result = -1 then retValue = retValue + 1
        
        ' third set - note this crosses months
        result = query2csv("insert into temp.some_data VALUES ('2000-01-31');")
        if result = -1 then retValue = retValue + 1
        result = query2csv("insert into temp.some_data VALUES ('2000-02-01');")
        if result = -1 then retValue = retValue + 1
        result = query2csv("insert into temp.some_data VALUES ('2000-02-01');")
        if result = -1 then retValue = retValue + 1
        result = query2csv("insert into temp.some_data VALUES ('2000-02-01');")
        if result = -1 then retValue = retValue + 1
        
        ' fourth set
        result = query2csv("insert into temp.some_data VALUES ('2000-02-07');")
        if result = -1 then retValue = retValue + 1

        ' fifth set
        result = query2csv("insert into temp.some_data VALUES ('2000-03-01');")
        if result = -1 then retValue = retValue + 1
        result = query2csv("insert into temp.some_data VALUES ('2000-03-02');")
        if result = -1 then retValue = retValue + 1
        result = query2csv("insert into temp.some_data VALUES ('2000-03-02');")
        if result = -1 then retValue = retValue + 1
        result = query2csv("insert into temp.some_data VALUES ('2000-03-02');")
        if result = -1 then retValue = retValue + 1
        
        q = _
            "WITH " &_
            "    edge_detect as ( " &_
            "        SELECT action, " &_
            "            julianday(action)-julianday(lag(action) OVER ()) <= 5 AS ingrp " &_
            "        FROM some_data " &_
            "        ORDER BY action " &_
            "    ), " &_
            "    build_grouping AS ( " &_
            "        SELECT action, COUNT(*) FILTER (WHERE ingrp IS NOT 1) " &_
            "            OVER (ROWS UNBOUNDED PRECEDING) AS g " &_
            "        FROM edge_detect " &_
            "    ) " &_
            "SELECT MIN(action) as Start, MAX(action) AS End " &_
            "FROM build_grouping " &_
            "GROUP BY g; "
        result = query2csv(q)
        if result <> 5 then retValue = retValue + 1
        if aQueryResults(1) <> "Start(adVarWChar),End(adVarWChar)" then retValue = retValue + 1
        if aQueryResults(2)(0) <> """2000-01-01"",""2000-01-04""" then retValue = retValue + 1
        if aQueryResults(2)(1) <> """2000-01-25"",""2000-01-25""" then retValue = retValue + 1
        if aQueryResults(2)(2) <> """2000-01-31"",""2000-02-01""" then retValue = retValue + 1
        if aQueryResults(2)(3) <> """2000-02-07"",""2000-02-07""" then retValue = retValue + 1
        if aQueryResults(2)(4) <> """2000-03-01"",""2000-03-02""" then retValue = retValue + 1
        
        closedb
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_sqlfcmp
        dim s: s = ""
        dim retValue: retValue = 0

        'https://sqlite.org/forum/forumpost/9040b9c0b1532b5c
        ' note the numbers here don't match website numbers exactly, but
        ' DO match output from locally compiled SQLite3

        log "******************************************************"
        log "sqlite_sqlfcmp (Keith Medcalf extension)"
        
        opendb "MEM  "
        log "load_extension('.\install\" & sBitPath & "\sqlfcmp.dll') will throw error but it does load the extension"
        query2csv("SELECT load_extension('.\install\" & sBitPath & "\sqlfcmp.dll') as ext_loaded")
        if instr(aQueryResults(3),"Function sequence error") = 0 then retValue = retValue+1

        query2csv("drop table if exists t;")
        query2csv("create table t(c real);")
        query2csv("insert into t values(.00001);")
        query2csv("insert into t values(.00001);")
        query2csv("insert into t values(.00001);")

        query2csv("select format('%.5f', total(c)) from t;")
        if aQueryResults(2)(0) <> """0.00003""" then retValue = retValue+1

        query2csv("select format('%!.26f', c) from t;")
        if aQueryResults(2).count <> 3 then retValue = retValue+1

        query2csv("select format('%!.26f', total(c)) from t;")
        if aQueryResults(2)(0) <> """0.00003000000000000000415""" then retValue = retValue+1

        ' https://sqlite.org/forum/forumpost/f46ded7529d456d7
        ' It is a scalar function that takes two arguments -- the first being the number to operate on, 
        ' and the second optional argument being the number of significant digits to maintain, with the default being 14.
        ' the "exactly rounded" result 
        query2csv("select format('%!.26f', sigdigits(total(c))) from t;")
        if aQueryResults(2)(0) <> """0.00003000000000000000076""" then retValue = retValue+1
        
        ' a diference of 1 ULP.
        query2csv("select ulps(sigdigits(total(c)), total(c)) from t;")
        if cint(aQueryResults(2)(0)) <> -1 then retValue = retValue+1

        ' the values differ by 1 ULP so they are not equal (0)
        query2csv("select total(c) == 3.0e-05 from t;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1

        query2csv("select format('%!.26f', 100.00000000000000001) as value union all select format('%!.26f', 100.00000000000001);")
        if aQueryResults(2)(0) <> """100.0""" then retValue = retValue+1
        if aQueryResults(2)(1) <> """100.0000000000000142""" then retValue = retValue+1

        query2csv("select ulps(100.00000000000000001,100.00000000000001) as u;")
        if cint(aQueryResults(2)(0)) <> -1 then retValue = retValue+1
        
        closedb

        opendb "SQL3-sqlfcmp"

        query2csv("drop table if exists t;")
        query2csv("create table t(c real);")
        query2csv("insert into t values(.00001);")
        query2csv("insert into t values(.00001);")
        query2csv("insert into t values(.00001);")

        query2csv("select format('%.5f', total(c)) from t;")
        if aQueryResults(2)(0) <> """0.00003""" then retValue = retValue+1

        query2csv("select format('%!.26f', c) from t;")
        if aQueryResults(2).count <> 3 then retValue = retValue+1

        query2csv("select format('%!.26f', total(c)) from t;")
        if aQueryResults(2)(0) <> """0.00003000000000000000415""" then retValue = retValue+1

        ' https://sqlite.org/forum/forumpost/f46ded7529d456d7
        ' It is a scalar function that takes two arguments -- the first being the number to operate on, 
        ' and the second optional argument being the number of significant digits to maintain, with the default being 14.
        ' the "exactly rounded" result 
        query2csv("select format('%!.26f', sigdigits(total(c))) from t;")
        if aQueryResults(2)(0) <> """0.00003000000000000000076""" then retValue = retValue+1
        
        ' a diference of 1 ULP.
        query2csv("select ulps(sigdigits(total(c)), total(c)) from t;")
        if cint(aQueryResults(2)(0)) <> -1 then retValue = retValue+1

        ' the values differ by 1 ULP so they are not equal (0)
        query2csv("select total(c) == 3.0e-05 from t;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1

        query2csv("select format('%!.26f', 100.00000000000000001) as value union all select format('%!.26f', 100.00000000000001);")
        if aQueryResults(2)(0) <> """100.0""" then retValue = retValue+1
        if aQueryResults(2)(1) <> """100.0000000000000142""" then retValue = retValue+1

        query2csv("select ulps(100.00000000000000001,100.00000000000001) as u;")
        if cint(aQueryResults(2)(0)) <> -1 then retValue = retValue+1

        closedb

        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_extension_fileio
        log "****************************************************************************"
        log "sqlite_extension_fileio"
        
        REM C:\Program Files (x86)\Windows Kits\10\Include\10.0.10240.0\ucrt\sys\stat.h
        REM #define _S_IFMT   0xF000 // File type mask
        REM #define _S_IFDIR  0x4000 // Directory
        REM #define _S_IFCHR  0x2000 // Character special
        REM #define _S_IFIFO  0x1000 // Pipe
        REM #define _S_IFREG  0x8000 // Regular
        REM #define _S_IREAD  0x0100 // Read permission, owner
        REM #define _S_IWRITE 0x0080 // Write permission, owner
        REM #define _S_IEXEC  0x0040 // Execute/search permission, owner

        logResult query2csv("SELECT load_extension('.\install\" & sBitPath & "\fileio.dll') as ext_loaded")

        log "fileio dll does not work with SELECT load_extension(), so need to load via connection string"
        opendb "SQL3-fileio"
        
        REM dump file info to table so we can look at schema
        log query("drop table if exists [fileInfo_C_Temp];") 
        log query("create table [fileInfo_C_Temp] as select * FROM fsdir('c:\temp');") 
        log query("PRAGMA table_info('fileInfo_C_Temp');")
        log query("select count(1) as numRecords from [fileInfo_C_Temp];")
        
        ' log "script crashes when blob is binary and large, and error handling does not catch issue"
        ' log "so you can't so a select * or select with data field if folder contains large binary file."
        ' log query("select name,mode,mtime,datetime(mtime, 'unixepoch','localtime') as reporttime from [fileInfo_C_Temp];")
        
        ' log "if $dir is provided, and $path is relative then $path interpreted relative to $dir"
        ' log "if folder does not exist then this will fail..."
        log query( _
            "SELECT name, " & _
            "printf('%08X',mode) as m, " & _
            "( printf('%04X',(mode & 0xF000)) == '8000' ) as isFile, " & _
            "( printf('%04X',(mode & 0xF000)) == '4000' ) as isDir, " & _
            "( printf('%04X',(mode & 0x0100)) == '0100' ) as isRead, " & _
            "( printf('%04X',(mode & 0x0080)) == '0080' ) as isWrite, " & _
            "( printf('%04X',(mode & 0x0040)) == '0040' ) as isExe, " & _
            "'' as theEnd " & _
            "FROM fsdir('..\','c:\temp\folder');" & _
            "")
        ' log "same but with explicit path only passed to fsdir()"
        log query( _
            "SELECT name, " & _
            "printf('%08X',mode) as m, " & _
            "( printf('%04X',(mode & 0xF000)) == '8000' ) as isFile, " & _
            "( printf('%04X',(mode & 0xF000)) == '4000' ) as isDir, " & _
            "( printf('%04X',(mode & 0x0100)) == '0100' ) as isRead, " & _
            "( printf('%04X',(mode & 0x0080)) == '0080' ) as isWrite, " & _
            "( printf('%04X',(mode & 0x0040)) == '0040' ) as isExe, " & _
            "'' as theEnd " & _
            "FROM fsdir('c:\temp');" & _
            "")
            
        ' log "read content of a text file where we know we can handle the blob in vbscript data types"
        log query( "select cast( readfile('test.csv') as text) as fileContent;" )
        
        ' log "db table column 'data' contains binary contennts of MSI.gif, write that out to a new file"
        ' log "if successful, writefile returns the number of bytes written. if not, empty recordset so nothing returned."
        ' log query( "SELECT writefile('c:\temp\MSI_new.gif',data) as bytesWritten FROM [fileInfo_C_Temp] WHERE name='c:\temp/MSI.gif';")
        
        closedb
    end function

    '********************************************
    public function sqlite_extension_decimal
        dim s: s = ""
        dim retValue: retValue = 0

        ' extension works but type is ignored as shoown below

        log "******************************************************"
        log "sqlite_extension_decimal"
        opendb "MEM  "
        log "load_extension() throws error but works"
        query2csv("SELECT load_extension('.\install\" & sBitPath & "\decimal.dll') as ext_loaded")
        if instr(aQueryResults(3),"Function sequence error") = 0 then retValue = retValue+1

        query2csv("SELECT decimal(pi());")
        query2csv("SELECT decimal_exp(pi())")
        query2csv("SELECT decimal(X'4055480000000000')")
        query2csv("SELECT decimal_mul(2, 3)")

        query2csv("CREATE TABLE MyTable(X real,Y integer,Z text);")
        query2csv("INSERT INTO MyTable VALUES (1.1,2,'100'),(2.2,4,'10'),(3.3,6,'1');")
        query2csv("SELECT decimal_sum(X) as dc FROM MyTable;")
        query2csv("SELECT decimal_add(X,Y) as d_a FROM MyTable;")
        query2csv("SELECT decimal_sub(X,Y) as d_s FROM MyTable;")
        query2csv("SELECT decimal_mul(X,Y) as d_m FROM MyTable;")
        query2csv("SELECT decimal_pow2(Y) as d_p FROM MyTable;")

        ' text converted to high precision decimal so...
        ' returns 1 --> first value is greater than second value
        query2csv("select decimal_cmp('100.00000000000000001', '100.00000000000000000') as d_c;")
        if cint(aQueryResults(2)(0)) <> 1 then retValue = retValue+1
        ' returns 0 --> first value is equal to second value
        query2csv("select decimal_cmp('100.00000000000000000', '100.00000000000000000') as d_c;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1
        ' returns -1 --> first value is less than second value
        query2csv("select decimal_cmp('100.00000000000000000', '100.00000000000000001') as d_c;")
        if cint(aQueryResults(2)(0)) <> -1 then retValue = retValue+1

        ' these numbers cannot be represented by double so with rounding these all return 0 (equal)
        query2csv("select decimal_cmp(100.00000000000000001, 100.00000000000000000) as d_c;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1
        query2csv("select decimal_cmp(100.00000000000000000, 100.00000000000000000) as d_c;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1
        query2csv("select decimal_cmp(100.00000000000000000, 100.00000000000000001) as d_c;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1

        closedb

        log "connection string"
        opendb "SQL3-decimal"

        query2csv("SELECT decimal(pi());")
        query2csv("SELECT decimal_exp(pi())")
        query2csv("SELECT decimal(X'4055480000000000')")
        query2csv("SELECT decimal_mul(2, 3)")

        query2csv("CREATE TABLE MyTable(X real,Y integer,Z text);")
        query2csv("INSERT INTO MyTable VALUES (1.1,2,'100'),(2.2,4,'10'),(3.3,6,'1');")
        query2csv("SELECT decimal_sum(X) as dc FROM MyTable;")
        query2csv("SELECT decimal_add(X,Y) as d_a FROM MyTable;")
        query2csv("SELECT decimal_sub(X,Y) as d_s FROM MyTable;")
        query2csv("SELECT decimal_mul(X,Y) as d_m FROM MyTable;")
        query2csv("SELECT decimal_pow2(Y) as d_p FROM MyTable;")

        ' text converted to high precision decimal so...
        ' returns 1 --> first value is greater than second value
        query2csv("select decimal_cmp('100.00000000000000001', '100.00000000000000000') as d_c;")
        if cint(aQueryResults(2)(0)) <> 1 then retValue = retValue+1
        ' returns 0 --> first value is equal to second value
        query2csv("select decimal_cmp('100.00000000000000000', '100.00000000000000000') as d_c;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1
        ' returns -1 --> first value is less than second value
        query2csv("select decimal_cmp('100.00000000000000000', '100.00000000000000001') as d_c;")
        if cint(aQueryResults(2)(0)) <> -1 then retValue = retValue+1

        ' these numbers cannot be represented by double so with rounding these all return 0 (equal)
        query2csv("select decimal_cmp(100.00000000000000001, 100.00000000000000000) as d_c;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1
        query2csv("select decimal_cmp(100.00000000000000000, 100.00000000000000000) as d_c;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1
        query2csv("select decimal_cmp(100.00000000000000000, 100.00000000000000001) as d_c;")
        if cint(aQueryResults(2)(0)) <> 0 then retValue = retValue+1

        closedb

        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_extension_vfsstat 
        dim s: s = ""
        dim retValue: retValue = 0

        log "****************************************************************************"
        log "sqlite_extension_vfsstat"
        
        REM ******************************
        REM NOTE
        REM ******************************
        REM this works in sqlite3.exe so the extension is ok, just doesn't work via ODBC
        REM could the problem be that you need to load the extension **BEFORE** opening DB?
        REM .load ./vfsstat
        REM .open test.db
        REM DROP TABLE IF EXISTS t1;
        REM CREATE TABLE t1(x,y);
        REM INSERT INTO t1 VALUES(123, randomblob(5000));
        REM CREATE INDEX t1x ON t1(x);
        REM DROP TABLE t1;
        REM VACUUM;
        REM SELECT * FROM vfsstat WHERE count>0;
        REM ┌──────────┬──────────────┬───────┐
        REM │   file   │     stat     │ count │
        REM ├──────────┼──────────────┼───────┤
        REM │ database │ bytes-in     │ 96    │
        REM │ database │ bytes-out    │ 40960 │
        REM │ database │ read         │ 9     │
        REM │ database │ write        │ 10    │
        REM │ database │ sync         │ 5     │
        REM │ database │ open         │ 1     │
        REM │ database │ lock         │ 32    │
        REM │ journal  │ bytes-out    │ 31348 │
        REM │ journal  │ read         │ 5     │
        REM │ journal  │ write        │ 31    │
        REM │ journal  │ sync         │ 10    │
        REM │ journal  │ open         │ 5     │
        REM │ *        │ access       │ 18    │
        REM │ *        │ delete       │ 5     │
        REM │ *        │ fullpathname │ 1     │
        REM │ *        │ randomness   │ 1     │
        REM └──────────┴──────────────┴───────┘

        dbSqlite3 = strFolder & "\testDBs\testfile.sqlite3"
        opendb "SQL3 "
        logResult query2csv("SELECT load_extension('.\install\" & sBitPath & "\vfsstat.dll') as ext_loaded")
        logResult query2csv("DROP TABLE IF EXISTS t1;")
        logResult query2csv("CREATE TABLE t1(x integer,y blob);")
        logResult query2csv("INSERT INTO t1 VALUES(123, randomblob(5000));")
        logResult query2csv("CREATE INDEX t1x ON t1(x);")
        logResult query2csv("SELECT name FROM sqlite_master;")
        logResult query2csv("select * from t1;")
        logResult query2csv("DROP TABLE IF EXISTS t1;")
        logResult query2csv("SELECT name FROM sqlite_master;")
        logResult query2csv("VACUUM;")
        logResult query2csv("SELECT * FROM vfsstat;")
        closedb
        
    end function

    '********************************************
    public function sqlite_extension_bfsvtab
        ' https://github.com/abetlen/sqlite3-bfsvtab-ext
        dim s: s = ""
        dim retValue: retValue = 0
        dim sSql

        log "******************************************************"
        log "sqlite_extension_bfsvtab"
        opendb "SQL3 "
        log "load_extension(bfsvtab.dll)"
        query2csv("SELECT load_extension('.\install\" & sBitPath & "\bfsvtab.dll') as ext_loaded")
        if aQueryResults(2)(0) <> "Null" then retValue = retValue+1

        query("create table edges(fromNode integer, toNode integer);")
        query("insert into edges(fromNode, toNode) values (1, 2),(1, 3),(2, 4),(3, 4);")

        ' Find the shortest path from node 1 to node 4
        sSql = _
            "select  " & _
            "  id, distance  " & _
            "  from bfsvtab  " & _
            "  where  " & _
            "    tablename  = 'edges'    and " & _
            "    fromcolumn = 'fromNode' and " & _
            "    tocolumn   = 'toNode'   and " & _
            "    root       = 1          and " & _
            "    id         = 4;"
        query2csv(sSql)
        if aQueryResults(2).count <> 1 then retValue = retValue+1

        ' Find the minimum distance from node 1 to all nodes
        sSql = _
            "select  " & _
            "  id, distance  " & _
            "  from bfsvtab  " & _
            "  where  " & _
            "    tablename  = 'edges'    and " & _
            "    fromcolumn = 'fromNode' and " & _
            "    tocolumn   = 'toNode'   and " & _
            "    root       = 1;"
        query2csv(sSql)
        if aQueryResults(2).count <> 4 then retValue = retValue+1

        ' Return the edge set of a spanning tree rooted at node 1
        sSql = _
            "select  " & _
            "  id, distance  " & _
            "  from bfsvtab  " & _
            "  where  " & _
            "    tablename  = 'edges'    and " & _
            "    fromcolumn = 'fromNode' and " & _
            "    tocolumn   = 'toNode'   and " & _
            "    root       = 1          and " & _
            "    parent     is not null;"
        query2csv(sSql)
        if aQueryResults(2).count <> 3 then retValue = retValue+1

        closedb

        log "connection string"
        opendb "SQL3-bfsvtab"

        ' Find the shortest path from node 1 to node 4
        sSql = _
            "select  " & _
            "  id, distance  " & _
            "  from bfsvtab  " & _
            "  where  " & _
            "    tablename  = 'edges'    and " & _
            "    fromcolumn = 'fromNode' and " & _
            "    tocolumn   = 'toNode'   and " & _
            "    root       = 1          and " & _
            "    id         = 4;"
        query2csv(sSql)
        if aQueryResults(2).count <> 1 then retValue = retValue+1

        ' Find the minimum distance from node 1 to all nodes
        sSql = _
            "select  " & _
            "  id, distance  " & _
            "  from bfsvtab  " & _
            "  where  " & _
            "    tablename  = 'edges'    and " & _
            "    fromcolumn = 'fromNode' and " & _
            "    tocolumn   = 'toNode'   and " & _
            "    root       = 1;"
        query2csv(sSql)
        if aQueryResults(2).count <> 4 then retValue = retValue+1

        ' Return the edge set of a spanning tree rooted at node 1
        sSql = _
            "select  " & _
            "  id, distance  " & _
            "  from bfsvtab  " & _
            "  where  " & _
            "    tablename  = 'edges'    and " & _
            "    fromcolumn = 'fromNode' and " & _
            "    tocolumn   = 'toNode'   and " & _
            "    root       = 1          and " & _
            "    parent     is not null;"
        query2csv(sSql)
        if aQueryResults(2).count <> 3 then retValue = retValue+1
        
        closedb

        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_extension_uuid 
        dim s: s = ""
        dim retValue: retValue = 0

        log "******************************************************"
        log "sqlite_extension_uuid"
        opendb "MEM  "
        log "load_extension() throws error but works"
        logResult query2csv("SELECT load_extension('.\install\" & sBitPath & "\uuid.dll') as ext_loaded")
        logResult query2csv("select uuid() as uv4")
        if aQueryResults(2).count <> 1 then retValue = retValue+1
        retValue = retValue+uuid_tests
        closedb

        log "connection string"
        opendb "SQL3-uuid"
        logResult query2csv("select uuid() as uv4")
        if aQueryResults(2).count <> 1 then retValue = retValue+1
        retValue = retValue+uuid_tests
        closedb

        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function uuid_tests
        log "****************************************"
        log "****************************************"
        log "****************************************"
        dim retValue: retValue = 0
        dim i
        dim aa
        dim a: a = Array( _
            "A0EEBC99-9C0B-4EF8-BB6D-6BB9BD380A11",_
            "{a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11}",_
            "a0eebc999c0b4ef8bb6d6bb9bd380a11",_
            "a0ee-bc99-9c0b-4ef8-bb6d-6bb9-bd38-0a11",_
            "{a0eebc99-9c0b4ef8-bb6d6bb9-bd380a11", _
            "{a0eebc99-9c0b4ef8-bb6d6bb9-bd380a11}", _
            "A0EEBC99-9C0B4EF8-BB6D-6BB9BD380A11",_
            "0EEBC99-9C0B-4EF8-BB6D-6BB9BD380A11",_
            "{a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11",_
            "a0eebc999c0b4ef8bb6d6bb9bd380a1",_
            "a0ee-bc99-9c0b-4ef8-bb6d-6bb9-a11",_
            "{a0eebc{99-9c0b4ef8-bb6d6bb9-bd380a11}" _ 
        )
        i=0
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> """a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11""" then retValue = retValue+1
        if aa(1) <> "A0EEBC999CB4EF8BB6D6BB9BD38A11" then retValue = retValue+1

        i=1
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> """a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11""" then retValue = retValue+1
        if aa(1) <> "A0EEBC999CB4EF8BB6D6BB9BD38A11" then retValue = retValue+1

        i=2
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> """a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11""" then retValue = retValue+1
        if aa(1) <> "A0EEBC999CB4EF8BB6D6BB9BD38A11" then retValue = retValue+1

        i=3
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> """a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11""" then retValue = retValue+1
        if aa(1) <> "A0EEBC999CB4EF8BB6D6BB9BD38A11" then retValue = retValue+1

        i=4
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> """a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11""" then retValue = retValue+1
        if aa(1) <> "A0EEBC999CB4EF8BB6D6BB9BD38A11" then retValue = retValue+1

        i=5
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> """a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11""" then retValue = retValue+1
        if aa(1) <> "A0EEBC999CB4EF8BB6D6BB9BD38A11" then retValue = retValue+1

        i=6
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> """a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11""" then retValue = retValue+1
        if aa(1) <> "A0EEBC999CB4EF8BB6D6BB9BD38A11" then retValue = retValue+1

        ' missing A at beginning"
        i=7
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> "Null" then retValue = retValue+1
        if aa(1) <> "Null" then retValue = retValue+1

        '  missing '}' but ok"
        i=8
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> """a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11""" then retValue = retValue+1
        if aa(1) <> "A0EEBC999CB4EF8BB6D6BB9BD38A11" then retValue = retValue+1

        ' number of digits < 16"
        i=9
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> "Null" then retValue = retValue+1
        if aa(1) <> "Null" then retValue = retValue+1

        ' number of digits < 16"
        i=10
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> "Null" then retValue = retValue+1
        if aa(1) <> "Null" then retValue = retValue+1

        ' stray '{'"
        i=11
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        aa = split(aQueryResults(2)(0),",")
        if aa(0) <> "Null" then retValue = retValue+1
        if aa(1) <> "Null" then retValue = retValue+1

        uuid_tests = retValue
    end function

    '********************************************
    public function sqlite_extension_functions_wholenumber
        dim s: s = ""
        dim retValue: retValue = 0

        opendb "MEM  "
        log "******************************************************"
        log "sqlite_extension_functions_wholenumber"
        log "wholenumber extension dll works with SELECT load_extension()"
        log query("SELECT load_extension('.\install\" & sBitPath & "\wholenumber.dll') as loaded;")
        log getSortedFunctionList("whole") & vbcrlf
        log query("drop table if exists nums;")
        log query("CREATE VIRTUAL TABLE nums USING wholenumber;")
        logResult query2csv("SELECT value FROM nums WHERE value<10;")
        if aQueryResults(2).count <> 9 then retValue = retValue+1
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        if aQueryResults(2)(8) <> 9 then retValue = retValue+1
        closedb

        opendb "SQL3-wholenumber"
        log getSortedFunctionList("whole") & vbcrlf
        log query("drop table if exists nums;")
        log query("CREATE VIRTUAL TABLE nums USING wholenumber;")
        logResult query2csv("SELECT value FROM nums WHERE value<10;")
        if aQueryResults(2).count <> 9 then retValue = retValue+1
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        if aQueryResults(2)(8) <> 9 then retValue = retValue+1
        closedb

        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_extension_functions_totype
        dim s: s = ""
        dim retValue: retValue = 0

        log "******************************************************"
        log "sqlite_extension_functions_totype"
        log "totype extension dll does not work with SELECT load_extension()"
        opendb "SQL3-totype"
        ' If X is an integer, real, or string value that can be
        ' losslessly represented as an integer, then tointeger(X)
        ' returns the corresponding integer value.
        ' If X is an 8-byte BLOB then that blob is interpreted as
        ' a signed two-compliment little-endian encoding of an integer
        ' and tointeger(X) returns the corresponding integer value.
        ' Otherwise tointeger(X) return NULL.
        log getSortedFunctionList("tointeger")
        logResult query2csv("SELECT tointeger(8) as i,tointeger(8) as r, tointeger('8') as s;")
        if aQueryResults(2)(0) <> "8,8,8" then retValue = retValue+1
        logResult query2csv("SELECT tointeger(8.1) as i,tointeger(8.1) as r, tointeger('8.1') as s;")
        if aQueryResults(2)(0) <> "Null,Null,Null" then retValue = retValue+1
        logResult query2csv("SELECT tointeger(8.9) as i,tointeger(8.9) as r, tointeger('8.9') as s;")
        if aQueryResults(2)(0) <> "Null,Null,Null" then retValue = retValue+1

        ' If X is an integer, real, or string value that can be
        ' convert into a real number, preserving at least 15 digits
        ' of precision, then toreal(X) returns the corresponding real value.
        ' If X is an 8-byte BLOB then that blob is interpreted as
        ' a 64-bit IEEE754 big-endian floating point value
        ' and toreal(X) returns the corresponding real value.
        ' Otherwise toreal(X) return NULL.
        log getSortedFunctionList("toreal")
        logResult query2csv("SELECT toreal(8) as i,toreal(8) as r, toreal('8') as s;")
        if aQueryResults(2)(0) <> "8,8,8" then retValue = retValue+1
        logResult query2csv("SELECT toreal(8.1) as i,toreal(8.1) as r, toreal('8.1') as s;")
        if aQueryResults(2)(0) <> "8.1,8.1,8.1" then retValue = retValue+1
        logResult query2csv("SELECT toreal(8.9) as i,toreal(8.9) as r, toreal('8.9') as s;")
        if aQueryResults(2)(0) <> "8.9,8.9,8.9" then retValue = retValue+1
        closedb

        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_extension_functions_sha
        dim s: s = ""
        dim retValue: retValue = 0
        dim result

        opendb "MEM  "
        log "******************************************************"
        log "sqlite_extension_functions_sha"
        log "load_extension() throws error, but sha3 methods are loaded (Function sequence error)"
        result = query2csv("SELECT load_extension('.\install\" & sBitPath & "\shathree.dll') as loaded;")
        if aQueryResults(2)(0) <> "Null" then retValue = retValue+1
        if instr(aQueryResults(3),"Function sequence error") = 0 then retValue = retValue+1
        
        ' examples from the source code https://www.sqlite.org/src/file/ext/misc/shathree.c
        query2csv("SELECT sha3(1) = sha3('1');")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("SELECT sha3('hello') = sha3(x'68656c6c6f');")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES('xyzzy')) SELECT sha3_agg(x) = sha3('T5:xyzzy') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES(x'010203')) SELECT sha3_agg(x) = sha3(x'42333a010203') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES(0x123456)) SELECT sha3_agg(x) = sha3(x'490000000000123456') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES(100.015625)) SELECT sha3_agg(x) = sha3(x'464059010000000000') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES(NULL)) SELECT sha3_agg(x) = sha3('N') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        closedb
        
        opendb "SQL3-sha"
        ' examples from the source code https://www.sqlite.org/src/file/ext/misc/shathree.c
        log "loading via connection string works"
        query2csv("SELECT sha3(1) = sha3('1');")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("SELECT sha3('hello') = sha3(x'68656c6c6f');")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES('xyzzy')) SELECT sha3_agg(x) = sha3('T5:xyzzy') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES(x'010203')) SELECT sha3_agg(x) = sha3(x'42333a010203') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES(0x123456)) SELECT sha3_agg(x) = sha3(x'490000000000123456') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES(100.015625)) SELECT sha3_agg(x) = sha3(x'464059010000000000') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        query2csv("WITH a(x) AS (VALUES(NULL)) SELECT sha3_agg(x) = sha3('N') FROM a;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        closedb
        
        ' https://sqlite.org/forum/forumpost/23b2e479a0
        ' The hash is computed over the database content, not its representation on disk. 
        ' This means, for example, that a VACUUM or similar data-preserving transformation does not change the hash.
        ' recursive CTE to create table with 1 million rows then hash a column/table with SHA3 functions
        dbSqlite3 = strFolder & "\testDBs\sha.sqlite3"
        if objFSO.FileExists(dbSqlite3) then 
            objFSO.DeleteFile(dbSqlite3)
        end if
        opendb "SQL3 "
        
        result = query2csv("SELECT load_extension('.\install\" & sBitPath & "\shathree.dll') as loaded;")
        if aQueryResults(2)(0) <> "Null" then retValue = retValue+1
        if instr(aQueryResults(3),"Function sequence error") = 0 then retValue = retValue+1
        
        result = query2csv("create table people (id INTEGER, income REAL, tax_rate REAL);")
        if result = -1 then retValue = retValue + 1
        
        dim q: q = _
            "WITH RECURSIVE person(x) AS ( " & _
            "SELECT 1 UNION ALL SELECT x+1 FROM person LIMIT 1000000 " & _
            ") " & _
            "INSERT INTO people ( id, income, tax_rate) " & _
            "SELECT x, 70+mod(x,15)*3, (15.0+(mod(x,5)*0.2)+mod(x,15))/100. FROM person;"
            
        result = query2csv(q)
        if result = -1 then retValue = retValue + 1
        result = query2csv("select count(1) from people;")
        if aQueryResults(2)(0) <> 1000000 then retValue = retValue + 1
        result = query2csv("select hex(sha3_agg(id)) from people;")
        if result <> 1 then retValue = retValue + 1
        if aQueryResults(2)(0) <> """6274E2DC85CDEB0A4E355B9FF79CFBCB95995779A4E7DDB90E3D312A7CF46278""" then retValue = retValue + 1
        
        log "you could use this to hash the sqlite_schema to prove db has same tables/schema"
        result = query2csv("CREATE TABLE t0(c0 INTEGER);")
        if result <> 0 then retValue = retValue + 1
        result = query2csv("CREATE TABLE t1(c0 INTEGER);")
        if result <> 0 then retValue = retValue + 1
        result = query2csv("CREATE TABLE t2(c0, c1 NOT NULL);")
        if result <> 0 then retValue = retValue + 1
        result = query2csv("CREATE TABLE t3(c0 INTEGER);")
        if result <> 0 then retValue = retValue + 1
        result = query2csv("select * from sqlite_schema;")
        if result <> 5 then retValue = retValue + 1
        
        ' computes a SHA3 hash of the content of sqlite_schema sql column
        result = query2csv("select hex(sha3_agg(sql)) from sqlite_schema;")
        if aQueryResults(2)(0) <> """C915284FCD9F6DE50CDB25B68A42A6EB92089C54CC7F2E5C9BCA1F26007DEA89""" then retValue = retValue + 1
        
        ' computes a SHA3 hash of the content of the people table (SQL must match so that result matches .selftest --init)
        result = query2csv("SELECT hex(sha3_query('SELECT * FROM ""people"" NOT INDEXED',224));")
        if aQueryResults(2)(0) <> """118C9A7D89C7A5A4A672E28C5CB72A4CF434A7D3A3DD0FA73D57CAF6""" then retValue = retValue + 1
        
        query2csv("vacuum;")
        
        ' verify vacuum does not change SHA3 hash
        result = query2csv("SELECT hex(sha3_query('SELECT * FROM ""people"" NOT INDEXED',224));")
        if aQueryResults(2)(0) <> """118C9A7D89C7A5A4A672E28C5CB72A4CF434A7D3A3DD0FA73D57CAF6""" then retValue = retValue + 1
        
        dim sha_result: sha_result = aQueryResults(2)(0)
        log "sha_result " & result & " " & sha_result
        
        closedb
        
        ' make a copy of the database
        log "copy " & strFolder & "\testDBs\sha.sqlite3" & " to " & strFolder & "\testDBs\sha_orig.sqlite3"
        helper_copy_file strFolder & "\testDBs\sha.sqlite3", strFolder & "\testDBs\sha_orig.sqlite3"
        dbSqlite3 = strFolder & "\testDBs\sha_orig.sqlite3"

        ' open the database copy
        opendb "SQL3-sha"
        
        ' verify SHA3 hash of the copy
        on error goto 0
        result = query2csv("SELECT hex(sha3_query('SELECT * FROM ""people"" NOT INDEXED',224));")
        
        if aQueryResults(2)(0) <> sha_result then retValue = retValue + 1
        
        closedb
        
        dbSqlite3 = strFolder & "\testDBs\testfile.sqlite3"
        
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function helper_copy_file(SourceFile,DestinationFile)

        'Check to see if the file already exists in the destination folder
        Dim wasReadOnly
        wasReadOnly = False
        If objFso.FileExists(DestinationFile) Then
        
            'Check to see if the file is read-only
            If objFso.GetFile(DestinationFile).Attributes And 1 Then 
                'The file exists and is read-only.
                'Remove the read-only attribute
                objFso.GetFile(DestinationFile).Attributes = objFso.GetFile(DestinationFile).Attributes - 1
                wasReadOnly = True
            End If

            objFso.DeleteFile DestinationFile, True
        End If

        'Copy the file...overwriting existing file if necessary
        objFso.CopyFile SourceFile, DestinationFile, True

        If wasReadOnly Then
            'Reapply the read-only attribute
            objFso.GetFile(DestinationFile).Attributes = objFso.GetFile(DestinationFile).Attributes + 1
        End If
        
    end function
    
    '********************************************
    public function sqlite_extension_functions_series
        dim s: s = ""
        dim retValue: retValue = 0

        opendb "SQL3 "

        logResult query2csv("select * from PRAGMA_function_list where name like '%series%';") & vbcrlf

        log "******************************************************"
        log "sqlite_extension_functions_series"
        log "not loaded..."
        log query("SELECT * FROM generate_series(0,100,5);")
        log query("SELECT load_extension('.\install\" & sBitPath & "\series.dll') as loaded;")
        log "but not showing up...maybe because it is table-valued-function using a virtual table. "
        s = getSortedFunctionList("series")
        if instr(s, "no items") = 0 then retValue = retValue + 1
        log "load_extension(series.dll) " & retValue
        'logResult query2csv("select * from PRAGMA_function_list where name like '%series%';")
        log "...but is is there...generate_series() expect values from 0 to 100 in steps of 5"
        query2csv("SELECT * FROM generate_series(0,100,5);")
        if aQueryResults(2).count <> 21 then retValue = retValue+1
        if aQueryResults(2)(0) <> 0 then retValue = retValue+1
        if aQueryResults(2)(20) <> 100 then retValue = retValue+1
        closedb
        
        opendb "SQL3-series"
        s = getSortedFunctionList("series")
        if instr(s, "no items") = 0 then retValue = retValue + 1
        log "connection string " & retValue
        log getSortedFunctionList("series") & vbcrlf
        'logResult query2csv("select * from PRAGMA_function_list where name like '%series%';")
        log "...but is is there...generate_series() expect values from 0 to 100 in steps of 5"
        query2csv("SELECT * FROM generate_series(0,100,5);")
        if aQueryResults(2).count <> 21 then retValue = retValue+1
        if aQueryResults(2)(0) <> 0 then retValue = retValue+1
        if aQueryResults(2)(20) <> 100 then retValue = retValue+1
        closedb
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_extension_functions_regex
        dim s: s = ""
        dim retValue: retValue = 0

        log "******************************************************"
        log "sqlite_extension_functions_regex"
        opendb "SQL3 "
        
        s =  getSortedFunctionList("regexp")
        if instr(s, "no items") = 0 then retValue = retValue + 1
        log "regexp not builtin for ODBC driver " & retValue
        
        log query("SELECT load_extension('.\install\" & sBitPath & "\regexp.dll') as loaded")
        s = getSortedFunctionList("regexp")
        if instr(s, "no items") > 0 then retValue = retValue + 1
        log "load_extension errors (Function sequence error) but REGEXP works " & retValue

        query2csv("select 'foobar' regexp 'foo';")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1

        query2csv("select 'foobar' regexp 'bar';")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1

        query2csv("select 'Retroactively relinquishing remunerations is reprehensible.' regexp ' \w{13} ';")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1

        query2csv("select 'Meet me at 10:30' regexp '\d+:\d+';")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1

        closedb

        opendb "SQL3-LoadExt-Regexp"
        s = getSortedFunctionList("regexp")
        if instr(s, "no items") > 0 then retValue = retValue + 1
        log "can be loaded via connection string " & retValue

        query2csv("select 'foobar' regexp 'foo';")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1

        query2csv("select 'foobar' regexp 'foo';")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1

        query2csv("select 'foobar' regexp 'bar';")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1

        query2csv("select 'Retroactively relinquishing remunerations is reprehensible.' regexp ' \w{13} ';")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1

        query2csv("select 'Meet me at 10:30' regexp '\d+:\d+';")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1

        closedb

        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_extension_functions_ieee754
        dim s: s = ""
        dim retValue: retValue = 0
        log "******************************************************"
        log "sqlite_extension_functions_ieee754"
        REM "ieee754.dll does not work with SELECT load_extension()"
        opendb "SQL3-ieee754"
        ' log query("select * from PRAGMA_function_list where name like '%ieee%';") & vbcrlf
        ' log getSortedFunctionList("ieee") & vbcrlf
        log ""
        query2csv("SELECT ieee754(45.25) as r;")
        if aQueryResults(2)(0) <> """ieee754(181,-2)""" then retValue = retValue+1
        query2csv("SELECT ieee754(181,-2) as r;")
        if aQueryResults(2)(0) <> "45.25" then retValue = retValue+1
        query2csv("SELECT ieee754_mantissa(45.25) as m, ieee754_exponent(45.25) as e;")
        if aQueryResults(2)(0) <> "181,-2" then retValue = retValue+1
        query2csv("SELECT ieee754_to_blob(1) as r;")
        if aQueryResults(2)(0) <> "3FF0000000000000" then retValue = retValue+1
        query2csv("SELECT ieee754_from_blob(x'3ff0000000000000') as r;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue+1
        closedb
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_extension_functions_checkfreelist
        dim s: s = ""
        dim retValue: retValue = 0
        log "****************************************************************************"
        log "checkfreelist - needs to be loaded via connection string"
        opendb "SQL3-checkfreelist"
        query2csv("SELECT checkfreelist('main');")
        if aQueryResults(2)(0) <> """ok""" then retValue = retValue+1
        closedb
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_extension_functions_tests
        dim retValue: retValue = 0
        log "******************************************************"
        log "sqlite_extension_functions_tests"
        opendb "MEM  "
        log "load_extension(extension-functions.dll)"
        query2csv("SELECT load_extension('.\install\" & sBitPath & "\extension-functions.dll') as ext_loaded")
        retValue = retValue+sqlite3_extension_functions
        closedb
        log "******************************************************"
        log "via connection string"
        opendb "SQL3-LoadExt-ExtFun"
        retValue = retValue+sqlite3_extension_functions
        closedb
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite3_extension_functions
        dim s: s = ""
        dim retValue: retValue = 0

        log "****************************************************************************"
        log "sqlite3_extension_functions"
        ' charindex(S1,S2)
        query2csv("select charindex('world','hello world!') as x;")
        if aQueryResults(2)(0) <> 7 then retValue = retValue+1
        ' charindex(S1,S2,N)
        query2csv("select charindex('world','hello world! hello world!',10) as x;")
        if aQueryResults(2)(0) <> 20 then retValue = retValue+1
        ' leftstr(S,N)
        query2csv("select leftstr('hello world!',5) as x;")
        if aQueryResults(2)(0) <> """hello""" then retValue = retValue+1
        ' padc(S,N)
        query2csv("select '|', padc('10',50) as x, '|';")
        if aQueryResults(2)(0) <> """|"",""                        10                        "",""|""" then retValue = retValue+1
        ' padl(S,N)
        query2csv("select '|', padl('10',50) as x, '|';")
        if aQueryResults(2)(0) <> """|"",""                                                10"",""|""" then retValue = retValue+1
        ' padr(S,N)
        query2csv("select '|', padr('10',50) as x, '|';")
        if aQueryResults(2)(0) <> """|"",""10                                                "",""|""" then retValue = retValue+1
        ' ' proper(S)
        query2csv("select proper('hElLo WoRlD!') as x;")
        if aQueryResults(2)(0) <> """Hello World!""" then retValue = retValue+1
        ' ' replicate(S,N)
        query2csv("select replicate('hElLo WoRlD fIvE tImEs!',5) as x;")
        if aQueryResults(2)(0) <> """hElLo WoRlD fIvE tImEs!hElLo WoRlD fIvE tImEs!hElLo WoRlD fIvE tImEs!hElLo WoRlD fIvE tImEs!hElLo WoRlD fIvE tImEs!""" then retValue = retValue+1
        ' ' reverse(S)
        query2csv("select reverse('hElLo WoRlD!') as x;")
        if aQueryResults(2)(0) <> """!DlRoW oLlEh""" then retValue = retValue+1
        ' ' rightstr(S,N)
        query2csv("select rightstr('hello world!',6) as x;")
        if aQueryResults(2)(0) <> """world!""" then retValue = retValue+1
        ' ' strfilter(S1,S2)
        query2csv("select strfilter('hello world!','o!') as x;")
        if aQueryResults(2)(0) <> """oo!""" then retValue = retValue+1

        query("SELECT load_extension('.\install\" & sBitPath & "\csv.dll') as loaded")
        query("CREATE VIRTUAL TABLE temp.t1 USING csv(filename='test.csv',header=true)")
        
        ' lower_quartile(X) 
        query2csv("select lower_quartile(col_1) as x from t1;")
        if aQueryResults(2)(0) <> 1.5 then retValue = retValue+1
        ' median(X)
        query2csv("select median(col_1) as x from t1;")
        if aQueryResults(2)(0) <> 2.5 then retValue = retValue+1
        ' mode(X)
        query2csv("select mode(col_2) as x from t1;")
        if aQueryResults(2)(0) <> 2 then retValue = retValue+1
        ' stdev(X)
        query2csv("select stdev(col_1) as x from t1;")
        if aQueryResults(2)(0) <> 1.29099444873581 then retValue = retValue+1
        ' upper_quartile(X)
        query2csv("select upper_quartile(col_1) as x from t1;")
        if aQueryResults(2)(0) <> 3.5 then retValue = retValue+1
        ' variance(X)
        query2csv("select variance(col_1) as x from t1;")
        if aQueryResults(2)(0) <> 1.66666666666667 then retValue = retValue+1
        
        query("drop table if exists t1.temp;")

        sqlite3_extension_functions = retValue
    end function

    '********************************************
    public function sqlite_extension_functions_csv
        dim s: s = ""
        dim retValue: retValue = 0
       
        log "******************************************************"
        log "sqlite_extension_functions_csv"
        opendb "SQL3 "
        log "load_extension(csv.dll) "
        log query("SELECT load_extension('.\install\" & sBitPath & "\csv.dll') as loaded")
        log query("CREATE VIRTUAL TABLE temp.t1 USING csv(filename='test.csv',header=true)")
        logResult query2csv(" SELECT * FROM t1")
        s = aQueryResults(2).count & "x" & ubound(split(aQueryResults(2)(0),","))+1
        if s <> "4x5" then retValue = retValue + 1
        closedb
        log ""
        
        opendb "SQL3-LoadExt-Csv"
        log "CSV Virtual Table via connection string"
        log query("CREATE VIRTUAL TABLE temp.t1 USING csv(filename='test.csv',header=true)")
        logResult query2csv(" SELECT * FROM t1")
        s = aQueryResults(2).count & "x" & ubound(split(aQueryResults(2)(0),","))+1
        if s <> "4x5" then retValue = retValue + 1
        closedb

        if retValue <> 0 then err.raise 1
    end function

    '********************************************
    public function sqlite_extension_crypto 
        log "******************************************************"
        log "sqlite_extension_crypto"

        dim ss: ss = ""
        dim retValue: retValue = 0

        opendb "MEM  "
        ss = getSortedFunctionList("md5")
        if instr(ss, "no items") = 0 then retValue = retValue + 1
        log retValue
        closedb
        log ""
  
        opendb "SQL3 "
        log "load_extension(crypto.dll)"
        query2csv("SELECT load_extension('.\install\" & sBitPath & "\crypto.dll') as ext_loaded")
        if aQueryResults(2)(0) <> "Null" then  retValue = retValue + 1
        query2csv("select hex(md5('abc'));")
        retValue = retValue + abs(strcomp(aQueryResults(2)(0),"""900150983CD24FB0D6963F7D28E17F72"""))
        log retValue & " " & aQueryResults(2)(0)
        query2csv("select hex(sha1('abc'));")
        retValue = retValue + abs(strcomp(aQueryResults(2)(0),"""A9993E364706816ABA3E25717850C26C9CD0D89D"""))
        log retValue & " " & aQueryResults(2)(0)
        query2csv("select hex(sha256('abc'));")
        retValue = retValue + abs(strcomp(aQueryResults(2)(0),"""BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD"""))
        log retValue & " " & aQueryResults(2)(0)
        query2csv("select hex(sha384('abc'));")
        retValue = retValue + abs(strcomp(aQueryResults(2)(0),"""CB00753F45A35E8BB5A03D699AC65007272C32AB0EDED1631A8B605A43FF5BED8086072BA1E7CC2358BAECA134C825A7"""))
        log retValue & " " & aQueryResults(2)(0)
        query2csv("select hex(sha512('abc'));")
        retValue = retValue + abs(strcomp(aQueryResults(2)(0),"""DDAF35A193617ABACC417349AE20413112E6FA4E89A97EA20A9EEEE64B55D39A2192992A274FC1A836BA3C23A3FEEBBD454D4423643CE80E2A9AC94FA54CA49F"""))
        log retValue & " " & aQueryResults(2)(0)
        closedb
        log ""

        log "crypto.dll via connection string."
        opendb "SQL3-crypto"
        query2csv("select hex(md5('abc'));")
        retValue = retValue +  strcomp(aQueryResults(2)(0),"""900150983CD24FB0D6963F7D28E17F72""")
        log retValue & " " & aQueryResults(2)(0)
        query2csv("select hex(sha1('abc'));")
        retValue = retValue +  strcomp(aQueryResults(2)(0),"""A9993E364706816ABA3E25717850C26C9CD0D89D""")
        log retValue & " " & aQueryResults(2)(0)
        query2csv("select hex(sha256('abc'));")
        retValue = retValue +  strcomp(aQueryResults(2)(0),"""BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD""")
        log retValue & " " & aQueryResults(2)(0)
        query2csv("select hex(sha384('abc'));")
        retValue = retValue +  strcomp(aQueryResults(2)(0),"""CB00753F45A35E8BB5A03D699AC65007272C32AB0EDED1631A8B605A43FF5BED8086072BA1E7CC2358BAECA134C825A7""")
        log retValue & " " & aQueryResults(2)(0)
        query2csv("select hex(sha512('abc'));")
        retValue = retValue +  strcomp(aQueryResults(2)(0),"""DDAF35A193617ABACC417349AE20413112E6FA4E89A97EA20A9EEEE64B55D39A2192992A274FC1A836BA3C23A3FEEBBD454D4423643CE80E2A9AC94FA54CA49F""")
        log retValue & " " & aQueryResults(2)(0)
        closedb
        log ""

        if retValue > 0 then err.raise retValue

    end function

    '********************************************
    public function sqlite_double_quoted_strings
        dim retValue: retValue = 0
        opendb "MEM  "
        log "The SQLITE_DBCONFIG_DQS_DML option activates or deactivates the legacy "
        log "double-quoted string literal misfeature for DML statements only, that is "
        log "DELETE, INSERT, SELECT, and UPDATE statements. The recommended setting is 0, "
        log "meaning that double-quoted strings are disallowed in all contexts. "
        log "However, the default setting is 3 for maximum compatibility with legacy applications."
        log "ODBC driver setting is 3 for maximum compatibility. "
        log "Hence both examples below work (1 row returned for table t0)."
        log ""
        query2csv("CREATE TABLE t0(c0 INTEGER);")
        
        query2csv("select * from sqlite_master where type=""table"";")
        if aQueryResults(2).count <> 1 then retValue = retValue+1

        query2csv("select * from sqlite_master where type='table';")
        if aQueryResults(2).count <> 1 then retValue = retValue+1

        closedb
        if retValue > 0 then err.raise retValue
    end function
    
    '********************************************
    public function sqlite_strftime
        ' The strftime() SQL function now supports %G, %g, %U, and %V.
        ' New conversion letters on the strftime() SQL function: %e %F %I %k %l %p %P %R %T %u
        dim retValue: retValue = 0
        opendb "MEM  "
        query2csv("SELECT strftime('%e -- %F -- %I -- %k -- %l -- %p -- %P -- %R -- %T -- %u -- %G -- %g -- %U -- %V', '2013-10-07T08:23:19.120') as r;")
        if aQueryResults(2)(0) <> """ 7 -- 2013-10-07 -- 08 --  8 --  8 -- AM -- am -- 08:23 -- 08:23:19 -- 1 -- 2013 -- 13 -- 40 -- 41""" then 
            retValue = retValue+1 
        end if
        closedb
        if retValue > 0 then err.raise retValue
    end function
    
    '********************************************
    public function sqlite_json
        dim retValue: retValue = 0
        opendb "MEM  "
        ' just a couple of examples to show that JSON works...
        
        '{"this":"is","a":["test"]}'
        query2csv("select json(' { ""this"" : ""is"", ""a"": [ ""test"" ] } ')")
        if aQueryResults(2)(0) <>  """{""""this"""":""""is"""",""""a"""":[""""test""""]}""" then 
            retValue = retValue+1 
        end if

        '[1,2,"3",4]'
        query2csv("select json_array(1,2,'3',4)") 
        if aQueryResults(2)(0) <>  """[1,2,""""3"""",4]""" then 
            retValue = retValue+1 
        end if

        '[""[1,2]"]'")
        query2csv("select json_array('[1,2]')") 
        if aQueryResults(2)(0) <>  """[""""[1,2]""""]""" then 
            retValue = retValue+1 
        end if
        
        '[[1,2]]'")
        query2csv("select json_array(json_array(1,2))") 
        if aQueryResults(2)(0) <>  """[[1,2]]""" then 
            retValue = retValue+1 
        end if
        
        '[1,null,"3","[4,5]","{\"six\":7.7}"]'")
        query2csv("select json_array(1,null,'3','[4,5]','{""six"":7.7}')") 
        if aQueryResults(2)(0) <>  """[1,null,""""3"""",""""[4,5]"""",""""{\""""six\"""":7.7}""""]""" then 
            retValue = retValue+1 
        end if
        
        '[1,null,"3",[4,5],{"six":7.7}]'")
        query2csv("select json_array(1,null,'3',json('[4,5]'),json('{""six"":7.7}'))") 
        if aQueryResults(2)(0) <>  """[1,null,""""3"""",[4,5],{""""six"""":7.7}]""" then 
            retValue = retValue+1 
        end if
        
        '$.c' → '[4,5,{"f":7}]'
        query2csv("select '{""a"":2,""c"":[4,5,{""f"":7}]}' -> '$.c'") 
        if aQueryResults(2)(0) <>  """[4,5,{""""f"""":7}]""" then 
            retValue = retValue+1 
        end if
        
        closedb
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_json_virtual_columns
        ' https://antonz.org/json-virtual-columns/
        
        dim retValue: retValue = 0
        opendb "MEM  "
    
        query2csv("create table events(value TEXT);")
        
        ' recursive CTE to create table
        dim q: q = _
            "WITH RECURSIVE event(x) AS ( " & _
            "     SELECT 1 UNION ALL SELECT x+1 FROM event LIMIT 10000000" & _
            ")" & _
            "INSERT INTO events ( value ) " & _
            "SELECT '{""timestamp"":""2022-05-15T09:31:00Z"",""object"":""user' || x || '"",""object_id"":' || x || ',""action"":""login"",""details"":{""ip"":""192.168.0.1""}}' FROM event;"

        oTimer.StartTimer
        query2csv(q)
        oTimer.StopTimer
        query2csv("select count(1) from events;")
        log "populate table (" & aQueryResults(2)(0) & ") " & oTimer.ResultString
        
        ' slow as you're calling json_extract alot for each query
        q = _
        "select " & _
          "json_extract(value, '$.object') as object, " & _
          "json_extract(value, '$.action') as action, " & _
          "json_extract(value, '$.object_id') as object_id " & _
        "from events " & _
        "where json_extract(value, '$.object_id') = [[[ID]]];"
        
        oTimer.StartTimer
        query2csv( replace(q,"[[[ID]]]","1000000") )
        oTimer.StopTimer
        log aQueryResults(2)(0)
        log "slow query object_id = " & split(aQueryResults(2)(0),",")(2) & " " & oTimer.ResultString

        oTimer.StartTimer
        query2csv( replace(q,"[[[ID]]]","5000000") )
        oTimer.StopTimer
        log aQueryResults(2)(0)
        log "slow query object_id = " & split(aQueryResults(2)(0),",")(2) & " " & oTimer.ResultString

        oTimer.StartTimer
        query2csv( replace(q,"[[[ID]]]","10000000") )
        oTimer.StopTimer
        log aQueryResults(2)(0)
        log "slow query object_id = " & split(aQueryResults(2)(0),",")(2) & " " & oTimer.ResultString

        
        ' pay up front creating virtual columns
        oTimer.StartTimer
        query "alter table events add column object_id integer as (json_extract(value, '$.object_id'));"
        query "alter table events add column object text as (json_extract(value, '$.object'));"
        query "alter table events add column action text as (json_extract(value, '$.action'));"
        query "create index events_object_id on events(object_id);"
        oTimer.StopTimer
        log "index " & oTimer.ResultString
        
        ' now queries are fast
        oTimer.StartTimer
        query2csv("select object, action, object_id from events where object_id = 1000000;")
        oTimer.StopTimer
        log aQueryResults(2)(0)
        log "fast query object_id = " & split(aQueryResults(2)(0),",")(2) & " " & oTimer.ResultString
        
        oTimer.StartTimer
        query2csv("select object, action, object_id from events where object_id = 5000000;")
        oTimer.StopTimer
        log aQueryResults(2)(0)
        log "fast query object_id = " & split(aQueryResults(2)(0),",")(2) & " " & oTimer.ResultString
        
        oTimer.StartTimer
        query2csv("select object, action, object_id from events where object_id = 10000000;")
        oTimer.StopTimer
        log aQueryResults(2)(0)
        log "fast query object_id = " & split(aQueryResults(2)(0),",")(2) & " " & oTimer.ResultString
        
        closedb
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_simple_virtual_columns
        ' https://antonz.org/generated-columns/
        
        on error goto 0
        dim retValue: retValue = 0
        dim result
        opendb "MEM  "
        
        result = query2csv("create table people (id INTEGER, income REAL, tax_rate REAL);")
        if result = -1 then retValue = retValue + 1
        
        ' recursive CTE to create table
        dim q: q = _
            "WITH RECURSIVE person(x) AS ( " & _
            "SELECT 1 UNION ALL SELECT x+1 FROM person LIMIT 100000 " & _
            ") " & _
            "INSERT INTO people ( id, income, tax_rate) " & _
            "SELECT x, 70+mod(x,15)*3, (15.0+(mod(x,5)*0.2)+mod(x,15))/100. FROM person;"
        result = query2csv(q)
        if result = -1 then retValue = retValue + 1
        
        result = query2csv("select count(1) as numPeople, min(income) as minIncome, max(income) as maxIncome, min(tax_rate) as minTaxRate, max(tax_rate) as maxTaxRate from people;")
        if split(aQueryResults(2)(0),",")(0) <> 100000 then retValue = retValue + 1

        q = "alter table people add column tax real as ( income * tax_rate );"
        result = query2csv(q)
        if result = -1 then retValue = retValue + 1
        
        result = query2csv("select id, income, tax_rate, round(tax,2) as tax from people limit 10;")
        if result <> 10 then retValue = retValue + 1

        closedb
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite_right_join
        log "******************************************************"
        log "sqlite_right_join (new feature in 3.39) based on issues found in forum"
        
        'https://sqlite.org/forum/forumpost/b40696f50145d21cd0f280a71a6e8bf15a540327635b37fd87c99429c49f1545
        opendb "MEM  "
        logResult query2csv("CREATE TABLE t0(c0 INTEGER);")
        logResult query2csv("INSERT INTO t0 VALUES('x');")
        logResult query2csv("CREATE TABLE t1(c0 INTEGER);")
        logResult query2csv("INSERT INTO t1 VALUES('y');")
        logResult query2csv("CREATE TABLE t2(c0, c1 NOT NULL);")
        logResult query2csv("INSERT INTO t2 VALUES('a', 'b');")
        logResult query2csv("CREATE TABLE t3(c0 INTEGER);")
        logResult query2csv("INSERT INTO t3 VALUES('c');")
        logResult query2csv("SELECT * FROM t3 LEFT OUTER JOIN t2 INNER JOIN t0 ON t2.c1 RIGHT OUTER JOIN t1 ON t2.c0;")
        logResult query2csv("SELECT * FROM t3 LEFT OUTER JOIN t2 INNER JOIN t0 ON t2.c1 RIGHT OUTER JOIN t1 ON t2.c0 WHERE (((t2.c1) ISNULL));")
        logResult query2csv("SELECT (((t2.c1) ISNULL) IS TRUE) FROM t3 LEFT OUTER JOIN t2 INNER JOIN t0 ON t2.c1 RIGHT OUTER JOIN t1 ON t2.c0;")
        closedb
   
        'https://sqlite.org/forum/forumpost/19ed84f6c5134705b20c94bc20553e2e25975b82491f66997732077828c05f0f
        opendb "MEM  "
        logResult query2csv("CREATE TABLE t1 (c0, c1);")
        logResult query2csv("INSERT INTO t1(c0) VALUES (2);")
        logResult query2csv("CREATE TABLE t2 (c0);")
        logResult query2csv("CREATE INDEX i0 ON t2 (c0) WHERE c0;")
        logResult query2csv("CREATE TABLE t3 (c0);")
        logResult query2csv("INSERT INTO t3 VALUES (1);")
        logResult query2csv("SELECT * FROM t2 RIGHT OUTER JOIN t3 ON t3.c0 LEFT OUTER JOIN t1 ON t2.c0 WHERE t1.c0;")
        logResult query2csv("SELECT (t1.c0 IS TRUE) FROM t2 RIGHT OUTER JOIN t3 ON t3.c0 LEFT OUTER JOIN t1 ON t2.c0;")
        closedb
        
        'https://sqlite.org/forum/forumpost/887a32bceb052f405ba23c9ca258da3d225cc933c0c4141797ff7a9dd765ade4
        opendb "MEM  "
        logResult query2csv("CREATE TABLE t1(a INT);")
        logResult query2csv("CREATE TABLE t2(b INT);")
        logResult query2csv("CREATE TABLE t3(c INTEGER PRIMARY KEY, d INT);")
        logResult query2csv("CREATE INDEX t3d ON t3(d);")
        logResult query2csv("INSERT INTO t3 VALUES(0, 0);")
        logResult query2csv("SELECT * FROM t1 JOIN t2 ON d>b RIGHT JOIN t3 ON true WHERE +d = 0;")
        logResult query2csv("SELECT * FROM t1 JOIN t2 ON d>b RIGHT JOIN t3 ON true WHERE d = 0;")
        closedb
        
        'https://sqlite.org/forum/forumpost/9d86153da92909b879bca2a37c7ad0f1c0c4321dc1d141c0a1766476feca63d1
        opendb "MEM  "
        logResult query2csv("CREATE TABLE t1(a INT,b BOOLEAN);")
        logResult query2csv("CREATE TABLE t2(c INT);  INSERT INTO t2 VALUES(NULL);")
        logResult query2csv("CREATE TABLE t3(d INT);")
        logResult query2csv("SELECT (b IS TRUE) FROM t1 JOIN t3 ON (b=TRUE) RIGHT JOIN t2 ON TRUE;")
        logResult query2csv("SELECT * FROM t1 JOIN t3 ON (b=TRUE) RIGHT JOIN t2 ON TRUE WHERE (b IS TRUE);")
        closedb
        
        'https://sqlite.org/forum/forumpost/a20cc714ca3dde757d488a18a6f9b9f45a0ace85c8f2c5af54c64bfe619bac34
        opendb "MEM  "
        logResult query2csv("CREATE TABLE t1 (ID INTEGER);")
        logResult query2csv("CREATE TABLE t2 (ID INTEGER);")
        logResult query2csv("INSERT INTO t1 VALUES (1);")
        logResult query2csv("INSERT INTO t2 VALUES (2);")
        logResult query2csv("SELECT * FROM t1 FULL JOIN t2 USING (ID);")
        closedb
        
        'https://sqlite.org/forum/forumpost/8f6790b0534cedcb48d000c440fdfaa6b1cb0c2ce0e4a1ee3a3d1768fd6c38b7
        opendb "MEM  "
        logResult query2csv("CREATE TABLE t1(a1 INT);")
        logResult query2csv("CREATE TABLE t2(b2 INT);")
        logResult query2csv("CREATE TABLE t3(c3 INT, d3 INT UNIQUE);")
        logResult query2csv("CREATE TABLE t4(e4 INT, f4 TEXT);")
        logResult query2csv("INSERT INTO t3(c3, d3) VALUES (2, 1);")
        logResult query2csv("INSERT INTO t4(f4) VALUES ('x');")
        logResult query2csv("CREATE INDEX i0 ON t3(c3) WHERE d3 ISNULL;")
        logResult query2csv("ANALYZE main;")
        logResult query2csv("SELECT * FROM t1 LEFT JOIN t2 ON true JOIN t3 ON (b2 IN (a1)) FULL JOIN t4 ON true;")
        logResult query2csv("SELECT 1 FROM t1 LEFT JOIN t2 ON true JOIN t3 ON (b2 IN (a1)) FULL JOIN t4 ON true;")
        closedb
        
    end function
    
    '********************************************
    public function sqlite_full_outer_join
        log "******************************************************"
        log "sqlite_full_outer_join: W3schools example with left and right first followed by full."
        
        'https://www.w3schools.com/sql/sql_join_right.asp
        opendb "MEM  "
        logResult query2csv("CREATE TABLE orders(OrderID integer,CustomerID integer,EmployeeID integer,OrderDate varchar,ShipperID integer);")
        logResult query2csv("INSERT INTO orders VALUES(10308,2,7,'1996-09-18',3);")
        logResult query2csv("INSERT INTO orders VALUES(10309,37,3,'1996-09-19',1);")
        logResult query2csv("INSERT INTO orders VALUES(10310,77,8,'1996-09-20',2);")
        logResult query2csv("CREATE TABLE employees(EmployeeID integer,LastName varchar,FirstName varchar,BirthDate varchar,Photo varchar);")
        logResult query2csv("INSERT INTO employees VALUES(1,'Davolio','Nancy','12/8/1968','EmpID1.pic');")
        logResult query2csv("INSERT INTO employees VALUES(2,'Fuller','Andrew','2/19/1952','EmpID2.pic');")
        logResult query2csv("INSERT INTO employees VALUES(3,'Leverling','Janet','8/30/1963','EmpID3.pic');")
        logResult query2csv("select * from orders;")
        logResult query2csv("select * from employees;")
        log "The following SQL statement will return all orders with employees who placed the order if known:"
        logResult query2csv(_
            "SELECT Orders.OrderID, Employees.LastName, Employees.FirstName " & _
            "FROM Orders " & _
            "LEFT JOIN Employees ON Orders.EmployeeID = Employees.EmployeeID " & _
            "ORDER BY Orders.OrderID;" _
        )
        log "The following SQL statement will return all employees, and any orders they might have placed:"
        logResult query2csv(_
            "SELECT Orders.OrderID, Employees.LastName, Employees.FirstName " & _
            "FROM Orders " & _
            "RIGHT JOIN Employees ON Orders.EmployeeID = Employees.EmployeeID " & _
            "ORDER BY Orders.OrderID;" _
        )
        log "The following SQL statement will return all orders and all employees including ones that did not order and orders placed by unknown employees:"
        logResult query2csv(_
            "SELECT Orders.OrderID, Employees.LastName, Employees.FirstName " & _
            "FROM Orders " & _
            "FULL OUTER JOIN Employees ON Orders.EmployeeID = Employees.EmployeeID " & _
            "ORDER BY Orders.OrderID;" _
        )
        closedb
    end function
    
    '********************************************
    public function sqlite_isDistinctFrom
        log "******************************************************"
        log "sqlite_isDistinctFrom (new in 3.39)"
        
        'The IS NOT DISTINCT FROM operator is an alternative spelling for the IS operator. Likewise, 
        'the IS DISTINCT FROM operator means the same thing as IS NOT. Standard SQL does not support 
        'the compact IS and IS NOT notation. Those compact forms are an SQLite extension. You have to 
        'use the prolix and much less readable IS NOT DISTINCT FROM and IS DISTINCT FROM operators on 
        'other SQL database engines.

        opendb "MEM  "
        log "these should both produce the same answer of 1"
        logResult query2csv("select 2 is true")
        logResult query2csv("select 2 IS NOT DISTINCT FROM true")
        
        log "these should both produce the same answer of 0"
        logResult query2csv("select 2 is not true")
        logResult query2csv("select 2 IS DISTINCT FROM true")
        closedb
    end function
    
    '********************************************
    public function sqlite_havingWithoutGroupBy
        log "******************************************************"
        log "sqlite_havingWithoutGroupBy (new in 3.39)"
        
        'https://sqlite.org/forum/forumpost/b27c979f3d115f3ce19dda945d108aeecd4305c0f999b182e18c4f4624045857
        opendb "MEM  "
        log "prior to 3.39 this would result in Parse error: a GROUP BY clause is required before HAVING"
        logResult query2csv("CREATE TABLE t1(a INT);")
        logResult query2csv("INSERT INTO t1 VALUES(1),(2),(3);")
        logResult query2csv("SELECT sum(a) FROM t1 HAVING sum(a)>0;")
        closedb
    end function
    
    '********************************************
    public function sqlite_localtimeModifierMaintainsFractSecs
        log "******************************************************"
        log "sqlite_localtimeModifierMaintainsFractSecs (fixed in 3.38.1)"
        
        opendb "MEM  "
        log "output should include fractional seconds"
        logResult query2csv("SELECT strftime('%Y-%m-%d %H:%M:%f', 1.234, 'unixepoch', 'localtime') as r;")
        closedb
    end function
    
    '********************************************
    public function sqlite_unixepochFunction
        log "******************************************************"
        log "sqlite_unixepochFunction (added 3.38.0)"
        
        opendb "MEM  "
        log "calling unixepoch() function works..."
        logResult query2csv("select unixepoch() as e;")
        log "t1 and t2 will both be integers. note that t2 ignores the fractional seconds as designed."
        logResult query2csv("select unixepoch('2004-01-01 02:34:56') as t1, unixepoch('2004-01-01 02:34:56.789') as t2;")
        closedb
    end function
    
    '********************************************
    public function sqlite_autoModifier
        log "******************************************************"
        log "sqlite_autoModifier (added 3.38.0)"
        
        opendb "MEM  "
        log "auto intreprets value based on magnitude of input"
        logResult query2csv("SELECT datetime(2459759.67224309,'auto') as j;")
        logResult query2csv("SELECT datetime(1092941466, 'auto') as u;")
        closedb
    end function
    
    '********************************************
    public function sqlite_juliandayModifier
        log "******************************************************"
        log "sqlite_juliandayModifier (added 3.38.0)"
        
        opendb "MEM  "
        log "default is to interpret number as julian day"
        logResult query2csv("SELECT datetime(2459759.67224309) as j;")
        log "force number to the left to be interpreted as unixepoch"
        logResult query2csv("SELECT datetime(1092941466, 'unixepoch') as u;")
        log "force number to the left to be interreted as julianday"
        logResult query2csv("SELECT datetime(2459759.67224309,'julianday') as j;")
        logResult query2csv("SELECT datetime(2459759,'julianday') as j;")
        closedb
    end function
    
    '********************************************
    public function sqlite_strictTable
        log "******************************************************"
        log "sqlite_strictTable (added 3.37)"
        
        'https://sqlite.org/stricttables.html
        
        opendb "MEM  "
        log "without strict"
        logResult query2csv("CREATE TABLE t1(a text);")
        logResult query2csv("INSERT INTO t1 VALUES('000123');")
        logResult query2csv("INSERT INTO t1 VALUES(123);")
        logResult query2csv("SELECT typeof(a), quote(a) FROM t1;")
        logResult query2csv("CREATE TABLE t2(a integer);")
        logResult query2csv("INSERT INTO t2 VALUES('000123');")
        logResult query2csv("INSERT INTO t2 VALUES(123);")
        logResult query2csv("SELECT typeof(a), quote(a) FROM t2;")
        closedb
        
        opendb "MEM  "
        log "strict will coerce values to the column typem, in this case text"
        logResult query2csv("CREATE TABLE t1(a text) STRICT;")
        logResult query2csv("INSERT INTO t1 VALUES('000123');")
        logResult query2csv("INSERT INTO t1 VALUES(123);")
        logResult query2csv("INSERT INTO t1 VALUES('eeeee');")
        logResult query2csv("SELECT typeof(a), quote(a) FROM t1;")
        
        log "strict will coerce values to the column typem, in this case integer"
        logResult query2csv("CREATE TABLE t2(a integer) STRICT;")
        logResult query2csv("INSERT INTO t2 VALUES('000123');")
        logResult query2csv("INSERT INTO t2 VALUES(123);")
        log "attempting to insert string into strict integer will fail...silently which is unfortunate..."
        logResult query2csv("INSERT INTO t1 VALUES('eeeee');")
        log "note no 'eeeee'..."
        logResult query2csv("SELECT typeof(a), quote(a) FROM t2;")
        
        log "Every column definition must specify a datatype for that column."
        logResult query2csv("CREATE TABLE t3(a) STRICT;")
        
        log "without strict you can specify a column without a datatype"
        logResult query2csv("CREATE TABLE t4(a);")
        logResult query2csv("INSERT INTO t4 VALUES('000123');")
        logResult query2csv("INSERT INTO t4 VALUES(123);")
        logResult query2csv("SELECT typeof(a), quote(a) FROM t4;")
        
        log "The datatype must be one of following:INT,INTEGER,REAL,TEXT,BLOB,ANY. varchar and double is not one of these..."
        logResult query2csv("CREATE TABLE t5(a varchar);")
        logResult query2csv("CREATE TABLE t6(a varchar) STRICT;")
        logResult query2csv("CREATE TABLE t7(d double);")
        logResult query2csv("CREATE TABLE t8(d double) STRICT;")
        logResult query2csv("CREATE TABLE t9(i int, ii integer, r real, t text, b blob, a any) STRICT;")
        closedb

        log "The ability to host any type of data in a single column has proven to be remarkably useful over the years. "
        log "In order to continue supporting this ability, even in STRICT tables, the new ANY datatype name is introduced. "
        opendb "MEM  "
        log "In a STRICT table, a column of type ANY always preserves the data exactly as it is received."
        log "result: text '000123'"
        logResult query2csv("CREATE TABLE t1(a ANY) STRICT;")
        logResult query2csv("INSERT INTO t1 VALUES('000123');")
        logResult query2csv("SELECT typeof(a), quote(a) FROM t1;")
        closedb
        
        opendb "MEM  "
        log "For an ordinary non-strict table, a column of type ANY will attempt to convert strings that look like numbers "
        log "into a numeric value, and if successful will store the numeric value rather than the original string."
        log "result: result: integer 123"
        logResult query2csv("CREATE TABLE t1(a ANY);")
        logResult query2csv("INSERT INTO t1 VALUES('000123');")
        logResult query2csv("SELECT typeof(a), quote(a) FROM t1;")
        closedb
    end function

    '********************************************
    private function generate_series_rcte(iStart,iEnd,iStep)
        log "******************************************************"
        dim sRCTE: sRCTE = _
            "WITH RECURSIVE " & _
            "  generate_series(value) AS ( " & _
            "    SELECT $start " & _
            "    UNION ALL " & _
            "    SELECT value+$step FROM generate_series " & _
            "      WHERE value+$step<=$end " & _
            "  ) SELECT value FROM generate_series"
        sRCTE = replace(sRCTE,"$start",iStart)
        sRCTE = replace(sRCTE,"$end",iEnd)
        sRCTE = replace(sRCTE,"$step",iStep)
        generate_series_rcte = sRCTE
    end function
    
    '********************************************
    public function generate_series 
        log "******************************************************"
        log "generate_series"
        opendb "MEM  "
        log "generate_series_rcte() method returns the SQL for a Recursive Common Table Expression"
        log "that generates a series of numbers. It can be used directly or in a subquery."
        logResult query2csv(generate_series_rcte(5,100,5))
        log ""
        logResult query2csv("select * from (" & generate_series_rcte(5,100,5) & ");")
        log ""
        closedb
    end function

    '********************************************
    public function calendarExamples 
        log "******************************************************"
        log "calendarExamples"
        opendb "MEM  "
        log "print out the name of the month given a date..."
        log "sqlite has no locale data so can't do this itself"
        logResult query2csv( _
            "select case strftime('%m', dateField) " & _
                "when '01' then 'January' " & _
                "when '02' then 'Febuary' " & _
                "when '03' then 'March' " & _
                "when '04' then 'April' " & _
                "when '05' then 'May' " & _
                "when '06' then 'June' " & _
                "when '07' then 'July' " & _
                "when '08' then 'August' " & _
                "when '09' then 'September' " & _
                "when '10' then 'October' " & _
                "when '11' then 'November' " & _
                "when '12' then 'December' " & _
                "else '' end as month " & _
            "from (select '2022-08-01' as dateField);" _
        )
        closedb
        log ""
        opendb "MEM  "
        log "this uses recursive CTE to generate dates for an entire month"
        log "and then joins this to some calendar events."
        logResult query2csv( _
            "CREATE TABLE Events(" & _
            "  EventDate NUMERIC," & _
            "  EventDescription TEXT" & _
            ");" _
        )
        logResult query2csv( _
            "INSERT INTO Events VALUES " & _
            " ('2022-02-13','John''s birthday')" & _
            ",('2022-02-14','Valentine''s day')" & _
            ",('2022-02-18','Fishing Trip with Jesse')" & _
            ",('2022-02-27','Deadline - Artemis project')" & _
            ";" _
        )
        logResult query2csv( _
            "WITH RECURSIVE " & _
            "Days(DayDate) AS ( " & _
            "    SELECT date('2022-02-01') " & _
            "    UNION ALL " & _
            "    SELECT date(DayDate,'+1 day') " & _
            "      FROM Days" & _
            "     WHERE DayDate < '2022-02-28' " & _
            "     LIMIT 31 " & _
            "), " & _
            "WeekDays (DayNo, DayName) AS ( " & _
            "    VALUES " & _
            "	   (0, 'Sunday'), " & _
            "	   (1, 'Monday'), " & _
            "	   (2, 'Tuesday'), " & _
            "	   (3, 'Wednesday'), " & _
            "	   (4, 'Thursday'), " & _
            "	   (5, 'Friday'), "  & _
            "	   (6, 'Saturday') " & _
            ") " & _
            "SELECT " & _
            "   WeekDays.Dayname,  " & _
            "   Days.DayDate,  " & _
            "   IFNULL(Events.EventDescription,'') AS Event " & _
            "FROM Days " & _
            "JOIN WeekDays " & _
            "ON WeekDays.DayNo = CAST(strftime('%w',DayDate) AS INT) " & _
            "  LEFT JOIN " & _
            "Events " & _
            "   ON Events.EventDate = Days.DayDate " & _
            "   ORDER BY Days.DayDate " & _
            ";" _
        )
        closedb
    end function
    
    '********************************************
    public function longSqlStringReturn
        REM log "******************************************************"
        REM log "longSqlString"
        REM log ""
        REM log "inspired by https://til.simonwillison.net/sqlite/column-combinations"
        REM log "This reads in a query from a file...the cool part of this SQL is that"
        REM log "it generates SQL that can then itself be executed. This can be handy"
        REM log "if you, say, want to get some kind of info about tables in your database."
        REM log "SQLite open this up as it has some function that allow the query to do"
        REM log "a little introspection of the DB itself. This would work great, except that"
        REM log "ADO defaults columns with strings to 255 characters by default. "
        REM log ""
        REM log "File longStringTest.sql contains a query that identifies patterns in DB rows in terms"
        REM log "of which columns are not null. (see link above for more details)."
        REM log ""
        dim sFileName: sFileName = ".\sql\longStringTest.sql"
        dim oFile: Set oFile = objFSO.OpenTextFile(sFileName)
        dim sFileContentsTemplate: sFileContentsTemplate = oFile.ReadAll
        oFile.close
        dim sSql,r
        set oFile = nothing
        log ""
        
        test 10,100,"TEXT",100,true,"journal=OFF",false,false
        dbSqlite3 = ".\testDBs\SQL3_True_10_100_100_T_OFF_64.sqlite3"
        opendb "SQL3 "
        log query("update test_table set myField_1 = NULL where id <= 5")
        sSql = replace(sFileContentsTemplate,"_T_","test_table")
        log "processing " & sFileName & " for table 'test_table' in " & dbSqlite3
        r = query(sSql)
        log len(r) & "-->" & r
        
        log ""
        log "The problem is that the returned result is truncated to 255 char. The full"
        log "sql query string is over 1000 char long. Lets try for a table with 2 columns."
        log ""
        
        test 10,1,"TEXT",100,true,"journal=OFF",false,false
        dbSqlite3 = ".\testDBs\SQL3_True_10_1_100_T_OFF_64.sqlite3"
        opendb "SQL3 "
        log query("update [test_table] set myField_1 = null where id <= 5")
        sSql = replace(sFileContentsTemplate,"_T_","test_table")
        log "processing " & sFileName & " for table 'test_table' in " & dbSqlite3
        r = query(sSql)
        log len(r) & "-->" & r
        
        log "OK the resulting query is approx 244 char so less than 255 char limit. "
        log "Use split to grab SQL and execute"
        r = split(r,":")(2)
        logResult query2csv(r)
        
        log ""
        log "Expected result:"
        log "returned 2 rows"
        log "columns(adVarWChar),num_rows(adInteger)"
        log """id, myField_1, "",5            <--- 5 records where both columns are not NULL"
        log """id, "",5                       <--- 5 records where only id is not NULL"
        
        log ""
        log "Issue I'm illustrating here is ADO default of 255 chars for strings "
        log "unless they are defined in a create table statement."
        log ""
        log "By the way, this will work fine in C# and other non-ADO places."
        log ""
        closedb
        
        log "just to prove the point, create a table with varchar(300) and fill with strings of increasing size"
        log ""
        opendb "MEM  "
        log query ( _
            "CREATE TABLE t ( " & _
                "len integer, " & _
                "str varchar(300)" & _
            ")" _
        )
        sSql = "INSERT INTO t ( len, str ) VALUES "
        dim i,j: for i = 0 to 303
            j = j & (i mod 10)
            if i >= 294 then sSql = sSql & "(" & i+1 & ",'" & j & "'),"
        next
        sSql = left(sSql,len(sSql)-1) ' remove the last comma
        query(sSql)
        logResult query2csv("select * from t;")
        log "scroll to right and see how returned string is truncated at length = 300 "
        closedb
        
    end function
    
    '********************************************
    public function unicodeCharacterTest 
        opendb "MEM  "
        log ""
        log query ( _
            "DROP TABLE IF EXISTS tblCountry; " _
        )
        log ""
        log query ( _
            "CREATE TABLE tblCountry ( " & _
            "    Country TEXT, " & _
            "    ISO2 TEXT, " & _
            "    ISO3 TEXT, " & _
            "    ISONum INT, " & _
            "    PRIMARY KEY ( " & _
            "        Country, " & _
            "        ISO2, " & _
            "        ISO3, " & _
            "        ISONum " & _
            "        ) " & _
            "); " _
        )
        log ""
        log "Note the second and third entries contain unicode fox: 🦊 A with circle on top: Å"
        log "You might need to change the font in your editor to see it properly. Try consolas."
        log "Everything works in sqlite/odbc just fine, but probably does not display correctly in CMD"
        log "With consolas, the A with circle on top should display, but fox will be box with ?."
        log ""
        log query ( _
            "INSERT INTO tblCountry ( " & _
            "	Country, " & _
            "	ISO2, " & _
            "	ISO3, " & _
            "	ISONum " & _
            ") VALUES " & _
            "('Afghanistan','AF','AFG',4)," & _
            "('🦊y Lady Land','XX','AFG',4)," & _
            "('Å Islands','AX','ALA',248); " _
        )
        
        log query("select * from  tblCountry;")
        log ""
        closedb    
    end function

    '********************************************
    public function varCharStringIssue 
        ' Issue with returning varchar strings via ADO Recordset
        ' Truncated results would not return correct buffer size which broke printing in VBScript
        ' Note this issue does not happen in C#/C++ examples
        ' References
        ' https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlextendedfetch-function?view=sql-server-ver15
        ' https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlgetdata-function?view=sql-server-ver15
        opendb "MEM  "

        dim sSql
        dim s: s = ""
        dim i
        log "*****************************************"
        log "varCharStringIssue"
        log ""
        log "create table"
        query( _
            "CREATE TABLE test( " & _
            "  typeText text, " & _
            "  typeVarchar varchar(64) " & _
            "); " _
        )
        log ""
        log "popupate table"
        for i = 1 to 70
            s = s & (i mod 10)
            if i > 60 then 
                sSql = "INSERT INTO test VALUES('" & s & "','" & s & "');"
                query(sSql)
            end if
        next
        log ""
        log "check to make sure table has correct count (should be 10)"
        log query("select count(1) as cnt from test;")
        log ""
        log "note the output type adLongVarWChar (203)"
        log "this will return all 10 entries for typeText."
        log query("select typeText from test;")
        log ""
        log ""
        log "Note the output type text (202)."
        log "After the fix, this should now return all ten entries."
        log query("select typeVarchar from test;")
        log ""
        log ""
        log "Here is a 'select [string];' query. In this case the default return string buffer is 512 bytes which"
        log "is 256 wchars. So everything after 255 is truncated."
        log "Note that length of string in SQL is increasing, however, returned string capped at 256 characters."
        log "3.47 it appears the buffer length is fixed"
        s = ""
        for i = 1 to 258
            s = s & (i mod 10)
            if i >= 250 then 
                dim r: r = query("select length('" & s & "') as len, '" & s & "' as longText;")
                log split(r,vbcrlf)(1)
            end if
        next
        log ""
        log ""
        
        closedb
    end function

    '********************************************
    public function graphExampleOne 
        opendb "MEM  "

        ' Recursive Common Table Expressions
        ' 3.4. Controlling Depth-First Versus Breadth-First Search Of a Tree Using ORDER BY
        ' from https://www.sqlite.org/lang_with.html
        log ""
        log "create the table (the order is scrambled on purpose here):"
        log ""
        log query( _
            "CREATE TABLE org( " & _
            "  name TEXT PRIMARY KEY, " & _
            "  boss TEXT REFERENCES org " & _
            ") WITHOUT ROWID; " & _
            "INSERT INTO org VALUES('Dave','Bob'); " & _
            "INSERT INTO org VALUES('Fred','Cindy'); " & _
            "INSERT INTO org VALUES('Gail','Cindy'); " & _
            "INSERT INTO org VALUES('Alice',NULL); " & _
            "INSERT INTO org VALUES('Cindy','Alice'); " & _
            "INSERT INTO org VALUES('Emma','Bob'); " & _
            "INSERT INTO org VALUES('Bob','Alice'); " & _
            "" _
        )
        log ""
        log "Here is a query to show the tree structure in a breadth-first pattern:"
        log ""
        log query( _
            "WITH RECURSIVE " & _
            "  under_alice(name,level) AS ( " & _
            "    VALUES('Alice',0) " & _
            "    UNION ALL " & _
            "    SELECT org.name, under_alice.level+1 " & _
            "      FROM org JOIN under_alice ON org.boss=under_alice.name " & _
            "     ORDER BY 2 " & _
            "  ) " & _
            "SELECT (substr('..........',1,level*3) || name) as result FROM under_alice; " _
        )
        log ""
        log "same as previous but without the ORDER BY 2 which also results in breadth-first pattern:"
        log ""
        log query( _
            "WITH RECURSIVE " & _
            "  under_alice(name,level) AS ( " & _
            "    VALUES('Alice',0) " & _
            "    UNION ALL " & _
            "    SELECT org.name, under_alice.level+1 " & _
            "      FROM org JOIN under_alice ON org.boss=under_alice.name " & _
            "  ) " & _
            "SELECT (substr('..........',1,level*3) || name) as result FROM under_alice; " _
        )
        log ""
        log "resulting in a depth-first search:"
        log ""
        log query( _
            "WITH RECURSIVE " & _
            "  under_alice(name,level) AS ( " & _
            "    VALUES('Alice',0) " & _
            "    UNION ALL " & _
            "    SELECT org.name, under_alice.level+1 " & _
            "      FROM org JOIN under_alice ON org.boss=under_alice.name " & _
            "     ORDER BY 2 DESC " & _
            "  ) " & _
            "SELECT (substr('..........',1,level*3) || name) as result FROM under_alice; " _
        )

        closedb
    end function
    
    '********************************************
    public function graphExampleTwo
        opendb "MEM  "

        ' Recursive Common Table Expressions
        ' 3.3. Queries Against A Graph
        ' from https://www.sqlite.org/lang_with.html
        log ""
        log "create the undirected graph table:"
        log ""
        log query( _
            "CREATE TABLE edge(aa INT, bb INT); " & _
            "CREATE INDEX edge_aa ON edge(aa); " & _
            "CREATE INDEX edge_bb ON edge(bb); " & _
            "INSERT INTO edge VALUES(1,2); " & _
            "INSERT INTO edge VALUES(1,3); " & _
            "INSERT INTO edge VALUES(1,4); " & _
            "INSERT INTO edge VALUES(4,5); " & _
            "INSERT INTO edge VALUES(6,7); " & _
            "INSERT INTO edge VALUES(7,8); " & _
            "INSERT INTO edge VALUES(7,9); " & _
            "INSERT INTO edge VALUES(10,9); " & _
            "INSERT INTO edge VALUES(11,12); " & _
            "" _
        )

        log "find all nodes of the graph that are connected to node 3"
        log query( _
            "WITH RECURSIVE nodes(x,y) AS ( " & _
            "   SELECT 3, 'base' as dir " & _
            "   UNION " & _
            "   SELECT aa , bb || '->' || aa as dir FROM edge JOIN nodes ON bb=x " & _
            "   UNION " & _
            "   SELECT bb, aa || '->' || bb as dir FROM edge JOIN nodes ON aa=x " & _
            ") " & _
            "SELECT x,y FROM nodes; " & _
            "" _
        )

        closedb
    end function
    
    '********************************************
    public function isValidIntOrFloat 
        ' Figuring out if a text value in SQLite is a valid integer or float
        ' https://til.simonwillison.net/sqlite/text-value-is-integer-or-float
        opendb "MEM  "
        log "Figuring out if a text value in SQLite is a valid integer or float"
        log "Note that for Null, ODBC returns type adVarWChar (202) for Field.Type," 
        log "and typename(Field.Value) = 'Null'. Result returned by query() will contain 'Null'."
        log ""
        log "float version - expect the following:"
        log "value(text):Null,is_valid_float(text):Null"
        log "value(text):1,is_valid_float(text):1"
        log "value(text):1.1,is_valid_float(text):1"
        log "value(text):dog,is_valid_float(text):0"
        log ""
        log query( _
            "select " & _
            "  value, " & _
            "  cast(cast(value AS REAL) AS TEXT) in (value, value || '.0') as is_valid_float " & _
            "from " & _
            "  ( " & _
            "    select " & _
            "      '1' as value " & _
            "    union " & _
            "    select " & _
            "      '1.1' as value " & _
            "    union " & _
            "    select " & _
            "      'dog' as value " & _
            "    union " & _
            "    select " & _
            "      null as value " & _
            "  ) " _
        )
        log ""
        log "integer version - expect the following:"
        log "value(text):Null,is_valid_int(text):Null"
        log "value(text):1,is_valid_int(text):1"
        log "value(text):1.1,is_valid_int(text):0"
        log "value(text):dog,is_valid_int(text):0"
        log ""
        log query( _
            "select " & _
            "  value, " & _
            "  cast(cast(value AS INTEGER) AS TEXT) = value as is_valid_int " & _
            "from " & _
            "  ( " & _
            "    select " & _
            "      '1' as value " & _
            "    union " & _
            "    select " & _
            "      '1.1' as value " & _
            "    union " & _
            "    select " & _
            "      'dog' as value " & _
            "    union " & _
            "    select " & _
            "      null as value " & _
            "  ) " _
        )
        log ""
        closedb
    end function

    '********************************************
    public function testDbInventory
        ' loop through a folder of sqlite database files 
        dim sPath: sPath = objFSO.GetAbsolutePathName(".\testDBs")
        dim oFolder: set oFolder = objFSO.GetFolder(sPath)
        log "path: " & oFolder.path & vbcrlf
        log "files:"
        dim oFile: for each oFile in oFolder.Files
            if objFSO.GetExtensionName(oFile.Path) = "sqlite3" then
                dbSqlite3 = oFile.path
                log dbSqlite3
                opendb "SQL3 "
                log query(" select name from sqlite_master where type = 'table';")
                closedb
            end if
        next
        log ""
    end function
    
    '********************************************
    public function recursiveCTE
            ' reproduce example from simplest recursive Common Table Expression (CTE)
            ' https://til.simonwillison.net/sqlite/simple-recursive-cte
            ' SQLite reference: https://www.sqlite.org/lang_with.html
            
            opendb "MEM  "
            
            ' really simple recursive CTE
            log query( _
                "with recursive counter(x) as ( " & _
                "  select 0 " & _
                "    union " & _
                "  select x + 1 from counter " & _
                ") " & _
                "select * from counter limit 5; " _
            )
            
            ' without the (x) - THIS WILL FAIL WITH ERROR no such column: x
            log query( _
                "with recursive counter as ( " & _
                "  select 0 " & _
                "    union " & _
                "  select x + 1 from counter " & _
                ") " & _
                "select * from counter limit 5; " _
            ) & vbcrlf
            
            ' "as x" fixes this. (x) is just a way of defining column names
            log query( _
                "with recursive counter as ( " & _
                "  select 0 as x " & _
                "    union " & _
                "  select x + 1 from counter " & _
                ") " & _
                "select * from counter limit 5; " _
            )

            ' two argument x and y example
            log query( _
                "with recursive counter(x,y) as ( " & _
                "  select 0 as x, 1 as y " & _
                "    union " & _
                "  select x + 1, y + 2 from counter " & _
                ") " & _
                "select * from counter limit 5; " _
            )
            
            ' similar example from https://www.sqlite.org/lang_with.html
            log "The sum of all natural numbers from 1 to 100 is 5050"
            log query( _
                "WITH RECURSIVE " & _
                "  cnt(x) AS ( " & _
                "     SELECT 1 " & _
                "     UNION ALL " & _
                "     SELECT x+1 " & _
                "     FROM cnt " & _
                "     LIMIT 100 " & _
                ") SELECT sum(x) as sumX FROM cnt;" _
            )

            closedb
    end function
    
    '********************************************
    public function sqlite_big_numbers
        log "******************************************************"
        log "sqlite_big_numbers"
        opendb "MEM  "
        
        ' reproducing some tests/examples from this site
        ' http://jakegoulding.com/blog/2011/02/06/sqlite-64-bit-integers/
        ' SQL to run in sqlite3.exe
        REM CREATE TABLE big_numbers (i INTEGER, r REAL, t TEXT, b BLOB);
        REM INSERT INTO big_numbers VALUES (9223372036854775807, 9223372036854775807, 9223372036854775807, 9223372036854775807); -- 2^63 - 1
        REM INSERT INTO big_numbers VALUES (9223372036854775808, 9223372036854775808, 9223372036854775808, 9223372036854775808); -- 2^63
        REM SELECT * FROM big_numbers;
        REM SELECT typeof(i),typeof(r),typeof(t),typeof(b) FROM big_numbers;
        REM SELECT i+1,r+1,t+1,b+1 FROM big_numbers;
        
        ' sqlite3 oputput:
        ' SQLite version 3.36.0 2021-06-18 18:36:39
        ' Enter ".help" for usage hints.
        ' Connected to a transient in-memory database.
        ' Use ".open FILENAME" to reopen on a persistent database.
        ' sqlite> CREATE TABLE big_numbers (i INTEGER, r REAL, t TEXT, b BLOB);
        ' sqlite> INSERT INTO big_numbers VALUES (9223372036854775807, 9223372036854775807, 9223372036854775807, 9223372036854775807); -- 2^63 - 1
        ' sqlite> INSERT INTO big_numbers VALUES (9223372036854775808, 9223372036854775808, 9223372036854775808, 9223372036854775808); -- 2^63
        ' sqlite> SELECT * FROM big_numbers;
        ' 9223372036854775807|9.22337203685478e+18|9223372036854775807|9223372036854775807
        ' 9.22337203685478e+18|9.22337203685478e+18|9.22337203685478e+18|9.22337203685478e+18
        ' sqlite> SELECT typeof(i),typeof(r),typeof(t),typeof(b) FROM big_numbers;
        ' integer|real|text|integer
        ' real|real|text|real
        ' this is somewhat unexpected...adding 1 to the assumed unsigned long value does not
        ' cause it to roll over to negative but instead returns real.
        ' sqlite> SELECT i+1,r+1,t+1,b+1 FROM big_numbers;
        ' 9.22337203685478e+18|9.22337203685478e+18|9.22337203685478e+18|9.22337203685478e+18
        ' 9.22337203685478e+18|9.22337203685478e+18|9.22337203685478e+18|9.22337203685478e+18

        ' however ODBC driver outputs  unexpected results...
        log query("drop table if exists big_numbers;")
        log query("CREATE TABLE big_numbers (i INTEGER, r REAL, t TEXT, b BLOB);")
        log "2^63 - 1"
        log query("INSERT INTO big_numbers VALUES (9223372036854775807, 9223372036854775807, 9223372036854775807, 9223372036854775807);")
        log "2^63"
        log query("INSERT INTO big_numbers VALUES (9223372036854775808, 9223372036854775808, 9223372036854775808, 9223372036854775808);")
        'log "doubles"
        'log query("INSERT INTO big_numbers VALUES (9.51e+18, 9.51e+18, 9.51e+18, 9.51e+18);")
        log ""
        log "expect this to match sqlite3.exe output, however integer differs"
        log "this is sqlite3 output:"
        log "i                     r                     t                     b"
        log "--------------------  --------------------  --------------------  --------------------"
        log "9223372036854775807   9.22337203685478e+18  9223372036854775807   9223372036854775807"
        log "9.22337203685478e+18  9.22337203685478e+18  9.22337203685478e+18  9.22337203685478e+18"
        log ""
        log "ODBC output via script (note the values of column i):"
        log query("SELECT * FROM big_numbers;")
        log ""
        log "The issue here is that vbscript does not have an unsigned integer, so 0x7FFFFFFFFFFFFFFF"
        log "intrepreted as signed integer == 2147483647 so this makes sense. Adding 1 causes SQLite to "
        log "return a real in scientific notation however the recordset is expecting a long. It would"
        log "seem that it is taking some number of chars from the left and casting to long removing the"
        log "fractional part? anyway the issue here is SQLite returning a real and how that value inserted"
        log "into long field in recordset is interpreted."
        log ""
        log "expect query output to match this sqlite3.exe output:"
        log "typeof(i)   typeof(r)   typeof(t)   typeof(b)"
        log "----------  ----------  ----------  ----------"
        log "integer     real        text        integer"
        log "real        real        text        real"
        log ""
        log query("SELECT typeof(i),typeof(r),typeof(t),typeof(b) FROM big_numbers;")
        log "same issue trying to shove real into long in recordset"
        log query("SELECT i+1,r+1,t+1,b+1 FROM big_numbers;")
        log query("SELECT i, i+1 FROM big_numbers;")
        log ""
        log "using cast() helps force the type of the value but then the integer value is now a real"
        log query("SELECT cast(i as real) as i, cast((i+1) as real) as 'i+1' FROM big_numbers;")

        closedb
    end function

    '********************************************
    Function BinaryToString(Binary,bHex)
        'Antonin Foller, http://www.motobit.com
        'Optimized version of a simple BinaryToString algorithm.

        Dim cl1, cl2, cl3, pl1, pl2, pl3
        Dim L,v
        cl1 = 1
        cl2 = 1
        cl3 = 1
        L = LenB(Binary)

        Do While cl1<=L
            v = AscB(MidB(Binary,cl1,1))
            if bHex then
                if v <> 0 then
                    pl3 = pl3 & Hex(v)
                else
                    pl3 = pl3 & "00"
                end if
            else
                if v <> 0 then
                    pl3 = pl3 & Chr(v)
                else
                    pl3 = pl3 & "."
                end if
            end if
            cl1 = cl1 + 1
            cl3 = cl3 + 1
            If cl3>300 Then
                pl2 = pl2 & pl3
                pl3 = ""
                cl3 = 1
                cl2 = cl2 + 1
                If cl2>200 Then
                    pl1 = pl1 & pl2
                    pl2 = ""
                    cl2 = 1
                End If
            End If
        Loop
        BinaryToString = pl1 & pl2 & pl3
    End Function

    '********************************************
    public function sqlite_extension_geopoly
        ' https://sqlite.org/geopoly.html
        ' https://sqlite.org/src/file/ext/rtree/visual01.txt
        ' https://sqlite.org/tmp/geopoly-demo.html
        log "****************************************************************************"
        log "sqlite_extension_geopoly (builtin to SQLiteODBC driver)"
        log ""
        dim sSql
        opendb "SQL3 "
        
        sSql = _
            "CREATE TEMP TABLE basis(name TEXT, jshape TEXT); " & _
            "INSERT INTO basis(name,jshape) VALUES " & _
            "  ('box-20','[[0,0],[20,0],[20,20],[0,20],[0,0]]'), " & _
            "  ('house-70','[[0,0],[50,0],[50,50],[25,70],[0,50],[0,0]]'), " & _
            "  ('line-40','[[0,0],[40,0],[40,5],[0,5],[0,0]]'), " & _
            "  ('line-80','[[0,0],[80,0],[80,7],[0,7],[0,0]]'), " & _
            "  ('arrow-50','[[0,0],[25,25],[0,50],[15,25],[0,0]]'), " & _
            "  ('triangle-30','[[0,0],[30,0],[15,30],[0,0]]'), " & _
            "  ('angle-30','[[0,0],[30,0],[30,30],[26,30],[26,4],[0,4],[0,0]]'), " & _
            "  ('star-10','[[1,0],[5,2],[9,0],[7,4],[10,8],[7,7],[5,10],[3,7],[0,8],[3,4],[1,0]]'); " & _
            "CREATE TEMP TABLE xform(A,B,C,D,clr); " & _
            "INSERT INTO xform(A,B,clr) VALUES " & _
            "  (1,0,'black'), " & _
            "  (0.707,0.707,'blue'), " & _
            "  (0.5,0.866,'red'), " & _
            "  (-0.866,0.5,'green'); " & _
            "CREATE TEMP TABLE xyoff(id1,id2,xoff,yoff,PRIMARY KEY(id1,id2,xoff,yoff)) " & _
            "  WITHOUT ROWID; " & _
            "INSERT INTO xyoff VALUES(1,1,811,659); " & _
            "INSERT INTO xyoff VALUES(1,1,235,550); " & _
            "INSERT INTO xyoff VALUES(1,1,481,620); " & _
            "INSERT INTO xyoff VALUES(1,1,106,494); " & _
            "INSERT INTO xyoff VALUES(1,1,487,106); " & _
            "INSERT INTO xyoff VALUES(1,1,817,595); " & _
            "INSERT INTO xyoff VALUES(1,1,240,504); " & _
            "INSERT INTO xyoff VALUES(1,1,806,457); " & _
            "INSERT INTO xyoff VALUES(1,1,608,107); " & _
            "INSERT INTO xyoff VALUES(1,1,768,662); " & _
            "INSERT INTO xyoff VALUES(1,2,808,528); " & _
            "INSERT INTO xyoff VALUES(1,2,768,528); " & _
            "INSERT INTO xyoff VALUES(1,2,771,171); " & _
            "INSERT INTO xyoff VALUES(1,2,275,671); " & _
            "INSERT INTO xyoff VALUES(1,2,326,336); " & _
            "INSERT INTO xyoff VALUES(1,2,690,688); " & _
            "INSERT INTO xyoff VALUES(1,2,597,239); " & _
            "INSERT INTO xyoff VALUES(1,2,317,528); " & _
            "INSERT INTO xyoff VALUES(1,2,366,223); " & _
            "INSERT INTO xyoff VALUES(1,2,621,154); " & _
            "INSERT INTO xyoff VALUES(1,3,829,469); " & _
            "INSERT INTO xyoff VALUES(1,3,794,322); " & _
            "INSERT INTO xyoff VALUES(1,3,358,387); " & _
            "INSERT INTO xyoff VALUES(1,3,184,444); " & _
            "INSERT INTO xyoff VALUES(1,3,729,500); " & _
            "INSERT INTO xyoff VALUES(1,3,333,523); " & _
            "INSERT INTO xyoff VALUES(1,3,117,595); " & _
            "INSERT INTO xyoff VALUES(1,3,496,201); " & _
            "INSERT INTO xyoff VALUES(1,3,818,601); " & _
            "INSERT INTO xyoff VALUES(1,3,541,343); " & _
            "INSERT INTO xyoff VALUES(1,4,603,248); " & _
            "INSERT INTO xyoff VALUES(1,4,761,649); " & _
            "INSERT INTO xyoff VALUES(1,4,611,181); " & _
            "INSERT INTO xyoff VALUES(1,4,607,233); " & _
            "INSERT INTO xyoff VALUES(1,4,860,206); " & _
            "INSERT INTO xyoff VALUES(1,4,310,231); " & _
            "INSERT INTO xyoff VALUES(1,4,727,539); " & _
            "INSERT INTO xyoff VALUES(1,4,660,661); " & _
            "INSERT INTO xyoff VALUES(1,4,403,133); " & _
            "INSERT INTO xyoff VALUES(1,4,619,331); " & _
            "INSERT INTO xyoff VALUES(2,1,712,578); " & _
            "INSERT INTO xyoff VALUES(2,1,567,313); " & _
            "INSERT INTO xyoff VALUES(2,1,231,423); " & _
            "INSERT INTO xyoff VALUES(2,1,490,175); " & _
            "INSERT INTO xyoff VALUES(2,1,898,353); " & _
            "INSERT INTO xyoff VALUES(2,1,589,483); " & _
            "INSERT INTO xyoff VALUES(2,1,188,462); " & _
            "INSERT INTO xyoff VALUES(2,1,720,106); " & _
            "INSERT INTO xyoff VALUES(2,1,793,380); " & _
            "INSERT INTO xyoff VALUES(2,1,154,396); " & _
            "INSERT INTO xyoff VALUES(2,2,324,218); " & _
            "INSERT INTO xyoff VALUES(2,2,120,327); " & _
            "INSERT INTO xyoff VALUES(2,2,655,133); " & _
            "INSERT INTO xyoff VALUES(2,2,516,603); " & _
            "INSERT INTO xyoff VALUES(2,2,529,572); " & _
            "INSERT INTO xyoff VALUES(2,2,481,212); " & _
            "INSERT INTO xyoff VALUES(2,2,802,107); " & _
            "INSERT INTO xyoff VALUES(2,2,234,509); " & _
            "INSERT INTO xyoff VALUES(2,2,501,269); " & _
            "INSERT INTO xyoff VALUES(2,2,349,553); " & _
            "INSERT INTO xyoff VALUES(2,3,495,685); " & _
            "INSERT INTO xyoff VALUES(2,3,897,372); " & _
            "INSERT INTO xyoff VALUES(2,3,350,681); " & _
            "INSERT INTO xyoff VALUES(2,3,832,257); " & _
            "INSERT INTO xyoff VALUES(2,3,778,149); " & _
            "INSERT INTO xyoff VALUES(2,3,683,426); " & _
            "INSERT INTO xyoff VALUES(2,3,693,217); " & _
            "INSERT INTO xyoff VALUES(2,3,746,317); " & _
            "INSERT INTO xyoff VALUES(2,3,805,369); " & _
            "INSERT INTO xyoff VALUES(2,3,336,585); " & _
            "INSERT INTO xyoff VALUES(2,4,890,255); " & _
            "INSERT INTO xyoff VALUES(2,4,556,565); " & _
            "INSERT INTO xyoff VALUES(2,4,865,555); " & _
            "INSERT INTO xyoff VALUES(2,4,230,293); " & _
            "INSERT INTO xyoff VALUES(2,4,247,251); " & _
            "INSERT INTO xyoff VALUES(2,4,730,563); " & _
            "INSERT INTO xyoff VALUES(2,4,318,282); " & _
            "INSERT INTO xyoff VALUES(2,4,220,431); " & _
            "INSERT INTO xyoff VALUES(2,4,828,336); " & _
            "INSERT INTO xyoff VALUES(2,4,278,525); " & _
            "INSERT INTO xyoff VALUES(3,1,324,656); " & _
            "INSERT INTO xyoff VALUES(3,1,625,362); " & _
            "INSERT INTO xyoff VALUES(3,1,155,570); " & _
            "INSERT INTO xyoff VALUES(3,1,267,433); " & _
            "INSERT INTO xyoff VALUES(3,1,599,121); " & _
            "INSERT INTO xyoff VALUES(3,1,873,498); " & _
            "INSERT INTO xyoff VALUES(3,1,789,520); " & _
            "INSERT INTO xyoff VALUES(3,1,656,378); " & _
            "INSERT INTO xyoff VALUES(3,1,831,601); " & _
            "INSERT INTO xyoff VALUES(3,1,256,471); " & _
            "INSERT INTO xyoff VALUES(3,2,332,258); " & _
            "INSERT INTO xyoff VALUES(3,2,305,463); " & _
            "INSERT INTO xyoff VALUES(3,2,796,341); " & _
            "INSERT INTO xyoff VALUES(3,2,830,229); " & _
            "INSERT INTO xyoff VALUES(3,2,413,271); " & _
            "INSERT INTO xyoff VALUES(3,2,269,140); " & _
            "INSERT INTO xyoff VALUES(3,2,628,441); " & _
            "INSERT INTO xyoff VALUES(3,2,747,643); " & _
            "INSERT INTO xyoff VALUES(3,2,584,435); " & _
            "INSERT INTO xyoff VALUES(3,2,784,314); " & _
            "INSERT INTO xyoff VALUES(3,3,722,233); " & _
            "INSERT INTO xyoff VALUES(3,3,815,421); " & _
            "INSERT INTO xyoff VALUES(3,3,401,267); " & _
            "INSERT INTO xyoff VALUES(3,3,451,650); " & _
            "INSERT INTO xyoff VALUES(3,3,329,485); " & _
            "INSERT INTO xyoff VALUES(3,3,878,370); " & _
            "INSERT INTO xyoff VALUES(3,3,162,616); " & _
            "INSERT INTO xyoff VALUES(3,3,844,183); " & _
            "INSERT INTO xyoff VALUES(3,3,161,216); " & _
            "INSERT INTO xyoff VALUES(3,3,176,676); " & _
            "INSERT INTO xyoff VALUES(3,4,780,128); " & _
            "INSERT INTO xyoff VALUES(3,4,566,121); " & _
            "INSERT INTO xyoff VALUES(3,4,646,120); " & _
            "INSERT INTO xyoff VALUES(3,4,223,557); " & _
            "INSERT INTO xyoff VALUES(3,4,251,117); " & _
            "INSERT INTO xyoff VALUES(3,4,139,209); " & _
            "INSERT INTO xyoff VALUES(3,4,813,597); " & _
            "INSERT INTO xyoff VALUES(3,4,454,538); " & _
            "INSERT INTO xyoff VALUES(3,4,616,198); " & _
            "INSERT INTO xyoff VALUES(3,4,210,159); " & _
            "INSERT INTO xyoff VALUES(4,1,208,415); " & _
            "INSERT INTO xyoff VALUES(4,1,326,665); " & _
            "INSERT INTO xyoff VALUES(4,1,612,133); " & _
            "INSERT INTO xyoff VALUES(4,1,537,513); " & _
            "INSERT INTO xyoff VALUES(4,1,638,438); " & _
            "INSERT INTO xyoff VALUES(4,1,808,269); " & _
            "INSERT INTO xyoff VALUES(4,1,552,121); " & _
            "INSERT INTO xyoff VALUES(4,1,100,189); " & _
            "INSERT INTO xyoff VALUES(4,1,643,664); " & _
            "INSERT INTO xyoff VALUES(4,1,726,378); " & _
            "INSERT INTO xyoff VALUES(4,2,478,409); " & _
            "INSERT INTO xyoff VALUES(4,2,497,507); " & _
            "INSERT INTO xyoff VALUES(4,2,233,148); " & _
            "INSERT INTO xyoff VALUES(4,2,587,237); " & _
            "INSERT INTO xyoff VALUES(4,2,604,166); " & _
            "INSERT INTO xyoff VALUES(4,2,165,455); " & _
            "INSERT INTO xyoff VALUES(4,2,320,258); " & _
            "INSERT INTO xyoff VALUES(4,2,353,496); " & _
            "INSERT INTO xyoff VALUES(4,2,347,495); " & _
            "INSERT INTO xyoff VALUES(4,2,166,622); " & _
            "INSERT INTO xyoff VALUES(4,3,461,332); " & _
            "INSERT INTO xyoff VALUES(4,3,685,278); " & _
            "INSERT INTO xyoff VALUES(4,3,427,594); " & _
            "INSERT INTO xyoff VALUES(4,3,467,346); " & _
            "INSERT INTO xyoff VALUES(4,3,125,548); " & _
            "INSERT INTO xyoff VALUES(4,3,597,680); " & _
            "INSERT INTO xyoff VALUES(4,3,820,445); " & _
            "INSERT INTO xyoff VALUES(4,3,144,330); " & _
            "INSERT INTO xyoff VALUES(4,3,557,434); " & _
            "INSERT INTO xyoff VALUES(4,3,254,315); " & _
            "INSERT INTO xyoff VALUES(4,4,157,339); " & _
            "INSERT INTO xyoff VALUES(4,4,249,220); " & _
            "INSERT INTO xyoff VALUES(4,4,391,323); " & _
            "INSERT INTO xyoff VALUES(4,4,589,429); " & _
            "INSERT INTO xyoff VALUES(4,4,859,592); " & _
            "INSERT INTO xyoff VALUES(4,4,337,680); " & _
            "INSERT INTO xyoff VALUES(4,4,410,288); " & _
            "INSERT INTO xyoff VALUES(4,4,636,596); " & _
            "INSERT INTO xyoff VALUES(4,4,734,433); " & _
            "INSERT INTO xyoff VALUES(4,4,559,549); " & _
            "INSERT INTO xyoff VALUES(5,1,549,607); " & _
            "INSERT INTO xyoff VALUES(5,1,584,498); " & _
            "INSERT INTO xyoff VALUES(5,1,699,116); " & _
            "INSERT INTO xyoff VALUES(5,1,525,524); " & _
            "INSERT INTO xyoff VALUES(5,1,304,667); " & _
            "INSERT INTO xyoff VALUES(5,1,302,232); " & _
            "INSERT INTO xyoff VALUES(5,1,403,149); " & _
            "INSERT INTO xyoff VALUES(5,1,824,403); " & _
            "INSERT INTO xyoff VALUES(5,1,697,203); " & _
            "INSERT INTO xyoff VALUES(5,1,293,689); " & _
            "INSERT INTO xyoff VALUES(5,2,199,275); " & _
            "INSERT INTO xyoff VALUES(5,2,395,393); " & _
            "INSERT INTO xyoff VALUES(5,2,657,642); " & _
            "INSERT INTO xyoff VALUES(5,2,200,655); " & _
            "INSERT INTO xyoff VALUES(5,2,882,234); " & _
            "INSERT INTO xyoff VALUES(5,2,483,565); " & _
            "INSERT INTO xyoff VALUES(5,2,755,640); " & _
            "INSERT INTO xyoff VALUES(5,2,810,305); " & _
            "INSERT INTO xyoff VALUES(5,2,731,655); " & _
            "INSERT INTO xyoff VALUES(5,2,466,690); " & _
            "INSERT INTO xyoff VALUES(5,3,563,584); " & _
            "INSERT INTO xyoff VALUES(5,3,491,117); " & _
            "INSERT INTO xyoff VALUES(5,3,779,292); " & _
            "INSERT INTO xyoff VALUES(5,3,375,637); " & _
            "INSERT INTO xyoff VALUES(5,3,253,553); " & _
            "INSERT INTO xyoff VALUES(5,3,797,514); " & _
            "INSERT INTO xyoff VALUES(5,3,229,480); " & _
            "INSERT INTO xyoff VALUES(5,3,257,194); " & _
            "INSERT INTO xyoff VALUES(5,3,449,555); " & _
            "INSERT INTO xyoff VALUES(5,3,849,630); " & _
            "INSERT INTO xyoff VALUES(5,4,329,286); " & _
            "INSERT INTO xyoff VALUES(5,4,640,197); " & _
            "INSERT INTO xyoff VALUES(5,4,104,150); " & _
            "INSERT INTO xyoff VALUES(5,4,438,272); " & _
            "INSERT INTO xyoff VALUES(5,4,773,226); " & _
            "INSERT INTO xyoff VALUES(5,4,441,650); " & _
            "INSERT INTO xyoff VALUES(5,4,242,340); " & _
            "INSERT INTO xyoff VALUES(5,4,301,435); " & _
            "INSERT INTO xyoff VALUES(5,4,171,397); " & _
            "INSERT INTO xyoff VALUES(5,4,541,619); " & _
            "INSERT INTO xyoff VALUES(6,1,651,301); " & _
            "INSERT INTO xyoff VALUES(6,1,637,137); " & _
            "INSERT INTO xyoff VALUES(6,1,765,643); " & _
            "INSERT INTO xyoff VALUES(6,1,173,296); " & _
            "INSERT INTO xyoff VALUES(6,1,263,192); " & _
            "INSERT INTO xyoff VALUES(6,1,791,302); " & _
            "INSERT INTO xyoff VALUES(6,1,860,601); " & _
            "INSERT INTO xyoff VALUES(6,1,780,445); " & _
            "INSERT INTO xyoff VALUES(6,1,462,214); " & _
            "INSERT INTO xyoff VALUES(6,1,802,207); " & _
            "INSERT INTO xyoff VALUES(6,2,811,685); " & _
            "INSERT INTO xyoff VALUES(6,2,533,531); " & _
            "INSERT INTO xyoff VALUES(6,2,390,614); " & _
            "INSERT INTO xyoff VALUES(6,2,260,580); " & _
            "INSERT INTO xyoff VALUES(6,2,116,377); " & _
            "INSERT INTO xyoff VALUES(6,2,860,458); " & _
            "INSERT INTO xyoff VALUES(6,2,438,590); " & _
            "INSERT INTO xyoff VALUES(6,2,604,562); " & _
            "INSERT INTO xyoff VALUES(6,2,241,242); " & _
            "INSERT INTO xyoff VALUES(6,2,667,298); " & _
            "INSERT INTO xyoff VALUES(6,3,787,698); " & _
            "INSERT INTO xyoff VALUES(6,3,868,521); " & _
            "INSERT INTO xyoff VALUES(6,3,412,587); " & _
            "INSERT INTO xyoff VALUES(6,3,640,131); " & _
            "INSERT INTO xyoff VALUES(6,3,748,410); " & _
            "INSERT INTO xyoff VALUES(6,3,257,244); " & _
            "INSERT INTO xyoff VALUES(6,3,411,195); " & _
            "INSERT INTO xyoff VALUES(6,3,464,356); " & _
            "INSERT INTO xyoff VALUES(6,3,157,339); " & _
            "INSERT INTO xyoff VALUES(6,3,434,505); " & _
            "INSERT INTO xyoff VALUES(6,4,480,671); " & _
            "INSERT INTO xyoff VALUES(6,4,519,228); " & _
            "INSERT INTO xyoff VALUES(6,4,404,513); " & _
            "INSERT INTO xyoff VALUES(6,4,120,538); " & _
            "INSERT INTO xyoff VALUES(6,4,403,663); " & _
            "INSERT INTO xyoff VALUES(6,4,477,677); " & _
            "INSERT INTO xyoff VALUES(6,4,690,154); " & _
            "INSERT INTO xyoff VALUES(6,4,606,498); " & _
            "INSERT INTO xyoff VALUES(6,4,430,665); " & _
            "INSERT INTO xyoff VALUES(6,4,499,273); " & _
            "INSERT INTO xyoff VALUES(7,1,118,526); " & _
            "INSERT INTO xyoff VALUES(7,1,817,522); " & _
            "INSERT INTO xyoff VALUES(7,1,388,638); " & _
            "INSERT INTO xyoff VALUES(7,1,181,265); " & _
            "INSERT INTO xyoff VALUES(7,1,442,332); " & _
            "INSERT INTO xyoff VALUES(7,1,475,282); " & _
            "INSERT INTO xyoff VALUES(7,1,722,633); " & _
            "INSERT INTO xyoff VALUES(7,1,104,394); " & _
            "INSERT INTO xyoff VALUES(7,1,631,262); " & _
            "INSERT INTO xyoff VALUES(7,1,372,392); " & _
            "INSERT INTO xyoff VALUES(7,2,600,413); " & _
            "INSERT INTO xyoff VALUES(7,2,386,223); " & _
            "INSERT INTO xyoff VALUES(7,2,839,174); " & _
            "INSERT INTO xyoff VALUES(7,2,293,410); " & _
            "INSERT INTO xyoff VALUES(7,2,281,391); " & _
            "INSERT INTO xyoff VALUES(7,2,859,387); " & _
            "INSERT INTO xyoff VALUES(7,2,478,347); " & _
            "INSERT INTO xyoff VALUES(7,2,646,690); " & _
            "INSERT INTO xyoff VALUES(7,2,713,234); " & _
            "INSERT INTO xyoff VALUES(7,2,199,588); " & _
            "INSERT INTO xyoff VALUES(7,3,389,256); " & _
            "INSERT INTO xyoff VALUES(7,3,349,542); " & _
            "INSERT INTO xyoff VALUES(7,3,363,345); " & _
            "INSERT INTO xyoff VALUES(7,3,751,302); " & _
            "INSERT INTO xyoff VALUES(7,3,423,386); " & _
            "INSERT INTO xyoff VALUES(7,3,267,444); " & _
            "INSERT INTO xyoff VALUES(7,3,243,182); " & _
            "INSERT INTO xyoff VALUES(7,3,453,658); " & _
            "INSERT INTO xyoff VALUES(7,3,126,345); " & _
            "INSERT INTO xyoff VALUES(7,3,120,472); " & _
            "INSERT INTO xyoff VALUES(7,4,359,654); " & _
            "INSERT INTO xyoff VALUES(7,4,339,516); " & _
            "INSERT INTO xyoff VALUES(7,4,710,452); " & _
            "INSERT INTO xyoff VALUES(7,4,810,560); " & _
            "INSERT INTO xyoff VALUES(7,4,644,692); " & _
            "INSERT INTO xyoff VALUES(7,4,826,327); " & _
            "INSERT INTO xyoff VALUES(7,4,465,462); " & _
            "INSERT INTO xyoff VALUES(7,4,310,456); " & _
            "INSERT INTO xyoff VALUES(7,4,577,613); " & _
            "INSERT INTO xyoff VALUES(7,4,502,555); " & _
            "INSERT INTO xyoff VALUES(8,1,601,620); " & _
            "INSERT INTO xyoff VALUES(8,1,372,683); " & _
            "INSERT INTO xyoff VALUES(8,1,758,399); " & _
            "INSERT INTO xyoff VALUES(8,1,485,552); " & _
            "INSERT INTO xyoff VALUES(8,1,159,563); " & _
            "INSERT INTO xyoff VALUES(8,1,536,303); " & _
            "INSERT INTO xyoff VALUES(8,1,122,263); " & _
            "INSERT INTO xyoff VALUES(8,1,836,435); " & _
            "INSERT INTO xyoff VALUES(8,1,544,146); " & _
            "INSERT INTO xyoff VALUES(8,1,270,277); " & _
            "INSERT INTO xyoff VALUES(8,2,849,281); " & _
            "INSERT INTO xyoff VALUES(8,2,563,242); " & _
            "INSERT INTO xyoff VALUES(8,2,704,463); " & _
            "INSERT INTO xyoff VALUES(8,2,102,165); " & _
            "INSERT INTO xyoff VALUES(8,2,797,524); " & _
            "INSERT INTO xyoff VALUES(8,2,612,426); " & _
            "INSERT INTO xyoff VALUES(8,2,345,372); " & _
            "INSERT INTO xyoff VALUES(8,2,820,376); " & _
            "INSERT INTO xyoff VALUES(8,2,789,156); " & _
            "INSERT INTO xyoff VALUES(8,2,321,466); " & _
            "INSERT INTO xyoff VALUES(8,3,150,332); " & _
            "INSERT INTO xyoff VALUES(8,3,136,152); " & _
            "INSERT INTO xyoff VALUES(8,3,468,528); " & _
            "INSERT INTO xyoff VALUES(8,3,409,192); " & _
            "INSERT INTO xyoff VALUES(8,3,820,216); " & _
            "INSERT INTO xyoff VALUES(8,3,847,249); " & _
            "INSERT INTO xyoff VALUES(8,3,801,267); " & _
            "INSERT INTO xyoff VALUES(8,3,181,670); " & _
            "INSERT INTO xyoff VALUES(8,3,398,563); " & _
            "INSERT INTO xyoff VALUES(8,3,439,576); " & _
            "INSERT INTO xyoff VALUES(8,4,123,309); " & _
            "INSERT INTO xyoff VALUES(8,4,190,496); " & _
            "INSERT INTO xyoff VALUES(8,4,571,531); " & _
            "INSERT INTO xyoff VALUES(8,4,290,255); " & _
            "INSERT INTO xyoff VALUES(8,4,244,412); " & _
            "INSERT INTO xyoff VALUES(8,4,264,596); " & _
            "INSERT INTO xyoff VALUES(8,4,253,420); " & _
            "INSERT INTO xyoff VALUES(8,4,847,536); " & _
            "INSERT INTO xyoff VALUES(8,4,120,288); " & _
            "INSERT INTO xyoff VALUES(8,4,331,639); " & _
            "/* Create the geopoly object from test data above */ " & _
            "DROP TABLE IF EXISTS geo1;" & _
            "CREATE VIRTUAL TABLE geo1 USING geopoly(type,clr); " & _
            "INSERT INTO geo1(_shape,type,clr) " & _
            "  SELECT geopoly_xform(jshape,A,B,-B,A,xoff,yoff), basis.name, xform.clr " & _
            "    FROM basis, xform, xyoff " & _
            "   WHERE xyoff.id1=basis.rowid AND xyoff.id2=xform.rowid; "
        objConn.execute sSql
        
        ' Query polygon
        sSql = _
            "CREATE TEMP TABLE querypoly(poly JSON, clr TEXT); " & _
            "INSERT INTO querypoly(clr, poly) VALUES " & _
            "   ('orange', '[[300,300],[400,350],[500,250],[480,500],[400,480],[300,550],[280,450],[320,400],[280,350],[300,300]]');"
        objConn.execute sSql

        log "overlap query"
        log query("SELECT * FROM geo1, querypoly WHERE geopoly_overlap(_shape, poly);")
        
        log "within query"
        log query("SELECT * FROM geo1, querypoly WHERE geopoly_within(_shape, poly);")
        
        log query("SELECT geopoly_blob(geopoly_regular(100,100,40,3)) as p;")
        log query("SELECT geopoly_json(geopoly_regular(100,100,40,3)) as p;")
        log query("SELECT geopoly_svg(geopoly_regular(100,100,40,3)) as p;")
        log query("SELECT geopoly_area(geopoly_regular(100,100,40,3)) as a;")
        
        closedb
    end function

    '********************************************
    public function helper_getStats(byref objConn)
        dim oRsStmt: set oRsStmt = objConn.execute("SELECT * FROM sqlite_stmt where busy!=1;")
        if oRsStmt.state = 1 then
            if (not oRsStmt.BOF) and (not oRsStmt.EOF)  then
                log "query: SELECT * FROM sqlite_stmt where busy!=1;"
                oRsStmt.MoveFirst
                do while oRsStmt.EOF = false
                    dim ff: for each ff in oRsStmt.Fields
                        log ff.Name & "(" & ff.type & "):" & ff.value
                    next
                    oRsStmt.MoveNext
                loop
                log ""
            else
                log "nothing to report"
            end if
            oRsStmt.close
        end if
    end function
    
    '********************************************
    public function helper_getResult(byref objCmd, bOutput)
        if bOutput then log ""
        dim ra
        dim oRS: set oRS = objCmd.execute(ra,,adCmdText)
        log "helper_getResult ra " & ra & " <-- " & objCmd.CommandText
        if oRs.state = 1 then
            if (not oRs.BOF) and (not oRs.EOF)  then
                oRs.MoveFirst
                do while oRs.EOF = false
                    dim ff: for each ff in oRs.Fields
                        if bOutput then log ff.Name & "(" & ff.type & "):" & ff.value
                    next
                    oRs.MoveNext
                loop
            end if
            oRs.close
        end if
        set oRs = nothing
        if bOutput then log ""
    end function
    
    '********************************************
    public function sqlite3_page_size_tests
        log "****************************************************************************"
        log "sqlite3_page_size_tests"
        opendb "SQL3 "
        dim sSql
        log query("PRAGMA page_size=8192")
        log query("VACUUM")
        log query("PRAGMA page_size=4096")
        log query("VACUUM")
        closedb
    end function

    '********************************************
    public function sqlite3_rtree_tests
        log "****************************************************************************"
        log "sqlite3_rtree_tests"
        opendb "SQL3 "
        
        dim sSql
        
        log vbcrlf & "sqlite3_rtree_tests"
        log query("PRAGMA compile_options;")
        
        query("drop table if exists demo_index")
        
        sSql = "CREATE VIRTUAL TABLE demo_index USING rtree( " & _
            "id, minX, maxX, minY, maxY " & _
        "); "
        query(sSql)
        
        sSql = _
            "INSERT INTO demo_index VALUES( " & _
            "    1," & _
            "    -80.7749, -80.7747, " & _
            "    35.3776, 35.3778" & _
            ");"
        query(sSql)
            
        sSql = _
            "INSERT INTO demo_index VALUES( " & _
            "    2," & _
            "    -81.0, -79.6," & _
            "    35.0, 36.2" & _
            "); "
        query(sSql)
        
        log query("SELECT * FROM demo_index WHERE id=1;")
        
        log query("SELECT id FROM demo_index WHERE minX>=-81.08 AND maxX<=-80.58 AND minY>=35.00  AND maxY<=35.44;")

        log query("SELECT id FROM demo_index WHERE maxX>=-81.08 AND minX<=-80.58 AND maxY>=35.00  AND minY<=35.44;")

        log query("SELECT id FROM demo_index WHERE maxY>=35.0  AND minY<=35.0;")

        query("drop table if exists demo_index2")

        sSql = _
            "CREATE VIRTUAL TABLE demo_index2 USING rtree(" & _
            "   id," & _
            "   minX, maxX," & _
            "   minY, maxY," & _
            "   +objname TEXT," & _
            "   +objtype TEXT," & _
            "   +boundary BLOB" & _
            ");"
        query(sSql)

        sSql = _
            "INSERT INTO demo_index2 VALUES( " & _
            "    1," & _
            "    -80.7749, -80.7747, " & _
            "    35.3776, 35.3778" & _
            ");"
        query(sSql)
            
        sSql = _
            "INSERT INTO demo_index2 VALUES( " & _
            "    2," & _
            "    -81.0, -79.6," & _
            "    35.0, 36.2" & _
            "); "
        query(sSql)

        sSql = "SELECT rtreecheck('demo_index');"
        log query(sSql)
        
        sSql = "SELECT rtreecheck('demo_index2');"
        log query(sSql)
        
        if objRS.state = 1 then objRS.close
        closedb
    end function

    '********************************************
    public function sqlite3_fts_tests
        log "****************************************************************************"
        log "sqlite3_fts_tests"
        opendb "SQL3 "
        
        ' FTS3 and FTS4 are nearly identical. They share most of their code in common, and their interfaces are the same.
        log vbcrlf & "fts4"
        query("DROP TABLE IF EXISTS pages;")
        query("CREATE VIRTUAL TABLE pages USING fts4(title, body);")
        
        log vbcrlf & "insert"
        query("INSERT INTO pages(docid, title, body) VALUES(53, 'Home Page', 'SQLite is a software...');")
        query("INSERT INTO pages(title, body) VALUES('Download', 'All SQLite source code...');")
        query("INSERT INTO pages(title, body) VALUES('Upload', 'Upload SQLite src code...');")
        
        log vbcrlf & "update"
        query("UPDATE pages SET title = 'Download SQLite' WHERE rowid = 54;")
        
        log vbcrlf & "optimize"
        query("INSERT INTO pages(pages) VALUES('optimize');")
        
        log vbcrlf & "query"
        logResult query2csv("SELECT * FROM pages WHERE pages MATCH 'sqlite';")
        logResult query2csv("SELECT * FROM pages WHERE pages MATCH 's* code';")


        log vbcrlf & "fts5"
        log query("CREATE VIRTUAL TABLE posts USING FTS5(title, body);")
        log query(_
            "INSERT INTO posts(title,body) " & _
            "VALUES('Learn SQlite FTS5','This tutorial teaches you how to perform full-text search using FTS5'), " & _
            "('Advanced SQlite Full-text Search','Show you some advanced techniques in SQLite full-text searching'), " & _
            "('SQLite SQLite SQLite SQLite ','SQLite SQLite SQLite SQLite SQLite SQLite SQLite SQLite SQLite SQLite '), " & _
            "('SQLite Tutorial','Help you learn SQLite quickly and use SQLite effectively'); ")
            
        log vbcrlf & "query"
        logResult query2csv("SELECT * FROM posts;")
        
        log vbcrlf & "full text search"
        logResult query2csv("SELECT * FROM posts WHERE posts MATCH 'fts5';")
        logResult query2csv("SELECT * FROM posts WHERE posts = 'fts5';")
        logResult query2csv("SELECT * FROM posts('fts5');")
        logResult query2csv("SELECT * FROM posts WHERE posts MATCH 'text' ORDER BY rank;")
        logResult query2csv("SELECT * FROM posts WHERE posts MATCH 'learn SQLite';")
        logResult query2csv("SELECT * FROM posts WHERE posts = 'search*';")
        logResult query2csv("SELECT * FROM posts WHERE posts MATCH 'learn NOT text';")
        logResult query2csv("SELECT * FROM posts WHERE posts MATCH 'learn OR text';")
        logResult query2csv("SELECT * FROM posts WHERE posts MATCH 'sqlite AND searching';")
        logResult query2csv("SELECT * FROM posts WHERE posts MATCH 'search AND sqlite OR help';")
        logResult query2csv("SELECT * FROM posts WHERE posts MATCH 'search AND (sqlite OR help)';")
        
        log vbcrlf & "highlight"
        logResult query2csv("SELECT  " & _
            "highlight(posts,0, '<b>', '</b>') title,  " & _
            "highlight(posts,1, '<b>', '</b>') body  " & _
            "FROM posts WHERE posts MATCH 'SQLite' " & _
            "ORDER BY rank;")
            
        log vbcrlf & "bm25()/rank"
        logResult query2csv("SELECT bm25(posts) as accuracy, * FROM posts WHERE posts MATCH 'sqlite' ORDER BY bm25(posts);")
        
        ' must have MATCH for rank to be non-null
        logResult query2csv("SELECT rank, * FROM posts WHERE posts MATCH 'sqlite' ORDER BY rank;")

        log vbcrlf & "snippet"
        logResult query2csv("SELECT  " & _
            "snippet(posts,0,'<<','>>','...',3) title,  " & _
            "snippet(posts,1,'<<','>>','...',6) body  " & _
            "FROM posts WHERE posts MATCH 'sqLite' ;")

        if objRS.state = 1 then objRS.close
        closedb
    end function

    '********************************************
    ' test the built in math functions
    public function sqlite3_BuiltIn_Tests
        log "****************************************************************************"
        log "sqlite3_BuiltIn_Tests"
        opendb "SQL3 "
        log "some of the functions are not available, log(x) is base 10"
        sqlite_math_functions
        closedb
    end function

    '********************************************
    public function sqlite_math_functions
        dim retValue: retValue = 0
        log "****************************************************************************"
        log "sqlite_math_functions"
        query2csv("select acos(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' acosh(X)
        query2csv("select acosh(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' asin(X)
        query2csv("select asin(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' asinh(X)
        query2csv("select asinh(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' atan(X)
        query2csv("select atan(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' atanh(X)
        query2csv("select atanh(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' atan2(X,Y)
        query2csv("select atan2(4,5) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1

        ' atn2(X,Y) FAILS function not available
        query2csv("select atn2(4,5) as x;")
        if instr(aQueryResults(3), "no such function") <= 0 then retValue = retValue + 1
        
        ' ceil(X)
        query2csv("select ceil(1.1) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' ceiling(X)
        query2csv("select ceiling(1.1) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' cos(X)
        query2csv("select cos(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' cosh(X)
        query2csv("select cosh(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1

        ' cot(X) FAILS function not available
        query2csv("select cot(30*3.14159/180) as x;")
        if instr(aQueryResults(3), "no such function") <= 0 then retValue = retValue + 1

        ' coth(X) FAILS function not available
        query2csv("select coth(30*3.14159/180) as x;")
        if instr(aQueryResults(3), "no such function") <= 0 then retValue = retValue + 1

        ' degrees(X)
        query2csv("select degrees(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' exp(X)
        query2csv("select exp(2) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' floor(X)
        query2csv("select floor(1.1) as w, floor(1.9) as x, floor(-1.1) as y, floor(-1.9) as z;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' ln(X)
        query2csv("select ln(10) as x;")
        if aQueryResults(2)(0) <> "2.30258509299405" then retValue = retValue + 1

        ' log(B,X)
        query2csv("select log(2,10) as x,log(10,10) as y,log(16,10) as z;")
        if aQueryResults(2)(0) <> "3.32192809488736,1,0.830482023721841" then retValue = retValue + 1

        ' log(X)
        query2csv("select log(10) as x;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue + 1

        ' log10(X)
        query2csv("select log10(10) as x;")
        if aQueryResults(2)(0) <> 1 then retValue = retValue + 1

        ' log2(X)
        query2csv("select log2(10) as x;")
        if aQueryResults(2)(0) <> "3.32192809488736" then retValue = retValue + 1

        ' mod(X,Y)
        query2csv("select mod(10,10) as x, mod(10,11) as y;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' pi()
        query2csv("select pi() as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' pow(X,Y)
        query2csv("select pow(2,2) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' power(X,Y)
        query2csv("select power(2,2) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' radians(X)
        query2csv("select radians(180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' sign(X)
        query2csv("select sign(-10) as x, sign(10) as y;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' sin(X)
        query2csv("select sin(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' sinh(X)
        query2csv("select sinh(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' sqrt(X)
        query2csv("select sqrt(4) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1

        ' square(X) FAILS function not available
        query2csv("select square(4) as x;")
        if instr(aQueryResults(3), "no such function") <= 0 then retValue = retValue + 1


        ' tan(X)
        query2csv("select tan(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1
        ' tanh(X)
        query2csv("select tanh(30*3.14159/180) as x;")
        if len(aQueryResults(3)) <> 0 then retValue = retValue + 1

        ' trunc(X)
        query2csv("select trunc(1.123) as x, trunc(1.9) as y;")
        if aQueryResults(2)(0) <> "1,1" then retValue = retValue + 1

        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function sqlite3_feature_tests
        log "****************************************************************************"
        log "sqlite3_feature_tests"
        log "test some other more recent features of sqlite or SQL used by sqlite"

        opendb "MEM  "
        dim retValue: retValue = 0
        dim result
        
        log "iif() is now included in SQLite SQL language"
        result = query2csv("select iif(1=2,'true','false') as w, iif(2=2,'true','false') as x, iif('hello' = 'world','true','false') as y, iif('same' = 'same','true','false') as z;")
        if aQueryResults(2)(0) <> """false"",""true"",""false"",""true""" then retValue = retValue + 1
        
        log "'alter table drop column' is now included in SQLite SQL language"
        result = query2csv("create table people (id INTEGER, income REAL, tax_rate REAL);")
        if result = -1 then retValue = retValue + 1
        dim q: q = _
            "WITH RECURSIVE person(x) AS ( " & _
            "SELECT 1 UNION ALL SELECT x+1 FROM person LIMIT 1000 " & _
            ") " & _
            "INSERT INTO people ( id, income, tax_rate) " & _
            "SELECT x, 70+mod(x,15)*3, (15.0+(mod(x,5)*0.2)+mod(x,15))/100. FROM person;"
        result = query2csv(q)
        if result = -1 then retValue = retValue + 1
        result = query2csv("select count(1) from people")
        if result <> 1 then retValue = retValue + 1
        
        result = query2csv("drop table if exists people_copy")
        if result <> 0 then retValue = retValue + 1
        
        result = query2csv("create table people_copy as select * from people")
        if result <> 0 then retValue = retValue + 1
        
        result = query2csv("select count(1) from people_copy")
        if result <> 1 then retValue = retValue + 1
        
        result = query2csv("alter table people_copy drop column tax_rate")
        if result <> 0 then retValue = retValue + 1
        
        result = query2csv("select * from people_copy limit 1")
        if result <> 1 then retValue = retValue + 1
        
        result = query2csv("drop table if exists people_copy")
        if result = -1 then retValue = retValue + 1

        closedb
        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    public function getTableInfoSimple
        log "****************************************************************************"
        log "getTableInfoSimple"
        opendb "SQL3 "
        
        log "want to get table info for table test_table, and in particular the column names"
        log "PRAGMA table_info() gets everything but you can't control what it gives you..." & vbcrlf
        log query("select * from test_table limit 3") & vbcrlf
        
        log "since the output of query is well known we can parse it to get long view of result..."  & vbcrlf
        dim sResult: sResult = query("PRAGMA table_info(test_table)")
        log sResult & vbcrlf
        log helper_longTableFormat(sResult) & vbcrlf
        log ""
        log "pragma_table_info lets you use a select query to pick the fields you want" & vbcrlf
        log query("select * from pragma_table_info('test_table') where name like '%MyField%'") & vbcrlf
        
        closedb
    end function

    '********************************************
    public function getTableInfoDetails
        log "****************************************************************************"
        log "getTableInfoDetails"
        
        log "List all columns in a SQLite database"
        log "based on https://til.simonwillison.net/sqlite/list-all-columns-in-a-database"
        
        opendb "SQL3 "
        
        log ""
        log "first recipe - contains all the details provided by pragma_table_info()"
        log ""
        log query( _
            "select " & _
            "  sqlite_master.name as table_name, " & _
            "  table_info.* " & _
            "from " & _
            "  sqlite_master " & _
            "  join pragma_table_info(sqlite_master.name) as table_info " & _
            "order by " & _
            "  sqlite_master.name " _
        )
        log ""        
        log "second recipe - simple...just table name and column name, ignores any ""sqlite_"" tables"
        log ""
        log query( _
            "SELECT " & _
            "  m.name as tableName, " & _
            "  p.name as columnName " & _
            "FROM " & _
            "  sqlite_master m " & _
            "  LEFT OUTER JOIN pragma_table_info((m.name)) p ON m.name <> p.name " & _
            "WHERE " & _
            "  m.type IN ('table', 'view') " & _
            "  AND m.name NOT LIKE 'sqlite_%' " & _
            "ORDER BY " & _
            "  tableName, " & _
            "  columnName " _
        )
        log ""
        
        closedb
        
    end function

    '********************************************
    public function helper_longTableFormat(s)
        dim sReturn
        dim aResult: aResult = split(s, vbcrlf)
        ' i = 0 is the SQL query string
        sReturn = aResult(0) & vbcrlf
        dim i: for i = 1 to ubound(aResult)-1
            sReturn = sReturn & "column " & i & vbcrlf
            dim sRow: for each sRow in split(aResult(i),",")
                sReturn = sReturn &  "    " & sRow & vbcrlf
            next
        next
        helper_longTableFormat = sReturn
    end function

    '********************************************
    function dumpPragma
        log "****************************************************************************"
        log "dumpPragma"
        opendb "MEM  "
        dim ss: ss = ""
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA analysis_limit")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA application_id")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA auto_vacuum")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA automatic_index")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA busy_timeout")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA cache_size")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA cache_spill")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA case_sensitive_like")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA cell_size_check")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA checkpoint_fullfsync")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA collation_list")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA compile_options")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA data_version")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA database_list")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA encoding")
        
        ' sorts the list so that comparing is easier
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & getSortedFunctionList("") & vbcrlf
        
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA integrity_check")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA hard_heap_limit")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA journal_mode")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA journal_size_limit")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA locking_mode")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA max_page_count")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA mmap_size")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA optimize")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA page_size")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA page_count")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA pragma_list")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA quick_check")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA shrink_memory")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA soft_heap_limit")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA synchronous")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA table_info('test_table')")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA table_xinfo('test_table')")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA temp_store")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA temp_store_directory")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA threads")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA trusted_schema")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA user_version")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA wal_autocheckpoint")
        ss = ss & "****************************************************************************" & vbcrlf
        ss = ss & query("PRAGMA wal_checkpoint")
        closedb
        dumpPragma = ss
    end function

    '********************************************
    public function getSortedFunctionList(f)
        ' assumes SQLite DB is already open
        dim sReturn
        Dim oRS : Set oRS = CreateObject("ADODB.Recordset")
        oRS.Fields.Append "name", 200, 50   'adVarChar
        oRS.Fields.Append "narg", 3         'adInteger
        oRS.Open
        dim s: s = query("PRAGMA function_list")
        dim rows: rows = split(s,vbcrlf)
        sReturn = rows(0) & vbcrlf
        dim r: for r = 1 to ubound(rows)-1
            dim cols: cols = split(rows(r),",")
            dim name: name = split(cols(0),":")(1)
            dim narg: narg = split(cols(4),":")(1)
            if len(f) > 0 then
                if instr( name, f) > 0 then 
                    oRS.AddNew
                    oRS.Fields("name").value = name
                    oRS.Fields("narg").value = narg
                    oRS.UpDate
                end if
            else
                oRS.AddNew
                oRS.Fields("name").value = name
                oRS.Fields("narg").value = narg
                oRS.UpDate
            end if
        next
        if oRS.recordcount > 0 then
            oRS.Sort = "name, narg"
            Dim aTable : aTable = oRS.GetRows()
            sReturn = sReturn &  oRs.Fields(0).name & "," & oRs.Fields(1).name  & vbcrlf
            Dim i: For i = 0 To UBound(aTable, 2)
                sReturn = sReturn & aTable(0, i) & "," & aTable(1, i) & vbcrlf
            Next
        else
            sReturn = sReturn & " no items found for """ & f &""""
        end if
        oRs.close
        set oRs = nothing
        getSortedFunctionList = sReturn
    end function

    '********************************************
    ' verify that various options all work for insert and select
    public function insertTests()
        
        REM mmap_size = SQLite has the option of accessing disk content directly using memory-mapped I/O 
        REM journal_mode = DELETE | TRUNCATE | PERSIST | MEMORY | WAL | OFF
        REM page_size = Query or set the page size of the database. The page size must be a power of two between 512 and 65536 inclusive.
        REM synchronous = 0 | OFF | 1 | NORMAL | 2 | FULL | 3 | EXTRA;
        REM temp_store = 0 | DEFAULT | 1 | FILE | 2 | MEMORY;
        REM locking_mode = NORMAL | EXCLUSIVE

        dim mmap_size:      mmap_size =         Array(0,1073741824,1073741824*2)                                ' off, 1 GB, 2 GB
        dim cache_size:     cache_size =        Array(-2000,-4000)                                              ' -2000 (default)
        dim journal_mode:   journal_mode =      Array("WAL","MEMORY","OFF","DELETE","TRUNCATE","PERSIST")       ' OFF, WAL (DELETE default)
        dim page_size:      page_size =         Array(4096,4096*4)                                              ' 4096 (default)
        dim synchronous:    synchronous =       Array("OFF","NORMAL","FULL","EXTRA")                            ' NORMAL (default)
        dim temp_store:     temp_store =        Array("MEMORY","FILE")                                          ' FILE
        dim locking_mode:   locking_mode =      Array("EXCLUSIVE","NORMAL")                                     ' NORMAL (default)
        
        REM dim mmap_size:      mmap_size =         Array(0)
        REM dim cache_size:     cache_size =        Array(-2000)
        REM dim journal_mode:   journal_mode =      Array("TRUNCATE")
        REM dim page_size:      page_size =         Array(4096*2)
        REM dim synchronous:    synchronous =       Array("NORMAL")
        REM dim temp_store:     temp_store =        Array("FILE")
        REM dim locking_mode:   locking_mode =      Array("NORMAL")
        
        dim apg: apg = Array()
        dim cs,jm,ps,s,ts,lm,mm
        
        for each ts in temp_store
            for each ps in page_size
                for each s in synchronous
                    for each lm in locking_mode
                        for each jm in journal_mode
                            for each cs in cache_size
                                for each mm in mmap_size
                                    redim preserve apg(ubound(apg)+1)
                                    apg(ubound(apg)) = "mmap_size="&mm&"|cache_size="&cs&"|journal_mode="&jm&"|page_size="&ps&"|synchronous="&s&"|temp_store="&ts&"|locking_mode="&lm
                                next
                            next
                        next
                    next
                next
            next
        next
        
        REM fixed rows and columns so transactions not important here
        dim types: types = Array("INTEGER","REAL","TEXT")
        dim aInsertsPerTransaction: aInsertsPerTransaction = Array(1000)
        dim aPrimaryKey: aPrimaryKey = Array(true,false)
        
        dim rtn, rps, max_rps_settings, max_rps
        max_rps_settings = ""
        max_rps = 0
        
        dim count: count = 0
        dim r: r = 100
        dim c: c = 25
        
        log "insertTests"
        
        dim t: for each t in types
            dim ipt: for each ipt in aInsertsPerTransaction
                dim pk: for each pk in aPrimaryKey
                    dim pg: for each pg in apg
                        rtn = test(r,c,t,ipt,pk,pg,true,false)
                        count = count + 1
                        log Right("       " & cstr(count), 7) & "," & rtn
                        rps = split(rtn,",")(0)
                        if rps > max_rps then 
                            max_rps = rps
                            max_rps_settings = rtn
                        end if
                    next
                next
            next
        next
        
        log vbcrlf & vbcrlf & "max_rps for "&r&"x"&c&" is " & max_rps_settings & vbcrlf
        
    end function

    '********************************************
    ' r = number of rows to insert
    ' c = number of columns per row
    ' t = column data type (INTEGER, REAL, TEXT)
    ' ipt = number of records per transaction
    ' pk = use primary key (true/false)
    ' pg = pipe "|" separated pragma string
    ' d = delete db after create (true/false)
    ' s = simulate only (don't create db)
    function test(r,c,t,ipt,pk,pg,d,s)
        dim ss: ss = ""
        dim retValue: retValue = 0
        dim pragmaHeader: pragmaHeader = ""
        dim pragma: pragma = ""
        
        if len(pg) > 0 then
            dim aa: aa = split(pg,"|")
            dim aaa
            for each aaa in aa
                pragmaHeader = pragmaHeader & split(aaa,"=")(0) & ","
                pragma = pragma & split(aaa,"=")(1) & ","
            next
        end if
        
        if bOpenFirstTime then
            log "run,bit,engine,pk,rows,columns,data_type," & pragmaHeader & "AddNewRate_rps,InsertTime_s,TransactionCount,numTranacions"
            bOpenFirstTime = false
        end if

        if s = false then
        
            dbSqlite3 = strFolder & "\testDBs\" & pk & "_" & r & "_" & c & "_" & ipt & "_" & left(t,1) & "_" & replace(replace(pragma,"-","N"),",","_") & iBitness & ".sqlite3"
            if objFSO.FileExists(dbSqlite3) then objFSO.DeleteFile(dbSqlite3)
            opendb "SQL3 "
            for each aaa in aa
                if bVerboseOutput then log "PRAGMA " & aaa & ";"
                objConn.execute "PRAGMA " & aaa & ";"
            next
        
            ' I'm not sure where this limit is coming from but it is there
            if c > 998 then c = 998
            
            dim sPk
            if pk = true then sPk = "pk_"
            
            dim sTableName: sTableName = "test_table"

            dim aHeader
            dim aData
            
            dim createTable: createTable = "CREATE TABLE " & sTableName & " ( "
                
            if pk = false then
                aHeader = Array("id")
                aData = Array(-1)
                createTable = createTable & _
                "    id     INTEGER"
            else
                aHeader = Array("id")
                aData = Array(-1)
                createTable = createTable & _
                "    id     INTEGER PRIMARY KEY"
            end if
            
            if c > 0 then createTable = createTable & ","
            
            dim i: for i = 1 to c
                createTable = createTable & "    myField_" & i & " " & t
                if i < c then createTable = createTable & ","
                redim preserve aHeader(ubound(aHeader)+1)
                aHeader(ubound(aHeader)) = "myField_" & i
                redim preserve aData(ubound(aData)+1)
                select case t
                    case "INTEGER"
                        aData(ubound(aData)) =  clng(i)
                    case "REAL"
                        aData(ubound(aData)) = cdbl(i+0.1)
                    case "TEXT"
                        aData(ubound(aData)) = "myField_" & i
                end select
            next
            
            createTable = createTable & ")"

            dim time: time = 0
            dim rps: rps = 0

            if typename(objConn) = "Empty" then 
                log "connection object is Empty"
                exit function
            end if
            objConn.execute createTable
        
            openRecordSet("select * from [" & sTableName & "]")
                
            dim iRecords: iRecords = r
            dim iTransactionCount: iTransactionCount = 0
            dim sHeader: sHeader = join(aHeader,",")
            dim sData
            dim dStart: dStart = timer
            objConn.BeginTrans
            for i = 1 to iRecords
                aData(0)= clng(i)
                
                ' AddNew is the way to go for quick inserts from scripts
                objRs.AddNew aHeader, aData
                
                ' insert into using objConn does not perform...at all.
                ' sData = join(aData,",")
                ' objConn.execute "insert into [" & sTableName & "] (" & sHeader & ") values(" & sData & ");"
                
                if i mod ipt = 0 then
                    objConn.CommitTrans
                    iTransactionCount = iTransactionCount + 1
                    objConn.BeginTrans
                end if
            next
            objConn.CommitTrans
            iTransactionCount = iTransactionCount + 1
        
            time = (timer - dStart)
            if time = 0 then time = 0.001
            rps = (r/cdbl(time))
                
            REM read back the last inserted record to verify table was inserted
            query2csv("select myField_1 from test_table where id = " & r)
            select case t
                case "INTEGER"
                     if 1 <> clng(aQueryResults(2)(0)) then
                        retValue = retValue + 1
                    end if
               case "REAL"
                    if 1.1 <> round(cdbl(aQueryResults(2)(0)),1) then
                        retValue = retValue + 1
                    end if
                case "TEXT"
                    if """myField_1""" <> aQueryResults(2)(0) then
                        retValue = retValue + 1
                    end if
            end select
            
            closedb

            if d then
                if objFSO.FileExists(dbSqlite3) then objFSO.DeleteFile(dbSqlite3)
                if objFSO.FileExists(dbSqlite3 & "-journal") then objFSO.DeleteFile(dbSqlite3 & "-journal")
            end if
            
        end if
        
        test = rps & "," & time & "," & iBitness & "," & pk & "," & r & "," & c & "," & t & "," & pragma & ipt & "," & iTransactionCount

        if retValue > 0 then err.raise retValue
    end function

    '********************************************
    sub opendb(p)
        Set objConn = CreateObject("ADODB.Connection")
        
        dim sConnStr

        select case p
            case "SQL3 "
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";"
            case "MEM  ":
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=:memory:;"
            case "SQL3-crypto"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\crypto.dll;"
            case "SQL3-LoadExt-Csv"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\csv.dll;"
            case "SQL3-LoadExt-ExtFun"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\extension-functions.dll;"
            case "SQL3-checkfreelist"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\checkfreelist.dll;"
            case "SQL3-ieee754"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\ieee754.dll;"
            case "SQL3-LoadExt-Regexp"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\regexp.dll;"
            case "SQL3-series"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\series.dll;"
            case "SQL3-sha"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\shathree.dll;"
            case "SQL3-totype"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\totype.dll;"
            case "SQL3-wholenumber"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\wholenumber.dll;"
            case "SQL3-uuid"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\uuid.dll;"
            case "SQL3-vfsstat"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\vfsstat.dll;"
            case "SQL3-bfsvtab"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\bfsvtab.dll;"
            case "SQL3-decimal"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\decimal.dll;"
            case "SQL3-fileio"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\fileio.dll;"
            case "SQL3-sqlfcmp"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=.\install\" & sBitPath & "\sqlfcmp.dll;"
        end select

        if bVerboseOutput then log p & " " & sConnStr & vbcrlf
                
        objConn.ConnectionString = sConnStr
        objConn.open
        Set ObjRS = CreateObject("ADODB.Recordset")
        ObjRS.CursorType = adOpenStatic
        Objrs.LockType = adLockOptimistic
        'ObjRS.CursorLocation = adUseClient
        Set objRS.ActiveConnection = objConn
        
    end sub

    '********************************************
    sub closedb
        if objRS.State = 1 then objRS.Close
        objConn.close
        Set objRS = nothing
        Set objConn = nothing
    end sub

    '********************************************
    function query(s)
        dim bOutputTextType: bOutputTextType = true
        dim ss: ss = "QUERY: " & s & vbcrlf
        on error resume next
        if typename(objConn) = "Empty" then
            log "No connection object. Did you call 'opendb()' first?"
            on error goto 0
            exit function
        end if
        dim oRs: set oRs = objConn.execute(s & ";")
        if err.number <> 0 then
            query = "QUERY ERROR --> " & s & ":" & err.description
            exit function
        else
            on error goto 0
            if oRs.state = 1 then
                if (not oRs.BOF) and (not oRs.EOF)  then
                    on error resume next
                    if err.number <> 0 then
                        log err.number & " " & err.description
                        err.clear
                    end if
                    ' if iBitness = 32 then oRs.MoveFirst
                    oRs.MoveFirst
                    on error goto 0
                    do while oRs.EOF = false
                        dim ff
                        on error resume next
                        dim fieldCount: fieldCount = oRs.fields.count
                        if err.number <> 0 then
                            log "0 ERROR " & err.number & "-->" & err.description
                            exit function
                        end if
                        on error goto 0
                        for each ff in oRs.Fields
                            on error resume next
                            select case ff.type
                                case 1      ' vbNull
                                    ss = ss & ff.Name & ":NULL "
                                case 128    ' adBinary
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & BinaryToString(ff.value,false) & " hex:" & BinaryToString(ff.value,true) & separator
                                    else
                                        ss = ss & ff.Name & "(adBinary):" & BinaryToString(ff.value,false) & " hex:" & BinaryToString(ff.value,true) & separator
                                    end if
                                case 2      ' vbInteger
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(vbInteger):" & ff.Value & separator
                                    end if
                                case 3      ' vbLong
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(vbLong):" & ff.Value & separator
                                    end if
                                case 4      ' vbSingle
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(vbSingle):" & ff.Value & separator
                                    end if
                                case 5      ' vbDouble
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(vbDouble):" & ff.Value & separator
                                    end if
                                case 135    ' adDBTimeStamp
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(adDBTimeStamp):" & cdbl(ff.Value) & " " & ff.Value & separator
                                    end if
                                case 202    ' adVarWChar
                                    if typename(ff.Value) = "Null" then
                                        if not bOutputTextType then
                                            ss = ss & ff.Name & "(" & ff.type & "):Null" & separator
                                        else
                                            ss = ss & ff.Name & "(text):Null" & separator
                                        end if
                                    else
                                        if not bOutputTextType then
                                            ss = ss & ff.Name & "(" & ff.type & ")(" & len(ff.Value) & "):" & ff.Value & separator
                                        else
                                            ss = ss & ff.Name & "(text)(" & len(ff.Value) & "):" & ff.Value & separator
                                        end if
                                    end if
                                    REM wscript.echo typename(ff.value) & " " & lenb(ff.value) & " [" & BinaryToString(ff.value,false) & "]"
                                case 203    ' adLongVarWChar (memo)
                                    if not bOutputTextType then
                                        ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    else
                                        ss = ss & ff.Name & "(adLongVarWChar):" & ff.Value & separator
                                    end if
                                case else
                                    REM ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                                    ss = ss & ff.Name & "(" & ff.type & "):" & ff.Value & separator
                            end select
                            if err.number <> 0 then
                                log "1 ERROR " & err.number & "-->" & err.description
                                log "BinaryToString " & BinaryToString(ff.value,true)
                                ss = ss & ff.Name  & "(" & ff.type & "):" &  "ERR "
                                err.clear
                            end if
                            on error goto 0
                        next
                        on error resume next
                        REM log "MoveNext BOF " & oRs.BOF & " EOF " & oRs.EOF
                        oRs.MoveNext
                        if err.number <> 0 then
                            log "2 ERROR " & err.number & "-->" & err.description
                            log "    query --> " & s
                        end if
                        on error goto 0
                        ' remove the last separator 
                        ss = left(ss,len(ss)-1)
                        ss = ss & vbcrlf
                    loop
                else
                    query = ss & "RS contains no records" & vbcrlf
                    oRs.close
                    exit function
                end if
            else
                query = ss & "RS is not open " & oRs.state & vbcrlf
                exit function
            end if
        end if
        oRs.close
        set oRs = nothing
        query= ss
    end function

    '********************************************
    function query2csv(s)
        query2csv = 0
        aQueryResults(2).removeall
        aQueryResults(3) = ""
        dim rowCount: rowCount = 0
        dim bOutputTextType: bOutputTextType = true
        dim ss
        aQueryResults(0) = s
        on error resume next
        if typename(objConn) = "Empty" then
            aQueryResults(3) = "No connection object. Did you call 'opendb()' first?"
            on error goto 0
            exit function
        end if
        dim oRs: set oRs = objConn.execute(s & ";")
        if err.number <> 0 then
            aQueryResults(3) = aQueryResults(3) & err.number & ":" & err.description & vbcrlf
            exit function
        else
            on error goto 0
            if oRs.state = 1 then
                dim ff
                ss = ""
                for each ff in oRs.Fields
                if bOutputTextType then
                    ss = ss & ff.Name & "(" & dDataTypes(ff.type) & ")" & separator

                else
                    ss = ss & ff.Name & "(" & ff.type & ")" & separator
                end if
                next
                ' remove the last separator 
                ss = left(ss,len(ss)-1)
                aQueryResults(1) = ss
                if (not oRs.BOF) and (not oRs.EOF)  then
                    on error resume next
                    if err.number <> 0 then
                        log err.number & " " & err.description
                        err.clear
                    end if
                    oRs.MoveFirst
                    on error goto 0
                    do while oRs.EOF = false
                        ss = ""
                        on error resume next
                        dim fieldCount: fieldCount = oRs.fields.count
                        if err.number <> 0 then
                            aQueryResults(3) = aQueryResults(3) & "0: " & err.number & ":" & err.description & "|"
                            exit function
                        end if
                        on error goto 0
                        for each ff in oRs.Fields
                            on error resume next
                            select case ff.type
                                case 1      ' vbNull
                                    ss = ss & ff.Name & ":NULL "
                                case 128    ' adBinary
                                    ss = ss & BinaryToString(ff.value,true) & separator
                                case 2      ' vbInteger
                                    ss = ss & ff.Value & separator
                                case 3      ' vbLong
                                    ss = ss & ff.Value & separator
                                case 4      ' vbSingle
                                    ss = ss & ff.Value & separator
                                case 5      ' vbDouble
                                    ss = ss & ff.Value & separator
                                case 135    ' adDBTimeStamp
                                    ss = ss & ff.Value & separator
                                case 202    ' adVarWChar
                                    if typename(ff.Value) = "Null" then
                                        ss = ss & "Null" & separator
                                    else
                                        ss = ss & chr(34) & replace(ff.Value,chr(34),chr(34)&chr(34)) & chr(34) & separator
                                    end if
                                case 203    ' adLongVarWChar (memo)
                                    ss = ss & chr(34) & replace(ff.Value,chr(34),chr(34)&chr(34)) & chr(34) & separator
                                case else
                                    ss = ss & ff.Value & separator
                            end select
                            if err.number <> 0 then
                                aQueryResults(3) = aQueryResults(3) & "1 " & err.number & ":" & err.description & "|"
                                query2csv = -1
                            end if
                            on error goto 0
                        next
                        on error resume next
                        oRs.MoveNext
                        if err.number <> 0 then
                            aQueryResults(3) = aQueryResults(3) & "2 " & err.number & ":" & err.description & "|"
                            query2csv = -1
                        end if
                        on error goto 0
                        ' remove the last separator 
                        ss = left(ss,len(ss)-1)
                        aQueryResults(2).add rowCount, ss
                        rowCount = rowCount + 1
                        query2csv = rowCount
                    loop
                else
                    'aQueryResults(3) = "RS contains no records" & "|"
                    oRs.close
                    set oRs = nothing
                    query2csv = 0
                    exit function
                end if
            else
                'aQueryResults(3) = "RS is not open " & oRs.state & "|"
                query2csv = 0
                set oRs = nothing
                exit function
            end if
        end if
        oRs.close
        set oRs = nothing
    end function

    '********************************************
    function openRecordSet(s)
        on error resume next
        ObjRS.Open(s)
        if err.number <> 0 then
            log err.number & " " & err.description
        end if
        on error goto 0
        if objRS.State = 1 then 
            openRecordSet = ObjRs.recordcount
        else
            openRecordSet = -1
        end if
    end function
    
    '********************************************
    public function log (s)
        on error resume next
        wscript.echo s
        if err.number = 500 then
            console.log s
        end if
        on error goto 0
    end function
    
    '********************************************
    public function logResult(r)
        ' log query2csv() result contained in aQueryResults to console
        if r >= 0 then
            log "QUERY  " & aQueryResults(0)
            log "returned " & (r) & " rows"
            if len(aQueryResults(3)) > 0  then log aQueryResults(3)
            log aQueryResults(1)
            dim vKey: for each vKey in aQueryResults(2)
                log aQueryResults(2).item(vKey)
            next
            log ""
        else
            log "QUERY  " & aQueryResults(0)
            log "returned no rows (" & r & ")"
            if len(aQueryResults(3)) > 0  then log aQueryResults(3)
            log ""
        end if
    end function
    
    '********************************************
    private function DoNothing()
        with oWShell
            .run "%comspec% /c exit", 0, false
        End With  
    end function

    '********************************************
    Public Function GetLogTime() 
        Dim strNow 
        strNow = Now() 
        GetLogTime = _
            Year(strNow) & "-" & _
            Pad(Month(strNow), 2, "0", True) & "-" & _
            Pad(Day(strNow), 2, "0", True) & "T" & _
            Pad(Hour(strNow), 2, "0", True) & ":" & _
            Pad(Minute(strNow), 2, "0", True) & ":" & _
            Pad(Second(strNow), 2, "0", True) 
    End Function

    '********************************************
    Public Function Pad(strText, nLen, strChar, bFront) 
        Dim nStartLen 
        If strChar = "" Then 
            strChar = "0" 
        End If 
        nStartLen = Len(strText) 
        If Len(strText) >= nLen Then 
            Pad = strText 
        Else 
            If bFront Then 
                Pad = String(nLen - Len(strText), strChar) & strText 
            Else 
                Pad = strText & String(nLen - Len(strText), strChar) 
            End If 
        End If 
    End Function 

end class
