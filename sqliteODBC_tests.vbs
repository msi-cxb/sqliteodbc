option explicit

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
    private dbAccess
    private dbDuck
    private dbDuckDSN
    

    '*************************************************************************
    sub class_initialize()
        bOpenFirstTime = true
        bVerboseOutput = true
        
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
        log "******************************************************"
        log "main"
        
        sqlite_version
        
        exit function
        
        opendb "MEM  "
        log query("PRAGMA compile_options")
        closedb
        
        ' sqlite only
        insertTests false,true
        
        if len(dbSqlite3) > 0 then 
            objFSO.CopyFile dbSqlite3, strFolder & "\testfile.sqlite3", true
        end if
        dbSqlite3 = strFolder & "\testfile.sqlite3"
        log "dbSqlite3 " & dbSqlite3
        
        log dumpPragma
        sqlite3Tests
        sqlite3_BuiltIn_Tests
        sqlite3_rtree_tests
        sqlite3_fts_tests
        sqlite3_page_size_tests
        sqlite3_feature_tests
        sqlite_big_numbers
        recursiveCTE
        generate_series
        calendarExamples
        
        ' only if msi dll has been built
        if true then
            sqlite3_msi_dll_tests
        end if
        
        ' as of 03.46.00 fileio extension is not working in ODBC
        ' file doesn't need to exist but more interesting if c:\temp has some content
        dbSqlite3 = "c:\temp\fileio.sqlite3"
        if objFso.FileExists(dbSqlite3) then
            log dbSqlite3 & " does not exist. Running test without it (file will be created)."
        end if
        sqlite_extension_fileio
        
        ' this file needs to exist generated by pcapToSqlite scripts needs to have withoutpayload
        dbSqlite3 = "M:\Apps\pcapToSqlite\trunk\scripts\packets.sqlite3"
        if objFso.FileExists(dbSqlite3) then
            sqlite3_SQLITE_STMT
            sqlite3_dbstat
        else
            log "could not run sqlite3_SQLITE_STMT/sqlite3_dbstat " & dbSqlite3 & " does not exist"
        end if
        
        dbSqlite3 = ".\testDBs\demoFieldBIT.sqlite3"
        isNull
        
        dbSqlite3 = ".\testDBs\tinyDB2.sqlite3"
        doubleExample
        
        dbSqlite3 = ".\testDBs\icmp.sqlite3"
        getTableInfoSimple
        getTableInfoDetails
        groupConcatExample
        windowFunctionExample
        
        ' these use MEM
        graphExampleOne
        graphExampleTwo
        isValidIntOrFloat
        varCharStringIssue
        unicodeCharacterTest

        dbSqlite3 = strFolder & "\testfile.sqlite3"
        sqlite_extension_functions_series
        sqlite_extension_functions_wholenumber
        sqlite_extension_functions_csv
        sqlite_extension_functions_regex
        sqlite_extension_functions_sha
        sqlite_extension_functions_ieee754
        sqlite_extension_functions_totype
        sqlite_extension_functions_tests
        sqlite_extension_functions_checkfreelist
        sqlite_extension_uuid
        sqlite_extension_crypto
        
        dbSqlite3 = strFolder & "\testDBs\testfile.sqlite3"
        sqlite_extension_geopoly
        
        ' in development
        sqlite_extension_spatialite
        
        ' 3.37, 3.38, 3.39 updates and improvements
        sqlite_right_join
        sqlite_full_outer_join
        sqlite_isDistinctFrom
        sqlite_havingWithoutGroupBy
        sqlite_localtimeModifierMaintainsFractSecs
        sqlite_unixepochFunction
        sqlite_autoModifier
        sqlite_juliandayModifier
        sqlite_printf_format
        sqlite_strictTable
        
        ' 3.46 updates and improvements
        sqlite_double_quoted_strings
        sqlite_strftime
        sqlite_json
        
        ' JSON to do
        ' sqlite_json_virtual_columns

        
    end function

    '********************************************
    public function sqlite_double_quoted_strings
        opendb "SQL3 "
        log "The SQLITE_DBCONFIG_DQS_DML option activates or deactivates the legacy "
        log "double-quoted string literal misfeature for DML statements only, that is "
        log "DELETE, INSERT, SELECT, and UPDATE statements. The recommended setting is 0, "
        log "meaning that double-quoted strings are disallowed in all contexts. "
        log "However, the default setting is 3 for maximum compatibility with legacy applications."
        log "ODBC driver setting is 3 for maximum compatibility. Hence both examples below work."
        log ""
        logResult query2csv("CREATE TABLE t0(c0 INTEGER);")
        logResult query2csv("select * from sqlite_master where type=""table"";")
        logResult query2csv("select * from sqlite_master where type='table';")
        closedb
    end function
    
    '********************************************
    public function sqlite_strftime
        ' The strftime() SQL function now supports %G, %g, %U, and %V.
        ' New conversion letters on the strftime() SQL function: %e %F %I %k %l %p %P %R %T %u
        opendb "MEM  "
        logResult query2csv("SELECT strftime('%e -- %F -- %I -- %k -- %l -- %p -- %P -- %R -- %T -- %u -- %G -- %g -- %U -- %V', '2013-10-07T08:23:19.120') as r;")
        log """ 7 -- 2013-10-07 -- 08 --  8 --  8 -- AM -- am -- 08:23 -- 08:23:19 -- 1 -- 2013 -- 13 -- 40 -- 41"""
        closedb
    end function
    
    '********************************************
    public function sqlite_json
        opendb "MEM  "
        ' just a couple of examples to show that JSON works...
        
        ''{"this":"is","a":["test"]}'
        logResult query2csv("select json(' { ""this"" : ""is"", ""a"": [ ""test"" ] } ')")
        '[1,2,"3",4]'
        logResult query2csv("select json_array(1,2,'3',4)") 
        '[""[1,2]"]'")
        logResult query2csv("select json_array('[1,2]')") 
        '[[1,2]]'")
        logResult query2csv("select json_array(json_array(1,2))") 
        '[1,null,"3","[4,5]","{\"six\":7.7}"]'")
        logResult query2csv("select json_array(1,null,'3','[4,5]','{""six"":7.7}')") 
        '[1,null,"3",[4,5],{"six":7.7}]'")
        logResult query2csv("select json_array(1,null,'3',json('[4,5]'),json('{""six"":7.7}'))") 
        '$.c' → '[4,5,{"f":7}]'
        logResult query2csv("select '{""a"":2,""c"":[4,5,{""f"":7}]}' -> '$.c'") 
        
        closedb
    end function

    '********************************************
    public function sqlite_json_virtual_columns
        ' https://antonz.org/json-virtual-columns/
    
        REM data
        REM {"timestamp":"2022-05-15T09:31:00Z","object":"user","object_id":11,"action":"login","details":{"ip":"192.168.0.1"}}
        REM {"timestamp":"2022-05-15T09:32:00Z","object":"account","object_id":12,"action":"deposit","details":{"amount":"1000","currency":"USD"}}
        REM {"timestamp":"2022-05-15T09:33:00Z","object":"company","object_id":13,"action":"edit","details":{"fields":["address","phone"]}}
    
        REM slow
        REM select
          REM json_extract(value, '$.object') as object,
          REM json_extract(value, '$.action') as action
        REM from events
        REM where json_extract(value, '$.object_id') = 11;
        
        REM fast using SQL virtual columns
        
        REM alter table events
        REM add column object_id integer
        REM as (json_extract(value, '$.object_id'));

        REM alter table events
        REM add column object text
        REM as (json_extract(value, '$.object'));

        REM alter table events
        REM add column action text
        REM as (json_extract(value, '$.action'));
        
        REM create index events_object_id on events(object_id);
        
        REM select object, action
        REM from events
        REM where object_id = 11;
        
    end function

    '********************************************
    public function sqlite_generated_columns
        ' https://antonz.org/generated-columns/
        
        REM id	tax
        REM 11	15.4
        REM 12	17.16
        REM 21	18.48
        REM 22	21.6
        REM 23	24.96
        REM 24	24.96
        REM 25	28.8
        REM 31	23.04
        REM 32	23.04
        REM 33	24

        REM -- add generated column to the table
        REM -- (comment this out after the first run)
        REM alter table people
        REM add column tax real as (
        REM income * tax_rate
        REM );

        REM -- query generated column as the regular one
        REM select id, round(tax,2) as tax
        REM from people;
        
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
    public function sqlite_printf_format
        log "******************************************************"
        log "sqlite_printf_format (added 3.38.0)"
        
        opendb "MEM  "
        log "this will throw error but it does load the extension"
        log query("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\sqlfcmp.dll') as ext_loaded")
        log "printf and format produce the same result"
        logResult query2csv( _
            "with sig(prec, sig) as " & _
            "     ( " & _
            "      values ('binary16',11), " & _
            "             ('binary32',24), " & _
            "             ('binary64',53), " & _
            "             ('binary128',113), " & _
            "             ('binary256',237) " & _
            "     ) " & _
            "select prec,  " & _
            "       format('%!.17f', ulp(360.0, sig)) as f, " & _
            "       printf('%!.17f', ulp(360.0, sig)) as p " & _
            "  from sig; " _
        )
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
    public function sqlite_extension_vfsstat 
        log "****************************************************************************"
        log "vfsstat - loads via connection string, but not providing much in the way of vfsstats"
        log "only count that updates is randomness..."
        log ""
        
        REM this works in sqlite3.exe so the extension is ok, just doesn't work via ODBC
        REM could the problem be that you need to load the extension before opening DB?
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

        REM opendb "SQL3-vfsstat"
        REM opendb "SQL3 "
        REM logResult query2csv("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\vfsstat.dll') as ext_loaded")
        REM logResult query2csv("DROP TABLE IF EXISTS t1;")
        REM logResult query2csv("CREATE TABLE t1(x integer,y blob);")
        REM logResult query2csv("INSERT INTO t1 VALUES(123, randomblob(50));")
        REM logResult query2csv("CREATE INDEX t1x ON t1(x);")
        REM logResult query2csv("SELECT name FROM sqlite_master;")
        REM logResult query2csv("select * from t1;")
        REM logResult query2csv("DROP TABLE IF EXISTS t1;")
        REM logResult query2csv("SELECT name FROM sqlite_master;")
        REM logResult query2csv("VACUUM;")
        REM logResult query2csv("SELECT * FROM vfsstat WHERE count>0;")
        REM closedb
        
        dim oConn: Set oConn = CreateObject("ADODB.Connection")
        oConn.Provider = "{SQLite3 ODBC Driver}"
        oConn.open
        log oConn.state
        dim iAffected
        dim oRs: set oRs = oConn.execute("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\vfsstat.dll') as ext_loaded",iAffected)
        log iAffected
        oConn.close
        
    end function

    '********************************************
    public function longSqlStringReturn
        ' inspired by https://til.simonwillison.net/sqlite/column-combinations
        log "******************************************************"
        log "longSqlString"
        log ""
        log "This reads in a query from a file...the cool part of this SQL is that"
        log "it generates SQL that can then itself be executed. This can be handy"
        log "if you, say, want to get some kind of info about tables in your database."
        log "SQLite open this up as it has some function that allow the query to do"
        log "a little introspection of the DB itself. This would work great, except that"
        log "ADO defaults columns with strings to 255 characters by default. "
        log ""
        log "File longStringTest.sql contains a query that identifies patterns in DB rows in terms"
        log "of which columns are not null. (see link in code for more details)."
        log ""
        dim sFileName: sFileName = ".\sql\longStringTest.sql"
        dim oFile: Set oFile = objFSO.OpenTextFile(sFileName)
        dim sFileContentsTemplate: sFileContentsTemplate = oFile.ReadAll
        oFile.close
        dim sSql,r
        sSql = replace(sFileContentsTemplate,"___TABLENAME___","pcap")
        set oFile = nothing
        log ""
        dbSqlite3 = ".\testDBs\icmp.sqlite3"
        opendb "SQL3 "
        log "processing " & sFileName & " for table 'pcap'"
        r = query(sSql)
        log r
        log ""
        log "The problem is that the returned result is truncated to 255 char. The full"
        log "result is over 1000 char long. Lets try for a table with less columns."
        log ""
        sSql = replace(sFileContentsTemplate,"___TABLENAME___","meta")
        log "processing " & sFileName & " for table 'meta'"
        r = query(sSql)
        log r
        log "OK the resulting query is approx 230 char so less than 255 char limit. Now execute the resulting SQL"
        r = split(r,":")(2)
        log query(r)
        log ""
        log "OK yeah not very exciting but that is only because the table has 2 non-null columns and one record."
        log "Yes there are other ways to do this... Issue I'm illustrating here is ADO default of 255 chars for strings "
        log "unless they are defined in a create table statement."
        log ""
        log "By the way, this will work in C# and other non ADO places."
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
        dim i,j: for i = 0 to 310
            j = j & (i mod 10)
            if i >= 290 then sSql = sSql & "(" & i & ",'" & j & "'),"
        next
        sSql = left(sSql,len(sSql)-1) ' remove the last comma
        log query(sSql)
        logResult query2csv("select * from t;")
        log "scroll to right and see how returned string is truncated at 300"
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
    public function groupConcatExample 
        opendb "SQL3 "
        
        log ""
        log "note that order of elements in result is based on 'presentation' hence the order by desc in CTE"
        log "note that there does appear to be a limit of 268 char to the size of the string generated by group_concat()...what?!?!?"
        log "limit is the same for both 32/64 bit"
        log query( _
            "with icmp as ( " & _
                "select " & _
                    "* " & _
                "from " & _
                    "pcap " & _
                "where Ipv4_Protocol = 'Icmp' " & _
                "and id < 10000 " & _
                "order by id desc" & _
            ") " & _
            "select " & _
                "length(group_concat(id)) as id_values_len, " & _
                "group_concat(id) as id_values " & _
            "from " & _
                "icmp " _
        )
        log ""
        
        closedb

    end function

    '********************************************
    public function windowFunctionExample
        if objFso.FileExists(dbSqlite3) then
            opendb "SQL3 "
            log ""
            log "window lag function - find all ICMP packets where there is more than a 5 sec gap between packets"
            log "uses a Common Table Expressions (CTE) first to get all ICMP packets, then window lag function to calc delta time between packets"
            log ""
            log query( _
                "select * from ( " & _
                    "with icmp as (" & _
                        "select " & _
                            "* " & _
                        "from " & _
                            "pcap " & _
                        "where Ipv4_Protocol = 'Icmp' " & _
                    ") " & _
                    "select " & _
                        "reporttime," & _
                        "(reporttime_dbl - lag(reporttime_dbl,1) OVER (order by reporttime_dbl)) * 86400 as delta_reporttime_dbl_sec " & _
                    "from " & _
                        "icmp " & _
                ") " & _
                "where " & _
                    "delta_reporttime_dbl_sec > 5 " _ 
            )
            log ""
            closedb
        else
            log dbSqlite3 & " does not exist"
        end if
    end function

    '********************************************
    public function testDbInventory
        dim sPath: sPath = objFSO.GetAbsolutePathName(".\testDBs")
        dim oFolder: set oFolder = objFSO.GetFolder(sPath)
        log "path: " & oFolder.path & vbcrlf
        log "files:"
        dim oFile: for each oFile in oFolder.Files
            dbSqlite3 = oFile.path
            log dbSqlite3
            opendb "SQL3 "
            log query(" select name from sqlite_master where type = 'table';")
            closedb
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
    public function harvest
        REM note that harvest can only be run 32 bit as it uses access.
        
        REM There is a memory limit that can be hit by copying large tables from sqlite to access
        REM via ODBC driver. Seems like a hard limit that has nothing to do with installed memory.
        REM So far have not found a way to get past this...
        
        REM Assumes that packets.sqlite in and packetsAccess.mdb in M:\Apps\pcapToSqlite\trunk\scripts
        REM and withoutpayload_9 table exists (this table has 900000 records and 87 columns)
        
        log "****************************************************************************"
        log "harvest"
        opendb "JET4 "
        dim sSql, params
        
        params = _
            "ID," & _
            "frameIdx," & _
            "srcFile," & _
            "location," & _
            "PKT_Col," & _
            "Time_UnixEpoch," & _
            "ReportTime," & _
            "ReportTime_dbl," & _
            "Time_msec," & _
            "Pkt_DataLength," & _
            "Pkt_LinkType," & _
            "ETH_Col," & _
            "src_mac," & _
            "dst_mac," & _
            "eth_type," & _
            "ARP_Col," & _
            "Arp_Operation," & _
            "Arp_SendHwdrAddr," & _
            "Arp_TgtHwdrAddr," & _
            "Arp_SendProtoAddr," & _
            "Arp_TgtProtoAddr," & _
            "Arp_ProtoAddrType," & _
            "Arp_ProtoAddrLen," & _
            "IPv4_Col," & _
            "Ipv4_Version," & _
            "Ipv4_HdrLen," & _
            "Ipv4_Id," & _
            "Ipv4_TotalLen," & _
            "Ipv4_PayloadLen," & _
            "Ipv4_TTL," & _
            "Ipv4_Protocol," & _
            "Ipv4_SrcAddr," & _
            "Ipv4_DestAddr," & _
            "Ipv4_FragOffset," & _
            "Ipv4_FragFlags," & _
            "IPv4_Payload," & _
            "IPv6_Col," & _
            "Ipv6_Version," & _
            "Ipv6_TrafficClass," & _
            "Ipv6_FlowLabel," & _
            "Ipv6_HdrLen," & _
            "Ipv6_TotalLen," & _
            "Ipv6_PayloadLen," & _
            "Ipv6_Protocol," & _
            "Ipv6_TTL," & _
            "Ipv6_SrcAddr," & _
            "Ipv6_DestAddr," & _
            "IPv6_Payload," & _
            "Tcp_Col," & _
            "Tcp_SrcPort," & _
            "Tcp_DestPort," & _
            "Tcp_SeqNum," & _
            "Tcp_AckNum," & _
            "Tcp_DataOffset," & _
            "Tcp_AllFlags," & _
            "Tcp_CWR," & _
            "Tcp_Urg," & _
            "Tcp_Ack," & _
            "Tcp_Psh," & _
            "Tcp_Rst," & _
            "Tcp_Syn," & _
            "Tcp_Fin," & _
            "Tcp_ECN," & _
            "Tcp_WindowSize," & _
            "Tcp_CheckSum," & _
            "Tcp_UrgentPtr," & _
            "Tcp_Options," & _
            "Tcp_Payload," & _
            "Udp_Col," & _
            "Udp_SrcPort," & _
            "Udp_DestPort," & _
            "Udp_Length," & _
            "Udp_CheckSum," & _
            "Udp_VldUdpChkSum," & _
            "Udp_VldChkSum," & _
            "Udp_Payload," & _
            "Icmpv4_Col," & _
            "Icmpv4_ID," & _
            "Icmpv4_TypeCode," & _
            "Icmpv4_Sequence," & _
            "Icmpv4_Payload," & _
            "Icmpv6_Col," & _
            "Icmpv6_Type," & _
            "Icmpv6_Code," & _
            "Icmpv6_Payload," & _
            "Pkt_Payload," & _
            "theend"
        
        REM comment the following line out to use selected parameters from list above
        params = "*"
        
        sSql = _
            "SELECT " & _
                params & " " & _
            "INTO " & _
                "cxb " & _
            "FROM " & _
            "( " & _
                "SELECT " & _
                    params & " " & _
                "FROM " & _
                    "[ODBC;DSN=SQLite3 Datasource;Database=" &dbSqlite3 & ";].[withoutpayload_25] " & _
            ")"
        
        REM log "drop the table if it exists"
        REM log query("drop table cxb") & vbcrlf
        
        REM log "build the table"
        REM log query(sSql) & vbcrlf

        log "dump the schema"
        log query("select * from pragma_table_info('cxb')") & vbcrlf
        
        REM log "dump the count of rows"
        REM dim r: r = split(split(query("select count(1) from cxb"),vbcrlf)(1),":")(1)
        
        REM log "dump the count of columns"
        REM dim c: c = (ubound(split(query("select count(1) as cnt from pragma_table_info('cxb')"),",")))+1
        REM log "rows " & r & " columns " & c
        
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
    public function doubleExample
        log "******************************************************"
        log "doubleExample"
        log ""
        if objFso.FileExists(dbSqlite3) then
            opendb "SQL3 "
            dim s: s = query("SELECT ReportTime AS fieldTime,ReportTime FROM [VIZ_VWS] ORDER BY ReportTime LIMIT 1;")
            dim a: a = split(s,vbcrlf)
            log s
            log "same " & (a(1) = "fieldTime(adDBTimeStamp):38949.1257924769 8/20/2006 3:01:08 AM,ReportTime(adDBTimeStamp):38949.1257924769 8/20/2006 3:01:08 AM") & vbcrlf
            closedb
        else
            log dbSqlite3 & " does not exist"
        end if
    end function
    
    '********************************************
    public function isNull 
        log "******************************************************"
        log "illustrate 'is [not] null'. AmplifyingRemarks contains null in all but two records in AC_C172_V1E1_C2."
        log "also comment on null and ODBC"
        log ""
        if objFso.FileExists(dbSqlite3) then
            opendb "SQL3 "
            log "is not null - expect two records returned"
            log ""
            log query("SELECT Reporttime,AmplifyingRemarks FROM [AC_C172_V1E1_C2] where AmplifyingRemarks is not null;")
            log ""
            log "is null"
            log "note the value returned from ODBC is type text so you can't test for null in scriptland."
            log ""
            log query("SELECT Reporttime,AmplifyingRemarks FROM [AC_C172_V1E1_C2] where AmplifyingRemarks is null;")
            log ""
            closedb
        else
            log dbSqlite3 & " does not exist"
        end if
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

    'https://www.gaia-gis.it/fossil/libspatialite/home
    'https://www.gaia-gis.it/gaia-sins/spatialite_topics.html
    'https://www.gaia-gis.it/gaia-sins/spatialite-tutorial-2.3.1.html


    '********************************************
    public function sqlite_extension_spatialite
        log "****************************************************************************"
        log "sqlite_extension_spatialite"
        log "TODO assuming spatialite can be build from source"
        log ""
    end function

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
    public function sqlite_extension_geopoly_DOESNOTWORK
        ' It would be cool if we could read in the geopoly.sql file...
        ' However, this won't work as is because the .sql file contains dot 
        ' commands (e.g. .print) which only work in the sqlite3.exe application.
        dim sFileName: sFileName = ".\sql\geopoly.sql"
        dim oFile: Set oFile = objFSO.OpenTextFile(sFileName)
        dim sFileContents: sFileContents = oFile.ReadAll
        oFile.close
        opendb "SQL3 "
        log query(sFileContents)
        closedb
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

        log "fileio dll does not work with SELECT load_extension(), so need to load via connection string"
        opendb "SQL3-fileio"
        
        REM dump file info to table so we can look at schema
        log query("drop table if exists [fileInfo_C_Temp];") 
        log query("create table [fileInfo_C_Temp] as select * FROM fsdir('c:\temp');") 
        log query("PRAGMA table_info('fileInfo_C_Temp');")
        log query("select count(1) as numRecords from [fileInfo_C_Temp];")
        
        log "script crashes when blob is binary and large, and error handling does not catch issue"
        log "so you can't so a select * or select with data field if folder contains large binary file."
        log query("select name,mode,mtime,datetime(mtime, 'unixepoch','localtime') as reporttime from [fileInfo_C_Temp];")
        
        log "if $dir is provided, and $path is relative then $path interpreted relative to $dir"
        log "if folder does not exist then this will fail..."
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
        log "same but with explicit path only passed to fsdir()"
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
            
        log "read content of a text file where we know we can handle the blob in vbscript data types"
        log query( "select cast( readfile('c:\temp\file.txt') as text) as fileContent;" )
        
        log "db table column 'data' contains binary contennts of MSI.gif, write that out to a new file"
        log "if successful, writefile returns the number of bytes written. if not, empty recordset so nothing returned."
        log query( "SELECT writefile('c:\temp\MSI_new.gif',data) as bytesWritten FROM [fileInfo_C_Temp] WHERE name='c:\temp/MSI.gif';")
        
        closedb
    end function

    '********************************************
    public function sqlite_extension_functions_ieee754
        log "******************************************************"
        log "sqlite_extension_functions_ieee754"
        log "ieee754 dll does not work with SELECT load_extension(), so need to load via connection string"
        opendb "SQL3-ieee754"
        REM can't filter the function list provided by PRAGMA for ieee methods...
        REM log query("PRAGMA function_list;")
        REM so we get a recordset and do it the "hard" way
        log getSortedFunctionList("ieee") & vbcrlf
        log ""
        REM exercise the various ieee methods
        log "expect r:ieee754(181,-2)"
        log query("SELECT ieee754(45.25) as r;")
        log "expect r:ieee754(181,-2)"
        log query("SELECT ieee754(181,-2) as r;")
        log "expect m:181 e:-2"
        log query("SELECT ieee754_mantissa(45.25) as m, ieee754_exponent(45.25) as e;")
        closedb
    end function

    '********************************************
    public function sqlite_extension_functions_series
        log "******************************************************"
        log "sqlite_extension_functions_series"
        opendb "SQL3 "
        log "series extension works with SELECT load_extension() result is loaded(202):"
        log query("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\series.dll') as loaded;")
        
        log "PRAGMA function_list does not see any series functions..."
        log getSortedFunctionList("series") & vbcrlf
        
        log "...but is is there...generate_series() expect values from 0 to 100 in steps of 5"
        log query("SELECT * FROM generate_series(0,100,5);")
        closedb
        
        opendb "SQL3-series"
        log "series extension via connection string ok"
        log "even when it works (32 bit) PRAGMA function_list still does not see the function"
        log getSortedFunctionList("series") & vbcrlf
        log "generate_series() expect values from 0 to 100 in steps of 5"
        log query("SELECT * FROM generate_series(0,100,5);")
        closedb
        log "END END END sqlite_extension_functions_series"
    end function

    '********************************************
    public function sqlite_extension_functions_totype
        log "******************************************************"
        log "sqlite_extension_functions_totype"
        log "totype extension dll does not work with SELECT load_extension()"
        opendb "SQL3-totype"
        log getSortedFunctionList("toreal") & vbcrlf
        log getSortedFunctionList("tointeger") & vbcrlf
        ' If X is an integer, real, or string value that can be
        ' losslessly represented as an integer, then tointeger(X)
        ' returns the corresponding integer value.
        ' If X is an 8-byte BLOB then that blob is interpreted as
        ' a signed two-compliment little-endian encoding of an integer
        ' and tointeger(X) returns the corresponding integer value.
        ' Otherwise tointeger(X) return NULL.
        log "totype dll extension does not work with SELECT load_extension(), so need to load via connection string"
        log "tointeger() expect values i:8 r:8 s:8"
        log query("SELECT tointeger(8) as i,tointeger(8) as r, tointeger('8') as s;")
        log "tointeger() expect values i: r: s:"
        log query("SELECT tointeger(8.1) as i,tointeger(8.1) as r, tointeger('8.1') as s;")
        log "tointeger() expect values i: r: s:"
        log query("SELECT tointeger(8.9) as i,tointeger(8.9) as r, tointeger('8.9') as s;")
        ' If X is an integer, real, or string value that can be
        ' convert into a real number, preserving at least 15 digits
        ' of precision, then toreal(X) returns the corresponding real value.
        ' If X is an 8-byte BLOB then that blob is interpreted as
        ' a 64-bit IEEE754 big-endian floating point value
        ' and toreal(X) returns the corresponding real value.
        ' Otherwise toreal(X) return NULL.
        log "toreal() expect values i:8 r:8 s:8"
        log query("SELECT toreal(8) as i,toreal(8) as r, toreal('8') as s;")
        log "toreal() expect values i:8.1 r:8.1 s:8.1"
        log query("SELECT toreal(8.1) as i,toreal(8.1) as r, toreal('8.1') as s;")
        log "toreal() expect values i:8.9 r:8.9 s:8.9"
        log query("SELECT toreal(8.9) as i,toreal(8.9) as r, toreal('8.9') as s;")
        closedb
    end function

    '********************************************
    public function sqlite_extension_functions_wholenumber
        opendb "SQL3 "
        log "******************************************************"
        log "sqlite_extension_functions_wholenumber"
        log "wholenumber extension dll works with SELECT load_extension()"
        log query("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\wholenumber.dll') as loaded;")
        log getSortedFunctionList("wholenumber") & vbcrlf
        log ""
        
        log query("drop table if exists nums;")
        log query("CREATE VIRTUAL TABLE nums USING wholenumber;")
        log "expect values from 1 to 9"
        log query("SELECT value FROM nums WHERE value>=1 and value<=9;")
        closedb
    end function

    '********************************************
    public function sqlite_extension_functions_sha
        dim s: s = "'now is the time for all good men to come to the aid of their country.'"
        opendb "MEM  "
        log "******************************************************"
        log "sqlite_extension_functions_sha"
        log "load_extension() does not complete successfully, only sha3 methods are loaded (Function sequence error)"
        log query("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\shathree.dll') as loaded;")
        log getSortedFunctionList("sha") & vbcrlf
        log "sha1/shathree dll extensiosn does not work with SELECT load_extension(), so need to load via connection string"
        log "sha1 EE0D681BAD6E6E4F0675A2E37438A2A196E17D16"
        log query("SELECT sha1(" & s & " ) as sha1;")
        log "224: ABA5FC2514949EF2064594330E1B493884239A8B809CE28DD436BE8A"
        log query("SELECT sha3(" & s & " ,224) as sha3_224;")
        log "256: 69778BBA85582DBD8754342CD6CDA7FFEFF622C2EC3EE8D0711F137DE2F97D61"
        log query("SELECT sha3(" & s & " ,256) as sha3_256;")
        log "384: 80DBBAC60521E64D88300C2F467AC0ADE123A7A4ADCBACBD1984231EB6A0C10F751A289AE0F8584A7DB08DDA91966263"
        log query("SELECT sha3(" & s & " ,384) as sha3_384;")
        log "512: AAFAF4A01580D32DAB323CDC93D1BE4F43BC2AAA21EF2915F97716215434D2F52FA0F601F1381BE52A038BC2E1C2EB2C470B45B4BDA8D97EBA6ADC1C0B9F82A8"
        log query("SELECT sha3(" & s & " ,512) as sha3_512;")
        closedb
        log ""
        
        opendb "SQL3-sha"
        log "loading via connection string ensures both sha1 and sha3 functions are loaded"
        log getSortedFunctionList("sha") & vbcrlf
        log "sha1/shathree dll extensiosn does not work with SELECT load_extension(), so need to load via connection string"
        log "sha1 EE0D681BAD6E6E4F0675A2E37438A2A196E17D16"
        log query("SELECT sha1(" & s & " ) as sha1;")
        log "224: ABA5FC2514949EF2064594330E1B493884239A8B809CE28DD436BE8A"
        log query("SELECT sha3(" & s & " ,224) as sha3_224;")
        log "256: 69778BBA85582DBD8754342CD6CDA7FFEFF622C2EC3EE8D0711F137DE2F97D61"
        log query("SELECT sha3(" & s & " ,256) as sha3_256;")
        log "384: 80DBBAC60521E64D88300C2F467AC0ADE123A7A4ADCBACBD1984231EB6A0C10F751A289AE0F8584A7DB08DDA91966263"
        log query("SELECT sha3(" & s & " ,384) as sha3_384;")
        log "512: AAFAF4A01580D32DAB323CDC93D1BE4F43BC2AAA21EF2915F97716215434D2F52FA0F601F1381BE52A038BC2E1C2EB2C470B45B4BDA8D97EBA6ADC1C0B9F82A8"
        log query("SELECT sha3(" & s & " ,512) as sha3_512;")
        closedb
        
    end function

    '********************************************
    public function sqlite_extension_functions_csv
        log "******************************************************"
        log "sqlite_extension_functions_csv"
        opendb "SQL3 "
        log "CSV Virtual Table via load_extension()"
        log "csv extension dll works with SELECT load_extension() "
        log query("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\csv.dll') as loaded")
        log query("CREATE VIRTUAL TABLE temp.t1 USING csv(filename='test.csv',header=true)")
        log query(" SELECT * FROM t1")
        closedb
        log ""
        
        opendb "SQL3-LoadExt-Csv"
        log "CSV Virtual Table via connection string"
        log query("CREATE VIRTUAL TABLE temp.t1 USING csv(filename='test.csv',header=true)")
        log query(" SELECT * FROM t1")
        closedb
    end function

    '********************************************
    public function sqlite_extension_functions_regex
        log "******************************************************"
        log "sqlite_extension_functions_regex"
        opendb "SQL3 "
        
        log "regexp not builtin for ODBC driver"
        log getSortedFunctionList("regexp") & vbcrlf
        log ""

        log "load_extension errors (Function sequence error) but REGEXP works"
        log query("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\regexp.dll') as loaded")
        log getSortedFunctionList("regexp") & vbcrlf
        log query("SELECT * FROM (select id from test_table where id < 100) where id REGEXP '1[2-3]{1,2}'")
        closedb
        log ""
        
        log "can be loaded via connection string"
        opendb "SQL3-LoadExt-Regexp"
        log getSortedFunctionList("regexp") & vbcrlf
        log query("SELECT * FROM (select id from test_table where id < 100) where id REGEXP '1[2-3]{1,2}'")
        closedb
    end function

    '********************************************
    ' test the extension-functions module that should now only have string and aggregate functions
    public function sqlite_extension_functions_tests
        log "******************************************************"
        log "sqlite_extension_functions_tests"
        log "without extension functions loaded these will all fail with QUERY ERROR" & vbcrlf
        opendb "SQL3 "
        sqlite3_extension_functions 
        closedb
        log "******************************************************"
        log "now extension functions are loaded so these queries will now work" & vbcrlf
        opendb "SQL3-LoadExt-ExtFun"
        sqlite3_extension_functions
        closedb
    end function

    '********************************************
    public function sqlite3_extension_functions
        log "****************************************************************************"
        log "sqlite3_extension_functions"
        ' charindex(S1,S2)
        log query("select charindex('world','hello world!') as x;")
        ' charindex(S1,S2,N)
        log query("select charindex('world','hello world! hello world!',10) as x;")
        ' leftstr(S,N)
        log query("select leftstr('hello world!',5) as x;")
        ' padc(S,N)
        log query("select '|', padc('10',50) as x, '|';")
        ' padl(S,N)
        log query("select '|', padl('10',50) as x, '|';")
        ' padr(S,N)
        log query("select '|', padr('10',50) as x, '|';")
        ' proper(S)
        log query("select proper('hElLo WoRlD!') as x;")
        ' replicate(S,N)
        log query("select replicate('hElLo WoRlD fIvE tImEs!',5) as x;")
        ' reverse(S)
        log query("select reverse('hElLo WoRlD!') as x;")
        ' rightstr(S,N)
        log query("select rightstr('hello world!',6) as x;")
        ' strfilter(S1,S2)
        log query("select strfilter('hello world!','o!') as x;")
        
        log "aggregate functions"
        ' lower_quartile(X) 
        log query("select lower_quartile(id) as x from (select id from test_table limit 100);")
        ' median(X)
        log query("select median(id) as x from (select id from test_table limit 100);")
        ' mode(X)
        log query("select mode(myField_2) as x from (select myField_2 from test_table limit 100);")
        ' stdev(X)
        log query("select stdev(id) as x from (select id from test_table limit 100);")
        ' upper_quartile(X)
        log query("select upper_quartile(id) as x from (select id from test_table limit 100);")
        ' variance(X)
        log query("select variance(id) as x from (select id from test_table limit 100);")
        
    end function

    '********************************************
    public function sqlite_extension_functions_checkfreelist
        log "****************************************************************************"
        log "sqlite_extension_functions_checkfreelist needs to be loaded via connection string"
        opendb "SQL3-checkfreelist"
        log query("SELECT checkfreelist('main');")
        closedb
    end function

    '********************************************
    public function sqlite_extension_uuid 
        log "******************************************************"
        log "sqlite_extension_uuid"
        opendb "MEM  "
        log "load_extension() does not work for uuid.dll"
        logResult query2csv("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\uuid.dll') as ext_loaded")
        closedb
        opendb "SQL3-uuid"
        logResult query2csv("select uuid() as uv4")
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
        dim i
        i=0: log i & " expect 'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11', 0xA0EEBC999CB4EF8BB6D6BB9BD38A11"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=1: log i & " expect 'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11', 0xA0EEBC999CB4EF8BB6D6BB9BD38A11"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=2: log i & " expect 'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11', 0xA0EEBC999CB4EF8BB6D6BB9BD38A11"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=3: log i & " expect 'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11', 0xA0EEBC999CB4EF8BB6D6BB9BD38A11"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=4: log i & " expect 'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11', 0xA0EEBC999CB4EF8BB6D6BB9BD38A11"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=5: log i & " expect 'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11', 0xA0EEBC999CB4EF8BB6D6BB9BD38A11"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=6: log i & " expect 'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11', 0xA0EEBC999CB4EF8BB6D6BB9BD38A11"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=7: log i & " expect null,null missing A at beginning "
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=8: log i & " expect 'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11', 0xA0EEBC999CB4EF8BB6D6BB9BD38A11 missing '}' but ok"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=9: log i & " expect null,null number of digits < 16"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=10: log i & " expect nul,null number of digits < 16"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        i=11: log i & " expect null,null stray '{'"
        logResult query2csv("select uuid_str('" & a(i) & "') as us, uuid_blob('" & a(i) & "') as ub")
        closedb
    end function
    
    '********************************************
    public function sqlite_extension_crypto 
        log "******************************************************"
        log "sqlite_extension_crypto"
        
        opendb "MEM  "
        log "load_extension() DOES work for crypto.dll"
        logResult query2csv("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\crypto.dll') as ext_loaded")
        logResult query2csv("drop table if exists customer;")
        logResult query2csv( _
            "CREATE TABLE IF NOT EXISTS customer AS " & _
            "    with recursive t as ( " & _
            "        select  " & _
            "            0 as x  " & _
            "            union all  " & _
            "            select x+1  " & _
            "            from t  " & _
            "            where x < 86400*10 " & _
            "    )  " & _
            "    select  " & _
            "        x as ID, " & _
            "        hex(md5(x)) as Name, " & _
            "        2459580.5+x/864000.0 as DueDate  " & _
            "    from t;" _
            )
            
        logResult query2csv("select * from customer limit 1;")
        closedb
        log "load_extension() via connection string."
        
        opendb "SQL3-crypto"
        logResult query2csv("drop table if exists customer;")
        logResult query2csv( _
            "CREATE TABLE IF NOT EXISTS customer AS " & _
            "    with recursive t as ( " & _
            "        select  " & _
            "            0 as x  " & _
            "            union all  " & _
            "            select x+1  " & _
            "            from t  " & _
            "            where x < 86400*10 " & _
            "    )  " & _
            "    select  " & _
            "        x as ID, " & _
            "        hex(md5(x)) as Name, " & _
            "        2459580.5+x/864000.0 as DueDate  " & _
            "    from t;" _
            )
            
        logResult query2csv("select * from customer limit 1;")
        closedb
    end function

    '********************************************
    public function sqlite3_SQLITE_STMT
        log "****************************************************************************"
        log "sqlite3_SQLITE_STMT"
        
        ' https://sqlite.org/stmt.html
        
        ' SQLITE_STMT provides information about all prepared statements 
        ' associated with the database connection. Stored in table [sqlite_stmt].
        ' also illustrate the use of the ADODB.Command object

        opendb "SQL3 "
        log "how many rows in withoutpayload table?"
        log query("SELECT count(1) as cnt FROM withoutpayload;")
        log "how many rows in sqlite_master table?"
        log query("SELECT count(1) as cnt FROM sqlite_master;")
        log "names of tables?"
        log query("SELECT name FROM sqlite_master where type = 'table';")
        closedb        

        ' manually build up a connection
        dim objConn: Set objConn = CreateObject("ADODB.Connection")
        dim sConnStr: sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";"
        log sConnStr
        objConn.Open sConnStr
        
        ' some queries to pass to our command objects
        dim aSqlStmts: aSqlStmts = Array( _
            "select type,name,rootpage from sqlite_master where type = 'table';" _
            ,"select * from SOURCES;" _
            ,"select name from sqlite_master;" _
            ,"select name from sqlite_master where type != 'table';" _
            ,"select distinct type from sqlite_master;" _
            ,"select * from withoutpayload limit 10000;" _
        )
        
        ' create an array of command objects all tied to same db connection
        dim aObjCmd: aObjCmd = Array()
        dim sSqlStmt
        for each sSqlStmt in aSqlStmts
            redim preserve aObjCmd(ubound(aObjCmd)+1)
            set aObjCmd(ubound(aObjCmd)) = CreateObject("ADODB.Command")
            aObjCmd(ubound(aObjCmd)).ActiveConnection = objConn
            aObjCmd(ubound(aObjCmd)).prepared = true
            aObjCmd(ubound(aObjCmd)).CommandText = sSqlStmt
        next
        
        log ""
        log "get stats from the sqlite_stmt for all prepared statements on the connection (before execute)"
        log ""
        helper_getStats(objConn)
        log ""
        
        dim ObjCmd
        ' now execute the commands
        log ""
        log "execute commands..."
        for each ObjCmd in aObjCmd
            helper_getResult ObjCmd,false
        next
        
        log ""
        log "get stats from the sqlite_stmt for all prepared statements on the connection (after one execute)"
        log ""
        helper_getStats(objConn)
        log ""

        ' again...
        log ""
        log "again..."
        for each ObjCmd in aObjCmd
            helper_getResult ObjCmd,false
        next

        log ""
        log "get stats from the sqlite_stmt for all prepared statements on the connection (after two executes)"
        log ""
        helper_getStats(objConn)
        log ""
        
        dim oRs
        ' one more time but use the connection object directly (does not affects stats)
        log ""
        log "execute sql directly with connection object only this time..."
        for each sSqlStmt in aSqlStmts
            set oRs = objConn.execute(sSqlStmt)
            log oRs.state & " <-- " & sSqlStmt
            if oRs.state = 1 then oRs.close: set oRs = nothing
        next

        log ""
        log "get stats from the sqlite_stmt for all prepared statements on the connection"
        log ""
        helper_getStats(objConn)
        log ""
        
        log ""
        log "close and reopen the connection this starts us with a new slate..."
        log ""
        objConn.close
        objConn.Open sConnStr

        ' execute with the connection object directly this will result in no prepared statements
        log ""
        log "execute sql directly with connection object..."
        for each sSqlStmt in aSqlStmts
            set oRs = objConn.execute(sSqlStmt)
            log oRs.state & " <-- " & sSqlStmt
            if oRs.state = 1 then oRs.close: set oRs = nothing
        next
        
        log ""
        log "get stats from the sqlite_stmt for all prepared statements on the connection (should be empty)"
        helper_getStats(objConn)
        log ""
        
        for each ObjCmd in aObjCmd
            set objCmd = nothing
        next
        
        objConn.close
        set objConn = nothing
                
    end function

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
    public function sqlite3_dbstat
        opendb "SQL3 "
        log query("SELECT name FROM sqlite_master where type = 'table';")
        closedb        

        opendb "SQL3-LoadExt-ExtFun"
        
        log "****************************************************************************"
        log "dbstat"
        
        ' https://www.sqlite.org/dbstat.html
        ' The DBSTAT virtual table is a read-only eponymous virtual table that returns information about 
        ' the amount of disk space used to store the content of an SQLite database. Example use cases for 
        ' the DBSTAT virtual table include the sqlite3_analyzer.exe utility program and the table size 
        ' pie-chart in the Fossil-implemented version control system for SQLite.

    
        ' limit 20 to keep the output small...
        log query("SELECT * FROM dbstat limit 20;")
        
        ' To find the total number of pages used to store table "withoutpayload" in schema "main", use either of the following 
        ' two queries (the first is the traditional way, and the second shows the use of the aggregated feature):
        log query("SELECT count(*) as pageno FROM dbstat('main') WHERE name='withoutpayload';")
        log query("SELECT pageno FROM dbstat('main',1) WHERE name='withoutpayload';")
        
        ' To see how efficiently the content of a table is stored on disk, compute the amount of space used to hold 
        ' actual content divided by the total amount of disk space used. The closer this number is to 100%, the more 
        ' efficient the packing. (In this example, the 'withoutpayload' table is assumed to be in the 'main' schema. Again, 
        ' there are two different versions that show the use of DBSTAT both without and with the new aggregated feature, respectively.)
        log query("SELECT sum(pgsize-unused)*100.0/sum(pgsize) as usage_percent FROM dbstat WHERE name='withoutpayload';")
        log query("SELECT (pgsize-unused)*100.0/pgsize as usage_percent FROM dbstat WHERE name='withoutpayload' AND aggregate=TRUE;")
        
        ' To find the average fan-out for a table, run:
        log query("SELECT avg(ncell) as avg_fan_out FROM dbstat WHERE name='withoutpayload' AND pagetype='internal';")

        ' Modern filesystems operate faster when disk accesses are sequential. Hence, SQLite will run faster if the content 
        ' of the database file is on sequential pages. To find out what fraction of the pages in a database are sequential 
        ' (and thus obtain a measurement that might be useful in determining when to VACUUM), run a query like the following:
        log query("CREATE TEMP TABLE s(rowid INTEGER PRIMARY KEY, pageno INT);")
        log query("INSERT INTO s(pageno) SELECT pageno FROM dbstat ORDER BY path;")
        log query("SELECT sum(s1.pageno+1==s2.pageno)*1.0/count(*) as sequential FROM s AS s1, s AS s2 WHERE s1.rowid+1=s2.rowid;")
        log query("DROP TABLE s;")
        
        dim sResult: sResult = query(  "SELECT " & _
                        " count(payload) as numPages " & _
                        ",avg(payload) as avgUsedBytesPerPage " & _
                        ",min(payload) as minPayloadPerPage " & _
                        ",max(payload) as maxPayloadPerPage " & _
                        ",stdev(payload) as stddevPayloadPerPage " & _
                        ",avg(unused) as avgUnusedBytesPerPage " & _
                        ",min(unused) as minUnusedPerPage " & _
                        ",max(unused) as maxUnusedPerPage " & _
                        ",stdev(unused) as stddevUnusedPerPage " & _
                    "FROM dbstat " & _
                        "WHERE name='small_01';" _
        )
        
        dim aResult: aResult = split(sResult,vbcrlf)
        log vbcrlf & aResult(0)
        log replace(aResult(1), ",", vbcrlf)
        log ""
        
        
        closedb
    end function

    '********************************************
    public function sqlite_version
        log "****************************************************************************"
        log ""
        opendb "MEM  "
        log query("SELECT sqlite_version() as vers, sqlite_source_id() as srcId;")
        closedb
    end function

    '********************************************
    public function sqlite3_msi_dll_tests
        log "****************************************************************************"
        log "sqlite3_msi_dll_tests"
        REM load the base SQLiteODBC driver without any extensions
        opendb "SQL3 "
        
        dim sSql
        dim oRs,ff

        log ""
        log "SQLITE SQLITE SQLITE SQLITE SQLITE SQLITE SQLITE SQLITE SQLITE SQLITE SQLITE SQLITE "
        log "load msi.dll using load_extension()"
        log query("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\msi.dll') as msi_loaded")

        log "WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT "
        log "execute query WITHOUT msi cdbl function "
        sSql = "select (32) as lat, (-76) as lon;"
        log sSql
        set oRs = objConn.execute(sSql)
        if oRs.state = 1 then
            log "opened without cdbl (note type = 5 is dbDouble, type = 3 is long)"
            for each ff in oRs.Fields
                log "    name:" & ff.Name & " value:" & ff.Value & " type:" & ff.Type
                on error resume next
                log "    typename " & typename(ff.Value)
                if err.number <> 0 then log "    " & err.number & " " & err.description & " (with local cdbl() " & cdbl(ff.Value) & ")"
                on error goto 0
            next
        else
            log "DID NOT open without cdbl"
        end if
        log ""
        if oRs.state = 1 then
            oRs.close
        end if 
        set oRs = nothing

        log "WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH "
        log "execute query WITH msi cdbl function"
        sSql = "select cdbl(32) as lat, cdbl(-76) as lon;"
        log sSql
        set oRs = objConn.execute(sSql)
        if oRs.state = 1 then
            log "opened with cdbl (note type = 5 is dbDouble, type = 3 is long)"
            for each ff in oRs.Fields
                log "    name:" & ff.Name & " value:" & ff.Value & " type:" & ff.Type
                on error resume next
                log "    typename " & typename(ff.Value)
                if err.number <> 0 then log "    " & err.number & " " & err.description & " (with local cdbl() " & cdbl(ff.Value) & ")"
                on error goto 0
            next
        else
            log "DID NOT open with cdbl"
        end if
        log ""
        
        if oRs.state = 1 then
            oRs.close
        end if 
        set oRs = nothing


        
        closedb

    end function

    '********************************************
    public function access_msi_dll_tests
        log "****************************************************************************"
        log "access_msi_dll_tests"
        opendb "JET4 "
        dim sSql
        dim oRs,ff

        log "JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 JET4 "

        log "WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT WITHOUT "
        log "execute query WITHOUT using builtin cdbl"
        sSql = "select (32) as lat, (-76) as lon"
        log sSql
        set oRs = objConn.execute(sSql)
        if oRs.state = 1 then
            log "opened without cdbl (note type = 5 is dbDouble, type = 3 is long, type = 131 is dbNumeric)"
            for each ff in oRs.Fields
                log "    name:" & ff.Name & " value:" & ff.Value & " type:" & ff.Type
                on error resume next
                log "    typename " & typename(ff.Value)
                if err.number <> 0 then log "    " & err.number & " " & err.description & " (with local cdbl() " & cdbl(ff.Value) & ")"
                on error goto 0
            next
        else
            log "DID NOT open without cdbl"
        end if
        log ""
        if oRs.state = 1 then
            oRs.close
        end if 
        set oRs = nothing

        log "WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH WITH "
        log "execute query WITH builtin cdbl"
        sSql = "select cdbl(32) as lat, cdbl(-76) as lon"
        log sSql
        set oRs = objConn.execute(sSql)
        if oRs.state = 1 then
            log "opened with cdbl (note type = 5 is dbDouble, type = 3 is long, type = 131 is dbNumeric)"
            for each ff in oRs.Fields
                log "    name:" & ff.Name & " value:" & ff.Value & " type:" & ff.Type
                on error resume next
                log "    typename " & typename(ff.Value)
                if err.number <> 0 then log "    " & err.number & " " & err.description & " (with local cdbl() " & cdbl(ff.Value) & ")"
                on error goto 0
            next
        else
            log "DID NOT open with cdbl"
        end if
        log ""
        
        if oRs.state = 1 then
            oRs.close
        end if 
        set oRs = nothing

        closedb
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
        
        log vbcrlf & "sqlite3_fts_tests"
        ' Create an FTS table
        query("DROP TABLE IF EXISTS pages;")
        query("CREATE VIRTUAL TABLE pages USING fts4(title, body);")
        ' Insert a row with a specific docid value.
        query("INSERT INTO pages(docid, title, body) VALUES(53, 'Home Page', 'SQLite is a software...');")
        ' Insert a row and allow FTS to assign a docid value using the same algorithm as
        ' SQLite uses for ordinary tables. In this case the new docid will be 54,
        ' one greater than the largest docid currently present in the table.
        query("INSERT INTO pages(title, body) VALUES('Download', 'All SQLite source code...');")
        query("INSERT INTO pages(title, body) VALUES('Upload', 'Upload SQLite src code...');")
        ' Example full-text-search queries
        log query("SELECT * FROM pages WHERE pages MATCH 'sqlite';")
        log query("SELECT * FROM pages WHERE pages MATCH 's* code';")

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
        opendb "SQL3-LoadExt-ExtFun"
        log "extension-functions.dll loaded, now missing functions work, and ExtFun log(x) is ln(x), but that is not what is happening..."
        sqlite_math_functions
        closedb
        opendb "SQL3 "
        log "now trying load_extension() to load extension-functions.dll...expect  log(x) is base 10 as"
        log "load_extension() does not support function overloading like loading from connection string does"
        logResult query2csv("SELECT load_extension('M:\Apps\SQLiteODBC\install\" & sBitPath & "\extension-functions.dll') as ext_loaded")
        log "missing functions work,  but log(x) is base 10"
        sqlite_math_functions
        closedb

    end function

    '********************************************
    public function sqlite_math_functions
        log "****************************************************************************"
        log "sqlite_math_functions"
        log query("select acos(30*3.14159/180) as x;")
        ' acosh(X)
        log query("select acosh(30*3.14159/180) as x;")
        ' asin(X)
        log query("select asin(30*3.14159/180) as x;")
        ' asinh(X)
        log query("select asinh(30*3.14159/180) as x;")
        ' atan(X)
        log query("select atan(30*3.14159/180) as x;")
        ' atanh(X)
        log query("select atanh(30*3.14159/180) as x;")
        ' atan2(X,Y)
        log query("select atan2(4,5) as x;")
        ' atn2(X,Y) FAILS function not available
        log query("select atn2(4,5) as x;")
        ' ceil(X)
        log query("select ceil(1.1) as x;")
        ' ceiling(X)
        log query("select ceiling(1.1) as x;")
        ' cos(X)
        log query("select cos(30*3.14159/180) as x;")
        ' cosh(X)
        log query("select cosh(30*3.14159/180) as x;")
        ' cot(X) FAILS function not available
        log query("select cot(30*3.14159/180) as x;")
        ' coth(X) FAILS function not available
        log query("select coth(30*3.14159/180) as x;")
        ' degrees(X)
        log query("select degrees(30*3.14159/180) as x;")
        ' exp(X)
        log query("select exp(2) as x;")
        ' floor(X)
        log query("select floor(1.1) as w, floor(1.9) as x, floor(-1.1) as y, floor(-1.9) as z;")
        ' ln(X)
        log query("select ln(10) as x;")
        ' log(B,X)
        log query("select log(2,10) as x,log(10,10) as y,log(16,10) as z;")
        ' log(X)
        log ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
        log query("select log(10) as x;")
        log "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
        ' log10(X)
        log query("select log10(10) as x;")
        ' log2(X)
        log query("select log2(10) as x;")
        ' mod(X,Y)
        log query("select mod(10,10) as x, mod(10,11) as y;")
        ' pi()
        log query("select pi() as x;")
        ' pow(X,Y)
        log query("select pow(2,2) as x;")
        ' power(X,Y)
        log query("select power(2,2) as x;")
        ' radians(X)
        log query("select radians(180) as x;")
        ' sign(X)
        log query("select sign(-10) as x, sign(10) as y;")
        ' sin(X)
        log query("select sin(30*3.14159/180) as x;")
        ' sinh(X)
        log query("select sinh(30*3.14159/180) as x;")
        ' sqrt(X)
        log query("select sqrt(4) as x;")
        ' square(X) FAILS function not available
        log query("select square(4) as x;")
        ' tan(X)
        log query("select tan(30*3.14159/180) as x;")
        ' tanh(X)
        log query("select tanh(30*3.14159/180) as x;")
        ' trunc(X)
        log query("select trunc(1.123) as x, trunc(1.9) as y;")

    end function

    '********************************************
    public function sqlite3_feature_tests
        log "****************************************************************************"
        log "sqlite3_feature_tests"
        log "test some other more recent features of sqlite or SQL used by sqlite"
        
        log "iif() is now included in SQLite SQL language"
        opendb "SQL3 "
        log query("select iif(1=2,'true','false') as w, iif(2=2,'true','false') as x, iif('hello' = 'world','true','false') as y, iif('same' = 'same','true','false') as z;")
        closedb
        
        log "'alter table drop column' is now included in SQLite SQL language"
        opendb "SQL3 "
        log query("drop table if exists cxb_copy")
        log query("create table cxb_copy as select * from test_table")
        log query("select count(1) from cxb_copy")
        log query("alter table cxb_copy drop column cxb_text")
        log query("select * from cxb_copy limit 1")
        log query("drop table if exists cxb_copy")
        closedb
        
    end function

    '********************************************
    public function sqlite3Tests
        log "****************************************************************************"
        log "sqlite3Tests"
        opendb "SQL3 "
        dim conStr
        
        log "sqlite doesn't do connection strings in SQL...see next two examples"
        log ""
        log "try to read from access DB via ODBC using a SQLite3 connection object"
        on error resume next
        conStr = "ODBC;DSN=MS Access Database;Database=" & dbAccess & ";"
        objConn.execute "select count(1) as cnt from [" & conStr & "].[test_table]"
        if err.number <> 0 then
            log err.number & " " & err.description & vbcrlf
        end if
        on error goto 0

        log "try to read from access DB via COM using a SQLite3 connection object [FAILS]"
        on error resume next
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbAccess & ";"
        objConn.execute "select count(1) as cnt from [" & conStr & "].[test_table]"
        if err.number <> 0 then
            log err.number & " " & err.description & vbcrlf
        end if
        on error goto 0
        
        log "get list of tables in database file same as .tables dot command in sqlite3.exe"
        log query("select name from sqlite_master where type = 'table' AND name NOT LIKE 'sqlite_%';")
        
        log "sqlite version of 'select * into t from...' is 'create table t as select * from...'"
        log query("drop table if exists [renameMe]")
        log query("create table [renameMe] as select * from test_table") 
        log query("select * from [renameMe] limit 3")
        
        log "alter table rename"
        log query("drop table if exists [cxb]")
        log query("alter table renameMe rename to [cxb];")
        log query("select * from cxb limit 3;")
        
        log "not case sensitive..."
        log query("select * from Cxb limit 3;")
        
        log "alter table add column"
        log query("alter table [cxb] add column cxb_text text;")
        log query("update [cxb] set cxb_text ='cxb';")
        log query("select * from cxb limit 3;")

        log "vacuum which is similar to compact/repair"
        log "VACUUM " & query("vacuum;")
        
        log "vacuum into a new db file, deleting the file first if it already exists"
        if objFSO.FileExists("vacuum.db3") then objFSO.DeleteFile("vacuum.db3")
        log "VACUUM into " & query("vacuum into 'vacuum.db3';")

        log "attach to currently opened db give it the name db2"
        log "attach '" & dbSqlite3 & "' as db2;"
        objConn.execute "attach '" & dbSqlite3 & "' as db2;"
        
        log "get count of records using attached db2 in table using query() that uses local oRs"
        log query("select count(1) as cnt from db2.test_table;")
        
        log "use the global objRS instead of the local query oRs to read attached db2"
        log openRecordSet("select * from db2.test_table")
        
        log "detach db2 and verify it worked"
        log "before detach " & query("select count(1) as cnt from db2.test_table;")
        objConn.execute "detach database 'db2';"
        log "after detach expect an error --> " & query("select count(1) as cnt from db2.test_table;")

        ' close global objRS
        if objRS.state = 1 then objRS.close
        closedb
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
    public function accessOdbcTests
        log "****************************************************************************"
        opendb "ODBC "
        log query("select top 3 MYFIELD_1, * from test_table")
        closedb
        
        opendb "JET4 "
        log query("select top 3 * from [ODBC;Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq='" & dbAccess & "';].[test_table]")
        closedb
    end function

    '********************************************
    public function accessTests
        log "****************************************************************************"
        log "accessTests"
        
        opendb "JET4 "
        log "note that SQL here must be valid Access SQL even though we are working on SQLite DB"
        log query("SELECT TOP 3 * FROM [ODBC;DSN=SQLite3 Datasource;Database=" & dbSqlite3 & ";].[test_table]")
        log "compare case sensitive query"
        log query("SELECT TOP 3 * FROM [ODBC;DSN=SQLite3 Datasource;Database=" & dbSqlite3 & ";].[cxb]")
        log query("SELECT TOP 3 * FROM [ODBC;DSN=SQLite3 Datasource;Database=" & dbSqlite3 & ";].[Cxb]")
        log "this one has no result as it is creating a new table"
        log query("SELECT top 3 *,CDbl(0) as newDoubleField,cdate(#5/6/2019#) as newDateField FROM test_table")
        closedb
        
        opendb "SQL3 "
        log query("drop table if exists [fromAccess]")
        closedb
        
        opendb "JET4 "
        log query("SELECT *,CDbl(0) as newDoubleField,cdate(#5/6/2019#) as newDateField INTO [ODBC;DSN=SQLite3 Datasource;Database=" & dbSqlite3 & ";].[fromAccess] FROM test_table")
        log query("SELECT TOP 3 cstr(myField_1) AS val1,newDateField,newDoubleField FROM [ODBC;DSN=SQLite3 Datasource;Database=" & dbSqlite3 & ";].[fromAccess]")
        closedb
    end function

    '********************************************
    function dumpPragma
        log "****************************************************************************"
        log "dumpPragma"
        opendb "MEM  "
        dim ss: ss = ""
        ss = ss & query("PRAGMA analysis_limit")
        ss = ss & query("PRAGMA application_id")
        ss = ss & query("PRAGMA auto_vacuum")
        ss = ss & query("PRAGMA automatic_index")
        ss = ss & query("PRAGMA busy_timeout")
        ss = ss & query("PRAGMA cache_size")
        ss = ss & query("PRAGMA cache_spill")
        ss = ss & query("PRAGMA case_sensitive_like")
        ss = ss & query("PRAGMA cell_size_check")
        ss = ss & query("PRAGMA checkpoint_fullfsync")
        ss = ss & query("PRAGMA collation_list")
        ss = ss & query("PRAGMA compile_options")
        ss = ss & query("PRAGMA data_version")
        ss = ss & query("PRAGMA database_list")
        ss = ss & query("PRAGMA encoding")
        
        ' sorts the list so that comparing is easier
        ss = ss & getSortedFunctionList("") & vbcrlf
        
        ss = ss & query("PRAGMA integrity_check")
        ss = ss & query("PRAGMA hard_heap_limit")
        ss = ss & query("PRAGMA journal_mode")
        ss = ss & query("PRAGMA journal_size_limit")
        ss = ss & query("PRAGMA locking_mode")
        ss = ss & query("PRAGMA max_page_count")
        ss = ss & query("PRAGMA mmap_size")
        ss = ss & query("PRAGMA optimize")
        ss = ss & query("PRAGMA page_size")
        ss = ss & query("PRAGMA page_count")
        ss = ss & query("PRAGMA pragma_list")
        ss = ss & query("PRAGMA quick_check")
        ss = ss & query("PRAGMA shrink_memory")
        ss = ss & query("PRAGMA soft_heap_limit")
        ss = ss & query("PRAGMA synchronous")
        ss = ss & query("PRAGMA table_info('test_table')")
        ss = ss & query("PRAGMA table_xinfo('test_table')")
        ss = ss & query("PRAGMA temp_store")
        ss = ss & query("PRAGMA temp_store_directory")
        ss = ss & query("PRAGMA threads")
        ss = ss & query("PRAGMA trusted_schema")
        ss = ss & query("PRAGMA user_version")
        ss = ss & query("PRAGMA wal_autocheckpoint")
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
    function OpenAccess(currentDb)
        dim oWsh: Set oWsh = CreateObject("WScript.Shell")
        on error resume next
        Dim RegKey: RegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE\Path"
        Dim MSAccPath: MSAccPath = oWsh.RegRead(RegKey)
        dim sCurrentPath: sCurrentPath = oWsh.CurrentDirectory
        oWsh.CurrentDirectory = MSAccPath
        dim sCmd
        if len(currentDb) > 0 then
            sCmd = """" & "MSACCESS.EXE" & """ """ & currentDb & """"
        else
            sCmd = """" & "MSACCESS.EXE"  & """"
        end if
        dim iReturn: iReturn = oWsh.Run(sCmd,7,false)
        if appDbViewer_LOGGING then console.log "OpenAccess: " & iReturn & " path " & MSAccPath & " cmd " & sCmd
        oWsh.CurrentDirectory = sCurrentPath
        on error goto 0
    end function

    '********************************************
    ' perform an array of tests to evaluate insert performance 
    ' given different settings
    public function insertTests(bAccess,bSqlite)
        
        REM journal_mode = DELETE | TRUNCATE | PERSIST | MEMORY | WAL | OFF
        REM page_size = Query or set the page size of the database. The page size must be a power of two between 512 and 65536 inclusive.
        REM synchronous = 0 | OFF | 1 | NORMAL | 2 | FULL | 3 | EXTRA;
        REM temp_store = 0 | DEFAULT | 1 | FILE | 2 | MEMORY;
        REM locking_mode = NORMAL | EXCLUSIVE

        REM dim cache_size:     cache_size =        Array(-2000,-4000,-8000)                                        ' -2000 (default)
        REM dim journal_mode:   journal_mode =      Array("WAL","MEMORY","OFF","DELETE","TRUNCATE","PERSIST")       ' OFF, WAL (DELETE default)
        REM dim page_size:      page_size =         Array(4096,4096*2,4096*4,4096*8)                                ' 4096 (default)
        REM dim synchronous:    synchronous =       Array("OFF","NORMAL","FULL","EXTRA")                            ' NORMAL (default)
        REM dim temp_store:     temp_store =        Array("MEMORY","FILE")                                          ' FILE
        REM dim locking_mode:   locking_mode =      Array("EXCLUSIVE","NORMAL")                                     ' NORMAL (default)
        
        ' these are the best performers
        dim cache_size:     cache_size =        Array(-2000)
        dim journal_mode:   journal_mode =      Array("WAL","OFF")
        dim page_size:      page_size =         Array(4096)
        dim synchronous:    synchronous =       Array("NORMAL")
        dim temp_store:     temp_store =        Array("FILE")
        dim locking_mode:   locking_mode =      Array("NORMAL")

        dim apg: apg = Array()
        dim cs,jm,ps,s,ts,lm
        
        for each ts in temp_store
            for each ps in page_size
                for each s in synchronous
                    for each lm in locking_mode
                        for each jm in journal_mode
                            for each cs in cache_size
                                redim preserve apg(ubound(apg)+1)
                                apg(ubound(apg)) = "cache_size="&cs&"|journal_mode="&jm&"|page_size="&ps&"|synchronous="&s&"|temp_store="&ts&"|locking_mode="&lm
                            next
                        next
                    next
                next
            next
        next
        
        dim aPrimaryKey ' set below in loop
        
        dim aInsertsPerTransaction: aInsertsPerTransaction = Array(50000)
        REM dim aInsertsPerTransaction: aInsertsPerTransaction = Array(50000,100000)
        REM dim aInsertsPerTransaction: aInsertsPerTransaction = Array(5000,10000,25000,50000,100000)
        
        REM dim aDbTypes: aDbTypes = Array("JET4 ","ACE12","ACE15","SQL3 ")
        dim aDbTypes
        if bAccess and bSqlite then
            aDbTypes = Array("JET4 ","SQL3 ")
        elseif bAccess then
            aDbTypes = Array("JET4 ")
        elseif bSqlite then
            aDbTypes = Array("SQL3 ")
        end if
        
        REM dim types: types = Array("INTEGER","REAL","TEXT")
        REM dim types: types = Array("INTEGER","TEXT")
        dim types: types = Array("INTEGER")
        
        if true then
            dim t: for each t in types
                dim r: for r = 100000 to 100000 step 100000
                    dim c: for c = 5 to 5 step 100
                        dim p: for each p in aDbTypes
                            if instr(p,"SQL") > 0 then 
                                aPrimaryKey = Array(true,false)
                            else
                                aPrimaryKey = Array(false)
                            end if
                            dim pk: for each pk in aPrimaryKey
                                dim ipt: for each ipt in aInsertsPerTransaction
                                    if instr(p,"SQL") > 0 then 
                                        dim pg: for each pg in apg
                                            REM log r & " " & c & " " & t & " " & p & " " & ipt & " " & pk
                                            test r,c,t,p,ipt,pk,pg
                                        next
                                    else
                                        REM log r & " " & c & " " & t & " " & p & " " & ipt & " " & pk
                                        test r,c,t,p,ipt,pk,""
                                    end if
                                next
                            next
                        next
                    next
                next
            next
        end if
    end function

    '********************************************
    ' r = number of rows to insert
    ' c = number of columns per row
    ' t = column data type (INTEGER, REAL, TEXT)
    ' p = which driver to use (SQL3 or 
    ' ipt = number of records per transaction
    ' pk = use primary key (true/false)
    ' pg = pipe "|" separated pragma string
    function test(r,c,t,p,ipt,pk,pg)
        dim fso: Set fso = CreateObject("Scripting.FileSystemObject") 
        ' set pragma values for SQLite3
        dim pragmaHeader: pragmaHeader = ""
        dim pragma: pragma = ""
        if p = "SQL3 " then
            if len(pg) > 0 then
                dim aa: aa = split(pg,"|")
                dim aaa
                for each aaa in aa
                    pragmaHeader = pragmaHeader & split(aaa,"=")(0) & ","
                    pragma = pragma & split(aaa,"=")(1) & ","
                next
                dbSqlite3 = strFolder & "\testDBs\" & trim(p) & "_" & pk & "_" & r & "_" & c & "_" & ipt & "_" & left(t,1) & "_" & replace(replace(pragma,"-","N"),",","_") & iBitness & ".sqlite3"
                if objFSO.FileExists(dbSqlite3) then objFSO.DeleteFile(dbSqlite3)
                opendb p
                for each aaa in aa
                    objConn.execute "PRAGMA " & aaa & ";"
                next
            end if
        else
            if instr(p,"JET") > 0 or instr(p,"ACE") > 0 then
                if bHasAccess then
                    dbAccess =  strFolder & "\testDBs\" & trim(p) & "_" & pk & "_" & r & "_" & c & "_" & ipt & "_" & left(t,1) & iBitness & ".mdb"
                    if objFSO.FileExists(dbAccess) then objFSO.DeleteFile(dbAccess)
                        
                    OpenAccess ""
                    on error resume next
                    dim oAccess: set oAccess = GetObject( , "Access.Application")
                    do while typename(oAccess) = "Empty"
                        sleep 100
                        set oAccess = GetObject( , "Access.Application")
                    loop
                    on error goto 0
                    dim dbVersion40: dbVersion40 = 64 ' Microsoft Jet database engine version 4.0
                    if fso.FileExists(dbAccess) = false then
                        oAccess.DBEngine.CreateDatabase dbAccess,";LANGID=0x0409;CP=1252;COUNTRY=0",dbVersion40
                    end if
                    oAccess.quit
                    set oAccess = nothing

                    opendb p
                else
                    log "cannot create database file ---> access not available! exiting test method."
                    exit function
                end if
            end if
        end if
        
        if bOpenFirstTime then
            log "bit,engine,pk,rows,columns,data_type," & pragmaHeader & "AddNewRate_rps,InsertTime_s,TransactionCount,numTranacions"
            bOpenFirstTime = false
        end if


        if c > 998 then c = 998
        
        dim sPk
        if pk = true then sPk = "pk_"
        
        
        dim sTableName: sTableName = "test_table"
        on error resume next
        REM objConn.execute "drop table [" & sTableName & "]"
        on error goto 0

        dim aHeader
        dim aData
        
        dim createTable: createTable = _
            "CREATE TABLE " & sTableName & " ( "
            
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

        if typename(objConn) = "Empty" then 
            log "connection object is Empty"
            exit function
        end if
        objConn.execute createTable
        
        openRecordSet("select * from [" & sTableName & "]")
            
        dim iRecords: iRecords = r+1
        dim iTransactionCount: iTransactionCount = 0
        dim sHeader: sHeader = join(aHeader,",")
        dim sData
        dim dStart: dStart = timer
        objConn.BeginTrans
        for i = 1 to iRecords
            REM if pk = false then aData(0)= clng(i)
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
        
        dim result: result = 0
        on error resume next
        objRs.movelast
        result = (timer - dStart)
        on error goto 0
        
        if p = "SQL3 " then
            log iBitness & "," & p & "," & pk & "," & r & "," & c & "," & t & "," & pragma & (r/cdbl(result)) & "," & result & "," & ipt & "," & iTransactionCount
        else
            log iBitness & "," & p & "," & pk & "," & r & "," & c & "," & t & "," & (r/cdbl(result)) & "," & result & "," & ipt & "," & iTransactionCount
        end if
        
        closedb
        
    end function

    '********************************************
    sub opendb(p)
        Set objConn = CreateObject("ADODB.Connection")
        
        dim sConnStr

        select case p
            case "JET4 ":
                sConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbAccess & ";"
            case "ACE12":
                sConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbAccess & ";"
            case "ACE15":
                sConnStr = "Provider=Microsoft.ACE.OLEDB.15.0;Data Source=" & dbAccess & ";"
            case "SQL3 "
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";"
            case "MEM  ":
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=:memory:;"
            case "SQL3-checkfreelist"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\sha1.dll,M:\Apps\SQLiteODBC\install\" & sBitPath & "\checkfreelist.dll;"
            case "SQL3-sha"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\sha1.dll,M:\Apps\SQLiteODBC\install\" & sBitPath & "\shathree.dll;"
            case "SQL3-series"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\series.dll;"
            case "SQL3-ieee754"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\ieee754.dll;"
            case "SQL3-fileio"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\fileio.dll;"
            case "SQL3-totype"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\totype.dll;"
            case "SQL3-LoadExt-ExtFun"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\extension-functions.dll;"
            case "SQL3-LoadExt-Regexp"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\regexp.dll;"
            case "SQL3-LoadExt-Csv"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\csv.dll;"
            case "SQL3-tracefile"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";tracefile=C:\Users\charlie\Desktop\SQLite3ODBC_trace.txt;"
            case "SQL3-vfsstat"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\vfsstat.dll;"
            case "SQL3-uuid"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\uuid.dll;"
            case "SQL3-crypto"
                sConnStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbSqlite3 & ";LoadExt=M:\Apps\SQLiteODBC\install\" & sBitPath & "\crypto.dll;"
            case "ODBC "
                sConnStr = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & dbAccess & ";"
            case "DUCKDB"
                sConnStr = "Driver=DuckDB Driver;Database=" & dbDuck & ";"
            case "DUCKDB-DSN"
                sConnStr = "DSN=" & dbDuckDSN & ";"
            case "DUCKDB-memory"
                sConnStr = "Driver=DuckDB Driver;"
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
        query2csv = -1
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
                                aQueryResults(3) = aQueryResults(3) & "1:" & err.number & ":" & err.description & "|"
                                err.clear
                            end if
                            on error goto 0
                        next
                        on error resume next
                        oRs.MoveNext
                        if err.number <> 0 then
                            aQueryResults(3) = aQueryResults(3) & "2 " & err.number & ":" & err.description & "|"
                        end if
                        on error goto 0
                        ' remove the last separator 
                        ss = left(ss,len(ss)-1)
                        aQueryResults(2).add rowCount, ss
                        query2csv = rowCount
                        rowCount = rowCount + 1
                    loop
                else
                    aQueryResults(3) = "RS contains no records" & "|"
                    oRs.close
                    query2csv = -1
                    exit function
                end if
            else
                aQueryResults(3) = "RS is not open " & oRs.state & "|"
                query2csv = -1
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
            log "return " & (r+1) & " rows"
            if len(aQueryResults(3)) > 0  then log aQueryResults(3)
            log aQueryResults(1)
            dim vKey: for each vKey in aQueryResults(2)
                log aQueryResults(2).item(vKey)
            next
            log ""
        else
            log "QUERY  " & aQueryResults(0)
            log "return no rows (" & r & ")"
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