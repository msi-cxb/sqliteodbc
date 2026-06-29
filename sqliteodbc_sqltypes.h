/*
 * sqliteodbc_sqltypes.h — thin wrapper over the Windows SDK sqltypes.h
 *
 * Adds SQLINTEGER32/SQLUINTEGER32 aliases needed for TIMESTAMP_STRUCT.fraction,
 * which must remain 32 bits on both 32-bit and 64-bit builds per the ODBC spec.
 *
 * DO NOT redefine SQLINTEGER or SQLUINTEGER here. Those types are defined by
 * the ODBC standard as 32-bit (long / unsigned long) on all platforms.
 * SQLLEN is the correct 64-bit signed length type on _WIN64.
 */
#pragma once

#include <sqltypes.h>   /* Windows SDK — provides SQLINTEGER, SQLUINTEGER, SQLLEN, etc. */

/* 32-bit integer aliases for struct fields that must not widen on _WIN64 */
typedef long            SQLINTEGER32;
typedef unsigned long   SQLUINTEGER32;
