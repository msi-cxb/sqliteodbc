[![CI](https://github.com/msi-cxb/sqliteodbc/actions/workflows/CI.yml/badge.svg)](https://github.com/msi-cxb/sqliteodbc/actions/workflows/CI.yml) [![sqlitte_version](https://github.com/msi-cxb/sqliteodbc/actions/workflows/sqlite_version.yml/badge.svg)](https://github.com/msi-cxb/sqliteodbc/actions/workflows/sqlite_version.yml)

Original README is [here](https://github.com/msi-cxb/sqliteodbc/blob/master/README).

This is the current version of SQLite3 used in the driver (values pulled automatically from sqlite3.c):

SQLITE_VERSION: 3.47.1

SQLITE_SOURCE_ID: 2024-10-21 16:30:22 03a9703e27c44437c39363d0baf82db4ebc94538a0f28411c85dda156f82636e

This is a fork of https://github.com/softace/sqliteodbc with modifications to build with Visual Studio 2015/2017/2019/2022 for both 32 and 64 bit. 

Features
- Builds in Visual Studio using nmake
- 32 and 64 bit builds work
- It includes a batch script [buildAndInstall.bat](https://github.com/msi-cxb/sqliteodbc/blob/master/buildAndInstall.bat) to build and install on Windows (10 and 11 tested).

TODO
- build and test both 32 and 64 bit drivers in CI
- possibly use cMake to build instead of nmake with .mak file
- figure out how to run [sqllogictest](https://sqlite.org/sqllogictest/info/trunk) tests in CI (I've got a VBScript that can do this but test files are huge)




