[![CI](https://github.com/msi-cxb/sqliteodbc/actions/workflows/CI.yml/badge.svg)](https://github.com/msi-cxb/sqliteodbc/actions/workflows/CI.yml) [![sqlitte_version](https://github.com/msi-cxb/sqliteodbc/actions/workflows/sqlite_version.yml/badge.svg)](https://github.com/msi-cxb/sqliteodbc/actions/workflows/sqlite_version.yml)

Original Christian Werner/SoftACE README is [here](https://github.com/msi-cxb/sqliteodbc/blob/master/README).

This is a fork of https://github.com/softace/sqliteodbc with modifications to build with Visual Studio 2015/2017/2019/2022 for both 32 and 64 bit. Also adds CI workflow that builds the ODBC driver and sqllogictest and runs the full set of tests. 

This is the current version of SQLite3 used in the driver (values pulled automatically from sqlite3.c):
- SQLITE_VERSION: 3.50.0
- SQLITE_SOURCE_ID: 2025-05-29 14:26:00 dfc790f998f450d9c35e3ba1c8c89c17466cb559f87b0239e4aab9d34e28f742

Features
- Builds with Visual Studio (2015/2017/2019/2022) using nmake
- 32 and 64 bit builds work and are tested by CI
- It includes a batch script [buildAndInstall.bat](https://github.com/msi-cxb/sqliteodbc/blob/master/buildAndInstall.bat) to build and install as admin on Windows (10 and 11 tested).
    - `buildAndInstall.bat [arch: 32 or 64] [install: true or false]`
- tested with SQLite3 developers `sqllogictest` full suite of tests

TODO
- possibly use cMake to build instead of nmake with .mak file




