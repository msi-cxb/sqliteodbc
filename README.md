[![CI](https://github.com/msi-cxb/sqliteodbc/actions/workflows/CI.yml/badge.svg)](https://github.com/msi-cxb/sqliteodbc/actions/workflows/CI.yml) [![sqlitte_version](https://github.com/msi-cxb/sqliteodbc/actions/workflows/sqlite_version.yml/badge.svg)](https://github.com/msi-cxb/sqliteodbc/actions/workflows/sqlite_version.yml)

Original README is [here](https://github.com/msi-cxb/sqliteodbc/blob/master/README).

This is a fork of https://github.com/softace/sqliteodbc with modifications to build with Visual Studio 2015/2017/2019/2022 for both 32 and 64 bit. 

Features
- Builds in Visual Studio using nmake
- 32 and 64 bit builds work
- It includes a batch script [buildAndInstall.bat](https://github.com/msi-cxb/sqliteodbc/blob/master/buildAndInstall.bat) to build and install on Windows (10 and 11 tested).

TODO
- build and test both 32 and 64 bit drivers in CI
- possibly use cMake to build instead of nmake with .mak file
- figure out how to run [sqllogictest](https://sqlite.org/sqllogictest/info/trunk) tests in CI (I've got a VBScript that can do this but test files are huge)




