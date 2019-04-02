@echo off
REM
REM (c) Copyright 2001-2009 SAP AG
REM All rights reserved.
REM
REM This script is used to launch the Core Tool of SAP Convergent Charging.
REM

IF ".%SAPCC_JAVA_HOME%" == "." (
  IF NOT ".%HIGHDEAL_JAVA_HOME%" == "." (
    ECHO Warning: using legacy environment variable HIGHDEAL_JAVA_HOME, please use SAPCC_JAVA_HOME instead.
    SET SAPCC_JAVA_HOME=%HIGHDEAL_JAVA_HOME%
  )
)

IF ".%SAPCC_JAVA_HOME%" == "." GOTO :javahome

if Windows_NT == %OS% goto :ntStart

echo "SAP Convergent Charging can run only on Windows NT/2000. Sorry."
goto :end

:ntStart
setlocal

IF NOT EXIST "%SAPCC_JAVA_HOME%/bin/java.exe" goto :javaPresent

"%SAPCC_JAVA_HOME%/bin/java.exe" -version 2>&1 | find "1.6" > nul
IF errorlevel 1 (
    echo The java specification version of the jvm used is not 1.6.
    echo Make sure the SAPCC_JAVA_HOME environment variable is set to the path of your SAP JVM 6 directory.
    del javaversion
    PAUSE
    goto :end
)

"%SAPCC_JAVA_HOME%/bin/java.exe" -version 2>&1 | find "SAP" > nul
IF errorlevel 1 (
    echo The java vendor of the jvm used is not SAP AG.
    echo Make sure the SAPCC_JAVA_HOME environment variable is set to the path of your SAP JVM 6 directory.
    del javaversion
    PAUSE
    goto :end
)

SET HIGHDEAL_HOME=..

SET HIGHDEAL_LIB=%HIGHDEAL_HOME%/jars
SET HIGHDEALCLASSES=%HIGHDEAL_LIB%/core_client.jar;%HIGHDEAL_LIB%/logging.jar;%HIGHDEAL_LIB%/sap.com~tc~logging~java.jar;%HIGHDEAL_LIB%/core_chargingplan.jar;%HIGHDEAL_LIB%/core_chargingprocess.jar;%HIGHDEAL_LIB%/core_tool_i18n.jar;%HIGHDEAL_LIB%/core_tool.jar;%HIGHDEAL_LIB%/common_message.jar;%HIGHDEAL_LIB%/common_util.jar
set CONFIG_FILE=%HIGHDEAL_HOME%/config/core_tool.config

SET JAVA=%SAPCC_JAVA_HOME%/bin/javaw

For /f "tokens=1,2,* delims=\=" %%a in (%CONFIG_FILE%) do IF "%%a" == "xml.verbose" IF "%%b" == "true" SET JAVA=%SAPCC_JAVA_HOME%/bin/java

SET CLASSPATH=%HIGHDEALCLASSES%

:: Clear TITLE variable
SET TITLE=
:: Set TITLE to "" only in Windows NT/2000/XP
ECHO/>NUL&SET TITLE=""
:: Cross-platform start with "quoted" path

start %TITLE% "%JAVA%" -Dhd.config=%CONFIG_FILE% -Dfile.encoding=UTF-8 -classpath %CLASSPATH% com.highdeal.uilauncher.PnRUILauncher %1 %2 %3 %4 %5 %6 %7 %8 %9
endLocal
GOTO end

:javahome
ECHO You must set the SAPCC_JAVA_HOME environment variable to the path of your Java root directory.
PAUSE
GOTO end

:javaPresent
ECHO java.exe not present in '%SAPCC_JAVA_HOME%\bin'.
ECHO Make sure the SAPCC_JAVA_HOME environment variable is set to the path of your Java root directory.
PAUSE
GOTO end

:end

