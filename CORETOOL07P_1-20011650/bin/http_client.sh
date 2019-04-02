#!/bin/sh
#
# (c) Copyright 2001-2009 SAP AG
# All rights reserved.
#
# This script is used to launch the HTTP Client program of SAP Convergent Charging.
#

HIGHDEAL_HOME=`/usr/bin/dirname $0`/..
export HIGHDEAL_HOME

# use old HIGHDEAL_JAVA_HOME variable if SAPCC_JAVA_HOME variable is not set
if [ "$HIGHDEAL_JAVA_HOME" != "" -a "$SAPCC_JAVA_HOME" = "" ] ; then
  echo Warning: using legacy environment variable HIGHDEAL_JAVA_HOME, please use SAPCC_JAVA_HOME instead.
  export SAPCC_JAVA_HOME=$HIGHDEAL_JAVA_HOME
fi

# check SAPCC_JAVA_HOME is set
if [ "$SAPCC_JAVA_HOME" = "" ] ; then
  echo You must set SAPCC_JAVA_HOME to point at your Java Development Kit directory
  echo Hit any key to continue...
  read
  exit 1
fi

if [ ! -f $SAPCC_JAVA_HOME/bin/java ] ; then
  echo "java not present in '$SAPCC_JAVA_HOME/bin'."
  echo Make sure the SAPCC_JAVA_HOME environment variable is set to the path of your Java root directory.
  echo Hit any key to continue...
  read
  exit 1
fi

$SAPCC_JAVA_HOME/bin/java -version > javaversion 2>&1
grep "1.6" javaversion > /dev/null
status=$?;
if test $status -eq 1; then
    echo The java specification version of the jvm used is not 1.6.
    echo Make sure the SAPCC_JAVA_HOME environment variable is set to the path of your SAP JVM 6 directory.
    rm javaversion
    read
    exit 1
fi

grep "SAP" javaversion > /dev/null
status=$?;
if test $status -eq 1; then
    echo The java vendor of the jvm used is not SAP AG.
    echo Make sure the SAPCC_JAVA_HOME environment variable is set to the path of your SAP JVM 6 directory.
    rm javaversion
    read
    exit 1
fi
rm javaversion

HIGHDEAL_LIB=$HIGHDEAL_HOME/jars
HIGHDEALCLASSES=$HIGHDEAL_LIB/core_client.jar:$HIGHDEAL_LIB/logging.jar:$HIGHDEAL_LIB/sap.com~tc~logging~java.jar:$HIGHDEAL_LIB/core_chargingplan.jar:$HIGHDEAL_LIB/core_chargingprocess.jar:$HIGHDEAL_LIB/common_message.jar:$HIGHDEAL_LIB/common_util.jar

CLASSPATH=$HIGHDEALCLASSES

JAVA="$SAPCC_JAVA_HOME/bin/java"

exec $JAVA -Dfile.encoding=UTF-8 -classpath $CLASSPATH com.highdeal.httpclient.HttpClient $*


