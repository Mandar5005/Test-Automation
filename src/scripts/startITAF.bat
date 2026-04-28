@echo off
setlocal ENABLEDELAYEDEXPANSION
rem #Windows batch file that helps to boot the engine.
rem #Provided without warranty.

rem #Include the path of the JDK installed.
rem #set JAVA_HOME=D:\java\64bit\jdk1.8.0_05

if not defined JAVA_HOME  (
   echo Edit this batch file and set the appropriate JAVA_HOME to JDK 1.8 installed on this machine
   goto :EXIT
)
rem #Set the classpath here
set cp=.\libs\*;%CP%

rem #Check the classpath here
echo cp=%CP%

rem #Pass the log4j-core-2x,log4j-api-2x version, libs path and the webdriver along with the project path and run parametersm(billing, Claims, PAS, DM) as given below. 
java -cp ".\libs\log4j-core-2.19.0.jar;.\libs\log4j-api-2.19.0.jar;.\libs\*;" -Dlog4j2.configurationFile=log4j2.xml com.majesco.itaf.main.ITAFWebDriver Claim\CommonResources\Config.xls Claims


:EXIT
