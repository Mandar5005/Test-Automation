
Steps for running the new iTAF framework:


1. Unzip the Fall'25-Update4-dist.zip file at any location in your P/C from where you want to execute iTAF.
2. Copy your test suite in the same directory where you have extracted the zip file.
3. Edit startITAF.bat for providing the Config.xls path as in your test suite in command line argument.
4. You can provide APPLICATIO_NAME in one of the below 2 possible ways.
	
	a. Add APPLICATION_NAME (Billing/Claims/PAS/DM/LNA/DevStudio(DS)) as second command line argument as per instruction provided inside startITAF.bat.
	
	For Example in case of Claims application:
	java -cp ".\libs\log4j-core-2.19.0.jar;.\libs\log4j-api-2.19.0.jar;.\libs\*;" -Dlog4j2.configurationFile=log4j2.xml 
	com.majesco.itaf.main.ITAFWebDriver  Claim\CommonResources\Config.xls Claims
	
	OR
	b. Add parameter APPLICATION_NAME in Config.xls with value as Billing/Claims/PAS/DM/LNA/DS as applicable in Config.xml.

5. Run the startITAF.bat.
6. Make sure the log4j2.xml is included in the src/main/resources folder along with the Log4j.properties file.