package com.majesco.itaf.rest.utils;

import io.restassured.response.Response;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Map.Entry;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.jayway.jsonpath.JsonPath;
import com.jayway.jsonpath.PathNotFoundException;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.main.WebHelper;
import com.majesco.itaf.main.WebHelperUtil;

import com.fasterxml.jackson.core.type.TypeReference;
import java.util.Collections;


public class JsonUtility {
	static String tagValue = null;
	static MainController controller = ObjectFactory.getMainController();
	private final static Logger log = LogManager.getLogger(JsonUtility.class.getName());
	public static List<String> jsonValueList = new ArrayList<String>();
	public static List<String> pathAsArray = null;
	public static LinkedHashMap<String, String> conditionsMap = null;
	public static boolean pathHasConditionFlag = false;
	public static boolean conditionsCheckedFlag = false;

	@SuppressWarnings("unchecked")
	private static void compareJson(Map<String, Object> holdingMap, Map<String, Object> clientMap, StringBuilder log,
			Map<String, Map<String, String>> resultMap) {
		Set<String> holdingMapKeySet = holdingMap.keySet();
		Set<String> clientMapKeySet = null;
		for (String holdingMapKey : holdingMapKeySet) {
			Object holdingMapKeyObject = holdingMap.get(holdingMapKey);
			if (holdingMapKeyObject != null && holdingMapKeyObject instanceof Map) {
				if (clientMap.containsKey(holdingMapKey)) {
					compareJson((Map<String, Object>) holdingMapKeyObject,
							(Map<String, Object>) clientMap.get(holdingMapKey), log, resultMap);
				} else {
					// put all holding map keys in log
					Set<String> childMapKeys = ((Map<String, Object>) holdingMapKeyObject).keySet();
					for (String childMapKey : childMapKeys) {
						log.append("Missing key " + childMapKey + "\n");
						resultMap.put(childMapKey, null);
					}
				}
			} else {
				Map<String, String> diffMap = new HashMap<String, String>();
				if (holdingMapKeyObject instanceof String) {
					String act = (String) holdingMapKeyObject.toString().replaceAll("\"", "");
					String exp = clientMap.get(holdingMapKey).toString().replaceAll("\"", "");
					if (CommonUtils.compareStrings(act, exp, true) != 0) {
						log.append("Key: " + holdingMapKey + "\texpected value: " + holdingMapKeyObject
								+ "\tactual value: " + clientMap.get(holdingMapKey) + "\n");
					}
				} else if (holdingMapKeyObject instanceof Number) {
					Number act = Double.parseDouble((String) holdingMapKeyObject.toString().replaceAll("\"", ""));// 08/06/2018
					Number exp = Double.parseDouble(clientMap.get(holdingMapKey).toString().replaceAll("\"", ""));// 08/06/2018
					if (!(act).equals(exp)) {
						log.append("Key: " + holdingMapKey + "\texpected value: " + holdingMapKeyObject
								+ "\tactual value: " + clientMap.get(holdingMapKey) + "\n");
					}
				}
				diffMap.put("ExpectedValue", holdingMapKeyObject.toString());
				diffMap.put("ActualValue", clientMap.get(holdingMapKey).toString());
				log.append("Key: To check array elements ::: " + "\texpected value: " + holdingMapKeyObject.toString()
						+ "\tactual value: " + clientMap.get(holdingMapKey).toString() + "\n");

				// resultMap.put(holdingMapKey, diffMap);
				resultMap.put(holdingMapKey, diffMap);
			}
		}
		// Checking if any of the response map keys are missing from holding
		// map.
		clientMapKeySet = clientMap.keySet();
		clientMapKeySet.removeAll(holdingMapKeySet);
		if (!clientMapKeySet.isEmpty()) {
			for (String missingKey : clientMapKeySet) {
				Map<String, String> diffMap = new HashMap<String, String>();
				log.append("Missing key: <" + missingKey + "> key not found in holding map.\n");
				diffMap.put("ExpectedValue", null);
				diffMap.put("ActualValue", clientMap.get(missingKey).toString());
				resultMap.put(missingKey, diffMap);
			}
		}
	}

	@SuppressWarnings("unchecked")
	public static <T extends Object> T compareJSON(String responseJson, String expectedJson, String returnType)
			throws Exception {
		T t = null;
		Map<String, Map<String, String>> resultMap = new HashMap<String, Map<String, String>>();

		StringBuilder log = new StringBuilder();

		Map<String, Object> holdingMap = (Map<String, Object>) new Gson().fromJson(expectedJson, Map.class);
		Map<String, Object> clientMap = (Map<String, Object>) new Gson().fromJson(responseJson, Map.class);
		// ***

		JsonUtility.compareJson(holdingMap, clientMap, log, resultMap);
		if ("String".equalsIgnoreCase(returnType)) {
			t = (T) log.toString();
		} else if ("Map".equalsIgnoreCase(returnType)) {
			t = (T) resultMap;
		} else {
			throw new Exception("Return type not supported");
		}
		return t;

	}

	 public static boolean validateFailedJsonContent(String json, String pathToNode, String expNodeDesc) {
	        boolean status = true;
	        String actualNodeDesc = null;

	        String[] possiblePaths = {
	            "$..Messages[*].Description",
	            "$..Message[*].Description",
	            "$..messages[*].Description",
	            "$..message[*].Description",
	            "$..message[*].user",
	            "$..messages[*].user",
	            "$..Messages[*].user",
	            "$..Message[*].user",
	            "$.ErrorMessage.messages[*]",//For Billing validation response - 18/02/2025
	            "$.ErrorMessage.message[*]"//For Billing validation response - 18/02/2025
	        };

	        try {
	            for (String path : possiblePaths) {
	                try {
	                    actualNodeDesc = JsonPath.read(json, path).toString();
	                    if (!actualNodeDesc.equals("[]")) {
	                        break; // Found a valid path with a description
	                    }
	                } catch (PathNotFoundException e) {
	                    // Continue to the next path
	                }
	            }

	            if (actualNodeDesc == null || actualNodeDesc.equals("[]")) {
	                WebHelper.restErrorResDesc = "Description not found in the JSON response.";
	                log.info(WebHelper.restErrorResDesc);
	                return false;
	            }

	            actualNodeDesc = actualNodeDesc.replaceAll("\\\\/", "/");
	            log.info("actualNodeDesc is: " + actualNodeDesc);

	            String[] expNodeDescList = expNodeDesc.split("\\|");
	            for (String tempStr : expNodeDescList) {
	                if (!actualNodeDesc.contains(tempStr)) {
	                    WebHelper.restErrorResDesc = "Failed: Message from the actual response: " + actualNodeDesc +
	                            " does not match expected response: " + expNodeDesc;
	                    log.info(WebHelper.restErrorResDesc);
	                    status = false;
	                }
	            }

	            if (status) {
	                if (actualNodeDesc.startsWith("[") && actualNodeDesc.endsWith("]")) {
	                    actualNodeDesc = actualNodeDesc.substring(1, actualNodeDesc.length() - 1);
	                }
	                WebHelper.restErrorResDesc = "Message from the actual response: " + actualNodeDesc +
	                        " matching expected response: " + expNodeDesc;
	                log.info("Message from the actual response: " + actualNodeDesc + " matching expected response: " + expNodeDesc);
	            }

	        } catch (Exception e) {
	            log.info("Exception occurred while parsing response json file: " + e.getMessage());
	            WebHelper.restErrorResDesc = "Exception occurred while parsing response json file: " + e.getMessage();
	            e.printStackTrace();
	            return false;
	        }

	        return status;
	    }
	
	
	/*
	 * public static boolean validateFailedJsonContent(String json, String
	 * pathToNode, String expNodeDesc) { try { // WebHelper.restErrorResDesc =
	 * "Following messages are not in Actual // Response:"; boolean status = true;
	 * 
	 * String actualNodeDesc = null;
	 * 
	 * try { actualNodeDesc = JsonPath.read(json,
	 * "$..Message[*].Description").toString(); if (actualNodeDesc.equals("[]")) {
	 * try { actualNodeDesc = JsonPath.read(json,
	 * "$..Message[*].Description").toString(); } catch (PathNotFoundException ex) {
	 * WebHelper.restErrorResDesc =
	 * "Neither 'Message' nor 'Messages' found in the JSON response.";
	 * log.info(WebHelper.restErrorResDesc); status = false; } } if
	 * (actualNodeDesc.equals("[]")) { try { actualNodeDesc = JsonPath.read(json,
	 * "$..message[*].user").toString(); } catch (PathNotFoundException ex) {
	 * WebHelper.restErrorResDesc =
	 * "Neither 'Message', 'Messages', nor 'message.user' found in the JSON response."
	 * ; log.info(WebHelper.restErrorResDesc); status = false; } }
	 * 
	 * } catch (PathNotFoundException e) { // If "Message" is not found, try
	 * "Messages", and then "message.user" - UPDATED try { actualNodeDesc =
	 * JsonPath.read(json, "$..Messages[*].Description").toString(); } catch
	 * (PathNotFoundException ex) { try { actualNodeDesc = JsonPath.read(json,
	 * "$..messages[*].user").toString(); } catch (PathNotFoundException ex2) {
	 * WebHelper.restErrorResDesc =
	 * "Neither 'Message', 'Messages', nor 'message.user' found in the JSON response."
	 * ; log.info(WebHelper.restErrorResDesc); status = false; } } }
	 * 
	 * actualNodeDesc = actualNodeDesc.replaceAll("\\\\/", "/");
	 * log.info("actualNodeDesc is: " + actualNodeDesc);
	 * 
	 * String[] expNodeDescList = expNodeDesc.split("\\|"); for (String tempStr :
	 * expNodeDescList) { if (!actualNodeDesc.contains(tempStr)) {
	 * 
	 * // WebHelper.restErrorResDesc = WebHelper.restErrorResDesc + "\n" + tempStr;
	 * WebHelper.restErrorResDesc = "Failed: Message from the actual response :" +
	 * actualNodeDesc + " does not match expected response :" + expNodeDesc;
	 * 
	 * log.info(WebHelper.restErrorResDesc); status = false; } } if (status == true)
	 * { if (actualNodeDesc.startsWith("[") && actualNodeDesc.endsWith("]")) {
	 * actualNodeDesc = actualNodeDesc.substring(1, actualNodeDesc.length() - 1); }
	 * WebHelper.restErrorResDesc = "Message from the actual response :" +
	 * actualNodeDesc + " matching expected response :" + expNodeDesc; //
	 * log.info("All messages are present in actual response");
	 * log.info("Message from the actual response :" + actualNodeDesc +
	 * " matching expected response :" + expNodeDesc); }
	 * 
	 * return status; } catch (Exception e) {
	 * log.info("Exception occured while parsing response json file: " +
	 * e.getMessage()); WebHelper.restErrorResDesc =
	 * "Exception occured while parsing response json file: " + e.getMessage();
	 * e.printStackTrace(); // Log the stack trace return false; } }
	 */

	public static boolean validateRestJson(String json) {
		try {
			// it will search for SuccessFlag tag in json and store its values
			// in successList
			List<String> successList = JsonPath.read(json, "$..SuccessFlag");
			boolean status = true;
			WebHelper.restErrorResDesc = "";

			if (successList.contains("FAILED") | successList.contains("failed") | successList.size() == 0) {
				status = false;
				// in case of failure it will search for Description tag and
				// stores its value in errDescList
				List<String> errDescList = JsonPath.read(json, "$..Description");
				if (!errDescList.isEmpty())
					WebHelper.restErrorResDesc = errDescList.get(0);
			}
			return status;
		} catch (Exception e) {
			log.info("Exception occured while parsing response json file: " + e.getMessage());
			WebHelper.restErrorResDesc = "Exception occured while parsing response json file: " + e.getMessage();
			return false;
		}
	}

/*	
	public static boolean validateOASJson(Response OASResponse) {
		try {
			String response = OASResponse.asString();
			boolean status = true;
			WebHelper.restErrorResDesc = "";
			String desc = "";
			desc = desc + "Http Status Code: " + OASResponse.getStatusCode() + " recieved";
			// Declaring successList outside the if statement
			List<String> successList = new ArrayList<>();
			List<String> readMessage = new ArrayList<>();

			if (!String.valueOf(OASResponse.getStatusCode()).startsWith("20")) {
				status = false;
				response = response.replace("\"", "\'");
				desc = desc + " (" + response + ")";
				log.info("validateOASJson response: " + desc);
			}

			// If present, Checking value of "SuccessFlag" key in Response.
			if (status && ITAFWebDriver.isBillingApplication()) {
				// List<String> successList = JsonPath.read(OASResponse.getBody().asString(),
				// "$..SuccessFlag");
				successList = JsonPath.read(OASResponse.getBody().asString(), "$..SuccessFlag");
				if (successList.size() > 0 && (successList.contains("FAILED") | successList.contains("failed")
						| successList.contains("NOT_PROCESSED")))// Added for Base Billing 14/4/2022

				readMessage = JsonPath.read(OASResponse.getBody().asString(), "$..Messages[*].Description");
				status = false;
				log.info("validateOASJson response: " + desc + readMessage);
			}

			if (successList.isEmpty()&& desc.contains("200") ) {
				WebHelper.restErrorResDesc = desc;
				status = true;
				
			}else if (successList.isEmpty()) {
				WebHelper.restErrorResDesc = desc;
				
			} else if(!successList.isEmpty() && successList.contains("SUCCESS") ) {
				WebHelper.restErrorResDesc = desc + ","+" response contains SUCCESS";
				status = true; // Resetting status to true if the response contains SUCCESS
			}
			else {
				WebHelper.restErrorResDesc = desc + " response contains : " + readMessage;
			}
			
			return status;
		} catch (Exception e) {
			log.info("Exception occured while parsing response json file: " + e.getMessage());
			WebHelper.restErrorResDesc = "Exception occured while parsing response json file: " + e.getMessage();
			return false;
		}
	}
*/
	
	public static boolean validateOASJson(Response OASResponse) {
	    try {
	        String response = OASResponse.asString();
	        boolean status = true;
	        WebHelper.restErrorResDesc = "";
	        
	        StringBuilder desc = new StringBuilder("Http Status Code: ")
	            .append(OASResponse.getStatusCode()).append(" received");
 
	        String responseBody = OASResponse.getBody().asString(); // Avoid multiple calls
	        
	        // Lists to hold extracted JSON values
	        List<String> successList = new ArrayList<>();
	        List<String> readMessage = new ArrayList<>();
 
	        // If status code is not in the 200s, log and mark failure
	        if (!String.valueOf(OASResponse.getStatusCode()).startsWith("20")) {
	            status = false;
	            response = response.replace("\"", "\'");
	            desc.append(" (").append(response).append(")");
	            log.info("validateOASJson response: " + desc);
	        }
 
	        // If present, check the value of "SuccessFlag" key in Response.
	        if (status && ITAFWebDriver.isBillingApplication()) {
	            successList = JsonPath.read(responseBody, "$..SuccessFlag");
 
	            if (successList != null && !successList.isEmpty() &&
	                (successList.contains("FAILED") || successList.contains("failed") || successList.contains("NOT_PROCESSED"))) {
	                
	                readMessage = JsonPath.read(responseBody, "$..Messages[*].Description");
	                status = false;
	                log.info("validateOASJson response: " + desc + readMessage);
	            }
	        }
 
	        // Final decision on status
	        if (successList == null || successList.isEmpty()) {
	            if (desc.toString().contains("200")) {
	                WebHelper.restErrorResDesc = desc.toString();
	                status = true;
	            } else {
	                WebHelper.restErrorResDesc = desc.toString();
	            }
	        } else if (successList.contains("SUCCESS")) {
	            WebHelper.restErrorResDesc = desc.append(", response contains SUCCESS").toString();
	            status = true;
	        } else {
	            WebHelper.restErrorResDesc = desc.append(" response contains: ").append(readMessage).toString();
	        }
 
	        return status;
 
	    } catch (Exception e) {
	        String errorMsg = "Exception occurred while parsing response JSON file: " + e.getMessage();
	        log.info(errorMsg);
	        WebHelper.restErrorResDesc = errorMsg;
	        return false;
	    }
	}	
	
//For DM team - 17/04/24
	/*
	 * public static String getValuesFromJson(String jsonFile, String jsonTag)
	 * throws Exception { JSONObject jsonObject = null; JSONArray jsonArray = null;
	 * tagValue = null; if (jsonFile.startsWith("{")) { jsonObject = new
	 * JSONObject(jsonFile); tagValue = readTagFromJsonObject(jsonObject, jsonTag);
	 * if (tagValue != null) { return tagValue; }else {
	 * log.info("Invalid Json!!! Json Tag-- " + jsonTag+
	 * " --not present in the response" ); } return null;
	 * 
	 * } else if (jsonFile.startsWith("[")) { jsonArray = new JSONArray(jsonFile);
	 * int jsonArrayLen = jsonArray.length(); for (int i = 0; i < jsonArrayLen; i++)
	 * { jsonObject = (JSONObject) jsonArray.get(i); tagValue =
	 * readTagFromJsonObject(jsonObject, jsonTag); if (tagValue != null) return
	 * tagValue; } } else { log.info("Invalid Json!!!"); } return null;
	 * 
	 * }
	 */

	// For DM team - 17/04/24
	public static String getValuesFromJson(String jsonFile, String jsonTag) throws Exception {
		JSONObject jsonObject = null;
		JSONArray jsonArray = null;
		String tagValue = null;
		if (jsonFile.startsWith("{")) {
			jsonObject = new JSONObject(jsonFile);
			tagValue = readTagFromJsonObject(jsonObject, jsonTag);
			if (tagValue != null) {
				return tagValue;
			} else {
				throw new JSONException("Invalid Json!!! Json Tag-- " + jsonTag + " --not present in the response");
			}
		} else if (jsonFile.startsWith("[")) {
			jsonArray = new JSONArray(jsonFile);
			int jsonArrayLen = jsonArray.length();
			for (int i = 0; i < jsonArrayLen; i++) {
				jsonObject = (JSONObject) jsonArray.get(i);
				tagValue = readTagFromJsonObject(jsonObject, jsonTag);
				if (tagValue != null)
					return tagValue;
			}
		} else {
			throw new JSONException("Invalid Json!!!");
		}
		throw new JSONException("Json Tag-- " + jsonTag + " --not present in the response");
	}

	public static String readTagFromJsonObject(JSONObject obj, String keyMain) throws Exception {
		// We need to know keys of Jsonobject
		@SuppressWarnings("rawtypes")
		Iterator iterator = obj.keys();
		String key = null;

		while (iterator.hasNext()) {
			key = (String) iterator.next();
			// if object is just string we change value in key
			if ((obj.optJSONArray(key) == null) && (obj.optJSONObject(key) == null)) {
				if ((key.equals(keyMain))) {
					tagValue = obj.get(keyMain).toString();
					break;
				}
			}

			// if it's jsonobject
			if (obj.optJSONObject(key) != null) {
				readTagFromJsonObject(obj.getJSONObject(key), keyMain);
			}

			// if it's jsonarray
			if (obj.optJSONArray(key) != null) {
				JSONArray jArray = obj.getJSONArray(key);
				for (int i = 0; i < jArray.length(); i++) {
					readTagFromJsonObject(jArray.getJSONObject(i), keyMain);
				}
			}
		}
		return tagValue;
	}

	/*
	 * public static String readFromUniqueNumberSheet(String testCaseId, String
	 * colomnName) {
	 * 
	 * Sheet uniqueNumberSheet = null; String uniqueTestcaseID = ""; HashMap<String,
	 * Integer> uniqueValuesHashMap = null; String uniqueValue = null;
	 * MainController controller = ObjectFactory.getMainController(); String
	 * SearchTestcaseID = null;
	 * 
	 * try { uniqueNumberSheet = WebHelperUtil.getSheet(Config.transactionInfo,
	 * "DataSheet"); uniqueValuesHashMap =
	 * WebHelperUtil.getValueFromHashMap(uniqueNumberSheet); int rowCount =
	 * uniqueNumberSheet.getPhysicalNumberOfRows();
	 * 
	 * if (testCaseId.equals("")) { SearchTestcaseID =
	 * controller.controllerTestCaseID.toString(); } else { SearchTestcaseID =
	 * testCaseId; }
	 * 
	 * for (int rIndex = 1; rIndex < rowCount; rIndex++) { uniqueTestcaseID =
	 * WebHelperUtil.getCellData("TestCaseID", uniqueNumberSheet, rIndex,
	 * uniqueValuesHashMap); if (SearchTestcaseID.equals(uniqueTestcaseID)) { return
	 * uniqueValue = WebHelperUtil.getCellData(colomnName, uniqueNumberSheet,
	 * rIndex, uniqueValuesHashMap); } }
	 * 
	 * } catch (Exception e) { log.error(e.getMessage(), e);
	 * controller.pauseFun(e.getMessage()); } return uniqueValue; }
	 */

	public static String readFromUniqueNumberSheet(String testCaseId, String columnName) throws Exception {
		Sheet uniqueNumberSheet = null;
		String uniqueTestcaseID = "";
		HashMap<String, Integer> uniqueValuesHashMap = null;
		String uniqueValue = null;
		MainController controller = ObjectFactory.getMainController();
		String SearchTestcaseID = null;

		/*
		 * try { uniqueNumberSheet = WebHelperUtil.getSheet(Config.transactionInfo,
		 * "DataSheet"); uniqueValuesHashMap =
		 * WebHelperUtil.getValueFromHashMap(uniqueNumberSheet); int rowCount =
		 * uniqueNumberSheet.getPhysicalNumberOfRows();
		 * 
		 * if (testCaseId.equals("")) { SearchTestcaseID =
		 * controller.controllerTestCaseID.toString(); } else { SearchTestcaseID =
		 * testCaseId; }
		 * 
		 * for (int rIndex = 1; rIndex < rowCount; rIndex++) { uniqueTestcaseID =
		 * WebHelperUtil.getCellData("TestCaseID", uniqueNumberSheet, rIndex,
		 * uniqueValuesHashMap);
		 * 
		 * if (SearchTestcaseID.equals(uniqueTestcaseID)) { uniqueValue =
		 * WebHelperUtil.getCellData(columnName, uniqueNumberSheet, rIndex,
		 * uniqueValuesHashMap); try { // Check if uniqueValue is empty or blank if
		 * (uniqueValue == null || uniqueValue.trim().isEmpty()) {
		 * log.info("Error: Value under Column " + "'"+ columnName +"'"+
		 * " in the unique no sheet is empty for the provided testCaseId: " +
		 * SearchTestcaseID);
		 * 
		 * throw new Exception("Value under Column " + columnName +
		 * " in the unique no sheet is empty for the provided testCaseId: " +
		 * SearchTestcaseID);
		 * 
		 * } } catch (Exception e) { // Handle the exception here, you can log it or
		 * perform any other action //e.printStackTrace(); }
		 * 
		 * return uniqueValue; } }
		 * 
		 * // No match found, throw specific exception throw new Exception(
		 * "No match found in the unique no sheet for the provided testCaseId: " +
		 * SearchTestcaseID);
		 * 
		 * } catch (IOException e) { String errorMessage = "Error: Value under Column "
		 * + columnName +
		 * " in the unique no sheet is empty for the provided testCaseId: " +
		 * SearchTestcaseID; log.error(errorMessage); throw new Exception(errorMessage);
		 * 
		 * } finally { if (uniqueNumberSheet != null) { try {
		 * uniqueNumberSheet.getWorkbook().close(); // Close workbook which implicitly
		 * closes sheet } catch (IOException e) { log.error("Error closing workbook",
		 * e); } } }
		 */

		try {
			uniqueNumberSheet = WebHelperUtil.getSheet(Config.transactionInfo, "DataSheet");
			uniqueValuesHashMap = WebHelperUtil.getValueFromHashMap(uniqueNumberSheet);
			int rowCount = uniqueNumberSheet.getPhysicalNumberOfRows();

			if (testCaseId.equals("")) {
				SearchTestcaseID = controller.controllerTestCaseID.toString();
			} else {
				SearchTestcaseID = testCaseId;
			}

			for (int rIndex = 1; rIndex < rowCount; rIndex++) {
				uniqueTestcaseID = WebHelperUtil.getCellData("TestCaseID", uniqueNumberSheet, rIndex,
						uniqueValuesHashMap);

				if (SearchTestcaseID.equals(uniqueTestcaseID)) {
					uniqueValue = WebHelperUtil.getCellData(columnName, uniqueNumberSheet, rIndex, uniqueValuesHashMap);
					// Check if uniqueValue is empty or blank
					if (uniqueValue == null || uniqueValue.trim().isEmpty()) {
						log.info("Error: Value under Column " + "'" + columnName + "'"
								+ " in the unique no sheet is empty for the provided testCaseId: " + SearchTestcaseID);

						throw new Exception("Value under Column " + "'" + columnName + "'"
								+ " in the unique no sheet is empty for the provided testCaseId: " + SearchTestcaseID);
					}

					return uniqueValue;
				}
			}

			// No match found, throw specific exception
			throw new Exception(
					"No match found in the unique no sheet for the provided testCaseId: " + SearchTestcaseID);
		} catch (IOException e) {
			String errorMessage = "Error: Value under Column " + columnName
					+ " in the unique no sheet is empty for the provided testCaseId: " + SearchTestcaseID;
			log.error(errorMessage);
			throw new Exception(errorMessage);
		} finally {
			if (uniqueNumberSheet != null) {
				try {
					uniqueNumberSheet.getWorkbook().close(); // Close workbook which implicitly closes sheet
				} catch (IOException e) {
					log.error("Error closing workbook", e);
				}
			}
		}

	}

	public static void readFromResponseJson(String restResponse1, String readFromResponse)
			throws JSONException, Exception {
		try {
			String tagValue = "";
			HashMap<String, String> resValuesMap = new HashMap<>();
			List<String> tagValueList = null;
			String[] resTags = readFromResponse.split(",");
			String key = null;
			for (String jsonPath : resTags) {
				key = null;
				jsonPath = jsonPath.trim();
				if (!jsonPath.startsWith("root") && jsonPath.contains("=")) {
					key = jsonPath.substring(0, jsonPath.indexOf("=")).trim();
					jsonPath = jsonPath.substring(jsonPath.indexOf("=") + 1).trim();
				}
				if (jsonPath.startsWith("root")) {
					// findjsonElement updated for DM team - 17/04/24
					tagValueList = findjsonElement(restResponse1, jsonPath.trim());
					if (key == null) {
						String[] temp = jsonPath.split("\\.");
						key = temp[temp.length - 1];
						if (key.contains("WHERE"))
							key = key.split("WHERE")[0].trim();
					}

					if (tagValueList.size() == 1)
						resValuesMap.put(key, tagValueList.get(0));
					else
						for (int i = 0; i < tagValueList.size(); i++)
							resValuesMap.put(key + i, tagValueList.get(i));
				} else {
					// getValuesFromJson updated for DM team - 17/04/24
					tagValue = JsonUtility.getValuesFromJson(restResponse1, jsonPath.trim());
					// resValuesMap.put(jsonPath.trim(), tagValue);
					if (tagValue != null) {
						resValuesMap.put(jsonPath.trim(), tagValue);
					} else {
						// Handle null tagValue here
						throw new Exception("Tag value is null for JSON path: " + jsonPath.trim());
					}
				}
				
				if (jsonValueList != null) {
				    jsonValueList.clear();
				} else {
				    jsonValueList = new ArrayList<>(); // Reinitialize if null
				}
			}
			com.majesco.itaf.rest.utils.ExcelUtility.fromAPIResponseToExcel(resValuesMap);

		} catch (Exception e) {
			// log.error(e.getMessage(), e);
			String errorMessage = e.getMessage();
			String firstLine = errorMessage.split("\\r?\\n")[0];
			log.error("Error :" + firstLine);
			throw e;
		}
	}

	public static String getTodaysDate() {
		LocalDateTime myDateObj = LocalDateTime.now();
		DateTimeFormatter myFormatObj = DateTimeFormatter.ofPattern("ddMMyyyyHH:mm");
		String todayDate = myDateObj.format(myFormatObj);
		log.info("Today's date: " + todayDate);
		return todayDate;
	}

	public static void updateReqJsonFile(String requestJsonFile, String updateReqBodyParam) throws Exception {

		String[] updateParamArray = updateReqBodyParam.split(",");
		Map<String, String> updatedMap = new HashMap<>();
		for (int i = 0; i < updateParamArray.length; i++) {
			String[] paramList = updateParamArray[i].split("=");
			String toBereadFromUniqueSheet = paramList[1].trim();
			String updatedValue = "";

			// call to new update method
			if (updateParamArray[i].trim().startsWith("root")) {
				jsonUpdater(requestJsonFile, updateParamArray[i]);
				continue;
			}

			if (toBereadFromUniqueSheet.contains("<>")) { // specific to claims team
				String[] tempArray = toBereadFromUniqueSheet.split("<>");
				updatedValue = JsonUtility.readFromUniqueNumberSheet("", tempArray[0].trim());
				if (tempArray[1].equalsIgnoreCase("today_date")) {
					tempArray[1] = getTodaysDate();
				}
				updatedValue = updatedValue + tempArray[1];
			} else {
				if (toBereadFromUniqueSheet.equals("@Date"))
					updatedValue = getDateFromOpenAPISheet();
				else

					updatedValue = JsonUtility.readFromUniqueNumberSheet("", toBereadFromUniqueSheet);
			}
			updatedMap.put(paramList[0].trim(), updatedValue);
		}
		updateJson(requestJsonFile, updatedMap);// Sel4

	}

	private static void updateJson(String requestJsonFile, Map<String, String> updatedMap) throws Exception {

		for (Map.Entry<String, String> entry : updatedMap.entrySet()) {
			String tagName = entry.getKey();
			String tagValue = entry.getValue();
			updateJsonTag(requestJsonFile, tagName.trim(), tagValue.trim());
		}
	}

	@SuppressWarnings("deprecation")
	private static void updateJsonTag(String requestJsonFile, String tagName, String tagValue) throws Exception {
		String jsonFile = new String(Files.readAllBytes(Paths.get(requestJsonFile)));

		JSONObject jsonObject = null;
		JSONArray jsonArray = null;
		String updatedJsonFile = null;
		if (jsonFile.startsWith("{")) {
			jsonObject = new JSONObject(jsonFile);
			jsonObject = updateJsonObject(jsonObject, tagName, tagValue);
			updatedJsonFile = jsonObject.toString();
		} else if (jsonFile.startsWith("[")) {
			jsonArray = new JSONArray(jsonFile);
			int jsonArrayLen = jsonArray.length();
			for (int i = 0; i < jsonArrayLen; i++) {
				jsonObject = (JSONObject) jsonArray.get(i);
				jsonObject = updateJsonObject(jsonObject, tagName, tagValue);
				jsonArray.put(i, jsonObject);
			}
			updatedJsonFile = jsonArray.toString();
		} else {
			log.info("Invalid Json present at path: " + requestJsonFile);
		}

		// To be further tested in Selenium 4 - 28/03/24
		/*
		 * // Parse the JSON string into a JsonObject using static method parseString
		 * JsonObject je = JsonParser.parseString(updatedJsonFile).getAsJsonObject(); //
		 * Create Gson instance with pretty printing enabled Gson gson = new
		 * GsonBuilder().setPrettyPrinting().create(); // Convert the JsonObject to a
		 * pretty-printed JSON string updatedJsonFile = gson.toJson(je);
		 */

		// convert to pretty json format//

		Gson gson = new GsonBuilder().setPrettyPrinting().create();
		JsonParser jp = new JsonParser();
		JsonElement je = jp.parse(updatedJsonFile);
		updatedJsonFile = gson.toJson(je);

		FileWriter writer = new FileWriter(new File(requestJsonFile));
		writer.write(updatedJsonFile);
		writer.close();
		log.info("JsonTag updated sucessfully");

	}

	@SuppressWarnings("unused")
	private static void updateJsonTag3(String requestJsonFile, String tagName, String tagValue) throws Exception {
		String jsonFile = new String(Files.readAllBytes(Paths.get(requestJsonFile)));

		JsonElement jsonElement = JsonParser.parseString(jsonFile);
		JsonObject jsonObject = jsonElement.getAsJsonObject();

		// Update JSON object recursively
		updateJsonObject(jsonObject, tagName, tagValue);

		// Convert JSON object back to String
		Gson gson = new GsonBuilder().setPrettyPrinting().create();
		String updatedJsonFile = gson.toJson(jsonObject);

		// Write updated JSON to file
		FileWriter writer = new FileWriter(new File(requestJsonFile));
		writer.write(updatedJsonFile);
		writer.close();

		// System.out.println("Done");
	}

	private static void updateJsonObject(JsonObject jsonObject, String tagName, String tagValue) {
		for (String key : jsonObject.keySet()) {
			JsonElement element = jsonObject.get(key);
			if (element.isJsonObject()) {
				updateJsonObject(element.getAsJsonObject(), tagName, tagValue);
			} else if (element.isJsonArray()) {
				for (JsonElement arrayElement : element.getAsJsonArray()) {
					updateJsonObject(arrayElement.getAsJsonObject(), tagName, tagValue);
				}
			} else if (key.equals(tagName)) {
				jsonObject.addProperty(key, tagValue);
			}
		}
	}

	public static JSONObject updateJsonObject(JSONObject obj, String keyMain, String newValue) throws Exception {
		// We need to know keys of Jsonobject
		@SuppressWarnings("unused")
		JSONObject json = new JSONObject();
		@SuppressWarnings("rawtypes")
		Iterator iterator = obj.keys();
		String key = null;

		while (iterator.hasNext()) {
			key = (String) iterator.next();
			// if object is just string we change value in key
			if ((obj.optJSONArray(key) == null) && (obj.optJSONObject(key) == null)) {
				if ((key.equals(keyMain))) {
					obj.put(key, newValue);
					return obj;
				}
			}
			// if it's jsonobject
			if (obj.optJSONObject(key) != null) {
				updateJsonObject(obj.getJSONObject(key), keyMain, newValue);
			}
			// if it's jsonarray
			if (obj.optJSONArray(key) != null) {
				JSONArray jArray = obj.getJSONArray(key);
				for (int i = 0; i < jArray.length(); i++) {
					updateJsonObject(jArray.getJSONObject(i), keyMain, newValue);
				}
			}
		}
		return obj;
	}

	public static String getDateFromOpenAPISheet() throws IOException {
		Workbook wb = null;
		Sheet tempSheet = null;

		File xlsxFile = new File(Config.inputDataFilePath + "OpenAPI\\OpenAPI.xlsx");
		File xlsFile = new File(Config.inputDataFilePath + "OpenAPI\\OpenAPI.xls");
		if (xlsxFile.exists()) {
			wb = new XSSFWorkbook(new FileInputStream(xlsxFile));
		} else {
			wb = new HSSFWorkbook(new FileInputStream(xlsFile));
		}
		tempSheet = wb.getSheet("Structure");
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
		// keep your date in 'Structure' sheet of openAPI file.
		CellReference cellReference = new CellReference("N2");
		Row row = tempSheet.getRow(cellReference.getRow());
		Cell cell = row.getCell(cellReference.getCol());
		CellValue cellValue = evaluator.evaluate(cell);
		String date = cellValue.getStringValue();
		log.info("Date From Open API Sheet is:" + date);
		wb.close();
		return date;
	}

	public static List<String> findjsonElement(String restResponse, String jsonElementPath1) throws Exception {
		String jsonElementPath = null;
		conditionsMap = new LinkedHashMap<String, String>();
		JSONObject rootjsonObject = null;
		pathHasConditionFlag = false;
		conditionsCheckedFlag = false;
		// String jsonElementPath1= "root.entities[*].entityId WHERE
		// sourceEntityType=IND_ADJ && cellPhoneNo=9194178804";

		// String jsonAsString= readJsonFile("getClaimEntityAll_Response.json");

		if (jsonElementPath1.contains("WHERE")) {
			pathHasConditionFlag = true;
			jsonElementPath = jsonElementPath1.split("WHERE")[0].trim();
			String condition = jsonElementPath1.split("WHERE")[1].trim();
			String[] conditionArray = condition.split("&&");
			for (int i = 0; i < conditionArray.length; i++) {
				String key = conditionArray[i].split("=")[0].trim();
				String value = conditionArray[i].split("=")[1].trim();
				if (value.equals("*"))
					value = readFromUniqueNumberSheet("", key);
				conditionsMap.put(key, value.trim());
			}
		} else
			jsonElementPath = jsonElementPath1;

		pathAsArray = Arrays.asList(jsonElementPath.split("\\."));

		//List<String> JsonElement=null;
		//public static List<String> jsonValueList = new ArrayList<String>();
		
		if (restResponse.startsWith("[")) {
			// if file is a JSONArray then name it "root" and put it in a JSONObject
			JSONArray jArray = new JSONArray(restResponse);
			rootjsonObject = new JSONObject();
			rootjsonObject.put("root", jArray);
			jsonValueList = getJsonElement(rootjsonObject.toString(), jsonElementPath);
		} else {
			// if file is json object
			rootjsonObject = new JSONObject(restResponse);
			jsonValueList = getJsonElement(rootjsonObject.toString(), jsonElementPath);
		}
		log.info("List of values found from json for " + jsonElementPath + " are: " + jsonValueList);
		return jsonValueList;
	}

    @SuppressWarnings("unchecked")
	public static List<String> getJsonElement(String jsonString, String jsonPath) throws Exception {
        try {
            // Convert JSON string to Map using TypeReference to ensure type safety
            ObjectMapper objectMapper = new ObjectMapper();
            jsonPath=jsonPath.replace("root", "$");
            Map<String, Object> jsonMap = objectMapper.readValue(jsonString, new TypeReference<Map<String, Object>>() {});

            // Use Jayway JSONPath to extract values dynamically
            Object result = JsonPath.read(jsonMap, jsonPath);

            // Handle different return types safely
            if (result instanceof List) {
                return new ArrayList<>((List<String>) result);  // Convert to mutable list
            } else if (result instanceof String) {
                return new ArrayList<>(Collections.singletonList((String) result)); // Convert single value to mutable list
            } else if (result instanceof Number || result instanceof Boolean) {
                return new ArrayList<>(Collections.singletonList(result.toString())); // Convert number/boolean to mutable list
            } else {
                throw new Exception("Unexpected JSONPath result type: " + result.getClass());
            }
            
        } catch (Exception e) {
        	String errorMessage = (e.getMessage() != null && !e.getMessage().isEmpty()) ? e.getMessage(): "No error message available";
			String firstLine = errorMessage.split("\\r?\\n", 2)[0];
			log.error("Error :" + firstLine);
			throw new Exception("Error while reading from JSON response: " + e.getMessage(), e);
        }
    }

/*	
	public static List<String> getJsonElement(JSONObject jsonObject, int pathIndex, String jsonElementPath)
			throws Exception {
		boolean checkAllElements = false;
		int checkThisIndex = 0;
		String keyTrimmed = null;

		// for json array
		if (pathAsArray.get(pathIndex).contains("[")) {
			keyTrimmed = StringUtils.substringBefore(pathAsArray.get(pathIndex), "[");
			if (pathAsArray.get(pathIndex).contains("*"))
				checkAllElements = true;
			else {
				checkAllElements = false;
				String temp = pathAsArray.get(pathIndex).substring(pathAsArray.get(pathIndex).indexOf("[") + 1,
						pathAsArray.get(pathIndex).indexOf("]"));
				checkThisIndex = Integer.parseInt(temp);
			}
		} else
			keyTrimmed = pathAsArray.get(pathIndex); // for json object or key-value pair

		if (jsonObject.has(keyTrimmed)) {
			// if its a key-value pair
			if ((jsonObject.optJSONArray(keyTrimmed) == null) && (jsonObject.optJSONObject(keyTrimmed) == null)) {
				// verify that key does not contain [] or [*]
				if (pathAsArray.get(pathIndex).contains("[")) {
					throw new Exception("Error while reading from json response: " + "[index] or [*] in "
							+ pathAsArray.get(pathIndex) + " is inavalid as " + keyTrimmed
							+ " is a key-value pair in path: " + jsonElementPath);

				} else// For DM issue - 17/04/2024
						// Check if the value is an integer
				if (jsonObject.optInt(keyTrimmed, Integer.MIN_VALUE) != Integer.MIN_VALUE) {
					jsonValueList.add(String.valueOf(jsonObject.getInt(keyTrimmed)));

				} else {
					if (pathHasConditionFlag) {
						for (Entry<String, String> entry : conditionsMap.entrySet()) {
							if (jsonObject.has(entry.getKey())
									&& jsonObject.get(entry.getKey()).toString().equals(entry.getValue()))
								conditionsCheckedFlag = true;
							else {
								conditionsCheckedFlag = false;
								break;
							}
						}
						if (conditionsCheckedFlag)
							jsonValueList.add(jsonObject.getString(pathAsArray.get(pathIndex)));
					} else if (!pathHasConditionFlag)
						jsonValueList.add(jsonObject.getString(pathAsArray.get(pathIndex)));
				}
			}

			// if it's JSONObject
			if (jsonObject.optJSONObject(keyTrimmed) != null) {
				// verify that key does not contain [] or [*]
				if (pathAsArray.get(pathIndex).contains("["))
					throw new Exception("Error while reading from json response: " + "[index] or [*] in "
							+ pathAsArray.get(pathIndex) + " is inavalid as " + keyTrimmed
							+ " is a JSONObject in path: " + jsonElementPath);
				else
					getJsonElement(jsonObject.getJSONObject(keyTrimmed), pathIndex + 1, jsonElementPath);
			}

			// if it's JSONArray
			if (jsonObject.optJSONArray(keyTrimmed) != null) {
				// verify that key mentions [] or [*] in array
				if (pathAsArray.get(pathIndex).contains("[")) {
					JSONArray jArray = jsonObject.getJSONArray(keyTrimmed);
					if (!checkAllElements) {
						if (checkThisIndex <= (jArray.length() - 1))
							getJsonElement(jArray.getJSONObject(checkThisIndex), pathIndex + 1, jsonElementPath);
						else
							throw new Exception("Error while reading from json response: "
									+ "Array Index limit exceeded for " + pathAsArray.get(pathIndex)
									+ " in following json path: " + jsonElementPath);
					} else
						for (int i = 0; i < jArray.length(); i++)
							getJsonElement(jArray.getJSONObject(i), pathIndex + 1, jsonElementPath);
				} else
					throw new Exception("Error while reading from json response: "
							+ "Please mention [index] or [*] after " + pathAsArray.get(pathIndex) + " as " + keyTrimmed
							+ " is a JSONArray for path: " + jsonElementPath);
			}
		} else
			throw new Exception("Error while reading from json response: " + pathAsArray.get(pathIndex)
					+ " does not exist at path: " + jsonElementPath);

		return jsonValueList;
	}

*/
	public static void jsonUpdater(String requestJsonFile, String setElementPath1) throws Exception {
		JSONObject updatedJson = null;
		String updatedJsonFile = null;
		conditionsCheckedFlag = false;
		pathHasConditionFlag = false;
		String setElementPath = null;

		conditionsMap = new LinkedHashMap<String, String>();
		JSONObject rootjsonObject;
		String setValue = null;

		if (setElementPath1.contains("WHERE")) {
			pathHasConditionFlag = true;
			setElementPath = setElementPath1.split("WHERE")[0].trim().split("=")[0].trim();
			setValue = setElementPath1.split("WHERE")[0].trim().split("=")[1].trim();
			String condition = setElementPath1.split("WHERE")[1].trim();
			String[] conditionArray = condition.split("&&");
			String key = null;
			String value = null;
			for (int i = 0; i < conditionArray.length; i++) {
				key = conditionArray[i].split("=")[0].trim();
				value = conditionArray[i].split("=")[1].trim();
				if (value.equals("*"))
					value = readFromUniqueNumberSheet("", key);
				conditionsMap.put(key, value);
			}
		} else {
			setElementPath = setElementPath1.split("=")[0].trim();
			setValue = setElementPath1.split("=")[1].trim();
		}
		if (setValue.contains(":")) { // Added for DM Application '7/6/2022

			setValue = readFromUniqueNumberSheet(setValue.split(":")[0].trim(), setValue.split(":")[1].trim());
		} else {
			setValue = readFromUniqueNumberSheet("", setValue);
		}

		pathAsArray = Arrays.asList(setElementPath.split("\\."));
		String jsonFile = new String(Files.readAllBytes(Paths.get(requestJsonFile)));
		if (jsonFile.startsWith("[")) {
			// if file is a JSONArray then name it "root" and put it in a JSONObject
			// String jsonFile = new String(Files.readAllBytes(Paths.get(requestJsonFile)));
			JSONArray jArray = new JSONArray(jsonFile);
			rootjsonObject = new JSONObject();
			rootjsonObject.put("root", jArray);
			updatedJson = setJsonElement(rootjsonObject, 0, setValue, setElementPath);
			updatedJsonFile = updatedJson.toString(4);
		} else {
			// String jsonFile = new String(Files.readAllBytes(Paths.get(requestJsonFile)));
			rootjsonObject = new JSONObject(jsonFile);
			updatedJson = setJsonElement(rootjsonObject, 1, setValue, setElementPath);
			updatedJsonFile = updatedJson.toString(4);
		}

		FileWriter writer = new FileWriter(new File(requestJsonFile));
		writer.write(updatedJsonFile);
		writer.close();
	}

	public static JSONObject setJsonElement(JSONObject jsonObject, int pathIndex, String setValue,
			String setElementPath) throws Exception {

		Iterator<String> keys = jsonObject.keys();
		boolean checkAllElements = false;
		int checkThisIndex = 0;
		String keyTrimmed = null;
		// for json array
		if (pathAsArray.get(pathIndex).contains("[")) {
			keyTrimmed = StringUtils.substringBefore(pathAsArray.get(pathIndex), "[");
			if (pathAsArray.get(pathIndex).contains("*"))
				checkAllElements = true;
			else {
				checkAllElements = false;
				String temp = pathAsArray.get(pathIndex).substring(pathAsArray.get(pathIndex).indexOf("[") + 1,
						pathAsArray.get(pathIndex).indexOf("]"));
				checkThisIndex = Integer.parseInt(temp);
			}
		} else
			keyTrimmed = pathAsArray.get(pathIndex); // for json object or key-value pair

		int exists = 0;
		while (keys.hasNext()) {
			String key = keys.next();
			if (key.equals(keyTrimmed)) {
				exists = 1;
				// if its a key-value pair
				if ((jsonObject.optJSONArray(key) == null) && (jsonObject.optJSONObject(key) == null)) {
					// verify that key does not contain [] or [*]
					if (pathAsArray.get(pathIndex).contains("["))
						throw new Exception("[index] or [*] in " + pathAsArray.get(pathIndex) + " is inavalid as "
								+ keyTrimmed + " is a key-value pair in path: " + setElementPath);
					else {
						if (pathHasConditionFlag) {
							for (Entry<String, String> entry : conditionsMap.entrySet()) {
								if (jsonObject.has(entry.getKey())
										&& jsonObject.get(entry.getKey()).toString().equals(entry.getValue()))
									conditionsCheckedFlag = true;
								else {
									conditionsCheckedFlag = false;
									break;
								}
							}
							if (conditionsCheckedFlag) {
								jsonObject.put(keyTrimmed, setValue);
							}

						} else if (!pathHasConditionFlag) {
							jsonObject.put(keyTrimmed, setValue);
						}

					}
					break;
				}

				// if it's JSONObject
				if (jsonObject.optJSONObject(key) != null) {
					// verify that key does not contain [] or [*]
					if (pathAsArray.get(pathIndex).contains("["))
						throw new Exception("[index] or [*] in " + pathAsArray.get(pathIndex) + " is inavalid as "
								+ keyTrimmed + " is a JSONObject in path: " + setElementPath);
					else
						setJsonElement(jsonObject.getJSONObject(key), pathIndex + 1, setValue, setElementPath);
				}

				// if it's JSONArray
				if (jsonObject.optJSONArray(key) != null) {
					// verify that key mentions [] or [*] in array

					if (pathAsArray.get(pathIndex).contains("[")) {
						JSONArray jArray = jsonObject.getJSONArray(key);
						if (!checkAllElements) {
							if (checkThisIndex <= (jArray.length() - 1))
								setJsonElement(jArray.getJSONObject(checkThisIndex), pathIndex + 1, setValue,
										setElementPath);
							else
								throw new Exception("Array Index limit exceeded for " + pathAsArray.get(pathIndex)
										+ " in following json path: " + setElementPath);
						} else
							for (int i = 0; i < jArray.length(); i++)
								setJsonElement(jArray.getJSONObject(i), pathIndex + 1, setValue, setElementPath);
					} else
						throw new Exception("Please mention [index] or [*] after " + pathAsArray.get(pathIndex) + " as "
								+ keyTrimmed + " is a JSONArray for path: " + setElementPath);
				}
			}
		}
		if (exists == 0)
			throw new Exception(pathAsArray.get(pathIndex) + " not found in " + setElementPath);

		return jsonObject;
	}
}
