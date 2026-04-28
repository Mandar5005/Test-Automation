package com.majesco.itaf.rest.service;

import java.io.File;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.TimeUnit;
import io.restassured.RestAssured;
import io.restassured.authentication.OAuthSignature;
import io.restassured.builder.RequestSpecBuilder;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.http.HttpMethod;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.rest.utils.JsonUtility;

public class RestService {
	private final static Logger log = LogManager.getLogger(RestService.class.getName());

	public static Response getRestResponse(String contentType, HttpMethod httpMethod, String url, String input,
			String authType, String queryParam, String pathParam, String multiPart, String headers) throws Exception {

		RequestSpecification requestSpec = null;
		Response restResponse = null;
		
		@SuppressWarnings("unused")
		String response = null;
		if (url.startsWith("https")) {
			log.info("URL starts with https,using Relaxed HTTPSValidation");
			RestAssured.useRelaxedHTTPSValidation();
		}

		// Build RequestSpecification based on Authentication type.
		if (Config.authEnabled.equalsIgnoreCase("TRUE")) {
			if (authType.equalsIgnoreCase("Auth2.0")) {
				requestSpec = RestAssured.given().auth().oauth2(getAuthToken(), OAuthSignature.HEADER);
				log.info("Authentication Type is Auth2.0");}
			else {
				requestSpec = RestAssured.given().auth().preemptive().basic(Config.basicAuthUsername,
						Config.basicAuthPassword);
			log.info("Authentication Type is BasicAuthUsernamePassword");}
		}
		else {
			requestSpec = RestAssured.given();
		}

		requestSpec.contentType(contentType);// imp line
		if (queryParam.contains("="))
			requestSpec.queryParams(createParamMap(queryParam));
		if (pathParam.contains("="))
			requestSpec.pathParams(createParamMap(pathParam));
		if (headers.contains("="))
			requestSpec.headers(createParamMap(headers));
		if (!input.equals(""))
			requestSpec.body(input);
		if (!multiPart.equals(""))
			createMultpartReqSpec(requestSpec, createParamMap(multiPart));

		switch (httpMethod) {
		case POST:
			restResponse = requestSpec.post(url);
			break;
		case PUT:
			restResponse = requestSpec.put(url);
			break;
		case GET:
			restResponse = requestSpec.get(url);
			break;
		case DELETE:
			restResponse = requestSpec.delete(url);
			break;
		default:
			log.info("invalid httpMethod: " + httpMethod);
			break;
		}
		response = restResponse.getBody().asString();
		log.info("Time Taken(in Milliseconds): " + restResponse.getTimeIn(TimeUnit.MILLISECONDS));
		return restResponse;
	}

	private static void createMultpartReqSpec(RequestSpecification requestSpec, Map<String, String> multiPartMap) throws Exception {
		for (Map.Entry<String, String> entry : multiPartMap.entrySet()) {
			String key = entry.getKey();
			String value = entry.getValue();
			if (value.equals("*"))
				value = JsonUtility.readFromUniqueNumberSheet("", key.trim());

			if (key.contains("(file)")) {
				key = key.substring(0, key.indexOf("(file)"));
				requestSpec.multiPart(key, new File(value));
			} else
				requestSpec.multiPart(key, value);
		}
	}

	private static Map<String, String> createParamMap(String reqParam) throws Exception {
		Map<String, String> queryParamMap = new HashMap<String, String>();
		String[] paramArray = reqParam.split(",");
		for (String param : paramArray) {
			String[] values = param.split("=");
			
			 if (values.length > 1 && values[1] != null) { // Add null check for values[1]
			//if (values.length > 1) {
				if (values[1].trim().equals("*")) {
					values[1] = JsonUtility.readFromUniqueNumberSheet("", values[0].trim());

				} else if (values[1].trim().contains("$")) { // Minaakshi : 01-04-2020 : Start
					String[] tempArray = values[1].trim().split("$");
					if (tempArray.length > 1)
						values[1] = JsonUtility.readFromUniqueNumberSheet(tempArray[0].trim(), tempArray[1].trim());
					else {
						String strTemp = tempArray[0].toString().replace("$", "");
						values[1] = JsonUtility.readFromUniqueNumberSheet("", strTemp);
					}
				} // Minaakshi : 01-04-2020 : End
				else {
					if (values[1].trim().contains("<>")) {
						String[] tempArray = values[1].trim().split("<>");
						values[1] = JsonUtility.readFromUniqueNumberSheet("", tempArray[0].trim());
						if (tempArray[1].equalsIgnoreCase("today_date")) {
							tempArray[1] = JsonUtility.getTodaysDate();
						}
						values[1] = values[1] + tempArray[1];
					}
				}
				queryParamMap.put(values[0].trim(), values[1].trim());
			}
		}
		return queryParamMap;
	}

	public static String getAuthToken() throws Exception {
		try {
			RequestSpecBuilder builder;
			RequestSpecification requestSpec;
			// add path parameters to request builder
			builder = new RequestSpecBuilder();
			builder.addQueryParam("username", Config.authTokenUsername);
			builder.addQueryParam("scope", Config.authScope);
			builder.addQueryParam("client_secret", Config.authTokenClientSecret);
			builder.addQueryParam("grant_type", Config.authGrantType);
			builder.addQueryParam("client_id", Config.authTokenClientId);
			builder.addQueryParam("password", Config.authTokenPassword);
			requestSpec = builder.build();

			// *** Below given condition added for envs with https - authtokenURLS
			if (Config.authTokenURL.startsWith("https"))
				RestAssured.useRelaxedHTTPSValidation();
			// ***end***

			Response response = RestAssured.given().spec(requestSpec).when().post(Config.authTokenURL).then().extract()
					.response();

			log.info("Authtoken Response: " + response.asString());
			JsonPath js = new JsonPath(response.asString());
			String authToken = js.getString("access_token");
			return authToken;
		} catch (Exception e) {
			log.info("Auth2.0 token not generated: " + e.getMessage());
			throw new Exception("Error in generating Auth2.0 token: " + e.getMessage());
		}
	}
}
