package com.majesco.itaf.rest.utils;

import org.json.JSONException;
import org.json.JSONObject;
import org.json.XML;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class XML2JsonCompare {
	private final static Logger log = LogManager.getLogger(XML2JsonCompare.class);
    public static boolean main(String expectedXmlPath, String responseXmlPath, List<String> ignoreKeyList)
            throws Exception {
    	
        // Convert XML files to JSON
        String expectedJsonPath = convertXmlToJson(expectedXmlPath);
        String responseJsonPath = convertXmlToJson(responseXmlPath);
        boolean result = ValidateJson.main(expectedJsonPath, responseJsonPath, ignoreKeyList);
        return result;
    }

	
	public static String convertXmlToJson(String xmlFilePath) throws IOException, JSONException {
		String xmlString = new String(Files.readAllBytes(Paths.get(xmlFilePath)));
		JSONObject jsonObject = XML.toJSONObject(xmlString);

		String jsonFilePath = xmlFilePath.replace(".xml", ".json");
		Files.write(Paths.get(jsonFilePath), jsonObject.toString().getBytes());

		String jsonFilePathFormatted = preetyprint(jsonFilePath);
		
		log.info("XML converted to JSON successfully");

		return jsonFilePathFormatted;
	}

    
    public static String preetyprint (String jsonFilePath) {

       {
            try {
                String jsonString = new String(Files.readAllBytes(Paths.get(jsonFilePath)));
                JSONObject jsonObject = new JSONObject(jsonString);
                String prettyJsonString = jsonObject.toString(4);
                Files.write(Paths.get(jsonFilePath), prettyJsonString.getBytes());
                log.info("Pretty printed JSON saved to: " + jsonFilePath);
            } catch (IOException | JSONException e) {
                e.printStackTrace();
            }
        }
	return jsonFilePath;
    }

    public static void deleteFile(String filePath) throws IOException {
        Files.deleteIfExists(Paths.get(filePath));
    }

}
