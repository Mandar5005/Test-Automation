package com.majesco.itaf.rest.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import org.apache.commons.lang3.math.NumberUtils;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class CommonUtils {
	public static int compareStrings(final String str1, final String str2, final boolean nullIsLess) {
		if (str1 == str2) {
			return 0;
		}
		if (str1 == null) {
			return nullIsLess ? -1 : 1;
		}
		if (str2 == null) {
			return nullIsLess ? 1 : -1;
		}
		return str1.compareTo(str2);
	}

	public static String WriteToFile(File file, byte[] data) throws IOException {
		// return writeFile(file, data);
		OutputStream os = null;
		os = new FileOutputStream(file);
		os.write(data);
		os.flush();
		os.close();
		return "success";
	}
	
	public static String toPrettyFormat(String jsonString) {
        JsonObject json = null;
        try {
            json = JsonParser.parseString(jsonString).getAsJsonObject();
        } catch (Exception e) {
            System.out.println(jsonString);
        }

        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        String prettyJson = gson.toJson(json);

        return prettyJson;
    }

	
	public static int compare(Object o1, Object o2) {
	    // Null handling
	    if (o1 == null && o2 == null) {
	        return 0;
	    } else if (o1 == null) {
	        return -1;
	    } else if (o2 == null) {
	        return 1;
	    }

	    // Numeric comparison
	    if (NumberUtils.isCreatable(o1.toString()) && NumberUtils.isCreatable(o2.toString())) {
	        BigDecimal value1 = new BigDecimal(o1.toString());
	        BigDecimal value2 = new BigDecimal(o2.toString());
	        return value1.compareTo(value2);
	    }

	    // String comparison
	    return o1.toString().compareTo(o2.toString());
	}
}