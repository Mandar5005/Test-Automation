package com.majesco.itaf.util;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.WebHelper;

public class RemoveExtraSpaces {
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private final static Logger log = LogManager.getLogger(WebHelper.class.getName());

	public static String removeSpace(String destPath) throws Exception {
		String inputFileName = destPath;
		if (destPath.endsWith("_Expected.txt")) {
			// For _Expected.txt files, remove trailing spaces and overwrite the same file.
			try {
				FileReader fileReader = new FileReader(inputFileName);
				BufferedReader bufferedReader = new BufferedReader(fileReader);
				String tempFileName = inputFileName + ".temp";
				FileWriter fileWriter = new FileWriter(tempFileName);

				String line;
				while ((line = bufferedReader.readLine()) != null) {
					// Remove extra trailing spaces and keep only one space
					String trimmedLine = line.replaceAll("\\s+", " ");
					fileWriter.write(trimmedLine);
					fileWriter.write(System.lineSeparator());
				}
				bufferedReader.close();
				fileWriter.close();

				// Replace the original file with the modified one
				replaceFile(inputFileName, tempFileName);

				// log.info("Trailing spaces removed successfully!");
			} catch (IOException e) {
				e.printStackTrace();
			}

		} else {
			// For other files, remove trailing spaces and save to a new file with "_New"
			// suffix.
			String outputFileName = getNewFileName(inputFileName);

			try {
				FileReader fileReader = new FileReader(inputFileName);
				BufferedReader bufferedReader = new BufferedReader(fileReader);
				FileWriter fileWriter = new FileWriter(outputFileName);

				String line;
				while ((line = bufferedReader.readLine()) != null) {
					// Remove extra trailing spaces and keep only one space
					String trimmedLine = line.replaceAll("\\s+", " ");
					fileWriter.write(trimmedLine);
					fileWriter.write(System.lineSeparator());
				}

				bufferedReader.close();
				fileWriter.close();

			} catch (IOException e) {
				e.printStackTrace();
			}

			return inputFileName + "_Actual";
		}
		return inputFileName;
	}

	// Helper method to replace the original file with the temporary file
	private static void replaceFile(String originalFileName, String tempFileName) {
		try {
			Files.delete(Paths.get(originalFileName));
			Files.move(Paths.get(tempFileName), Paths.get(originalFileName));
			log.info("File replaced successfully!");
		} catch (IOException e) {
			e.printStackTrace();
			System.err.println("Error replacing the file!");
		}
	}

	// Helper method to generate the new output file name with '_New' suffix
	private static String getNewFileName(String inputFileName) {
		int lastDotIndex = inputFileName.lastIndexOf('.');
		if (lastDotIndex != -1) {
			String fileNameWithoutExtension = inputFileName.substring(0, lastDotIndex);
			String fileExtension = inputFileName.substring(lastDotIndex);
			return fileNameWithoutExtension + "_Actual" + fileExtension;
		} else {
			// If there's no file extension, simply append "_New" at the end
			return inputFileName + "_Actual";
		}
	}
}
