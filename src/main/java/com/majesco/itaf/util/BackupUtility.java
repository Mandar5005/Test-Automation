package com.majesco.itaf.util;

import java.io.*;
import java.util.Arrays;
import java.util.List;
import java.util.zip.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.itaf.main.Config;
import java.text.SimpleDateFormat;
import java.util.Date;

public class BackupUtility {

	private final static Logger log = LogManager.getLogger(BackupUtility.class);

	public static String initBackup(String backupFiles, String backupLocation) throws IOException {

		String[] backupFileArray = backupFiles.split("\\s*,\\s*");
		List<String> backupFileList = Arrays.asList(backupFileArray);
		String zipFileName = zipAndBackupFiles(backupFileList, backupLocation);
		return zipFileName;
	}

	private static String zipAndBackupFiles(List<String> fileList, String backupLocation) throws IOException {
		String timeString = new SimpleDateFormat("dd.MM.yyyy-HH.mm").format(new Date());
		String zipFileName = backupLocation + "\\iTAFBackup_" + Config.projectName + "_" + timeString + ".zip";
		FileOutputStream fos = new FileOutputStream(zipFileName);
		ZipOutputStream zipOut = new ZipOutputStream(fos);
		for (String filename : fileList) {
			File file = new File(filename);
			zipFile(file, file.getName(), zipOut);
		}
		zipOut.close();
		fos.close();
		return zipFileName;
	}

	private static void zipFile(File fileToZip, String fileName, ZipOutputStream zipOut) {
		if (fileToZip.isHidden())
			return;

		try {
			if (fileToZip.isDirectory()) {
				if (fileName.endsWith("/")) {
					zipOut.putNextEntry(new ZipEntry(fileName));
					zipOut.closeEntry();
				} else {
					zipOut.putNextEntry(new ZipEntry(fileName + "/"));
					zipOut.closeEntry();
				}
				File[] children = fileToZip.listFiles();
				for (File childFile : children) {
					zipFile(childFile, fileName + "/" + childFile.getName(), zipOut);
				}
				return;
			}
			FileInputStream fis = new FileInputStream(fileToZip);
			ZipEntry zipEntry = new ZipEntry(fileName);
			zipOut.putNextEntry(zipEntry);
			byte[] bytes = new byte[1024];
			int length;
			while ((length = fis.read(bytes)) >= 0) {
				zipOut.write(bytes, 0, length);
			}
			fis.close();
		} catch (Exception exObj) {
			log.info("Exception occured while zipping file: " + fileToZip + "--" + exObj.getMessage());
		}
	}
}