package com.majesco.itaf.util;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.jcraft.jsch.*;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.WebHelper;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Vector;

public class FTPFileTransferLinux {

	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private final static Logger log = LogManager.getLogger(WebHelper.class.getName());
	public static String file_names;
	public static String file_name;

	public static void FTPLinux(String filesnames, String destinationfolder, String cycleDate) throws Exception {
		cycleDate = cycleDate.replace("/", "_");
		JSch jsch = new JSch();
		Session session = null;
		ChannelSftp channelSftp = null;
		String portString = Config.ftpserverport.trim();
		int port = Integer.parseInt(portString);

		log.info("Initiating SSH (secure shell)connection with the FTP server...");
		try {
			log.info("FTP Server Host : " + Config.ftpserverhost);
			session = jsch.getSession(Config.ftpserveruserid.trim(), Config.ftpserverhost.trim(), port);
			session.setPassword(Config.ftpserverpassword.trim());
			session.setConfig("StrictHostKeyChecking", "no");
			session.connect();

			channelSftp = (ChannelSftp) session.openChannel("sftp");
			channelSftp.connect();
			log.info("Connection successful");

			log.info("Checking if Destination Folder exists...");
			try {
				SftpATTRS attrs = channelSftp.stat(destinationfolder);
				if (attrs.isDir()) {
					log.info("Destination folder exists: [" + destinationfolder + "]");
				} else {
					throw new SftpException(ChannelSftp.SSH_FX_FAILURE, "Not a directory: " + destinationfolder);
				}
			} catch (SftpException ex) {
				// Folder does not exist, throw an exception
				log.info("Destination folder does not exist: [" + destinationfolder + "]");
				throw new SftpException(ChannelSftp.SSH_FX_FAILURE,
						"Destination folder does not exist: " + destinationfolder);

			}
			String[] fileNamesArray = filesnames.trim().split("\\s+");
			log.info("Initiating File Upload...");
			for (String fileName : fileNamesArray) {
				fileName = fileName.trim(); // Trim to remove any leading/trailing spaces
				String localFilePath = Config.inputDataFilePath + "CopyFlatFile\\FlatFiles\\" + cycleDate + "\\"
						+ fileName;

				// Upload the file to the remote server
				channelSftp.put(localFilePath, destinationfolder);

				// Print a message for each file upload
				log.info("File uploaded: [" + fileName + "]");
			}

			// Add folder refresh code here
			log.info("Verifying if the files exists in the Destination Folder...");
			refreshFolder(filesnames, destinationfolder);

			log.info("All the flatfiles are uploaded successfully.");
			// webDriver.getReport().setMessage("Files Copied :" + filesnames);

		} catch (JSchException | SftpException e) {
			e.printStackTrace();
			log.info("Failed to upload the file.");
		} finally {

			if (channelSftp != null) {
				channelSftp.disconnect();
			}
			if (session != null) {
				session.disconnect();
			}
			log.info("SSH (secure shell)connection disconnected successfully");
		}
	}

	private static void refreshFolder(String filesnames, String destinationfolder) {
		String portString = Config.ftpserverport.trim();
		int port = Integer.parseInt(portString);

		try {
			JSch jsch = new JSch();
			Session session = jsch.getSession(Config.ftpserveruserid, Config.ftpserverhost, port);
			session.setPassword(Config.ftpserverpassword);
			session.setConfig("StrictHostKeyChecking", "no");
			session.connect();

			ChannelSftp channelSftp = (ChannelSftp) session.openChannel("sftp");
			channelSftp.connect();

			// Upload your files here...
			// Check if files exist with a delay
			int maxAttempts = 3; // Maximum attempts (adjust as needed)
			int delayMillis = 5000; // Delay in milliseconds (adjust as needed)

			for (int attempt = 1; attempt <= maxAttempts; attempt++) {
				List<String> csvFiles = getCsvFilesInDirectory(channelSftp, destinationfolder);
				if (!csvFiles.isEmpty()) {
					log.info("Below is the list of all the csv files present in the remote directory:");
					for (String csvFile : csvFiles) {
						log.info("- " + csvFile);
					}
					// Perform your refresh actions here...
					// break;
				} else {
					log.info("CSV files not detected. Waiting for the next attempt...");
					Thread.sleep(delayMillis);
				}

				if (areAllFilesPresent(channelSftp, destinationfolder, filesnames)) {
					log.info("Uploaded flatfiles are present in the remote directory.");
					// Perform your refresh actions here...
					break;
				} else {
					log.info("Not all the uploaded flatfiles are present. Waiting for the next attempt...");
					Thread.sleep(delayMillis);
				}
			}

			// Disconnect
			channelSftp.disconnect();
			session.disconnect();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static List<String> getCsvFilesInDirectory(ChannelSftp channelSftp, String remoteDirectory)
			throws SftpException {
		List<String> csvFiles = new ArrayList<>();
		try {
			// Attempt to list files in the remote directory
			Vector<ChannelSftp.LsEntry> lsEntries = channelSftp.ls(remoteDirectory);
			for (ChannelSftp.LsEntry entry : lsEntries) {
				if (!entry.getFilename().equals(".") && !entry.getFilename().equals("..")
						&& entry.getFilename().toLowerCase().endsWith(".csv")) {
					csvFiles.add(entry.getFilename());
				}
			}
		} catch (SftpException e) {
			// Directory or files not found
		}
		return csvFiles;
	}

	private static boolean areAllFilesPresent(ChannelSftp channelSftp, String remoteDirectory, String filesnames)
			throws SftpException {
		List<String> expectedFiles = Arrays.asList(filesnames.trim().split("\\s+"));
		List<String> actualFiles = getCsvFilesInDirectory(channelSftp, remoteDirectory);

		// Check if all expected files are present in the actual list
		return actualFiles.containsAll(expectedFiles);

	}
}
