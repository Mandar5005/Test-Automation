package com.majesco.itaf.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Vector;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.jcraft.jsch.Channel;
import com.jcraft.jsch.ChannelSftp;
import com.jcraft.jsch.ChannelSftp.LsEntry;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.Session;
import com.jcraft.jsch.SftpException;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.main.WebHelper;
import com.majesco.itaf.main.WebHelperBilling;
import com.majesco.itaf.webservice.WebService;

public class OutboundTransferFiles

{
	private final static Logger log = LogManager.getLogger(OutboundTransferFiles.class.getName());
	public static String FlatFileResponse = null;
	public static int failedSC = -1;
	public static String report_msg;
	public static String report_status;

	private static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private static MainController controller = ObjectFactory.getMainController();

	public void copyFiles_Nav(String remote_path_out, String local_path, String file_to_be_converted, String extension,
			String xfl_filename, String archive, String validateTag, String validationMsg, String file_cycledate)

	{
		if (Config.EnvironmentType.equalsIgnoreCase("Windows")) {
			log.info("Environment Type is: " + Config.EnvironmentType);
			try {

				copyFilesWindows(remote_path_out, local_path, file_to_be_converted, extension, xfl_filename, archive,
						validateTag, validationMsg, file_cycledate);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		} else if (Config.EnvironmentType.equalsIgnoreCase("Linux")) {
			log.info("Environment Type is: " + Config.EnvironmentType);
			try {
				copyFilesLinux(remote_path_out, local_path, file_to_be_converted, extension, xfl_filename, archive,
						validateTag, validationMsg, file_cycledate);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		} else
			log.info("Invalid Environment Type or it is Null,kindly check config.xlsm");
	}

	public void copyFilesWindows(String copyFrom, String copyTo, String xmlFileName, String extn, String xflFile,
			String archive, String validateTag, String validationMsg, String file_cycledate) throws IOException {

		File remoteFile = new File(copyFrom);
		File[] listOfFiles = remoteFile.listFiles();

		@SuppressWarnings("unused")
		String searchFile;
		if (WebHelperBilling.stransactionType.contains("Outbound")) {
			searchFile = xmlFileName + "_";
		} else {
			searchFile = xmlFileName + "_" + file_cycledate;
		}

		for (int i = 0; i < listOfFiles.length; i++) {

			Boolean b = listOfFiles[i].getName().endsWith(extn);

			if (b == true) {
				try {
					File ff_dir = new File(copyTo);
					FileUtils.forceMkdir(ff_dir);
					log.info("Directory created " + ff_dir);
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					log.info("Creating direcotory failed");
				}

				String fileName = listOfFiles[i].getName();
				log.info(listOfFiles[i].getName());
				String response_filename = fileName.replace(extn, ".xml");
				@SuppressWarnings("unused")
				String scno_responseFile = response_filename.substring(0, (response_filename.indexOf('_')));
				response_filename = copyTo + "/" + response_filename;

				try {
					InputStream in = new FileInputStream(new File(copyFrom + "/" + listOfFiles[i].getName()));
					OutputStream out = new FileOutputStream(new File(copyTo + "/" + listOfFiles[i].getName()));

					byte[] buffer = new byte[1024];

					int len;

					while ((len = in.read(buffer)) > 0) {
						out.write(buffer, 0, len);
					}

					in.close();
					out.close();

					// Flat to Xml conversion

					FlatToXmlConversion fx = new FlatToXmlConversion();
					fx.convertFlatToXml(copyTo, fileName, extn, xflFile);

					// sonali- to save outbound file as response 08/23/2019
					File oldFile = new File(response_filename);
					File newFile = new File(copyTo + "/" + xmlFileName + "_" + file_cycledate + "_Response.xml");

					log.info("Oubound File craeted : " + copyTo + "/" + xmlFileName + "_" + file_cycledate
							+ "_Response.xml");

					if (oldFile.renameTo(newFile)) {
						log.info("Renamed");
					} else {
						log.info("Failed Renamed");
					}

				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

		// archiving files
		log.info("Do you want to archive Files? Y or N");
		if (archive.equalsIgnoreCase("Y")) {
			ArchiveFiles af = new ArchiveFiles();
			af.archiveWindows(copyFrom, extn);
		}
	}

	public void copyFilesLinux(String copyFrom, String copyTo, String xmlFileName, String extn, String xflFile,
			String archive, String validateTag, String validationMsg, String file_cycledate) throws IOException {

		// Get data from Config
		String hostName = Config.flatFileHostName;
		String userName = Config.flatFileUserName;
		String password = Config.flatFilePassword;
		@SuppressWarnings("unused")
		int port = Integer.parseInt(Config.flatFilePort);

		//

		Boolean flatfile_found = false;

		JSch jsch = new JSch();

		Session session = null;
		log.info("Trying to connect.....");

		try {
			session = jsch.getSession(userName, hostName, 22);
			session.setConfig("StrictHostKeyChecking", "no");
			session.setConfig("PreferredAuthentications", "password"); // For keberos error

			session.setPassword(password);
			session.connect();

			Channel channel = session.openChannel("sftp");
			channel.connect();
			ChannelSftp sftpChannel = (ChannelSftp) channel;

			log.info("Done !!");

			// FlatFile copying from remote Linux server to local machine

			sftpChannel.cd(copyFrom);
			@SuppressWarnings("rawtypes")
			Vector filelist = sftpChannel.ls(copyFrom);

			report_msg = "";
			report_status = "";
			@SuppressWarnings("unused")
			String searchFile;

			if (WebHelperBilling.stransactionType.contains("Outbound")) {
				searchFile = xmlFileName + "_";
			} else {
				searchFile = xmlFileName + "_" + file_cycledate;
			}

			for (int i = 0; i < filelist.size(); i++) {
				LsEntry entry = (LsEntry) filelist.get(i);

				Boolean b = entry.getFilename().endsWith(extn);

				if (b == true) {

					String fileName = entry.getFilename();
					String response_filename = fileName.replace(extn, ".xml");

					String scno_responseFile;
					if (response_filename.contains("_")) {
						scno_responseFile = response_filename.substring(2, (response_filename.indexOf('_')));
						// Meghna---for Flatfile Sc failure

					} else {
						scno_responseFile = response_filename;
					}

					response_filename = copyTo + "/" + response_filename;

					// Converting Flat file to xml

					if (WebHelperBilling.stransactionType.contains("Outbound")) {
						if (fileName.startsWith(xmlFileName + "_")) {
							flatfile_found = true;
							try {
								File ff_dir = new File(copyTo);
								FileUtils.forceMkdir(ff_dir);
								log.info("Directory created " + ff_dir);
							} catch (Exception e) {
								log.error(e.getMessage(), e);
								log.info("Creating direcotory failed");
							}

							sftpChannel.get(entry.getFilename(), copyTo);
							FlatToXmlConversion fx = new FlatToXmlConversion();
							fx.convertFlatToXml(copyTo, fileName, extn, xflFile);

							File oldFile = new File(response_filename);
							File newFile = new File(
									copyTo + "/" + xmlFileName + "_" + file_cycledate + "_Response.xml");

							log.info("Oubound File created : " + copyTo + "/" + xmlFileName + "_" + file_cycledate
									+ "_Response.xml");

							if (oldFile.renameTo(newFile)) {
								log.info("Renamed");
							} else {
								log.info("Failed Renamed");
							}

							break;
						} else if (fileName.startsWith(xmlFileName)) {
							flatfile_found = true;
							try {
								File ff_dir = new File(copyTo);
								FileUtils.forceMkdir(ff_dir);
								log.info("Directory created " + ff_dir);
							} catch (Exception e) {
								log.error(e.getMessage(), e);
								log.info("Creating direcotory failed ");
							}

							sftpChannel.get(entry.getFilename(), copyTo);
							// Added by samir
							if (WebHelperBilling.stransactionType.contains("Convert")) {
								String newFileName = copyTo + "\\" + fileName;
								if (WebHelperBilling.stransactionType.contains("Collection")) {
									boolean isEmpty = VariableToFixedLengthConverter.convertToFixedLength(newFileName,
											"\\^");
									if (isEmpty == true) {
										System.out.println("Found Blank File, Contiue with next file");
										continue;
									}
								} else {
									boolean isEmpty = VariableToFixedLengthConverter.convertToFixedLength(newFileName,
											",");
									if (isEmpty == true) {
										System.out.println("Found Blank File, Continue with next file");
										continue;
									}
								}
							}
							// end

							FlatToXmlConversion fx = new FlatToXmlConversion();
							fx.convertFlatToXml(copyTo, fileName, extn, xflFile);

							File oldFile = new File(response_filename);
							File newFile = new File(
									copyTo + "/" + xmlFileName + "_" + file_cycledate + "_Response.xml");

							log.info("Oubound File created : " + copyTo + "/" + xmlFileName + "_" + file_cycledate
									+ "_Response.xml");

							if (oldFile.renameTo(newFile)) {
								log.info("Renamed");
							} else {
								log.info("Failed Renamed");
							}

							break;
						}
					} else if (fileName.startsWith(xmlFileName + "_" + file_cycledate)) // ---For
					// Inbound
					// Response
					// if(fileName.startsWith(searchFile))
					{

						flatfile_found = true;
						sftpChannel.get(entry.getFilename(), copyTo);
						FlatToXmlConversion fx = new FlatToXmlConversion();
						fx.convertFlatToXml(copyTo, fileName, extn, xflFile);

						// Meghna--Verify Response XML//
						try {
							String Tag_Name = "*";
							String[] Node_Value = new String[2];
							Node_Value[1] = "ErrorCode";
							Node_Value[0] = "Description";

							int index = 0;

							// WebService.getXMLResponseTagValue(responseXml,Tag_Name,Node_Value1,index);
							FlatFileResponse = WebService.getXMLResponseStatusFlatFile(response_filename, Tag_Name,
									Node_Value, index, validateTag, validationMsg);

							if (FlatFileResponse.equalsIgnoreCase("SUCCESS")) {

								// iTAFSeleniumWeb.WebDriver.report.setStrMessage("SUCCESS");//Mandar--
								report_msg = "SUCCESS:" + response_filename;
								report_status = "PASS";
								WebHelper.success = true;
								WebHelper.description = "SUCCESS";
							}

							else if (FlatFileResponse.equalsIgnoreCase("FAILED") && WebHelper.success != true) {
								// WebService.getXMLResponseTagValue(responseXml,Tag_Name,Node_Value2,index);
								Node_Value[1] = "ErrorCode";
								WebHelper.FailedResponseTagValue = WebService.getXMLResponseTagValue(response_filename,
										Tag_Name, Node_Value[1], index);
								if (webDriver.getReport().getMessage() == null
										|| webDriver.getReport().getMessage() == "") {
									// to do//Mandar --20/09/2017
								}
								// Meghna--Flat File
								report_msg = "FAILED:" + response_filename;
								report_status = "FAIL";
								// Meghna--Flat File

								WebHelper.description = WebHelper.FailedResponseTagValue;
								WebHelper.failed = true;
							}
							// String Node_Value1 = "ProcessStatusFlag";
							// String ProcessStatusFlag =
							// WebService.getXMLResponseTagValue(responseXml,Tag_Name,Node_Value1,index);
							log.info("Tag value from Response file is:" + WebHelper.FailedResponseTagValue);

							if (FlatFileResponse == null || FlatFileResponse.equalsIgnoreCase("FAILED")
									|| FlatFileResponse.equalsIgnoreCase("BLANK")) {
								// Meghna--For marking SC in MainCont in case of
								// failures
								failedSC = Integer.parseInt(scno_responseFile);

								controller.recoveryhandler();
								log.info("Flat File response failed for : " + failedSC);
							}
							WebHelper.success = false;
						}

						catch (Exception e) {
							log.error(e.getMessage(), e);
						}
						break;
					}

				}
			}

			if (!flatfile_found) {
				report_msg = report_msg + ("FAILED: Flat File not found");
				// report_status = "FAIL";
			}

			// Meghna---Write Final message and status to report//
			if (webDriver.getReport().getMessage() == null || webDriver.getReport().getMessage() == "") {
				webDriver.getReport().setMessage(report_msg);
			}

			if (report_status.contains("FAIL")) {
				webDriver.getReport().setStatus("FAIL");
			} else {
				webDriver.getReport().setStatus("PASS");
			}

			// ---Archiving------//
			log.info("Do you want to archive Files? Y or N");
			// String archive = sc.nextLine();

			if (archive.equalsIgnoreCase("Y")) {
				ArchiveFiles af = new ArchiveFiles();
				af.archiveLinux(copyFrom, extn);
			}
			sftpChannel.disconnect();
			sftpChannel.exit();
			session.disconnect();

		} catch (JSchException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		} catch (SftpException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}

	}

}
