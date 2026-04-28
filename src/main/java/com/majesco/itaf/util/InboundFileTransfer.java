package com.majesco.itaf.util;

import java.io.File;
import java.io.IOException;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.jcraft.jsch.Channel;
import com.jcraft.jsch.ChannelSftp;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.Session;
import com.jcraft.jsch.SftpException;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;

public class InboundFileTransfer {

	private final static Logger log = LogManager.getLogger(InboundFileTransfer.class.getName());
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	public Boolean ValidateSuccess = false;

	public void copyFilesToRemote_Nav(String local_path, String remote_path_in, String file_to_be_converted,
			String string, String xfl_filename)

	{
		if (Config.EnvironmentType.equalsIgnoreCase("Windows")) {
			log.info("Environment Type is: " + Config.EnvironmentType);
			copyFilesToRemoteWindows(local_path, remote_path_in, file_to_be_converted, string, xfl_filename);

		} else if (Config.EnvironmentType.equalsIgnoreCase("Linux")) {
			log.info("Environment Type is: " + Config.EnvironmentType);
			copyFilesToRemoteLinux(local_path, remote_path_in, file_to_be_converted, string, xfl_filename);

		} else
			log.info("Invalid Environment Type or it is Null,kindly check config.xlsm");
	}

	public void copyFilesToRemoteWindows(String copyFromLocal, String copyToRemote, String xmlFileName, String extn,
			String xflFile) {

		try {
			File ff_dir = new File(copyToRemote);
			if (ff_dir.exists()) {
				log.info("Directory Already Exist " + ff_dir);
			} else {
				FileUtils.forceMkdir(ff_dir);
				log.info("Directory created " + ff_dir);
			}

			// call to xml converter
			xmlFileName = xmlFileName + ".xml";
			XmlToFlatConversion.convertXmlToFlat(copyFromLocal, copyToRemote, xmlFileName, extn, xflFile);

			// Copy flat from Local to remote windows server
			xmlFileName = xmlFileName.replaceAll(".xml", extn);
			File sourceFile = new File(copyFromLocal + "/" + xmlFileName);

			File targetFile = new File(copyToRemote + "/" + xmlFileName);
			sourceFile.renameTo(targetFile);

		} catch (IOException e) {
			// log.error(e.getMessage(), e);
			log.info("Unable to create directory ");
			webDriver.getReport().setStatus("FAIL");
			webDriver.getReport().setMessage("Direcorty Does Not Exist on server");

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
			// log.info("Unable to create directory ");
			webDriver.getReport().setStatus("FAIL");
			webDriver.getReport().setMessage("Direcorty Does Not Exist on server");

		}
	}

	public void copyFilesToRemoteLinux(String copyFromLocal, String copyToRemote, String xmlFileName, String extn,
			String xflFile) {

		String hostName = Config.flatFileHostName;
		String userName = Config.flatFileUserName;
		String password = Config.flatFilePassword;
		int port = Integer.parseInt(Config.flatFilePort);
		JSch jsch = new JSch();
		Session session = null;
		System.out.println("Trying to connect.....");

		try {

			session = jsch.getSession(userName, hostName, port);
			session.setConfig("StrictHostKeyChecking", "no");
			session.setConfig("PreferredAuthentications", "password");//For kerberos error.
			session.setPassword(password);
			session.connect();
			Channel channel = session.openChannel("sftp");
			channel.connect();
			ChannelSftp sftpChannel = (ChannelSftp) channel;
			System.out.println("Done !!");

			// Convert xml to flat
			sftpChannel.lcd(copyFromLocal);

			// call to xml converter
			xmlFileName = xmlFileName + ".xml";
			XmlToFlatConversion.convertXmlToFlat(copyFromLocal, copyToRemote, xmlFileName, extn, xflFile);

			// Copy flat from Local to remote linux server
			sftpChannel.cd(copyToRemote);
			sftpChannel.put(copyFromLocal + "/" + xmlFileName.replaceAll(".xml", extn), copyToRemote);

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
