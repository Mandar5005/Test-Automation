package com.majesco.itaf.util;

import java.util.ArrayList;
import java.util.List;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.jcraft.jsch.ChannelSftp;
import com.jcraft.jsch.ChannelSftp.LsEntry;
import com.jcraft.jsch.ChannelSftp.LsEntrySelector;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.Session;
import com.jcraft.jsch.SftpException;

public class SFTPUtil {
	private static final Logger logger = LogManager.getLogger(SFTPUtil.class);
	public static final String SFTP_CHANNEL = "sftp";
	public static final String HOST_KEY_CHECKING = "StrictHostKeyChecking";
	public static final String HOST_KEY_CHECKING_VAL = "no";
	private final String SFTP_UID;
	private final String SFTP_PWD;
	private final String SFTP_HOST;
	private final int SFTP_PORT;

	public SFTPUtil(String sftpUid, String sftpPwd, String sftpHost, int sftpPort) {
		this.SFTP_UID = sftpUid;
		this.SFTP_PWD = sftpPwd;
		this.SFTP_HOST = sftpHost;
		this.SFTP_PORT = sftpPort;
	}

	public Session getSftpSession() throws JSchException {
		JSch jsch = new JSch();
		Session session = jsch.getSession(SFTP_UID, SFTP_HOST, SFTP_PORT);
		session.setConfig(HOST_KEY_CHECKING, HOST_KEY_CHECKING_VAL);
		session.setPassword(SFTP_PWD);

		session.connect();
		logger.info("SFTP Connected");

		return session;
	}

	public void downloadSFTPFile(Session session, String remotePath, String localPath)
			throws JSchException, SftpException {// PARAMS INCLUDING FILE NAME
		ChannelSftp sftpChannel = null;
		try {
			sftpChannel = (ChannelSftp) session.openChannel(SFTP_CHANNEL);
			sftpChannel.connect();

			sftpChannel.get(remotePath, localPath);
		} finally {
			if (null != sftpChannel && sftpChannel.isConnected()) {
				sftpChannel.disconnect();
				sftpChannel.exit();
			}
		}
	}

	public void moveSFTPFile(Session session, String remoteOldPath, String remoteNewPath, String fileName)
			throws JSchException, SftpException {// PARAMS INCLUDING FILE NAME
		ChannelSftp sftpChannel = null;
		try {
			sftpChannel = (ChannelSftp) session.openChannel(SFTP_CHANNEL);
			sftpChannel.connect();

			sftpChannel.rename(remoteOldPath + fileName, remoteNewPath + fileName);
		} finally {
			if (null != sftpChannel && sftpChannel.isConnected()) {
				sftpChannel.disconnect();
				sftpChannel.exit();
			}
		}
	}

	public void uploadSFTPFile(Session session, String remotePath, String localPath)
			throws JSchException, SftpException {
		ChannelSftp sftpChannel = null;
		try {
			sftpChannel = (ChannelSftp) session.openChannel(SFTP_CHANNEL);
			sftpChannel.connect();

			sftpChannel.put(localPath, remotePath);
		} finally {
			if (null != sftpChannel && sftpChannel.isConnected()) {
				sftpChannel.disconnect();
				sftpChannel.exit();
			}
		}
	}

	public List<String> getSFTPFiles(Session session, String directoryPath, String fileExtn)
			throws JSchException, SftpException {
		List<String> fileList = new ArrayList<>();
		ChannelSftp sftpChannel = null;

		try {
			sftpChannel = (ChannelSftp) session.openChannel(SFTP_CHANNEL);
			sftpChannel.connect();
			logger.info("SFTP Channel Connected");

			LsEntrySelector selector = new LsEntrySelector() {

				@Override
				public int select(LsEntry entry) {
					final String fileName = entry.getFilename();
					if (".".equals(fileName) || "..".equals(fileName)) {
						return CONTINUE;
					} else if (!entry.getAttrs().isLink() && !entry.getAttrs().isDir()
							&& fileName.toLowerCase().endsWith(fileExtn)) {
						fileList.add(fileName);
					}

					return CONTINUE;
				}
			};
			sftpChannel.ls(directoryPath, selector);
		} finally {
			if (null != sftpChannel && sftpChannel.isConnected()) {
				sftpChannel.disconnect();
				sftpChannel.exit();
			}
		}

		return fileList;
	}
}
