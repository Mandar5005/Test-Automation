package com.majesco.itaf.outlookmail;

import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.net.InetAddress;
import java.net.URL;
import java.net.UnknownHostException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Base64;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.servlet.http.HttpServletRequest;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.itaf.main.Config;

public class SendMail {
	private final static Logger log = LogManager.getLogger(SendMail.class);

	public static void SendMailWebOutlook(String target_path, String summaryReportName) throws Exception {

		// System.out.println("Sending mail...");
		log.info("Sending Web Outlook Email...");
		log.info("Initiating SMTP connection...");
		Properties props = new Properties();
		props.put("mail.smtp.host", "smtp.office365.com");
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.port", "587");
		props.put("mail.smtp.starttls.enable", "true");
		log.info("Initiating UserID/PasssAuthentication...");

		Session mailSession = Session.getDefaultInstance(props, new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(Config.outlookUsername, Config.outlookApppassword);
			}
		});

		try {
			Message message = new MimeMessage(mailSession);
			message.setFrom(new InternetAddress(Config.outlookemailFrom));
			message.setSubject(Config.outlookemailSubject);
			message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(Config.outlookemailTo));
			message.setRecipients(Message.RecipientType.CC, InternetAddress.parse(Config.outlookemailCC));

			MimeMultipart multipart = new MimeMultipart("related");

			// first part (the html)
			BodyPart messageBodyPart = new MimeBodyPart();

			@SuppressWarnings("deprecation")
			URL fileURL = new URL("file://" + target_path);
			
			// Read the image file into a byte array
						byte[] imageBytes = Files.readAllBytes(Paths.get(Config.executionStatusReportPath + "\\Outlook\\" + summaryReportName));
						
						// Convert the image data to a base64-encoded string
			            String imageData = Base64.getEncoder().encodeToString(imageBytes);
			            
			            // Embed the image data directly in the HTML code
			            //String htmlContent = "<html><body><img src='data:image/jpeg;base64," + imageData + "' width='100%' height='100%'></body></html>\";"
			            String htmlContent = "<html><body><img src='data:image/jpeg;base64," + imageData + "' width='100%' height='100%'></body></html>";

			String htmlText = "<P>Hi All,<P>" + "<P>PFA run results.</P>"
					+ "<P>Below given is the status for the unattended run.</P>"
					+ "<P> The results and input files are shared at the below given location:.</P>" 
					+ "<p><a href=\"" + fileURL + "\">" + fileURL + "</a></p>"
					+ "<P> <P>"
					+  htmlContent;

			//+ "<html><body><img src='data:image/jpeg;base64," + imageData + "width=\"100%\" height=\"100%\"'></body></html>";
			//+ "<img src=\"cid:image\" width=\"100%\" height=\"100%\">";
			messageBodyPart.setContent(htmlText, "text/html");

			// add it
			multipart.addBodyPart(messageBodyPart);

			// second part (the image)
			//messageBodyPart = new MimeBodyPart();
			//DataSource fds = new FileDataSource(Config.executionStatusReportPath + "\\Outlook\\" + summaryReportName);
			
			//messageBodyPart.setDataHandler(new DataHandler(fds));
			//messageBodyPart.setHeader("Content-ID", "<image>");
			
			// messageBodyPart.setHeader("Content-Transfer-Encoding", "base64");
			// messageBodyPart.setDisposition(MimeBodyPart.INLINE+ "; size=0");
			// messageBodyPart.setDisposition(MimeBodyPart.INLINE);

			// add image to the multipart
			multipart.addBodyPart(messageBodyPart);

			// Add multiple
			// attachments//"D:\\Worspace_SVN_Live_Branch\\iTAF_Live_Branch\\iTAF_SVN\\Claim_HO\\Claim_Loss\\Resources\\Results\\SummaryResults.csv","SummaryResults.csv");
			addAttachment(multipart, Config.resultOutput, "SummaryResults.csv");
			addAttachment(multipart, Config.executionStatusReportUtility, "ExecutionStatusReport.xlsm");
			addAttachment(multipart, Config.controllerFilePath, "MainController.xls");
			addAttachment(multipart, Config.transactionInfo, "UniqueNumber.xls");
			
			message.setContent(multipart);
			// Send message
			Transport.send(message);

			log.info("Email sent successfully....");
			// System.out.println("Email sent successfully....");

		} catch (MessagingException e) {
			throw new RuntimeException(e);
		}
	}

	private static void addAttachment(Multipart multipart, String Attachment, String fileName) throws Exception {
		DataSource source = new FileDataSource(Attachment);
		BodyPart messageBodyPart = new MimeBodyPart();
		messageBodyPart.setDataHandler(new DataHandler(source));
		messageBodyPart.setFileName(fileName);
		multipart.addBodyPart(messageBodyPart);
	}

	@SuppressWarnings("unused")
	private static BufferedImage resize(BufferedImage img, int height, int width) {
		Image tmp = img.getScaledInstance(width, height, Image.SCALE_SMOOTH);
		BufferedImage resized = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);

		Graphics2D g2d = resized.createGraphics();
		g2d.drawImage(tmp, 0, 0, null);
		g2d.dispose();
		return resized;
	}

	public static String addAddress(HttpServletRequest request) {
		String fullUrl = request.getRequestURL().toString();
		String url = fullUrl.split("/addAddress/")[1];
		System.out.println(url);
		return url;
	}

	public static String hostname() {
  	String hostname = "Unknown";

	try
	{
	    InetAddress addr;
	    addr = InetAddress.getLocalHost();
	    hostname = addr.getHostName();
	    log.info(hostname);
	 
	}
	catch (UnknownHostException ex)
	{
	    System.out.println("Hostname can not be resolved");
	}
	return hostname;
	
	
}
}