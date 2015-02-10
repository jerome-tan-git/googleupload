package com.ask.inv.Util;

import java.io.IOException;
import java.util.Date;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Address;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.util.ByteArrayDataSource;
import org.apache.log4j.Logger;

public class MailUtilCom {
	private static final Logger logger = Logger.getLogger(MailUtilCom.class);
	private static MailUtilCom mail = null;

	private String host;
	private String user;
	private String password;
	private String mail_auth;
	private Session session;
	
	private MailUtilCom(String host, String user, String password) {
		this.host = host;
		this.user = user;
		this.password = password;
		this.mail_auth = "false";
		init();
	}
	
	private MailUtilCom(String host, String user, String password, String mail_auth) {
		this.host = host;
		this.user = user;
		this.password = password;
		this.mail_auth = mail_auth;
		init();
	}

	private void init() {
		Properties props = System.getProperties();
		if("ture".equalsIgnoreCase(mail_auth)) {
			props.put("mail.smtp.auth", "true");
			Authenticator auth = new Email_Autherticator(user,password); 
			session = Session.getInstance(props, auth);
		} else {
			props.put("mail.smtp.host", host);
			props.put("mail.smtp.auth", "false");
			session = Session.getInstance(props);
		}
		
	}

	public static MailUtilCom newInstance(String host, String user, String password) {
		if (mail == null) {
			mail = new MailUtilCom(host, user, password);
		}
		return mail;
	}
	
	public static MailUtilCom newInstance(String host, String user, String password, String mail_auth) {
		if (mail == null) {
			mail = new MailUtilCom(host, user, password, mail_auth);
		}
		return mail;
	}

	public void send(String from, String to, String subject, String content)
			throws AddressException, MessagingException, IOException {
		send(from, to, null, null, subject, content);
	}

	public void send(String from, String to, String cc, String bcc,
			String subject, String content) throws AddressException,
			MessagingException, IOException {
		if(from==null || from.length()==0){
			throw new MessagingException("From mail address is null!");
		}
		if(to==null || to.length()==0){
			throw new MessagingException("To mail address is null!");
		}
		Message msg = new MimeMessage(session);
		if (from != null) {
			msg.setFrom(new InternetAddress(from));
			msg.setReplyTo(new Address[] { new InternetAddress(from) });
		}
		
		msg.setRecipients(Message.RecipientType.TO,
					InternetAddress.parse(to, false));
		if (cc != null) {
				msg.setRecipients(Message.RecipientType.CC,
					InternetAddress.parse(cc, false));
		}
		if (bcc != null) {
			msg.setRecipients(Message.RecipientType.BCC,
					InternetAddress.parse(bcc, false));
		}

		msg.setSubject(subject);
		msg.setDataHandler(new DataHandler(new ByteArrayDataSource(content,
				"text/html")));

		msg.setHeader("X-Mailer", "Jmail");
		msg.setSentDate(new Date());

		Transport.send(msg);
	}
	
	public int send(String from, String to, String cc, String bcc,
			String subject, String body, String attachments) {
		return send(from, to, cc, bcc, subject, body, attachments, null);
	}
	
	public int send(String from, String to, String cc, String bcc,
			String subject, String body, String attachments, String contentType) {
		int errorStatus = 0;

		try {
			Message msg = new MimeMessage(session);

			msg.addFrom(InternetAddress.parse(from));
			msg.addRecipients(Message.RecipientType.TO, InternetAddress
					.parse(to));

			if (null != cc) {
				msg.addRecipients(Message.RecipientType.CC, InternetAddress
						.parse(cc));
			}

			if (null != bcc) {
				msg.addRecipients(Message.RecipientType.BCC, InternetAddress
						.parse(bcc));
			}

			// Attach the files to the message
			if (null != attachments) {
				Multipart mp = new MimeMultipart();

				int startIndex = 0;
				int posIndex = 0;
				while (-1 != (posIndex = attachments.indexOf("///", startIndex))) {
					// Create and fill other message parts;
					MimeBodyPart mbp = new MimeBodyPart();
					FileDataSource fds = new FileDataSource(attachments
							.substring(startIndex, posIndex));
					
					logger.info("++++++++++++++ attachments: "
							+ attachments.substring(startIndex, posIndex));
					
					mbp.setDataHandler(new DataHandler(fds));
					mbp.setFileName(fds.getName());
					mp.addBodyPart(mbp);
					posIndex += 3;
					startIndex = posIndex;
				}
				// Last, or only, attachment file;
				if (startIndex < attachments.length()) {
					MimeBodyPart mbp = new MimeBodyPart();
					FileDataSource fds = new FileDataSource(attachments
							.substring(startIndex));
					
					logger.info(".............. attachments: "
							+ attachments.substring(startIndex));
					
					mbp.setDataHandler(new DataHandler(fds));
					mbp.setFileName(fds.getName());
					mp.addBodyPart(mbp);
				}

				BodyPart boaypart = new MimeBodyPart();
				boaypart.setText(body);
				boaypart.setContent(body,"text/html;charset=utf-8");
				mp.addBodyPart(boaypart);
				
				msg.setContent(mp);
				
			} else {
				if (contentType != null) {
					msg.setContent(body, contentType);
				} else {
					msg.setText(body);
				}
			}

			// Set the Date: header
			// Subject field
			msg.setSubject(subject);
			
			logger.info(".............. subject: " + subject);

			msg.setSentDate(new Date());
			msg.saveChanges();

			// Send the message;
			Transport.send(msg);
		} catch (MessagingException MsgException) {
			MsgException.printStackTrace();
			errorStatus = 1;
		}
		return errorStatus;
	}
	
	class Email_Autherticator extends Authenticator {
		private String user;
		private String pwd;

		public Email_Autherticator() {
			super();
		}

		public Email_Autherticator(String user, String pwd) {
			super();
			this.user = user;
			this.pwd = pwd;
		}

		public PasswordAuthentication getPasswordAuthentication() {
			return new PasswordAuthentication(user, pwd);
		}
	}
}
