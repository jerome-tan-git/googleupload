package com.ask.inv.Util;

import java.io.File;
import java.io.IOException;
import java.io.StringWriter;
import java.util.HashMap;
import java.util.Properties;
import org.apache.log4j.Logger;
import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.Template;
import freemarker.template.TemplateException;

public class SendMailUtil {
	private static Logger logger = Logger.getLogger (SendMailUtil.class.getName());
	private Properties prop;
	private static Configuration file_cfg = new Configuration();
	private String mail_smtp_host;
	private String mail_auth;
	private String mail_user;
	private String mail_password;
	private String mail_from;
	private String mail_to;
	private String mail_cc;
	private String mail_subject;
	private String mail_content_hi;
	private String mail_content;
	private String mail_content_signature;
	private String mail_template;
	private String template_path;
	
	public SendMailUtil(Properties prop) {
		this.prop = prop;
		mail_smtp_host = prop.getProperty("mail.smtp.host");
		mail_auth = prop.getProperty("mail.password.auth");
		mail_user = prop.getProperty("mail.user");
		mail_password = prop.getProperty("mail.password");
		mail_from = prop.getProperty("mail.from");
		mail_to = prop.getProperty("mail.to");
		mail_cc = prop.getProperty("mail.cc");
		mail_subject = prop.getProperty("mail.subject");
		mail_content_hi = prop.getProperty("mail.content.hi");
		mail_content = prop.getProperty("mail.content");
		mail_content_signature = prop.getProperty("mail.content.signature");
		mail_template = prop.getProperty("template.email");
		template_path = prop.getProperty("freemarker.template.path");
		setFile_cfg();
	}
	
	public SendMailUtil(String fileName) {
		prop = new Properties();
		try {
			prop.load(this.getClass().getClassLoader().getResourceAsStream("GAMarketingDataRequestReport.properties"));
		} catch (IOException e) {
			e.printStackTrace();
		}
		mail_smtp_host = prop.getProperty("mail.smtp.host");
		mail_auth = prop.getProperty("mail.password.auth");
		mail_user = prop.getProperty("mail.user");
		mail_password = prop.getProperty("mail.password");
		mail_from = prop.getProperty("mail.from");
		mail_to = prop.getProperty("mail.to");
		mail_cc = prop.getProperty("mail.cc");
		mail_subject = prop.getProperty("mail.subject");
		mail_content_hi = prop.getProperty("mail.content.hi");
		mail_content = prop.getProperty("mail.content");
		mail_content_signature = prop.getProperty("mail.content.signature");
		mail_template = prop.getProperty("template.email");
		template_path = prop.getProperty("freemarker.template.path");
		setFile_cfg();
	}
	private void setFile_cfg() {
		try {
			Properties p = new Properties();
			p.load(SendMailUtil.class.getClassLoader().getResourceAsStream("freemarker.properties"));
			
			file_cfg.setDirectoryForTemplateLoading(new File(template_path));
			file_cfg.setSettings(p);
			
//			ClassTemplateLoader ctl = new ClassTemplateLoader(TemplateFileFreemarkerUtil.class, "/template");
//			TemplateLoader tl = file_cfg.getTemplateLoader();
//			TemplateLoader[] loaders = new TemplateLoader[] { tl, ctl };
//			MultiTemplateLoader mtl = new MultiTemplateLoader(loaders);
//			file_cfg.setTemplateLoader(mtl);
		} catch (IOException e) {
			logger.error(e);
		} catch (TemplateException e) {
			logger.error(e);
		}
		file_cfg.setObjectWrapper(new DefaultObjectWrapper());
	}
	public Properties getProp() {
		return prop;
	}

	public void setProp(Properties prop) {
		this.prop = prop;
	}

	public String getMail_smtp_host() {
		return mail_smtp_host;
	}
	public void setMail_smtp_host(String mail_smtp_host) {
		this.mail_smtp_host = mail_smtp_host;
	}
	public String getMail_auth() {
		return mail_auth;
	}
	public void setMail_auth(String mail_auth) {
		this.mail_auth = mail_auth;
	}
	public String getMail_user() {
		return mail_user;
	}
	public void setMail_user(String mail_user) {
		this.mail_user = mail_user;
	}
	public String getMail_password() {
		return mail_password;
	}
	public void setMail_password(String mail_password) {
		this.mail_password = mail_password;
	}
	public String getMail_from() {
		return mail_from;
	}
	public void setMail_from(String mail_from) {
		this.mail_from = mail_from;
	}
	public String getMail_to() {
		return mail_to;
	}
	public void setMail_to(String mail_to) {
		this.mail_to = mail_to;
	}
	public String getMail_cc() {
		return mail_cc;
	}
	public void setMail_cc(String mail_cc) {
		this.mail_cc = mail_cc;
	}
	public String getMail_subject() {
		return mail_subject;
	}
	public void setMail_subject(String mail_subject) {
		this.mail_subject = mail_subject;
	}
	public String getMail_content_hi() {
		return mail_content_hi;
	}
	public void setMail_content_hi(String mail_content_hi) {
		this.mail_content_hi = mail_content_hi;
	}
	public String getMail_content() {
		return mail_content;
	}
	public void setMail_content(String mail_content) {
		this.mail_content = mail_content;
	}
	public String getMail_content_signature() {
		return mail_content_signature;
	}
	public void setMail_content_signature(String mail_content_signature) {
		this.mail_content_signature = mail_content_signature;
	}
	public String getMail_template() {
		return mail_template;
	}
	public void setMail_template(String mail_template) {
		this.mail_template = mail_template;
	}
	public String getTemplate_path() {
		return template_path;
	}
	public void setTemplate_path(String template_path) {
		this.template_path = template_path;
	}
	public void sendBatchReport(String subject, String attachments) {
		HashMap<String, String> rootMap = new HashMap<String,String>();
		rootMap.put("hi", mail_content_hi);
		rootMap.put("content", mail_content);
		rootMap.put("signature", mail_content_signature);
		try {
			String content = getContent(mail_template, rootMap);
			logger.info("get mail content ---->"+content.length());
			MailUtilCom mail = MailUtilCom.newInstance(mail_smtp_host, mail_user, mail_password, mail_auth);
			mail.send(mail_from, mail_to,mail_cc,null, subject, content, attachments);
			logger.info("send mail Success");
		} catch (IOException e) {
			logger.error(e.getMessage(),e);
		} catch (TemplateException e) {
			logger.error(e.getMessage(),e);
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
		}
	}
	public void sendBatchReport(String attachments) {
		sendBatchReport(mail_subject,attachments);
	}
	public void sendMail() {
		HashMap<String, String> rootMap = new HashMap<String,String>();
		rootMap.put("hi", mail_content_hi);
		rootMap.put("content", mail_content);
		rootMap.put("signature", mail_content_signature);
		try {
			String content = getContent(mail_template, rootMap);
			logger.info("get mail content ---->"+content.length());
			MailUtilCom mail = MailUtilCom.newInstance(mail_smtp_host, mail_user, mail_password, mail_auth);
			mail.send(mail_from, mail_to,mail_cc,null, mail_subject, content);
			logger.info("send mail Success");
		} catch (IOException e) {
			logger.error(e.getMessage(),e);
		} catch (TemplateException e) {
			logger.error(e.getMessage(),e);
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
		}
	}
	public String getContent(String templateName,Object data) throws IOException, TemplateException{
		StringWriter w = new StringWriter();
		Template temp = file_cfg.getTemplate(templateName);
		temp.process(data, w);
		w.flush();
		return w.toString();
	}
}
