package com.ask.inv.google.dirve;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Properties;
import java.util.Set;

import com.ask.inv.web.download.DownLoadBook;
import com.ask.inv.web.download.DriveCommandLine;
import com.sun.org.apache.xml.internal.security.exceptions.Base64DecodingException;
import com.sun.org.apache.xml.internal.security.utils.Base64;

public class DownloadUpload {
	public static void main(String[] args) throws Exception {
		if (args.length < 2) {
			System.out.println("Please input phantomJs path and config path");
			System.exit(1);
		}
		
		Properties prop = new Properties();
		try {
			InputStream in = new BufferedInputStream(new FileInputStream(
					args[1]));
			prop.load(in);

		} catch (Exception e) {
			e.printStackTrace();
		}

		/**
		 * tableau_account=jetan tableau_password=123!@#QWEqweE
		 * tableau_url=http://tableau001.vcbrands.net/auth/
		 * google_account=jerome.tan@ask.com google_password=123!@#QWEqwe
		 * google_doc_filename=./Master.twbx
		 * download_url=http://tableau001.vcbrands
		 * .net/workbooks/MasterJerometest?format=twb&errfmt=html
		 */
		byte[] key = Base64.decode("/imPE5FUDowZE4ltFWHsSVekH7nav0zl");
		
		String tableau_account = prop.getProperty("tableau_account");
		String tableau_password = new String(DESedeCoder.decrypt(Base64.decode(prop.getProperty("tableau_password")), key));
		String tableau_url = prop.getProperty("tableau_url");
		String google_account = prop.getProperty("google_account");
		String google_password = new String(DESedeCoder.decrypt(Base64.decode(prop.getProperty("google_password")), key));
		String google_doc_filename = prop.getProperty("google_doc_filename");
		String download_url = prop.getProperty("download_url");
//		System.out.println(b);
		// Set keyValue = prop.keySet();
		// for (Iterator it = keyValue.iterator(); it.hasNext();) {
		// String key = (String) it.next();
		// }

		String fileName = "./Master_tmp.twbx";
		File f = new File(fileName);
		if (f.exists()) {
			f.delete();
		}
		DownLoadBook dlb = new DownLoadBook(tableau_url, tableau_account,
				tableau_password, download_url, fileName);
		
		dlb.DownLoadfile();
		if (f.exists()) {
			DriveCommandLine driveCommandLine = new DriveCommandLine(fileName,"Master.twbx");
			driveCommandLine.uploadFile();
		}

	}
}
