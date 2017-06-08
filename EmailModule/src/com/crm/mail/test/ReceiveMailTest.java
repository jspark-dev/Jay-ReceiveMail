package com.crm.mail.test;

import java.util.Hashtable;

import com.crm.mail.receive.ReceiveMail;

public class ReceiveMailTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String protocol = "imaps";
		String emailHost = "imap.mail.us-east-1.awsapps.com";
		String emailID = "jknight@flexiblebydesign.com";
		String emailPassword = "Jlkj0646!23";
		String downloadPath = "C:\\Users\\Administrator\\Documents\\FlexClientFolders\\MorningStarFiber\\";
		String emailIndex = "1";
		boolean getAttachment = true;
		String inputTimestamp = "";
		
		
		ReceiveMail receiveMail = new ReceiveMail();
		
		try {
			String result = null;
			result = receiveMail.getEmailCount(protocol, emailHost, emailID, emailPassword);
			System.out.println(result);

			result = receiveMail.getEmail(protocol, emailHost, emailID, emailPassword, downloadPath, emailIndex, getAttachment);
			System.out.println(result);
			
			result = receiveMail.getLatestEmail(protocol, emailHost, emailID, emailPassword, downloadPath, getAttachment);
			System.out.println(result);
			
			Hashtable<String, String> resultHT = receiveMail.getEmails(protocol, emailHost, emailID, emailPassword, downloadPath, inputTimestamp, getAttachment);
			System.out.println(resultHT.size());
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
