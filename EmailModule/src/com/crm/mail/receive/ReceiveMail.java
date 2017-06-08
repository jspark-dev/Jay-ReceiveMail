package com.crm.mail.receive;

import com.sun.mail.util.FolderClosedIOException;
import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.Hashtable;
import java.util.List;
import java.util.Properties;
import javax.mail.BodyPart;
import javax.mail.Folder;
import javax.mail.FolderClosedException;
import javax.mail.Message;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.ContentType;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMultipart;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

public class ReceiveMail {
	private int attachmentCount = 0;
	private boolean getAttachment = false;
	private String attachmentPath = null;
	private String downloadPath = null;
	private String emailSentDateFolder = null;
	private List<String> filteredExtensionList = null;
	private StringBuilder filteredAttachment = null;

	private void setFilteredExtensionList() {
		this.filteredExtensionList = new ArrayList();
		this.filteredExtensionList.add("ACTION");
		this.filteredExtensionList.add("APK");
		this.filteredExtensionList.add("APP");
		this.filteredExtensionList.add("BAT");
		this.filteredExtensionList.add("BIN");
		this.filteredExtensionList.add("CMD");
		this.filteredExtensionList.add("COM");
		this.filteredExtensionList.add("COMMAND");
		this.filteredExtensionList.add("CPL");
		this.filteredExtensionList.add("CSH");
		this.filteredExtensionList.add("EXE");
		this.filteredExtensionList.add("GADGET");
		this.filteredExtensionList.add("INF");
		this.filteredExtensionList.add("INS");
		this.filteredExtensionList.add("INX");
		this.filteredExtensionList.add("IPA");
		this.filteredExtensionList.add("ISU");
		this.filteredExtensionList.add("JOB");
		this.filteredExtensionList.add("JSE");
		this.filteredExtensionList.add("KSH");
		this.filteredExtensionList.add("LNK");
		this.filteredExtensionList.add("MSC");
		this.filteredExtensionList.add("MSI");
		this.filteredExtensionList.add("MSP");
		this.filteredExtensionList.add("MST");
		this.filteredExtensionList.add("OSX");
		this.filteredExtensionList.add("OUT");
		this.filteredExtensionList.add("PAF");
		this.filteredExtensionList.add("PIF");
		this.filteredExtensionList.add("PRG");
		this.filteredExtensionList.add("PS1");
		this.filteredExtensionList.add("REG");
		this.filteredExtensionList.add("RGS");
		this.filteredExtensionList.add("RUN");
		this.filteredExtensionList.add("SCR");
		this.filteredExtensionList.add("SCT");
		this.filteredExtensionList.add("SHB");
		this.filteredExtensionList.add("SHS");
		this.filteredExtensionList.add("U3P");
		this.filteredExtensionList.add("VB");
		this.filteredExtensionList.add("VBE");
		this.filteredExtensionList.add("VBS");
		this.filteredExtensionList.add("VBSCRIPT");
		this.filteredExtensionList.add("WORKFLOW");
		this.filteredExtensionList.add("WS");
		this.filteredExtensionList.add("WSF");
		this.filteredExtensionList.add("WSH");
	}

	public String getEmailCount(String protocol, String emailHost, String emailID, String emailPassword) throws Exception {
		String strCount = null;
		Properties properties = new Properties();
		properties.setProperty("mail.store.protocol", protocol);

		Session session = Session.getInstance(properties, null);
		Store store = session.getStore();
		store.connect(emailHost, emailID, emailPassword);
		Folder inbox = store.getFolder("INBOX");
		inbox.open(1);

		strCount = Integer.toString(inbox.getMessageCount());

		return strCount;
	}

	public String getEmail(String protocol, String emailHost, String emailID, String emailPassword, String downloadPath, String emailIndex, boolean getAttachment) throws Exception {
		this.getAttachment = getAttachment;
		this.downloadPath = downloadPath;

		if (this.getAttachment) {
			setFilteredExtensionList();
		}
		Properties properties = new Properties();
		properties.setProperty("mail.store.protocol", protocol);
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");

		Session session = Session.getInstance(properties, null);
		Store store = session.getStore();
		store.connect(emailHost, emailID, emailPassword);
		Folder inbox = store.getFolder("INBOX");
		inbox.open(1);

		stringBuilder.append("<Emails>");
		this.attachmentCount = 0;
		this.attachmentPath = "";
		Message message = inbox.getMessage(Integer.parseInt(emailIndex));
		if (message != null) {
			stringBuilder.append("<Email>");
			stringBuilder.append("<From>").append(((InternetAddress) message.getFrom()[0]).getAddress()).append("</From>");
			stringBuilder.append("<ReceivedDate>").append(message.getReceivedDate().toString()).append("</ReceivedDate>");
			stringBuilder.append("<ReceivedDateInMilliseconds>").append(message.getReceivedDate().getTime()).append("</ReceivedDateInMilliseconds>");
			stringBuilder.append("<Subject>").append(escapeXML(message.getSubject())).append("</Subject>");

			this.emailSentDateFolder = Long.toString(message.getReceivedDate().getTime());
			String content = getContentFromMessage(message);
			if ((content != null) && (!(content.isEmpty())))
				stringBuilder.append("<Content>").append(escapeXML(content)).append("</Content>");
			else {
				stringBuilder.append("<Content></Content>");
			}

			if (this.getAttachment) {
				stringBuilder.append(this.attachmentPath);

				if (this.filteredAttachment != null) {
					stringBuilder.append(this.filteredAttachment.toString());
					this.filteredAttachment = null;
				}
			}

			stringBuilder.append("<AttachmentCount>").append(Integer.toString(this.attachmentCount)).append("</AttachmentCount>");
			stringBuilder.append("</Email>");
			stringBuilder.append("</Emails>");
		}

		return stringBuilder.toString();
	}

	public Hashtable<String, String> getEmails(String protocol, String emailHost, String emailID, String emailPassword, String downloadPath, String inputTimestamp, boolean getAttachment) throws Exception {
		Hashtable listOfEmails = new Hashtable();
		this.getAttachment = getAttachment;
		this.downloadPath = downloadPath;

		if (this.getAttachment) {
			setFilteredExtensionList();
		}
		Properties properties = new Properties();
		properties.setProperty("mail.store.protocol", protocol);
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");

		Session session = Session.getInstance(properties, null);
		Store store = session.getStore();
		store.connect(emailHost, emailID, emailPassword);
		Folder inbox = store.getFolder("INBOX");
		inbox.open(1);

		if ((inbox.getMessageCount() > 0) && (inbox.getMessages() != null)) {
			stringBuilder.append("<Emails>");
			
			int numberOfIteration = 500;
			if(inbox.getMessageCount() < numberOfIteration) {
				numberOfIteration = inbox.getMessageCount();
			}
			for (int j = inbox.getMessageCount(); j > 0; --j) {
				StringBuilder stringBuilderTemp = new StringBuilder();
				try {
					this.attachmentCount = 0;
					this.attachmentPath = "";

					Message message = inbox.getMessage(j);
					if (message != null) {
						if ((inputTimestamp != null) && (!(inputTimestamp.isEmpty()))) {
							long tempTimestamp = Long.parseLong(inputTimestamp);
							if (message.getReceivedDate().getTime() <= tempTimestamp) {
								break;
							}
						}

						if ((j == inbox.getMessageCount()) && (message.getReceivedDate() != null)) {
							listOfEmails.put("LAST_EMAIL_TIMESTAMP", Long.toString(message.getReceivedDate().getTime()));
						}

						stringBuilderTemp.append("<Email>");
						stringBuilderTemp.append("<From>").append(((InternetAddress) message.getFrom()[0]).getAddress()).append("</From>");
						stringBuilderTemp.append("<ReceivedDate>").append((message.getReceivedDate() != null) ? message.getReceivedDate().toString() : "").append("</ReceivedDate>");
						stringBuilderTemp.append("<Subject>").append(escapeXML(message.getSubject())).append("</Subject>");

						this.emailSentDateFolder = Long.toString(message.getReceivedDate().getTime());
						String content = getContentFromMessage(message);
						if ((content != null) && (!(content.isEmpty())))
							stringBuilderTemp.append("<Content>").append(escapeXML(content)).append("</Content>");
						else {
							stringBuilderTemp.append("<Content></Content>");
						}

						if (this.getAttachment) {
							stringBuilderTemp.append(this.attachmentPath);

							if (this.filteredAttachment != null) {
								stringBuilderTemp.append(this.filteredAttachment.toString());
								this.filteredAttachment = null;
							}
						}

						stringBuilderTemp.append("<AttachmentCount>").append(Integer.toString(this.attachmentCount)).append("</AttachmentCount>");
						stringBuilderTemp.append("</Email>");
						stringBuilder.append(stringBuilderTemp.toString());
					}
				} catch (FolderClosedException e) {
					inbox = store.getFolder("INBOX");
					inbox.open(1);
					--j;
				} catch (FolderClosedIOException e) {
					inbox = store.getFolder("INBOX");
					inbox.open(1);
					--j;
				}
			}
			stringBuilder.append("</Emails>");
			listOfEmails.put("EMAIL_LIST", stringBuilder.toString());
		}

		return listOfEmails;
	}

	public String getLatestEmail(String protocol, String emailHost, String emailID, String emailPassword, String downloadPath, boolean getAttachment) throws Exception {
		this.getAttachment = getAttachment;
		this.downloadPath = downloadPath;

		if (this.getAttachment) {
			setFilteredExtensionList();
		}
		Properties properties = new Properties();
		properties.setProperty("mail.store.protocol", protocol);
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");

		Session session = Session.getInstance(properties, null);
		Store store = session.getStore();
		store.connect(emailHost, emailID, emailPassword);
		Folder inbox = store.getFolder("INBOX");
		inbox.open(1);

		stringBuilder.append("<Emails>");
		this.attachmentCount = 0;
		this.attachmentPath = "";
		Message message = inbox.getMessage(inbox.getMessageCount());
		if (message != null) {
			stringBuilder.append("<Email>");
			stringBuilder.append("<From>").append(((InternetAddress) message.getFrom()[0]).getAddress()).append("</From>");
			stringBuilder.append("<ReceivedDate>").append(message.getReceivedDate().toString()).append("</ReceivedDate>");
			stringBuilder.append("<ReceivedDateInMilliseconds>").append(message.getReceivedDate().getTime()).append("</ReceivedDateInMilliseconds>");
			stringBuilder.append("<Subject>").append(escapeXML(message.getSubject())).append("</Subject>");

			this.emailSentDateFolder = Long.toString(message.getSentDate().getTime());
			String content = getContentFromMessage(message);
			if ((content != null) && (!(content.isEmpty())))
				stringBuilder.append("<Content>").append(escapeXML(content)).append("</Content>");
			else {
				stringBuilder.append("<Content></Content>");
			}

			if (this.getAttachment) {
				stringBuilder.append(this.attachmentPath);

				if (this.filteredAttachment != null) {
					stringBuilder.append(this.filteredAttachment.toString());
					this.filteredAttachment = null;
				}
			}

			stringBuilder.append("<AttachmentCount>").append(Integer.toString(this.attachmentCount)).append("</AttachmentCount>");
			stringBuilder.append("</Email>");
			stringBuilder.append("</Emails>");
		}

		return stringBuilder.toString();
	}

	private String getContentFromMessage(Message message) throws Exception {
		String content = null;

		if (message.isMimeType("text/plain")) {
			content = message.getContent().toString();
		} else if (message.isMimeType("text/html")) {
			String contentHTML = message.getContent().toString();
			content = Jsoup.parse(contentHTML).text();
			content = content.replace("&", "&amp;");
		} else if (message.isMimeType("multipart/*")) {
			MimeMultipart multipart = (MimeMultipart) message.getContent();
			content = getContentFromMimeMultipart(multipart);
		}

		return content;
	}

	private String getContentFromMimeMultipart(MimeMultipart multipart) throws Exception {
		String content = "";

		int count = multipart.getCount();
		if (count == 0) {
			return content;
		}
		boolean multipartAlt = new ContentType(multipart.getContentType()).match("multipart/alternative");
		if (multipartAlt) {
			return getContentFromMimeBodyPart(multipart.getBodyPart(count - 1));
		}
		for (int i = 0; i < count; ++i) {
			BodyPart bodypart = multipart.getBodyPart(i);
			content = content + getContentFromMimeBodyPart(bodypart);
		}

		return content;
	}

	private String getContentFromMimeBodyPart(BodyPart bodypart) throws Exception {
		String content = "";

		if ((bodypart.isMimeType("text/plain")) && (!("attachment".equalsIgnoreCase(bodypart.getDisposition())))) {
			content = bodypart.getContent().toString();
		} else if ((bodypart.isMimeType("text/html")) && (!("attachment".equalsIgnoreCase(bodypart.getDisposition())))) {
			String contentHTML = bodypart.getContent().toString();
			content = Jsoup.parse(contentHTML).text();
		} else if (bodypart.getContent() instanceof MimeMultipart) {
			if (("attachment".equalsIgnoreCase(bodypart.getDisposition())) || ((bodypart.getFileName() != null) && (!(bodypart.getFileName().isEmpty())))) {
				if (this.getAttachment) {
					getAttachment((MimeBodyPart) bodypart);
				}
				this.attachmentCount += 1;
			}
			content = getContentFromMimeMultipart((MimeMultipart) bodypart.getContent());
		} else if (("attachment".equalsIgnoreCase(bodypart.getDisposition())) || ((bodypart.getFileName() != null) && (!(bodypart.getFileName().isEmpty())))) {
			if (this.getAttachment) {
				getAttachment((MimeBodyPart) bodypart);
			}
			this.attachmentCount += 1;
		}

		return content;
	}

	private void getAttachment(MimeBodyPart bodyPart) throws Exception {
		String filename = bodyPart.getFileName();
		String newFileName = filename;

		String fileExtension = filename.substring(filename.lastIndexOf(".") + 1);
		boolean isIgnore = false;

		for (String executable : this.filteredExtensionList) {
			if (fileExtension.equalsIgnoreCase(executable)) {
				isIgnore = true;
				if (this.filteredAttachment == null) {
					this.filteredAttachment = new StringBuilder();
				}

				this.filteredAttachment.append("<Warning>").append(filename).append(" is an executable.</Warning>");
				break;
			}
		}

		if (!(isIgnore)) {
			File filePath = new File(this.downloadPath + this.emailSentDateFolder);
			if (!(filePath.exists())) {
				filePath.mkdirs();
			}

			File file = new File(this.downloadPath + File.separator + this.emailSentDateFolder + File.separator + newFileName);

			if (!(file.exists())) {
				bodyPart.saveFile(file);
			}
			this.attachmentPath = this.attachmentPath + "<Attachment>" + this.downloadPath + File.separator + this.emailSentDateFolder + File.separator + newFileName + "</Attachment>";
		}
	}
	
	private String escapeXML(String message) {
		String result = null;
		
		result = message.replaceAll("&", "&amp;");
		result = result.replaceAll("<", "&lt;");
		result = result.replaceAll(">", "&gt;");
		
		return result;
	}
}