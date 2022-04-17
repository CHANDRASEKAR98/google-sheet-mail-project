package com.sheet.mail;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.UnknownHostException;
import java.security.GeneralSecurityException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.atomic.AtomicInteger;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
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

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

/*
 * GoogleSheetMail class is used to send mails to the employees
 *  
 */
public class GoogleSheetMail {
	
	private static final String APPLICATION_NAME = "Google Sheets Mail Project";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";
    private static final List<String> SCOPES = Collections.singletonList(
    		SheetsScopes.SPREADSHEETS_READONLY);
    private static final String CREDENTIALS_FILE_PATH = "/client_secret.json";
    
    /*
     * getCredentials method is used to authorize the google API credentials
     * @param HTTP_TRANSPORT
     * @return Credential
     * 
     */
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = GoogleSheetMail.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, 
        		new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }
    
    public static void main(String... args) throws IOException, GeneralSecurityException {
    	try {
	        // Build a new authorized API client service.
    		AtomicInteger count = new AtomicInteger(0);
	        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
	        final String spreadsheetId = "<paste_your_spreadsheet_url>";
	        final String range = "Sheet1!A2:D";
	        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
	                .setApplicationName(APPLICATION_NAME)
	                .build();
	        ValueRange response = service.spreadsheets().values()
	                .get(spreadsheetId, range)
	                .execute();
	        List<List<Object>> values = response.getValues();
	        if (values == null || values.isEmpty()) {
	            System.out.println("No data found.");
	        } else {
	        	values.stream().forEach(row -> {
	        		try {
	            		String today = new SimpleDateFormat("yyyy/MM/dd").format(new Date());
						Date birthDate = new SimpleDateFormat("yyyy/MM/dd").parse(String.valueOf(row.get(2)));
						String birthDay = new SimpleDateFormat("yyyy/MM/dd").format(birthDate);
						if(birthDay.equals(today)) {
							System.out.println(row.get(1) + " has birthday today!!");
							try {
								sendBirthDayMailNotification(String.valueOf(row.get(1)), 
										String.valueOf(row.get(2)), String.valueOf(row.get(3)));
								count.incrementAndGet();
								System.out.println("Mail sent successfully!!!");
							} catch (AddressException e) {
								System.out.println("Invalid Mail Address : "+ e);
							} catch (MessagingException e) {
								System.out.println("Unable to send the mail : "+ e);
							}
						}
	        		} catch (ParseException e) {
	        			System.out.println("Date Parsing Error: " + e);
	        		}
	        	});
	        	if(count.get() == 0) {
	        		System.out.println("No one has birthday today!!!");
	        	}
	        }
    	} catch(UnknownHostException e) {
    		System.out.println("Couldn't connect to the Network. Check your network and try again. " + e);
    	}
    }

    /*
     * sendBirthDayMailNotification method is used to send mail
     * @param employeeName
     * @param birthDay
     * @param emailId
     * @return 
     * 
     */
	private static void sendBirthDayMailNotification(String employeeName, String birthDay, 
			String emailId) throws AddressException, MessagingException {
		Properties prop = new Properties();
		prop.put("mail.smtp.host", "smtp.gmail.com");
        prop.put("mail.smtp.port", "465");
        prop.put("mail.smtp.auth", "true");
        prop.put("mail.smtp.socketFactory.port", "465");
        prop.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
		Session session = Session.getInstance(prop, new Authenticator() {
		    @Override
		    protected PasswordAuthentication getPasswordAuthentication() {
		    	String username = "<paste_your_mail_id>";
		    	String password = "<paste_your_mail_password>";
		        return new PasswordAuthentication(username, password);
		    }
		});
		
		Message message = new MimeMessage(session);
		message.setFrom(new InternetAddress("<sender_mail_address>"));
		message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(emailId));
		message.setSubject(setMailSubject());
		
		MimeBodyPart mimeBodyPart = new MimeBodyPart();
		mimeBodyPart.setContent(setMailBody(employeeName), "text/html; charset=utf-8");
		
		MimeBodyPart mimeattachmentPart = new MimeBodyPart();
		DataSource dataSource = new FileDataSource(new File("F:/hearticon.jpg"));
		mimeattachmentPart.setDataHandler(new DataHandler(dataSource));
		mimeattachmentPart.setFileName(dataSource.getName());
		
		Multipart multiPart = new MimeMultipart();
		multiPart.addBodyPart(mimeBodyPart);
		multiPart.addBodyPart(mimeattachmentPart);
		message.setContent(multiPart);
		
		Transport.send(message);
	}
	
	/*
	 * setMailSubject method is used to set a subject for the mail
	 * return String
	 * 
	 */
	private static String setMailSubject() {
		return "Wishing you Happy BirthDay!!!";
	}
	
	/*
	 * setMailBody method is used to set body content of the mail
	 * @param employeeName
	 * return String
	 * 
	 */
	private static String setMailBody(String employeeName) {
		String greet = "Dear <b>" + employeeName + "</b>, <br><br>";
		String wishes = "We are so much delighted to celebrate this wonderful day with you.<br>"
				+ "May this year brings you all the happiness you want!!<br>"
				+ "We would like to thank you for your great contribution to our organization's growth"
				+ " and your continuous support even through the hard times.<br>"
				+ "On the whole, we are so excited to be working with you!!<br><br>"
				+ "Have wonderful day celebrating with your family and friends!";
		return greet.concat(wishes);
		
	}
}
