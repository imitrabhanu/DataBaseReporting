package com.novopay.masteronboarding.report.prod; 
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.log4j.Logger;
public class OnBoardingEmail {
	
	private static Logger logger=Logger.getLogger("OnBoardingEmail.class");
	public void sendEmail(String Date)
	{
		 Properties properties=new Properties();
		   FileInputStream fileInput=null;
		   
	try{	 
		fileInput =new FileInputStream("./Properties/EmailProperties.properties");
		properties.load(fileInput);
		fileInput.close();	
	}
	catch (IOException IOe )
	{
		logger.debug("Io Excption",IOe);
	}
	catch (Exception ie )
	{
		logger.debug("Excption",ie);
	}
		       String to =properties.getProperty("ToEmail");
			   String from =properties.getProperty("FromEmail");
			   final String username =properties.getProperty("Username");
			   final String password =properties.getProperty("Password");
			   String cc=properties.getProperty("CcEmail");
			  
			   
	      // Assuming you are sending email through relay.jangosmtp.net
	      Properties props = new Properties();
			props.put("mail.smtp.host", "smtp.gmail.com");
			props.put("mail.smtp.socketFactory.port", "465");
			props.put("mail.smtp.socketFactory.class",
					"javax.net.ssl.SSLSocketFactory");
			props.put("mail.smtp.auth", "true");
			props.put("mail.smtp.port", "465");

	      Session session = Session.getInstance(props,
	         new javax.mail.Authenticator() {
	            protected PasswordAuthentication getPasswordAuthentication() 
	            {
	               return new PasswordAuthentication(username, password);
	            }
	         });

	      try {
	    	  // Create a default MimeMessage object.
	          Message message = new MimeMessage(session);

	          // Set From: header field of the header.
	          message.setFrom(new InternetAddress(from));

	          // Set To: header field of the header.
	          message.setRecipients(Message.RecipientType.TO,
	             InternetAddress.parse(to));
	          
	          //Set CC
	          
	          message.setRecipients(Message.RecipientType.CC,
	                  InternetAddress.parse(cc));

	          // Set Subject: header field
	          message.setSubject("Daily On-Boarding Master Tracker as on "+ Date);

	          // This mail has 2 part, the BODY and the embedded image
	          MimeMultipart multipart = new MimeMultipart("related");

	    	    // first part  (the html)
	    	    BodyPart messageBodyPart = new MimeBodyPart();
	    	    String htmlText = "<p> Dear All, <br>Please find the Daily On-Boarding Master data below.</p>"
	    	    +"<h3>Zone/Area Wise Overall Status:-</h3>";
	    	    htmlText+="<p align=left><img src=\"cid:image1\"> </p>";
	    	    htmlText+="<h3>Partner Wise Status:-</h3>";
	    	    htmlText+="<p align=left><img src=\"cid:image2\"> </p>";
	    	    htmlText+="<h3>Pendency in Sales Bin:-</h3>";
	    	    htmlText+="<p align=left><img src=\"cid:image3\"> </p>";
	    	    htmlText+="<h3>Axis partner, awaiting for Device Numbers:-</h3>";
	    	    htmlText+="<p align=left><img src=\"cid:image4\"> </p>";
	    	    htmlText+="<p>Regards,<br>reports_dev@novopay.in<br></p>";	  
	    	    messageBodyPart.setContent(htmlText, "text/html");

	    	    // add it
	    	    multipart.addBodyPart(messageBodyPart);

	    	    // second part (1st image)
	    	    messageBodyPart = new MimeBodyPart();
	    	    DataSource fds1 = new FileDataSource
	    	    ("./Image/Zone & Area Wise Data.jpeg");
	    	    messageBodyPart.setDataHandler(new DataHandler(fds1));
	    	    messageBodyPart.addHeader("Content-ID","<image1>");
	    	    // add it
	    	    multipart.addBodyPart(messageBodyPart);
	            //2nd image
	    	    messageBodyPart = new MimeBodyPart();
	    	    DataSource fds2 = new FileDataSource
	    	    ("./Image/Partner Wise Data.jpeg");
	    	    messageBodyPart.setDataHandler(new DataHandler(fds2));
	    	    messageBodyPart.addHeader("Content-ID","<image2>");
	    	    // add it
	    	    multipart.addBodyPart(messageBodyPart);
	    	   //3rd image
	    	    messageBodyPart = new MimeBodyPart();
	    	    DataSource fds3 = new FileDataSource
	    	    ("./Image/Pendency in Sales Bin.jpeg");
	    	    messageBodyPart.setDataHandler(new DataHandler(fds3));
	    	    messageBodyPart.addHeader("Content-ID","<image3>");
	    	    // add it
	    	    multipart.addBodyPart(messageBodyPart);
	            //24th image
	    	    messageBodyPart = new MimeBodyPart();
	    	    DataSource fds4 = new FileDataSource
	    	    ("./Image/Axis Partner, Awaiting for Devi.jpeg");
	    	    messageBodyPart.setDataHandler(new DataHandler(fds4));
	    	    messageBodyPart.addHeader("Content-ID","<image4>");
	    	    // add it
	    	    multipart.addBodyPart(messageBodyPart); 
	    	    
	    	 /*// Add Attachment
	            messageBodyPart = new MimeBodyPart();
	            String filename = "./Report/Report_"+reportGenerationDate+".xlsx";
	            DataSource source = new FileDataSource(filename);
	            messageBodyPart.setDataHandler(new DataHandler(source));
	            messageBodyPart.setFileName(filename);
	            multipart.addBodyPart(messageBodyPart);*/
	    	 	    
	          // put everything together
	    	    message.setContent(multipart);

	         // Send message
	         Transport.send(message);

	      } 
	      catch (MessagingException me) 
	      {
	      logger.error("Messaging Exception",me);
	      }
	      
	      catch (Exception e) 
	      {
	      logger.error("Exception",e);
	      }
	      logger.debug("Email Sent Successfully");
	   }		
}