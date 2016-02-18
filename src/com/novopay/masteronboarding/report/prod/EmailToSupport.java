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
/**
* The class sends an email to specified users with an attachment.
* 
* @author  Mitrabhanu
* @version 1.0
* @since   2016-02-17 
*/
public class EmailToSupport {
	private static Logger logger=Logger.getLogger("OnBoardingEmail.class");
	
	/**
     * This is a method to send an email  to specified users along with the image in the body of the email.
     * 
     * @return Nothing
     */	
	public void sendEmailToSupport(String Date,String reportGenerationDate){

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
catch (Exception e )
{
	logger.debug("Excption",e);
}
	       String ToSupportEmail =properties.getProperty("ToSupportEmail");
		   String from =properties.getProperty("FromEmail");
		   final String username =properties.getProperty("Username");
		   final String password =properties.getProperty("Password");
		   String ccSupport=properties.getProperty("CcSupportEmail");
	   
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
            protected PasswordAuthentication getPasswordAuthentication() {
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
             InternetAddress.parse(ToSupportEmail));
          
          //Set CC
          message.setRecipients(Message.RecipientType.CC,
                  InternetAddress.parse(ccSupport));

          // Set Subject: header field
          message.setSubject("Daily On-Boarding Master Tracker Sheet "+Date);

          // This mail has 2 part, the BODY and the embedded image
          MimeMultipart multipart = new MimeMultipart("related");

    	    // first part  (the html)
    	    BodyPart messageBodyPart = new MimeBodyPart();
    	    String htmlText = "<p> Hi Team, <br>Please find the Daily On-Boarding Master Tracker Sheet attached. Kindly upload the same to Google Drive.</p>";
    	    	    htmlText+="<p>Regards,<br>reports_dev@novopay.in<br></p>";	  
    	    	    messageBodyPart.setContent(htmlText, "text/html");
    	    // add it
    	    multipart.addBodyPart(messageBodyPart); 
    	    messageBodyPart = new MimeBodyPart();
            String filename = "./Report/Report_"+reportGenerationDate+".xlsx";
            DataSource source = new FileDataSource(filename);
            messageBodyPart.setDataHandler(new DataHandler(source));
            messageBodyPart.setFileName(filename);
            multipart.addBodyPart(messageBodyPart);
    	   
    	    // put everything together
    	    message.setContent(multipart);

         // Send message
         Transport.send(message);

      } catch (MessagingException me) 
      {
      logger.error("Messaging Exception",me);
      }
      
      catch (Exception e) 
      {
      logger.error("Exception",e);
      }
      logger.debug("Email Sent to Support Team Successfully");
   }
}

