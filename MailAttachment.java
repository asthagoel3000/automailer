
package com.ericsson.softhuman.rpa;

import javax.mail.*;
import javax.mail.internet.*;
import java.util.*;
import java.io.IOException;
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MailAttachment
{
	static String INPATH ="";
	static String OUTPATH ="";
	
	public static int doProcess1(String InPath, String OutPath) 
    {		
		
		INPATH =InPath;
		OUTPATH =OutPath;
		String from, path, user_name, to, cc, subject, body, signature, attachment;
		int i, j, row_count;		
		XSSFWorkbook myExcelBook;
		XSSFSheet myExcelSheet;
		
		try 
		{	
			// Assuming you are sending email from localhost
		    String host = "mail2.internal.ericsson.com";

			// Get system properties
			Properties properties = System.getProperties();

			// Setup mail server
			properties.setProperty("mail.smtp.host", host);

			// Get the default Session object.
			Session session = Session.getDefaultInstance(properties);
			
			//getting system user name
			user_name=System.getProperty("user.name");	
			
			//getting From email address			
			from= "C:\\Users\\" + user_name + "\\AppData\\Local\\Microsoft\\Outlook";
			File folder = new File(from);
			File[] listOfFiles = folder.listFiles();
			for (i = 0; i < listOfFiles.length; i++)
			{
				if (listOfFiles[i].isFile())
				{
					if(listOfFiles[i].getName().endsWith("nst"))
					{
						from=listOfFiles[i].getName();
						from=from.substring(0,(from.length()-4));
					}
				}
			}
			
			//Sending Input file path to readFromExcel function
			path= /*"C:\\Users\\" + user_name*/ INPATH + "Email_Bot_inputs.xlsx";
			
			myExcelBook = new XSSFWorkbook(new FileInputStream(path));
		    myExcelSheet = myExcelBook.getSheet("Sheet1");
		    row_count = myExcelSheet.getLastRowNum();
		    System.out.println("Total emails to be sent: " + row_count);    
			
		    for (j=1; j<=row_count;j++)
		    {
				//input To email address
				to= readFromExcel(path, "to", j);	
				
				//input Cc email address
				cc= readFromExcel(path, "cc", j);	
				
				//input Subject
				subject= readFromExcel(path, "subject", j);
				
				//input Body
				body= readFromExcel(path, "body", j);
				
				//input Signature
				signature= readFromExcel(path, "signature", j);
				
				//input attachment
		    	attachment= readFromExcel(path, "attachment", j);
				attachment= attachment + "\\";			
				
				// Create a default MimeMessage object.
				MimeMessage message = new MimeMessage(session);
	
				// Set From: header field of the header.
				
				String s1= from.substring(0,from.indexOf("@")+1);
				String s2=s1+"ericsson.com";
				
				message.setFrom(new InternetAddress(s2));
	
				message.addRecipients(Message.RecipientType.TO, InternetAddress.parse(to));
				message.addRecipients(Message.RecipientType.CC, InternetAddress.parse(cc));
				
				message.setSubject(subject);
	
				// Create the message part 
				BodyPart messageBodyPart = new MimeBodyPart();
	
				// Create a multipart message
				Multipart multipart = new MimeMultipart();
				messageBodyPart = new MimeBodyPart();
	
				//mail body with HTML formatting
				body = "<body><div>Hi,<br><br>" + body + "<br><br></div>";
			    signature = "<body>Thanks & Regards<br><B>" + signature + "</B>";			
				body= body + signature;				
	
				// Send the complete message parts
				messageBodyPart.setContent(body, "text/html");
				multipart.addBodyPart(messageBodyPart);
				
				//attaching all files present in the selected path
				System.out.println("Email-"+ j + " attachment list:-");
				File folder2 = new File(attachment);
				File[] listOfFiles2 = folder2.listFiles();
				for (i = 0; i < listOfFiles2.length; i++)
				{		
					if (listOfFiles2[i].isFile())
					{	
						MimeBodyPart att = new MimeBodyPart();
						System.out.println(listOfFiles2[i].getName());
						att.attachFile(attachment+listOfFiles2[i].getName());
						multipart.addBodyPart(att);					
					}
				}	
				message.setContent(multipart);
	
				// Send message
				Transport.send(message);
				System.out.println("Email-"+ j + " sent successfully.");
		    }
		    myExcelBook.close();
		}				
		catch(Exception e)
		{       
			e.printStackTrace();
			return -1;
		}	
		return 0;
	}
	
	public static String readFromExcel(String file, String category, int row_count) throws IOException
	{			
	    XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
	    XSSFSheet myExcelSheet = myExcelBook.getSheet("Sheet1");
	    XSSFRow row = myExcelSheet.getRow(row_count);
	    try
	    {
	    	if( category == "to")
		    {
		    	return row.getCell(0).getStringCellValue();
		    }
		    else if( category == "cc")
		    {
		    	return row.getCell(1).getStringCellValue();
		    }
		    else if( category == "subject")
		    {
		    	return row.getCell(2).getStringCellValue();
		    }
		    else if( category == "body")
		    {
		    	return row.getCell(3).getStringCellValue();
		    }
		    else if( category == "signature")
		    {
		    	return row.getCell(4).getStringCellValue();
		    }
		    else if( category == "attachment")
		    {
		    	return row.getCell(5).getStringCellValue();
		    }   
		    else
		    {
		    	return "INVALID INPUT";
		    }
	    }
	    finally
	    {
	    	myExcelBook.close();
	    }
	}
}
