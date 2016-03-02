package com.novopay.masteronboarding.report.prod;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.log4j.Logger;
/**
* The program implements an application that
* automates the process of report generation from database 
* and sends the report to all the required personals.
* 
* 
* @author  Mitrabhanu
* @version 1.0
* @since   2016-02-17 
*/
public class MainReportGeneration {
private static Logger logger=Logger.getLogger(MainReportGeneration.class);
	
	       /**
	        * This is a method to get the current date in ddMMMyyyy format.
	        * 
	        * @param today This is the parameter to store the date.
	        * @return String This returns the current date in ddMMMyyyy format
	        */	
			private String todayDate()
			{
				String today=new SimpleDateFormat("ddMMMyyyy").format(new Date());
				return today;
			}
			/**
		     * This is a method to get the current date in yyyyMMMdd_HH format.
	         * 
             * @param reportDateToday This is the parameter to store the date.
   	         * @return String This returns the current date in yyyyMMMdd_HH format
		     */	
			 private String reportDateToday()
			 {
				 String reportDateToday=new SimpleDateFormat("yyyyMMMdd_HH").format(new Date());
				 return reportDateToday;
			 }	
    /**
     * This is the main method.
     * 
     * @param args Unused.
     * @return Nothing.
     */												
	public static void main(String[] args) {
		logger.debug("Report generation started");
		
		MainReportGeneration MainObject=new MainReportGeneration();
		ReportGeneratorImplementation reportGeneratorImplementation=new ReportGeneratorImplementation();
		reportGeneratorImplementation.workBook();
		
		ArrayList<Integer> sheetRowCountList =reportGeneratorImplementation.ReportGenerator(MainObject.reportDateToday());
		ExcelToImageConvertion conversion=new ExcelToImageConvertion();
		conversion.ImageConversion(MainObject.reportDateToday());
		
		ImageCropping cropping=new ImageCropping();
		cropping.conversionImplementation(sheetRowCountList);
		OnBoardingEmail email=new OnBoardingEmail();
		
		email.sendEmail(MainObject.todayDate());
		EmailToSupport emailSupport=new EmailToSupport();
		emailSupport.sendEmailToSupport(MainObject.todayDate(),MainObject.reportDateToday());
		
		logger.debug("Report generation successfully");
	}

}
