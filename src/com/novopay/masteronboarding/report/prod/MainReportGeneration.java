package com.novopay.masteronboarding.report.prod;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.log4j.Logger;


public class MainReportGeneration {
	private static Logger logger=Logger.getLogger(MainReportGeneration.class);
	//Method to get Todays Date
			private String todayDate()
			{
				String today=new SimpleDateFormat("ddMMMyyyy").format(new Date());
				return today;
			}
			//Method to get Date for File Name if the generation date is today
				private String reportDateToday()
				{
					String reportDateToday=new SimpleDateFormat("yyyyMMMdd_HH").format(new Date());
					return reportDateToday;
				}
	
	public static void main(String[] args) {
		logger.debug("Report generation started");
		MainReportGeneration MainObject=new MainReportGeneration();
		ReportGeneratorImplementation reportGeneratorImplementation=new ReportGeneratorImplementation();
		reportGeneratorImplementation.workBook();
		
		reportGeneratorImplementation.ReportGenerator(MainObject.reportDateToday());
		ExcelToImageConvertion conversion=new ExcelToImageConvertion();
		conversion.ImageConversion(MainObject.reportDateToday());
		ImageCropping cropping=new ImageCropping();
		cropping.conversionImplementation();
		OnBoardingEmail email=new OnBoardingEmail();
		email.sendEmail(MainObject.todayDate());
		EmailToSupport emailSupport=new EmailToSupport();
		emailSupport.sendEmailToSupport(MainObject.todayDate(),MainObject.reportDateToday());
		logger.debug("Report generation successfully");
	}

}
