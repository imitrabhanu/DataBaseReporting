package com.novopay.masteronboarding.report.prod;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
* The class implements various method to generate excel report from database 
* and stores the report in a specified location.
* It uses Apache POI Api for storing the data in excel sheet.
* 
* 
* @author  Mitrabhanu
* @version 1.0
* @since   2016-02-17 
*/


public class ReportGeneratorImplementation {
	private static Logger logger=Logger.getLogger("ReportGeneratorImplementation.class");
	     //WorkBook Creation
		 static XSSFWorkbook wb=new XSSFWorkbook();
		 //Sheet Creation
		 XSSFSheet sheet1 =(XSSFSheet) wb.createSheet("Master OnBoarding Data");
		 XSSFSheet sheet2 =(XSSFSheet) wb.createSheet("Zone & Area Wise Data");
		 XSSFSheet sheet3 =(XSSFSheet) wb.createSheet("Partner Wise Data");
		 XSSFSheet sheet4 =(XSSFSheet) wb.createSheet("Pendency in Sales Bin");
		 XSSFSheet sheet5 =(XSSFSheet) wb.createSheet("Axis Partner, Awaiting for Device Numbers");
		 /**
	        * This is a method to create header for each sheet, assigning cell style and cell values.
	        * 
	        * @return Nothing
	        */
		 public void workBook()
		 {   
			 XSSFRow header1 =(XSSFRow) sheet1.createRow(0);
			 XSSFRow header2 =(XSSFRow) sheet2.createRow(0);
			 XSSFRow header3 =(XSSFRow) sheet3.createRow(0);
			 XSSFRow header4 =(XSSFRow) sheet4.createRow(0);
			 XSSFRow header5 =(XSSFRow) sheet5.createRow(0);		 	 
			 //Sheet 1
			 Cell cell1=header1.createCell(0);
			 cell1.setCellValue("Partner");
			 cell1.setCellStyle(csHeader());
			 
			 Cell cell2=header1.createCell(1);
			 cell2.setCellValue("Retailer Name");
			 cell2.setCellStyle(csHeader());
			 
			 Cell cell3=header1.createCell(2);
			 cell3.setCellValue("Area");
			 cell3.setCellStyle(csHeader());
			 
			 Cell cell4=header1.createCell(3);
			 cell4.setCellValue("Territory");
			 cell4.setCellStyle(csHeader());
			 
			 Cell cell5=header1.createCell(4);
			 cell5.setCellValue("Current Status");
			 cell5.setCellStyle(csHeader());
			 
			 Cell cell6=header1.createCell(5);
			 cell6.setCellValue("Last Action Date");
			 cell6.setCellStyle(csHeader());
			 
			 Cell cell7=header1.createCell(6);
			 cell7.setCellValue("Currently pending with");
			 cell7.setCellStyle(csHeader());
			 
			 Cell cell8=header1.createCell(7);
			 cell8.setCellValue("Created By");
			 cell8.setCellStyle(csHeader());
			 
			 Cell cell9=header1.createCell(8);
			 cell9.setCellValue("Created by Role");
			 cell9.setCellStyle(csHeader());
			 
			 Cell cell10=header1.createCell(9);
			 cell10.setCellValue("Lead Creation Date");
			 cell10.setCellStyle(csHeader());
			 
			 Cell cell11=header1.createCell(10);
			 cell11.setCellValue("Lead Approval Date");
			 cell11.setCellStyle(csHeader());
			 
			 Cell cell12=header1.createCell(11);
			 cell12.setCellValue("Lead approved by");
			 cell12.setCellStyle(csHeader());
			 
			 Cell cell13=header1.createCell(12);
			 cell13.setCellValue("Date of Onboarding Request");
			 cell13.setCellStyle(csHeader());
			 
			 Cell cell14=header1.createCell(13);
			 cell14.setCellValue("Docs Current Status");
			 cell14.setCellStyle(csHeader());
			 
			 Cell cell15=header1.createCell(14);
			 cell15.setCellValue("Doc Return count");
			 cell15.setCellStyle(csHeader());
			 
			 Cell cell16=header1.createCell(15);
			 cell16.setCellValue("Last upload by sales team");
			 cell16.setCellStyle(csHeader());
			 
			 Cell cell17=header1.createCell(16);
			 cell17.setCellValue("Docs Verification Date (by NP Ops)");
			 cell17.setCellStyle(csHeader());
			 
			 Cell cell18=header1.createCell(17);
			 cell18.setCellValue("Latest Rejection/Sent Back Remarks");
			 cell18.setCellStyle(csHeader());
			 
			 Cell cell19=header1.createCell(18);
			 cell19.setCellValue("Doc Upload Date in RBL Portal");
			 cell19.setCellStyle(csHeader());
			 
			 Cell cell20=header1.createCell(19);
			 cell20.setCellValue("Bank Status");
			 cell20.setCellStyle(csHeader());
			 
			 Cell cell21=header1.createCell(20);
			 cell21.setCellValue("Bank Remarks");
			 cell21.setCellStyle(csHeader());
			 
			 Cell cell22=header1.createCell(21);
			 cell22.setCellValue("NP Portal Status");
			 cell22.setCellStyle(csHeader());
			 
			 Cell cell23=header1.createCell(22);
			 cell23.setCellValue("Onboarding fees");
			 cell23.setCellStyle(csHeader());
			 
			 Cell cell24=header1.createCell(23);
			 cell24.setCellValue("Amount receipt confirmation date");
			 cell24.setCellStyle(csHeader());
			 
			 Cell cell25=header1.createCell(24);
			 cell25.setCellValue("Activation date");
			 cell25.setCellStyle(csHeader());
			 
			 Cell cell26=header1.createCell(25);
			 cell26.setCellValue("VAN");
			 cell26.setCellStyle(csHeader());
			 
			 Cell cell27=header1.createCell(26);
			 cell27.setCellValue("First Deposit Amount");
			 cell27.setCellStyle(csHeader());
			 
			 Cell cell28=header1.createCell(27);
			 cell28.setCellValue("First Cash In Date");
			 cell28.setCellStyle(csHeader());
			 
			 Cell cell29=header1.createCell(28);
			 cell29.setCellValue("msisdn");
			 cell29.setCellStyle(csHeader());
			 
			 Cell cell30=header1.createCell(29);
			 cell30.setCellValue("zone");
			 cell30.setCellStyle(csHeader());
						 
			 //Sheet 2
			 Cell cell31=header2.createCell(0);
			 cell31.setCellValue("Zone");
			 cell31.setCellStyle(csHeader());
			 
			 Cell cell32=header2.createCell(1);
			 cell32.setCellValue("Area");
			 cell32.setCellStyle(csHeader());
			 
			 Cell cell33=header2.createCell(2);
			 cell33.setCellValue("Data Collection Pending");
			 cell33.setCellStyle(csHeader());
			 
			 Cell cell34=header2.createCell(3);
			 cell34.setCellValue("Pending for document verification");
			 cell34.setCellStyle(csHeader());
			 
			 Cell cell35=header2.createCell(4);
			 cell35.setCellValue("Pending with Sales Team - Sent Back");
			 cell35.setCellStyle(csHeader());
			 
			 Cell cell36=header2.createCell(5);
			 cell36.setCellValue("Documents verified");
			 cell36.setCellStyle(csHeader());
			 
			 Cell cell37=header2.createCell(6);
			 cell37.setCellValue("Waiting for OD account number");
			 cell37.setCellStyle(csHeader());
			 
			 Cell cell38=header2.createCell(7);
			 cell38.setCellValue("Activated");
			 cell38.setCellStyle(csHeader());
			 
			 Cell cell39=header2.createCell(8);
			 cell39.setCellValue("Total");
			 cell39.setCellStyle(csHeader());
			
			 //Sheet 3
			 Cell cell40=header3.createCell(0);
			 cell40.setCellValue("Partner");
			 cell40.setCellStyle(csHeader());
			 
			 Cell cell41=header3.createCell(1);
			 cell41.setCellValue("Data Collection Pending");
			 cell41.setCellStyle(csHeader());
			 
			 Cell cell42=header3.createCell(2);
			 cell42.setCellValue("Pending for document verification");
			 cell42.setCellStyle(csHeader());
			 
			 Cell cell43=header3.createCell(3);
			 cell43.setCellValue("Pending with Sales Team - Sent Back");
			 cell43.setCellStyle(csHeader());
			 
			 Cell cell44=header3.createCell(4);
			 cell44.setCellValue("Documents verified");
			 cell44.setCellStyle(csHeader());
			 
			 Cell cell45=header3.createCell(5);
			 cell45.setCellValue("Waiting for OD account number");
			 cell45.setCellStyle(csHeader());
			 
			 Cell cell46=header3.createCell(6);
			 cell46.setCellValue("Activated");
			 cell46.setCellStyle(csHeader());
			 
			 Cell cell47=header3.createCell(7);
			 cell47.setCellValue("Total");
			 cell47.setCellStyle(csHeader());
			 
			 //Sheet 4
			 Cell cell48=header4.createCell(0);
			 cell48.setCellValue("Zone");
			 cell48.setCellStyle(csHeader());
			 
			 Cell cell49=header4.createCell(1);
			 cell49.setCellValue("0-1 Days");
			 cell49.setCellStyle(csHeader());
			 
			 Cell cell50=header4.createCell(2);
			 cell50.setCellValue("1-2 Days");
			 cell50.setCellStyle(csHeader());
			 
			 Cell cell51=header4.createCell(3);
			 cell51.setCellValue("2-3 Days");
			 cell51.setCellStyle(csHeaderLongPending());
			 
			 Cell cell52=header4.createCell(4);
			 cell52.setCellValue("3-4 Days");
			 cell52.setCellStyle(csHeaderLongPending());
			 
			 Cell cell53=header4.createCell(5);
			 cell53.setCellValue("4-5 Days");
			 cell53.setCellStyle(csHeaderLongPending());
			 
			 Cell cell54=header4.createCell(6);
			 cell54.setCellValue("5-6 Days");
			 cell54.setCellStyle(csHeaderLongPending());
			 
			 Cell cell55=header4.createCell(7);
			 cell55.setCellValue("6-7 Days");
			 cell55.setCellStyle(csHeaderLongPending());
			 
			 Cell cell56=header4.createCell(8);
			 cell56.setCellValue("7-8 Days");
			 cell56.setCellStyle(csHeaderLongPending());
			 
			 Cell cell57=header4.createCell(9);
			 cell57.setCellValue("8-9 Days");
			 cell57.setCellStyle(csHeaderLongPending());
			 
			 Cell cell58=header4.createCell(10);
			 cell58.setCellValue("9-10 Days");
			 cell58.setCellStyle(csHeaderLongPending());
			 
			 Cell cell59=header4.createCell(11);
			 cell59.setCellValue("10-15 Days");
			 cell59.setCellStyle(csHeaderLongPending());
			 
			 Cell cell60=header4.createCell(12);
			 cell60.setCellValue("15-20 Days");
			 cell60.setCellStyle(csHeaderLongPending());
			 
			 Cell cell61=header4.createCell(13);
			 cell61.setCellValue("20-30 Days");
			 cell61.setCellStyle(csHeaderLongPending());
			 
			 Cell cell62=header4.createCell(14);
			 cell62.setCellValue("30-60 Days");
			 cell62.setCellStyle(csHeaderLongPending());
			 
			 Cell cell63=header4.createCell(15);
			 cell63.setCellValue("60-100 Days");
			 cell63.setCellStyle(csHeaderLongPending());
			 
			 Cell cell64=header4.createCell(16);
			 cell64.setCellValue(">100 Day");
			 cell64.setCellStyle(csHeaderLongPending());
			 
			 Cell cell65=header4.createCell(17);
			 cell65.setCellValue("Total");
			 cell65.setCellStyle(csHeader());
			 
			 //Sheet 5
			 Cell cell66=header5.createCell(0);
			 cell66.setCellValue("Zone");
			 cell66.setCellStyle(csHeader());
			 
			 Cell cell67=header5.createCell(1);
			 cell67.setCellValue("0-1 Days");
			 cell67.setCellStyle(csHeader());
			 
			 Cell cell68=header5.createCell(2);
			 cell68.setCellValue("1-2 Days");
			 cell68.setCellStyle(csHeader());
			 
			 Cell cell69=header5.createCell(3);
			 cell69.setCellValue("2-3 Days");
			 cell69.setCellStyle(csHeaderLongPending());
			 
			 Cell cell70=header5.createCell(4);
			 cell70.setCellValue("3-4 Days");
			 cell70.setCellStyle(csHeaderLongPending());
			 
			 Cell cell71=header5.createCell(5);
			 cell71.setCellValue("4-5 Days");
			 cell71.setCellStyle(csHeaderLongPending());
			 
			 Cell cell72=header5.createCell(6);
			 cell72.setCellValue("5-6 Days");
			 cell72.setCellStyle(csHeaderLongPending());
			 
			 Cell cell73=header5.createCell(7);
			 cell73.setCellValue("6-7 Days");
			 cell73.setCellStyle(csHeaderLongPending());
			 
			 Cell cell74=header5.createCell(8);
			 cell74.setCellValue("7-8 Days");
			 cell74.setCellStyle(csHeaderLongPending());
			 
			 Cell cell75=header5.createCell(9);
			 cell75.setCellValue("8-9 Days");
			 cell75.setCellStyle(csHeaderLongPending());
			 
			 Cell cell76=header5.createCell(10);
			 cell76.setCellValue("9-10 Days");
			 cell76.setCellStyle(csHeaderLongPending());
			 
			 Cell cell77=header5.createCell(11);
			 cell77.setCellValue("10-15 Days");
			 cell77.setCellStyle(csHeaderLongPending());
			 
			 Cell cell78=header5.createCell(12);
			 cell78.setCellValue("15-20 Days");
			 cell78.setCellStyle(csHeaderLongPending());
			 
			 Cell cell79=header5.createCell(13);
			 cell79.setCellValue("20-30 Days");
			 cell79.setCellStyle(csHeaderLongPending());
			 
			 Cell cell80=header5.createCell(14);
			 cell80.setCellValue("30-60 Days");
			 cell80.setCellStyle(csHeaderLongPending());
			 
			 Cell cell81=header5.createCell(15);
			 cell81.setCellValue("60-100 Days");
			 cell81.setCellStyle(csHeaderLongPending());
			 
			 Cell cell82=header5.createCell(16);
			 cell82.setCellValue(">100 Day");
			 cell82.setCellStyle(csHeaderLongPending());
			 
			 Cell cell83=header5.createCell(17);
			 cell83.setCellValue("Total");
			 cell83.setCellStyle(csHeader());
			 
			 
		 }
		 			
			
		//Method to Create CellStyle for Header
		 
		 /**
	        * This is a method to to create CellStyle for Header.
	        * 
	        * @return CellStyle this returns the cell style for header.
	        */
		    private CellStyle csHeader()
		    {
		    	
				 CellStyle cs =wb.createCellStyle();
				 cs.setFillForegroundColor(IndexedColors.GOLD.getIndex());
				 cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
				
				 Font font=wb.createFont();
				 font.setColor(IndexedColors.BLACK.getIndex());
				 cs.setFont(font);
				 
				 cs.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
				 cs.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
				 cs.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
				 cs.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
				 return cs;
                 }
		        /**
		        * This is a method to to create CellStyle for Header with long pending data.
		        * 
		        * @return CellStyle this returns the cell style for header with long pending data.
		        */
		    private CellStyle csHeaderLongPending()
		    {
		    	 CellStyle cs1 =wb.createCellStyle();
		    	 cs1.setFillForegroundColor(IndexedColors.RED.getIndex());
				 cs1.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    	 
				 Font font1=wb.createFont();
				 font1.setColor(IndexedColors.WHITE.getIndex());
				 cs1.setFont(font1);
				 			 
				 cs1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				 cs1.setBorderRight(XSSFCellStyle.BORDER_THIN);
				 cs1.setBorderTop(XSSFCellStyle.BORDER_THIN);
				 cs1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				 return cs1;		   	
		    }
		    //Method to Add Borders to Cells
		    /**
		        * This is a method to to add Borders to Cells.
		        * 
		        * @return CellStyle this returns the cell style to add Borders to Cells.
		        */
		    
		    private CellStyle csBorder()
		    {
		    	 CellStyle cs2 =wb.createCellStyle();
				 cs2.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				 cs2.setBorderRight(XSSFCellStyle.BORDER_THIN);
				 cs2.setBorderTop(XSSFCellStyle.BORDER_THIN);
				 cs2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				 return cs2;		   	
		    }
		    
		  
		        /**
		        * This is a method to store the data in excel in a specified location.
		        * 
		        * @return Nothing
		        */
		    private void FileWrite(String reportGenerationDate)
		    {
		    	try{
	      FileOutputStream fileOut = new FileOutputStream("./Report/Report_"+reportGenerationDate+".xlsx");
	      wb.write(fileOut);
	      fileOut.close();
		    	}
		    	catch (IOException ie) 
		    	{      
		    		logger.error("IO Exception",ie);     
			    }
		    	
		    	catch (Exception e) 
		    	{      
		    		logger.error("Exception",e);     
			    }
		    }
		    
		      /**
		        * This is a method to read data from database and write it to excel.
		        * 
		        * @return Nothing
		        */
		
		    
		    ArrayList<Integer> ReportGenerator(String reportDate )
			{
				    PreparedStatement stmt1,stmt2,stmt3,stmt4,stmt5=null; 
					ResultSet rs1,rs2,rs3,rs4,rs5=null;
					ArrayList<Integer> sheetRowCountList=new ArrayList<Integer>();
					
					String s1="SELECT partner AS 'Partner',r.name AS 'Retailer Name', areas.name AS 'Area',territory.name AS 'Territory', CASE WHEN s.description ='Approved' AND roh.source_status_id=5 THEN 'Pending with Sales Team' WHEN s.description ='Approved' AND roh.source_status_id=13 THEN 'Pending with Sales Team - Sent Back' WHEN s.description ='Data Verification Failed' THEN 'Pending with Sales Team - Sent Back' WHEN s.description ='Pending Onboard' THEN 'Pending for document verification' WHEN s.description ='Onboarded' AND r.partner='RBL' THEN 'Documents verified' WHEN s.description ='Onboarded' AND r.partner='BOI' THEN 'Waiting for OD account number' ELSE s.description END AS 'Current Status', roh.action_performed_date AS 'Last Action Date', CASE WHEN s.description ='Approved' AND roh.source_status_id=5 THEN 'Sales' WHEN s.description ='Approved' AND roh.source_status_id=13 THEN 'Sales' WHEN s.description ='Pending Onboard' THEN 'Ops' WHEN s.description ='Onboarded' AND r.partner='RBL' THEN 'Ops' WHEN s.description ='Onboarded' AND r.partner='BOI' THEN 'Partner' WHEN s.description ='Data Verification Failed' THEN 'Sales' ELSE 'N/A' END AS 'Currently pending with', IFNULL(TRIM(u.first_name),'') AS 'Created By',rm.role_title AS 'Created by Role', lc.action_performed_date AS 'Lead Creation Date',IFNULL(la.action_performed_date,'') AS 'Lead Approval Date', IFNULL(la.first_name,'') AS 'Lead approved by', IFNULL(appr.action_performed_date,'') AS 'Date of Onboarding Request', CASE WHEN onb.action_performed_date IS NOT NULL THEN 'Good' WHEN onb.action_performed_date IS NULL AND return_count IS NULL AND appr.action_performed_date IS NOT NULL THEN 'To Be Verified' WHEN onb.action_performed_date IS  NULL AND return_count>0 THEN 'Returned' ELSE '' END AS 'Docs Current Status', IFNULL(return_count,'') AS 'Doc Return count',IFNULL(upl.action_performed_date,'') AS 'Last upload by sales team', IFNULL(onb.action_performed_date,'') AS 'Docs Verification Date (by NP Ops)', CASE WHEN (roh.source_status_id='13' AND roh.target_status_id='7') OR (roh.source_status_id='5' AND roh.target_status_id='9') OR (roh.source_status_id='13' AND roh.target_status_id='9') OR (roh.source_status_id='13' AND roh.target_status_id='24') THEN roh.comments ELSE '' END AS 'Latest Rejection/Sent Back Remarks',  '' AS 'Doc Upload Date in RBL Portal','' AS 'Bank Status','' AS 'Bank Remarks', CASE WHEN act.action_performed_date IS NOT NULL THEN 'Activated' WHEN act.action_performed_date IS NULL AND onb.action_performed_date IS NOT NULL THEN 'Pending activation' ELSE '' END AS 'NP Portal Status', '' AS 'Onboarding fees','' AS 'Amount receipt confirmation date', IFNULL(act.action_performed_date,'') AS 'Activation date', IFNULL(a1.attr_value,'') AS 'VAN',IFNULL(fc.amount,'') AS 'First Deposit Amount',IFNULL(fc.first_cash_in_date,'') AS 'First Cash In Date', msisdn,region.name AS 'zone' FROM np_sales.retailer r JOIN (	SELECT r.retailer_id,r.source_status_id,r.target_status_id,CONCAT('\"',r.comments,'\"') AS comments,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN 	(SELECT retailer_id,MAX(action_performed_date) AS action_performed_date FROM np_sales.retailer_onboarding_history GROUP BY retailer_id) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	WHERE r.target_status_id IN (9,11,13,15,17,24) OR (r.target_status_id = 7 AND r.source_status_id=13) ) roh ON r.id=roh.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r WHERE (source_status_id=3 AND target_status_id=5)  OR (source_status_id IS NULL AND target_status_id=5) ) lc ON r.id=lc.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date,u.first_name 	FROM np_sales.retailer_onboarding_history r 	JOIN master.user_attribute a ON a.attr_key='UUID' AND action_performed_by=a.attr_value 	JOIN master.user u ON a.user_id=u.id 	WHERE source_status_id=5 AND target_status_id=7 ) la ON r.id=la.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MIN(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=7 AND target_status_id = 13 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) appr ON r.id=appr.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history 		WHERE (source_status_id=7 AND target_status_id = 13) 		OR (source_status_id=24 AND target_status_id = 13) 		GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) upl ON r.id=upl.retailer_id LEFT JOIN (	SELECT roh.retailer_id,CASE 	WHEN rh.return_count>0 THEN 'Returned' 	WHEN ro.retailer_id IS NOT NULL THEN 'Good' 	ELSE '' END AS doc_status, 	return_count 	FROM (SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history) roh 	LEFT JOIN ( 	SELECT DISTINCT(retailer_id),COUNT(*) AS return_count FROM np_sales.retailer_onboarding_history roh 	WHERE (source_status_id=13 AND target_status_id = 7) OR (source_status_id=13 AND target_status_id = 24) GROUP BY retailer_id 	)rh ON roh.retailer_id=rh.retailer_id 	LEFT JOIN ( 	SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history roh 	WHERE source_status_id=13 AND target_status_id = 15 	) ro ON roh.retailer_id=ro.retailer_id ) ond ON r.id=ond.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=13 AND target_status_id = 15 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date	 ) onb ON r.id=onb.retailer_id AND r.status>=15 LEFT JOIN (	SELECT retailer_id,action_performed_date 	FROM np_sales.retailer_onboarding_history WHERE 	(source_status_id=15 AND target_status_id = 17)	 	OR (source_status_id=19 AND target_status_id = 17) 	OR (source_status_id=13 AND target_status_id = 17) ) act ON r.id=act.retailer_id LEFT JOIN master.organization org ON r.org_code=org.code LEFT JOIN master.organization_attribute a1 ON a1.orgnization_id=org.id AND a1.attr_key='VIRTUAL_ACC_NUM' LEFT JOIN master.organization_attribute a2 ON a2.orgnization_id=org.id AND a2.attr_key='WALLET_ACCOUNT_NUMBER' LEFT JOIN 		( 			SELECT msa.account_no AS account_no, 			'No' AS first_cash_in_done, '' AS first_cash_in_date,'' AS amount 			FROM wallet.m_savings_account msa 			WHERE msa.id NOT IN ( 				SELECT msat.savings_account_id 				FROM wallet.m_savings_account_transaction msat 				JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 				JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 				WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 				GROUP BY msat.savings_account_id 			) AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) UNION 			SELECT msa.account_no AS account_no, 			'Yes' AS first_cash_in_done, MIN(msat.transaction_date) AS first_cash_in_date,TRUNCATE(msat.amount,2) AS amount 			FROM wallet.m_savings_account msa 			JOIN wallet.m_savings_account_transaction msat ON msa.id=msat.savings_account_id AND msat.transaction_type_enum=1 			JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 			JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 			WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 			AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) 			GROUP BY msat.savings_account_id 		) fc ON fc.account_no=a2.attr_value LEFT JOIN np_sales.status_master s ON roh.target_status_id=s.id LEFT JOIN master.user_attribute ua ON ua.attr_key='UUID' AND ua.attr_value=r.created_by LEFT JOIN master.user u ON u.id=ua.user_id LEFT JOIN (SELECT DISTINCT USER,role FROM master.mapping_user_role) urmap ON u.id=urmap.user LEFT JOIN master.role_master rm ON urmap.role=rm.id LEFT JOIN master.geo_heirarchy_type ght ON ght.code=r.geo_hierarchy_type LEFT JOIN master.geo_heirarchy territory ON r.geo_hierarchy_type='TERRITORY' AND territory.code=r.geo_hierarchy_code AND ght.id=territory.type_id LEFT JOIN master.geo_heirarchy areas ON (territory.parent_id IS NOT NULL AND territory.parent_id=areas.id) 		OR (r.geo_hierarchy_type='AREA' AND areas.code=r.geo_hierarchy_code AND ght.id=areas.type_id) LEFT JOIN master.geo_heirarchy region ON (areas.parent_id IS NOT NULL AND areas.parent_id=region.id) 		OR (r.geo_hierarchy_type='REGION' AND region.code=r.geo_hierarchy_code AND ght.id=region.type_id) WHERE r.name!='Test' AND MONTH(appr.action_performed_date)=MONTH(NOW()+INTERVAL -1 DAY) AND DATE(appr.action_performed_date)<=DATE(NOW()+INTERVAL -1 DAY) ORDER BY (appr.action_performed_date IS NULL),appr.action_performed_date DESC;";
					String s2="SELECT region.name AS 'zone',areas.name AS 'Area', COUNT(DISTINCT CASE WHEN s.description ='Approved' AND roh.source_status_id=7 THEN r.msisdn ELSE NULL END) AS 'Data Collection Pending', COUNT(DISTINCT CASE WHEN s.description ='Pending Onboard' THEN r.msisdn ELSE NULL END) AS 'Pending for document verification', COUNT(DISTINCT CASE WHEN s.description ='Data Verification Failed' THEN r.msisdn ELSE NULL END) AS 'Pending with Sales Team - Sent Back', COUNT(DISTINCT CASE WHEN s.description ='Onboarded' AND r.partner !='BOI' THEN r.msisdn ELSE NULL END) AS'Onboarded', COUNT(DISTINCT CASE WHEN s.description ='Onboarded' AND r.partner='BOI' THEN r.msisdn ELSE NULL END) AS'Waiting for OD account number', COUNT(DISTINCT CASE WHEN s.description ='Activated' THEN r.msisdn ELSE NULL END) AS 'Activated',COUNT(DISTINCT CASE WHEN s.description !='Rejected' THEN r.msisdn ELSE NULL END) AS total FROM np_sales.retailer r JOIN (	SELECT r.retailer_id,r.source_status_id,r.target_status_id,CONCAT('\"',r.comments,'\"') AS comments,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN 	(SELECT retailer_id,MAX(action_performed_date) AS action_performed_date FROM np_sales.retailer_onboarding_history GROUP BY retailer_id) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	WHERE r.target_status_id IN (9,11,13,15,17,24) OR (r.target_status_id = 7 AND r.source_status_id=13) ) roh ON r.id=roh.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r WHERE (source_status_id=3 AND target_status_id=5)  OR (source_status_id IS NULL AND target_status_id=5) ) lc ON r.id=lc.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date,u.first_name 	FROM np_sales.retailer_onboarding_history r 	JOIN master.user_attribute a ON a.attr_key='UUID' AND action_performed_by=a.attr_value 	JOIN master.user u ON a.user_id=u.id 	WHERE source_status_id=5 AND target_status_id=7 ) la ON r.id=la.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MIN(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=7 AND target_status_id = 13 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) appr ON r.id=appr.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history 		WHERE (source_status_id=7 AND target_status_id = 13) 		OR (source_status_id=24 AND target_status_id = 13) 		GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) upl ON r.id=upl.retailer_id LEFT JOIN (	SELECT roh.retailer_id,CASE 	WHEN rh.return_count>0 THEN 'Returned' 	WHEN ro.retailer_id IS NOT NULL THEN 'Good' 	ELSE '' END AS doc_status, 	return_count 	FROM (SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history) roh 	LEFT JOIN ( 	SELECT DISTINCT(retailer_id),COUNT(*) AS return_count FROM np_sales.retailer_onboarding_history roh 	WHERE (source_status_id=13 AND target_status_id = 7) OR (source_status_id=13 AND target_status_id = 24) GROUP BY retailer_id 	)rh ON roh.retailer_id=rh.retailer_id 	LEFT JOIN ( 	SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history roh 	WHERE source_status_id=13 AND target_status_id = 15 	) ro ON roh.retailer_id=ro.retailer_id ) ond ON r.id=ond.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=13 AND target_status_id = 15 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date	 ) onb ON r.id=onb.retailer_id AND r.status>=15 LEFT JOIN (	SELECT retailer_id,action_performed_date 	FROM np_sales.retailer_onboarding_history WHERE 	(source_status_id=15 AND target_status_id = 17)	 	OR (source_status_id=19 AND target_status_id = 17) 	OR (source_status_id=13 AND target_status_id = 17) ) act ON r.id=act.retailer_id LEFT JOIN master.organization org ON r.org_code=org.code LEFT JOIN master.organization_attribute a1 ON a1.orgnization_id=org.id AND a1.attr_key='VIRTUAL_ACC_NUM' LEFT JOIN master.organization_attribute a2 ON a2.orgnization_id=org.id AND a2.attr_key='WALLET_ACCOUNT_NUMBER' LEFT JOIN 		( 			SELECT msa.account_no AS account_no, 			'No' AS first_cash_in_done, '' AS first_cash_in_date,'' AS amount 			FROM wallet.m_savings_account msa 			WHERE msa.id NOT IN ( 				SELECT msat.savings_account_id 				FROM wallet.m_savings_account_transaction msat 				JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 				JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 				WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 				GROUP BY msat.savings_account_id 			) AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) UNION 			SELECT msa.account_no AS account_no, 			'Yes' AS first_cash_in_done, MIN(msat.transaction_date) AS first_cash_in_date,TRUNCATE(msat.amount,2) AS amount 			FROM wallet.m_savings_account msa 			JOIN wallet.m_savings_account_transaction msat ON msa.id=msat.savings_account_id AND msat.transaction_type_enum=1 			JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 			JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 			WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 			AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) 			GROUP BY msat.savings_account_id 		) fc ON fc.account_no=a2.attr_value LEFT JOIN np_sales.status_master s ON roh.target_status_id=s.id LEFT JOIN master.user_attribute ua ON ua.attr_key='UUID' AND ua.attr_value=r.created_by LEFT JOIN master.user u ON u.id=ua.user_id LEFT JOIN (SELECT DISTINCT USER,role FROM master.mapping_user_role) urmap ON u.id=urmap.user LEFT JOIN master.role_master rm ON urmap.role=rm.id LEFT JOIN master.geo_heirarchy_type ght ON ght.code=r.geo_hierarchy_type LEFT JOIN master.geo_heirarchy territory ON r.geo_hierarchy_type='TERRITORY' AND territory.code=r.geo_hierarchy_code AND ght.id=territory.type_id LEFT JOIN master.geo_heirarchy areas ON (territory.parent_id IS NOT NULL AND territory.parent_id=areas.id) 		OR (r.geo_hierarchy_type='AREA' AND areas.code=r.geo_hierarchy_code AND ght.id=areas.type_id) LEFT JOIN master.geo_heirarchy region ON (areas.parent_id IS NOT NULL AND areas.parent_id=region.id) 		OR (r.geo_hierarchy_type='REGION' AND region.code=r.geo_hierarchy_code AND ght.id=region.type_id)	 WHERE r.name!='Test' AND MONTH(appr.action_performed_date)=MONTH(NOW()+INTERVAL -1 DAY) AND DATE(appr.action_performed_date)<=DATE(NOW()+INTERVAL -1 DAY) GROUP BY areas.name ORDER BY region.name;";
					String s3="SELECT partner AS 'Partner', COUNT(DISTINCT CASE WHEN s.description ='Approved' AND roh.source_status_id=7 THEN r.msisdn ELSE NULL END) AS 'Data Collection Pending', COUNT(DISTINCT CASE WHEN s.description ='Pending Onboard' THEN r.msisdn ELSE NULL END) AS 'Pending for document verification', COUNT(DISTINCT CASE WHEN s.description ='Data Verification Failed' THEN r.msisdn ELSE NULL END) AS 'Pending with Sales Team - Sent Back', COUNT(DISTINCT CASE WHEN s.description ='Onboarded' AND r.partner !='BOI' THEN r.msisdn ELSE NULL END) AS'Onboarded', COUNT(DISTINCT CASE WHEN s.description ='Onboarded' AND r.partner='BOI' THEN r.msisdn ELSE NULL END) AS'Waiting for OD account number', COUNT(DISTINCT CASE WHEN s.description ='Activated' THEN r.msisdn ELSE NULL END) AS 'Activated',COUNT(DISTINCT CASE WHEN s.description !='Rejected' THEN r.msisdn ELSE NULL END) AS total FROM np_sales.retailer r JOIN (	SELECT r.retailer_id,r.source_status_id,r.target_status_id,CONCAT('\"',r.comments,'\"') AS comments,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN 	(SELECT retailer_id,MAX(action_performed_date) AS action_performed_date FROM np_sales.retailer_onboarding_history GROUP BY retailer_id) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	WHERE r.target_status_id IN (9,11,13,15,17,24) OR (r.target_status_id = 7 AND r.source_status_id=13) ) roh ON r.id=roh.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r WHERE (source_status_id=3 AND target_status_id=5)  OR (source_status_id IS NULL AND target_status_id=5) ) lc ON r.id=lc.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date,u.first_name 	FROM np_sales.retailer_onboarding_history r 	JOIN master.user_attribute a ON a.attr_key='UUID' AND action_performed_by=a.attr_value 	JOIN master.user u ON a.user_id=u.id 	WHERE source_status_id=5 AND target_status_id=7 ) la ON r.id=la.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MIN(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=7 AND target_status_id = 13 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) appr ON r.id=appr.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history 		WHERE (source_status_id=7 AND target_status_id = 13) 		OR (source_status_id=24 AND target_status_id = 13) 		GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) upl ON r.id=upl.retailer_id LEFT JOIN (	SELECT roh.retailer_id,CASE 	WHEN rh.return_count>0 THEN 'Returned' 	WHEN ro.retailer_id IS NOT NULL THEN 'Good' 	ELSE '' END AS doc_status, 	return_count 	FROM (SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history) roh 	LEFT JOIN ( 	SELECT DISTINCT(retailer_id),COUNT(*) AS return_count FROM np_sales.retailer_onboarding_history roh 	WHERE (source_status_id=13 AND target_status_id = 7) OR (source_status_id=13 AND target_status_id = 24) GROUP BY retailer_id 	)rh ON roh.retailer_id=rh.retailer_id 	LEFT JOIN ( 	SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history roh 	WHERE source_status_id=13 AND target_status_id = 15 	) ro ON roh.retailer_id=ro.retailer_id ) ond ON r.id=ond.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=13 AND target_status_id = 15 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date	 ) onb ON r.id=onb.retailer_id AND r.status>=15 LEFT JOIN (	SELECT retailer_id,action_performed_date 	FROM np_sales.retailer_onboarding_history WHERE 	(source_status_id=15 AND target_status_id = 17)	 	OR (source_status_id=19 AND target_status_id = 17) 	OR (source_status_id=13 AND target_status_id = 17) ) act ON r.id=act.retailer_id LEFT JOIN master.organization org ON r.org_code=org.code LEFT JOIN master.organization_attribute a1 ON a1.orgnization_id=org.id AND a1.attr_key='VIRTUAL_ACC_NUM' LEFT JOIN master.organization_attribute a2 ON a2.orgnization_id=org.id AND a2.attr_key='WALLET_ACCOUNT_NUMBER' LEFT JOIN 		( 			SELECT msa.account_no AS account_no, 			'No' AS first_cash_in_done, '' AS first_cash_in_date,'' AS amount 			FROM wallet.m_savings_account msa 			WHERE msa.id NOT IN ( 				SELECT msat.savings_account_id 				FROM wallet.m_savings_account_transaction msat 				JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 				JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 				WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 				GROUP BY msat.savings_account_id 			) AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) UNION 			SELECT msa.account_no AS account_no, 			'Yes' AS first_cash_in_done, MIN(msat.transaction_date) AS first_cash_in_date,TRUNCATE(msat.amount,2) AS amount 			FROM wallet.m_savings_account msa 			JOIN wallet.m_savings_account_transaction msat ON msa.id=msat.savings_account_id AND msat.transaction_type_enum=1 			JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 			JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 			WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 			AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) 			GROUP BY msat.savings_account_id 		) fc ON fc.account_no=a2.attr_value LEFT JOIN np_sales.status_master s ON roh.target_status_id=s.id LEFT JOIN master.user_attribute ua ON ua.attr_key='UUID' AND ua.attr_value=r.created_by LEFT JOIN master.user u ON u.id=ua.user_id LEFT JOIN (SELECT DISTINCT USER,role FROM master.mapping_user_role) urmap ON u.id=urmap.user LEFT JOIN master.role_master rm ON urmap.role=rm.id LEFT JOIN master.geo_heirarchy_type ght ON ght.code=r.geo_hierarchy_type LEFT JOIN master.geo_heirarchy territory ON r.geo_hierarchy_type='TERRITORY' AND territory.code=r.geo_hierarchy_code AND ght.id=territory.type_id LEFT JOIN master.geo_heirarchy areas ON (territory.parent_id IS NOT NULL AND territory.parent_id=areas.id) 		OR (r.geo_hierarchy_type='AREA' AND areas.code=r.geo_hierarchy_code AND ght.id=areas.type_id) LEFT JOIN master.geo_heirarchy region ON (areas.parent_id IS NOT NULL AND areas.parent_id=region.id) 		OR (r.geo_hierarchy_type='REGION' AND region.code=r.geo_hierarchy_code AND ght.id=region.type_id) WHERE r.name!='Test' AND MONTH(appr.action_performed_date)=MONTH(NOW()+INTERVAL -1 DAY) AND DATE(appr.action_performed_date)<=DATE(NOW()+INTERVAL -1 DAY) GROUP BY partner;";
					String s4="SELECT region.name AS 'zone', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 1 THEN r.msisdn ELSE NULL END) AS 'Aging 0-1 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 2 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 1 THEN r.msisdn ELSE NULL END) AS 'Aging 1-2 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 3 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 2 THEN r.msisdn ELSE NULL END) AS 'Aging 2-3 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 4 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 3 THEN r.msisdn ELSE NULL END) AS 'Aging 3-4 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 5 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 4 THEN r.msisdn ELSE NULL END) AS 'Aging 4-5 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 6 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 5 THEN r.msisdn ELSE NULL END) AS 'Aging 5-6 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 7 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 6 THEN r.msisdn ELSE NULL END) AS 'Aging 6-7 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 8 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 7 THEN r.msisdn ELSE NULL END) AS 'Aging 7-8 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 9 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 8 THEN r.msisdn ELSE NULL END) AS 'Aging 8-9 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 10 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 9 THEN r.msisdn ELSE NULL END) AS 'Aging 9-10 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 15 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 10 THEN r.msisdn ELSE NULL END) AS 'Aging 10-15 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 20 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 15 THEN r.msisdn ELSE NULL END) AS 'Aging 15-20 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 30 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 20 THEN r.msisdn ELSE NULL END) AS 'Aging 20-30 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 60 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 30 THEN r.msisdn ELSE NULL END) AS 'Aging 30-60 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 100 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 60 THEN r.msisdn ELSE NULL END) AS 'Aging 60-100 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) > 100  THEN r.msisdn ELSE NULL END) AS 'Aging >100 Day', COUNT(DISTINCT msisdn) AS Total  FROM np_sales.retailer r JOIN (	SELECT r.retailer_id,r.source_status_id,r.target_status_id,CONCAT('\"',r.comments,'\"') AS comments,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN 	(SELECT retailer_id,MAX(action_performed_date) AS action_performed_date FROM np_sales.retailer_onboarding_history GROUP BY retailer_id) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	WHERE r.target_status_id IN (9,11,13,15,17,24) OR (r.target_status_id = 7 AND r.source_status_id=13) ) roh ON r.id=roh.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r WHERE (source_status_id=3 AND target_status_id=5)  OR (source_status_id IS NULL AND target_status_id=5) ) lc ON r.id=lc.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date,u.first_name 	FROM np_sales.retailer_onboarding_history r 	JOIN master.user_attribute a ON a.attr_key='UUID' AND action_performed_by=a.attr_value 	JOIN master.user u ON a.user_id=u.id 	WHERE source_status_id=5 AND target_status_id=7 ) la ON r.id=la.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MIN(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=7 AND target_status_id = 13 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) appr ON r.id=appr.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history 		WHERE (source_status_id=7 AND target_status_id = 13) 		OR (source_status_id=24 AND target_status_id = 13) 		GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) upl ON r.id=upl.retailer_id LEFT JOIN (	SELECT roh.retailer_id,CASE 	WHEN rh.return_count>0 THEN 'Returned' 	WHEN ro.retailer_id IS NOT NULL THEN 'Good' 	ELSE '' END AS doc_status, 	return_count 	FROM (SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history) roh 	LEFT JOIN ( 	SELECT DISTINCT(retailer_id),COUNT(*) AS return_count FROM np_sales.retailer_onboarding_history roh 	WHERE (source_status_id=13 AND target_status_id = 7) OR (source_status_id=13 AND target_status_id = 24) GROUP BY retailer_id 	)rh ON roh.retailer_id=rh.retailer_id 	LEFT JOIN ( 	SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history roh 	WHERE source_status_id=13 AND target_status_id = 15 	) ro ON roh.retailer_id=ro.retailer_id ) ond ON r.id=ond.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=13 AND target_status_id = 15 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date	 ) onb ON r.id=onb.retailer_id AND r.status>=15 LEFT JOIN (	SELECT retailer_id,action_performed_date 	FROM np_sales.retailer_onboarding_history WHERE 	(source_status_id=15 AND target_status_id = 17)	 	OR (source_status_id=19 AND target_status_id = 17) 	OR (source_status_id=13 AND target_status_id = 17) ) act ON r.id=act.retailer_id LEFT JOIN master.organization org ON r.org_code=org.code LEFT JOIN master.organization_attribute a1 ON a1.orgnization_id=org.id AND a1.attr_key='VIRTUAL_ACC_NUM' LEFT JOIN master.organization_attribute a2 ON a2.orgnization_id=org.id AND a2.attr_key='WALLET_ACCOUNT_NUMBER' LEFT JOIN 		( 			SELECT msa.account_no AS account_no, 			'No' AS first_cash_in_done, '' AS first_cash_in_date,'' AS amount 			FROM wallet.m_savings_account msa 			WHERE msa.id NOT IN ( 				SELECT msat.savings_account_id 				FROM wallet.m_savings_account_transaction msat 				JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 				JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 				WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 				GROUP BY msat.savings_account_id 			) AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) UNION 			SELECT msa.account_no AS account_no, 			'Yes' AS first_cash_in_done, MIN(msat.transaction_date) AS first_cash_in_date,TRUNCATE(msat.amount,2) AS amount 			FROM wallet.m_savings_account msa 			JOIN wallet.m_savings_account_transaction msat ON msa.id=msat.savings_account_id AND msat.transaction_type_enum=1 			JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 			JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 			WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 			AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) 			GROUP BY msat.savings_account_id 		) fc ON fc.account_no=a2.attr_value LEFT JOIN np_sales.status_master s ON roh.target_status_id=s.id LEFT JOIN master.user_attribute ua ON ua.attr_key='UUID' AND ua.attr_value=r.created_by LEFT JOIN master.user u ON u.id=ua.user_id LEFT JOIN (SELECT DISTINCT USER,role FROM master.mapping_user_role) urmap ON u.id=urmap.user LEFT JOIN master.role_master rm ON urmap.role=rm.id LEFT JOIN master.geo_heirarchy_type ght ON ght.code=r.geo_hierarchy_type LEFT JOIN master.geo_heirarchy territory ON r.geo_hierarchy_type='TERRITORY' AND territory.code=r.geo_hierarchy_code AND ght.id=territory.type_id LEFT JOIN master.geo_heirarchy areas ON (territory.parent_id IS NOT NULL AND territory.parent_id=areas.id) 		OR (r.geo_hierarchy_type='AREA' AND areas.code=r.geo_hierarchy_code AND ght.id=areas.type_id) LEFT JOIN master.geo_heirarchy region ON (areas.parent_id IS NOT NULL AND areas.parent_id=region.id) 		OR (r.geo_hierarchy_type='REGION' AND region.code=r.geo_hierarchy_code AND ght.id=region.type_id) WHERE r.name!='Test' AND s.description ='Data Verification Failed' GROUP BY region.name ORDER BY region.name;";
					String s5="SELECT areas.name AS 'Area', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 1 THEN r.msisdn ELSE NULL END) AS 'Aging 0-1 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 2 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 1 THEN r.msisdn ELSE NULL END) AS 'Aging 1-2 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 3 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 2 THEN r.msisdn ELSE NULL END) AS 'Aging 2-3 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 4 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 3 THEN r.msisdn ELSE NULL END) AS 'Aging 3-4 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 5 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 4 THEN r.msisdn ELSE NULL END) AS 'Aging 4-5 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 6 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 5 THEN r.msisdn ELSE NULL END) AS 'Aging 5-6 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 7 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 6 THEN r.msisdn ELSE NULL END) AS 'Aging 6-7 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 8 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 7 THEN r.msisdn ELSE NULL END) AS 'Aging 7-8 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 9 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 8 THEN r.msisdn ELSE NULL END) AS 'Aging 8-9 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 10 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 9 THEN r.msisdn ELSE NULL END) AS 'Aging 9-10 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 15 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 10 THEN r.msisdn ELSE NULL END) AS 'Aging 10-15 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 20 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 15 THEN r.msisdn ELSE NULL END) AS 'Aging 15-20 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 30 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 20 THEN r.msisdn ELSE NULL END) AS 'Aging 20-30 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 60 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 30 THEN r.msisdn ELSE NULL END) AS 'Aging 30-60 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) <= 100 AND DATEDIFF(CURDATE(),appr.action_performed_date) > 60 THEN r.msisdn ELSE NULL END) AS 'Aging 60-100 Days', COUNT(DISTINCT CASE WHEN DATEDIFF(CURDATE(),appr.action_performed_date) > 100  THEN r.msisdn ELSE NULL END) AS 'Aging >100 Day', COUNT(DISTINCT msisdn) AS Total FROM np_sales.retailer r JOIN (	SELECT r.retailer_id,r.source_status_id,r.target_status_id,CONCAT('\"',r.comments,'\"') AS comments,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN 	(SELECT retailer_id,MAX(action_performed_date) AS action_performed_date FROM np_sales.retailer_onboarding_history GROUP BY retailer_id) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	WHERE r.target_status_id IN (9,11,13,15,17,24) OR (r.target_status_id = 7 AND r.source_status_id=13) ) roh ON r.id=roh.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r WHERE (source_status_id=3 AND target_status_id=5)  OR (source_status_id IS NULL AND target_status_id=5) ) lc ON r.id=lc.retailer_id LEFT JOIN ( 	SELECT r.retailer_id,r.action_performed_date,u.first_name 	FROM np_sales.retailer_onboarding_history r 	JOIN master.user_attribute a ON a.attr_key='UUID' AND action_performed_by=a.attr_value 	JOIN master.user u ON a.user_id=u.id 	WHERE source_status_id=5 AND target_status_id=7 ) la ON r.id=la.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MIN(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=7 AND target_status_id = 13 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) appr ON r.id=appr.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history 		WHERE (source_status_id=7 AND target_status_id = 13) 		OR (source_status_id=24 AND target_status_id = 13) 		GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date 	 ) upl ON r.id=upl.retailer_id LEFT JOIN (	SELECT roh.retailer_id,CASE 	WHEN rh.return_count>0 THEN 'Returned' 	WHEN ro.retailer_id IS NOT NULL THEN 'Good' 	ELSE '' END AS doc_status, 	return_count 	FROM (SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history) roh 	LEFT JOIN ( 	SELECT DISTINCT(retailer_id),COUNT(*) AS return_count FROM np_sales.retailer_onboarding_history roh 	WHERE (source_status_id=13 AND target_status_id = 7) OR (source_status_id=13 AND target_status_id = 24) GROUP BY retailer_id 	)rh ON roh.retailer_id=rh.retailer_id 	LEFT JOIN ( 	SELECT DISTINCT retailer_id FROM np_sales.retailer_onboarding_history roh 	WHERE source_status_id=13 AND target_status_id = 15 	) ro ON roh.retailer_id=ro.retailer_id ) ond ON r.id=ond.retailer_id LEFT JOIN (	SELECT r.retailer_id,r.action_performed_date 	FROM np_sales.retailer_onboarding_history r INNER JOIN ( 		SELECT retailer_id,MAX(action_performed_date) AS action_performed_date 		FROM np_sales.retailer_onboarding_history WHERE source_status_id=13 AND target_status_id = 15 GROUP BY retailer_id 	) roh 	ON r.retailer_id=roh.retailer_id AND r.action_performed_date=roh.action_performed_date	 ) onb ON r.id=onb.retailer_id AND r.status>=15 LEFT JOIN (	SELECT retailer_id,action_performed_date 	FROM np_sales.retailer_onboarding_history WHERE 	(source_status_id=15 AND target_status_id = 17)	 	OR (source_status_id=19 AND target_status_id = 17) 	OR (source_status_id=13 AND target_status_id = 17) ) act ON r.id=act.retailer_id LEFT JOIN master.organization org ON r.org_code=org.code LEFT JOIN master.organization_attribute a1 ON a1.orgnization_id=org.id AND a1.attr_key='VIRTUAL_ACC_NUM' LEFT JOIN master.organization_attribute a2 ON a2.orgnization_id=org.id AND a2.attr_key='WALLET_ACCOUNT_NUMBER' LEFT JOIN 		( 			SELECT msa.account_no AS account_no, 			'No' AS first_cash_in_done, '' AS first_cash_in_date,'' AS amount 			FROM wallet.m_savings_account msa 			WHERE msa.id NOT IN ( 				SELECT msat.savings_account_id 				FROM wallet.m_savings_account_transaction msat 				JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 				JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 				WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 				GROUP BY msat.savings_account_id 			) AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) UNION 			SELECT msa.account_no AS account_no, 			'Yes' AS first_cash_in_done, MIN(msat.transaction_date) AS first_cash_in_date,TRUNCATE(msat.amount,2) AS amount 			FROM wallet.m_savings_account msa 			JOIN wallet.m_savings_account_transaction msat ON msa.id=msat.savings_account_id AND msat.transaction_type_enum=1 			JOIN wallet.m_payment_detail mpd ON msat.payment_detail_id = mpd.id 			JOIN wallet.m_code_value mcv ON mpd.payment_type_cv_id = mcv.id 			WHERE mcv.code_value IN ('CASHIN','TRANSFER') AND msat.is_reversed = FALSE 			AND msa.product_id IN ( 				SELECT msp.id 				FROM wallet.m_savings_product msp 				WHERE msp.name = 'BoI Agent' OR msp.name='RBL Agent' 			) 			GROUP BY msat.savings_account_id 		) fc ON fc.account_no=a2.attr_value LEFT JOIN np_sales.status_master s ON roh.target_status_id=s.id LEFT JOIN master.user_attribute ua ON ua.attr_key='UUID' AND ua.attr_value=r.created_by LEFT JOIN master.user u ON u.id=ua.user_id LEFT JOIN (SELECT DISTINCT USER,role FROM master.mapping_user_role) urmap ON u.id=urmap.user LEFT JOIN master.role_master rm ON urmap.role=rm.id LEFT JOIN master.geo_heirarchy_type ght ON ght.code=r.geo_hierarchy_type LEFT JOIN master.geo_heirarchy territory ON r.geo_hierarchy_type='TERRITORY' AND territory.code=r.geo_hierarchy_code AND ght.id=territory.type_id LEFT JOIN master.geo_heirarchy areas ON (territory.parent_id IS NOT NULL AND territory.parent_id=areas.id) 		OR (r.geo_hierarchy_type='AREA' AND areas.code=r.geo_hierarchy_code AND ght.id=areas.type_id) LEFT JOIN master.geo_heirarchy region ON (areas.parent_id IS NOT NULL AND areas.parent_id=region.id) 		OR (r.geo_hierarchy_type='REGION' AND region.code=r.geo_hierarchy_code AND ght.id=region.type_id) WHERE r.name!='Test' AND s.description ='Onboarded' AND r.partner='AXIS' GROUP BY areas.name ORDER BY areas.name;";
					try{		 
						 Connection connection = DatabaseConnection.getConnection();
						 
						//Master OnBoarding Data
						 stmt1=connection.prepareStatement(s1); 
						 //Area/Zone Wise Data
						 stmt2=connection.prepareStatement(s2); 
						//Partner Wise Data
						 stmt3=connection.prepareStatement(s3);
						 //Pendency data
						 stmt4=connection.prepareStatement(s4);
						//Axis data
						 stmt5=connection.prepareStatement(s5);
						 
						 
						 logger.debug("Query Execution Started");
						 //Master Onboarding data
						 rs1=stmt1.executeQuery();
						 int index1=1;
						 while(rs1.next()) 
						 {
						 XSSFRow row = (XSSFRow) sheet1.createRow(index1);
						 Cell cell1a=row.createCell(0);
						 cell1a.setCellValue(rs1.getString(1));
						 cell1a.setCellStyle(csBorder());
						 
						 Cell cell1b=row.createCell(1);
						 cell1b.setCellValue(rs1.getString(2));
						 cell1b.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1c=row.createCell(2);
						 cell1c.setCellValue(rs1.getString(3));
						 cell1c.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1d=row.createCell(3);
						 cell1d.setCellValue(rs1.getString(4));
						 cell1d.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1e=row.createCell(4);
						 cell1e.setCellValue(rs1.getString(5));
						 cell1e.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1f=row.createCell(5);
						 cell1f.setCellValue(rs1.getString(6));
						 cell1f.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1g=row.createCell(6);
						 cell1g.setCellValue(rs1.getString(7));
						 cell1g.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1h=row.createCell(7);
						 cell1h.setCellValue(rs1.getString(8));
						 cell1h.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1i=row.createCell(8);
						 cell1i.setCellValue(rs1.getString(9));
						 cell1i.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1j=row.createCell(9);
						 cell1j.setCellValue(rs1.getString(10));
						 cell1j.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1k=row.createCell(10);
						 cell1k.setCellValue(rs1.getString(11));
						 cell1k.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1l=row.createCell(11);
						 cell1l.setCellValue(rs1.getString(12));
						 cell1l.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1m=row.createCell(12);
						 cell1m.setCellValue(rs1.getString(13));
						 cell1m.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1n=row.createCell(13);
						 cell1n.setCellValue(rs1.getString(14));
						 cell1n.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1o=row.createCell(14);
						 cell1o.setCellValue(rs1.getString(15));
						 cell1o.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1p=row.createCell(15);
						 cell1p.setCellValue(rs1.getString(16));
						 cell1p.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1q=row.createCell(16);
						 cell1q.setCellValue(rs1.getString(17));
						 cell1q.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1r=row.createCell(17);
						 cell1r.setCellValue(rs1.getString(18));
						 cell1r.setCellStyle(cell1a.getCellStyle());
						 
						 
						 Cell cell1s=row.createCell(18);
						 cell1s.setCellValue(rs1.getString(19));
						 cell1s.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1t=row.createCell(19);
						 cell1t.setCellValue(rs1.getString(20));
						 cell1t.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1u=row.createCell(20);
						 cell1u.setCellValue(rs1.getString(21));
						 cell1u.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1v=row.createCell(21);
						 cell1v.setCellValue(rs1.getString(22));
						 cell1v.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1w=row.createCell(22);
						 cell1w.setCellValue(rs1.getString(23));
						 cell1w.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1x=row.createCell(23);
						 cell1x.setCellValue(rs1.getString(24));
						 cell1x.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1y=row.createCell(24);
						 cell1y.setCellValue(rs1.getString(25));
						 cell1y.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1z=row.createCell(25);
						 cell1z.setCellValue(rs1.getString(26));
						 cell1z.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1a1=row.createCell(26);
						 cell1a1.setCellValue(rs1.getString(27));
						 cell1a1.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1a2=row.createCell(27);
						 cell1a2.setCellValue(rs1.getString(28));
						 cell1a2.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1a3=row.createCell(28);
						 cell1a3.setCellValue(rs1.getString(29));
						 cell1a3.setCellStyle(cell1a.getCellStyle());
						 
						 Cell cell1a4=row.createCell(29);
						 cell1a4.setCellValue(rs1.getString(30));
						 cell1a4.setCellStyle(cell1a1.getCellStyle());
						 index1++;					 
						 } 
						 rs1.close();
						 stmt1.close();
						 
						 
						 //Area/Zone Wise Data
						 rs2=stmt2.executeQuery();
						 int index2=1;
						 while(rs2.next()) 
						 {
						 XSSFRow row = (XSSFRow) sheet2.createRow(index2);
						 Cell cell0=row.createCell(0);
						 cell0.setCellValue(rs2.getString(1));
						 cell0.setCellStyle(csBorder());
						 
						 Cell cell1=row.createCell(1);
						 cell1.setCellValue(rs2.getString(2));
						 cell1.setCellStyle(csBorder());
						 
						 Cell cell2=row.createCell(2);
						 cell2.setCellValue(rs2.getString(3));
						 cell2.setCellStyle(csBorder());
						 
						 Cell cell3=row.createCell(3);
						 cell3.setCellValue(rs2.getString(4));
						 cell3.setCellStyle(csBorder());
						 
						 Cell cell4=row.createCell(4);
						 cell4.setCellValue(rs2.getString(5));
						 cell4.setCellStyle(csBorder());
						 
						 Cell cell5=row.createCell(5);
						 cell5.setCellValue(rs2.getString(6));
						 cell5.setCellStyle(csBorder());
						 
						 Cell cell6=row.createCell(6);
						 cell6.setCellValue(rs2.getString(7));
						 cell6.setCellStyle(csBorder());
						 
						 Cell cell7=row.createCell(7);
						 cell7.setCellValue(rs2.getString(8));
						 cell7.setCellStyle(csBorder());
						 
						 Cell cell8=row.createCell(8);
						 cell8.setCellValue(rs2.getString(9));
						 cell8.setCellStyle(csBorder());
						 index2++;					 
						 } 
						 
						 rs2.close();
						 stmt2.close();
						 
						 //Partner Wise Data
						 rs3=stmt3.executeQuery(); 
						 int index3=1;
						 while(rs3.next()) 
						 { 
						 XSSFRow row = (XSSFRow) sheet3.createRow(index3);
						 Cell cell9=row.createCell(0);
						 cell9.setCellValue(rs3.getString(1));
						 cell9.setCellStyle(csBorder());
						 
						 Cell cell10=row.createCell(1);
						 cell10.setCellValue(rs3.getString(2));
						 cell10.setCellStyle(csBorder());
						 
						 Cell cell11=row.createCell(2);
						 cell11.setCellValue(rs3.getString(3));
						 cell11.setCellStyle(csBorder());
						 
						 Cell cell12=row.createCell(3);
						 cell12.setCellValue(rs3.getString(4));
						 cell12.setCellStyle(csBorder());
						 
						 Cell cell13=row.createCell(4);
						 cell13.setCellValue(rs3.getString(5));
						 cell13.setCellStyle(csBorder());
						 
						 Cell cell14=row.createCell(5);
						 cell14.setCellValue(rs3.getString(6));
						 cell14.setCellStyle(csBorder());
						 
						 Cell cell15=row.createCell(6);
						 cell15.setCellValue(rs3.getString(7));
						 cell15.setCellStyle(csBorder());
						 
						 Cell cell16=row.createCell(7);
						 cell16.setCellValue(rs3.getString(8));
						 cell16.setCellStyle(csBorder());
						 
						 index3++;
						 } 
						 rs3.close();
				         stmt3.close();
			            
				       //Pendency data
						 rs4=stmt4.executeQuery();
						 int index4=1;
						 while(rs4.next()) 
						 {
						 XSSFRow row = (XSSFRow) sheet4.createRow(index4);
						 Cell cell17=row.createCell(0);
						 cell17.setCellValue(rs4.getString(1));
						 cell17.setCellStyle(csBorder());
						 
						 Cell cell18=row.createCell(1);
						 cell18.setCellValue(rs4.getString(2));
						 cell18.setCellStyle(csBorder());
						 
						 Cell cell19=row.createCell(2);
						 cell19.setCellValue(rs4.getString(3));
						 cell19.setCellStyle(csBorder());
						 
						 Cell cell20=row.createCell(3);
						 cell20.setCellValue(rs4.getString(4));
						 cell20.setCellStyle(csBorder());
						 
						 Cell cell21=row.createCell(4);
						 cell21.setCellValue(rs4.getString(5));
						 cell21.setCellStyle(csBorder());
						 
						 Cell cell22=row.createCell(5);
						 cell22.setCellValue(rs4.getString(6));
						 cell22.setCellStyle(csBorder());
						 
						 Cell cell23=row.createCell(6);
						 cell23.setCellValue(rs4.getString(7));
						 cell23.setCellStyle(csBorder());
						 
						 Cell cell24=row.createCell(7);
						 cell24.setCellValue(rs4.getString(8));
						 cell24.setCellStyle(csBorder());
						 
						 Cell cell25=row.createCell(8);
						 cell25.setCellValue(rs4.getString(9));
						 cell25.setCellStyle(csBorder());
						 
						 Cell cell26=row.createCell(9);
						 cell26.setCellValue(rs4.getString(10));
						 cell26.setCellStyle(csBorder());
						 
						 Cell cell27=row.createCell(10);
						 cell27.setCellValue(rs4.getString(11));
						 cell27.setCellStyle(csBorder());
						 
						 Cell cell28=row.createCell(11);
						 cell28.setCellValue(rs4.getString(12));
						 cell28.setCellStyle(csBorder());
						 
						 Cell cell29=row.createCell(12);
						 cell29.setCellValue(rs4.getString(13));
						 cell29.setCellStyle(csBorder());
						 
						 Cell cell30=row.createCell(13);
						 cell30.setCellValue(rs4.getString(14));
						 cell30.setCellStyle(csBorder());
						 
						 Cell cell31=row.createCell(14);
						 cell31.setCellValue(rs4.getString(15));
						 cell31.setCellStyle(csBorder());
						 
						 Cell cell32=row.createCell(15);
						 cell32.setCellValue(rs4.getString(16));
						 cell32.setCellStyle(csBorder());
						 
						 Cell cell33=row.createCell(16);
						 cell33.setCellValue(rs4.getString(17));
						 cell33.setCellStyle(csBorder());
						 
						 Cell cell34=row.createCell(17);
						 cell34.setCellValue(rs4.getString(18));
						 cell34.setCellStyle(csBorder());
						 index4++;					 
						 } 
						 rs4.close();
						 stmt4.close();
						 
						 //Axis Data
						 rs5=stmt5.executeQuery(); 
						 int index5=1;
						 while(rs5.next()) 
						 { 
						 XSSFRow row = (XSSFRow) sheet5.createRow(index5);

						 Cell cell35=row.createCell(0);
						 cell35.setCellValue(rs5.getString(1));
						 cell35.setCellStyle(csBorder());
						 
						 Cell cell36=row.createCell(1);
						 cell36.setCellValue(rs5.getString(2));
						 cell36.setCellStyle(csBorder());
						 
						 Cell cell37=row.createCell(2);
						 cell37.setCellValue(rs5.getString(3));
						 cell37.setCellStyle(csBorder());
						 
						 Cell cell38=row.createCell(3);
						 cell38.setCellValue(rs5.getString(4));
						 cell38.setCellStyle(csBorder());
						 
						 Cell cell39=row.createCell(4);
						 cell39.setCellValue(rs5.getString(5));
						 cell39.setCellStyle(csBorder());
						 
						 Cell cell40=row.createCell(5);
						 cell40.setCellValue(rs5.getString(6));
						 cell40.setCellStyle(csBorder());
						 
						 Cell cell41=row.createCell(6);
						 cell41.setCellValue(rs5.getString(7));
						 cell41.setCellStyle(csBorder());
						 
						 Cell cell42=row.createCell(7);
						 cell42.setCellValue(rs5.getString(8));
						 cell42.setCellStyle(csBorder());
						 
						 Cell cell43=row.createCell(8);
						 cell43.setCellValue(rs5.getString(9));
						 cell43.setCellStyle(csBorder());
						 
						 Cell cell44=row.createCell(9);
						 cell44.setCellValue(rs5.getString(10));
						 cell44.setCellStyle(csBorder());
						 
						 Cell cell45=row.createCell(10);
						 cell45.setCellValue(rs5.getString(11));
						 cell45.setCellStyle(csBorder());
						 
						 Cell cell46=row.createCell(11);
						 cell46.setCellValue(rs5.getString(12));
						 cell46.setCellStyle(csBorder());
						 
						 Cell cell47=row.createCell(12);
						 cell47.setCellValue(rs5.getString(13));
						 cell47.setCellStyle(csBorder());
						 
						 Cell cell48=row.createCell(13);
						 cell48.setCellValue(rs5.getString(14));
						 cell48.setCellStyle(csBorder());
						 
						 Cell cell49=row.createCell(14);
						 cell49.setCellValue(rs5.getString(15));
						 cell49.setCellStyle(csBorder());
						 
						 Cell cell50=row.createCell(15);
						 cell50.setCellValue(rs5.getString(16));
						 cell50.setCellStyle(csBorder());
						 
						 Cell cell51=row.createCell(16);
						 cell51.setCellValue(rs5.getString(17));
						 cell51.setCellStyle(csBorder());
						 
						 Cell cell52=row.createCell(17);
						 cell52.setCellValue(rs5.getString(18));
						 cell52.setCellStyle(csBorder());
						 
						 
						 index5++;
						 } 
						 rs5.close();
				         stmt5.close();
				         
				        //Set Column Width 
				         
				         for(int i=0;i<30;i++)
				         {
				         sheet1.autoSizeColumn(i);
				         }
				         
				         for(int i=0;i<9;i++)
				         {
				         sheet2.autoSizeColumn(i);
				         }
				         
				         for(int i=0;i<8;i++)
				         {
				         sheet3.autoSizeColumn(i);
				         }
				         
				         for(int i=0;i<18;i++)
				         {
				         sheet4.autoSizeColumn(i);
				         }
				         
				         for(int i=0;i<18;i++)
				         {
				         sheet5.autoSizeColumn(i);
				         }
				         sheetRowCountList.add(sheet2.getPhysicalNumberOfRows());
				         sheetRowCountList.add(sheet3.getPhysicalNumberOfRows());
				         sheetRowCountList.add(sheet4.getPhysicalNumberOfRows());
				         sheetRowCountList.add(sheet5.getPhysicalNumberOfRows());
				         
				         
						//Write File	         
			            FileWrite(reportDate);
			            connection.close(); 
			            
						 }         
					catch (SQLException se) 
				     {      
				     logger.error("SQL Exception",se);    
				     }
				
					 catch (Exception e) 
					 {      
					 logger.error("Exception",e);     
					 }
					 logger.debug("Report Generation Successful");
					 return sheetRowCountList;
					 }
		}
