package com.novopay.masteronboarding.report.prod;

import java.io.IOException;

import org.apache.log4j.Logger;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
/**
* The class converts all the sheets in an excel file to images.
* and stores the images in a specified location.
* 
* 
* @author  Mitrabhanu
* @version 1.0
* @since   2016-02-17 
*/

public class ExcelToImageConvertion 
{
	private static Logger logger=Logger.getLogger(ExcelToImageConvertion.class);
	/**
     * This is a method to convert excel sheets to image.This user Aspose Cells API for the conversion.
     * 
     * @return Nothing
     */
	
	void ImageConversion(String reportGenerationDate)
	{
	try
    {   
		//Sets the path for font which is calibri by default.
		CellsHelper.setFontDir("/usr/share/fonts/truetype/msttcorefonts/");
		//Instantiate a new Workbook object
		//Open template
		Workbook book = new Workbook("./Report/Report_"+reportGenerationDate+".xlsx");

		//Iterate over all worksheets in the workbook
		for (int i = 1; i < book.getWorksheets().getCount(); i++)
		{
			Worksheet sheet = book.getWorksheets().get(i);
			
			//Apply different Image and Print options
			ImageOrPrintOptions options = new ImageOrPrintOptions();

			//Set Horizontal Resolution
			options.setHorizontalResolution(100);

			//Set Vertical Resolution
			options.setVerticalResolution(100);

			//Set Image Format
			options.setImageFormat(ImageFormat.getJpeg());

			//If you want entire sheet as a single image
			options.setOnePagePerSheet(true);
			
			//Render to image
		    SheetRender sr = new SheetRender(sheet, options);
		    sr.toImage(0,"./Image/"+sheet.getName()+".jpeg");
		}
	}
	catch(IOException ie)
	{
	logger.error("IOExxception",ie);	
		
	}
	
	catch (Exception e)
	{
		logger.error("IOExxception",e);
	}
	logger.debug("Excel coversion to Image Complete");
}}

