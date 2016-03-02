package com.novopay.masteronboarding.report.prod;

import java.awt.image.BufferedImage;
import java.awt.image.ColorConvertOp;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import javax.imageio.ImageIO;

import org.apache.log4j.Logger;
/**
* This class crops the image to a specified size.
* 
* @author  Mitrabhanu
* @version 1.0
* @since   2016-02-17 
*/

public class ImageCropping {
	private static Logger logger=Logger.getLogger(ImageCropping.class);
	
	 /**
     * This is a method to convert CMYK image to RGB image.
     * 
     * 
     * @return BufferedImage This returns image in RBG format.
     * 
     */	
	
	private BufferedImage convertCMYK2RGB(BufferedImage image) throws IOException{
	   
		//Create a new RGB image
	    BufferedImage rgbImage = new BufferedImage(image.getWidth(), image.getHeight(),
	    BufferedImage.TYPE_3BYTE_BGR);
	    //color conversion
	    ColorConvertOp op = new ColorConvertOp(null);
	    op.filter(image, rgbImage);
	    return rgbImage;
	}
	
	 /**
     * This is a method to crop images to a specified size.
     * 
     * 
     * @return Nothing
     */	
	void conversionImplementation(ArrayList<Integer> sheetRowCountList)
	{
		try {
			BufferedImage imgage1 = ImageIO.read(new File("./Image/Zone & Area Wise Data.jpeg"));
			BufferedImage imgage2 = ImageIO.read(new File("./Image/Partner Wise Data.jpeg"));
			BufferedImage imgage3 = ImageIO.read(new File("./Image/Pendency in Sales Bin.jpeg"));
			BufferedImage imgage4 = ImageIO.read(new File("./Image/Axis Partner, Awaiting for Devi.jpeg"));
			
			ImageCropping image=new ImageCropping();
			BufferedImage convertedImage1=image.convertCMYK2RGB(imgage1);
			BufferedImage convertedImage2=image.convertCMYK2RGB(imgage2);
			BufferedImage convertedImage3=image.convertCMYK2RGB(imgage3);
			BufferedImage convertedImage4=image.convertCMYK2RGB(imgage4);
			
			BufferedImage CroppedImgage1 = convertedImage1.getSubimage(58, 58, 1400, (sheetRowCountList.get(0)+2)*20);
			BufferedImage CroppedImgage2 = convertedImage2.getSubimage(61, 61, 1270, (sheetRowCountList.get(1)+2)*20);
			BufferedImage CroppedImgage3 = convertedImage3.getSubimage(60, 65, 1285, (sheetRowCountList.get(2)+2)*20);
			BufferedImage CroppedImgage4 = convertedImage4.getSubimage(58, 60, 1360, (sheetRowCountList.get(3)+2)*20);
			

			File outputfile1 = new File("./Image/Zone & Area Wise Data.jpeg");
			ImageIO.write(CroppedImgage1, "jpeg", outputfile1);
			File outputfile2 = new File("./Image/Partner Wise Data.jpeg");
			ImageIO.write(CroppedImgage2, "jpeg", outputfile2);
			File outputfile3 = new File("./Image/Pendency in Sales Bin.jpeg");
			ImageIO.write(CroppedImgage3, "jpeg", outputfile3);
			File outputfile4 = new File("./Image/Axis Partner, Awaiting for Devi.jpeg");
			ImageIO.write(CroppedImgage4, "jpeg", outputfile4);

		} 
		
		catch (IOException ie) 
		{
			logger.error("IOException",ie);
		}
		
		catch (Exception e) 
		{
			logger.error("Exception",e);
		}
		logger.debug("Image Cropping Complete");
	}
}