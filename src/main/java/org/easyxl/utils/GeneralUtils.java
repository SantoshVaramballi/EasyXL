package org.easyxl.utils;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GeneralUtils {

	
	 public static void writeXLToFile(XSSFWorkbook xlWorkbook, String path, String fileName) {
		
		
  	  FileOutputStream fos = null;
  	
			try {
				File xlfile = new File(path + System.getProperty("file.separator")+fileName + ".xlsx");
				xlfile.getParentFile().mkdirs();
				xlfile.createNewFile();
				fos = new FileOutputStream(xlfile);
				xlWorkbook.write(fos);
		  	    fos.flush();
		  	    fos.close();
		  	    
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				
			}
	}
	 
 	public static byte[] hex2Rgb(String hexColorStr) {
 		
 		String formatedColorString = validateAndFormatString(hexColorStr);
 		
 		if(formatedColorString != null) {

		 int r= Integer.valueOf(hexColorStr.substring(1,3), 16);
		 int g= Integer.valueOf(hexColorStr.substring(3,5), 16);
		 int b= Integer.valueOf(hexColorStr.substring(5,7), 16);

		 return new byte[] {(byte)r,(byte)g,(byte)b};
 		}
 		else {
 			System.out.println("Invalid Color String ");
 			return null;
 		}
}

	private static String validateAndFormatString(String hexColorStr) {
		
		//length 7 and #
		if (hexColorStr.length()==7 && hexColorStr.startsWith("#")) {
			return hexColorStr.replaceFirst("#", "");
		}
		//length 6 without #
		else if(hexColorStr.length()==6 && !hexColorStr.startsWith("#")) {
			return hexColorStr;
		}
		else {
			System.out.println("Invalid Color String");
			return null;
		}
		
		
		
	}
	
}
