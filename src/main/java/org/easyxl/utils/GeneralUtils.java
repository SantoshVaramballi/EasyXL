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
try {
		 int r= Integer.valueOf(formatedColorString.substring(0,2), 16);
		 int g= Integer.valueOf(formatedColorString.substring(2,4), 16);
		 int b= Integer.valueOf(formatedColorString.substring(4,6), 16);

		 return new byte[] {(byte)r,(byte)g,(byte)b};
}
catch (Exception e) {
	System.out.println("Error in parsing Color Value :" + hexColorStr + "; Error " +e);
	return null;
}
 		}
 		else {
 			System.out.println("Invalid Color String ");
 			return null;
 		}
}

	private static String validateAndFormatString(String hexColorStr) {
		
		if (hexColorStr.length()==7 && hexColorStr.startsWith("#")) {
			hexColorStr = hexColorStr.replaceFirst("#", "");
			return hexColorStr;
		}
		else if(hexColorStr.length()==6 && !hexColorStr.startsWith("#")) {
			return hexColorStr;
		}
		else {
			System.out.println("Invalid Color String");
			return null;
		}
		
		
		
	}
	
}
