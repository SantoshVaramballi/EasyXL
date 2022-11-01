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
}
