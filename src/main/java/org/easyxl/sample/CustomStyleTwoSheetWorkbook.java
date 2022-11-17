package org.easyxl.sample;

import java.util.ArrayList;

import org.easyxl.ExcelGenerator;
import org.easyxl.ExcelGenerator.horizontalAllignment;

public class CustomStyleTwoSheetWorkbook {

		public static void main(String[] args) {
		// TODO Auto-generated method stub
	
		String[] mainHeader1 = new String[]{"Header1","Header2", "Header3"};
		String[] subHeader1 = new String[]{"SubHeader1","SubHeader2", "SubHeader3","SubHeader4","SubHeader5", "SubHeader6"};
		
		ArrayList<String[]> dataList = new ArrayList<String[]>();
		
		dataList.add( new String[]{"data11","data12", "data13","data14","data15", "data16"});
		dataList.add( new String[]{"data21","data22", "data23","data24","data25", "data26"});
		dataList.add( new String[]{"data31","data32", "data33","data34","data35", "data36"});
		dataList.add( new String[]{"data41","data42", "data43","data44","data45", "data46"});
		dataList.add( new String[]{"data51","data52", "data53","data54","data55", "data56"});
		
		
		
	
		ExcelGenerator egm = new ExcelGenerator();
		
		///Add custom Styles
		
		egm.addCustomStyle("customTitleStyle", "56887d", "fff0f5", "Harrington", 14, horizontalAllignment.CENTER, true, true, true);
		egm.addCustomStyle("customHeaderStyle1", "#8B7d99", "0A4859", null, 12, horizontalAllignment.CENTER, true, false, false);
		egm.addCustomStyle("customHeaderStyle2", "#8B7d99", "BDD7EE", null, 12, horizontalAllignment.CENTER, true, false, false);
		egm.addCustomStyle("customDataStyle", "fff0f5", "#548235", null, 11, horizontalAllignment.LEFT, false, false, false);
		
		
		
		egm.createNewExcelSheet("FirstSheet");
		
		
		egm.newRow();
		egm.addCustomStyledData("SAMPLE TITLE BAR 1",6,"customTitleStyle");
		
		egm.newRow(2);
		egm.addCustomStyledData("Header","customHeaderStyle1");
		egm.newRow();
		egm.addCustomStyledData("data11","customDataStyle");
		egm.newRow();
		egm.addCustomStyledData("data21","customDataStyle");
		egm.newRow();
		egm.addCustomStyledData("data31","customDataStyle");
		
		
		egm.dataStylToggleReset();
		
		egm.newRow(5);
		egm.addCustomStyledData("SAMPLE TITLE BAR 2",8,"customTitleStyle");
		
		egm.newRow(2);
		egm.addCustomStyledData(mainHeader1,2,"customHeaderStyle1");
		
		
		egm.newRow();
		egm.addCustomStyledData(subHeader1,"customHeaderStyle2");
		egm.newRow();
		
		for(String[] dataRow : dataList) {
			egm.addCustomStyledData(dataRow,"customDataStyle");
			egm.dataStylToggle();
			egm.newRow();
		}
    	
		
		
		egm.createNewExcelSheet("secondSheet");
		
		
		egm.newRow();
		egm.addCustomStyledData("SHEET 2 SAMPLE TITLE BAR 1",6,"customTitleStyle");
		
		egm.newRow(2);
		egm.addCustomStyledData("Header","customHeaderStyle1");
		egm.newRow();
		egm.addCustomStyledData("data11","customDataStyle");
		egm.newRow();
		egm.addCustomStyledData("data21","customDataStyle");
		egm.newRow();
		egm.addCustomStyledData("data31","customDataStyle");
		
		
		egm.newRow(5);
		egm.addCustomStyledData("SHEET 2 SAMPLE TITLE BAR 2",8,"customTitleStyle");
		egm.newRow(2);
		egm.addCustomStyledData(mainHeader1,2,"customHeaderStyle1");
		egm.newRow();
		egm.addCustomStyledData(subHeader1,"customHeaderStyle2");
		egm.newRow();
		
		for(String[] dataRow : dataList) {
			egm.addCustomStyledData(dataRow,"customDataStyle");
			egm.newRow();
		}
		
    	
		egm.saveFile("X:/path","CustomStyleTwoSheetWorkbook");
			

		
	}





}
