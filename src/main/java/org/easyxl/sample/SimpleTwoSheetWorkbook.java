package org.easyxl.sample;

import java.util.ArrayList;

import org.easyxl.ExcelGenerator;
import org.easyxl.ExcelGenerator.hdStl;

public class SimpleTwoSheetWorkbook {

	



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
		
		egm.createNewExcelSheet("FirstSheet");
		
		
		egm.newRow();
		egm.addHeaderData("SAMPLE TITLE BAR 1",6,hdStl.TITLE);
		
		egm.newRow(2);
		egm.addHeaderData("Header",hdStl.STYLE1);
		egm.newRow();
		egm.addData("data11");
		egm.newRow();
		egm.dataStylToggle();
		egm.addData("data21");
		egm.newRow();
		egm.dataStylToggle();
		egm.addData("data31");
		
		
		egm.dataStylToggleReset();
		
		egm.newRow(5);
		egm.addHeaderData("SAMPLE TITLE BAR 2",8,hdStl.TITLE);
		
		egm.newRow(2);
		egm.addHeaderData(mainHeader1,2,hdStl.STYLE1);
		
		
		egm.newRow();
		egm.addHeaderData(subHeader1,hdStl.STYLE2);
		egm.newRow();
		
		for(String[] dataRow : dataList) {
			egm.addData(dataRow);
			egm.dataStylToggle();
			egm.newRow();
		}
		
		

		

    	
    	
		egm.createNewExcelSheet("secondSheet");
		
		
		egm.newRow();
		egm.addHeaderData("SHEET 2 SAMPLE TITLE BAR 1",6,hdStl.TITLE);
		
		egm.newRow(2);
		egm.addHeaderData("Header",hdStl.STYLE1);
		egm.newRow();
		egm.addData("data11");
		egm.newRow();
		egm.dataStylToggle();
		egm.addData("data21");
		egm.newRow();
		egm.dataStylToggle();
		egm.addData("data31");
		
		
		egm.newRow(5);
		egm.addHeaderData("SHEET 2 SAMPLE TITLE BAR 2",8,hdStl.TITLE);
		egm.newRow(2);
		egm.addHeaderData(mainHeader1,2,hdStl.STYLE1);
		egm.newRow();
		egm.addHeaderData(subHeader1,hdStl.STYLE2);
		egm.newRow();
		
		for(String[] dataRow : dataList) {
			egm.addData(dataRow);
			egm.dataStylToggle();
			egm.newRow();
		}
		
    	
		egm.saveFile("X:/path","SimpleTwoSheetWorkbook");
			

		
	}





}
