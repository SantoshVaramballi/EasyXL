package org.excel;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelGenerator {
//test comment

    StringBuffer xl = new StringBuffer();
    XSSFRow row = null;


    protected XSSFSheet sheet = null;
    protected XSSFWorkbook workBook = new XSSFWorkbook();
    protected int titleRowNo = 0;
    int columnCount = 0;
    int sheetCount = 1;
    int newRowIndex = 0;
    int rowIndex = 0;
    int maxColumnCount = 0;
    String userDate = null;
    String font = "calibri";
    private HashMap<String,XSSFCellStyle> customStylesMap  = new HashMap<String,XSSFCellStyle>();


	public ExcelGenerator(){}





	private XSSFColor  xssfColorBlack = new XSSFColor(new byte[]{0, 0,0},new DefaultIndexedColorMap());

   private boolean dataStyleToggle = false;


	public enum hdStl { style1, style2, Title }
	XSSFCellStyle defaultCellStyle = null;

    XSSFCellStyle cellStyleHeaderTitle = null;

    XSSFCellStyle cellStyleHeader1 = null;
    XSSFCellStyle cellStyleHeader2 = null;
    XSSFCellStyle cellStyleData1 = null;
    XSSFCellStyle cellStyleData2 = null;





    public XSSFWorkbook getExcelWorkBook() {

    	return  this.workBook;


    }


    public void createNewExcelSheet (String sheetName) {
    	 newRowIndex = 0;
    	 maxColumnCount = 0;
    	try {

    	    sheet = this.workBook.createSheet(sheetName);
    	    sheetCount = 1;
    	    sheet.setDisplayGridlines(false);
    	    sheet.setDefaultColumnWidth(20);
    	    dataStylToggleReset();

    	    /*--------------------------------------------------Assigning Different Cell Styles ------------------------------------------------------------------*/
    	    //
    	    System.out.println("Title - setting format data.");
    	    byte[] hexBGColorHedTitle= new  byte[] { (byte) 255, (byte) 127, (byte) 80 };
    	    byte[] hexFontColorHedTitle = new  byte[]  { (byte) 0, (byte) 0, (byte) 0 };
    	    String fontHedTitle= "calibri";
    	    Boolean isBoldHedTitle = true;
    	    Boolean isItalicHedTitle = false;
    	    Boolean isUnderlineHedTitle= false;
    	    double fontSizeHedTitle= 11;
    	    String allignmentHedTitle= "Left";

    	    cellStyleHeaderTitle = setCellStyle(hexBGColorHedTitle, hexFontColorHedTitle, fontHedTitle, isBoldHedTitle, isItalicHedTitle, isUnderlineHedTitle, fontSizeHedTitle, allignmentHedTitle);

    	    // Header1

    	    byte[] hexBGColorHed1 = new  byte[] { (byte) 0, (byte) 82, (byte) 224 };
    	    byte[] hexFontColorHed1= new  byte[] { (byte) 255, (byte) 255, (byte) 255 };
    	    String fontHed1= font;
    	    Boolean isBoldHed1= true;
    	    Boolean isItalicHed1 = false;
    	    Boolean isUnderlineHed1 = false;
    	    double fontSizeHed1 = 11;
    	    String allignmentHed1 ="Center";

    	    cellStyleHeader1 = setCellStyle(hexBGColorHed1, hexFontColorHed1, fontHed1, isBoldHed1, isItalicHed1, isUnderlineHed1, fontSizeHed1, allignmentHed1);

    	    // Header2
    	    System.out.println("Header2 - setting format data.");
    	    byte[] hexBGColorHed2 = new  byte[] { (byte) 0, (byte) 166, (byte) 184 };
    	    byte[] hexFontColorHed2 = new  byte[] { (byte) 0, (byte) 0, (byte) 0 };
    	    String fontHed2= font;
    	    Boolean isBoldHed2 = true;
    	    Boolean isItalicHed2= false;
    	    Boolean isUnderlineHed2= false;
    	    double fontSizeHed2 = 11;
    	    String allignmentHed2 = "Center";

    	    cellStyleHeader2 = setCellStyle(hexBGColorHed2, hexFontColorHed2, fontHed2, isBoldHed2, isItalicHed2, isUnderlineHed2, fontSizeHed2, allignmentHed2);



    	    // Data1
    	    System.out.println("Data1 - setting format data.");
    	    byte[] hexBGColorData1 = new  byte[] { (byte) 255, (byte) 255, (byte) 255 };
    	    byte[] hexFontColorData1= new  byte[] { (byte) 0, (byte) 0, (byte) 0 };
    	    String fontData1= font;
    	    Boolean isBoldData1 = false;
    	    Boolean isItalicData1= false;
    	    Boolean isUnderlineData1= false;
    	    double fontSizeData1 = 10.5;
    	    String allignmentData1 = "Left";

    	    cellStyleData1 = setCellStyle(hexBGColorData1, hexFontColorData1, fontData1, isBoldData1, isItalicData1, isUnderlineData1, fontSizeData1, allignmentData1);


    	    // Data2
    	    System.out.println("Data2 - setting format data.");
    	    byte[] hexBGColorData2= new  byte[]  { (byte) 232, (byte) 238, (byte) 248 };
    	    byte[] hexFontColorData2  = new  byte[]{ (byte) 0, (byte) 0, (byte) 0 };
    	    String fontData2= font;
    	    Boolean isBoldData2= false;
    	    Boolean isItalicData2= false;
    	    Boolean isUnderlineData2= false;
    	    double fontSizeData2 = 10.5;
    	    String allignmentData2 = "Left";

    	    cellStyleData2 = setCellStyle(hexBGColorData2, hexFontColorData2, fontData2, isBoldData2, isItalicData2, isUnderlineData2, fontSizeData2, allignmentData2);


    	    /*--------------------------------------------------Done Assigning Different Cell Styles ------------------------------------------------------------------*/

    	} catch (Exception ex) {
    	     System.out.println("generateExcelFile" + ex);
    	    ex.printStackTrace();
    	}
        }



    protected void createStyledCellNew(XSSFRow row, String value, hdStl headerStyleType) {
	XSSFCell contentCell = null;

	int cellCount = (row.getLastCellNum() < 0 ? 0 : row.getLastCellNum());
	contentCell = row.createCell(cellCount);
	contentCell.setCellValue(new XSSFRichTextString(value));

	if (headerStyleType.equals(hdStl.style1)) {
	    contentCell.setCellStyle(cellStyleHeader1);

	} else if (headerStyleType.equals(hdStl.Title)) {
	    contentCell.setCellStyle(cellStyleHeaderTitle);

	} else {
	    contentCell.setCellStyle(cellStyleHeader2);

	}

    }

    protected void createCustomStyledCellNew(XSSFRow row, String value, String customStyleName) {

	//Check Style exists ?
	if(customStylesMap.containsKey(customStyleName)) {
		//If yes Continue
		createStyledCell(row, value, customStylesMap.get(customStyleName));
	}

	else {
		//Else Print Debug and Set default.
		System.out.println("Style " +customStyleName+" Not defined, Using Default style.");
		createStyledCell(row, value, defaultCellStyle);
	}


    }

    protected void createStyledCell(XSSFRow row, String value, XSSFCellStyle xssfCellStyle) {
	XSSFCell contentCell = null;

	int cellCount = (row.getLastCellNum() < 0 ? 0 : row.getLastCellNum());
	contentCell = row.createCell(cellCount);
	contentCell.setCellValue(new XSSFRichTextString(value));
	contentCell.setCellStyle(xssfCellStyle);
    }

    protected void createNonStyledCellNew(XSSFRow row, String value) {

	XSSFCell contentCell = null;
	int cellCount = (row.getLastCellNum() < 0 ? 0 : row.getLastCellNum());
	contentCell = row.createCell(cellCount);
	contentCell.setCellValue(new XSSFRichTextString(value));
	if (dataStyleToggle) {
	    contentCell.setCellStyle(cellStyleData2);
	} else {
	    contentCell.setCellStyle(cellStyleData1);
	}

    }


    protected XSSFCellStyle setCellStyle(byte[] hexBGColor, byte[] hexFontColor, String font, Boolean bold,
	    Boolean italic, Boolean Underline, double size, String allignment) {
	XSSFCellStyle cellStyle = workBook.createCellStyle();
	XSSFFont cellFont = workBook.createFont();


	cellStyle.setLeftBorderColor(xssfColorBlack);
	cellStyle.setBorderLeft(BorderStyle.THIN);

	cellStyle.setRightBorderColor(xssfColorBlack);
	cellStyle.setBorderRight(BorderStyle.THIN);

	cellStyle.setTopBorderColor(xssfColorBlack);
	cellStyle.setBorderTop(BorderStyle.THIN);

	cellStyle.setBottomBorderColor(xssfColorBlack);
	cellStyle.setBorderBottom(BorderStyle.THIN);

	cellStyle.setFillForegroundColor(new XSSFColor(hexBGColor));
	cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

	if (allignment.equalsIgnoreCase("Left")) {
	    cellStyle.setAlignment(HorizontalAlignment.LEFT);
	} else if (allignment.equalsIgnoreCase("Right")) {
	    cellStyle.setAlignment(HorizontalAlignment.RIGHT);
	} else {
	    cellStyle.setAlignment(HorizontalAlignment.CENTER);
	}

	cellFont.setColor(new XSSFColor(hexFontColor));
	cellFont.setBold(bold);
	cellFont.setItalic(italic);
	if (Underline) {
	    cellFont.setUnderline(Font.U_SINGLE);
	}
	// cellFont.setFontName("calibri");
	cellFont.setFontName(font);
	cellFont.setFontHeight(size);

	cellStyle.setFont(cellFont);

	return cellStyle;
    }

    /**
     * Creates new row and returns to the newest row.
     *
     * @return
     */
    public XSSFRow createRow() {
	XSSFRow row = null;
	if (((rowIndex + 1) % 1048576) == 0) {

	    newRowIndex = 1;
	    sheetCount++;
	    if(sheetCount == 2) {
	    	this.workBook.setSheetName(workBook.getSheetIndex(sheet), sheet.getSheetName()+"-1");
	    }
	    sheet = workBook.createSheet(sheet.getSheetName()+"-"+ sheetCount);
	}
	row = sheet.createRow(newRowIndex);
	newRowIndex++;
	rowIndex++;
	return row;
    }

    /**
     * Creates new rows and returns to the newest row.
     *
     * @param numOfRows
     * @return
     */
    public XSSFRow createRow(int numOfRows) {
	XSSFRow newRow = null;
	for (int i = 0; i < numOfRows; i++) {
	    newRow = createRow();

	}
	return newRow;
    }

    public void newRow(int noOfRows){
	row = createRow(noOfRows);
    }

    public void newRow(){
	row = createRow();
    }
    /**
     * Adds data into single cell without styling(data styling).
     *
     * @param data(String) -> String data into single cell
     */
    public void addData(String data) {
	xl.setLength(0);
	xl.append(data);
	createNonStyledCellNew(row, xl.toString());
	updateMaxColumnCount();

    }

    /**
     * Adds data into single cell without styling(data styling).
     *
     * @param dataArray(String[]) -> Each string data into single cell
     */
    public void addData(String[] dataArray) {

	for (String data : dataArray) {
	    xl.setLength(0);
	    xl.append(data);
	    createNonStyledCellNew(row, xl.toString());
	    updateMaxColumnCount();
	}

    }


    public void addData(String[] dataArray,int lastColumn) {

	for (String data : dataArray) {

		addData(data,lastColumn);
	}

    }
    public void addHeaderData(String data, hdStl headerStyleType) {
	xl = new StringBuffer("");
	xl.append(data);
	createStyledCellNew(row, xl.toString(), headerStyleType);
	updateMaxColumnCount();

    }

    public void addCustomStyledData(String data, String customStyleName) {
	xl = new StringBuffer("");
	xl.append(data);
	createCustomStyledCellNew(row, xl.toString(), customStyleName);
	updateMaxColumnCount();

    }


    protected void updateMaxColumnCount() {
	if (row.getLastCellNum() > this.maxColumnCount) {
	    this.maxColumnCount = row.getLastCellNum();

	}

    }

    protected void updateMaxColumnCount(int numOfColumn) {
 	if (row.getLastCellNum()+numOfColumn > this.maxColumnCount) {
 	    this.maxColumnCount = row.getLastCellNum()+numOfColumn;

 	}

     }
    /**
     * Add styled(Heading/Title) data to single merged cell
     *
     * @param content(String) -> Content of cell
     * @param numberOfColumn(int) -> Number of columns to be merged
     * @param headerStyleType(enum hdStl) -> Style of cell
     *
     *
     */
    public void addHeaderData(String content, int numberOfColumn, hdStl headerStyleType) {
	int rowNum = row.getRowNum();

	int firstColumn = row.getLastCellNum();
	firstColumn = (firstColumn < 0 ? 0 : firstColumn);

	addHeaderData(content, headerStyleType);
	numberOfColumn = firstColumn + numberOfColumn-1;

	sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn));
	RegionUtil.setBorderBottom( BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	updateMaxColumnCount();
    }


    /**
     * Add styled(Heading/Title) data to single cell
     *
     * @param dataArray (String[]) -> Adds each string object to 1 cell
     * @param headerStyleType (enum hdStl) -> Style of cell
     */
    public void addHeaderData(String[] dataArray, hdStl headerStyleType) {
	for (String data : dataArray) {
	    xl = new StringBuffer("");
	    xl.append(data);
	    createStyledCellNew(row, xl.toString(), headerStyleType);
	    updateMaxColumnCount();
	}

    }

    public void addHeaderData(String[] dataArray, int columnWidth, hdStl headerStyleType) {
    	int numberOfColumn=1;
	for (String data : dataArray) {
	    int rowNum = row.getRowNum();

	    int firstColumn = row.getLastCellNum();

	    firstColumn = (firstColumn < 0 ? 0 : firstColumn);

	    addHeaderData(data, headerStyleType);
	    numberOfColumn = firstColumn + columnWidth-1;

	   int X = row.getLastCellNum();

	    sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn));
	    int Y = row.getLastCellNum();
	    updateMaxColumnCount();
	    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	}
    }
    public void addData(String content, int lastColumn) {
	int rowNum = row.getRowNum();

	int firstColumn = row.getLastCellNum();

	if (firstColumn < 0) {
	    firstColumn = 0;
	}
	lastColumn = firstColumn + lastColumn-1;
	addData(content);
	sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, firstColumn, lastColumn));
	RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, lastColumn), sheet);
		RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, lastColumn), sheet);
		RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, lastColumn), sheet);
		RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, lastColumn), sheet);
		updateMaxColumnCount();


    }

    protected void nextCell(int numOfCells) {
	row.createCell(row.getLastCellNum() + (numOfCells));
    }

    /**
     * Switches between the 2 styles of data (valid only for addData not addHeaderData)
     */
    public void dataStylToggle() {
	dataStyleToggle = !dataStyleToggle;
    }

    public void dataStylToggleReset() {
	dataStyleToggle = false;
    }




    	public static byte[] hex2Rgb(String colorStr) {

    		 int r= Integer.valueOf(colorStr.substring(1,3), 16);
    		 int g= Integer.valueOf(colorStr.substring(3,5), 16);
    		 int b= Integer.valueOf(colorStr.substring(5,7), 16);

    		 return new byte[] {(byte)r,(byte)g,(byte)b};
    }

}
