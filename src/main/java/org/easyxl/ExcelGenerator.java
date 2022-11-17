package org.easyxl;

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
import org.easyxl.utils.GeneralUtils;

/**
 * @author SANTOSHVARAMBALLI
 *
 */
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

    
	private XSSFColor  xssfColorBlack = new XSSFColor(new byte[]{0, 0,0},new DefaultIndexedColorMap());

	   private boolean dataStyleToggle = false;

	   public enum horizontalAllignment { LEFT, RIGHT, CENTER }
		public enum hdStl { STYLE1, STYLE2, TITLE }
		XSSFCellStyle defaultCellStyle = null;

	    XSSFCellStyle cellStyleHeaderTitle = null;

	    XSSFCellStyle cellStyleHeader1 = null;
	    XSSFCellStyle cellStyleHeader2 = null;
	    XSSFCellStyle cellStyleData1 = null;
	    XSSFCellStyle cellStyleData2 = null;


	public ExcelGenerator(){
		
		
		
	}








    public XSSFWorkbook getExcelWorkBook() {
    	return  this.workBook;
    }
    
    
public void addCustomStyle (String StyleName,String  hexBGColor,String  hexFontColor,String font,double fontSize,horizontalAllignment allignment,
		Boolean isBold, Boolean isItalic, Boolean isUnderline) {
	 XSSFCellStyle newCustomCellStyle = null;
	//Convert Hex to byte
    byte[] bgColor= new  byte[] { (byte) 255, (byte) 255, (byte) 255 };
    byte[] fontColor = new  byte[]  { (byte) 0, (byte) 0, (byte) 0 };
    String finalIAllignment= "Left";
    
    
    byte[] bgColorTemp = GeneralUtils.hex2Rgb(hexBGColor);
    byte[] fontColorTemp = GeneralUtils.hex2Rgb(hexFontColor);
    
    if(bgColorTemp!=null && fontColorTemp != null) {
    	bgColor = bgColorTemp;
    	fontColor=fontColorTemp;
    }
    else {
    	System.out.println(StyleName + "- Setting default values for font and background color and supplied values are not valid Hex Code.");
    }
    
    if(allignment.equals(horizontalAllignment.RIGHT)) {
    	finalIAllignment= "Right";
    }
    else if(allignment.equals(horizontalAllignment.CENTER)) {
    	finalIAllignment= "Center";
    }
    
    
    
    newCustomCellStyle = setCellStyle(bgColor, fontColor, font, isBold, isItalic, isUnderline, fontSize, finalIAllignment);
	
    customStylesMap.put(StyleName, newCustomCellStyle);
    System.out.println( StyleName + "- Saved");
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
    	    byte[] bgColorHedTitle= new  byte[] { (byte) 255, (byte) 127, (byte) 80 };
    	    byte[] fontColorHedTitle = new  byte[]  { (byte) 0, (byte) 0, (byte) 0 };
    	    String fontHedTitle= "calibri";
    	    Boolean isBoldHedTitle = true;
    	    Boolean isItalicHedTitle = false;
    	    Boolean isUnderlineHedTitle= false;
    	    double fontSizeHedTitle= 11;
    	    String allignmentHedTitle= "Left";

    	    cellStyleHeaderTitle = setCellStyle(bgColorHedTitle, fontColorHedTitle, fontHedTitle, isBoldHedTitle, isItalicHedTitle, isUnderlineHedTitle, fontSizeHedTitle, allignmentHedTitle);

    	    // Header1

    	    byte[] bgColorHed1 = new  byte[] { (byte) 0, (byte) 82, (byte) 224 };
    	    byte[] fontColorHed1= new  byte[] { (byte) 255, (byte) 255, (byte) 255 };
    	    String fontHed1= font;
    	    Boolean isBoldHed1= true;
    	    Boolean isItalicHed1 = false;
    	    Boolean isUnderlineHed1 = false;
    	    double fontSizeHed1 = 11;
    	    String allignmentHed1 ="Center";

    	    cellStyleHeader1 = setCellStyle(bgColorHed1, fontColorHed1, fontHed1, isBoldHed1, isItalicHed1, isUnderlineHed1, fontSizeHed1, allignmentHed1);

    	    // Header2
    	    System.out.println("Header2 - setting format data.");
    	    byte[] bgColorHed2 = new  byte[] { (byte) 0, (byte) 166, (byte) 184 };
    	    byte[] fontColorHed2 = new  byte[] { (byte) 0, (byte) 0, (byte) 0 };
    	    String fontHed2= font;
    	    Boolean isBoldHed2 = true;
    	    Boolean isItalicHed2= false;
    	    Boolean isUnderlineHed2= false;
    	    double fontSizeHed2 = 11;
    	    String allignmentHed2 = "Center";

    	    cellStyleHeader2 = setCellStyle(bgColorHed2, fontColorHed2, fontHed2, isBoldHed2, isItalicHed2, isUnderlineHed2, fontSizeHed2, allignmentHed2);



    	    // Data1
    	    System.out.println("Data1 - setting format data.");
    	    byte[] bgColorData1 = new  byte[] { (byte) 255, (byte) 255, (byte) 255 };
    	    byte[] fontColorData1= new  byte[] { (byte) 0, (byte) 0, (byte) 0 };
    	    String fontData1= font;
    	    Boolean isBoldData1 = false;
    	    Boolean isItalicData1= false;
    	    Boolean isUnderlineData1= false;
    	    double fontSizeData1 = 10.5;
    	    String allignmentData1 = "Left";

    	    cellStyleData1 = setCellStyle(bgColorData1, fontColorData1, fontData1, isBoldData1, isItalicData1, isUnderlineData1, fontSizeData1, allignmentData1);


    	    // Data2
    	    System.out.println("Data2 - setting format data.");
    	    byte[] bgColorData2= new  byte[]  { (byte) 232, (byte) 238, (byte) 248 };
    	    byte[] fontColorData2  = new  byte[]{ (byte) 0, (byte) 0, (byte) 0 };
    	    String fontData2= font;
    	    Boolean isBoldData2= false;
    	    Boolean isItalicData2= false;
    	    Boolean isUnderlineData2= false;
    	    double fontSizeData2 = 10.5;
    	    String allignmentData2 = "Left";

    	    cellStyleData2 = setCellStyle(bgColorData2, fontColorData2, fontData2, isBoldData2, isItalicData2, isUnderlineData2, fontSizeData2, allignmentData2);


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

	if (headerStyleType.equals(hdStl.STYLE1)) {
	    contentCell.setCellStyle(cellStyleHeader1);

	} else if (headerStyleType.equals(hdStl.TITLE)) {
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


	protected XSSFCellStyle setCellStyle(byte[] bgColor, byte[] fontColor, String font, Boolean bold,
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

	cellStyle.setFillForegroundColor(new XSSFColor(bgColor));
	cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

	if (allignment.equalsIgnoreCase("Left")) {
	    cellStyle.setAlignment(HorizontalAlignment.LEFT);
	} else if (allignment.equalsIgnoreCase("Right")) {
	    cellStyle.setAlignment(HorizontalAlignment.RIGHT);
	} else {
	    cellStyle.setAlignment(HorizontalAlignment.CENTER);
	}

	cellFont.setColor(new XSSFColor(fontColor));
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
    


    /**
     * @param data
     * @param customStyleName
     */
    public void addCustomStyledData(String data, String customStyleName) {
	xl = new StringBuffer("");
	xl.append(data);
	createCustomStyledCellNew(row, xl.toString(), customStyleName);
	updateMaxColumnCount();
    }

    
    /**
     * @param data
     * @param columnWidth
     * @param customStyleName
     */
    public void addCustomStyledData(String data, int columnWidth, String customStyleName) {
    	int rowNum = row.getRowNum();

    	int firstColumn = row.getLastCellNum();
    	firstColumn = (firstColumn < 0 ? 0 : firstColumn);

    	addCustomStyledData(data, customStyleName);
    	columnWidth = firstColumn + columnWidth-1;

    	sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, firstColumn, columnWidth));
    	RegionUtil.setBorderBottom( BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, columnWidth), sheet);
    	RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, columnWidth), sheet);
    	RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, columnWidth), sheet);
    	RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, columnWidth), sheet);
    	updateMaxColumnCount();
        }
    
    
    /**
     * @param dataArray
     * @param customStyleName
     */
    public void addCustomStyledData(String[] dataArray, String customStyleName) {
	for (String data : dataArray) {
	    xl = new StringBuffer("");
	    xl.append(data);
	    createCustomStyledCellNew(row, xl.toString(), customStyleName);
	    updateMaxColumnCount();
	}

    }

    /**
     * @param dataArray
     * @param columnWidth
     * @param customStyleName
     */
    public void addCustomStyledData(String[] dataArray, int columnWidth, String customStyleName) {
    	//int numberOfColumn=1;
	for (String data : dataArray) {
		
		addCustomStyledData(data, columnWidth,customStyleName);
		/*
	    int rowNum = row.getRowNum();

	    int firstColumn = row.getLastCellNum();

	    firstColumn = (firstColumn < 0 ? 0 : firstColumn);

	    addCustomStyledData(data, customStyleName);
	    numberOfColumn = firstColumn + columnWidth-1;

	   int X = row.getLastCellNum();

	    sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn));
	    int Y = row.getLastCellNum();
	    updateMaxColumnCount();
	    RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	    RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	    RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
	*/}
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
    public void addHeaderData(String data, int numberOfColumn, hdStl headerStyleType) {
	int rowNum = row.getRowNum();

	int firstColumn = row.getLastCellNum();
	firstColumn = (firstColumn < 0 ? 0 : firstColumn);

	addHeaderData(data, headerStyleType);
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
    public void addData(String data, int numberOfColumn) {
	int rowNum = row.getRowNum();

	int firstColumn = row.getLastCellNum();

	if (firstColumn < 0) {
	    firstColumn = 0;
	}
	numberOfColumn = firstColumn + numberOfColumn-1;
	addData(data);
	sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn));
	RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
		RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
		RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
		RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, firstColumn, numberOfColumn), sheet);
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


    	
    	public void saveFile(String path, String fileName){
    		XSSFWorkbook xlWorkbook = getExcelWorkBook();
    		GeneralUtils.writeXLToFile(xlWorkbook,path, fileName);
    		System.out.println("File "+fileName+".xlsx is Saved");
    	}

}
