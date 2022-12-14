![GitHub](https://img.shields.io/github/license/leroy-merlin-br/logstash-exporter)  ![Output](https://img.shields.io/badge/Excel-xlsx-008000) ![JavaVersion](https://img.shields.io/badge/Java-%3E=%201.6-4A7493)


# EasyXL
EasyXL is a tool kit which helps to Generate Formatted Excel WorkBooks easily.



##### Table of Contents  

[Description](#description)  
[Sample Output](#sample-output)  
[JAR File](#jar-file)  
[Dependencies](#dependencies)  
[Usage](#usage)  
[Primary Methods](#primary-methods)  


# Description
EasyXL is a tool kit that helps to Generate Formatted Excel WorkBooks easily.
It is built on top of the Apache POI library.
It aims to 
1) Reduce the effort involved in creating code for generating Formatted Excel reports in JAVA. 
2) Increase code efficiency by creating the Format object once and reusing it for the entire workbook creation.

# Sample Output
![Sample_excel_file.JPG](https://github.com/SantoshVaramballi/EasyXL/blob/main/Sample%20outputs/Sample_excel_file.JPG)

# JAR File
Jar file is available [here](https://github.com/SantoshVaramballi/EasyXL/tree/main/EasyXL_Jars).

# Dependencies 
The File needs apache POI and its related dependencies.
The below entries need to be added to the pom.xml file.

```xml
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>5.2.2</version>
</dependency>
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.2</version>
</dependency>
```

# Usage
1) Add the [EasyXL.jar](https://github.com/SantoshVaramballi/EasyXL/tree/main/EasyXL_Jars) file to your project dependencies. 
2) Import the required modules as below. 
```java
import org.easyxl.ExcelGenerator;
import org.easyxl.ExcelGenerator.hdStl;
```

3) Create a Excel generator object using the code below.
```java
ExcelGenerator egm = new ExcelGenerator();
```
4) Populate the data using the methods mentioned in [Primary Methods](#primary-methods) 
5) Save the File to disk using the "saveFile" metod.
```java
egm.saveFile("File path","file_name");
```

NOTE: You can find the full code sample here [SimpleTwoSheetWorkbook](https://github.com/SantoshVaramballi/EasyXL/blob/main/src/main/java/org/easyxl/sample/SimpleTwoSheetWorkbook.java)


# Primary Methods
```java

/*Creating new sheet*/
void createNewExcelSheet (String sheetName);
```

```java
/*Navigating to new row*/
void newRow();
void newRow(int noOfRows);
```
```java
/*Adding header data*/
void addHeaderData(String data, hdStl headerStyleType);
void addHeaderData(String data, int numberOfColumn, hdStl headerStyleType);
void addHeaderData(String[] dataArray, hdStl headerStyleType);
```
```java
/*Adding non-header data*/
void addData(String data);
void addData(String[] dataArray);
void addData(String data, int numberOfColumn) 
```
```java
/*Switching non-header data format*/
void dataStylToggle();
void dataStylToggleReset();
```
```java
/*Saving file to disk*/
void saveFile(String path, String fileName);

```













