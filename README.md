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
1) Add the jar file to your project dependencies. 
2) Import the required modules as below. 
```
import org.easyxl.ExcelGenerator;
import org.easyxl.ExcelGenerator.hdStl;
```

3) Create a Excel generator object using the code below.
```
ExcelGenerator egm = new ExcelGenerator();
```
4) Populate the data using the methods mentioned in [Primary Methods](#primary-methods) 
5) Save the File to disk using the "saveFile" metod.
```
egm.saveFile("File path","file_name");
```

NOTE: You can find the full code sample here [SimpleTwoSheetWorkbook](https://github.com/SantoshVaramballi/EasyXL/blob/main/src/main/java/org/easyxl/sample/SimpleTwoSheetWorkbook.java)


# Primary Methods
<To be completed>













