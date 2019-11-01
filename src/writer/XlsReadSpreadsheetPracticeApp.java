package writer;



/*
https://howtodoinjava.com/library/readingwriting-excel-files-in-java-poi-tutorial/
https://www.apache.org/dyn/closer.lua/poi/release/src/poi-src-4.1.1-20191023.zip
https://mvnrepository.com/artifact/org.apache.poi/poi/3.11-beta2
Required Files to run this code:
1:
dom4j-1.6.1.jar
https://mvnrepository.com/artifact/dom4j/dom4j/1.6.1

2:
poi-3.9-20121203.jar
https://mvnrepository.com/artifact/org.apache.poi/poi/3.9

3:
poi-ooxml-3.9-20121203.jar
https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml/3.9

4:
poi-ooxml-schemas-3.9-20121203.jar
https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml-schemas/3.9

5:
xmlbeans-2.3.0.jar
https://mvnrepository.com/artifact/org.apache.xmlbeans/xmlbeans/2.3.0

Add by going to: Files>Project structure>Libraries>(plus sign)>Java>(navigate to folders & select)

Reading an excel file using POI is also very simple if we divide this in steps.

Create workbook instance from excel sheet
Get to the desired sheet
Increment row number
iterate over all cells in a row
repeat step 3 and 4 until all data is read
 */

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class XlsReadSpreadsheetPracticeApp {

    public static void main(String[] args)
    {
        try
        {
            FileInputStream file = new FileInputStream(new File("SalesDashboard-9-20-19.xlsx"));

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext())
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "t");
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }



}//end class









