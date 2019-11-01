package writer;

//import com.sun.java.util.jar.pack.Package;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/*
https://howtodoinjava.com/library/readingwriting-excel-files-in-java-poi-tutorial/

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
 */

public class XlsWriterPracticeApp {

    public static void main( String[] args ) throws IOException {


        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
        data.put("2", new Object[] {1, "Amit", "Shukla"});
        data.put("3", new Object[] {2, "Lokesh", "Gupta"});
        data.put("4", new Object[] {3, "John", "Adwards"});
        data.put("5", new Object[] {4, "Brian", "Schultz"});

        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
                org.apache.poi.ss.usermodel.Cell cell = row.createCell(cellnum++);
                if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("howtodoinjava_demo.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }


        //****************************************************************

//       File myFile = new File("C://temp/Employee.xlsx");
//        FileInputStream fis = null;
//        try {
//            fis = new FileInputStream(myFile);
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        }
//
//        // Finds the workbook instance for XLSX file
//        XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
//
//        // Return first sheet from the XLSX workbook
//        XSSFSheet mySheet = myWorkBook.getSheetAt(0);
//
//        // Get iterator to all the rows in current sheet
//        Iterator<Row> rowIterator = mySheet.iterator();
//
//        // Traversing over each row of XLSX file
//        while (rowIterator.hasNext()) {
//            org.apache.poi.ss.usermodel.Row row = rowIterator.next();
//
//            // For each row, iterate through each columns
//            Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//
//                org.apache.poi.ss.usermodel.Cell cell = cellIterator.next();
//
//                switch (cell.getCellType()) {
//                    case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING:
//                        System.out.print(cell.getStringCellValue() + "\t");
//                        break;
//                    case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC:
//                        System.out.print(cell.getNumericCellValue() + "\t");
//                        break;
//                    case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BOOLEAN:
//                        System.out.print(cell.getBooleanCellValue() + "\t");
//                        break;
//                    case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA:
//                        System.out.println("this option might not work...");
//                        System.out.println(cell.getCellFormula() + "\t");
//                        break;
//                    default :
//
//                }
//            }
//            System.out.println("");
//        }

//*******************************************************************************
















    }// end main
    //**************************************888
}//end class
