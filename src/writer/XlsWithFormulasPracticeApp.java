package writer;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
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

public class XlsWithFormulasPracticeApp {
    public static void main(String[] args)
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Calculate Simple Interest");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Pricipal");
        header.createCell(1).setCellValue("RoI");
        header.createCell(2).setCellValue("T");
        header.createCell(3).setCellValue("Interest (P r t)");

        Row dataRow = sheet.createRow(1);
        dataRow.createCell(0).setCellValue(14500d);
        dataRow.createCell(1).setCellValue(9.25);
        dataRow.createCell(2).setCellValue(3d);
        dataRow.createCell(3).setCellFormula("A2*B2*C2");

        try {
            FileOutputStream out =  new FileOutputStream(new File("formulaDemo.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Excel with foumula cells written successfully");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }




}
