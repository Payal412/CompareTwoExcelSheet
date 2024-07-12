package javaPackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Compare_Two_ExcelSheet {

    public static void main(String[] args) throws IOException 
    {
        // Set the path to the ChromeDriver
        System.setProperty("webdriver.chrome.driver", "E:\\chromedriver_win32\\chromedriver.exe");

        // Initialize the WebDriver
        WebDriver driver = new ChromeDriver();

        // Specify the location of Excel files
        File src1 = new File("E:\\Two Excel sheet\\Sheet1.xlsx");
        File src2 = new File("E:\\Two Excel sheet\\Sheet2.xlsx");

        // Load File
        FileInputStream fis1 = new FileInputStream(src1);
        FileInputStream fis2 = new FileInputStream(src2);

        Workbook wb1 = WorkbookFactory.create(fis1);
        Sheet sheet1 = wb1.getSheetAt(0);

        Workbook wb2 = WorkbookFactory.create(fis2);
        Sheet sheet2 = wb2.getSheetAt(0);

        compareSheets(sheet1, sheet2);

        // Close the WebDriver
        driver.quit();

        // Close the file input streams
        fis1.close();
        fis2.close();
    }

    public static void compareSheets(Sheet sheet1, Sheet sheet2) 
    {
        int firstRow1 = sheet1.getFirstRowNum();
        int lastRow1 = sheet1.getLastRowNum();

        int firstRow2 = sheet2.getFirstRowNum();
        int lastRow2 = sheet2.getLastRowNum();

        for (int i = firstRow1; i <= lastRow1; i++) 
        {
            Row row1 = sheet1.getRow(i);
            Row row2 = sheet2.getRow(i);

            //
            if (row1 == null && row2 == null) {
                continue;
            }
            
            if ((row1 == null && row2 != null) || (row1 != null && row2 == null)) 
            {
                System.out.println("Row " + i + " is different.");
                continue;
            }

            int firstCell1 = row1.getFirstCellNum();
            int lastCell1 = row1.getLastCellNum();

            int firstCell2 = row2.getFirstCellNum();
            int lastCell2 = row2.getLastCellNum();

            for (int j = firstCell1; j <= lastCell1; j++) 
            {
                Cell cell1 = row1.getCell(j);
                Cell cell2 = row2.getCell(j);

                if (cell1 == null && cell2 == null) 
                {
                    continue;
                }
                
                if ((cell1 == null && cell2 != null) || (cell1 != null && cell2 == null)) 
                {
                    System.out.println("Cell (" + i + ", " + j + ") is different.");
                    continue;
                }

                String value1 = cell1.toString();
                String value2 = cell2.toString();

                if (!value1.equals(value2)) 
                {
                    System.out.println("Cell (" + i + ", " + j + ") is different. Value in sheet1: " + value1 + ", Value in sheet2: " + value2);
                }
           }
       }
   }
}
