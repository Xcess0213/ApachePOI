import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class StudentsTest {

    @Test
    public void test() throws IOException {

        // Input Stream from your excel file
        FileInputStream studentsExcelFile = new FileInputStream("src/test/students.xlsx");

        // use input stream for creating workbook
        Workbook wb = new XSSFWorkbook((studentsExcelFile));

        // Get the sheet
        Sheet dataSheet = wb.getSheet("data");

        // Select row
        Row row = dataSheet.getRow(6);

        // select cell (column)
        Cell cell = row.getCell(1);

        System.out.println(cell.toString());
    }

    @Test
    public void test2() throws IOException {
        FileInputStream excelFile = new FileInputStream("src/test/students.xlsx");

        Workbook studentsExcel = new XSSFWorkbook(excelFile);

        Sheet dataSheet = studentsExcel.getSheet("data");

        Row row = dataSheet.getRow(20);

        for (int cell = 0; cell < 6; cell++) {

            Cell column = row.getCell(cell);

            System.out.println(column.toString());
        }
    }
    @Test
    public void test3() throws IOException {
        FileInputStream excelFile = new FileInputStream("src/test/students.xlsx");

        Workbook studentsExcel = new XSSFWorkbook(excelFile);

        Sheet dataSheet = studentsExcel.getSheet("data");

        Assert.assertNotEquals(dataSheet, null);

        Row row = dataSheet.getRow(34);

        for (int i = 0; i < 6; i++) {

            Cell column = row.getCell(i);

            System.out.println(column.toString());
        }
    }

    @Test
    public void  test4() throws IOException {

        FileInputStream excelFile = new FileInputStream("src/test/students.xlsx");

        Workbook studentsExcel = new XSSFWorkbook(excelFile);

        Sheet dataSheet = studentsExcel.getSheet("data");

        int firstRow = dataSheet.getFirstRowNum();
        int lastRow = dataSheet.getLastRowNum();

        int rowCount = lastRow - firstRow;
        for (int i = 0; i < rowCount; i++ ){

            Row currentRow = dataSheet.getRow(i);

            for (int j = 0; j < 6; j++) {

                Cell cell = currentRow.getCell(j);

                System.out.print(cell.toString() + " ");
            }
            System.out.println();

        }



    }

}
