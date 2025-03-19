package test;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class Excell {
	@Test
	public void getExcell() throws IOException {
		// Every row of excel shold be sent to 1 array
				FileInputStream fis = new FileInputStream("E:\\Projeler\\Testing\\ExcelDataProvider\\TestData.xlsx");
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				XSSFSheet sheet = workbook.getSheetAt(0);
				int rowCount = sheet.getPhysicalNumberOfRows(); // row count
				XSSFRow row = sheet.getRow(0);
				int columnCount = row.getLastCellNum(); // column count
				Object data[][] = new Object[rowCount - 1][columnCount];
				/*
				 * data[0][0]="Hallo"; data[0][1]="Hallo"; data[0][2]="Hallo";
				 */
				for (int i = 0; i < rowCount; i++) {  
					System.out.println("Loop started...");
				    row = sheet.getRow(i); // i-1 yerine direkt i kullan
				    if (row != null) {  // Satırın null olup olmadığını kontrol et
				        for (int j = 0; j < columnCount; j++) {  
				            System.out.println(row.getCell(j));  
				            data[i][j] = row.getCell(j);
				        }
				    } else {
				        System.out.println("Empty line skipped: " + i);
				    }
				    System.out.println("Loop Endeded...");
				}
	}
}
