package test;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataProvide {
    DataFormatter formatter = new DataFormatter();

    @Test(dataProvider = "driveTest")
    public void testCaseData(String greeting, String communication, String id) {
        System.out.println(greeting + " " + communication + " " + id);
    }

    @DataProvider(name = "driveTest")
    public Object[][] getData() throws IOException {
        FileInputStream fis = new FileInputStream("E:\\Projeler\\Testing\\ExcelDataProvider\\TestData.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rowCount = sheet.getPhysicalNumberOfRows(); // toplam satır sayısı
        XSSFRow row = sheet.getRow(0);
        int columnCount = row.getLastCellNum(); // toplam sütun sayısı

        Object[][] data = new Object[rowCount - 1][columnCount]; // Başlık satırını atlıyoruz

        for (int i = 1; i < rowCount; i++) { // 1'den başlat, çünkü ilk satır başlık
            row = sheet.getRow(i);
            if (row == null) continue; // Boş satır varsa atla
            
            for (int j = 0; j < columnCount; j++) {
                XSSFCell cell = row.getCell(j);
                data[i - 1][j] = (cell == null) ? "" : formatter.formatCellValue(cell);
            }
        }

        workbook.close();
        fis.close();

        return data;
    }
}
