package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class UploadDownload {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String fruitName = "Apple";
		String updatedValue = "1230";
		String fileName = "C:\\\\Users\\\\albat\\\\Downloads\\\\download.xlsx";
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));
		driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");
		// Download
		driver.findElement(By.cssSelector("#downloadButton")).click();
		// Edit Excel File -getColumnNumber of Price- getRownumber of Apple -> update
		// excel with row, column
		int col = getColumnNumber(fileName, "price");
		int row = getRowNumber(fileName, "Apple");
		Assert.assertTrue(updateCell(fileName, row, col, "1250"));

		// Upload
		WebElement upload = driver.findElement(By.xpath("//*[@id=\"fileinput\"]"));
		upload.sendKeys("C:\\Users\\albat\\Downloads\\download.xlsx"); // C:\Users\albat\Downloads\download.xlsx

		// Wait for success message to show up and wait for it to disappear
		By toastLocator = By.cssSelector(".Toastify__toast-body div:nth-child(2)"); // Toast message
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		wait.until(ExpectedConditions.visibilityOfElementLocated(toastLocator));
		String toastMessage = driver.findElement(toastLocator).getText();
		System.out.println(toastMessage);
		Assert.assertEquals("Updated Excel Data Successfully", toastMessage);

		wait.until(ExpectedConditions.invisibilityOfElementLocated(toastLocator));

		// Verify updated excel data showing in the web table
		String priceColumn = driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id"); // //*[@id="root"]/div[1]/div[1]/div/div/div[1]/div/div[4]/div/div
		String actualPrice = driver.findElement(By.xpath(
				"//div[text()='" + fruitName + "']/parent::div/parent[@id='cell-" + priceColumn + "-undefined']"))
				.getText();
		System.out.println(actualPrice);
		Assert.assertEquals("900", actualPrice);
	}

	private static boolean updateCell(String fileName, int row, int col, String updatedValue) throws IOException {
		FileInputStream fis = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Sheet1");

		Row rowField = sheet.getRow(row);
		if (rowField == null) {
			rowField = sheet.createRow(row);
		}

		Cell cellField = rowField.getCell(col);
		if (cellField == null) {
			cellField = rowField.createCell(col);
		}

		cellField.setCellValue(updatedValue);

		// ÇIKIŞ AKIŞINI DOĞRU KULLANMAK
		fis.close(); // İlk olarak input stream'i kapatıyoruz
		FileOutputStream fos = new FileOutputStream(fileName);
		workbook.write(fos); // Çıktıyı yazdır
		fos.close(); // Çıkış akışını kapat
		workbook.close(); // Workbook'u kapat

		return true;
	}

	private static int getRowNumber(String fileName, String rowName) throws IOException {
		FileInputStream fis = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		try {
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			Iterator<Row> rows = sheet.iterator();

			int rowIndex = 0;

			while (rows.hasNext()) {
				Row row = rows.next();
				Iterator<Cell> cells = row.cellIterator();

				while (cells.hasNext()) {
					Cell cell = cells.next();
					if (cell.getStringCellValue().equalsIgnoreCase(rowName)) {
						workbook.close();
						fis.close();
						return rowIndex; // Doğru satırı bulduğunda döndür
					}
				}
				rowIndex++;
			}

			return -1; // Eğer satır bulunamazsa -1 döndür
		} finally {
			workbook.close();
			fis.close();
		}
	}

	private static int getColumnNumber(String fileName, String colName) throws IOException {
		FileInputStream fis = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		try {
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			Row firstRow = sheet.getRow(0); // İlk satırı doğrudan al

			if (firstRow == null) {
				return -1; // Eğer ilk satır yoksa -1 döndür
			}

			int column = -1;
			for (Cell cell : firstRow) {
				if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(colName)) {
					column = cell.getColumnIndex();
					break;
				}
			}

			return column;
		} finally {
			workbook.close();
			fis.close();
		}
	}

	private static String getLatestDownloadedFile(String folderPath) {
		File folder = new File(folderPath);
		File[] files = folder.listFiles((dir, name) -> name.endsWith(".xlsx"));

		if (files == null || files.length == 0)
			return null;

		File latestFile = files[0];
		for (File file : files) {
			if (file.lastModified() > latestFile.lastModified()) {
				latestFile = file;
			}
		}
		return latestFile.getAbsolutePath();
	}

}
