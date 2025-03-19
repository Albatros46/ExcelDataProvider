package test;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;


public class UploadDownload {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String fruitName = "Apple";
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));
		driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");
		//Download
		driver.findElement(By.cssSelector("#downloadButton")).click();
		//Edit Excel File
		
		//Upload
		WebElement upload= driver.findElement(By.xpath("//*[@id=\"fileinput\"]"));
		upload.sendKeys("C:\\Users\\albat\\Downloads\\download.xlsx"); //C:\Users\albat\Downloads\download.xlsx
		
		
		//Wait for success message to show up and wait for it to disappear
		By toastLocator = By.cssSelector(".Toastify__toast-body div:nth-child(2"); //Toast message
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		wait.until(ExpectedConditions.visibilityOfElementLocated(toastLocator));
		String toastMessage = driver.findElement(toastLocator).getText();
		System.out.println(toastMessage);
		Assert.assertEquals("Updated Excel Data Successfully", toastMessage);
		
		wait.until(ExpectedConditions.invisibilityOfElementLocated(toastLocator));
		
		//Verify updated excel data showing in the web table
		String priceColumn=driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id"); //  //*[@id="root"]/div[1]/div[1]/div/div/div[1]/div/div[4]/div/div
		String actualPrice=driver.findElement(By.xpath("//div[text()='"+fruitName+"']/parent::div/parent[@id='cell-"+priceColumn+"-undefined']")).getText();
		System.out.println(actualPrice);
		Assert.assertEquals("900", actualPrice);
	}

}
