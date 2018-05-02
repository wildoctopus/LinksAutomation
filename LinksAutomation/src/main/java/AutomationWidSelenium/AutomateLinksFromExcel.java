package AutomationWidSelenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;

public class AutomateLinksFromExcel {
	
	public static WebDriver createInstance(){
        DesiredCapabilities capabilities = DesiredCapabilities.chrome();
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--incognito");
        capabilities.setCapability(ChromeOptions.CAPABILITY, options);
        System.setProperty("webdriver.chrome.driver","E:\\Softwares\\chromedriver_win32\\chromedriver.exe");
        WebDriver driver = new ChromeDriver(capabilities);
        
        return driver;
    }

	public void readExcel(String filePath, String fileName, String sheetName) throws IOException {

		/*System.setProperty("webdriver.chrome.driver", "E:\\Softwares\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		*/
		WebDriver driver = createInstance();;

		//driver.manage().window().maximize();

		driver.manage().timeouts().implicitlyWait(6, TimeUnit.SECONDS);
		
		// Create an object of File class to open xlsx file

		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file

		FileInputStream inputStream = new FileInputStream(file);

		Workbook hrefsWorkbook = null;

		// Find the file extension by splitting file name in substring and getting only
		// extension name

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file

		if (fileExtensionName.equals(".xlsx")) {

			// If it is xlsx file then create object of XSSFWorkbook class

			hrefsWorkbook = new XSSFWorkbook(inputStream);

		}

		// Check condition if the file is xls file

		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of XSSFWorkbook class

			hrefsWorkbook = new HSSFWorkbook(inputStream);

		}

		// Read sheet inside the workbook by its name

		Sheet hrefsSheet = hrefsWorkbook.getSheet(sheetName);

		// Find number of rows in excel file

		int rowCount = hrefsSheet.getLastRowNum() - hrefsSheet.getFirstRowNum();

		// Create a loop over all the rows of excel file to read it

		for (int i = 0; i < rowCount + 1; i++) {

			Row row = hrefsSheet.getRow(i);

			// Create a loop to fetch cell values in a row

			for (int j = 0; j < row.getLastCellNum(); j++) {

				// navigate to each value

				System.out.print(row.getCell(j).getStringCellValue());
				driver.navigate().to(row.getCell(j).getStringCellValue());

			}

			//System.out.println();

		}
		
		driver.close();

	}

	

	

}