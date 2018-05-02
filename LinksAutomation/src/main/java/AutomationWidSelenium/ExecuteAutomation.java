package AutomationWidSelenium;

import java.io.IOException;

public class ExecuteAutomation {

public static void main(String[] strings) throws IOException {

		

		// Create an object of AutomateLinksFromExcel class

		AutomateLinksFromExcel objExcelFile = new AutomateLinksFromExcel();

		// Prepare the path of excel file

		String filePath = "E:\\selenium";

		// Call read file method of the class to read data
		
		for (int i = 0; i<400; i++)
		{
			objExcelFile.readExcel(filePath, "LinksExcel.xlsx", "ExcelLinksSheet");
			
		}

		//objExcelFile.readExcel(filePath, "LinksExcel.xlsx", "ExcelLinksSheet");

	}

}
