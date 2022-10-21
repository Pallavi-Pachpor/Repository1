package xmlread;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;
public class ExcelReading {

	@Test
	public void Readsheetdetails() throws IOException {
//		create instance of fileinputstream class and path the required exel file location to its constructor
		FileInputStream fis = new FileInputStream(".\\src\\test\\resources\\23rdjune.xlsx");
//		Create an instance of workbook class and path FileInputStream
//		Instance to its constructor
		Workbook workbook = new XSSFWorkbook(fis);
//		get the number of sheet present in excel 
		int sheetcount = workbook.getNumberOfSheets();
		System.out.println("sheet count ="+sheetcount);
//		print all the sheet name from excel
		for(int i=0;i<sheetcount;i++) {
			String sheetname = workbook.getSheetName(i);
			System.out.println("sheet name is: "+sheetname);
		}
//		get required sheet from excel
		Sheet sheet = workbook.getSheet("Sheet1");
	}
	@Test
	public void getrowdetails() throws IOException {
//		create instance of fileinputstream class and path the required exel file location to its constructor
		FileInputStream fis = new FileInputStream("C:\\Users\\pachp\\Desktop\\newworkspace\\Maven1\\src\\test\\resources\\23rdjune.xlsx");
//		Create an instance of workbook class and path FileInputStream
//		Instance to its constructor
		Workbook workbook = new XSSFWorkbook(fis);
//		get required sheet from excel
		Sheet sheet = workbook.getSheet("Sheet1");
//		get the required sheet present in excel 
		int rowcount = sheet.getLastRowNum();
		System.out.println("row count: "+rowcount);
//		get speciic row from sheet 
		Row row = sheet.getRow(1);
	}
}

