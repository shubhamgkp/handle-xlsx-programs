package handleXLSX;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class XLSX_OutCrrspndnceRowCol {

	public static void main(String[] args) throws IOException{
			
		rowColOut(4, 2);
	}

	public static void rowColOut(int rowNum, int colNum) throws IOException{
		File file = new File("E:/STUDY_STUFF/SQA/SELENIUM_TESTING/Gurgaon_Class/Selenium/testdata.xlsx");
		FileInputStream fs = new FileInputStream(file);

		XSSFWorkbook wk =  new XSSFWorkbook(fs);
		XSSFSheet sheet = wk.getSheetAt(0);
		int numRow = sheet.getPhysicalNumberOfRows();

		XSSFRow row = sheet.getRow(rowNum);
		int numCell = row.getPhysicalNumberOfCells();

		XSSFCell cell = row.getCell(colNum);
		System.out.print(cell.getStringCellValue()+" ");
	}
}
