package handleXLSX;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class XLSX_Read {

	public static void main(String[] args) throws IOException {

		File file = new File("E:/STUDY_STUFF/SQA/SELENIUM_TESTING/Gurgaon_Class/Selenium/testdata.xlsx");
		FileInputStream fs = new FileInputStream(file);

		XSSFWorkbook wk =  new XSSFWorkbook(fs);
		XSSFSheet sheet = wk.getSheetAt(0);
		int numRow = sheet.getPhysicalNumberOfRows();

		for(int n=0;n<numRow;n++){

			XSSFRow row = sheet.getRow(n);
			int numCell = row.getPhysicalNumberOfCells();
			for(int m=0;m<numCell;m++){

				XSSFCell cell = row.getCell(m);
				System.out.print(cell.getStringCellValue()+" ");
			}
			System.out.println();
		}
	}
}
