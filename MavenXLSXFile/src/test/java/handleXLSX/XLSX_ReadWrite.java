package handleXLSX;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class XLSX_ReadWrite {

	public static void main(String[] args) throws IOException {

		File file1 = new File("E:/STUDY_STUFF/SQA/SELENIUM_TESTING/Gurgaon_Class/Selenium/testdata.xlsx");
		FileInputStream fs = new FileInputStream(file1);
		
		File file2 = new File("E:/STUDY_STUFF/SQA/SELENIUM_TESTING/Gurgaon_Class/Selenium/testdatawrite.xlsx");
		FileOutputStream fo = new FileOutputStream(file2);
				
		XSSFWorkbook wk1 =  new XSSFWorkbook(fs);
		XSSFSheet sheet1 = wk1.getSheetAt(0);
		int numRow = sheet1.getPhysicalNumberOfRows();
		
		XSSFWorkbook wk2 = new XSSFWorkbook();
		XSSFSheet sheet2 = wk2.createSheet("skv");

		for(int n=0;n<numRow;n++){

			XSSFRow row1 = sheet1.getRow(n);
			int numCell = row1.getPhysicalNumberOfCells();
			
			XSSFRow row2 = sheet2.createRow(n);
			for(int m=0;m<numCell;m++){
				
				XSSFCell cell1 = row1.getCell(m);
				String cell1Val = cell1.getStringCellValue();
				
				XSSFCell cell2 = row2.createCell(m);
				cell2.setCellValue(cell1Val);
			}			
		}
		wk2.write(fo);
		fo.flush();
		fo.close();
	}
}
