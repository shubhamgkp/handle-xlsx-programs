package handleXLSX;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSX_UserInputWrite {

	public static void main(String[] args) throws IOException {

		File file = new File("E:/STUDY_STUFF/SQA/SELENIUM_TESTING/Gurgaon_Class/Selenium/testdatawrite.xlsx");
		FileOutputStream fs = new FileOutputStream(file);				

		System.out.println("Please enter the test data!");
		Scanner sc = new Scanner(System.in);
		
		XSSFWorkbook wk = new XSSFWorkbook();
		XSSFSheet sheet = wk.createSheet("skv");
		
		for(int i=0;i<2;i++){
			
			XSSFRow row = sheet.createRow(i);
			for(int j=0;j<2;j++){
				
				String inpdata = sc.nextLine();
				XSSFCell cell = row.createCell(j);
				cell.setCellValue(inpdata);
			}			
		}
		
		wk.write(fs);
		fs.flush();
		fs.close();
	}
}
