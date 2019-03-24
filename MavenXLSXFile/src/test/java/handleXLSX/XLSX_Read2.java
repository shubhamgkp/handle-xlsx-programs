package handleXLSX;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSX_Read2 {

	public static void main(String[] args) throws IOException {

		File file = new File("./TestData/writetestdata.xlsx");
		FileInputStream fs = new FileInputStream(file);

		XSSFWorkbook wk = new XSSFWorkbook(fs);
		XSSFSheet sheet = wk.getSheet("Sheet1");
		int numRow = sheet.getPhysicalNumberOfRows();

		System.out.println(numRow);

		XSSFRow row = sheet.createRow(2);
		row.createCell(0).setCellValue("Abhishek Bhai");
		row.createCell(1).setCellValue("Rajesh Yadav");
		fs.close();
		
		FileOutputStream fo = new FileOutputStream(file);
		
		wk.write(fo);
		fo.close();
		System.out.println(" is successfully written!");
		
	}
}