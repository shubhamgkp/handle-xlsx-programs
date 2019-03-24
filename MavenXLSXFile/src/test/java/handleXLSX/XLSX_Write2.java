package handleXLSX;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSX_Write2 {

	public static void main(String[] args) throws IOException {

		File file = new File("./TestData/testdatawrite.xlsx");
		FileOutputStream fs = new FileOutputStream(file);				

		XSSFWorkbook wk = new XSSFWorkbook();
		//XSSFSheet sheet = wk.createSheet("skv");
		XSSFSheet sheet = wk.getSheet("skv");


		//sheet.getRow(0).getCell(0).setCellValue("Shubham");;
		sheet.createRow(2).createCell(0).setCellValue("Rajesh Yadav");
		
		//sheet.createRow(1).createCell(0).setCellValue("Rohit Pandey");
		/*sheet.createRow(2).createCell(0).setCellValue("Rajesh Yadav");
		 * sheet.createRow(3).createCell(0).setCellValue("Abhishek Roy");
		 * sheet.createRow(4).createCell(0).setCellValue("Ajay Yadav");
		 * sheet.createRow(5).createCell(0).setCellValue("Hamara Pandit");
		 */
		/*
		 * sheet.createRow(6).createCell(0).setCellValue("SP Sir");
		 * sheet.createRow(7).createCell(0).setCellValue("Neha Nehra");
		 * sheet.createRow(8).createCell(0).setCellValue("Ajay Choudhary");
		 */
		wk.write(fs);
		fs.flush();
		fs.close();
	}
}
