package handleXLSX;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSX_Write {

	public static void main(String[] args) throws IOException {

		File file = new File("./TestData/testdatawrite.xlsx");
		FileOutputStream fs = new FileOutputStream(file);				

		XSSFWorkbook wk = new XSSFWorkbook();
		XSSFSheet sheet = wk.createSheet("skv");
		
		for(int i=0;i<3;i++){
			
			XSSFRow row = sheet.createRow(i);
			for(int j=0;j<2;j++){
				
				XSSFCell cell1 = row.createCell(j);
				cell1.setCellValue("Shubham");
			}			
		}
		
		wk.write(fs);
		fs.flush();
		fs.close();
	}
}
