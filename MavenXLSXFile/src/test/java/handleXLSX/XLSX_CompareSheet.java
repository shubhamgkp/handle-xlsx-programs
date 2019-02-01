package handleXLSX;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSX_CompareSheet {

	static File file1,file2;
	static FileInputStream fs1,fs2;
	static XSSFWorkbook wb1,wb2;
	static XSSFSheet sheet1,sheet2;
	static XSSFRow row1,row2;
	static XSSFCell cell1,cell2;

	public static void main(String[] args) throws IOException {

		file1 = new File("E:/STUDY_STUFF/SQA/SELENIUM_TESTING/Gurgaon_Class/Selenium/firstxl.xlsx");
		fs1 = new FileInputStream(file1);

		file2 = new File("E:/STUDY_STUFF/SQA/SELENIUM_TESTING/Gurgaon_Class/Selenium/secondxl.xlsx");
		fs2 = new FileInputStream(file2);

		wb1 = new XSSFWorkbook(fs1);
		wb2 = new XSSFWorkbook(fs2);

		sheet1=wb1.getSheet("Sheet1");
		sheet2=wb2.getSheet("Sheet1");

		int numRow1=sheet1.getPhysicalNumberOfRows();
		int numRow2=sheet2.getPhysicalNumberOfRows();

		if(numRow1==numRow2){
			for(int i=0;i<numRow1;i++){

				for(int j=0;j<2;j++){

					row1=sheet1.getRow(i);
					cell1=row1.getCell(j);
					String cellVal1=cell1.getStringCellValue();

					row2=sheet2.getRow(i);
					cell2=row2.getCell(j);
					String cellVal2=cell2.getStringCellValue();

					if(cellVal1.equals(cellVal2)){

						System.out.println(cellVal1+" "+cellVal2+" is equal!");
					}
					else{
						System.out.println(cellVal1+" "+cellVal2+" is not equal!");
					}
				}
			}			
		}

		else{
			System.out.println("Sheet rows are not equal!");
		}
	}
}
