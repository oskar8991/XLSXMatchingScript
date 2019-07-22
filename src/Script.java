import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


// goes into xlsx 1 and check for key cell in first row, then extracts wanted cells
// cycles through all key cells in xlsx 2 looking for same String (label), then extracts wanted cells
// combines the key cell and wanted cells into a new xlsx document as output


public class Script {
	
	public static void main(String[] args) throws InvalidFormatException, FileNotFoundException, IOException {
		
		int rowCounter3 = 0;
		
		// input xlsx file 1
		XSSFWorkbook workbook1 = new XSSFWorkbook(new FileInputStream("1coordinatesTest.xlsx"));
		XSSFSheet sheet1 = workbook1.getSheetAt(0);
		XSSFRow row1 = sheet1.getRow(0);;

		// input xlsx file 2
		XSSFWorkbook workbook2 = new XSSFWorkbook(new FileInputStream("2valuesTest.xlsx"));
		XSSFSheet sheet2 = workbook2.getSheetAt(0);
		XSSFRow row2 = sheet2.getRow(1);
		
		// output xlsx file
		XSSFWorkbook workbook3 = new XSSFWorkbook();
		XSSFSheet sheet3 = workbook3.createSheet("sheet");
		
		
		for(int i = 0; i <= sheet1.getLastRowNum(); i++) {
			XSSFRow currentRow1 = sheet1.getRow(i);
			XSSFCell currentFirstCell1 = currentRow1.getCell(0);
			XSSFCell currentSecondCell1 = currentRow1.getCell(1);
			String currentFirstCellString1 = currentFirstCell1.toString();
			String currentSecondCellSring1 = currentSecondCell1.toString();

			
			for(int x = 0; x <= sheet2.getLastRowNum(); x++) {
				XSSFRow currentRow2 = sheet2.getRow(x);
				XSSFCell currentFirstCell2 = currentRow2.getCell(0);
				XSSFCell currentSecondCell2 = currentRow2.getCell(1);
				
				if(currentFirstCell2 != null) {
					String currentFirstCellString2 = currentFirstCell2.toString();
					String currentSecondCellString2 = currentSecondCell2.toString();
					
					if(currentFirstCellString1.equals(currentFirstCellString2)) {
						XSSFRow row3 = sheet3.createRow(rowCounter3);
						XSSFCell cellOne3 = row3.createCell(0);
						XSSFCell cellTwo3 = row3.createCell(1);
						XSSFCell cellThree3 = row3.createCell(2);
						cellOne3.setCellValue(currentFirstCellString1);
						cellTwo3.setCellValue(currentSecondCellSring1);
						cellThree3.setCellValue(currentSecondCellString2);
						workbook3.write(new FileOutputStream("output.xlsx"));
						rowCounter3++;
					}
				}
				
			}
			
		}
		
		workbook3.close();
		
	}

}
