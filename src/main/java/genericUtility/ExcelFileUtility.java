package genericUtility;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtility {
	public String readDataFromExcelFile(String sheetNum, int rowNum, int cellNum) throws Throwable {
		FileInputStream fileInputStream = new FileInputStream("./CommonData/TestData.xlsx");
	    Workbook workBook = WorkbookFactory.create(fileInputStream);
	    Sheet sheet = workBook.getSheet(sheetNum);
	    Row row = sheet.getRow(rowNum);
	    Cell cell = row.getCell(cellNum);
	    String value = cell.getStringCellValue();
		return value;
	}
	public int lastRowNum(String sheetNum) throws Throwable {
		FileInputStream fileInputStream = new FileInputStream("./CommonData/TestData.xlsx");
		Workbook workBook = WorkbookFactory.create(fileInputStream);
		return workBook.getSheet(sheetNum).getLastRowNum();
	}
}