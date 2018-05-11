package nonused;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Case5 {

	public static void main(String[] args) throws IOException {
		LinkedHashMap<String,LinkedHashMap<String,String>> entire = new LinkedHashMap<String,LinkedHashMap<String,String>>(); 
		String excelPath = "C:\\Users\\User\\Desktop\\reading1.xlsx";
		File f = new File(excelPath);
		FileInputStream fip = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fip);
		XSSFSheet sh = wb.getSheetAt(0);
		Row row = null;
		Cell cell = null;
		int colC = sh.getRow(1).getLastCellNum();
		int rowC = sh.getLastRowNum();
		XSSFRow firstRow = sh.getRow(0);
		for(int i=0;i<rowC;i++)  {
			row = sh.getRow(i);
			String empno = String.valueOf(row.getCell(0).getNumericCellValue());
			LinkedHashMap<String,String> rowData = new LinkedHashMap<String,String>(); 
			for(int j=0;j<colC;j++)  {
				cell = row.getCell(j);
				String key = firstRow.getCell(j).getStringCellValue();
				String value = null;
				switch(cell.getCellType())  {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.print(cell.getNumericCellValue());
					value = String.valueOf(cell.getNumericCellValue());
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.print(cell.getStringCellValue());
					value = String.valueOf(cell.getStringCellValue());
					break;
				}
				System.out.print("\t");
				rowData.put(key, value);
			}
			System.out.println();
			entire.put(empno, rowData);
			

	}
		Case51 cas = new Case51();
		cas.readData(entire);

}
}
