package nonused;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Case1 {
	public static void main(String[] args) {
		String excelPath = "C:\\Users\\User\\Desktop\\reading1.xlsx";
		System.out.println(excelPath);		
		File f = new File(excelPath);
		FileInputStream fip;
		XSSFWorkbook wb = null;
		try {
			fip = new FileInputStream(f);
			wb = new XSSFWorkbook(fip);
		} catch (IOException e) {
			e.printStackTrace();
		}
		LinkedHashMap<String,LinkedHashMap<String,String>> sheetData = new LinkedHashMap<String,LinkedHashMap<String,String>>(); 
		XSSFSheet sh = wb.getSheetAt(0);
		XSSFRow row = null;
		XSSFCell cell = null;
		int rowCount = sh.getPhysicalNumberOfRows();
		XSSFRow firstRow = sh.getRow(0);
		for(int i=1;i<rowCount;i++)  {			
			row = sh.getRow(i);
			String empno = String.valueOf(row.getCell(0).getNumericCellValue());
			LinkedHashMap<String , String> rowData = new LinkedHashMap<String , String>();
			int columnCount = row.getLastCellNum();
			for(int j=1;j<columnCount;j++)  {
				String key = firstRow.getCell(j).getStringCellValue();
				String value = null;
					cell=row.getCell(j);
				switch(cell.getCellType())  {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.print(cell.getNumericCellValue());
					value =  String.valueOf(cell.getNumericCellValue());
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.print(cell.getStringCellValue());
					value = cell.getStringCellValue();
					//System.out.println("\t");
				}
				System.out.print("\t");
           		rowData.put(key, value);	
			}
			System.out.println();
			sheetData.put(empno, rowData);
		}
		//System.out.println(sheetData);
		Case2 cs = new Case2();
		cs.readMyExcelData(sheetData);
		
	}

}
