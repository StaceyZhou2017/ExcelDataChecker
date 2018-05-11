package nonused;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Case3 {

	public static void main(String[] args) throws IOException {
		ArrayList<ArrayList> enlist = new ArrayList<ArrayList>();
		String excelPath = "C:\\Users\\User\\Desktop\\reading1.xlsx";
		File f = new File(excelPath);
		FileInputStream fip = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fip);
		XSSFSheet sh = wb.getSheetAt(0);
		int rowC = sh.getLastRowNum();
		Row row = null;
		Cell cell = null;
		int colC = sh.getRow(1).getLastCellNum();
		for(int i=0;i<rowC;i++)  {
			row = sh.getRow(i);
			ArrayList list1 = new ArrayList<>();
			for(int j=0;j<colC;j++)  {
				cell = row.getCell(j);
				switch(cell.getCellType())  {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.print(cell.getNumericCellValue());
					list1.add(cell.getNumericCellValue());
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.print(cell.getStringCellValue());
					list1.add(cell.getStringCellValue());
					break;
				}
				System.out.print("\t");
			}
			System.out.println();
			enlist.add(list1);
			
		}
		System.out.println(enlist);
		
			}

}
