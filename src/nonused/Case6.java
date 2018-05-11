package nonused;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Case6 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//ArrayList<ArrayList> list = new ArrayList<ArrayList>();
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
			for(int j=0;j<colC;j++)  {
				cell = row.getCell(j);
				switch(cell.getCellType())  {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.print(cell.getStringCellValue());
					break;
				}
				System.out.print("\t");
			}
			System.out.println();
			

	}
	}

}
