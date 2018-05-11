package nonused;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Case4 {

	public static void main(String[] args) throws IOException {
		ArrayList<ArrayList> entire = new ArrayList<ArrayList>(); 
		String excelPath = "C:\\Users\\User\\Desktop\\reading1.xlsx";
		File f = new File(excelPath);
		FileInputStream fip = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fip);
		XSSFSheet sh = wb.getSheetAt(0);
		int rowC = sh.getLastRowNum();
		int colC = sh.getRow(1).getLastCellNum();
		XSSFRow row = null;
		XSSFCell cell = null;
		for(int i=0;i<rowC;i++)   {
			row = sh.getRow(i);
			ArrayList list = new ArrayList<>();
			for(int j =0;j<colC;j++)  {
				cell = row.getCell(j);
				switch(cell.getCellType()){
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue());
						list.add(cell.getStringCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue());
						list.add(cell.getNumericCellValue());
						break;
			}
				System.out.print("\t");
		}	
			System.out.println();
			entire.add(list);

		}
		System.out.println(entire);
		}
	}


