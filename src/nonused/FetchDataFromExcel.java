package nonused;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FetchDataFromExcel {

	public static void main(String[] args) {
		
		boolean isFound = checkForDataPresent("C:\\Users\\admin\\Desktop\\config\\test.xlsx","raheem");
		System.out.println(isFound);
	}
	
	public static boolean checkForDataPresent(String excelPath, String item){
		boolean isFound = false;
		File f = new File(excelPath);
		Workbook book = null;
		try {
			 book = new XSSFWorkbook(new FileInputStream(f));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		outerLoop:
		for(int sheetNo=0;sheetNo<book.getNumberOfSheets();sheetNo++){
			Sheet currSheet = book.getSheetAt(sheetNo);
			String value = "";
			for(int rowNo=0;rowNo<currSheet.getPhysicalNumberOfRows();rowNo++){
				Row currRow = currSheet.getRow(rowNo);
				for(int cellNo=0;cellNo<currRow.getLastCellNum();cellNo++){
					Cell currCell = currRow.getCell(cellNo);
					
					switch(currCell.getCellTypeEnum()){
					
					case STRING:
						value = currCell.getStringCellValue();
						break;
					case NUMERIC:
						double oriValue = currCell.getNumericCellValue();
						long modifiedValue= Math.round(oriValue);
						if(oriValue == modifiedValue){
							value = String.valueOf(modifiedValue);
						}else{
							value = String.valueOf(oriValue);
						}
						break;
					case BOOLEAN:
						boolean bool = currCell.getBooleanCellValue();
						value = String.valueOf(bool);
						break;
					}
					
					if(item.equals(value)){
						isFound = true;
						break outerLoop;
					}
				}
			}
		}
		
		
		return isFound;
		
	}

}
