package excelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	static String filePath = "C://numbers.xlsx";
	
	public static String getString(){
		String result = "";
		File myFile = new File(filePath);
	    FileInputStream fis = null;
		try {
			fis = new FileInputStream(myFile);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	
	    // Finds the workbook instance for XLSX file
	    XSSFWorkbook myWorkBook = null;
		try {
			myWorkBook = new XSSFWorkbook (fis);
		} catch (IOException e) {
			e.printStackTrace();
		}
	   
	    // Return first sheet from the XLSX workbook
	    XSSFSheet mySheet = myWorkBook.getSheetAt(0);
	   
	    // Get iterator to all the rows in current sheet
	    Iterator<Row> rowIterator = mySheet.iterator();
	   
	    // Traversing over each row of XLSX file
	    while (rowIterator.hasNext()) {
	        Row row = rowIterator.next();
	
	        // For each row, iterate through each columns
	        Iterator<Cell> cellIterator = row.cellIterator();
	        while (cellIterator.hasNext()) {
	
	            Cell cell = cellIterator.next();
	
	            if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
	                int num = (int) cell.getNumericCellValue();
	                result += num;
	            }
	            else if (cell.getCellType() == Cell.CELL_TYPE_STRING){
	            	String string = cell.getStringCellValue();
	            	Scanner in = new Scanner(string).useDelimiter("[^0-9]+");
	            	int integer = in.nextInt();
	            	result += integer;
	            }
	        }
	    }
	    return result;
	}
	
	public static void main(String[] args) {
		String numbers = ExcelReader.getString();
		System.out.println(numbers);
	}
}
