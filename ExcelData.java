package week5.day2.assessment;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelData extends LeadsBase{
	
	public static String[][] fetchCreate(String sheet1) throws IOException{
		XSSFWorkbook wBook = new XSSFWorkbook("./data/Lead.xlsx");
		XSSFSheet ws = wBook.getSheet(sheet1);
		int rows = ws.getLastRowNum();
		System.out.println(rows);
		short cells = ws.getRow(0).getLastCellNum();
		System.out.println(cells);		
		String[][] data = new String[rows][cells];
		for (int i = 1; i <= rows; i++) {
			XSSFRow row = ws.getRow(i);
			for (int j = 0; j < cells; j++) {
				XSSFCell cell = row.getCell(j);
				String cellValue = cell.getStringCellValue();
				System.out.println(cellValue);
				data[i-1][j] = cellValue;
			}
		}
		wBook.close();		
		return data;	
	}
		
	public static String[][] fetchEdit(String sheet2) throws IOException{
		
		XSSFWorkbook wBook = new XSSFWorkbook("./data/Lead.xlsx");
		XSSFSheet ws = wBook.getSheet(sheet2);
		int rows = ws.getLastRowNum();
		short cells = ws.getRow(0).getLastCellNum();
		String[][] data = new String[rows][cells];
		for(int i=1;i<=rows;i++)
		{
			XSSFRow row = ws.getRow(i);
			for(int j=0;j<cells;j++) {
				XSSFCell cell = row.getCell(j);
				String cellValue = cell.getStringCellValue();	
				System.out.println(cellValue);
				data[i-1][j] = cellValue;
			}
		}
		wBook.close();		
		return data;		
	}
	public static String[][] fetchMerge(String sheet3) throws IOException{
		
		XSSFWorkbook wBook = new XSSFWorkbook("./data/Lead.xlsx");
		XSSFSheet ws = wBook.getSheet(sheet3);
		int rows = ws.getLastRowNum();
		short cells = ws.getRow(0).getLastCellNum();
		String[][] data = new String[rows][cells];
		for(int i=1;i<=rows;i++)
		{
			XSSFRow row = ws.getRow(i);  
			
			
			for(int j=0;j<cells;j++) {
				XSSFCell cell = row.getCell(j);
				String cellValue = cell.getStringCellValue();	
				System.out.println(cellValue);
				data[i-1][j] = cellValue;
			}
		}
		wBook.close();		
		return data;		
	}
	public static String[][] fetchDuplicate(String sheet4) throws IOException{
		
		XSSFWorkbook wBook = new XSSFWorkbook("./data/Lead.xlsx");
		XSSFSheet ws = wBook.getSheet(sheet4);
		int rows = ws.getLastRowNum();
		short cells = ws.getRow(0).getLastCellNum();
		String[][] data = new String[rows][cells];
		for(int i=1;i<=rows;i++)
		{
			XSSFRow row = ws.getRow(i);
			for(int j=0;j<cells;j++) {
				XSSFCell cell = row.getCell(j);
				String cellValue = cell.getStringCellValue();
				System.out.println(cellValue);
				data[i-1][j] = cellValue;
			}
		}
		wBook.close();		
		return data;		
	}
	public static String[][] fetchDelete(String sheet5) throws IOException{
		
		XSSFWorkbook wBook = new XSSFWorkbook("./data/Lead.xlsx");
		XSSFSheet ws = wBook.getSheet(sheet5);
		int rows = ws.getLastRowNum();
		short cells = ws.getRow(0).getLastCellNum();
		String[][] data = new String[rows][cells];
		for(int i=1;i<=rows;i++)
		{
			XSSFRow row = ws.getRow(i);
			for(int j=0;j<cells;j++) {
				XSSFCell cell = row.getCell(j);
				String cellValue = cell.getStringCellValue();
				System.out.println(cellValue);
				data[i-1][j] = cellValue;
			}
		}	
		wBook.close();	
		return data;		
	}
}
