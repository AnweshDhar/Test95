package test.Testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPrac12 {
	static FileInputStream inpstr = null;
	static XSSFWorkbook workbook = null;
	
	public static String[] getRowData(String filepath, String sheetname, int colIndex, String TCName)
	{
		boolean bTag = false;
		String[] arr = null ;//= new String[10];
		try{					
			File f1 = new File(filepath);
			inpstr = new FileInputStream(f1);
			workbook = new XSSFWorkbook(inpstr);
								
			XSSFSheet ws1 = workbook.getSheet(sheetname);	
			int rowcount = ws1.getLastRowNum()+1;
			for(int i=1; i<rowcount; i++)
			{
				Row r1 = ws1.getRow(i);
				String exlTCName = r1.getCell(0).getStringCellValue();					
				if(TCName.trim().equals(exlTCName))						
				{
					int colCount = r1.getLastCellNum();
					System.out.println("column number: " + colCount);
					arr = new String[colCount];
					for(int j=colIndex; j<colCount; j++)
					{
						Cell cell = r1.getCell(j);
						if(cell.getCellType()==Cell.CELL_TYPE_STRING) {
							arr[j-colIndex]  = cell.getStringCellValue();
						}else if (cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA) {
							arr[j-colIndex]  = (String.valueOf(cell.getNumericCellValue()));
						//arr[j-colIndex]  = r1.getCell(j).getStringCellValue();
					}
					}
					
												
					bTag = true;
					break;
				}						
			}
			
			if(bTag==false)
			{
				System.out.println("Test case Name not found in test data sheet");
			}
		}
		catch(IOException e)	{
			System.out.println("File not found or unable to read/write data..");
		}
		
		catch(Exception e)	{
			System.out.println("unknown exception");
		}	
		
		return arr;
	}
	
	public static void main(String[] args) {
		
String[] strRowData = getRowData("D:\\testing masters\\Automation_ORS\\TestData\\TestDataExcel.xlsx", "TestData", 1,"Tc_04");
		
		System.out.println(strRowData[0]);
		System.out.println(strRowData[1]);
		System.out.println(strRowData[2]);
		System.out.println(strRowData[3]);
		System.out.println(strRowData[4]);
		System.out.println(strRowData[5]);
		
	}


}
