package com.excelrowcol;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class Excel_Reader {

	public String path;
	FileInputStream fis;
	HSSFWorkbook workbook;
	HSSFSheet sheet;

	public Excel_Reader(String path) {
		this.path = path;
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public String getExcelData(String sheetName,String colName,int rowNo) {

		try{
		int col_Num = 0;
	//	int index = workbook.getSheetIndex(sheetName);
		//sheet = workbook.getSheetAt(index);
		 sheet = workbook.getSheet(sheetName);
		HSSFRow row = sheet.getRow(0);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			if (row.getCell(i).getStringCellValue().equals(colName)) {
				col_Num = i;
			}

		}
		// to get row no
		row = sheet.getRow(rowNo - 1);
		HSSFCell cell = row.getCell(col_Num);
		
		if(cell.getCellType()== Cell.CELL_TYPE_STRING){
			return cell.getStringCellValue();
		}
		else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
			return String.valueOf(cell.getNumericCellValue());
		}
		
		else if(cell.getCellType()==Cell.CELL_TYPE_BOOLEAN){
			return String.valueOf(cell.getBooleanCellValue());
					
		}
		else if(cell.getCellType()==Cell.CELL_TYPE_BLANK){
			return "";
		}
		}
		catch(Exception e){
			e.printStackTrace();
		}

		return null;

	}

	
	public static void main(String[] args) {
		
		
		Excel_Reader excel_Reader=new Excel_Reader("/src/com/excelrowcol/Auto.xls");
		System.out.println((excel_Reader.getExcelData("Auto","weight",3)));
		
	}
}
