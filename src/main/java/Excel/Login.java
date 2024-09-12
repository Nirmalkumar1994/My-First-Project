package Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Login {
	public static  void main (String[]args) throws IOException {
		
		// Path of Excel
	File file =	new File("E:\\Nirmal Java\\Eclipse Programs\\Excel\\Excel\\Book1.xlsx");
		// Read the object from file - FileInputStream Class		
		FileInputStream stream  = new FileInputStream(file);
	
	  //		Mention the workbook --- Collection of sheet
	Workbook workbook = new XSSFWorkbook(stream);
	
	 //  sheet name
	Sheet sheet = workbook.getSheet("Login");
	
	//Get row Detail
	Row row = sheet.getRow(3);
	
	// Get Cell
	
	Cell cell = row.getCell(2);
	System.out.println(cell);
	
	}

	
}
