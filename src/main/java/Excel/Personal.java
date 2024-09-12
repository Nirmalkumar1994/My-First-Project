package Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.values.XmlValueOutOfRangeException;

public class Personal {
	public static void main(String[] args) throws IOException {

		File file = new File("E:\\Nirmal Java\\Eclipse Programs\\Excel\\Excel\\Personal.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Login");
		// Row row = sheet.getRow(i);
		// int count = sheet.getPhysicalNumberOfRows();
		// System.out.println(count);
		// int count1 = row.getPhysicalNumberOfCells();
		// System.out.println(count1);
//		for(int i=0;i<row.getPhysicalNumberOfCells();i++) {
//			Cell c = row.getCell(i);
//			System.out.println(c);	
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row2 = sheet.getRow(i);

			for (int j = 0; j < row2.getPhysicalNumberOfCells(); j++) {
				Cell cell = row2.getCell(j);
				CellType type = cell.getCellType();

				switch (type) {

				case STRING:
					String value = cell.getStringCellValue();
					System.out.println(value);
					break;
				case NUMERIC:

					if (DateUtil.isCellDateFormatted(cell)) {
						Date dateCellValue = cell.getDateCellValue();

						SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyy");
						String format = dateFormat.format(dateCellValue);
						System.out.println(format);
					} else {

//						double numericCellValue = cell.getNumericCellValue();
//						BigDecimal b = BigDecimal.valueOf(numericCellValue);
//						String num = b.toString();
//						System.out.println(num);
						double numericCellValue = cell.getNumericCellValue();
						long round = Math.round(numericCellValue);
						if (numericCellValue == round) {
							String valueOf = String.valueOf(round);
							System.out.println(valueOf);
						}else {
							String valueOf = String.valueOf(numericCellValue);
							System.out.println(valueOf);
						}
						}
					break;

				default:
					break;

				}

			}

		}
	}
}
