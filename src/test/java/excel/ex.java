package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ex {

	
	public static void main(String[] args) throws Throwable{
		File f = new File("/Users/gowtham/Desktop/book1.xlxs");
		FileInputStream f1 = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(f1);
		
		//.xls -- HSSSFWorkbook  , .xlsx -- XSSFWorkbook
		Sheet s = w.getSheet("sheet1");
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				int cellType = c.getCellType();
				// 0 -- Number or Date 1 -- String
				if (cellType==1) {
					String value = c.getStringCellValue();
					System.out.println(value);
				}
				else if (cellType==0) {
					if (DateUtil.isCellDateFormatted(c)) {
						Date d = c.getDateCellValue();
						SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyyy");
						String value = sd.format(d);
						System.out.println(value);
					}
					else {
						double d = c.getNumericCellValue();
						long l = (long)d;
						String value = String.valueOf(l);
						System.out.println(value);
					}
				}
			}
		}
		
}
}
