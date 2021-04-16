package Excel;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;
public class WriteXLSXFile {
	
	public static XSSFWorkbook book;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;
	
 @Test
 public void write() throws IOException {
	 File file =new File("./TestData/testdata.xlsx");
		if(!file.exists()) {
			file.createNewFile();
		}
		
		FileOutputStream fos = new FileOutputStream(file);
		
		book = new XSSFWorkbook();
		sheet = book.createSheet("Data");
		
		//1st row -->1st column
		row = sheet.createRow(0);
		cell = row.createCell(0);
		cell.setCellType(CellType.STRING);
		cell.setCellValue("Name");
		
		//1st row -->2nd column
		cell = row.createCell(1);
		cell.setCellType(CellType.STRING);
		cell.setCellValue("Place");
		
		
		//1st row -->3rd column
		cell = row.createCell(2);
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue("PinCode");
		
		
		//Updating the data
		WriteData(1, "John", "New York", 200234);
		WriteData(2, "Harry", "Virginia", 10023);
		WriteData(3, "Chris", "New Jersey", 523345);
		

		book.write(fos);
		book.close();
	 
 }
 
 public static void WriteData(int rownum,String name,String place,int pincode) {
		row = sheet.createRow(rownum);
		cell = row.createCell(0);
		cell.setCellType(CellType.STRING);
		cell.setCellValue(name);
		
		//1st row -->2nd column
		cell = row.createCell(1);
		cell.setCellType(CellType.STRING);
		cell.setCellValue(place);
		
		
		//1st row -->3rd column
		cell = row.createCell(2);
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(pincode);
	}
	
}
