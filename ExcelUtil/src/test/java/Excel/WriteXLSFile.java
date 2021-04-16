package Excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import org.testng.annotations.Test;

public class WriteXLSFile {
	
	public static HSSFWorkbook book;
	public static HSSFSheet sheet;
	public static HSSFRow row;
	public static HSSFCell cell;
	
	@Test
	public void write() throws IOException {
		
		File file =new File("./TestData/testdata1.xls");
		if(!file.exists()) {
			file.createNewFile();
		}
		
		FileOutputStream fos = new FileOutputStream(file);
		book = new HSSFWorkbook();
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
	
	
	
	
	
	
	
	


