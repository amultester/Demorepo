package Read_write;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilterInputStream;
import java.io.FilterOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_read_write {
	
	public static void main (String[]args) throws Exception {
	
		File src= new File ("C:\\Users\\innobot-user-1.LAPTOP-9DDO4JSH\\Downloads\\Users-Bulk-Import-Status (1).xlsx");
		
		FileInputStream input= new FileInputStream(src);
		XSSFWorkbook wb= new XSSFWorkbook(input);
		XSSFSheet sh1 = wb.getSheetAt(0);
	    String	 sheetvalue = sh1.getRow(0).getCell(0).getStringCellValue();
        System.out.println("The excel value is" +sheetvalue );
       
       
       //write
        //first commit
       
      sh1.getRow(0).createCell(17).setCellValue("25");
      FileOutputStream src1 = new FileOutputStream (new File("C:\\Users\\innobot-user-1.LAPTOP-9DDO4JSH\\Downloads\\Untitled 4.xlsx"));
      wb.write(src1);
      System.out.println("The new value is entered ");
		
	}
}

