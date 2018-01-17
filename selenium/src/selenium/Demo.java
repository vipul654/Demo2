package selenium;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.binary.XSSFBCommentsTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {

	public static void main(String[] args) throws IOException
	{
		FileInputStream fi = new FileInputStream("‪C:\\Users\\pc\\Desktop\\rt.xlsx");
		Workbook wb= new XSSFWorkbook(fi);
		wb.createSheet("newsheet");
		FileOutputStream fo = new FileOutputStream("‪C:\\Users\\pc\\Desktop\\rt.xlsx");
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();
		
		
		
		
		

	}

}
