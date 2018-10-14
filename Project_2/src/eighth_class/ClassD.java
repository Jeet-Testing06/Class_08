// Read from 1 excel and write in another excel


package eighth_class;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ClassD {
	
	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\jitender.ahuja\\Desktop\\Latest1.xlsx");
		File f1 = new File("C:\\Users\\jitender.ahuja\\Desktop\\Latest2.xlsx");
		
		FileInputStream fi = new FileInputStream(f);
		XSSFWorkbook xw = new XSSFWorkbook(fi);
		XSSFSheet xs = xw.getSheetAt(0);
		
		int r = xs.getPhysicalNumberOfRows();
		
		FileOutputStream fo = new FileOutputStream(f1);
		XSSFWorkbook xw1 = new XSSFWorkbook();
		XSSFSheet xs1 = xw1.createSheet("Jeet");
		
		
		
		for (int i =0; i<r; i++)
		{
			XSSFRow xr1 = xs1.createRow(i);
			XSSFRow xr = xs.getRow(i);
			for(int j=0; j<xr.getPhysicalNumberOfCells(); j++)
			{
		
				XSSFCell xc = xr.getCell(j);
			 // System.out.println(xc.getStringCellValue());
				
				XSSFCell xc1 = xr1.createCell(j);
				xc1.setCellValue(xc.getStringCellValue());
				
			}
			
			
		}
		
		
		xw1.write(fo);
		fo.flush();
		fo.close();
		
		
	}

}
