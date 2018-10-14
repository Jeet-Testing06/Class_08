// Read data from .xlsx Excel file


package eighth_class;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ClassA {
	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\jitender.ahuja\\Desktop\\Latest.xlsx");
		FileInputStream fi = new FileInputStream(f);
		
		XSSFWorkbook xw = new XSSFWorkbook(fi);
		XSSFSheet xs = xw.getSheetAt(0);
		
		int r = xs.getPhysicalNumberOfRows();
		
		for(int i=0; i<r; i++)
		{
			
			XSSFRow xr = xs.getRow(i);
			for(int j=0; j<xr.getPhysicalNumberOfCells(); j++)
			{
				
				XSSFCell xc = xr.getCell(j);
				System.out.println(xc.getStringCellValue());
				
			}	
		}		
	}
}