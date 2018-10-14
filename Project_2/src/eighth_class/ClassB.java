// Write data in .xlsx Excel


package eighth_class;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ClassB {
	public static void main(String[] args) throws IOException {
	
		File f = new File("C:\\Users\\jitender.ahuja\\Desktop\\Latest1.xlsx");
		FileOutputStream fo = new FileOutputStream(f);
		
		XSSFWorkbook xw = new XSSFWorkbook();
		XSSFSheet xs = xw.createSheet("Jeet");
		
		for (int i=0; i<5; i++)
		{
			
			XSSFRow xr = xs.createRow(i);
			for (int j=0; j<5; j++)
			{
				
				XSSFCell xc = xr.createCell(j);
				xc.setCellValue("Jitender");
			}
				
				
			
		}
				
		xw.write(fo);
		fo.flush();
		fo.close();
		
	}
}