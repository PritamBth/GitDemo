import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriver {

	public static void main(String[] args) throws IOException 
	{
		FileInputStream fl=new FileInputStream("C:\\Users\\Admin\\workspace\\PrimeMo\\xlxl\\data1\\Book1.xlsx");
		XSSFWorkbook wk=new XSSFWorkbook(fl);
		int count=wk.getNumberOfSheets();
		System.out.println(count);
		for(int i=0;i<count;i++)
		{
			if(wk.getSheetName(i).equalsIgnoreCase("pk"))
			{
			
			XSSFSheet sheet=wk.getSheetAt(i);
			String name=sheet.getSheetName();
			System.out.println(name);
			//identify Test-case column by scanning the entire 1st row
			Iterator<Row> rows=sheet.iterator();
			Row firstrow=rows.next();
		    Iterator<Cell> ce=firstrow.cellIterator();
		    
		    int k=0;
		    int col = 0; 
		    while(ce.hasNext())
		    {
		    	Cell v=ce.next();
		    	if(v.getStringCellValue().equalsIgnoreCase("Data1"))
		    	{
		    		col=k;
		    	}
		    	k++;
		    	}
		    System.out.println(col);
			}
			
		}
		
		
	}

}
