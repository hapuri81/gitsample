import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	public ArrayList<String> getData(String testCaseName) throws IOException
	{
		// TODO Auto-generated method stub
         ArrayList<String> a=new ArrayList();
		FileInputStream fis = new FileInputStream("C:\\Users\\harik\\OneDrive\\Desktop\\demoData.xlsx");
		XSSFWorkbook workBook = new XSSFWorkbook(fis);
		int sheets=workBook.getNumberOfSheets();
		for(int i=0;i<sheets;i++)
		{
			if(workBook.getSheetName(i).equalsIgnoreCase("Sheet1"))
					{
	                    XSSFSheet sheet=workBook.getSheetAt(i);	
	              Iterator<Row>  rows= sheet.iterator();
	                    Row firstrow=rows.next();
	                    Iterator<Cell> ce= firstrow.cellIterator();
	                   // ce.next();
	                    int k=0;
	                    int column=0;
	                    while(ce.hasNext())
	                    {
	                    	Cell value=ce.next();
	                    	if(value.getStringCellValue().equalsIgnoreCase("TestCases"))
	                    	{
	                    		column=k;
	                    
	                    	}
	                    		k++;	
	                    }
	                
	                    while(rows.hasNext())
	                    {
	                    	Row r=rows.next();
	                    	if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName))
	                    	{
	                    		Iterator<Cell> cv=r.cellIterator();
	                    		while(cv.hasNext())
	                    		{
	                    			Cell c=cv.next();
	                    			if(c.getCellType()==CellType.STRING)
	                    			{
	                    			a.add(c.getStringCellValue());
	                    			}
	                    			else
	                    			{ 
	                    				a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
	                    			}
	                    		}
	                    	}
	                    }
				
					}
		}
		return a;
	}

}
