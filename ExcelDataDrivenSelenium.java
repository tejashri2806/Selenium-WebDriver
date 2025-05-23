import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataDrivenSelenium {

	public static void main(String[] args) throws IOException   {
		// TODO Auto-generated method stub
		
		// FileInputStream class in Java is useful for reading data from a file in the form of a Java sequence of bytes.
		FileInputStream file = new FileInputStream("C:\\Users\\tejas\\Downloads\\Sample Data.xlsx");
		
		//XSSFWorkbook is a class that is used to represent both high and low level Excel file formats
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		int sheets = workbook.getNumberOfSheets();
		for(int i=0;i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("Office Supply List"))
			{
				////Identify Country column by scanning the entire 1st row
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				/// sheet is collection of rows
				Iterator<Row> Rows = sheet.iterator();
				Row Firstrow = Rows.next();								
				
//				/////row is collection of cells
				Iterator<Cell>ce=Firstrow.cellIterator();
				int k =0;
				int column =0;
				while(ce.hasNext()) 
				{
					Cell Value = ce.next();
					if(Value.getStringCellValue().equalsIgnoreCase("Country"))
					{
						column=k;
					}
					k++;
				}
				System.out.println(column);
				
				//////once column is identified then scan entire Country column to identify row with "India" value in it
				while(Rows.hasNext())
				{
					Row r = Rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("India"))
					{
						Iterator<Cell>cv=r.cellIterator();
						while(cv.hasNext()) 
						{
							System.out.println(cv.next().getStringCellValue());
						}
					}
				}
				
				
				
			}
		}
		
		
	}

}
