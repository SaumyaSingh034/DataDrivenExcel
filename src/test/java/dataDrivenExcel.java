import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDrivenExcel {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream("‪D://Selenium//ExcelDOc//DemoData.xlsx");
		//Get control over Excel
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		//Number of sheets
		int sheets = workbook.getNumberOfSheets();
		for(int i=0;i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata"))
				
			{
				int k = 0, column = 0;
				//Get Control over the particular sheet
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				//get access to entire row
				Iterator<Row> rows = sheet.iterator();
				Row firstrow = rows.next();
				
				//get access to specific row and cell
				Iterator<Cell> cel = firstrow.cellIterator();
				while(cel.hasNext())
				{
					//Store value of cell in first row
					Cell value = cel.next();
					if(value.getStringCellValue().equalsIgnoreCase("TestCase"))
					{
						column = k;
						
					}
					k++;
					System.out.println(column);
				}
				
			}
		}

	}

}
