package basepackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.Properties;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class BaseClass {
	
	private static String excelPath = 
			System.getProperty("user.dir") + "\\src\\main\\\\java\\files\\data.xlsx"; //Enter path of Excel
	
	private static String propPath = 
			System.getProperty("user.dir") + "\\src\\main\\java\\\\basepackage\\config.properties";
	
	private static String txtpath = System.getProperty("user.dir") + "\\src\\main\\java\\files\\data.txt";
	
	public static void excelToTxt(String sheetName)
	{
		FileInputStream fis;
		XSSFWorkbook book;
		Sheet sheet;
		Row row;
		
		Properties prop;
		
		StringBuilder sb = new StringBuilder();
		
		try {
			File source = new File(excelPath);
			fis = new FileInputStream(source);
			book = new XSSFWorkbook(fis);
			sheet =  book.getSheet(sheetName);
			
			prop = new Properties();
			prop.load(new FileInputStream(propPath));
			
			
			Row headerRow = sheet.getRow(0);
			
			int lastRowNu = sheet.getLastRowNum();
			int lastColNu = sheet.getRow(0).getLastCellNum();
			System.out.println("rows: " + lastRowNu + "  cols: " + lastColNu );
			
			for(int i = 1 ; i < lastRowNu+1; i++) 
			{
				row = book.getSheet(sheetName).getRow(i);
				
				for(int j = 0 ; j < lastColNu ; j++)
				{
					DataFormatter format = new DataFormatter();
					String data = format.formatCellValue(row.getCell(j));
			
					String colName = format.formatCellValue(headerRow.getCell(j));
					String propRead = prop.getProperty(colName);
					int reqdLength = Integer.parseInt(propRead); //careful
					
					data = formatString(data,reqdLength);
					
					sb.append(data);
				}

			}
			
			sb.append("\n");
			
			Path path = Paths.get(txtpath);
			Files.writeString(path, sb.toString(),StandardOpenOption.APPEND);
			
			sb.setLength(0);
			book.close();
			fis.close();
			} 
		catch (FileNotFoundException e)
		{
			e.printStackTrace();
		} 
		catch (IOException e) 
		{
			e.printStackTrace();
		} 
		catch(NumberFormatException e)
		{
			System.out.println("Exception while parsing property for col name from property file.");
			e.printStackTrace();
		}
		finally
		{
			book=null;
			fis=null;
		}
		
	}
	
	  public static String formatString(String s, int reqdSize)
	    {
		  	String output = "";
	        StringBuffer sbuffer = new StringBuffer(s.trim());
	        int orgLen = sbuffer.length();
	       
	        if(orgLen > reqdSize)
	        {
	            sbuffer.setLength(reqdSize);
	            output = sbuffer.toString();
	        }
	        else
	        {
		        int spacesReqd = reqdSize - orgLen;
		        output = sbuffer + " ".repeat(spacesReqd);
	        }
	       
	        return output;

	    }

	
	public static void main(String[] args) throws IOException {
		
		excelToTxt("header");
		excelToTxt("detailRecord");
		
		
	}

}
