package Apache.POI;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Workbook -->Sheet-->Row-->Cells
public class App 
{
    public static void main( String[] args ) throws IOException
    {
    	XSSFWorkbook workbook=new XSSFWorkbook();
    	XSSFSheet sheet=workbook.createSheet("Emp Info");
    	
    	/*
    	Object empdata[][]= {
    			{"EmpId","Name","Job"},
    			{101,"Abhay","PAT"},
    			{102,"Amit","PA"},
    			{103,"Ayush","PAT"},
    	};
    	*/
    	
    	ArrayList<Object[]> empdata=new ArrayList<Object[]>();
    	empdata.add(new Object[] {"Empid","Name","Job"});
    	empdata.add(new Object[] {101,"Ashwani","PA"});
    	empdata.add(new Object[] {102,"Abhay","PAT"});
    	empdata.add(new Object[] {102,"Raj","PAT"});
    	
    	//using for loop
    	 /*int rows=empdata.length;
    	 int cols=empdata[0].length;
    	 System.out.println(rows);
    	 System.out.println(cols);
    	 
    	 for(int r=0;r<rows;r++)
    	 {
    		 XSSFRow row=sheet.createRow(r);
    		 for(int c=0;c<cols;c++)
    		 {
    			 XSSFCell cell=row.createCell(c);
    			 Object value =empdata[r][c];
    			 
    			 if(value instanceof String)
    				 cell.setCellValue((String)value);
    			 if(value instanceof Integer)
    				 cell.setCellValue((Integer)value);
    			 if(value instanceof Boolean)
    				 cell.setCellValue((Boolean)value);
    		 }
    	 }*/
    	
    	// using for each loop
    	/*
    	int rowCount=0;
    	
    	for(Object emp[]:empdata)// first row will capture to emp
    	{
    		XSSFRow row=sheet.createRow(rowCount++); //post inc
    		int columnCount=0;
    		for(Object value:emp) 
    		{
    			XSSFCell cell=row.createCell(columnCount++);
    			if(value instanceof String)
    				cell.setCellValue((String)value);
    			if(value instanceof Integer)
    				cell.setCellValue((Integer)value);
   			 	if(value instanceof Boolean)
   			 		cell.setCellValue((Boolean)value);
    			
    		}
    	}*/
    	
    	int rowCount=0;
    	
    	for(Object[] emp:empdata)
    	{
    		XSSFRow row=sheet.createRow(rowCount++); //post inc
    		int columnCount=0;
    		for(Object value:emp)
    		{

    			XSSFCell cell=row.createCell(columnCount++);
    			if(value instanceof String)
    				cell.setCellValue((String)value);
    			if(value instanceof Integer)
    				cell.setCellValue((Integer)value);
   			 	if(value instanceof Boolean)
   			 		cell.setCellValue((Boolean)value);
    		}
    	}
    	 
    	 String filepath=".\\data\\employee.xlsx";
    	 FileOutputStream outStream=new FileOutputStream(filepath);
    	 workbook.write(outStream);
    	 outStream.close();
    	 
    	 System.out.println("Emplyee file writen successfully:");
    }
}
