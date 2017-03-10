import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

// here my library import that i took from internet

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadandCreate { 

	// my main class
	public static void main(String[] args) throws IOException
	{
	
			//Blank workbook
			XSSFWorkbook workbook = new XSSFWorkbook(); 
			
			//Create a blank sheet
			XSSFSheet sheet = workbook.createSheet();
			 
			//This data needs to be written (Object[])
			// it include my excel file.
			Map<String, Object[]> data = new TreeMap<String, Object[]>();
			data.put("1", new Object[] {"ID", "Dept", "Grade", "Letter Grade"});
			data.put("2", new Object[] {1234, "CMPE", 90, "AB"});
			data.put("3", new Object[] {2345, "EEEN", 95, "AA"});
			data.put("4", new Object[] {4353, "EEEN", 90, "AB"});
			data.put("5", new Object[] {3424, "CMPE", 95, "AA"});
			data.put("6", new Object[] {3423, "CMPE", 100, "AA"});
		
			
			//Iterate over data and write to sheet
			Set<String> keyset = data.keySet();
			int rownum = 0;
			for (String key : keyset)
			{
				
			    Row row = sheet.createRow(rownum++);
			    Object [] objArr = data.get(key);
			    int cellnum = 0;
			    for (Object obj : objArr){
			    	
			       Cell cell = row.createCell(cellnum++);
			       if(obj instanceof String)
			            cell.setCellValue((String)obj);
			       
			        else if(obj instanceof Integer)
			            cell.setCellValue((Integer)obj);
			     }
			}
			try 
			{
				//Write the workbook in file system
			    FileOutputStream out = new FileOutputStream(new File("grades.xls"));
			    workbook.write(out);
			    out.close();
			    System.out.println("okundu");
			} 
			catch (Exception e){
			    e.printStackTrace();
			}
			
	
	
		// here my EXCEL reader.
		try
		{
			FileInputStream file = new FileInputStream(new File("grades.xls"));
			XSSFWorkbook workbook1 = new XSSFWorkbook(file);
			XSSFSheet sheet1 = workbook1.getSheetAt(0);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet1.iterator();
			
			while (rowIterator.hasNext()){
				Row row = rowIterator.next();
				
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while (cellIterator.hasNext()){
					Cell cell = cellIterator.next();
					
					//Check the cell type and format accordingly
					switch (cell.getCellType()){
						case Cell.CELL_TYPE_NUMERIC:
							System.out.print(cell.getNumericCellValue() + " ");
							break;
						case Cell.CELL_TYPE_STRING:
							System.out.print(cell.getStringCellValue() + " ");
							break;
					}
				}
				System.out.println("");
			}
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		
		//end of reading excel
		
		
			// reading Txt file.
			// arraylist for my data.
		 	ArrayList<ArrayList<String>> all = null;
	        ArrayList<String> oneRow = null;
	        String currentLine;
	        FileInputStream fis = new FileInputStream("person.txt");
	        DataInputStream myInput = new DataInputStream(fis);
	        
	        // loop starting.
	        int i = 0;
	        all = new ArrayList<ArrayList<String>>();
	       
	        // Go on until you can not find any data in txt file.
	        while ((currentLine = myInput.readLine()) != null)
	        {
	            oneRow = new ArrayList<String>();
	            String oneData[] = currentLine.split(",");
	            for (int j = 0; j < oneData.length; j++){
	                oneRow.add(oneData[j]);
	            }
	            
	            all.add(oneRow);
	            System.out.println();
	            i++;
	        }
	        
	     // writing part.
	     try 
	     {
	         HSSFWorkbook workBook1 = new HSSFWorkbook();
	         HSSFSheet sheet1 = workBook1.createSheet("sheet1");
	         
	         // loop for my array.
	         for (int i1 = 0; i1 < all.size(); i1++)
	         {
	           ArrayList<?> ardata = (ArrayList<?>) all.get(i1);
	           HSSFRow row = sheet1.createRow((short) 0 + i1);
	           
	           // loop for my data.
	           for (int k = 0; k < ardata.size(); k++)
	           {
	                System.out.print(ardata.get(k));
	                HSSFCell cell = row.createCell((short) k);
	                cell.setCellValue(ardata.get(k).toString());	           
	                }
	           System.out.println();
	        }
	       FileOutputStream fileOutputStream =  new FileOutputStream("merged.xls");
	       workBook1.write(fileOutputStream);
	       fileOutputStream.close();
	     } 
	     catch (Exception e) {
	    	 
	     }
	     
	     // end of text reader
	     
	}// end of main
}// end of class
