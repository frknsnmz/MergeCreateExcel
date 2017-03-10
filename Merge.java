import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

// my importation that i directly took from internet.
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Merge {

	//	Main method 
	public static void main(String args[]) throws IOException {
		
		// my arraylist.
		ArrayList<ArrayList<String>> all = new ArrayList<ArrayList<String>>();
		ArrayList<String> oneRow = null;
		String currentLine;
		
		FileInputStream fis = new FileInputStream("person.txt");
		DataInputStream myInput = new DataInputStream(fis);

		while ((currentLine = myInput.readLine()) != null){
			
			oneRow = new ArrayList<String>();
			String oneData[] = currentLine.split(",");

			for (int j = 0; j < oneData.length; j++){
				oneRow.add(oneData[j]);
				}
			all.add(oneRow);
			}

		System.out.println("txt dosya okundu");
		
		//These operations for reading the data
		FileInputStream file = new FileInputStream(new File("grades.xls"));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//important for my row numbers.
		int rown=-1;
		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()){
			rown=rown+1;
			Row row = rowIterator.next();
			}
		
//	Checking id for insertion.
		for(int q=0; q<all.size(); q++){
			for(int w=1; w<=rown; w++){
				if(Double.parseDouble(all.get(q).get(0))==sheet.getRow(w).getCell(0).getNumericCellValue()){
					for(int e=1; e<=3; e++){
						if (sheet.getRow(w).getCell(e).getCellType()==Cell.CELL_TYPE_NUMERIC)
							all.get(q).add(Double.toString(sheet.getRow(w).getCell(e).getNumericCellValue()));
						else if (sheet.getRow(w).getCell(e).getCellType()== Cell.CELL_TYPE_STRING)
							all.get(q).add(sheet.getRow(w).getCell(e).getStringCellValue());
						
						}
					}
				else;
			}		
		}
		/*
		 * 
		 * for(int e=1; e<=3; e++){
						if (sheet.getRow(w).getCell(e).getCellType()==Cell.CELL_TYPE_NUMERIC)
							all.get(q).add(Double.toString(sheet.getRow(w).getCell(e).getNumericCellValue()));
						else if (sheet.getRow(w).getCell(e).getCellType()== Cell.CELL_TYPE_STRING)
							all.get(q).add(sheet.getRow(w).getCell(e).getStringCellValue());
						
		 * 
		 */
		

		file.close();
		System.out.println("xls dosyasý okundu");
		
		
		//	Lets combine all of them !
		try {
			
			HSSFWorkbook workbookfinal = new HSSFWorkbook();
			HSSFSheet sheetfinal = workbookfinal.createSheet("sheet1");

			//My rows are ; 
			String a[] ={"Id", "Name", "Surname", "Gender", "Age", "Dept", "Grade", "Letter Grade" };
			HSSFRow rowfirst= sheetfinal.createRow(0);
			
			for(int x=0; x<a.length; x++){
				HSSFCell cellfirst=rowfirst.createCell(x);
				cellfirst.setCellValue(a[x]);
				}

			for (int s = 0; s < all.size(); s++){
				ArrayList<?> ardata = (ArrayList<?>) all.get(s);
				HSSFRow row = sheetfinal.createRow((short) 1 + s);
				
				for (int k = 0; k < ardata.size(); k++){
					HSSFCell cell = row.createCell((short) k);
					cell.setCellValue(ardata.get(k).toString());	           
					}
	        }
	         
	       FileOutputStream out =  new FileOutputStream("merged.xls");
	       workbookfinal.write(out);
	       out.close();
	     } 
		
	     catch (Exception e) {
	    	 
	     }
		
		 System.out.println("merged.xls dosyasý yazýldý.");
    }// end of the main
}// end of the class
