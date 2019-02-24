package first;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.io.File;
import java.io.FileOutputStream;

import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingFromFile {

	public static void main(String[] args) throws Exception {
		// "C:\\Users\\lenovo pc\\Downloads\\samplefile.txt"

		String data = readFileAsString("C:\\Users\\lenovo pc\\Downloads\\samplefile.txt");
		int counter = 0;
		
		String row1 ,row2 ,row3 = null;
		String actual_Data = data.substring(3);
		if (counter == 0) {
			// System.out.println("D0:"+actual_Data);
			counter++;
			String[] splength = actual_Data.split("\\n");
			System.out.println("L:" + splength.length);
			for (int i = 0; i < 2; i++) {
				String row_data = splength[i];
				row1 = row_data.substring(0, 6);
				row2 = row_data.substring(6, 10);
				row3 = row_data.substring(10, 30);
				String value = "" + i + 1;
				 
				//Create blank workbook
			      XSSFWorkbook workbook = new XSSFWorkbook();
			      
			      //Create a blank sheet
			      XSSFSheet spreadsheet = workbook.createSheet( " Employee Info ");

			      //Create row object
			      XSSFRow row;
			      
			    //This data needs to be written (Object[])
			      Map < String, Object[] > empinfo = new TreeMap < String, Object[] >();
			      empinfo.put( value, new Object[] {
			         "ROW1", "ROW2", "ROW3" });
			      
			      empinfo.put( "2", new Object[] {
					         row1, row2, row3 });
			      
			    //Iterate over data and write to sheet
			      Set < String > keyid = empinfo.keySet();

			      int rowid = 0;
			      
			      for (String key : keyid) {
			         row = spreadsheet.createRow(rowid++);
			         Object [] objectArr = empinfo.get(key);
			         int cellid = 0;
			         
			         for (Object obj : objectArr){
			            Cell cell = row.createCell(cellid++);
			            cell.setCellValue((String)obj);
			         }
			      }
			      //Write the workbook in file system
			      FileOutputStream out = new FileOutputStream(
			         new File("C:\\Users\\lenovo pc\\Downloads\\Writesheet.xlsx"));
			      
			      workbook.write(out);
			      out.close();
			      System.out.println("Writesheet.xlsx written successfully");
//			      workbook.close();
			}
		}
}

	public static String readFileAsString(String fileName) throws Exception {
		String data = "";
		data = new String(Files.readAllBytes(Paths.get(fileName)));
		return data;
	}

	public static void length(String st) {
		String[] allVar = null;
		allVar = st.substring(2).split("\\n");
		System.out.println("L:" + allVar.length);
	}
}
