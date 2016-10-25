package pl.seweryn;

import java.io.IOException;

import org.apache.poi.hssf.usermodel.*;

public class Main {

	public static void main(String[] args) {
		//Example test
		final String exampleOpenPath = "example\\309969A.xls";
		final String exampleSavePath = "C:\\Users\\Kuzyn\\Desktop\\test.xls";
		
		//Creating reference
		ReadWriteXLS xls = new ReadWriteXLS();
		HSSFWorkbook wb = null;
		
		//Reading file from path
		try {
			wb = xls.readFile(exampleOpenPath);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		//Checking if it works
		//wb.createSheet("new sheet");
		System.out.println(xls.getAllSheets().get(0).getSheetName());
		System.out.println(xls.getActiveLeftHeader(xls.getAllSheets().get(0)));
		
		
		//Saving file
		/*try {
			xls.saveSheetToFile(exampleSavePath);
		} catch (IOException e) {
			e.printStackTrace();
		}*/
	}

}
