package pl.seweryn;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;

public class ReadWriteXLS {
	private HSSFWorkbook wb = new HSSFWorkbook();
	private ArrayList<HSSFSheet> sheets;

	//Read file
	public HSSFWorkbook readFile(String filename) throws IOException {
		FileInputStream fis = new FileInputStream(filename);
		try {
			wb = new HSSFWorkbook(fis);
			System.out.println("Opened file.");
			return wb;
		} finally {
			fis.close();
		}
	}

	//Finding all sheets
	public ArrayList<HSSFSheet> getAllSheets() {
		int NumberOfSheets = wb.getNumberOfSheets();

		ArrayList<HSSFSheet> sheets = new ArrayList<HSSFSheet>();

		for (int i = 0; i < NumberOfSheets; i++) {
			if (!wb.getSheetName(i).contains("Macro")) { //dont need sheet with name Macro
				sheets.add(wb.getSheetAt(i));
			}
		}

		this.sheets = sheets;
		return sheets;
	}
	
	//Get left header string from sheet
	public String getActiveLeftHeader(HSSFSheet sheet) {
		String headerText = "";
		HSSFHeader header = sheet.getHeader();
		headerText = header.getLeft();
		return headerText;
	}

	//Save file
	public void saveSheetToFile(String outputFilename) throws IOException {
		FileOutputStream out = new FileOutputStream(outputFilename);
		try {
			wb.write(out);
		} finally {
			out.close();
		}
		System.out.println("File Saved.");
	}

}
