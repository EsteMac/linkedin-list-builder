import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

import javax.swing.JFileChooser;

/**
 * Author: Estevan McCalley
 * Date: 5/20/16
 * Description: This class allows user to select an Excel workbook and search for contacts
 * that have email addresses that appear on both sheets and highlights them
 */
public class ExcelCompareLists {
	// Locations of data to compare on first sheet
	public final static int REGISTERED_LIST_SHEET_INDEX = 0;
	public final static int REGISTERED_EMAIL_COLUMN_INDEX = 1;
	
	// Locations of data to compare on second sheet
	public final static int MARKETO_UPLOAD_LIST_SHEET_INDEX = 1;
	public final static int MARKETO_EMAIL_COLUMN_INDEX = 3;

	public static void main(String[] args) {
		
		JFileChooser fileChooser = new JFileChooser();
		int returnValue = fileChooser.showOpenDialog(null);
		
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			try {
				Workbook workbook = new XSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
				Sheet registeredSheet = workbook.getSheetAt(REGISTERED_LIST_SHEET_INDEX);
				Sheet marketoSheet = workbook.getSheetAt(MARKETO_UPLOAD_LIST_SHEET_INDEX);
				
				// Keeps track of # of duplicates
				int duplicateContacts = 0;
				
				// style white fonts
		        Font whiteFont = workbook.createFont();
		        whiteFont.setColor(IndexedColors.WHITE.getIndex());
		        whiteFont.setBoldweight(Font.BOLDWEIGHT_NORMAL);
		        whiteFont.setFontHeightInPoints((short)12);
		        
		        // style blue cells with white fonts
		        CellStyle styleBlueWhite = workbook.createCellStyle();
		        styleBlueWhite.setFillForegroundColor(IndexedColors.AQUA.getIndex());
		        styleBlueWhite.setFillPattern(CellStyle.SOLID_FOREGROUND);
		        styleBlueWhite.setAlignment(CellStyle.ALIGN_CENTER);
		        styleBlueWhite.setFont(whiteFont);
		        styleBlueWhite.setWrapText(true);
		        
				// Remove rows with no email address from the first sheet
				for (int r = 1; r <= registeredSheet.getLastRowNum(); r++) {
					Row row = registeredSheet.getRow(r);
					Cell registeredEmailCell = row.getCell(REGISTERED_EMAIL_COLUMN_INDEX);
					// Check for empty rows
				    if(registeredEmailCell == null){
				    	registeredSheet.shiftRows(r + 1, registeredSheet.getLastRowNum(), -1);
				        r--;
				        continue;
				    }
				}
				
				// Find and highlight contacts appearing in both sheets
				System.out.println("SEARCHING FOR CONTACTS ON BOTH SHEETS...");
				for (int r = 1; r <= registeredSheet.getLastRowNum(); r++) {
					Row row = registeredSheet.getRow(r);
					// Check for duplicate entries and highlight them
					Cell registeredEmailCell = row.getCell(REGISTERED_EMAIL_COLUMN_INDEX);
					for (int c = 1; c <= marketoSheet.getLastRowNum(); c++) {
						Row compareRow = marketoSheet.getRow(c);
						Cell marketoEmailCell = compareRow.getCell(MARKETO_EMAIL_COLUMN_INDEX);
						// Compare content of both cells
						if ((marketoEmailCell.toString() == registeredEmailCell.toString())
								&& !registeredEmailCell.toString().isEmpty()) {
							duplicateContacts++;
							System.out.println("FOUND COMMON ENTRY: " + registeredEmailCell.toString());
							// Highlight common contact row on registered list
							registeredEmailCell.setCellStyle(styleBlueWhite);
						} 
					}
				}
				
				// Log stats to the console
				System.out.println();
				System.out.println("Found " + duplicateContacts + " common entries");
	        
				// Create new file from input file data
		        try {
		            FileOutputStream output = new FileOutputStream("compare-lists-output.xlsx");
		            workbook.write(output);
		            output.close();
		        } catch(Exception e) {
		            e.printStackTrace();
		        }
		        
		        // Close the workbook
		        try {
					workbook.close();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
		        
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		}

	}
	
}
