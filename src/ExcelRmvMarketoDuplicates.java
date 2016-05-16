import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

import javax.swing.JFileChooser;

/**
 * Author: Estevan McCalley
 * Date: 5/16/16
 * Description: This class allows user to select an Excel sheet formatted for 
 * Marketo upload, and removes duplicates
 */
public class ExcelRmvMarketoDuplicates {

	public static void main(String[] args) {
		
		JFileChooser fileChooser = new JFileChooser();
		int returnValue = fileChooser.showOpenDialog(null);
		
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			try {
				Workbook workbook = new XSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
				Sheet marketoSheet = workbook.getSheetAt(0);
				
				// Keeps track of # of duplicates
				int duplicateContacts = 0;
				
				// Remove duplicate rows
				System.out.println("REMOVING DUPLICATE ENTRIES...");
				for (int r = 1; r <= marketoSheet.getLastRowNum(); r++) {
					Row row = marketoSheet.getRow(r);
					// Check for duplicate entries and remove them
					Cell checkCell = row.getCell(0);
					Cell currentAdjacentCell = row.getCell(1);
					for (int c = (r + 1); c <= marketoSheet.getLastRowNum(); c++) {
						Row compareRow = marketoSheet.getRow(c);
						Cell compareCell = compareRow.getCell(0);
						Cell compareAdjacentCell = compareRow.getCell(1);
						// Compare content of 2 cells per contact to ensure it's a duplicate
						if ((checkCell.toString() == compareCell.toString())
								&& (currentAdjacentCell.toString() == compareAdjacentCell.toString())) {
							duplicateContacts++;
							System.out.println("REMOVED DUPLICATE ENTRY: " + compareCell.toString());
							// Clear duplicate row
							marketoSheet.removeRow(compareRow);
							if (marketoSheet.getRow(c + 1) != null) {
								// Remove cleared empty row
								marketoSheet.shiftRows(c + 1, marketoSheet.getLastRowNum(), -1);
							}
							// Reset row counter
							r--;
						} 
					}
				}
				
				// Log stats to the console
				System.out.println();
				System.out.println("Removed " + duplicateContacts + " duplicate entries");
	        
				// Create new file from input file data
		        try {
		            FileOutputStream output = new FileOutputStream("rmv-duplicates-output.xlsx");
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
