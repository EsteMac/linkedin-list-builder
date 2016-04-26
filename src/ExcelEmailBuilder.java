import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.stream.Stream;
import javax.swing.JFileChooser;

/**
 * Author: Estevan McCalley
 * Date: 4/14/16
 * Description: This class allows user to select a DataMiner output file to be processed
 * Currently only supports xlsx input files
 */
public class ExcelEmailBuilder {

	public static void main(String[] args) {
		// Set column index containing contact full names. All other columns are referred to relative to this one
		final int nameColumnIndex = 0; 
		
		JFileChooser fileChooser = new JFileChooser();
		int returnValue = fileChooser.showOpenDialog(null);
		
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			try {
				Workbook workbook = new XSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
				Sheet sheet = workbook.getSheetAt(0);
				
				// Strings used to build email addresses
				String fullName;
				String firstName;
				String middleName;
				String lastName;
				String accountName;
				String domainName;
				String currentName;
				String titleName;
				
				// HashMap to store the number of unique domain / email identified per named account
			    Map<String, Integer> map = new HashMap<String, Integer>();
			    map.put("aarp.org", 0);
			    map.put("advance-auto.com", 0);
			    map.put("ahss.org", 0);
			    map.put("aflac.com", 0);
			    map.put("altisource.com", 0);
			    map.put("amerisourcebergen.com", 0);
			    map.put("astrazeneca.com", 0);
			    map.put("autonation.com", 0);
			    map.put("avidxchange.com", 0);
			    map.put("baycare.org", 0);
			    map.put("bbandt.com", 0);
			    map.put("benefitfocus.com", 0);
			    map.put("blackbaud.com", 0);
			    map.put("blackboard.com", 0);
			    map.put("bcbsfl.com", 0);
			    map.put("bcbsnc.com", 0);
			    map.put("bdpint.com", 0);
			    map.put("carecorenational.com", 0);
			    map.put("carefirst.com", 0);
			    map.put("carmax.com", 0);
			    map.put("carnival.com", 0);
			    map.put("catalinamarketing.com", 0);
			    map.put("citrix.com", 0);
			    map.put("comscore.com", 0);
			    map.put("autotrader.com", 0);
			    map.put("csx.com", 0);
			    map.put("danaher.com", 0);
			    map.put("darden.com", 0);
			    map.put("dell.com", 0);
			    map.put("delta.com", 0);
			    map.put("dollartree.com", 0);
			    map.put("dominionenterprises.com", 0);
			    map.put("duke-energy.com", 0);
			    map.put("usa.dupont.com", 0);
			    map.put("equifax.com", 0);
			    map.put("footballfanatics.com", 0);
			    map.put("fanniemae.com", 0);
			    map.put("fnf.com", 0);
			    map.put("fiserv.com", 0);
			    map.put("fpl.com", 0);
			    map.put("freddiemac.com", 0);
			    map.put("freedommortgage.com", 0);
			    map.put("fticonsulting.com", 0);
			    map.put("geico.com", 0);
			    map.put("gdit.com", 0);
			    map.put("ge.com", 0);
			    map.put("genworth.com", 0);
			    map.put("harris.com", 0);
			    map.put("healthesystems.com", 0);
			    map.put("hilton.com", 0);
			    map.put("ihg.com", 0);
			    map.put("imshealth.com", 0);
			    map.put("inovalon.com", 0);
			    map.put("jmfamily.com", 0);
			    map.put("labcorp.com", 0);
			    map.put("lfg.com", 0);
			    map.put("lowes.com", 0);
			    map.put("macys.com", 0);
			    map.put("manh.com", 0);
			    map.put("markelcorp.com", 0);
			    map.put("marriott.com", 0);
			    map.put("effem.com", 0);
			    map.put("masonite.com", 0);
			    map.put("mckesson.com", 0);
			    map.put("merck.com", 0);
			    map.put("microstrategy.com", 0);
			    map.put("moffitt.org", 0);
			    map.put("mohawkind.com", 0);
			    map.put("nascar.com", 0);
			    map.put("underarmour.com", 0);
			    map.put("navyfederal.org", 0);
			    map.put("ncr.com", 0);
			    map.put("neustar.biz", 0);
			    map.put("nielsen.com", 0);
			    map.put("ngc.com", 0);
			    map.put("officedepot.com", 0);
			    map.put("praintl.com", 0);
			    map.put("publix.com", 0);
			    map.put("raymondjames.com", 0);
			    map.put("transcore.com", 0);
			    map.put("rccl.com", 0);
			    map.put("sas.com", 0);
			    map.put("sita.aero", 0);
			    map.put("subaru.com", 0);
			    map.put("sungardas.com", 0);
			    map.put("sykes.com", 0);
			    map.put("synchronoss.com", 0);
			    map.put("syniverse.com", 0);
			    map.put("troweprice.com", 0);
			    map.put("chicos.com", 0);
			    map.put("email.chop.edu", 0);
			    map.put("na.ko.com", 0);
			    map.put("hersheys.com", 0);
			    map.put("homedepot.com", 0);
			    map.put("vanguard.com", 0);
			    map.put("hcsc.net", 0);
			    map.put("travelport.com", 0);
			    map.put("tsys.com", 0);
			    map.put("ultimatesoftware.com", 0);
			    map.put("ugcorp.com", 0);
			    map.put("ups.com", 0);
			    map.put("verisign.com", 0);
			    map.put("vertexinc.com", 0);
			    map.put("wfu.edu", 0);
			    map.put("wellcare.com", 0);
			    map.put("NONE FOUND", 0);
				
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
				
				// Create the title row
				Row titleRow = sheet.getRow(0);
				Cell firstNameTitleCell = titleRow.createCell(nameColumnIndex + 8);
				firstNameTitleCell.setCellValue("FirstName");
				firstNameTitleCell.setCellStyle(styleBlueWhite);
				Cell middleNameTitleCell = titleRow.createCell(nameColumnIndex + 9);
				middleNameTitleCell.setCellValue("MiddleName");
				middleNameTitleCell.setCellStyle(styleBlueWhite);
				Cell lastNameTitleCell = titleRow.createCell(nameColumnIndex + 10);
				lastNameTitleCell.setCellValue("LastName");
				lastNameTitleCell.setCellStyle(styleBlueWhite);
				Cell domainTitleCell = titleRow.createCell(nameColumnIndex + 11);
				domainTitleCell.setCellValue("Domain");
				domainTitleCell.setCellStyle(styleBlueWhite);
				Cell emailTitleCell = titleRow.createCell(nameColumnIndex + 12);
				emailTitleCell.setCellValue("Email");
				emailTitleCell.setCellStyle(styleBlueWhite);
				
				// Keeps track of # of empty rows
				int emptyRows = 0;
				
				// Keeps track of # of unidentified domain names
				int domainsNotFound = 0;
				
				// Keeps track of # of rows where domain & email are identified
				int nonEmptyRows = 0;
				
				// Keeps track of # of rows without any mutual remote connections
				// where the name shows up as "LinkedIn Member"
				int linkedinMembers = 0;
				
				// Keeps track of # of duplicates
				int duplicateContacts = 0;
				
				// Remove empty rows
				for (int r = 1; r <= sheet.getLastRowNum(); r++) {
					Row row = sheet.getRow(r);
					// Check for empty rows
				    if(row == null){
				        sheet.shiftRows(r + 1, sheet.getLastRowNum(), -1);
				        emptyRows++;
				        r--;
				        continue;
				    }
				}
				
				// Remove "LinkedIn Member" rows
				for (int r = 1; r <= sheet.getLastRowNum(); r++) {
					Row row = sheet.getRow(r);
					// Check for entries with names listed as "LinkedIn" and remove them
					Cell linkedinCell = row.getCell(nameColumnIndex);
					if (linkedinCell.toString().toLowerCase().contains("linkedin")) {
						linkedinMembers++;
						System.out.println("NON-CONNECTED 'LinkedIn Member' ENTRY REMOVED: " + linkedinCell.toString());
						// Clear row with no name or connections
						sheet.removeRow(row);
						// Remove cleared empty row
						sheet.shiftRows(r + 1, sheet.getLastRowNum(), -1);
						// Reset row counter
						r--;
						continue;
					}
				}
				
				// Remove duplicate rows
				for (int r = 1; r <= sheet.getLastRowNum(); r++) {
					Row row = sheet.getRow(r);
					// Check for duplicate entries and remove them
					Cell checkCell = row.getCell(nameColumnIndex);
					Cell currentAdjacentCell = row.getCell(nameColumnIndex + 1);
					for (int c = (r + 1); c <= sheet.getLastRowNum(); c++) {
						Row compareRow = sheet.getRow(c);
						Cell compareCell = compareRow.getCell(nameColumnIndex);
						Cell compareAdjacentCell = compareRow.getCell(nameColumnIndex + 1);
						// Compare content of 2 cells per contact to ensure it's a duplicate
						if ((checkCell.toString() == compareCell.toString())
								&& (currentAdjacentCell.toString() == compareAdjacentCell.toString())) {
							duplicateContacts++;
							System.out.println("DUPLICATE ENTRY REMOVED: " + compareCell.toString());
							// Clear duplicate row
							sheet.removeRow(compareRow);
							// Remove cleared empty row
							sheet.shiftRows(c + 1, sheet.getLastRowNum(), -1);
							// Reset row counter
							r--;
						} 
					}
				}
				
				// Iterate row by row and assign domain / email addresses
				for (int r = 1; r <= sheet.getLastRowNum(); r++) {
					Row row = sheet.getRow(r);		
				    
					// Split full name into separate cells
					fullName = row.getCell(nameColumnIndex).toString();
					String[] names = fullName.split("\\s+");
					for (int i = 0; i < names.length; i++) {
					    // Check for a non-word character and replace
					    names[i] = names[i].replaceAll("[^\\w]", "");
					}					
					Cell firstCell = row.createCell(nameColumnIndex + 8);
					firstCell.setCellStyle(styleBlueWhite);
					Cell middleCell = row.createCell(nameColumnIndex + 9);
					middleCell.setCellStyle(styleBlueWhite);
					Cell lastCell = row.createCell(nameColumnIndex + 10);
					lastCell.setCellStyle(styleBlueWhite);
					if (names.length == 3) {
						firstName = names[0];
						middleName = names[1];
						lastName = names[2];
					} else if (names.length == 2) {
						firstName = names[0];
						middleName = null;
						lastName = names[1];
					} else {
						firstName = names[0];
						middleName = null;
						lastName = null;
					}
					firstCell.setCellValue(firstName); 
					middleCell.setCellValue(middleName);
					lastCell.setCellValue(lastName);
					
					// Create domain names based on "company" field if info available
					Cell domainCell = row.createCell(nameColumnIndex + 11);
					domainCell.setCellStyle(styleBlueWhite);
					Cell accountCell = row.getCell(nameColumnIndex + 3);
					if (accountCell != null) {
						accountName = accountCell.toString().toLowerCase();
					} else {
						accountName = "";
					}
					
					// Create domain names based on "current" field if info available
					Cell currentCell = row.getCell(nameColumnIndex + 6);
					if (currentCell != null) {
						currentName = currentCell.toString().toLowerCase();
					} else {
						currentName = "";
					}
					
					// Create domain names based on "title" field if info available
					Cell titleCell = row.getCell(nameColumnIndex + 1);
					if (titleCell != null) {
						titleName = titleCell.toString().toLowerCase();
					} else {
						titleName = "";
					}
					
					if (accountName.contains("aarp")
							|| currentName.contains("aarp")
							|| titleName.contains("aarp")) {
						domainName = "aarp.org";
					} else if (accountName.contains("advance auto") 
							|| accountName.contains("advance-auto") 
							|| accountName.contains("advanceauto")
							|| currentName.contains("advance auto") 
							|| currentName.contains("advance-auto") 
							|| currentName.contains("advanceauto")
							|| titleName.contains("advance auto") 
							|| titleName.contains("advance-auto") 
							|| titleName.contains("advanceauto")) {
						domainName = "advance-auto.com";
					} else if (accountName.contains("adventist")
							|| currentName.contains("adventist")
							|| titleName.contains("adventist")) {
						domainName = "ahss.org";
					} else if (accountName.contains("aflac")
							|| currentName.contains("aflac")
							|| titleName.contains("aflac")) {
						domainName = "aflac.com";
					} else if (accountName.contains("altisource")
							|| currentName.contains("altisource")
							|| titleName.contains("altisource")) {
						domainName = "altisource.com";
					} else if (accountName.contains("amerisourcebergen")
							|| currentName.contains("amerisourcebergen")
							|| titleName.contains("amerisourcebergen")) {
						domainName = "amerisourcebergen.com";
					} else if (accountName.contains("astrazeneca")
							|| currentName.contains("astrazeneca")
							|| titleName.contains("astrazeneca")) {
						domainName = "astrazeneca.com";
					} else if (accountName.contains("autonation")
							|| currentName.contains("autonation")
							|| titleName.contains("autonation")) {
						domainName = "autonation.com";
					} else if (accountName.contains("avidxchange")
							|| currentName.contains("avidxchange")
							|| titleName.contains("avidxchange")) {
						domainName = "avidxchange.com";
					} else if (accountName.contains("baycare")
							|| currentName.contains("baycare")
							|| titleName.contains("baycare")) {
						domainName = "baycare.org";
					} else if (accountName.contains("bb&t") 
							|| accountName.contains("branch banking")
							|| currentName.contains("bb&t") 
							|| currentName.contains("branch banking")
							|| titleName.contains("bb&t") 
							|| titleName.contains("branch banking")) {
						domainName = "bbandt.com";
					} else if (accountName.contains("benefitfocus")
							|| currentName.contains("benefitfocus")
							|| titleName.contains("benefitfocus")) {
						domainName = "benefitfocus.com";
					} else if (accountName.contains("blackbaud")
							|| currentName.contains("blackbaud")
							|| titleName.contains("blackbaud")) {
						domainName = "blackbaud.com";
					} else if (accountName.contains("blackboard")
							|| currentName.contains("blackboard")
							|| titleName.contains("blackboard")) {
						domainName = "blackboard.com";
					} else if ((accountName.contains("blue cross") 
							&& accountName.contains("florida")) 
							|| accountName.contains("florida blue")
							|| (currentName.contains("blue cross") 
							&& currentName.contains("florida")) 
							|| currentName.contains("florida blue")
							|| (titleName.contains("blue cross") 
							&& titleName.contains("florida")) 
							|| titleName.contains("florida blue")) {
						domainName = "bcbsfl.com";
					} else if ((accountName.contains("blue cross") 
							&& accountName.contains("north carolina")) 
							|| accountName.contains("bcbsnc")
							|| (currentName.contains("blue cross") 
							&& currentName.contains("north carolina")) 
							|| currentName.contains("bcbsnc")
							|| (titleName.contains("blue cross") 
							&& titleName.contains("north carolina")) 
							|| titleName.contains("bcbsnc")) {
						domainName = "bcbsnc.com";
					} else if (accountName.contains("bdp")
							|| currentName.contains("bdp")
							|| titleName.contains("bdp")) {
						domainName = "bdpint.com";
					} else if (accountName.contains("carecore")
							|| currentName.contains("carecore")
							|| titleName.contains("carecore")) {
						domainName = "carecorenational.com";
					} else if (accountName.contains("carefirst")
							|| currentName.contains("carefirst")
							|| titleName.contains("carefirst")) {
						domainName = "carefirst.com";
					} else if (accountName.contains("carmax")
							|| currentName.contains("carmax")
							|| titleName.contains("carmax")) {
						domainName = "carmax.com";
					} else if (accountName.contains("carnival")
							|| currentName.contains("carnival")
							|| titleName.contains("carnival")) {
						domainName = "carnival.com";
					} else if (accountName.contains("catalina")
							|| currentName.contains("catalina")
							|| titleName.contains("catalina")) {
						domainName = "catalinamarketing.com";
					} else if (accountName.contains("citrix")
							|| currentName.contains("citrix")
							|| titleName.contains("citrix")) {
						domainName = "citrix.com";
					} else if (accountName.contains("comscore")
							|| currentName.contains("comscore")
							|| titleName.contains("comscore")) {
						domainName = "comscore.com";
					} else if (accountName.contains("cox") 
							|| accountName.contains("autotrader") 
							|| accountName.contains("ready auto")
							|| currentName.contains("cox") 
							|| currentName.contains("autotrader") 
							|| currentName.contains("ready auto")
							|| titleName.contains("cox") 
							|| titleName.contains("autotrader") 
							|| titleName.contains("ready auto")) {
						domainName = "autotrader.com";
					} else if (accountName.contains("csx")
							|| currentName.contains("csx")
							|| titleName.contains("csx")) {
						domainName = "csx.com";
					} else if (accountName.contains("danaher")
							|| currentName.contains("danaher")
							|| titleName.contains("danaher")) {
						domainName = "danaher.com";
					} else if (accountName.contains("darden")
							|| currentName.contains("darden")
							|| titleName.contains("darden")) {
						domainName = "darden.com";
					} else if (accountName.contains("dell") 
							|| accountName.contains("secureworks")
							|| currentName.contains("dell") 
							|| currentName.contains("secureworks")
							|| titleName.contains("dell") 
							|| titleName.contains("secureworks")) {
						domainName = "dell.com";
					} else if (accountName.contains("delta")
							|| currentName.contains("delta")
							|| titleName.contains("delta")) {
						domainName = "delta.com";
					} else if (accountName.contains("dollar")
							|| currentName.contains("dollar")
							|| titleName.contains("dollar")) {
						domainName = "dollartree.com";
					} else if (accountName.contains("dominion")
							|| currentName.contains("dominion")
							|| titleName.contains("dominion")) {
						domainName = "dominionenterprises.com";
					} else if (accountName.contains("duke energy")
							|| currentName.contains("duke energy")
							|| titleName.contains("duke energy")) {
						domainName = "duke-energy.com";
					} else if (accountName.contains("dupont")
							|| currentName.contains("dupont")
							|| titleName.contains("dupont")) {
						domainName = "usa.dupont.com";
					} else if (accountName.contains("equifax")
							|| currentName.contains("equifax")
							|| titleName.contains("equifax")) {
						domainName = "equifax.com";
					} else if (accountName.contains("fanatics")
							|| currentName.contains("fanatics")
							|| titleName.contains("fanatics")) {
						domainName = "footballfanatics.com";
					} else if (accountName.contains("fannie")
							|| currentName.contains("fannie")
							|| titleName.contains("fannie")) {
						domainName = "fanniemae.com";
					} else if (accountName.contains("fidelity")
							|| currentName.contains("fidelity")
							|| titleName.contains("fidelity")) {
						domainName = "fnf.com";
					} else if (accountName.contains("fiserv")
							|| currentName.contains("fiserv")
							|| titleName.contains("fiserv")) {
						domainName = "fiserv.com";
					} else if (accountName.contains("florida power") 
							|| accountName.contains("fp&l")
							|| currentName.contains("florida power") 
							|| currentName.contains("fp&l")
							|| titleName.contains("florida power") 
							|| titleName.contains("fp&l")) {
						domainName = "fpl.com";
					} else if (accountName.contains("freddie")
							|| currentName.contains("freddie")
							|| titleName.contains("freddie")) {
						domainName = "freddiemac.com";
					} else if (accountName.contains("freedom")
							|| currentName.contains("freedom")
							|| titleName.contains("freedom")) {
						domainName = "freedommortgage.com";
					} else if (accountName.contains("fti")
							|| currentName.contains("fti")
							|| titleName.contains("fti")) {
						domainName = "fticonsulting.com";
					} else if (accountName.contains("geico")
							|| currentName.contains("geico")
							|| titleName.contains("geico")) {
						domainName = "geico.com";
					} else if (accountName.contains("general dynamic") 
							|| accountName.contains("gd")
							|| currentName.contains("general dynamic") 
							|| currentName.contains("gd")
							|| titleName.contains("general dynamic") 
							|| titleName.contains("gd")) {
						domainName = "gdit.com";
					} else if (accountName.contains("general electric") 
							|| accountName.contains("ge appliances")
							|| accountName.contains("ge aviation")
							|| accountName.contains("ge digital")
							|| accountName.contains("ge capital")
							|| accountName.contains("ge energy")
							|| accountName.contains("ge healthcare")
							|| accountName.contains("ge oil")
							|| accountName.contains("ge power")
							|| accountName.contains("ge transportation")
							|| accountName.contains("ge global")
							|| currentName.contains("general electric") 
							|| currentName.contains("ge appliances")
							|| currentName.contains("ge aviation")
							|| currentName.contains("ge digital")
							|| currentName.contains("ge capital")
							|| currentName.contains("ge energy")
							|| currentName.contains("ge healthcare")
							|| currentName.contains("ge oil")
							|| currentName.contains("ge power")
							|| currentName.contains("ge transportation")
							|| currentName.contains("ge global")
							|| titleName.contains("general electric") 
							|| titleName.contains("ge appliances")
							|| titleName.contains("ge aviation")
							|| titleName.contains("ge digital")
							|| titleName.contains("ge capital")
							|| titleName.contains("ge energy")
							|| titleName.contains("ge healthcare")
							|| titleName.contains("ge oil")
							|| titleName.contains("ge power")
							|| titleName.contains("ge transportation")
							|| titleName.contains("ge global")) {
						domainName = "ge.com";
					} else if (accountName.contains("genworth")
							|| currentName.contains("genworth")
							|| titleName.contains("genworth")) {
						domainName = "genworth.com";
					} else if (accountName.contains("harris")
							|| currentName.contains("harris")
							|| titleName.contains("harris")) {
						domainName = "harris.com";
					} else if (accountName.contains("health e")
							|| currentName.contains("health e")
							|| titleName.contains("health e")) {
						domainName = "healthesystems.com";
					} else if (accountName.contains("hilton")
							|| currentName.contains("hilton")
							|| titleName.contains("hilton")) {
						domainName = "hilton.com";
					} else if (accountName.contains("ihg") 
							|| accountName.contains("intercontinental hotels")
							|| currentName.contains("ihg") 
							|| currentName.contains("intercontinental hotels")
							|| titleName.contains("ihg") 
							|| titleName.contains("intercontinental hotels")) {
						domainName = "ihg.com";
					} else if (accountName.contains("ims health")
							|| currentName.contains("ims health")
							|| titleName.contains("ims health")) {
						domainName = "imshealth.com";
					} else if (accountName.contains("inovalon")
							|| currentName.contains("inovalon")
							|| titleName.contains("inovalon")) {
						domainName = "inovalon.com";
					} else if (accountName.contains("jm family") 
							|| accountName.contains("jmfamily")
							|| currentName.contains("jm family") 
							|| currentName.contains("jmfamily")
							|| titleName.contains("jm family") 
							|| titleName.contains("jmfamily")) {
						domainName = "jmfamily.com";
					} else if (accountName.contains("labcorp")
							|| currentName.contains("labcorp")
							|| titleName.contains("labcorp")) {
						domainName = "labcorp.com";
					} else if (accountName.contains("lincoln financial") 
							|| accountName.contains("lfg")
							|| currentName.contains("lincoln financial") 
							|| currentName.contains("lfg")
							|| titleName.contains("lincoln financial") 
							|| titleName.contains("lfg")) {
						domainName = "lfg.com";
					} else if (accountName.contains("lowe's") 
							|| accountName.contains("lowes")
							|| currentName.contains("lowe's") 
							|| currentName.contains("lowes")
							|| titleName.contains("lowe's") 
							|| titleName.contains("lowes")) {
						domainName = "lowes.com";
					} else if (accountName.contains("macy's") 
							|| accountName.contains("macys")
							|| currentName.contains("macy's") 
							|| currentName.contains("macys")
							|| titleName.contains("macy's") 
							|| titleName.contains("macys")) {
						domainName = "macys.com";
					} else if (accountName.contains("manhattan")
							|| currentName.contains("manhattan")
							|| titleName.contains("manhattan")) {
						domainName = "manh.com";
					} else if (accountName.contains("markel")
							|| currentName.contains("markel")
							|| titleName.contains("markel")) {
						domainName = "markelcorp.com";
					} else if (accountName.contains("marriott")
							|| currentName.contains("marriott")
							|| titleName.contains("marriott")) {
						domainName = "marriott.com";
					} else if (accountName.contains("mars")
							|| currentName.contains("mars")
							|| titleName.contains("mars")) {
						domainName = "effem.com";
					} else if (accountName.contains("masonite")
							|| currentName.contains("masonite")
							|| titleName.contains("masonite")) {
						domainName = "masonite.com";
					} else if (accountName.contains("mckesson")
							|| currentName.contains("mckesson")
							|| titleName.contains("mckesson")) {
						domainName = "mckesson.com";
					} else if (accountName.contains("merck")
							|| currentName.contains("merck")
							|| titleName.contains("merck")) {
						domainName = "merck.com";
					} else if (accountName.contains("microstrategy")
							|| currentName.contains("microstrategy")
							|| titleName.contains("microstrategy")) {
						domainName = "microstrategy.com";
					} else if (accountName.contains("moffitt")
							|| currentName.contains("moffitt")
							|| titleName.contains("moffitt")) {
						domainName = "moffitt.org";
					} else if (accountName.contains("mohawkind")
							|| currentName.contains("mohawkind")
							|| titleName.contains("mohawkind")) {
						domainName = "mohawkind.com";
					} else if (accountName.contains("nascar")
							|| currentName.contains("nascar")
							|| titleName.contains("nascar")) {
						domainName = "nascar.com";
					} else if (accountName.contains("under armour")
							|| accountName.contains("mapmyfitness")
							|| currentName.contains("under armour")
							|| currentName.contains("mapmyfitness")
							|| titleName.contains("under armour")
							|| titleName.contains("mapmyfitness")) {
						domainName = "underarmour.com";
					} else if (accountName.contains("navy federal") 
							|| accountName.contains("nfcu")
							|| currentName.contains("navy federal") 
							|| currentName.contains("nfcu")
							|| titleName.contains("navy federal") 
							|| titleName.contains("nfcu")) {
						domainName = "navyfederal.org";
					} else if (accountName.contains("ncr")
							|| currentName.contains("ncr")
							|| titleName.contains("ncr")) {
						domainName = "ncr.com";
					} else if (accountName.contains("neustar")
							|| currentName.contains("neustar")
							|| titleName.contains("neustar")) {
						domainName = "neustar.biz";
					} else if (accountName.contains("nielsen")
							|| currentName.contains("nielsen")
							|| titleName.contains("nielsen")) {
						domainName = "nielsen.com";
					} else if (accountName.contains("northrop")
							|| currentName.contains("northrop")
							|| titleName.contains("northrop")) {
						domainName = "ngc.com";
					} else if (accountName.contains("office depot")
							|| currentName.contains("office depot")
							|| titleName.contains("office depot")) {
						domainName = "officedepot.com";
					} else if (accountName.contains("pra intl") 
							|| accountName.contains("pra international")
							|| accountName.contains("pra group")
							|| currentName.contains("pra intl") 
							|| currentName.contains("pra international")
							|| currentName.contains("pra group")
							|| titleName.contains("pra intl") 
							|| titleName.contains("pra international")
							|| titleName.contains("pra group")) {
						domainName = "praintl.com";
					} else if (accountName.contains("publix")
							|| currentName.contains("publix")
							|| titleName.contains("publix")) {
						domainName = "publix.com";
					} else if (accountName.contains("raymond")
							|| currentName.contains("raymond")
							|| titleName.contains("raymond")) {
						domainName = "raymondjames.com";
					} else if (accountName.contains("roper") 
							|| accountName.contains("transcore")
							|| currentName.contains("roper") 
							|| currentName.contains("transcore")
							|| titleName.contains("roper") 
							|| titleName.contains("transcore")) {
						domainName = "transcore.com";
					} else if (accountName.contains("caribbean")
							|| currentName.contains("caribbean")
							|| titleName.contains("caribbean")) {
						domainName = "rccl.com";
					} else if (accountName.contains("sas")
							|| currentName.contains("sas")
							|| titleName.contains("sas")) {
						domainName = "sas.com";
					} else if (accountName.contains("sita")
							|| currentName.contains("sita")
							|| titleName.contains("sita")) {
						domainName = "sita.aero";
					} else if (accountName.contains("subaru")
							|| currentName.contains("subaru")
							|| titleName.contains("subaru")) {
						domainName = "subaru.com";
					} else if (accountName.contains("sungard")
							|| currentName.contains("sungard")
							|| titleName.contains("sungard")) {
						domainName = "sungardas.com";
					} else if (accountName.contains("sykes")
							|| currentName.contains("sykes")
							|| titleName.contains("sykes")) {
						domainName = "sykes.com";
					} else if (accountName.contains("synchronoss")
							|| currentName.contains("synchronoss")
							|| titleName.contains("synchronoss")) {
						domainName = "synchronoss.com";
					} else if (accountName.contains("syniverse")
							|| currentName.contains("syniverse")
							|| titleName.contains("syniverse")) {
						domainName = "syniverse.com";
					} else if (accountName.contains("rowe")
							|| currentName.contains("rowe")
							|| titleName.contains("rowe")) {
						domainName = "troweprice.com";
					} else if (accountName.contains("chico")
							|| currentName.contains("chico")
							|| titleName.contains("chico")) {
						domainName = "chicos.com";
					} else if (accountName.contains("children's hospital of philadelphia") 
							|| accountName.contains("childrens hospital of philadelphia") 
							|| accountName.contains("chop")
							|| currentName.contains("children's hospital of philadelphia") 
							|| currentName.contains("childrens hospital of philadelphia") 
							|| currentName.contains("chop")
							|| titleName.contains("children's hospital of philadelphia") 
							|| titleName.contains("childrens hospital of philadelphia") 
							|| titleName.contains("chop")) {
						domainName = "email.chop.edu";
					} else if (accountName.contains("coca-cola") 
							|| accountName.contains("coca cola")
							|| currentName.contains("coca-cola") 
							|| currentName.contains("coca cola")
							|| titleName.contains("coca-cola") 
							|| titleName.contains("coca cola")) {
						domainName = "na.ko.com";
					} else if (accountName.contains("hershey")
							|| currentName.contains("hershey")
							|| titleName.contains("hershey")) {
						domainName = "hersheys.com";
					} else if (accountName.contains("home depot") 
							|| accountName.contains("thd")
							|| currentName.contains("home depot") 
							|| currentName.contains("thd")
							|| titleName.contains("home depot") 
							|| titleName.contains("thd")) {
						domainName = "homedepot.com";
					} else if (accountName.contains("vanguard")
							|| currentName.contains("vanguard")
							|| titleName.contains("vanguard")) {
						domainName = "vanguard.com";
					} else if (accountName.contains("tmg")
							|| currentName.contains("tmg")
							|| titleName.contains("tmg")) {
						domainName = "hcsc.net";
					} else if (accountName.contains("travelport")
							|| currentName.contains("travelport")
							|| titleName.contains("travelport")) {
						domainName = "travelport.com";
					} else if (accountName.contains("tsys")
							|| currentName.contains("tsys")
							|| titleName.contains("tsys")) {
						domainName = "tsys.com";
					} else if (accountName.contains("ultimate")
							|| currentName.contains("ultimate")
							|| titleName.contains("ultimate")) {
						domainName = "ultimatesoftware.com";
					} else if (accountName.contains("united guaranty")
							|| currentName.contains("united guaranty")
							|| titleName.contains("united guaranty")) {
						domainName = "ugcorp.com";
					} else if (accountName.contains("united parcel") 
							|| accountName.contains("ups")
							|| currentName.contains("united parcel") 
							|| currentName.contains("ups")
							|| titleName.contains("united parcel") 
							|| titleName.contains("ups")) {
						domainName = "ups.com";
					} else if (accountName.contains("verisign")
							|| currentName.contains("verisign")
							|| titleName.contains("verisign")) {
						domainName = "verisign.com";
					} else if (accountName.contains("vertex")
							|| currentName.contains("vertex")
							|| titleName.contains("vertex")) {
						domainName = "vertexinc.com";
					} else if (accountName.contains("wake forest")
							|| currentName.contains("wake forest")
							|| titleName.contains("wake forest")) {
						domainName = "wfu.edu";
					} else if (accountName.contains("wellcare")
							|| currentName.contains("wellcare")
							|| titleName.contains("wellcare")) {
						domainName = "wellcare.com";
					} else {
						domainName = "NONE FOUND";
						domainsNotFound++;
					}
					// Populate company domain
					domainCell.setCellValue(domainName);
					
					// Update the HashMap with number of domain names identified
					map.put(domainName, map.get(domainName) + 1);
					
					// Create email addresses
					Cell emailCell = row.createCell(nameColumnIndex + 12);
					emailCell.setCellStyle(styleBlueWhite);
					String email = null;
					switch (domainName)
					{
						// Cases with FirstInitial + LastName@domainName
						case "aarp.org":
						case "aflac.com":
						case "amerisourcebergen.com":
						case "avidxchange.com":
						case "bbandt.com":
						case "carecorenational.com":
						case "carnival.com":
						case "comscore.com":
						case "darden.com":
						case "footballfanatics.com":
						case "healthesystems.com":
						case "imshealth.com":
						case "inovalon.com":
						case "manh.com":
						case "markelcorp.com":
						case "microstrategy.com":
						case "underarmour.com":
						case "rccl.com":
						case "subaru.com":
						case "na.ko.com":
						case "hersheys.com":
						case "tsys.com":
						case "ups.com":
						case "verisign.com":
						case "masonite.com":
						case "dollartree.com":
						case "geico.com":
						case "nascar.com":
							nonEmptyRows++;
							email = firstName.substring(0, 1) + lastName + "@" + domainName; 
							break;
							
						// Cases with FirstName.MiddleInitial(if available).LastName@domainName
						case "delta.com":
						case "lowes.com":
							nonEmptyRows++;
							if (middleName != null) {
								email = firstName + "." + middleName.substring(0, 1) + lastName + "@" + domainName;
							} else {
								email = firstName + "." + lastName + "@" + domainName;
							}
							break;
						
						// Cases with FirstName.LastName@domainName
						case "advance-auto.com":
						case "ahss.org":
						case "altisource.com":
						case "astrazeneca.com":
						case "baycare.org":
						case "bdpinternational.com":
						case "benefitfocus.com":
						case "blackbaud.com":
						case "blackboard.com":
						case "bcbsfl.com":
						case "bcbsnc.com":
						case "carefirst.com":
						case "catalinamarketing.com":
						case "chicos.com":
						case "citrix.com":
						case "autotrader.com":
						case "danaher.com":
						case "dominionenterprises.com":
						case "duke-energy.com":
						case "usa.dupont.com":
						case "equifax.com":
						case "fnf.com":
						case "fiserv.com":
						case "fpl.com":
						case "freedommortgage.com":
						case "fticonsulting.com":
						case "gdit.com":
						case "ge.com":
						case "genworth.com":
						case "harris.com":
						case "hilton.com":
						case "ihg.com":
						case "jmfamily.com":
						case "lfg.com":
						case "macys.com":
						case "marriott.com":
						case "effem.com":
						case "mckesson.com":
						case "moffitt.org":
						case "ncr.com":
						case "neustar.biz":
						case "nielsen.com":
						case "ngc.com":
						case "officedepot.com":
						case "publix.com":
						case "raymondjames.com":
						case "sas.com":
						case "sita.aero":
						case "sungardas.com":
						case "sykes.com":
						case "synchronoss.com":
						case "syniverse.com":
						case "transcore.com":
						case "travelport.com":
						case "ugcorp.com":
						case "vertexinc.com":
						case "wellcare.com":
							nonEmptyRows++;
							email = firstName + "." + lastName + "@" + domainName; 
							break;
														
						// Cases with LastName + FirstInitial@domainName
						case "autonation.com":
						case "email.chop.edu":
						case "wfu.edu":
							nonEmptyRows++;
							email = lastName + firstName.substring(0, 1) + "@" + domainName;
							break;
							
						// Cases with LastName + FirstName@domainName
						case "praintl.com":
							nonEmptyRows++;
							email = lastName + firstName + "@" + domainName;
							break;
							
						// Cases with FirstName_LastName@domainName
						case "carmax.com":
						case "csx.com":
						case "dell.com":
						case "fanniemae.com":
						case "freddiemac.com":
						case "merck.com":
						case "mohawkind.com":
						case "navyfederal.org":
						case "troweprice.com":
						case "homedepot.com":
						case "ultimatesoftware.com":
						case "vanguard.com":
						case "hcsc.net":
							nonEmptyRows++;
							email = firstName + "_" + lastName + "@" + domainName;
							break;
							
						// Cases with first 6 letters of LastName + FirstInitial@domainName
						case "labcorp.com":
							nonEmptyRows++;
							if (lastName == null) {
								email = firstName + "@" + domainName;
							} else if (lastName.length() > 6) {
								email = lastName.substring(0, 6) + firstName.substring(0, 1) + "@" + domainName;
							} else {
								email = lastName + firstName.substring(0, 1) + "@" + domainName;
							}
							break;
						
						// Default case
						default:
							System.out.println("UNABLE TO ID DOMAIN FOR: 'company': " + accountName + 
									"  'current': " + currentName + "   'title': " + titleName); 
					}
					emailCell.setCellValue(email);
				}
				
				// Format success percentage
				DecimalFormat df = new DecimalFormat("#.##");
				df.setRoundingMode(RoundingMode.HALF_UP);
				
				// Sort map of count of successful domain matches
				Map<String, Integer> ascendingMap = sortByValue(map);
				
				// Log stats to the console
				int goodContacts = nonEmptyRows 
						- linkedinMembers - duplicateContacts;
				System.out.println();
				System.out.println("Successfully matched domain/email for " 
						+ df.format(100 - (((double) domainsNotFound 
								/ (goodContacts + domainsNotFound)) * 100)) 
								+ "% of contacts (" + goodContacts + " matched with "
								+ domainsNotFound + " not matched).");
				System.out.println("Removed " + linkedinMembers + " 'LinkedIn Members', " 
						+ emptyRows + " empty rows and " + duplicateContacts 
						+ " duplicate entries.");
				
				// Log total number of results per account
				System.out.println();
				for (String name : ascendingMap.keySet()) {
					String key = name;
					String value = ascendingMap.get(name).toString();
					System.out.println(key + ": " + value);
				}
				
				// Resize new columns to fit data
				sheet.autoSizeColumn(nameColumnIndex + 8);
				sheet.autoSizeColumn(nameColumnIndex + 9);
				sheet.autoSizeColumn(nameColumnIndex + 10);
				sheet.autoSizeColumn(nameColumnIndex + 11);
				sheet.autoSizeColumn(nameColumnIndex + 12);
	        
				// Create new file from input file data
		        try {
		            FileOutputStream output = new FileOutputStream("email-builder-output.xlsx");
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
	
	// Helper method to sort Map by value from smallest to largest
	public static <K, V extends Comparable<? super V>> Map<K, V> sortByValue(Map<K, V> map) {
	    Map<K, V> result = new LinkedHashMap<>();
	    Stream<Map.Entry<K, V>> st = map.entrySet().stream();
	
	    st.sorted(Map.Entry.comparingByValue())
	        .forEachOrdered(e -> result.put(e.getKey(), e.getValue()));
	
	    return result;
	}
	
}
