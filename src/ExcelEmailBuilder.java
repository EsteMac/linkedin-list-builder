import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
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
				Sheet dataMinerSheet = workbook.getSheetAt(0);
				Sheet accountsSheet = workbook.getSheetAt(1);
				Sheet summarySheet = workbook.createSheet("Summary");
				
				// Strings used to build email addresses
				String fullName;
				String firstName;
				String middleName;
				String lastName;
				String accountName;
				String currentName;
				String titleName;
				String domainName = null;
				
				// Cells to populate
				Cell domainCell;
				Cell emailCell;
				
				// HashMap to store the number of unique domain / email identified per named account
			    Map<String, Integer> map = new HashMap<String, Integer>();
				for (int r = 1; r <= accountsSheet.getLastRowNum(); r++) {
					Row row = accountsSheet.getRow(r);
					Cell companyDomainCell = row.getCell(0);
					map.put(companyDomainCell.toString(), 0);
				}
				
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
				Row titleRow = dataMinerSheet.getRow(0);
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
				for (int r = 1; r <= dataMinerSheet.getLastRowNum(); r++) {
					Row row = dataMinerSheet.getRow(r);
					// Check for empty rows
				    if(row == null){
				        dataMinerSheet.shiftRows(r + 1, dataMinerSheet.getLastRowNum(), -1);
				        emptyRows++;
				        r--;
				        continue;
				    }
				}
				
				// Remove "LinkedIn Member" rows
				for (int r = 1; r <= dataMinerSheet.getLastRowNum(); r++) {
					Row row = dataMinerSheet.getRow(r);
					// Check for entries with names listed as "LinkedIn" and remove them
					Cell linkedinCell = row.getCell(nameColumnIndex);
					if (linkedinCell.toString().toLowerCase().contains("linkedin")) {
						linkedinMembers++;
						System.out.println("NON-CONNECTED 'LinkedIn Member' ENTRY REMOVED: " + linkedinCell.toString());
						if (r != dataMinerSheet.getLastRowNum()) {
							// Clear row with no name or connections
							dataMinerSheet.removeRow(row);
							// Remove cleared empty row
							dataMinerSheet.shiftRows(r + 1, dataMinerSheet.getLastRowNum(), -1);
						} else {
							dataMinerSheet.removeRow(row);
						}
						// Reset row counter
						r--;
						continue;
					}
				}
				
				// Remove duplicate rows
				for (int r = 1; r <= dataMinerSheet.getLastRowNum(); r++) {
					Row row = dataMinerSheet.getRow(r);
					// Check for duplicate entries and remove them
					Cell checkCell = row.getCell(nameColumnIndex);
					Cell currentAdjacentCell = row.getCell(nameColumnIndex + 1);
					for (int c = (r + 1); c <= dataMinerSheet.getLastRowNum(); c++) {
						Row compareRow = dataMinerSheet.getRow(c);
						Cell compareCell = compareRow.getCell(nameColumnIndex);
						Cell compareAdjacentCell = compareRow.getCell(nameColumnIndex + 1);
						// Compare content of 2 cells per contact to ensure it's a duplicate
						if ((checkCell.toString() == compareCell.toString())
								&& (currentAdjacentCell.toString() == compareAdjacentCell.toString())) {
							duplicateContacts++;
							System.out.println("DUPLICATE ENTRY REMOVED: " + compareCell.toString());
							// Clear duplicate row
							dataMinerSheet.removeRow(compareRow);
							if (dataMinerSheet.getRow(c + 1) != null) {
								// Remove cleared empty row
								dataMinerSheet.shiftRows(c + 1, dataMinerSheet.getLastRowNum(), -1);
							}
							// Reset row counter
							r--;
						} 
					}
				}
				
				// Iterate row by row and assign domain / email addresses
				for (int r = 1; r <= dataMinerSheet.getLastRowNum(); r++) {
					Row row = dataMinerSheet.getRow(r);		
				    
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
					domainCell = row.createCell(nameColumnIndex + 11);
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
					
					// Search across rows in the input data sheet to find a match
					search:
						for (int i = 1; i <= accountsSheet.getLastRowNum(); i++) {
							Row accountRow = accountsSheet.getRow(i);
							domainName = null;
							// for each row, iterate across account name columns to find a match
							for(int accountColumn = 2; accountColumn < accountRow.getLastCellNum(); accountColumn++) {
								Cell accountNameCell = accountRow.getCell(accountColumn);
								String accountNameString = accountNameCell.toString().toLowerCase();
								// check for boolean operator AND in the input sheet account name cells
								if (accountNameString.contains(" and ")) {
									String[] accountBooleanSplit = accountNameString.split(" and ", 2);
									String leftAccount = accountBooleanSplit[0];
									String rightAccount = accountBooleanSplit[1];
									if ((accountName.contains(leftAccount) && accountName.contains(rightAccount))
											|| (currentName.contains(leftAccount) && currentName.contains(rightAccount))
											|| (titleName.contains(leftAccount) && titleName.contains(rightAccount))) {
										domainName = accountRow.getCell(0).toString();
										break search;
									}
								} else if (accountName.contains(accountNameString)
											|| currentName.contains(accountNameString)
											|| titleName.contains(accountNameString)) {
										domainName = accountRow.getCell(0).toString();
										break search;
								} else {
									continue;
								}
							}
						}
					// Check for matches unable to make
					if (domainName == null) {
						domainName = "NONE FOUND";
						domainsNotFound++;
					}
					// Populate company domain
					domainCell.setCellValue(domainName);
					
					// Update the HashMap with number of domain names identified
					if (!domainName.contains("NONE FOUND")) {
						map.put(domainName, map.get(domainName) + 1);
					}
					
					// Create email addresses
					emailCell = row.createCell(nameColumnIndex + 12);
					emailCell.setCellStyle(styleBlueWhite);
					String email = null;
					switch (domainName)
					{
						// Cases with FirstInitial + LastName@domainName
						case "aarp.org":
						case "aflac.com":
						case "amerisourcebergen.com":
						case "audiusa.com":
						case "avidxchange.com":
						case "bbandt.com":
						case "carecorenational.com":
						case "carnival.com":
						case "chubb.com":
						case "comscore.com":
						case "darden.com":
						case "ebay.com":
						case "footballfanatics.com":
						case "hanloninvest.com":
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
						case "sbgnet.com":
						case "southernco.com":
						case "tsys.com":
						case "ups.com":
						case "verisign.com":
						case "masonite.com":
						case "seic.com":
						case "dollartree.com":
						case "geico.com":
						case "nascar.com":
						case "urbanout.com":
						case "wlgore.com":
							nonEmptyRows++;
							email = firstName.substring(0, 1) + lastName + "@" + domainName; 
							break;
							
						// Cases with FirstName.MiddleInitial(if available).LastName@domainName
						case "delta.com":
						case "lowes.com":
						case "gsk.com":
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
						case "ellucian.com":
						case "equifax.com":
						case "fmc.com":
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
						case "iassoftware.com":
						case "ibx.com":
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
						case "qvc.com":
						case "raymondjames.com":
						case "rovicorp.com":
						case "sas.com":
						case "siemens.com":
						case "sig.com":
						case "sita.aero":
						case "sungardas.com":
						case "sykes.com":
						case "synchronoss.com":
						case "syniverse.com":
						case "towerswatson.com":
						case "transcore.com":
						case "travelport.com":
						case "tokiom.com":
						case "tycoelectronics.com":
						case "ugcorp.com":
						case "vertexinc.com":
						case "vw.com":
						case "wawa.com":
						case "wellcare.com":
							nonEmptyRows++;
							email = firstName + "." + lastName + "@" + domainName; 
							break;
							
						// Cases with LastName.FirstName@domainName
						case "endo.com":
							nonEmptyRows++;
							email = lastName + "." + firstName + "@" + domainName; 
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
							
						// Cases with LastName-FirstName@domain.com
						case "aramark.com":
							nonEmptyRows++;
							email = lastName + "-" + firstName + "@" + domainName;
							break;
							
						// Cases with first 6 letters of LastName + FirstInitial@domainName
						case "labcorp.com":
						case "slhn.org":
							nonEmptyRows++;
							if (lastName == null) {
								email = firstName + "@" + domainName;
							} else if (lastName.length() > 6) {
								email = lastName.substring(0, 6) + firstName.substring(0, 1) + "@" + domainName;
							} else {
								email = lastName + firstName.substring(0, 1) + "@" + domainName;
							}
							break;
							
						// Cases with first 6 letters of LastName + FirstInitial + MiddleInitial@domainName
						case "airproducts.com":
							nonEmptyRows++;
							if (lastName == null) {
								email = firstName + "@" + domainName;
							} else if (lastName.length() > 6 && middleName != null) {
								email = lastName.substring(0, 6) + firstName.substring(0, 1) + middleName.substring(0, 1) + "@" + domainName;
							} else if (lastName.length() > 6 && middleName == null) {
								email = lastName.substring(0, 6) + firstName.substring(0, 1) + "@" + domainName;
							} else if (lastName.length() <= 6 && middleName != null) {
								email = lastName + firstName.substring(0, 1) + middleName.substring(0, 1) + "@" + domainName;
							} else if (lastName.length() <= 6 && middleName == null) {
								email = lastName + firstName.substring(0, 1) + "@" + domainName;
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
				
				// Sort map of count of successful domain matches (from largest to smallest)
				Map<String, Integer> descendingMap = sortByValue(map);
				
				// Log stats to the console
				int goodContacts = nonEmptyRows 
						- linkedinMembers - duplicateContacts;
				System.out.println();
				System.out.println("Successfully matched domain/email for " 
						+ df.format(100 - (((double) domainsNotFound 
								/ (goodContacts + domainsNotFound)) * 100)) 
								+ "% of contacts (" + goodContacts + " matched with "
								+ domainsNotFound + " not matched)");
				System.out.println("Removed " + emptyRows + " empty rows, " 
						+ duplicateContacts + " duplicate entries, and " 
						+ linkedinMembers + " 'LinkedIn Members'");
				
				// Log total number of results per account (from largest to smallest)
				System.out.print("Top accounts for this search: ");
				for (String name : descendingMap.keySet()) {
					String key = name;
					int valueInt = descendingMap.get(name);
					String valueString = descendingMap.get(name).toString();
					if (valueInt != 0 && !key.contains("NONE FOUND")) {
						System.out.print(key + ": " + valueString + " --> ");
					}
				}
				
				// Resize new columns on main sheet to fit data
				dataMinerSheet.autoSizeColumn(nameColumnIndex + 8);
				dataMinerSheet.autoSizeColumn(nameColumnIndex + 9);
				dataMinerSheet.autoSizeColumn(nameColumnIndex + 10);
				dataMinerSheet.autoSizeColumn(nameColumnIndex + 11);
				dataMinerSheet.autoSizeColumn(nameColumnIndex + 12);
				
				// Create the title row for summary sheet
				Row summaryTitleRow = summarySheet.createRow(0);
				summaryTitleRow.createCell(0).setCellValue("Domain");
				summaryTitleRow.createCell(1).setCellValue("Total Found");
				
				// In summary sheet enter total number of results per account (from largest to smallest)
				Iterator<Entry<String, Integer>> it = descendingMap.entrySet().iterator();
				for (int r = 1; it.hasNext(); r++) {
					@SuppressWarnings("rawtypes")
					Map.Entry pair = (Map.Entry) it.next();
					Row row = summarySheet.createRow(r);
					
					// No need to rank NONE FOUND
				    if(pair.getKey().toString().contains("NONE FOUND")){
				        r--;
				        continue;
				    }
				    
					Cell domainSummaryCell = row.createCell(0);
					domainSummaryCell.setCellValue(pair.getKey().toString());
					
					Cell totalCell = row.createCell(1);
					totalCell.setCellValue((int) pair.getValue());
				}
				
				// Resize columns on summary sheet to fit data
				summarySheet.autoSizeColumn(0);
				summarySheet.autoSizeColumn(1);
	        
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
	
	// Helper method to sort Map by value from largest to smallest
	public static <K, V extends Comparable<? super V>> Map<K, V> sortByValue(Map<K, V> map) {
	    Map<K, V> result = new LinkedHashMap<>();
	    Stream<Map.Entry<K, V>> st = map.entrySet().stream();
	    st.sorted(Map.Entry.comparingByValue(Comparator.reverseOrder()))
	        .forEachOrdered(e -> result.put(e.getKey(), e.getValue()));
	    return result;
	}
	
}
