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
				Sheet marketoSheet = workbook.createSheet("Marketo");
				
				// Constants to populate for every contact in marketo tab
				final String country = accountsSheet.getRow(0).getCell(1).toString();
				final String originalLeadSource = accountsSheet.getRow(1).getCell(1).toString();
				final String originalLeadSourceDescription = accountsSheet.getRow(2).getCell(1).toString();
				final String mostRecentLeadSource = accountsSheet.getRow(3).getCell(1).toString();
				final String mostRecentLeadSourceDescription = accountsSheet.getRow(4).getCell(1).toString();
				final String agile = accountsSheet.getRow(0).getCell(4).toString();
				final String productData = accountsSheet.getRow(1).getCell(4).toString();
				final String paas = accountsSheet.getRow(2).getCell(4).toString();
				final String labs = accountsSheet.getRow(3).getCell(4).toString();
				final String pws = accountsSheet.getRow(4).getCell(4).toString();
				
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
			    Map<String, Integer> resultsDistributionMap = new HashMap<String, Integer>();
				for (int r = 1; r <= accountsSheet.getLastRowNum(); r++) {
					Row row = accountsSheet.getRow(r);
					Cell companyDomainCell = row.getCell(1);
					resultsDistributionMap.put(companyDomainCell.toString(), 0);
				}
				
				// HashMap to store the company names associated with domains
			    Map<String, String> domainAccountMap = new HashMap<String, String>();
				for (int r = 1; r <= accountsSheet.getLastRowNum(); r++) {
					Row row = accountsSheet.getRow(r);
					String companyDomainCell = row.getCell(1).toString();
					String companyNameCell = row.getCell(0).toString();
					domainAccountMap.put(companyDomainCell, companyNameCell);
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
				
				// Keeps track of contacts with good data that were also successfully matched with an email / domain
				int goodContacts = 0;
				
				// Keeps track of # of empty rows
				int emptyRows = 0;
				
				// Keeps track of # of unidentified domain names
				int domainsNotFound = 0;
				
				// Keeps track of # of rows without any mutual remote connections
				// where the name shows up as "LinkedIn Member"
				int linkedinMembers = 0;
				
				// Keeps track of # of rows where the last name is missing
				int nameInclomplete = 0;
				
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
				System.out.println("REMOVING 'LinkedIn Member' ENTRIES...");
				for (int r = 1; r <= dataMinerSheet.getLastRowNum(); r++) {
					Row row = dataMinerSheet.getRow(r);
					// Check for entries with names listed as "LinkedIn" and remove them
					Cell linkedinCell = row.getCell(nameColumnIndex);
					if (linkedinCell.toString().toLowerCase().contains("linkedin")) {
						linkedinMembers++;
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
				System.out.println("REMOVING DUPLICATE ENTRIES...");
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
							System.out.println("REMOVED DUPLICATE ENTRY: " + compareCell.toString());
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
					
					// HashMap to store the email structure type per named account
				    Map<String, Integer> emailStructureMap = new HashMap<String, Integer>();
				    
					// Search across rows in the input data sheet to find a match
					search:
						for (int i = 6; i <= accountsSheet.getLastRowNum(); i++) {
							Row accountRow = accountsSheet.getRow(i);
							domainName = null;
							// for each row, iterate across account name columns to find a match
							for(int c = 3; c < accountRow.getLastCellNum(); c++) {
								Cell accountNameCell = accountRow.getCell(c);
								if (cellIsEmpty(accountNameCell)) {
									// No need to continue to check empty cells if there 
									// are no empty cells between Account Name cells
									// Move on and check the next row of account names
									break;
								}
								String accountNameString = accountNameCell.toString().toLowerCase();
								// check for boolean operator AND in the input sheet account name cells
								if (accountNameString.contains(" and ")) {
									if (splitCellCheck(accountNameString, accountName, currentName, titleName) == true) {
										domainName = accountRow.getCell(1).toString();
										// This cell associates the correct email format structure with each domain
										// which is stored in the emailStructureMap 
										Cell emailTypeCell = accountRow.getCell(2);
										emailStructureMap.put(domainName, (int) emailTypeCell.getNumericCellValue());
										// Success! Exit this loop for this once contact check
										break search;
									};
								} else if (cellCheck(accountNameString, accountName, currentName, titleName) == true) {
										domainName = accountRow.getCell(1).toString();
										// This cell associates the correct email format structure with each domain
										// which is stored in the emailStructureMap 
										Cell emailTypeCell = accountRow.getCell(2);
										emailStructureMap.put(domainName, (int) emailTypeCell.getNumericCellValue());
										// Success! Exit this loop for this once contact check
										break search;
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
					if (!domainName.contains("NONE FOUND") && !(lastName == null) 
							&& !(lastName.length() == 1) && !(firstName.length() == 1)) {
						resultsDistributionMap.put(domainName, resultsDistributionMap.get(domainName) + 1);
					}
					
					// Create email addresses
					emailCell = row.createCell(nameColumnIndex + 12);
					emailCell.setCellStyle(styleBlueWhite);
					String email = null;
					
					// Build email addresses based on email structure types stored in emailStructuremap		
					if (!domainName.contains("NONE FOUND")) {
						switch (emailStructureMap.get(domainName))
						{
							// Cases with FirstInitial + LastName@domainName
							case 1:
								email = firstName.substring(0, 1) + lastName + "@" + domainName; 
								break;
								
							// Cases with FirstName.MiddleInitial(if available).LastName@domainName
							case 2:
								if (middleName != null) {
									email = firstName + "." + middleName.substring(0, 1) + "." + lastName + "@" + domainName;
								} else {
									email = firstName + "." + lastName + "@" + domainName;
								}
								break;
							
							// Cases with FirstName.LastName@domainName
							case 3:
								email = firstName + "." + lastName + "@" + domainName; 
								break;
								
							// Cases with LastName.FirstName@domainName
							case 4:
								email = lastName + "." + firstName + "@" + domainName; 
								break;							
															
							// Cases with LastName + FirstInitial@domainName
							case 5:
								email = lastName + firstName.substring(0, 1) + "@" + domainName;
								break;
								
							// Cases with LastName + FirstName@domainName
							case 6:
								email = lastName + firstName + "@" + domainName;
								break;
								
							// Cases with FirstName_LastName@domainName
							case 7:
								email = firstName + "_" + lastName + "@" + domainName;
								break;
								
							// Cases with LastName-FirstName@domain.com
							case 8:
								email = lastName + "-" + firstName + "@" + domainName;
								break;
								
							// Cases with first 6 letters of LastName + FirstInitial@domainName
							case 9:
								if (lastName == null) {
									email = firstName + "@" + domainName;
								} else if (lastName.length() > 6) {
									email = lastName.substring(0, 6) + firstName.substring(0, 1) + "@" + domainName;
								} else {
									email = lastName + firstName.substring(0, 1) + "@" + domainName;
								}
								break;
								
							// Cases with first 6 letters of LastName + FirstInitial + MiddleInitial@domainName
							case 10:
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
								
							// Cases with FirstName_MiddleInitial(if available)_LastName@domainName
							case 11:
								if (middleName != null) {
									email = firstName + "_" + middleName.substring(0, 1) + "_" + lastName + "@" + domainName;
								} else {
									email = firstName + "_" + lastName + "@" + domainName;
								}
								break;
							
							// Cases with first 6 letters of LastName + FirstInitial@domainName
							case 12:
								if (lastName == null) {
									email = firstName + "@" + domainName;
								} else if (lastName.length() > 6) {
									email = lastName.substring(0, 6) + firstName.substring(0, 1) + "@" + domainName;
								} else if (lastName.length() <= 6) {
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
					}
					emailCell.setCellValue(email);
					
					// Remove rows with incomplete names
					if (lastName == null || lastName.length() == 1 || firstName.length() == 1) {
						nameInclomplete++;
						if (r == 1) {
							System.out.println("REMOVING ENTRIES WITH INCOMPLETE NAMES...");
						}
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
				
				// Sort map of count of successful domain matches (from largest to smallest)
				Map<String, Integer> descendingMap = sortByValue(resultsDistributionMap);
				
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
				
				// Create the title row for the marketo sheet
				Row marketoTitleRow = marketoSheet.createRow(0);
				marketoTitleRow.createCell(0).setCellValue("First Name");
				marketoTitleRow.createCell(1).setCellValue("Last Name");
				marketoTitleRow.createCell(2).setCellValue("Company Name");
				marketoTitleRow.createCell(3).setCellValue("Email Address");
				marketoTitleRow.createCell(4).setCellValue("Phone Number");
				marketoTitleRow.createCell(5).setCellValue("Job Title");
				marketoTitleRow.createCell(6).setCellValue("City");
				marketoTitleRow.createCell(7).setCellValue("State");
				marketoTitleRow.createCell(8).setCellValue("Postal Code");
				marketoTitleRow.createCell(9).setCellValue("Country");
				marketoTitleRow.createCell(10).setCellValue("Original Lead Source Description");
				marketoTitleRow.createCell(11).setCellValue("Original Lead Source");
				marketoTitleRow.createCell(12).setCellValue("Most Recent Lead Source");
				marketoTitleRow.createCell(13).setCellValue("Most Recent Lead Source Description");
				marketoTitleRow.createCell(14).setCellValue("Lead Action");
				marketoTitleRow.createCell(15).setCellValue("Agile");
				marketoTitleRow.createCell(16).setCellValue("Product: Data");
				marketoTitleRow.createCell(17).setCellValue("PaaS");
				marketoTitleRow.createCell(18).setCellValue("Labs");
				marketoTitleRow.createCell(19).setCellValue("PWS");
				
				// Populate Marketo sheet
				System.out.println();
				System.out.println("POPULATING MARKETO SHEET...");
				int rowDelta = 0;
				for (int r = 1; r <= dataMinerSheet.getLastRowNum(); r++) {
					Row dataMinerRow = dataMinerSheet.getRow(r);
					Row marketoRow = marketoSheet.createRow(r - rowDelta);
					
					// Create cells in marketo sheet to populate
					Cell marketoFirstName = marketoRow.createCell(0);
					Cell marketoLastName = marketoRow.createCell(1);
					Cell marketoCompanyName = marketoRow.createCell(2);
					Cell marketoEmail = marketoRow.createCell(3);
					Cell marketoCountry = marketoRow.createCell(9);
					Cell marketoOriginalLeadSourceDescription = marketoRow.createCell(10);
					Cell marketoOriginalLeadSource = marketoRow.createCell(11);
					Cell marketoMostRecentLeadSource = marketoRow.createCell(12);
					Cell marketoMostRecentLeadSourceDescription = marketoRow.createCell(13);
					Cell marketoAgile = marketoRow.createCell(15);
					Cell marketoProductData = marketoRow.createCell(16);
					Cell marketoPaaS = marketoRow.createCell(17);
					Cell marketoLabs = marketoRow.createCell(18);
					Cell marketoPWS = marketoRow.createCell(19);
					
					// Populate constants for all contacts
					marketoCountry.setCellValue(country);
					marketoOriginalLeadSourceDescription.setCellValue(originalLeadSourceDescription);
					marketoOriginalLeadSource.setCellValue(originalLeadSource);
					marketoMostRecentLeadSource.setCellValue(mostRecentLeadSource);
					marketoMostRecentLeadSourceDescription.setCellValue(mostRecentLeadSourceDescription);
					if (!agile.isEmpty()) {
						marketoAgile.setCellValue(agile);
					}
					if (!productData.isEmpty()) {
						marketoProductData.setCellValue(productData);
					}
					if (!paas.isEmpty()) {
						marketoPaaS.setCellValue(paas);
					}
					if (!labs.isEmpty()) {
						marketoLabs.setCellValue(labs);
					}
					if (!pws.isEmpty()) {
						marketoPWS.setCellValue(pws);
					}
					
					// Get contact info from data miner sheet to populate marketo sheet
					String dataMinerFirstName = dataMinerRow.getCell(8).toString();
					String dataMinerLastName = dataMinerRow.getCell(10).toString();
					String dataMinerEmail = dataMinerRow.getCell(12).toString();
					String dataMinerDomain = dataMinerRow.getCell(11).toString();
					
					// Get company name from domain name in HashMap to populate marketo sheet		
					for (Map.Entry<String, String> entry : domainAccountMap.entrySet()) {
					    String domain = entry.getKey();
					    String company = entry.getValue();
					    if (dataMinerEmail.contains(domain)) {
					    	String companyName = company;
					    	marketoCompanyName.setCellValue(companyName);
					    	break;
					    }
					}
					
					// Populate contact info in marketo sheet
					if (dataMinerDomain.contains("NONE FOUND")) {
						rowDelta++;
					} else {
						marketoFirstName.setCellValue(dataMinerFirstName);
						marketoLastName.setCellValue(dataMinerLastName);
						marketoEmail.setCellValue(dataMinerEmail);
						goodContacts++;
					}
				}
				
				// Resize columns in Marketo sheet to fit data
				for (int col = 0; col < marketoTitleRow.getLastCellNum(); col++) {
					marketoSheet.autoSizeColumn(col);
				}
				
				// Format success percentage
				DecimalFormat df = new DecimalFormat("#.##");
				df.setRoundingMode(RoundingMode.HALF_UP);
				
				// Log stats to the console
				System.out.println();
				System.out.println("Successfully matched domain/email for " 
						+ df.format(100 - (((double) (domainsNotFound + linkedinMembers + nameInclomplete)
								/ (goodContacts + domainsNotFound + linkedinMembers + nameInclomplete)) * 100)) 
								+ "% of contacts (" + goodContacts + " matched with "
								+ (domainsNotFound + linkedinMembers + nameInclomplete) + " not matched)");
				System.out.println("Removed " + emptyRows + " empty rows, " 
						+ duplicateContacts + " duplicate entries, " 
						+ linkedinMembers + " 'LinkedIn Members', and " + nameInclomplete + " entries with incomplete names");
				
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
	
    // Helper method to check if a cell is empty
    private static boolean cellIsEmpty(Cell c) {
    	boolean isEmpty = false;
    	if (c == null || c.toString().isEmpty()) {
    		isEmpty = true;
    	}
    	return isEmpty;
    }
    
    // Helper method to check if accountName or currentName or titleName 
    // match up with any of the user provided account data
    private static boolean cellCheck(String userAccount, 
    		String accountName, String currentName, String titleName) {
    	boolean cellCheckFound = false;
		if (accountName.contains(userAccount) || currentName.contains(userAccount)
				|| currentName.contains(userAccount) || titleName.contains(userAccount)) {
			cellCheckFound = true;
		}
    	return cellCheckFound;
    }
    
    // Helper method to split " AND " separated account names and check if 
    // accountName or currentName or titleName match up with any of the user provided account data
    private static boolean splitCellCheck(String userAccount, 
    		String accountName, String currentName, String titleName) {
    	boolean splitCheckFound = false;
    	if (!userAccount.contains(" and ")) {
    		System.out.println("This string does not contain 'and' separator!");
    	} else {
    		String[] accountBooleanSplit = userAccount.split(" and ", 2);
    		String leftAccount = accountBooleanSplit[0];
    		String rightAccount = accountBooleanSplit[1];
    		if ((accountName.contains(leftAccount) && accountName.contains(rightAccount))
    				|| (currentName.contains(leftAccount) && currentName.contains(rightAccount))
    				|| (titleName.contains(leftAccount) && titleName.contains(rightAccount))) {
    			splitCheckFound = true;
    		}
    	}
    	return splitCheckFound;
    }
	
}
