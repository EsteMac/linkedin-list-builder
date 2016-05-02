# linkedin-list-builder

## Synopsis

This project allows users to take an output file of LinkedIn contacts scraped with the DataMiner (https://data-miner.io/) Chrome extension, attach email addresses to contacts, and deliver the output in an xlsx Excel file, along with a data summary which ranks named accounts for each LinkedIn search term used. In its present iteration, the DataMiner output must first be saved in an .xlsx file in order to be compatable with the imported Apache POI (https://poi.apache.org/) libraries which are used to aid in manipulating the Excel file. The program delivers output data in the project file email-builder-output.xlsx.

All program logic is contained in the ExcelEmailBuilder.java file. Input data specific to each region must be entered into an additional sheet within the Excel file containing output from DataMiner. This data must be entered in a pre-defined format in order for the program to execute successfully. Input data unique to each region is typically static and should be copy / pasted into the added sheet before DataMiner output can be processed. Once this data is built for the region it can quickly be re-used for each list, although updates are necessary any time the region named account list changes. Email addresses are appended to each contact based on researched common email patterns for each company domain associated with the contact.

The program also iterates across the DataMiner output file and removes empty rows, duplicate rows, and rows which contain contacts listed as "LinkedIn Member" due to a lack of connections between the LinkedIn user and the contact. Contacts who were not able to be matched with an email address are not removed by the program. Once these are manually checked to ensure there is not an error in failed program logic, or input data, they should be manually removed from within the output Excel file itself.

## Code Example

In order to identify the company domain associated with each contact, the program checks multiple fields associated with the contact against input named account data to first determine the contact's employer.
Example:
```
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
					// This cell associates the correct email format structure with each domain
					// which is stored in the emailStructureMap 
					Cell emailTypeCell = accountRow.getCell(1);
					emailStructureMap.put(domainName, (int) emailTypeCell.getNumericCellValue());
					break search;
				}
			} else if (accountName.contains(accountNameString)
						|| currentName.contains(accountNameString)
						|| titleName.contains(accountNameString)) {
					domainName = accountRow.getCell(0).toString();
					// This cell associates the correct email format structure with each domain
					// which is stored in the emailStructureMap 
					Cell emailTypeCell = accountRow.getCell(1);
					emailStructureMap.put(domainName, (int) emailTypeCell.getNumericCellValue());
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
```

Once the contact is associated with a domain name, the program then constructs an email address for the contact via the use of a switch statement.
Example:
```
// Build email addresses based on email structure types stored in emailStructuremap		
if (!domainName.contains("NONE FOUND")) {
	switch (emailStructureMap.get(domainName))
	{
		// Cases with FirstInitial + LastName@domainName
		case 1:
			nonEmptyRows++;
			email = firstName.substring(0, 1) + lastName + "@" + domainName; 
			break;
			
		// Cases with FirstName.MiddleInitial(if available).LastName@domainName
		case 2:
			nonEmptyRows++;
			if (middleName != null) {
				email = firstName + "." + middleName.substring(0, 1) + lastName + "@" + domainName;
			} else {
				email = firstName + "." + lastName + "@" + domainName;
			}
			break;
		
		// Cases with FirstName.LastName@domainName
		case 3:
			nonEmptyRows++;
			email = firstName + "." + lastName + "@" + domainName; 
			break;
			
		// Cases with LastName.FirstName@domainName
		case 4:
			nonEmptyRows++;
			email = lastName + "." + firstName + "@" + domainName; 
			break;							
										
		// Cases with LastName + FirstInitial@domainName
		case 5:
			nonEmptyRows++;
			email = lastName + firstName.substring(0, 1) + "@" + domainName;
			break;
			
		// Cases with LastName + FirstName@domainName
		case 6:
			nonEmptyRows++;
			email = lastName + firstName + "@" + domainName;
			break;
			
		// Cases with FirstName_LastName@domainName
		case 7:
			nonEmptyRows++;
			email = firstName + "_" + lastName + "@" + domainName;
			break;
			
		// Cases with LastName-FirstName@domain.com
		case 8:
			nonEmptyRows++;
			email = lastName + "-" + firstName + "@" + domainName;
			break;
			
		// Cases with first 6 letters of LastName + FirstInitial@domainName
		case 9:
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
		case 10:
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
}
emailCell.setCellValue(email);

```

## Motivation

This project was created in order to elliminate the tedious task of manually assigning email addresses to contacts scraped from LinkedIn.

## Installation and Running

This projected was created in Spring Tool Suite. All that is needed is to run the ExcelEmailBuilder.java class which will prompt the user to select the DataMiner output xlsx file for processing. Before running the program, the user should ensure that the 0 based index of the DataMiner output file column which contains the contact's full name matches up with the program global variable nameColumnIndex, which in most cases should be set to 0 as shown below:
```
// Set column index containing contact full names. All other columns are referred to relative to this one
		final int nameColumnIndex = 0;
```
All other columns in DataMiner output sheet are located in relation to the nameColumnIndex.
The user must also create a new input sheet within the DataMiner xlsx file.

<h4>Input Data</h4>

Once the DataMiner output file is saved as an xlsx file, a new sheet within this file should be created as the second sheet (left to right) in the file. This sheet should be populated with the following data types in this format, including one title row:

1. <h5>COLUMN A:</h5> Domain - This is the email domain name for each named account. Ex. advance-auto.com

2. <h5>COLUMN B:</h5> Email Structure Type - For each email domain, the user must denote the most common email format type for that account. This is entered as a numerical value into this field ranging from 1 - 10 by using the following legend:

	| Email Structure                                                                         | Type  | 
	| :---------------------------------------------------------------------------------------|:------|
	| FirstInitial + LastName @domainName                                                     | 1     | 
	| FirstName.MiddleInitial(if available).LastName @domainName                              | 2     |
	| FirstName.LastName @domainName                                                          | 3     |
	| LastName.FirstName @domainName                                                          | 4     | 
	| LastName + FirstInitial @domainName                                                     | 5     |
	| LastName + FirstName @domainName                                                        | 6     |
	| FirstName_LastName @domainName                                                          | 7     | 
	| LastName-FirstName @domain.com                                                          | 8     |
	| LastName + FirstInitial @domainName                                                     | 9     |
	| first 6 letters of LastName + FirstInitial + MiddleInitial @domainName                  | 10    | 
	
3. <h5>COLUMN C:</h5> Account Name 1 - This field should contain a unique account name (or partial name) for the account that the user would expect to find in the LinkedIn profile of a contact employed at that account. Care must be taken to ensure this field is unique to this account and likely not found in profiles of LinkedIn contacts employed at other accounts. Account name fields also support the "AND" boolean operator to assist in this task. There should always be a space between operator "AND" and its neighboring words. Ex. to differentiate between Blue Cross Blue Shield of North Carolina and Blue Cross Blue Shield of Florida, the user might enter "blue cross AND north carolina" (without quotations) for the North Carolina account, and "blue cross AND florida" for the Florida account. Boolean "OR" operator is not supported within account named fields (as "AND" is supported), instead the use of additional account named columns are used for this purpose.

4. <h5>COLUMN D:</h5> Account Name 2 - This column allows the user to enter a second common spelling of the account name that the user anticipates could be used by LinkedIn contacts for that named account. The program searches for each account name permutation using "OR" logic, meaning if either the name in this field is used in a LinkedIn profile OR the name in another account name field for this profile is used, the contact is matched with the account and domain name in question.

5. <h5>COLUMN E:</h5> Account Name 3 - This column provides the user with yet another opportunity to enter in another common spelling of the account name that the user anticipates could be used by LinkedIn contacts for that named account.

6. <h5>COLUMN F:</h5> Account Name 4 - Provides the same function as other account name columns. 

Additional account name columns may be appended to this sheet without limit, if the user chooses to include additional account name spellings. Only the first account name column is required to be populated for each row.

Example input sheet:

| Domain           | Email Structure Type | Acct Name 1      | Acct Name 2    | Acct Name 3    | Acct Name 4    |
| :----------------|:---------------------|:-----------------|:---------------|:---------------|:---------------|
| advance-auto.com | 3                    | advance auto     | advance-auto   | advanceauto    |                |
| ge.com           | 3                    | general electric | ge appliances  | ge aviation    | ge digital     |
| lowes.com        | 2                    | lowe's           | lowes          |                |                |

Once the user provided named account sheet is populated as described, the user may now run the application for data processing.

<h4>Data Output</h4>

The application generates an excel file titled email-builder-output.xlsx which should contain at least 3 sheets within the workbook.

1. When opened, the first sheet in this file contains the processed DataMiner output, and associated email addresses. Columns with data generated by the application (email addresses etc...) are styled in blue. 
2. The second sheet should contain the user provided account input data. 
3. The third sheet should contain the application generated summary, which shows a stack ranking of accounts for the LinkedIn search keywords
