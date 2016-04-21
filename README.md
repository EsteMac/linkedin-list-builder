# linkedin-list-builder

## Synopsis

This project allows users to take an output file of LinkedIn contacts scraped with the DataMiner (https://data-miner.io/) Chrome extension, attach email addresses to contacts, and deliver the output in an xlsx Excel file. In its present iteration, the DataMiner output must first be saved in an .xlsx file in order to be compatable with the imported Apache POI (https://poi.apache.org/) libraries which are used to aid in manipulating the Excel file. The program delivers output data in the project file email-builder-output.xlsx.

All logic and data for the program is contained in the ExcelEmailBuilder.java class which runs on a single thread and currently does not implement the use of objects or a database. Hard coded data is specific to named accounts in the Southeast Region, and must be updated any time the named account list changes. Email addresses are appended to each contact based on researched common email patterns for each company domain associated with the contact.

The program also iterates across the DataMiner output file and removes empty rows, duplicate rows, and rows which contain contacts listed as "LinkedIn Member" due to a lack of connections between the LinkedIn user and the contact. Contacts who were not able to be matched with an email address are not removed by the program. Once these are manually checked to ensure there is not an error in failed program logic, they are manually removed within the output Excel file itself.

## Code Example

In order to identify the company domain associated with each contact, the program checks multiple fields associated with the contact to first determine the contact's employer via the use of if / else statements.
Example:
```
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
```

Once the contact is associated with a domain name, the program then constructs an email address for the contact via the use of a switch statement.
Example:
```
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
```

## Motivation

This project was created in order to elliminate the tedious task of manually assigning email addresses to contacts scraped from LinkedIn.

## Installation

This projected was created in Spring Tool Suite. All that is needed is to run the ExcelEmailBuilder.java class which will prompt the user to select the DataMiner output xlsx file for processing. Before running the program, the user should ensure that the 0 based index of the DataMiner output file column which contains the contact's full name matches up with the program global variable nameColumnIndex, which in most cases should be set to 0 as shown below:
```
// Set column index containing contact full names. All other columns are referred to relative to this one
		final int nameColumnIndex = 0;
```
All other columns are located in relation to the nameColumnIndex.
