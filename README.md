# vba-sign-in-tool
**Excel Sign-In Sheet with FT# Lookup**
An automated Excel sign-in sheet built with VBA, designed to streamline user information lookup via Active Directory (AD) and track sign-in and resolution metrics.

**Features**
**FT# Lookup:**
Automatically retrieves user details (name, email, phone) from Active Directory using their FT#.
**Time Tracking:**
Logs sign-in and resolution times for each user.
**Metrics Reporting:**
Generates reports with metrics like total sign-ins, average resolution time, and unique users.
**Automation:**
Uses VBA macros for real-time data entry and AD integration.

**How to Use**
**Open the Sheet:**
Use sign_in_sheet.xlsm in Excel.
**Sign-In Process:**
Enter the FT# in column D.
The sheet automatically populates the following:
Name in column E
Phone in column G
Email in column H
**Resolution Tracking:**
Mark "Resolved?" in column I as Yes or No.
The sheet records the date and time in column J.
**Generate Reports:**
Run the GenerateReportAndChart macro for detailed analytics and visualizations.

**Setup Instructions**
**Enable Macros in Excel:**
Go to File > Options > Trust Center > Trust Center Settings > Macro Settings.
Select "Enable all macros."
**Configure LDAP for FT# Lookup:**
Open the VBA editor (Alt + F11).
Update the domain variable in the code to match your Active Directory settings:
Example: DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL
**Save the sheet and start using it!**

**Requirements**
Microsoft Excel 2016 or later
Access to Active Directory (for FT# lookup)

**Repository Structure**

sign-in-sheet/
├── LICENSE                # License file
├── README.md              # Documentation
├── sign_in_sheet.xlsm     # Excel file with VBA macros

License
See the LICENSE file for more details.

Credits
Created by Joseph Simpson.
