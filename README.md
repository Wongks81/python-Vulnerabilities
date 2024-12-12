Summary
=======
This is a small project to help the team consolidating data that were passed by other departments

Data
====
1. Report.xlsx - Excel sheet were received from the security team, generated from Tableau
2. Intune.xlsx - Excel sheet generated from Intune
3. Result.xlsx - Excel sheet generated by program in dist\main\main.exe

Scenario
========
Security team generate an list of vulnerabilities from tableau to excel sheet.
On our end, we need to fix the vulnerabilities listed in the sheet.

As tableau only contain laptop hostname and not the owner of the laptop, we need a way to segregate the laptops according to the
country / region for the respective teams to work on them.

How to run the program
======================
Generate the Intune.xlsx by:
  1. Going into Intune and select "Devices" > "All Devices" on the left
  2. Click on "Export"
  3. Under Export data for all managed devices, click on "Include all inventory data in the exported file"

Copy the Report.xlsx and Intune.xlsx to where main.exe is located
Run main.exe
