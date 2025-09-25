# How to Generate the Monthly SLA Report:

Assuming you have already completed the set-up steps below, generating monthly SLA reports is as simple as the following:
1. Download Excel file from each of the 4 SLA filters on Teamwork. They should download with the names `All Tasks Report - YY MMM.xlsx` or `All Tasks Report - YY MMM (1-3).xlsx`.
2. Double click the shortcut that says `Generate Monthly SLA Report`.



# What happens when I run this program?

When you double click the `Generate Monthly SLA Report` shortcut, it runs the Python program called `updateSLA.py`, which does the following:
1. It makes a copy of the `SLA Report Template.xlsx` file in your downloads folder and renames it according to the current month.
2. It looks through the 4 files you downloaded from Teamwork and renames them depending on what type of task information they contain. (Prototypes, PSIAs, etc.)
3. It deletes any tasks that don't have the typical task name, which is to clean the data to just include "normal" a11y tasks.
4. It copies the remaining data from those 4 files into their respective sheets on the monthly report that it created in step 1.
5. It cleans the monthly report, ensuring that there are no extra rows and that the tables fit the size of the data.
6. It moves the monthly report to the correct folder, creating a folder if it needs to.
7. It modifies the overall SLA Report Overview.xlsx document to include a row for the current month.
8. It deletes the 4 files downloaded from Teamwork from your downloads folder.



# How do I set up my computer to use this process?

To use the `Generate Monthly SLA Report` program, these steps must first be followed as setup:
1. Install Python from the Microsoft Store: [Microsoft Store - Python](https://apps.microsoft.com/detail/9PJPW5LDXLZ5?hl=en-us&gl=US&ocid=pdpshare)
2. You will need to install several packages for Python that the program uses. To do so, open the windows terminal and type the following: (pressing enter after each line)
    1. ```pip install pandas```
    2. ```pip install openpyxl```
    3. ```pip install pywin32```
3. Add the SLA folder as a trusted location in Excel:
    1. Open Excel.
    2. Click **Options** on bottom left.
    3. Click **Trust Center** from the list.
    4. Click **Trust Center Settings**.
    5. Click **Trusted Locations** from the list.
    6. Check **Allow Trusted Locations on my network**.
    7. Click **Add new location...**.
    8. Paste the following in the Path: `\\byu.local\dcedfsroot\isdata\IS\Quality Assurance\ACCESSIBILITY\SLA Monthly Reports\`
    9. Check **Subfolders of this location are also trusted**.
    10. Click **OK** on all Excel pop-up windows.

To use the Tampermonkey script to download the 4 files from Teamwork, these steps must first be followed as setup:
1.  Add the Tampermonkey Chrome extension: [Chrome Extension - Tampermonkey](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
2.  (Optional) Click the puzzle icon in the upper-right of chrome and pin the Tampermonkey extension.
3.  Right click the Tampermonkey extension -> **Manage extension**.
4.  Toggle **Developer mode** on in the upper-right corner.
5.  Toggle **Allow User Scripts** on in the middle of the page.
6.  Click the Tampermonkey extension -> **Dashboard**.
7.  Click the **Utilities** tab in the upper-right corner.
8.  Click **Choose File** by "Import from file".
9.  Navigate to `N:\IS\Quality Assurance\ACCESSIBILITY\SLA Monthly Reports` and select the file `Download SLA Spreadsheets.user.js`.
10. Click **Install**.
11. On [Teamwork's task page](https://byuis.teamwork.com/app/everything/tasks), make sure you have 4 SLA filters named `SLA - Prototypes`, `SLA - 50% Reviews`, `SLA - PSIAs`, and `SLA - Peer Verifications`.
12. Setup is now complete. To use the script, click the Tampermonkey extension -> **Download SLA Documents**.
13. The first time you do this, Chrome may block multiple pop-ups. In this case, click the icon in the address bar and say **Always allow pop-ups and redirects from https꞉//byuis․teamwork․com**.



# What is the purpose of each file?

| Document                               | Purpose                                                                                                          |
|----------------------------------------|------------------------------------------------------------------------------------------------------------------|
|`README.md`                             | The file you are currently reading, which contains instructions for using this program.                          |
|`Generate Monthly SLA Report`           | This shortcut simply runs the updateSLA.py file when you double click it.                                        |
|`updateSLA.py`                          | Holds the code for the automation of the SLA report.                                                             |
|`SLA Update Program Log.txt`            | This logs information including warnings and errors from the program in case something goes wrong.               |
|`SLA Report Overview.xlsx`              | This is the overall SLA report. It shows all historical data from all monthly SLA reports.                       |
|`(YEAR) SLA Folders`                    | Each folder contains the monthly SLA reports for that year.                                                      |
|`SLA Report Template.xlsx`              | This is the template that is copied for each monthly report in step 1 of the program.                            |
|`Download SLA Spreadsheets.user.js`     | This is a Tampermonkey script that can be used in Chrome to automate downloading the 4 Excel files from Teamwork.|