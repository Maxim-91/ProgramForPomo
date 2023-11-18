# Program for "pomo" (boss in Finnish)
Checks the site pages from the company’s customer base for the presence of a sale and compares the contract date with today.  
If, when checking now, the property on the website is listed as sold and the contract date is currently less than now, in this case a certain result is recorded, otherwise (if the contract date has expired) a different result is recorded.

Briefly about the key elements of the code:  
- **void AddToStartup()** // Add the program to Windows startup  
- **void TimerCallback (object state)** // The program repeats its process at 23:00 (in case the computer was not turned off)  
- **void ImportExcelToDB(DataGridView db, int dd)** // Import the client database from an Excel file into DataGridView; if yesterday's file does not exist, the program looks for the previous file, etc. up to 15 days  
- **void CheckWebsiteStatus(DataGridView dbase)** // Проверка на наличие продажи жилья, реалезована через ошибку сайта 410  
- **void ExportDBToExcel(DataGridView db)** // Export the result from the DataGridView to a new Excel file  
