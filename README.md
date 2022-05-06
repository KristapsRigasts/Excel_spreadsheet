## Create Excel spreadsheet 
Create Excel spreadsheet with work schedules and total worked hour count using phpoffice/phpspreadsheet.

Script is executed with console command, where you can pass month and year.

In Excel file script will add different cell color for weekend days.

![Excel spreadsheet](public/workingSheet.jpg?raw=true "Excel spreadsheet")

### How to run script

* run command `composer install`
* run command `php artisan createExcel month-year` - replace month-year with month and year you want to create Excel spreadsheet.

For example: `php artisan createExcel 05-2022` - will create Excel file for May 2022.
* Created Excel file will be saved in public folder, with generated name workHourSheet_month-year.xlsx 
