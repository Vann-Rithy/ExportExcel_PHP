# Export Data from SQL to Excel Using PHP

In this guide, we will demonstrate how to export data from a SQL database to an Excel spreadsheet using PHP. We will be using the `phpoffice/phpspreadsheet` library to simplify Excel file generation. To follow along, make sure you have Composer installed and a web server environment set up for your PHP project.

## Prerequisites

1. **PHP**: Ensure you have PHP installed on your system. You can download and install PHP from the official [PHP website](https://www.php.net/downloads.php).

2. **Composer**: If you don't have Composer installed, you can download it from the official [Composer website](https://getcomposer.org/download/).

3. **Web Server**: Set up a local web server (e.g., Apache or Nginx) to run your PHP scripts. You can also use development tools like XAMPP or MAMP, which include a web server.

4. **SQL Database**: You should have a SQL database set up with the data you want to export to Excel.

## Step 1: Create a PHP Project

1. Create a new directory for your project:

   ```bash
   mkdir sql-to-excel-export
   cd sql-to-excel-export
   ```

2. Initialize a new Composer project:

   ```bash
   composer init
   ```

   Follow the prompts to create a `composer.json` file. Make sure to specify a project name and author.

3. Install the `phpoffice/phpspreadsheet` library:

   ```bash
   composer require phpoffice/phpspreadsheet --ignore-platform-req=ext-gd
   ```

   The `--ignore-platform-req=ext-gd` flag is used to bypass certain requirements, which can be useful for development environments.

## Step 2: Create PHP Script for Export

Create a new PHP file, e.g., `export.php`, in your project directory. This script will connect to your SQL database, retrieve data, and export it to an Excel file.

```php
<?php
require 'vendor/autoload.php'; // Include the Composer autoloader

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Database configuration
$host = 'localhost';
$username = 'your_db_username';
$password = 'your_db_password';
$database = 'your_database_name';

// Create a new spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Connect to the database
$pdo = new PDO("mysql:host=$host;dbname=$database", $username, $password);
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

// SQL query to fetch data from your table
$query = "SELECT * FROM your_table_name";
$stmt = $pdo->query($query);
$data = $stmt->fetchAll(PDO::FETCH_ASSOC);

// Define column headers
$columns = array_keys($data[0]);
$col = 'A';
foreach ($columns as $column) {
    $sheet->setCellValue($col . '1', $column);
    $col++;
}

// Populate data
$row = 2;
foreach ($data as $record) {
    $col = 'A';
    foreach ($record as $value) {
        $sheet->setCellValue($col . $row, $value);
        $col++;
    }
    $row++;
}

// Save the spreadsheet as an Excel file
$writer = new Xlsx($spreadsheet);
$filename = 'exported_data.xlsx';
$writer->save($filename);

// Download the Excel file
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"$filename\"");
readfile($filename);
unlink($filename); // Delete the temporary file
```

Make sure to replace the placeholders in the script with your own database credentials and table information.

## Step 3: Run the PHP Script

You can run the PHP script by accessing it through your web server. Place the script in your web server's document root or virtual host directory, and then access it in your web browser, e.g., `http://localhost/sql-to-excel-export/export.php`.

The script will connect to your database, fetch the data, create an Excel file, and initiate the download of the Excel file.

That's it! You've successfully exported data from a SQL database to an Excel file using PHP and the `phpoffice/phpspreadsheet` library.
