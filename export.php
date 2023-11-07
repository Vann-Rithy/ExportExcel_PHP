<?php
require 'vendor/autoload.php'; // Include the PHPSpreadsheet library

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    // Database connection details
    $dbHost = "localhost";
    $dbUser = "plexustrust_vcardapp";
    $dbPassword = "Vcardapp$@168";
    $dbName = "plexustrust_vcardapp";

    // Create a database connection
    $conn = new mysqli($dbHost, $dbUser, $dbPassword, $dbName);

    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }

    $companyName = $_POST['company_name'];

    // Create a new Excel spreadsheet
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'Khmer Name');
    $sheet->setCellValue('B1', 'Name');
    $sheet->setCellValue('C1', 'Email');
    $sheet->setCellValue('D1', 'Phone Number');
    $sheet->setCellValue('E1', 'Company Name');
    $sheet->setCellValue('F1', 'Position');
    $sheet->setCellValue('G1', 'Website');
    $sheet->setCellValue('H1', 'Location');
    $sheet->setCellValue('I1', 'Facebook');
    $sheet->setCellValue('J1', 'Telegram');
    $sheet->setCellValue('K1', 'LinkedIn');
    $sheet->setCellValue('L1', 'Instagram');
    $sheet->setCellValue('M1', 'Twitter');
    $sheet->setCellValue('N1', 'YouTube');

    // Fetch data from the database based on the selected company
    if ($companyName === "All Company") {
        $sql = "SELECT name_khmer, name_english, email, phone_number, company_name, position, website, location, facebook, telegram, linkedin, instagram, twitter, youtube FROM person
            INNER JOIN contact ON person.person_id = contact.person_id
            INNER JOIN company ON person.person_id = company.person_id
            INNER JOIN social ON person.person_id = social.person_id";
    } else {
        $sql = "SELECT name_khmer, name_english, email, phone_number, company_name, position, website, location, facebook, telegram, linkedin, instagram, twitter, youtube FROM person
            INNER JOIN contact ON person.person_id = contact.person_id
            INNER JOIN company ON person.person_id = company.person_id
            INNER JOIN social ON person.person_id = social.person_id
            WHERE company_name = ?";
    }
// ...


    // Prepare and execute the SQL statement with a parameter (if not "All Company")
    if ($companyName !== "All Company") {
        $stmt = $conn->prepare($sql);
        $stmt->bind_param("s", $companyName);
        $stmt->execute();
        $result = $stmt->get_result();
    } else {
        $result = $conn->query($sql);
    }

    // Loop through the results and add data to the spreadsheet
    $row = 2;
    while ($row_data = $result->fetch_assoc()) {
        $sheet->setCellValue('A' . $row, $row_data['name_khmer']);
        $sheet->setCellValue('B' . $row, $row_data['name_english']);
        $sheet->setCellValue('C' . $row, $row_data['email']);
        $sheet->setCellValue('D' . $row, $row_data['phone_number']);
        $sheet->setCellValue('E' . $row, $row_data['company_name']);
        $sheet->setCellValue('F' . $row, $row_data['position']);
        $sheet->setCellValue('G' . $row, $row_data['facebook']);
        $sheet->setCellValue('H' . $row, $row_data['telegram']);
        $sheet->setCellValue('I' . $row, $row_data['linkedin']);
        $sheet->setCellValue('J' . $row, $row_data['instagram']);
        $sheet->setCellValue('K' . $row, $row_data['twitter']);
        $sheet->setCellValue('L' . $row, $row_data['youtube']);
        $row++;
    }

    // Save the Excel file to a local path
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $filename = 'exported_data.xlsx';
    $writer->save($filename);

    // Close the database connection
    $conn->close();

    // Serve the Excel file for download
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="' . $filename . '"');
    readfile($filename);
    unlink($filename); // Remove the file after serving

    exit();
} else {
    echo "Invalid request.";
}
?>
