<?php

require 'vendor/autoload.php';

use GuzzleHttp\Client;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Function to scrape data from a website
function scrapeData($url) {
    $client = new Client();
    $response = $client->request('GET', $url);
    $html = (string) $response->getBody();

    $dom = new DOMDocument();
    @$dom->loadHTML($html);
    $xpath = new DOMXPath($dom);

    // Extract data from the table with id 'data-table'
    $data = [];
    $rows = $xpath->query('//table[@id="data-table"]//tr');
    foreach ($rows as $rowIndex => $row) {
        $rowData = [];
        $cells = $xpath->query('td', $row);
        if ($cells->length > 0) { // Skip header row
            $rowData['Name'] = trim($cells->item(0)->textContent);
            $rowData['Age'] = trim($cells->item(1)->textContent);
            $rowData['City'] = trim($cells->item(2)->textContent);
            $data[] = $rowData;
        }
    }
    return $data;
}

// Function to convert column index to letter (e.g., 1 -> A, 2 -> B)
function columnIndexToLetter($index) {
    $letter = '';
    while ($index > 0) {
        $index--;
        $letter = chr(65 + ($index % 26)) . $letter;
        $index = floor($index / 26);
    }
    return $letter;
}

// URL to scrape data from
$url = 'http://localhost/test2/test.html'; // Replace with your actual URL

// Scrape the data
$data = scrapeData($url);

// Create a new Spreadsheet object
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Write headers to the first row
$headers = ['Name', 'Age', 'City'];
foreach ($headers as $colIndex => $header) {
    $cellCoordinate = columnIndexToLetter($colIndex + 1) . '1';
    $sheet->setCellValue($cellCoordinate, $header);
}

// Write data to the spreadsheet
foreach ($data as $rowIndex => $row) {
    foreach ($headers as $colIndex => $header) {
        $cellCoordinate = columnIndexToLetter($colIndex + 1) . ($rowIndex + 2);
        $sheet->setCellValue($cellCoordinate, $row[$header]);
    }
}

// Save the spreadsheet as an Excel file
$writer = new Xlsx($spreadsheet);
$writer->save('test.xlsx');

echo "Data has been scraped and saved to scraped_data.xlsx\n";

?>
