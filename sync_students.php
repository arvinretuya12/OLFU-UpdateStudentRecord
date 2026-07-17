<?php
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use Google\Client;
use Google\Service\Sheets;
use Google\Service\Sheets\BatchUpdateSpreadsheetRequest;
use Google\Service\Sheets\ValueRange;

// ================= CONFIGURATION =================
$spreadsheetId = '1cMXTUT--dhMvVuP6AiohfJ5dnZc7mzVk2ioglACMfq8'; // <--- PASTE YOUR ID HERE
// =================================================

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    die("Please use the form to access this script.");
}

echo "<h3>Processing...</h3>";

try {
    // 1. Handle File Upload
    if (!isset($_FILES['excel_file']) || $_FILES['excel_file']['error'] != 0) {
        throw new Exception("File upload failed.");
    }

    $targetSheetName = $_POST['sheet_name'];
    $directory = 'uploads/';

    // 1. Create the directory if it doesn't exist
    if (!is_dir($directory)) {
        mkdir($directory, 0755, true);
    }

    // 2. Clean the filename (optional but recommended for Linux)
    // This replaces spaces and brackets with underscores
    $safeName = preg_replace('/[^a-zA-Z0-9._-]/', '_', $_FILES['excel_file']['name']);
    $uploadedFile = $directory . $safeName;

    // 3. Move the file
    if (!move_uploaded_file($_FILES['excel_file']['tmp_name'], $uploadedFile)) {
        throw new Exception("Could not move uploaded file. Check if 'uploads/' is writable.");
    }

    // 2. Setup Google Client
    $client = new Client();
    $client->setApplicationName('Student Sync Script');
    $client->setScopes([Sheets::SPREADSHEETS]);
    // Logic: Use Environment Variable if on Render, fallback to local file if on PC
    $googleJson = getenv('GOOGLE_AUTH_JSON');

    if ($googleJson) {
        // We are on Render (Production)
        $authConfig = json_decode($googleJson, true);
        $client->setAuthConfig($authConfig);
    } else {
        // We are on your Local Machine (Development)
        $client->setAuthConfig('credentials.json');
    }

    $service = new Sheets($client);

    // 3. Read Excel
    $spreadsheet = IOFactory::load($uploadedFile);
    $excelData = $spreadsheet->getActiveSheet()->toArray();
    
    $excelNames = [];
    for ($i = 12; $i < count($excelData); $i++) {
        $name = trim(strtoupper($excelData[$i][2] ?? ''));
        if (!empty($name)) {
            $excelNames[] = $name;
        }
    }
    sort($excelNames);

    // 4. Read Google Sheet
    $range = $targetSheetName . "!B6:B"; 
    $response = $service->spreadsheets_values->get($spreadsheetId, $range);
    $gsheetRows = $response->getValues();

    $gsheetNames = [];
    if ($gsheetRows) {
        foreach ($gsheetRows as $row) {
            $gsheetNames[] = trim(strtoupper($row[0] ?? ''));
        }
    }

    // 5. Sync Logic
    $sheetRowIndex = 0; 
    $physicalRowPointer = 6; 
    
    $sheetId = getSheetIdByName($service, $spreadsheetId, $targetSheetName);

    $structureRequests = [];
    $addedStudents = [];

    echo "<div style='font-family:monospace; background:#f4f4f4; padding:15px; border-radius:5px;'>";

    // --- PHASE 1: QUEUE THE CHANGES ---
    foreach ($excelNames as $excelName) {
        
        while (isset($gsheetNames[$sheetRowIndex]) && $gsheetNames[$sheetRowIndex] === '') {
            $sheetRowIndex++;
            $physicalRowPointer++;
        }

        $currentGSheetName = $gsheetNames[$sheetRowIndex] ?? null;

        if ($excelName === $currentGSheetName) {
            $sheetRowIndex++;     
            $physicalRowPointer++; 
        } else {
            echo "QUEUEING: <strong>$excelName</strong> at row " . ($physicalRowPointer) . "<br>";
            
            $newRowIndex = $physicalRowPointer - 1; // 0-based index for API

            // Queue A: Insert Empty Row
            $structureRequests[] = [
                'insertDimension' => [
                    'range' => [
                        'sheetId' => $sheetId,
                        'dimension' => 'ROWS',
                        'startIndex' => $newRowIndex, 
                        'endIndex' => $newRowIndex + 1
                    ],
                    'inheritFromBefore' => false 
                ]
            ];
            
            // Queue B: Copy Formulas & Formatting
            $sourceRowIndex = ($newRowIndex > 5) ? ($newRowIndex - 1) : ($newRowIndex + 1);
            $structureRequests[] = [
                'copyPaste' => [
                    'source' => [
                        'sheetId' => $sheetId,
                        'startRowIndex' => $sourceRowIndex,
                        'endRowIndex' => $sourceRowIndex + 1,
                    ],
                    'destination' => [
                        'sheetId' => $sheetId,
                        'startRowIndex' => $newRowIndex,
                        'endRowIndex' => $newRowIndex + 1,
                    ],
                    'pasteType' => 'PASTE_NORMAL' 
                ]
            ];
            
            // Track the student for the value sanitization phase
            $addedStudents[] = [
                'name' => $excelName,
                'rowNumber' => $physicalRowPointer
            ];

            $physicalRowPointer++;
        }
    }

    // --- PHASE 2: BATCH EXECUTE ROW STRUCTURE (1 API CALL) ---
    if (!empty($structureRequests)) {
        $batchRequest = new BatchUpdateSpreadsheetRequest(['requests' => $structureRequests]);
        $service->spreadsheets->batchUpdate($spreadsheetId, $batchRequest);
    }

    // --- PHASE 3: BATCH SANITIZE & WRITE DATA (2 API CALLS) ---
    if (!empty($addedStudents)) {
        
        // 1. Fetch the newly copied rows in ONE bulk request
        $minRow = $addedStudents[0]['rowNumber'];
        $maxRow = $addedStudents[count($addedStudents) - 1]['rowNumber'];
        $rangeToFetch = "$targetSheetName!$minRow:$maxRow";
        
        $getParams = ['valueRenderOption' => 'FORMULA']; 
        $response = $service->spreadsheets_values->get($spreadsheetId, $rangeToFetch, $getParams);
        $allFetchedRows = $response->getValues();

        $valueUpdateRequests = [];

        // 2. Clean the data locally in PHP
        foreach ($addedStudents as $student) {
            $rowNumber = $student['rowNumber'];
            $excelName = $student['name'];
            $arrayIndex = $rowNumber - $minRow; 

            if (isset($allFetchedRows[$arrayIndex])) {
                $rowData = $allFetchedRows[$arrayIndex];
                $cleanedData = [];

                foreach ($rowData as $colIndex => $cellValue) {
                    $cellValue = (string)$cellValue;
                    
                    if ($colIndex === 1) {
                        // Hardcode the student name into Column B (Index 1)
                        $cleanedData[] = $excelName;
                    } elseif (substr($cellValue, 0, 1) === '=') {
                        // Keep formulas
                        $cleanedData[] = $cellValue;
                    } else {
                        // Clear hardcoded data/empty cells
                        $cleanedData[] = ""; 
                    }
                }

                // Ensure the array has the name even if the copied row didn't stretch to Col B
                if (!isset($cleanedData[1])) {
                    $cleanedData[1] = $excelName;
                }

                $valueUpdateRequests[] = new ValueRange([
                    'range' => "$targetSheetName!A$rowNumber",
                    'values' => [$cleanedData]
                ]);
            }
        }

        // 3. Write all the cleaned rows back to Google Sheets in ONE bulk update
        if (!empty($valueUpdateRequests)) {
            $batchValuesRequest = new \Google\Service\Sheets\BatchUpdateValuesRequest([
                'valueInputOption' => 'USER_ENTERED',
                'data' => $valueUpdateRequests
            ]);
            $service->spreadsheets_values->batchUpdate($spreadsheetId, $batchValuesRequest);
        }
    }

    echo "</div>";
    echo "<h3 style='color:green'>Sync Complete!</h3>";
    echo "<a href='index.php'>Go Back</a>";

    unlink($uploadedFile);

} catch (Exception $e) {
    echo '<h3 style="color:red">Error: ' . $e->getMessage() . '</h3>';
}

// ================= HELPER FUNCTIONS =================
// Note: We deleted insertRow, copyFullRow, sanitizeRow, and writeCell because 
// we built them directly into the high-speed batch loops above!

function getSheetIdByName($service, $spreadsheetId, $sheetName) {
    $meta = $service->spreadsheets->get($spreadsheetId);
    foreach ($meta->getSheets() as $sheet) {
        if ($sheet->getProperties()->getTitle() === $sheetName) {
            return $sheet->getProperties()->getSheetId();
        }
    }
    throw new Exception("Sheet name '$sheetName' not found.");
}
?>