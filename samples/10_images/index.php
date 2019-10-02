<?php

require( __DIR__ . '/../Bootstrap.php');

use alhimik1986\PhpExcelTemplator\PhpExcelTemplator;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;


// This sample will work only with template.xlsx file. If you want to insert images to your
// template file, please use the file "template.xlsx" in this directory as template.
// If you create template file through the context menu,
// then the output file will have an error. It's because the file has
// relationship "PrinterSettings", which have same id as your first inserted image.
// The file template.xlsx in this directory does not have relationship "PrinterSettings"
// and have not relationship id, therefore it will work correctly.
$templateFile = './template.xlsx';
$fileName = './exported_file.xlsx';

$params = [
	'{current_date}' => date('d-m-Y'),
	'{department}' => 'Sales department',
    '[sales_manager]' => [
        'Nalty A.',
        'Ochoa S.',
        'Patel O.',
    ],
	'[[images]]' => [
		['1', '2', '3', '4', '5'],
	],
	'[[sales_amount_by_hours]]' => [
		['10000', '2000', '300', '40000', '500'],
		['1000', '200000', '3000', '400', '5000'],
		['10000', '2000', '30000', '40000', '500000'],
	],
];
/*
// If you insert images like here, then these images will shift to the right,
// because the parameter [[sales_amount_by_hours]] creates additional columns.
$callbacks = [
    '[[images]]' => function(CallbackParam $param) {
        $sheet = $param->sheet;
        $row_index = $param->row_index;
        $col_index = $param->col_index;
        $cell_coordinate = $param->coordinate;
        $amount = $param->param[$row_index][$col_index];

        $drawing = new Drawing();
        $drawing->setPath(__DIR__.'/images/' . ((int)$amount) . '.png');
        $drawing->setCoordinates($cell_coordinate);
        $drawing->setWorksheet($sheet);
    },
];
*/
// To get around this bug, we have to insert pictures after inserting the parameters.
// For this we will use events.
$events = [
    PhpExcelTemplator::AFTER_INSERT_PARAMS => function(Worksheet $sheet, array $templateVarsArr) {
        $imageVarCol = null;
        $imageVarColIndex = null;
        $imageVarRow = null;
        foreach ($templateVarsArr as $rowKey => $row) {
            foreach ($row as $colKey => $colContent) {
                if ($colContent == '[[images]]') {
                    $imageVarColIndex = $colKey + 1;
                    $imageVarRow = $rowKey + 1;
                    $imageVarCol = Coordinate::stringFromColumnIndex($imageVarColIndex);
                }
            }
        }
        $col_width = $sheet->getColumnDimension($imageVarCol)->getWidth();

        for ($i = 0; $i < 5; $i++) {
            $colIndex = $imageVarColIndex + $i;
            $coordinate = Coordinate::stringFromColumnIndex($colIndex) . $imageVarRow;
            $drawing = new Drawing();
            $drawing->setPath(__DIR__ . '/images/' . ($i + 1) . '.png');
            $drawing->setCoordinates($coordinate);
            $drawing->setWorksheet($sheet);
            // The templator copy style of cell, but not width. Let's make it manually
            $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($colIndex))->setWidth($col_width);
            // Clear cell, which must contain just an image
            $sheet->getCell($coordinate)->setValue(null);
        }
    },
];
$callbacks = [];

PhpExcelTemplator::saveToFile($templateFile, $fileName, $params, $callbacks, $events);
// PhpExcelTemplator::outputToFile($templateFile, $fileName, $params, $callbacks); // to download the file from web page
