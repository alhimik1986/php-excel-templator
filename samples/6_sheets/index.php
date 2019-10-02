<?php

require( __DIR__ . '/../Bootstrap.php');

use alhimik1986\PhpExcelTemplator\PhpExcelTemplator;
use PhpOffice\PhpSpreadsheet\IOFactory;

$templateFile = './template.xlsx';
$fileName = './exported_file.xlsx';
$params = require('./params.php');

$spreadsheet = IOFactory::load($templateFile);
$templateVarsArr = $spreadsheet->getActiveSheet()->toArray();
$templateSheet = clone $spreadsheet->getActiveSheet();
$callbacks = [];
$events = [];

$sheet1 = $spreadsheet->getSheet(0);
PhpExcelTemplator::renderWorksheet($sheet1, $templateVarsArr, $params, $callbacks, $events);

$sheet2 = clone $templateSheet;
$sheet2->setTitle('Workshet 2');
$spreadsheet->addSheet($sheet2);
PhpExcelTemplator::renderWorksheet($sheet2, $templateVarsArr, $params, $callbacks, $events);

$sheet3 = clone $templateSheet;
$sheet3->setTitle('Workshet 3');
$spreadsheet->addSheet($sheet3);
PhpExcelTemplator::renderWorksheet($sheet3, $templateVarsArr, $params, $callbacks, $events);

PhpExcelTemplator::saveSpreadsheetToFile($spreadsheet, $fileName);
// PhpExcelTemplator::outputSpreadsheetToFile($spreadsheet, $fileName); // to download the file from web page
