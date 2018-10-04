<?php
use alhimik1986\PhpExcelTemplator\PhpExcelTemplator;

require( __DIR__ . '/../Bootstrap.php');

PhpExcelTemplator::saveToFile('./template.xlsx', './exported_file.xlsx', [
	'{current_date}' => date('d-m-Y'),
	'{department}' => 'Sales department',
]);

/*
// to download the file from web page
PhpExcelTemplator::outputToFile('./template.xlsx', './exported_file.xlsx', [
	'{current_date}' => date('d-m-Y'),
	'{department}' => 'Sales department',
]);
*/
