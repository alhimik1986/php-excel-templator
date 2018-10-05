<?php

require( __DIR__ . '/../Bootstrap.php');

use alhimik1986\PhpExcelTemplator\PhpExcelTemplator;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\setters\CellSetterStringValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArrayValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArray2DValue;

$templateFile = './template.xlsx';
$fileName = './exported_file.xlsx';

$now = new DateTime();
$dateArr = [
	'01-06-2018',
	'02-06-2018',
	'03-06-2018',
	'04-06-2018',
	'05-06-2018',
];
$codeArr = [
	'0001543',
	'0003274',
	'000726',
	'0012553',
	'0008245',
];
$managerArr = [
	'Adams D.',
	'Baker A.',
	'Clark H.',
	'Davis O.',
	'Evans P.',
];
$salesAmountArr = [
	'10 230 $',
	'45 100 $',
	'70 500 $',
	'362 180 $',
	'5 900 $',
];
$salesManagerArr = [
	'Nalty A.',
	'Ochoa S.',
	'Patel O.',
];
$hoursArr = [
	['01', '02', '03', '04', '05', '06', '07', '08'],
];
$numOfSalesByHours = [
	['100', '200', '300', '400', '500', '600', '700', '800'],
	['1000', '2000', '3000', '4000', '5000', '6000', '7000', '8000'],
	['10000', '20000', '30000', '40000', '50000', '60000', '70000', '80000'],
];

$params = [
	'{current_date}' => new ExcelParam(CellSetterStringValue::class, $now->format('d-m-Y')),
	'{department}' => new ExcelParam(CellSetterStringValue::class, 'Sales department'),

	'[date]' => new ExcelParam(CellSetterArrayValue::class, $dateArr),
	'[code]' => new ExcelParam(CellSetterArrayValue::class, $codeArr),
	'[manager]' => new ExcelParam(CellSetterArrayValue::class, $managerArr),
	'[sales_amount]' => new ExcelParam(CellSetterArrayValue::class, $salesAmountArr),

	'[sales_manager]' => new ExcelParam(CellSetterArrayValue::class, $salesManagerArr),
	'[[hours]]' => new ExcelParam(CellSetterArray2DValue::class, $hoursArr),
	'[[sales_amount_by_hours]]' => new ExcelParam(CellSetterArray2DValue::class, $numOfSalesByHours),
];
PhpExcelTemplator::saveToFile($templateFile, $fileName, $params);
// PhpExcelTemplator::outputToFile($templateFile, $fileName, $params); // to download the file from web page
