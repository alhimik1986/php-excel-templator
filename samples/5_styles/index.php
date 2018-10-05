<?php

require( __DIR__ . '/../Bootstrap.php');

use alhimik1986\PhpExcelTemplator\PhpExcelTemplator;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;
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
$salesByHoursArr = [
	['10000', '2000', '300', '40000', '500', '600', '700', '800000'],
	['1000', '200000', '3000', '400', '5000', '6000', '7000', '8000'],
	['10000', '2000', '30000', '40000', '500000', '60', '70000', '800'],
];

define('STRING_TYPE', CellSetterStringValue::class);
define('ARRAY_TYPE', CellSetterArrayValue::class);
define('ARRAY_2D_TYPE', CellSetterArray2DValue::class);

$params = [
	'{current_date}' => new ExcelParam(STRING_TYPE, $now->format('d-m-Y')),
	'{department}' => new ExcelParam(STRING_TYPE, 'Sales department'),

	'[date]' => new ExcelParam(ARRAY_TYPE, $dateArr),
	'[code]' => new ExcelParam(ARRAY_TYPE, $codeArr),
	'[manager]' => new ExcelParam(ARRAY_TYPE, $managerArr),
	'[sales_amount]' => new ExcelParam(ARRAY_TYPE, $salesAmountArr, function(CallbackParam $param) {
		$sheet = $param->sheet;
		$row_index = $param->row_index;
		$cell_coordinate = $param->coordinate;
		$amount = $param->param[$row_index];
		$amount = preg_replace('/[\s\$]/', '', $amount);
		if ($amount > 50000) {
			$sheet->getStyle($cell_coordinate)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFB5FFA8');
			$sheet->getStyle($cell_coordinate)->getFont()->setBold(true);
		}
	}),

	'[sales_manager]' => new ExcelParam(ARRAY_TYPE, $salesManagerArr),
	'[[hours]]' => new ExcelParam(ARRAY_2D_TYPE, $hoursArr),
	'[[sales_amount_by_hours]]' => new ExcelParam(ARRAY_2D_TYPE, $salesByHoursArr, function(CallbackParam $param) {
		$sheet = $param->sheet;
		$row_index = $param->row_index;
		$col_index = $param->col_index;
		$cell_coordinate = $param->coordinate;
		$amount = $param->param[$row_index][$col_index];
		if ($amount > 50000) {
			$sheet->getStyle($cell_coordinate)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFB5FFA8');
			$sheet->getStyle($cell_coordinate)->getFont()->setBold(true);
		}
	}),
];
PhpExcelTemplator::saveToFile($templateFile, $fileName, $params);
// PhpExcelTemplator::outputToFile($templateFile, $fileName, $params); // to download the file from web page
