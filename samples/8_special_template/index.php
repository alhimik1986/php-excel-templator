<?php

require( __DIR__ . '/../Bootstrap.php');

use alhimik1986\PhpExcelTemplator\PhpExcelTemplator;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArrayValueSpecial;

$templateFile = './template.xlsx';
$fileName = './exported_file.xlsx';
define('SPECIAL_ARRAY_TYPE', CellSetterArrayValueSpecial::class);

$params = [
	'[product_name]' => new ExcelParam(SPECIAL_ARRAY_TYPE, [
		'product 1',
		'product 2',
		'product 3',
		'product 4',
		'product 5',
	]),
	'[product_count]' => new ExcelParam(SPECIAL_ARRAY_TYPE, [
		'1',
		'2',
		'3',
		'4',
		'5',
	]),
	'[product_price]' => new ExcelParam(SPECIAL_ARRAY_TYPE, [
		'100',
		'200',
		'300',
		'400',
		'500',
	]),
	'[product_summ]' => new ExcelParam(SPECIAL_ARRAY_TYPE, [
		'100',
		'400',
		'900',
		'1 600',
		'2 500',
	]),
	'[product_oem]' => new ExcelParam(SPECIAL_ARRAY_TYPE, [
		'Product OEM 1',
		'Product OEM 2',
		'Product OEM 3',
		'Product OEM 4',
		'Product OEM 5',
	]),
	'[product_arrival_date]' => new ExcelParam(SPECIAL_ARRAY_TYPE, [
		'2018-01-01',
		'2018-01-02',
		'2018-01-03',
		'2018-01-04',
		'2018-01-05',
	]),
	'[product_currency]' => new ExcelParam(SPECIAL_ARRAY_TYPE, [
		'р',
		'р',
		'р',
		'р',
		'р',
	]),
	'{all_summ}' => '5 500',
	'{dolg}' => '0',
	'{document_id}' => '1111111',
	'{document_date}' => date('Y-m-d'),
];
$callbacks = [
	'{all_summ}' => function(CallbackParam $param) {
		$sheet = $param->sheet;
		$cell_coordinate = $param->coordinate;
		$sheet->getStyle($cell_coordinate)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFB5FFA8');
	},
];
PhpExcelTemplator::saveToFile($templateFile, $fileName, $params, $callbacks);
// PhpExcelTemplator::outputToFile($templateFile, $fileName, $params, $callbacks); // to download the file from web page
