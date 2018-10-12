<?php

require( __DIR__ . '/../Bootstrap.php');

use alhimik1986\PhpExcelTemplator\PhpExcelTemplator;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;

$templateFile = './template.xlsx';
$fileName = './exported_file.xlsx';

$params = [
	'{current_date}' => date('d-m-Y'),
	'{department}' => 'Sales department',
	'[date]' => [
		'01-06-2018',
		'02-06-2018',
		'03-06-2018',
		'04-06-2018',
		'05-06-2018',
	],
	'[code]' => [
		'0001543',
		'0003274',
		'000726',
		'0012553',
		'0008245',
	],
	'[manager]' => [
		'Adams D.',
		'Baker A.',
		'Clark H.',
		'Davis O.',
		'Evans P.',
	],
	'[sales_amount]' => [
		'10 230 $',
		'45 100 $',
		'70 500 $',
		'362 180 $',
		'5 900 $',
	],
		'[sales_manager]' => [
		'Nalty A.',
		'Ochoa S.',
		'Patel O.',
	],
	'[[hours]]' => [
		['01', '02', '03', '04', '05', '06', '07', '08'],
	],
	'[[sales_amount_by_hours]]' => [
		['10000', '2000', '300', '40000', '500', '600', '700', '800000'],
		['1000', '200000', '3000', '400', '5000', '6000', '7000', '8000'],
		['10000', '2000', '30000', '40000', '500000', '60', '70000', '800'],
	],
];
$callbacks = [
	'[sales_amount]' => function(CallbackParam $param) {
		$sheet = $param->sheet;
		$row_index = $param->row_index;
		$cell_coordinate = $param->coordinate;
		$amount = $param->param[$row_index];
		$amount = preg_replace('/[\s\$]/', '', $amount);
		if ($amount > 50000) {
			$sheet->getStyle($cell_coordinate)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFB5FFA8');
			$sheet->getStyle($cell_coordinate)->getFont()->setBold(true);
		}
	},
	'[[sales_amount_by_hours]]' => function(CallbackParam $param) {
		$sheet = $param->sheet;
		$row_index = $param->row_index;
		$col_index = $param->col_index;
		$cell_coordinate = $param->coordinate;
		$amount = $param->param[$row_index][$col_index];
		if ($amount > 50000) {
			$sheet->getStyle($cell_coordinate)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFB5FFA8');
			$sheet->getStyle($cell_coordinate)->getFont()->setBold(true);
		}
	},
];
PhpExcelTemplator::saveToFile($templateFile, $fileName, $params, $callbacks);
// PhpExcelTemplator::outputToFile($templateFile, $fileName, $params, $callbacks); // to download the file from web page
