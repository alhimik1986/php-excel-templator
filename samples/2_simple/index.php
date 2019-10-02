<?php

require( __DIR__ . '/../Bootstrap.php');

use alhimik1986\PhpExcelTemplator\PhpExcelTemplator;

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
		['100', '200', '300', '400', '500', '600', '700', '800'],
		['1000', '2000', '3000', '4000', '5000', '6000', '7000', '8000'],
		['10000', '20000', '30000', '40000', '50000', '60000', '70000', '80000'],
	],
];
PhpExcelTemplator::saveToFile($templateFile, $fileName, $params);
// PhpExcelTemplator::outputToFile($templateFile, $fileName, $params); // to download the file from web page
