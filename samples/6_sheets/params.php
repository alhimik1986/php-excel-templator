<?php
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\setters\CellSetterStringValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArrayValue;

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

define('STRING_TYPE', CellSetterStringValue::class);
define('ARRAY_TYPE', CellSetterArrayValue::class);

$params = [
	'{current_date}' => new ExcelParam(STRING_TYPE, $now->format('d-m-Y')),
	'{department}' => new ExcelParam(STRING_TYPE, 'Sales department'),

	'[date]' => new ExcelParam(ARRAY_TYPE, $dateArr),
	'[code]' => new ExcelParam(ARRAY_TYPE, $codeArr),
	'[manager]' => new ExcelParam(ARRAY_TYPE, $managerArr),
	'[sales_amount]' => new ExcelParam(ARRAY_TYPE, $salesAmountArr),
];

return $params;
