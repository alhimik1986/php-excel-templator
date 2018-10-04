<?php
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\setters\CellSetterSingleValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArrayValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArray2DValue;

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

$params = [
	'{current_date}' => new ExcelParam(CellSetterSingleValue::class, $now->format('d-m-Y')),
	'{department}' => new ExcelParam(CellSetterSingleValue::class, 'Sales department'),

	'[date]' => new ExcelParam(CellSetterArrayValue::class, $dateArr),
	'[code]' => new ExcelParam(CellSetterArrayValue::class, $codeArr),
	'[manager]' => new ExcelParam(CellSetterArrayValue::class, $managerArr),
	'[sales_amount]' => new ExcelParam(CellSetterArrayValue::class, $salesAmountArr),
];

return $params;
