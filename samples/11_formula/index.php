<?php

require( __DIR__ . '/../Bootstrap.php');

use alhimik1986\PhpExcelTemplator\PhpExcelTemplator;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArrayValueSpecial;
use alhimik1986\PhpExcelTemplator\setters\CellSetterFormulaValue;
use alhimik1986\PhpExcelTemplator\setters\DTO\FormulaValue;

$templateFile = './template.xlsx';
$fileName = './exported_file.xlsx';
define('SPECIAL_ARRAY_TYPE', CellSetterArrayValueSpecial::class);
define('FORMULA_TYPE', CellSetterFormulaValue::class);

$params = [
	'[col1]' => new ExcelParam(SPECIAL_ARRAY_TYPE, [1, 2, 3, 4]),
    '[col2]' => new ExcelParam(SPECIAL_ARRAY_TYPE, [2, 3, 4, 5]),
    '[col3]' => new ExcelParam(FORMULA_TYPE, new FormulaValue('=(%-2,0%)+(%-1,0%)', 4)),
];

PhpExcelTemplator::saveToFile($templateFile, $fileName, $params);
// PhpExcelTemplator::outputToFile($templateFile, $fileName, $params); // to download the file from web page
