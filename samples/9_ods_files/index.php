<?php
use alhimik1986\PhpExcelTemplator\PhpExcelTemplatorOds;

require( __DIR__ . '/../Bootstrap.php');

PhpExcelTemplatorOds::saveToFile('./template.ods', './exported_file.ods', [
	'{current_date}' => date('d-m-Y'),
	'{department}' => 'Sales department',
]);
