<?php

namespace alhimik1986\PhpExcelTemplator\setters;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use alhimik1986\PhpExcelTemplator\InsertedCells;
use alhimik1986\PhpExcelTemplator\params\SetterParam;

interface ICellSetter
{
	/**
	 * Установить значения в указанной ячейке.
	 * @param SetterParam $setter_param
	 * @param InsertedCells $inserted_cells Объект, хранящий в себе число вставленных ячеек
	 * @return InsertedCells Объект, хранящий в себе число вставленных ячеек
	 */
	public function setCellValue(SetterParam $setter_param, InsertedCells $inserted_cells);
}
