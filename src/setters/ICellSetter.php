<?php

namespace alhimik1986\PhpExcelTemplator\setters;

use alhimik1986\PhpExcelTemplator\InsertedCells;
use alhimik1986\PhpExcelTemplator\params\SetterParam;

interface ICellSetter
{
	/**
     * Set values in the specified cell.
	 * @param SetterParam $setter_param
	 * @param InsertedCells $inserted_cells The object that stores the count of inserted cells
	 * @return InsertedCells
	 */
	public function setCellValue(SetterParam $setter_param, InsertedCells $inserted_cells): InsertedCells;
}
