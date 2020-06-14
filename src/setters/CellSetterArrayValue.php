<?php

namespace alhimik1986\PhpExcelTemplator\setters;

use alhimik1986\PhpExcelTemplator\InsertedCells;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\SetterParam;
use alhimik1986\PhpExcelTemplator\ReferenceHelper;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

class CellSetterArrayValue implements ICellSetter
{
    /**
     * {@inheritDoc}
     * @throws SpreadsheetException
     */
	public function setCellValue(SetterParam $setterParam, InsertedCells $insertedCells): InsertedCells
    {
		$sheet = $setterParam->sheet;
		$row_key = $setterParam->row_key;
		$col_key = $setterParam->col_key;
		$tpl_var_name = $setterParam->tpl_var_name;
		$param = $setterParam->params[$tpl_var_name];
		if ( ! $this->_validateValue($param->value)) {
			return $insertedCells;
		}

		$pColumn = $insertedCells->getCurrentCol($row_key, $col_key);
		$pColumnIndex = $insertedCells->getCurrentColIndex($row_key, $col_key);
		$pRow = $insertedCells->getCurrentRowIndex($row_key, $col_key);
		$values = $param->value;
		$this->_insertNewRowsIfNeed($sheet, $values, $pColumnIndex, $pRow);

		foreach($values as $row_index=>$value) {
			$currCellCoordinates = $pColumn.($pRow + $row_index);

			$sheet->setCellValue($currCellCoordinates, $value);
			if ($param->callback) {
				$callbackParam = new CallbackParam([
					'sheet'=>$sheet, 'coordinate'=>$currCellCoordinates, 'param'=>$param->value,
					'tpl_var_name'=>$tpl_var_name, 'row_index'=>$row_index, 'col_index'=>0,
				]);
				call_user_func($param->callback, $callbackParam);
			}
		}

		$insertedCells->addInsertedRows($row_key, $col_key, count($values) - 1);

		return $insertedCells;
	}

    /**
     * @param mixed $value
     * @return boolean
     * @throws RuntimeException
     */
	private function _validateValue($value): bool
	{
		if ( ! is_array($value)) {
            throw new RuntimeException('In the '.ExcelParam::class.' class the field "value" must be an array, when the setter '.__CLASS__.' is used.');
		}
		return count($value) > 0;
	}

    /**
     * @param Worksheet $sheet
     * @param String[] $values
     * @param integer $pColumnIndex The current column index
     * @param integer $pRow The current row index
     * @throws SpreadsheetException
     */
	private function _insertNewRowsIfNeed(Worksheet $sheet, $values, $pColumnIndex, $pRow): void
	{
		$rowsToInsert = count($values) - 1;
		if ($rowsToInsert > 0) {
			$objReferenceHelper = ReferenceHelper::getInstance();
			$pCol = Coordinate::stringFromColumnIndex($pColumnIndex);
			$coordinate1 = $pCol . ($pRow+1);
			$coordinate2 = $pCol . ($sheet->getHighestRow()+$rowsToInsert);
			$objReferenceHelper->insertNewBefore($coordinate1, 0, $rowsToInsert, $sheet, $coordinate2);
		}
	}
}
