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

class CellSetterArray2DValue implements ICellSetter
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

		$pColumn = $insertedCells->getCurrentColIndex($row_key, $col_key);
		$pRow = $insertedCells->getCurrentRowIndex($row_key, $col_key);
		$values = $param->value;
		$this->_insertNewRowsAndColsIfNeed($sheet, $values, $pColumn, $pRow);

		foreach($values as $row_index => $value_arr) {
			foreach($value_arr as $col_index=>$value) {
				$pColumnWord = Coordinate::stringFromColumnIndex($pColumn + $col_index);
				$currCellCoordinates = $pColumnWord.($pRow + $row_index);

				$sheet->setCellValue($currCellCoordinates, $value);
				if ($param->callback) {
					$callbackParam = new CallbackParam([
						'sheet'=>$sheet, 'coordinate'=>$currCellCoordinates,
						'param'=>$param->value, 'tpl_var_name'=>$tpl_var_name,
						'row_index'=>$row_index, 'col_index'=>$col_index,
					]);
					call_user_func($param->callback, $callbackParam);
				}
			}
		}

		$firstValue = reset($values);
        $insertedCells->addInsertedCols($row_key, $col_key, count($firstValue) - 1);
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

        foreach($value as $key=>$val) {
            if ( ! is_array($value[$key])) {
                throw new RuntimeException('In the '.ExcelParam::class.' class the field "value" with index "'.$key.'" must be an array, when the setter '.__CLASS__.' is used.');
            }
        }
        return count($value) > 0;
	}

    /**
     * @param Worksheet $sheet
     * @param array $values
     * @param integer $pCol The current column index
     * @param integer $pRow
     * @throws SpreadsheetException
     */
	private function _insertNewRowsAndColsIfNeed(Worksheet $sheet, $values, $pCol, $pRow): void
    {
		$objReferenceHelper = ReferenceHelper::getInstance();

		foreach($values as $row_index=>$valueArr) {
			$colsToInsert = count($valueArr) - 1;
			//$highestColumn = Coordinate::columnIndexFromString($sheet->getHighestColumn());

			// Inserting row
			if ($row_index > 0) {
				$rowsToInsert = 1;
				$pCol1 = Coordinate::stringFromColumnIndex($pCol);
				$pCol2 = Coordinate::stringFromColumnIndex($pCol);
				$coordinate1 = $pCol1 . ($pRow + $row_index);
				$coordinate2 = $pCol2 . ($sheet->getHighestRow()+$rowsToInsert);
				$objReferenceHelper->insertNewBefore($coordinate1, 0, $rowsToInsert, $sheet, $coordinate2);
				//echo 'Row    '.$coordinate1.' '.$coordinate2."\n";
			}

			// Inserting columns
			$pCol_1 = Coordinate::stringFromColumnIndex($pCol+1);
			$pCol_2 = Coordinate::stringFromColumnIndex($pCol+$colsToInsert);
			$coordinate1 = $pCol_1 . ($pRow + $row_index);
			$coordinate2 = $pCol_2 . ($pRow + $row_index);
			$objReferenceHelper->insertNewBefore($coordinate1, $colsToInsert, 0, $sheet, $coordinate2);
			//echo 'Column '.$coordinate1.' '.$coordinate2.' '.$colsToInsert."\n";
		}
	}
}
