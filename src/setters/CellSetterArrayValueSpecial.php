<?php

namespace alhimik1986\PhpExcelTemplator\setters;

use alhimik1986\PhpExcelTemplator\InsertedCells;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\SetterParam;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

class CellSetterArrayValueSpecial implements ICellSetter
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
		$this->_insertNewRowsIfNeed($sheet, $values, $insertedCells, $row_key, $pColumnIndex, $pRow);

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

		$maxInsertedRows = $this->_getMaxInsertedRows($row_key, $insertedCells);
		$maxInsertedRows = (count($values) - 1 > $maxInsertedRows) ? count($values) - 1 : $maxInsertedRows;
		$maxColumnIndex = Coordinate::columnIndexFromString($sheet->getHighestColumn());
		for($i=0; $i<=$maxColumnIndex; $i++) {
			$insertedCells->setInsertedRows($row_key, $i, $maxInsertedRows);
		}

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
     * @param InsertedCells $insertedCells
     * @param integer $row_key The row index, where was template variable
     * @param integer $pColumnIndex The current column index
     * @param integer $pRow The current row index
     * @throws SpreadsheetException
     */
	private function _insertNewRowsIfNeed(Worksheet $sheet, $values, $insertedCells, $row_key, $pColumnIndex, $pRow): void
    {
		$maxInsertedRows = $this->_getMaxInsertedRows($row_key, $insertedCells);
		$rowsToInsert = (count($values) - 1 > $maxInsertedRows) ? count($values) - 1 : 0;
		$maxInsertedRows = ($maxInsertedRows > $rowsToInsert) ? $maxInsertedRows : $rowsToInsert;
		if ($rowsToInsert > 0) {
			$sheet->insertNewRowBefore($pRow+1, $rowsToInsert);
		}
		$this->_mergeColumnsIfNeed($sheet, $pColumnIndex, $pRow, $maxInsertedRows);
	}

	/**
	 * @param InsertedCells $insertedCells
	 * @param integer $row_key The row index, where was template variable
	 * @return integer Maximum number of inserted rows in all columns of the specified row
	 */
	private function _getMaxInsertedRows($row_key, $insertedCells): int
    {
		$maxInsertedRows = 0;
		foreach($insertedCells->inserted_rows as $col_key=>$insertedRowsInCol) {
			$insertedRows = $insertedCells->getInsertedRows($row_key, $col_key);
			if ($insertedRows > $maxInsertedRows) {
				$maxInsertedRows = $insertedRows;
			}
		}

		return $maxInsertedRows;
	}

    /**
     * @param Worksheet $sheet
     * @param integer $pColumnIndex Current column index
     * @param integer $pRow Current row index
     * @param integer $maxInsertedRows Maximum number of inserted rows in the current row
     * @throws SpreadsheetException
     */
	private function _mergeColumnsIfNeed($sheet, $pColumnIndex, $pRow, $maxInsertedRows): void
    {
		$pCol = Coordinate::stringFromColumnIndex($pColumnIndex);
		$coordinate = $pCol.$pRow;
		$mergedCellsCount = $this->_getMergedCellsCount($sheet, $coordinate);
		if ($mergedCellsCount > 0) {
			for($rowOffset=0; $rowOffset<$maxInsertedRows; $rowOffset++) {
				$rowIndex = $pRow + 1 + $rowOffset;
				$coordinate1 = $pCol.$rowIndex;
				$pCol2 = Coordinate::stringFromColumnIndex($pColumnIndex + $mergedCellsCount);
				$coordinate2 = $pCol2.$rowIndex;
				$mergeRange = Coordinate::buildRange([[$coordinate1, $coordinate2]]);
				$sheet->mergeCells($mergeRange);
			}
		}
	}

    /**
     * @param Worksheet $sheet
     * @param string $coordinate Current cell coordinate
     * @return integer Number of merged cells in the specified cell coordinate
     * @throws SpreadsheetException
     */
	private function _getMergedCellsCount($sheet, $coordinate): int
    {
		$mergedCellsCount = 0;
		$cell = $sheet->getCell($coordinate);
		foreach ($sheet->getMergeCells() as $cells) {
            /** @noinspection NullPointerExceptionInspection */
            if ($cell->isInRange($cells)) {
				$cellsRangeArr = Coordinate::splitRange($cells);
				$cellsArr = $cellsRangeArr[0];
				$coord1 = Coordinate::coordinateFromString($cellsArr[0]);
				$coord2 = Coordinate::coordinateFromString($cellsArr[1]);
				$colIndex1 = Coordinate::columnIndexFromString($coord1[0]);
				$colIndex2 = Coordinate::columnIndexFromString($coord2[0]);
				$mergedCellsCount = abs($colIndex2 - $colIndex1);
			}
		}

		return $mergedCellsCount;
	}
}
