<?php

namespace alhimik1986\PhpExcelTemplator\setters;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use alhimik1986\PhpExcelTemplator\InsertedCells;
use alhimik1986\PhpExcelTemplator\params\SetterParam;
use alhimik1986\PhpExcelTemplator\ReferenceHelper;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;
use Exception;

class CellSetterArray2DValue implements ICellSetter
{
	/**
	 * {@inheritdoc}
     * @throws Exception
	 */
	public function setCellValue(SetterParam $setterParam, InsertedCells $insertedCells) {
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
		$this->_insertNewRowsAndColsIfNeed($sheet, $values, $insertedCells, $col_key, $row_key, $pColumn, $pRow);

		foreach($values as $row_index=>$value_arr) {
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

		foreach($values as $row_index=>$valueArr) {
			$insertedCells->addInsertedCols($row_key, $col_key, count($valueArr) - 1);
			break;
		}
		$insertedCells->addInsertedRows($row_key, $col_key, count($values)-1);

		return $insertedCells;
	}

    /**
     * @param mixed $value
     * @return boolean
     * @throws Exception
     */
	private function _validateValue($value)
	{
		if ( ! is_array($value)) {
			throw new Exception('В классе '.ExcelParam::class.' поле "value" должно быть массивом, когда используется сеттер '.__CLASS__.'.');
		} else {
			foreach($value as $key=>$val) {
				if ( ! is_array($value[$key])) {
					throw new Exception('В классе '.ExcelParam::class.' поле "value" с ключом "'.$key.'" должно быть массивом, когда используется сеттер '.__CLASS__.'.');
				}
			}
		}
		return count($value) > 0;
	}

    /**
     * @param Worksheet $sheet
     * @param array $values
     * @param InsertedCells $insertedCells
     * @param integer $col_key Столбец таблицы, в котором была шаблонная переменная
     * @param integer $row_key Строка таблицы, в которой была шаблонная переменная
     * @param integer $pCol Текущий столбец таблицы
     * @param integer $pRow
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
	private function _insertNewRowsAndColsIfNeed(Worksheet $sheet, $values, $insertedCells, $col_key, $row_key, $pCol, $pRow)
	{
		$objReferenceHelper = ReferenceHelper::getInstance();

		foreach($values as $row_index=>$valueArr) {
			$colsToInsert = count($valueArr) - 1;
			//$highestColumn = Coordinate::columnIndexFromString($sheet->getHighestColumn());

			// Вставляю строку
			if ($row_index > 0) {
				$rowsToInsert = 1;
				$pCol1 = Coordinate::stringFromColumnIndex($pCol);
				$pCol2 = Coordinate::stringFromColumnIndex($pCol);
				$coordinate1 = $pCol1 . ($pRow+$row_index);
				$coordinate2 = $pCol2 . ($sheet->getHighestRow()+$rowsToInsert);
				$objReferenceHelper->insertNewBefore($coordinate1, 0, $rowsToInsert, $sheet, $coordinate2);
				//echo 'Row    '.$coordinate1.' '.$coordinate2."\n";
			}

			// Вставляю колонки
			$pCol1 = Coordinate::stringFromColumnIndex($pCol+1);
			$pCol2 = Coordinate::stringFromColumnIndex($pCol+$colsToInsert);
			$coordinate1 = $pCol1 . ($pRow+$row_index);
			$coordinate2 = $pCol2 . ($pRow+$row_index);
			$objReferenceHelper->insertNewBefore($coordinate1, $colsToInsert, 0, $sheet, $coordinate2);
			//echo 'Column '.$coordinate1.' '.$coordinate2.' '.$colsToInsert."\n";
		}
	}
}
