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

class CellSetterArrayValue implements ICellSetter
{
    /**
     * {@inheritdoc}
     * @throws \PhpOffice\PhpSpreadsheet\Exception
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

		$pColumn = $insertedCells->getCurrentCol($row_key, $col_key);
		$pColumnIndex = $insertedCells->getCurrentColIndex($row_key, $col_key);
		$pRow = $insertedCells->getCurrentRowIndex($row_key, $col_key);
		$values = $param->value;
		$this->_insertNewRowsIfNeed($sheet, $values, $insertedCells, $col_key, $row_key, $pColumnIndex, $pRow);

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
     * @throws Exception
     */
	private function _validateValue($value)
	{
		if ( ! is_array($value)) {
			throw new Exception('В классе '.ExcelParam::class.' поле "value" должно быть массивом, когда используется сеттер '.__CLASS__.'.');
		}
		return count($value) > 0;
	}

    /**
     * @param Worksheet $sheet
     * @param String[] $values
     * @param InsertedCells $insertedCells
     * @param integer $col_key Столбец таблицы, в котором была шаблонная переменная
     * @param integer $row_key Строка таблицы, в которой была шаблонная переменная
     * @param integer $pColumnIndex Текущий столбец таблицы
     * @param integer $pRow Текущая строка таблицы
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
	private function _insertNewRowsIfNeed(Worksheet $sheet, $values, $insertedCells, $col_key, $row_key, $pColumnIndex, $pRow)
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
