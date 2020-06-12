<?php

namespace alhimik1986\PhpExcelTemplator\setters;

use Exception;
use alhimik1986\PhpExcelTemplator\InsertedCells;
use alhimik1986\PhpExcelTemplator\params\SetterParam;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;
use RuntimeException;

class CellSetterStringValue implements ICellSetter
{
    /**
     * {@inheritDoc}
     * @throws Exception
     */
	public function setCellValue(SetterParam $setterParam, InsertedCells $insertedCells): InsertedCells
    {
		$sheet = $setterParam->sheet;
		$row_key = $setterParam->row_key;
		$col_key = $setterParam->col_key;
		$current_col_content = $setterParam->col_content;
		$tpl_var_name = $setterParam->tpl_var_name;
		$param = $setterParam->params[$tpl_var_name];
		if ( ! $this->_validateValue($param->value)) {
			return $insertedCells;
		}

		$coordinate = $insertedCells->getCurrentCellCoordinate($row_key, $col_key);
		$col_value = strtr($current_col_content, [$tpl_var_name => $param->value]);
		$sheet->setCellValue($coordinate, $col_value);
		if ($param->callback) {
			$callbackParam = new CallbackParam([
				'sheet'        => $sheet,
                'coordinate'   => $coordinate,
                'param'        => $param->value,
				'tpl_var_name' => $tpl_var_name,
                'row_index'    => 0,
                'col_index'    => 0,
			]);
			call_user_func($param->callback, $callbackParam);
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
		if (is_array($value)) {
            throw new RuntimeException('In the '.ExcelParam::class.' class the field "value" must be an array, when the setter '.__CLASS__.' is used.');
		}
		return true;
	}
}
