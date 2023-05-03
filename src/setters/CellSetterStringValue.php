<?php

namespace alhimik1986\PhpExcelTemplator\setters;

use alhimik1986\PhpExcelTemplator\InsertedCells;
use alhimik1986\PhpExcelTemplator\params\SetterParam;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;
use RuntimeException;

class CellSetterStringValue implements ICellSetter
{
    public function setCellValue(SetterParam $setterParam, InsertedCells $insertedCells): InsertedCells
    {
        $sheet             = $setterParam->sheet;
        $rowKey            = $setterParam->rowKey;
        $colKey            = $setterParam->colKey;
        $currentColContent = $setterParam->colContent;
        $tplVarName        = $setterParam->tplVarName;
        $param             = $setterParam->params[$tplVarName];
        if (!$this->_validateValue($param->value)) {
            return $insertedCells;
        }

        $coordinate = $insertedCells->getCurrentCellCoordinate($rowKey, $colKey);
        $col_value  = strtr($currentColContent, [$tplVarName => $param->value]);
        $sheet->setCellValue($coordinate, $col_value);
        if ($param->callback) {
            $callbackParam = new CallbackParam([
                'sheet'        => $sheet,
                'coordinate'   => $coordinate,
                'param'        => $param->value,
                'tpl_var_name' => $tplVarName,
                'row_index'    => 0,
                'col_index'    => 0,
            ]);
            call_user_func($param->callback, $callbackParam);
        }

        return $insertedCells;
    }

    /**
     * @param mixed $value
     *
     * @return bool
     */
    private function _validateValue($value): bool
    {
        if (is_array($value)) {
            throw new RuntimeException('In the ' . ExcelParam::class . ' class the field "value" must be an array, when the setter ' . __CLASS__ . ' is used.');
        }
        return true;
    }
}
