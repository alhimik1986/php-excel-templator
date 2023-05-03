<?php

namespace alhimik1986\PhpExcelTemplator\setters;

use alhimik1986\PhpExcelTemplator\InsertedCells;
use alhimik1986\PhpExcelTemplator\params\CallbackParam;
use alhimik1986\PhpExcelTemplator\params\SetterParam;
use alhimik1986\PhpExcelTemplator\ReferenceHelper;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class CellSetterArrayValue implements ICellSetter
{
    /**
     * @throws SpreadsheetException
     */
    public function setCellValue(SetterParam $setterParam, InsertedCells $insertedCells): InsertedCells
    {
        $sheet      = $setterParam->sheet;
        $rowKey     = $setterParam->rowKey;
        $colKey     = $setterParam->colKey;
        $tplVarName = $setterParam->tplVarName;
        $param      = $setterParam->params[$tplVarName];
        if (!$this->_validateValue($param->value)) {
            return $insertedCells;
        }

        $pColumn      = $insertedCells->getCurrentCol($rowKey, $colKey);
        $pColumnIndex = $insertedCells->getCurrentColIndex($rowKey, $colKey);
        $pRow         = $insertedCells->getCurrentRowIndex($rowKey, $colKey);
        $values       = $param->value;
        $this->_insertNewRowsIfNeed($sheet, $values, $pColumnIndex, $pRow);

        foreach ($values as $rowIndex => $value) {
            $currCellCoordinates = $pColumn . ($pRow + $rowIndex);

            $sheet->setCellValue($currCellCoordinates, $value);
            if ($param->callback) {
                $callbackParam = new CallbackParam([
                    'sheet'        => $sheet,
                    'coordinate'   => $currCellCoordinates,
                    'param'        => $param->value,
                    'tpl_var_name' => $tplVarName,
                    'row_index'    => $rowIndex,
                    'col_index'    => 0,
                ]);
                call_user_func($param->callback, $callbackParam);
            }
        }

        $insertedCells->addInsertedRows($rowKey, $colKey, count($values) - 1);

        return $insertedCells;
    }

    private function _validateValue(array $value): bool
    {
        return count($value) > 0;
    }

    /**
     * @param Worksheet $sheet
     * @param String[] $values
     * @param integer $pColumnIndex The current column index
     * @param integer $pRow The current row index
     */
    private function _insertNewRowsIfNeed(Worksheet $sheet, array $values, int $pColumnIndex, int $pRow): void
    {
        $rowsToInsert = count($values) - 1;
        if ($rowsToInsert > 0) {
            $objReferenceHelper = ReferenceHelper::getInstance();
            $pCol               = Coordinate::stringFromColumnIndex($pColumnIndex);
            $coordinate1        = $pCol . ($pRow + 1);
            $coordinate2        = $pCol . ($sheet->getHighestRow() + $rowsToInsert);
            $objReferenceHelper->insertNewBefore($coordinate1, 0, $rowsToInsert, $sheet, $coordinate2);
        }
    }
}
