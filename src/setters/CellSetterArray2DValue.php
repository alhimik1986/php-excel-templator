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

        $pColumn = $insertedCells->getCurrentColIndex($rowKey, $colKey);
        $pRow    = $insertedCells->getCurrentRowIndex($rowKey, $colKey);
        $values  = $param->value;
        $this->_insertNewRowsAndColsIfNeed($sheet, $values, $pColumn, $pRow);

        foreach ($values as $rowIndex => $value_arr) {
            foreach ($value_arr as $colIndex => $value) {
                $pColumnWord         = Coordinate::stringFromColumnIndex($pColumn + $colIndex);
                $currCellCoordinates = $pColumnWord . ($pRow + $rowIndex);

                $sheet->setCellValue($currCellCoordinates, $value);
                if ($param->callback) {
                    $callbackParam = new CallbackParam([
                        'sheet'        => $sheet,
                        'coordinate'   => $currCellCoordinates,
                        'param'        => $param->value,
                        'tpl_var_name' => $tplVarName,
                        'row_index'    => $rowIndex,
                        'col_index'    => $colIndex,
                    ]);
                    call_user_func($param->callback, $callbackParam);
                }
            }
        }

        $firstValue = reset($values);
        $insertedCells->addInsertedCols($rowKey, $colKey, count($firstValue) - 1);
        $insertedCells->addInsertedRows($rowKey, $colKey, count($values) - 1);

        return $insertedCells;
    }

    private function _validateValue(array $value): bool
    {
        foreach ($value as $key => $val) {
            if (!is_array($val)) {
                throw new RuntimeException('In the ' . ExcelParam::class . ' class the field "value" with index "' . $key . '" must be an array, when the setter ' . __CLASS__ . ' is used.');
            }
        }
        return count($value) > 0;
    }

    /**
     * @param Worksheet $sheet
     * @param array     $values
     * @param integer   $pCol The current column index
     * @param integer   $pRow
     */
    private function _insertNewRowsAndColsIfNeed(Worksheet $sheet, array $values, int $pCol, int $pRow): void
    {
        $objReferenceHelper = ReferenceHelper::getInstance();

        foreach ($values as $rowIndex => $valueArr) {
            $rowIndex     = (int)$rowIndex;
            $colsToInsert = count($valueArr) - 1;
            //$highestColumn = Coordinate::columnIndexFromString($sheet->getHighestColumn());

            // Inserting row
            if ($rowIndex > 0) {
                $rowsToInsert = 1;
                $pCol1        = Coordinate::stringFromColumnIndex($pCol);
                $pCol2        = Coordinate::stringFromColumnIndex($pCol);
                $coordinate1  = $pCol1 . ($pRow + $rowIndex);
                $coordinate2  = $pCol2 . ($sheet->getHighestRow() + $rowsToInsert);
                $objReferenceHelper->insertNewBefore($coordinate1, 0, $rowsToInsert, $sheet, $coordinate2);
                //echo 'Row    '.$coordinate1.' '.$coordinate2."\n";
            }

            // Inserting columns
            $pCol_1      = Coordinate::stringFromColumnIndex($pCol + 1);
            $pCol_2      = Coordinate::stringFromColumnIndex($pCol + $colsToInsert);
            $coordinate1 = $pCol_1 . ($pRow + $rowIndex);
            $coordinate2 = $pCol_2 . ($pRow + $rowIndex);
            $objReferenceHelper->insertNewBefore($coordinate1, $colsToInsert, 0, $sheet, $coordinate2);
            //echo 'Column '.$coordinate1.' '.$coordinate2.' '.$colsToInsert."\n";
        }
    }
}
