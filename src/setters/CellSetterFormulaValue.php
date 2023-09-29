<?php

namespace alhimik1986\PhpExcelTemplator\setters;

use alhimik1986\PhpExcelTemplator\InsertedCells;
use alhimik1986\PhpExcelTemplator\params\SetterParam;
use alhimik1986\PhpExcelTemplator\setters\DTO\FormulaValue;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class CellSetterFormulaValue implements ICellSetter
{

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

        /** @var FormulaValue $value */
        $value        = $param->value;

        $this->_insertNewRowsIfNeed($sheet, $value, $insertedCells, $rowKey, $pColumnIndex, $pRow);

        $formulaValue = $value->getFormula();
        $quantity = $value->getQuantity();

        for($i = 0; $i <= $quantity - 1; $i++) {
            $currCellCoordinates = $pColumn . ($pRow + $i);

            $value = $this->modifyFormula($rowKey + $i, $colKey, $formulaValue, $insertedCells);
            $sheet->setCellValue($currCellCoordinates, $value);
        }

        $maxInsertedRows = $this->_getMaxInsertedRows($rowKey, $insertedCells);
        $maxColumnIndex  = Coordinate::columnIndexFromString($sheet->getHighestColumn());
        for ($i = 0; $i <= $maxColumnIndex; $i++) {
            $insertedCells->setInsertedRows($rowKey, $i, $maxInsertedRows);
        }

        return $insertedCells;
    }


    private function _getMaxInsertedRows(int $rowKey, InsertedCells $insertedCells): int
    {
        $maxInsertedRows = 0;
        foreach ($insertedCells->inserted_rows as $col_key => $insertedRowsInCol) {
            $insertedRows = $insertedCells->getInsertedRows($rowKey, $col_key);
            if ($insertedRows > $maxInsertedRows) {
                $maxInsertedRows = $insertedRows;
            }
        }

        return $maxInsertedRows;
    }

    private function _validateValue($value): bool
    {
        return $value instanceof FormulaValue;
    }

    private function _insertNewRowsIfNeed(Worksheet $sheet, FormulaValue $value, InsertedCells $insertedCells, int $row_key, int $pColumnIndex, int $pRow): void
    {
        $maxInsertedRows = $this->_getMaxInsertedRows($row_key, $insertedCells);
        $rowsToInsert    = ($value->getQuantity() - 1 > $maxInsertedRows) ? $value->getQuantity() - 1 : 0;
        if ($rowsToInsert > 0) {
            $sheet->insertNewRowBefore($pRow + 1, $rowsToInsert);
        }
    }


    protected function modifyFormula($rowIndex, $columnIndex, $value, InsertedCells $insertedCells): string
    {
        $matches = [];
        preg_match_all("/(\(%[-,+]?\d,[-,+]?\d%\))/", $value, $matches);

        foreach ($matches[1] as $formula) {
            $modified = str_replace(['(%', '%)'], '', $formula);
            list($colModifier, $rowModifier) = explode(',', $modified);
            $colModifier = (int)$colModifier;
            $rowModifier = (int)$rowModifier;

            $finalCol = $columnIndex + $colModifier;
            $finalRow = $rowIndex + $rowModifier;
            $cellCoordinate = Coordinate::stringFromColumnIndex($finalCol + 1) . (string)($finalRow+1);

            $value = str_replace($formula, $cellCoordinate, $value);
        }

        return $value;
    }
}