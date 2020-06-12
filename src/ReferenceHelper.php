<?php

namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;

class ReferenceHelper extends \PhpOffice\PhpSpreadsheet\ReferenceHelper
{
    /**
     * Instance of this class.
     *
     * @var ReferenceHelper
     */
    private static $instance;

    /**
     * Get an instance of this class.
     *
     * @return ReferenceHelper
     */
    public static function getInstance(): ReferenceHelper
    {
        if (!isset(self::$instance) || (self::$instance === null)) {
            self::$instance = new self();
        }

        return self::$instance;
    }

    /**
     * Insert a new column or row, updating all possible related data.
     *
     * @param string $pBefore Insert before this cell address (e.g. 'A1')
     * @param string $pAfter Insert until this cell address (e.g. 'C1')
     * @param int $pNumCols Number of columns to insert/delete (negative values indicate deletion)
     * @param int $pNumRows Number of rows to insert/delete (negative values indicate deletion)
     * @param Worksheet $pSheet The worksheet that we're editing
     *
     * @throws SpreadsheetException
     */
    public function insertNewBefore($pBefore, $pNumCols, $pNumRows, Worksheet $pSheet, $pAfter=null): void
    {
        $remove = ($pNumCols < 0 || $pNumRows < 0);
        $allCoordinates = $pSheet->getCoordinates();

        // Get coordinate of $pBefore
        list($beforeColumn, $beforeRow) = Coordinate::coordinateFromString($pBefore);
        $beforeColumnIndex = Coordinate::columnIndexFromString($beforeColumn);

        // Clear cells if we are removing columns or rows
        $highestColumn = $pSheet->getHighestColumn();
        $highestRow = $pSheet->getHighestRow();

//******************************* My changes ********************************************
        // Get coordinate of $pAfter
        if ($pAfter !== null) {
            list($afterColumn, $afterRow) = Coordinate::coordinateFromString($pAfter);
            $afterColumnIndex = Coordinate::columnIndexFromString($afterColumn);

            $highestColumn = Coordinate::stringFromColumnIndex($afterColumnIndex);
            $highestRow = $afterRow+1;
        }
//***************************************************************************************

        // 1. Clear column strips if we are removing columns
        if ($pNumCols < 0 && $beforeColumnIndex - 2 + $pNumCols > 0) {
            for ($i = 1; $i <= $highestRow - 1; ++$i) {
                for ($j = $beforeColumnIndex - 1 + $pNumCols; $j <= $beforeColumnIndex - 2; ++$j) {
                    $coordinate = Coordinate::stringFromColumnIndex($j + 1) . $i;
                    $pSheet->removeConditionalStyles($coordinate);
                    if ($pSheet->cellExists($coordinate)) {
                        $pSheet->getCell($coordinate)->setValueExplicit('', DataType::TYPE_NULL);
                        $pSheet->getCell($coordinate)->setXfIndex(0);
                    }
                }
            }
        }

        // 2. Clear row strips if we are removing rows
        if ($pNumRows < 0 && $beforeRow - 1 + $pNumRows > 0) {
            for ($i = $beforeColumnIndex - 1; $i <= Coordinate::columnIndexFromString($highestColumn) - 1; ++$i) {
                for ($j = $beforeRow + $pNumRows; $j <= $beforeRow - 1; ++$j) {
                    $coordinate = Coordinate::stringFromColumnIndex($i + 1) . $j;
                    $pSheet->removeConditionalStyles($coordinate);
                    if ($pSheet->cellExists($coordinate)) {
                        $pSheet->getCell($coordinate)->setValueExplicit('', DataType::TYPE_NULL);
                        $pSheet->getCell($coordinate)->setXfIndex(0);
                    }
                }
            }
        }

        // Loop through cells, bottom-up, and change cell coordinate
        if ($remove) {
            // It's faster to reverse and pop than to use unshift, especially with large cell collections
            $allCoordinates = array_reverse($allCoordinates);
        }
        while ($coordinate = array_pop($allCoordinates)) {
            $cell = $pSheet->getCell($coordinate);
            $cellIndex = Coordinate::columnIndexFromString($cell->getColumn());

            if ($cellIndex - 1 + $pNumCols < 0) {
                continue;
            }

            // New coordinate
            $newCoordinate = Coordinate::stringFromColumnIndex($cellIndex + $pNumCols) . ($cell->getRow() + $pNumRows);

            // Should the cell be updated? Move value and cellXf index from one cell to another.
            if (($cellIndex >= $beforeColumnIndex) && ($cell->getRow() >= $beforeRow) &&
                ($cellIndex <= $afterColumnIndex) && ($cell->getRow() <= $afterRow)
            ) {
                // Update cell styles
                $pSheet->getCell($newCoordinate)->setXfIndex($cell->getXfIndex());

                // Insert this cell at its new location
                if ($cell->getDataType() == DataType::TYPE_FORMULA) {
                    // Formula should be adjusted
                    $pSheet->getCell($newCoordinate)
                        ->setValue($this->updateFormulaReferences($cell->getValue(), $pBefore, $pNumCols, $pNumRows, $pSheet->getTitle()));
                } else {
                    // Formula should not be adjusted
                    $pSheet->getCell($newCoordinate)->setValue($cell->getValue());
                }

                // Clear the original cell
                $pSheet->getCellCollection()->delete($coordinate);
            } else {
                /*    We don't need to update styles for rows/columns before our insertion position,
                        but we do still need to adjust any formulae    in those cells                    */
                if ($cell->getDataType() == DataType::TYPE_FORMULA) {
                    // Formula should be adjusted
                    $cell->setValue($this->updateFormulaReferences($cell->getValue(), $pBefore, $pNumCols, $pNumRows, $pSheet->getTitle()));
                }
            }
        }

        // Duplicate styles for the newly inserted cells
//******************************* My changes ********************************************
        //$highestColumn = $pSheet->getHighestColumn();
        //$highestRow = $pSheet->getHighestRow();
//***************************************************************************************

        if ($pNumCols > 0 && $beforeColumnIndex - 2 > 0) {
            for ($i = $beforeRow; $i <= $highestRow - 1; ++$i) {
                // Style
                $coordinate = Coordinate::stringFromColumnIndex($beforeColumnIndex - 1) . $i;
                if ($pSheet->cellExists($coordinate)) {
                    $xfIndex = $pSheet->getCell($coordinate)->getXfIndex();
                    $conditionalStyles = $pSheet->conditionalStylesExists($coordinate) ?
                        $pSheet->getConditionalStyles($coordinate) : false;
                    for ($j = $beforeColumnIndex; $j <= $beforeColumnIndex - 1 + $pNumCols; ++$j) {
                        $pSheet->getCellByColumnAndRow($j, $i)->setXfIndex($xfIndex);
                        if ($conditionalStyles) {
                            $cloned = [];
                            foreach ($conditionalStyles as $conditionalStyle) {
                                $cloned[] = clone $conditionalStyle;
                            }
                            $pSheet->setConditionalStyles(Coordinate::stringFromColumnIndex($j) . $i, $cloned);
                        }
                    }
                }
            }
        }

        if ($pNumRows > 0 && $beforeRow - 1 > 0) {
            for ($i = $beforeColumnIndex; $i <= Coordinate::columnIndexFromString($highestColumn); ++$i) {
                // Style
                $coordinate = Coordinate::stringFromColumnIndex($i) . ($beforeRow - 1);
                if ($pSheet->cellExists($coordinate)) {
                    $xfIndex = $pSheet->getCell($coordinate)->getXfIndex();
                    $conditionalStyles = $pSheet->conditionalStylesExists($coordinate) ?
                        $pSheet->getConditionalStyles($coordinate) : false;
                    for ($j = $beforeRow; $j <= $beforeRow - 1 + $pNumRows; ++$j) {
                        $pSheet->getCell(Coordinate::stringFromColumnIndex($i) . $j)->setXfIndex($xfIndex);
                        if ($conditionalStyles) {
                            $cloned = [];
                            foreach ($conditionalStyles as $conditionalStyle) {
                                $cloned[] = clone $conditionalStyle;
                            }
                            $pSheet->setConditionalStyles(Coordinate::stringFromColumnIndex($i) . $j, $cloned);
                        }
                    }
                }
            }
        }

        // Update worksheet: column dimensions
        $this->adjustColumnDimensions($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: row dimensions
        $this->adjustRowDimensions($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        //    Update worksheet: page breaks
        $this->adjustPageBreaks($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        //    Update worksheet: comments
        $this->adjustComments($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: hyperlinks
        $this->adjustHyperlinks($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: data validations
        $this->adjustDataValidations($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: merge cells
        $this->adjustMergeCells($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: protected cells
        $this->adjustProtectedCells($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: autofilter
        $autoFilter = $pSheet->getAutoFilter();
        $autoFilterRange = $autoFilter->getRange();
        if (!empty($autoFilterRange)) {
            if ($pNumCols != 0) {
                $autoFilterColumns = $autoFilter->getColumns();
                if (count($autoFilterColumns) > 0) {
                    sscanf($pBefore, '%[A-Z]%d', $column, $row);
                    $columnIndex = Coordinate::columnIndexFromString($column);
                    list($rangeStart, $rangeEnd) = Coordinate::rangeBoundaries($autoFilterRange);
                    if ($columnIndex <= $rangeEnd[0]) {
                        if ($pNumCols < 0) {
                            //    If we're actually deleting any columns that fall within the autofilter range,
                            //        then we delete any rules for those columns
                            $deleteColumn = $columnIndex + $pNumCols - 1;
                            $deleteCount = abs($pNumCols);
                            for ($i = 1; $i <= $deleteCount; ++$i) {
                                if (isset($autoFilterColumns[Coordinate::stringFromColumnIndex($deleteColumn + 1)])) {
                                    $autoFilter->clearColumn(Coordinate::stringFromColumnIndex($deleteColumn + 1));
                                }
                                ++$deleteColumn;
                            }
                        }
                        $startCol = ($columnIndex > $rangeStart[0]) ? $columnIndex : $rangeStart[0];

                        //    Shuffle columns in autofilter range
                        if ($pNumCols > 0) {
                            $startColRef = $startCol;
                            $endColRef = $rangeEnd[0];
                            $toColRef = $rangeEnd[0] + $pNumCols;

                            do {
                                $autoFilter->shiftColumn(Coordinate::stringFromColumnIndex($endColRef), Coordinate::stringFromColumnIndex($toColRef));
                                --$endColRef;
                                --$toColRef;
                            } while ($startColRef <= $endColRef);
                        } else {
                            //    For delete, we shuffle from beginning to end to avoid overwriting
                            $startColID = Coordinate::stringFromColumnIndex($startCol);
                            $toColID = Coordinate::stringFromColumnIndex($startCol + $pNumCols);
                            $endColID = Coordinate::stringFromColumnIndex($rangeEnd[0] + 1);
                            do {
                                $autoFilter->shiftColumn($startColID, $toColID);
                                ++$startColID;
                                ++$toColID;
                            } while ($startColID != $endColID);
                        }
                    }
                }
            }
            $pSheet->setAutoFilter($this->updateCellReference($autoFilterRange, $pBefore, $pNumCols, $pNumRows));
        }

        // Update worksheet: freeze pane
        if ($pSheet->getFreezePane()) {
            $splitCell = $pSheet->getFreezePane();
            $topLeftCell = $pSheet->getTopLeftCell();

            $splitCell = $this->updateCellReference($splitCell, $pBefore, $pNumCols, $pNumRows);
            $topLeftCell = $this->updateCellReference($topLeftCell, $pBefore, $pNumCols, $pNumRows);

            $pSheet->freezePane($splitCell, $topLeftCell);
        }

        // Page setup
        if ($pSheet->getPageSetup()->isPrintAreaSet()) {
            $pSheet->getPageSetup()->setPrintArea($this->updateCellReference($pSheet->getPageSetup()->getPrintArea(), $pBefore, $pNumCols, $pNumRows));
        }

        // Update worksheet: drawings
        $aDrawings = $pSheet->getDrawingCollection();
        foreach ($aDrawings as $objDrawing) {
            $newReference = $this->updateCellReference($objDrawing->getCoordinates(), $pBefore, $pNumCols, $pNumRows);
            if ($objDrawing->getCoordinates() != $newReference) {
                $objDrawing->setCoordinates($newReference);
            }
        }

        // Update workbook: named ranges
        if (count($pSheet->getParent()->getNamedRanges()) > 0) {
            foreach ($pSheet->getParent()->getNamedRanges() as $namedRange) {
                if ($namedRange->getWorksheet()->getHashCode() == $pSheet->getHashCode()) {
                    $namedRange->setRange($this->updateCellReference($namedRange->getRange(), $pBefore, $pNumCols, $pNumRows));
                }
            }
        }

        // Garbage collect
        $pSheet->garbageCollect();
    }
}
