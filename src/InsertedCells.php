<?php

namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class InsertedCells
{
    public array $inserted_cols = [];
    public array $inserted_rows = [];

    /**
     * @param integer[][] $inserted_cols Inserted columns, where
     * key 1 - index of column,
     * key 2 - index of row,
     * value - count of inserted rows in the row of template variable
     *
     * @param integer[][] $inserted_rows Inserted rows, where
     * key 1 - index of row,
     * key 2 - index of column,
     * value - count of inserted columns in the column of template variable
     */
    public function __construct(array $inserted_cols = [], array $inserted_rows = [])
    {
        $this->inserted_rows = $inserted_rows;
        $this->inserted_cols = $inserted_cols;
    }

    /**
     * @param $rowKey integer Current index of row
     * @param $colKey integer Current index of column
     *
     * @return integer Total count of inserted rows, that were inserted from 0 row to current row
     */
    public function getInsertedRowsCount(int $rowKey, int $colKey): int
    {
        $this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $colKey, $rowKey);

        $result = 0;
        for ($i = 0; $i < $rowKey; $i++) {
            $result += $this->inserted_rows[$colKey][$i];
        }
        return $result;
    }

    /**
     * @param $rowKey integer Current index of row
     * @param $colKey integer Current index of column
     *
     * @return integer Total count of inserted columns, that were inserted from 0 column to current column
     */
    public function getInsertedColsCount(int $rowKey, int $colKey): int
    {
        $this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $rowKey, $colKey);

        $result = 0;
        for ($i = 0; $i < $colKey; $i++) {
            $result += $this->inserted_cols[$rowKey][$i];
        }
        return $result;
    }

    /**
     * @param $rowKey integer The row index, where was template variable
     * @param $colKey integer The column index, where was template variable
     * @param $insertedColsCount integer
     */
    public function setInsertedCols(int $rowKey, int $colKey, int $insertedColsCount): void
    {
        $this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $rowKey, $colKey);
        $this->inserted_cols[$rowKey][$colKey] = $insertedColsCount;
    }

    /**
     * @param $rowKey integer The row index, where was template variable
     * @param $colKey integer The column index, where was template variable
     * @param $insertedRowsCount integer
     */
    public function setInsertedRows(int $rowKey, int $colKey, int $insertedRowsCount): void
    {
        $this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $colKey, $rowKey);
        $this->inserted_rows[$colKey][$rowKey] = $insertedRowsCount;
    }

    /**
     * @param $rowKey integer The row index, where was template variable
     * @param $colKey integer The column index, where was template variable
     * @param $insertedColsCount integer
     */
    public function addInsertedCols(int $rowKey, int $colKey, int $insertedColsCount): void
    {
        $this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $rowKey, $colKey);
        $this->inserted_cols[$rowKey][$colKey] += $insertedColsCount;
    }

    public function addInsertedRows(int $rowKey, int $colKey, int $insertedRowsCount): void
    {
        $this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $colKey, $rowKey);
        $this->inserted_rows[$colKey][$rowKey] += $insertedRowsCount;
    }

    public function getInsertedRows(int $rowKey, int $colKey): int
    {
        $this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $colKey, $rowKey);
        return $this->inserted_rows[$colKey][$rowKey];
    }

    public function getInsertedCols(int $rowKey, int $colKey): int
    {
        $this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $rowKey, $colKey);
        return $this->inserted_cols[$rowKey][$colKey];
    }

    /**
     * Inserts indexes to specified array. The array filled with indexes from 0 to $key_index, if the index does not exist.
     */
    private function _addKeysToArrayIfNotExists(array &$array, int $keyIndex): void
    {
        for ($i = 0; $i <= $keyIndex; $i++) {
            if (!array_key_exists($i, $array)) {
                $array[$i] = 0;
            }
        }
    }

    /**
     * Inserts indexes to specified 2D array. The array filled with indexes from 0 to $key_index, if the indexes does not exist.
     */
    private function _addKeysTo2DArrayIfNotExists(array &$array, int $i_max, int $j_max): void
    {
        for ($i = 0; $i <= $i_max; $i++) {
            for ($j = 0; $j <= $j_max; $j++) {
                if (!array_key_exists($i, $array)) {
                    $array[$i] = [];
                }
                if (!array_key_exists($j, $array[$i])) {
                    $array[$i][$j] = 0;
                }
            }
        }
    }

    /**
     * @param $rowKey integer The row index, where was template variable
     * @param $colKey integer The column index, where was template variable
     *
     * @return integer Index of row, considering the count of inserted columns and rows
     */
    public function getCurrentRowIndex(int $rowKey, int $colKey): int
    {
        $inserted_rows = $this->getInsertedRowsCount($rowKey, $colKey);
        return $rowKey + 1 + $inserted_rows;
    }

    /**
     * @return integer Index of column, considering the count of inserted columns and rows
     */
    public function getCurrentColIndex(int $rowKey, int $colKey): int
    {
        $inserted_cols = $this->getInsertedColsCount($rowKey, $colKey);
        return $colKey + 1 + $inserted_cols;
    }

    /**
     * @return string Index of column (as a letter), considering the count of inserted columns and rows
     */
    public function getCurrentCol(int $rowKey, int $colKey): string
    {
        $col_index = $this->getCurrentColIndex($rowKey, $colKey);
        return Coordinate::stringFromColumnIndex($col_index);
    }

    /**
     * @return string Cell coordinate, considering the count of inserted columns and rows
     */
    public function getCurrentCellCoordinate(int $rowKey, int $colKey): string
    {
        $col       = $this->getCurrentCol($rowKey, $colKey);
        $row_index = $this->getCurrentRowIndex($rowKey, $colKey);
        return $col . $row_index;
    }
}
