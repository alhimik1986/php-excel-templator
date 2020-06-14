<?php

namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class InsertedCells
{
	/**
	 * @var integer[][] Inserted rows, where
	 * key 1 - index of column,
	 * key 2 - index of row,
	 * value - count of inserted rows in the row of template variable
	 */
	public $inserted_rows;

	/**
	 * @var integer[][] Inserted columns, where
	 * ключ 1 - index of row,
	 * ключ 2 - index of column,
     * value - count of inserted columns in the column of template variable
	 */
	public $inserted_cols;

	/**
	 * @param integer[] $inserted_cols
	 * @param integer[] $inserted_rows
	 */
	public function __construct($inserted_cols=[], $inserted_rows=[])
	{
		$this->inserted_cols = $inserted_cols;
		$this->inserted_rows = $inserted_rows;
	}

	/**
	 * @param $row_key integer Current index of row
	 * @param $col_key integer Current index of column
	 * @return integer Total count of inserted rows, that were inserted from 0 row to current row
	 */
	public function getInsertedRowsCount($row_key, $col_key): int
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $col_key, $row_key);

		$result = 0;
		for($i=0; $i < $row_key; $i++) {
			$result += $this->inserted_rows[$col_key][$i];
		}
		return $result;
	}

	/**
	 * @param $row_key integer Current index of row
	 * @param $col_key integer Current index of column
	 * @return integer Total count of inserted columns, that were inserted from 0 column to current column
	 */
	public function getInsertedColsCount($row_key, $col_key): int
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $row_key, $col_key);

		$result = 0;
		for($i=0; $i < $col_key; $i++) {
			$result += $this->inserted_cols[$row_key][$i];
		}
		return $result;
	}

	/**
	 * @param $row_key integer The row index, where was template variable
	 * @param $col_key integer The column index, where was template variable
	 * @param $inserted_cols_count integer
	 */
	public function setInsertedCols($row_key, $col_key, $inserted_cols_count): void
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $row_key, $col_key);
		$this->inserted_cols[$row_key][$col_key] = $inserted_cols_count;
	}

	/**
	 * @param $row_key integer The row index, where was template variable
	 * @param $col_key integer The column index, where was template variable
	 * @param $inserted_rows_count integer
	 */
	public function setInsertedRows($row_key, $col_key, $inserted_rows_count): void
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $col_key, $row_key);
		$this->inserted_rows[$col_key][$row_key] = $inserted_rows_count;
	}

	/**
	 * @param $row_key integer The row index, where was template variable
	 * @param $col_key integer The column index, where was template variable
	 * @param $inserted_cols_count integer
	 */
	public function addInsertedCols($row_key, $col_key, $inserted_cols_count): void
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $row_key, $col_key);
		$this->inserted_cols[$row_key][$col_key] += $inserted_cols_count;
	}

	/**
	 * @param $row_key integer
	 * @param $col_key integer
	 * @param $inserted_rows_count integer
	 */
	public function addInsertedRows($row_key, $col_key, $inserted_rows_count): void
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $col_key, $row_key);
		$this->inserted_rows[$col_key][$row_key] += $inserted_rows_count;
	}

	/**
	 * @param $row_key integer
	 * @param $col_key integer
	 * @return integer
	 */
	public function getInsertedRows($row_key, $col_key): int
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $col_key, $row_key);
		return $this->inserted_rows[$col_key][$row_key];
	}

	/**
	 * @param $row_key integer
	 * @param $col_key integer
	 * @return integer
	 */
	public function getInsertedCols($row_key, $col_key): int
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $row_key, $col_key);
		return $this->inserted_cols[$row_key][$col_key];
	}

	/**
     * Inserts indexes to specified array. The array filled with indexes from 0 to $key_index, if the index does not exist.
	 * @param $array array
	 * @param $key_index integer
	 */
	private function _addKeysToArrayIfNotExists(&$array, $key_index): void
	{
		for($i=0; $i <= $key_index; $i++) {
			if ( ! array_key_exists($i, $array)) {
				$array[$i] = 0;
			}
		}
	}

	/**
     * Inserts indexes to specified 2D array. The array filled with indexes from 0 to $key_index, if the indexes does not exist.
	 * @param $array array
	 * @param $i_max integer
	 * @param $j_max integer
	 */
	private function _addKeysTo2DArrayIfNotExists(&$array, $i_max, $j_max): void
	{
		for($i=0; $i <= $i_max; $i++) {
			for($j=0; $j <= $j_max; $j++) {
				if ( ! array_key_exists($i, $array)) {
					$array[$i] = [];
				}
				if ( ! array_key_exists($j, $array[$i])) {
					$array[$i][$j] = 0;
				}
			}
		}
	}

	/**
	 * @param $row_key integer The row index, where was template variable
	 * @param $col_key integer The column index, where was template variable
	 * @return integer Index of row, considering the count of inserted columns and rows
	 */
	public function getCurrentRowIndex($row_key, $col_key): int
	{
		$inserted_rows = $this->getInsertedRowsCount($row_key, $col_key);
		return $row_key + 1 + $inserted_rows;
	}

	/**
	 * @param $row_key integer
	 * @param $col_key integer
	 * @return integer Index of column, considering the count of inserted columns and rows
	 */
	public function getCurrentColIndex($row_key, $col_key): int
	{
		$inserted_cols = $this->getInsertedColsCount($row_key, $col_key);
		return $col_key + 1 + $inserted_cols;
	}

	/**
	 * @param $row_key integer
	 * @param $col_key integer
	 * @return string Index of column (as a letter), considering the count of inserted columns and rows
	 */
	public function getCurrentCol($row_key, $col_key): string
	{
		$col_index = $this->getCurrentColIndex($row_key, $col_key);
		return Coordinate::stringFromColumnIndex($col_index);
	}

	/**
	 * @param $row_key integer
	 * @param $col_key integer
	 * @return string Cell coordinate, considering the count of inserted columns and rows
	 */
	public function getCurrentCellCoordinate($row_key, $col_key): string
	{
		$col = $this->getCurrentCol($row_key, $col_key);
		$row_index = $this->getCurrentRowIndex($row_key, $col_key);
		return $col . $row_index;
	}
}
