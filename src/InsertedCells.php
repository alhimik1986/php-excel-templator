<?php

namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class InsertedCells
{
	/**
	 * @var integer[][] Вставленные строки, где
	 * ключ 1 - номер столбца, 
	 * ключ 2 - номер строки,
	 * значение - количество вставленных строк относительно этой строки
	 */
	public $inserted_rows;

	/**
	 * @var integer[][] Вставленные столбцы, где
	 * ключ 1 - номер строки,
	 * ключ 2 - номер столбца,
	 * значение - количество вставленных столбцов относительно этого столбца
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
	 * @param $row_key integer Текущая строка таблицы
	 * @param $col_key integer Текущая колонка таблицы
	 * @return integer Общее число вставленных строк,
	 * которые были вставлены от начала таблицы до текущей строки
	 */
	public function getInsertedRowsCount($row_key, $col_key)
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $col_key, $row_key);

		$result = 0;
		for($i=0; $i < $row_key; $i++) {
			$result += $this->inserted_rows[$col_key][$i];
		}
		return $result;
	}

	/**
	 * @param $row_key integer Текущая строка таблицы
	 * @param $col_key integer Текущая колонка таблицы
	 * @return integer Общее число столбцов,
	 * которые были вставлены от начала таблицы до текущего столбца
	 */
	public function getInsertedColsCount($row_key, $col_key)
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $row_key, $col_key);

		$result = 0;
		for($i=0; $i < $col_key; $i++) {
			$result += $this->inserted_cols[$row_key][$i];
		}
		return $result;
	}

	/**
	 * Подменяет добавленные столбцы относительно указанной строки и столбца
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @param $inserted_cols_count integer Количество добавленных столбцов
	 */
	public function setInsertedCols($row_key, $col_key, $inserted_cols_count)
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $row_key, $col_key);
		$this->inserted_cols[$row_key][$col_key] = $inserted_cols_count;
	}

	/**
	 * Подменяет добавленные строки относительно указанной строки
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @param $inserted_rows_count integer Количество добавленных строк
	 */
	public function setInsertedRows($row_key, $col_key, $inserted_rows_count)
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $col_key, $row_key);
		$this->inserted_rows[$col_key][$row_key] = $inserted_rows_count;
	}

	/**
	 * Регистрирует добавленные столбцы относительно указанной строки и столбца
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @param $inserted_cols_count integer Количество добавленных столбцов
	 */
	public function addInsertedCols($row_key, $col_key, $inserted_cols_count)
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $row_key, $col_key);
		$this->inserted_cols[$row_key][$col_key] += $inserted_cols_count;
	}

	/**
	 * Регистрирует добавленные строки относительно указанной строки
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @param $inserted_rows_count integer Количество добавленных строк
	 */
	public function addInsertedRows($row_key, $col_key, $inserted_rows_count)
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $col_key, $row_key);
		$this->inserted_rows[$col_key][$row_key] += $inserted_rows_count;
	}

	/**
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @return integer Число вставленных строк относительно указанной строки
	 */
	public function getInsertedRows($row_key, $col_key)
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_rows, $col_key, $row_key);
		return $this->inserted_rows[$col_key][$row_key];
	}

	/**
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @return integer Число вставленных столбцов относительно указанной строки и столбца
	 */
	public function getInsertedCols($row_key, $col_key)
	{
		$this->_addKeysTo2DArrayIfNotExists($this->inserted_cols, $row_key, $col_key);
		return $this->inserted_cols[$row_key][$col_key];
	}

	/**
	 * Вставляет индексы в указанный массив от нуля до указанного максимального индекса
	 * @param $array array Массив, в который нужно добавить индексы, если они не существуют
	 * @param $key_index integer Максимальный индекс
	 */
	private function _addKeysToArrayIfNotExists(&$array, $key_index)
	{
		for($i=0; $i <= $key_index; $i++) {
			if ( ! array_key_exists($i, $array)) {
				$array[$i] = 0;
			}
		}
	}

	/**
	 * Вставляет индексы в указанный массив от нуля до указанного максимального индекса массива и подмассива
	 * @param $array array Массив, в который нужно добавить индексы, если они не существуют
	 * @param $i_max integer Максимальный индекс массива
	 * @param $j_max integer Максимальный индекс подмассива
	 */
	private function _addKeysTo2DArrayIfNotExists(&$array, $i_max, $j_max)
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
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @return integer Строка ячейки с учётом вставленных столбцов и строк (отсчёт с нуля)
	 */
	public function getCurrentRowIndex($row_key, $col_key)
	{
		$inserted_rows = $this->getInsertedRowsCount($row_key, $col_key);
		return $row_key + 1 + $inserted_rows;
	}

	/**
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @return integer Столбец ячейки с учётом вставленных столбцов и строк (отсчёт с нуля)
	 */
	public function getCurrentColIndex($row_key, $col_key)
	{
		$inserted_cols = $this->getInsertedColsCount($row_key, $col_key);
		return $col_key + 1 + $inserted_cols;
	}

	/**
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @return string Столбец ячейки с учётом вставленных столбцов и строк (буквой)
	 */
	public function getCurrentCol($row_key, $col_key)
	{
		$col_index = $this->getCurrentColIndex($row_key, $col_key);
		return Coordinate::stringFromColumnIndex($col_index);
	}

	/**
	 * @param $row_key integer Строка таблицы, в которой была шаблонная переменная
	 * @param $col_key integer Столбец таблицы, в котором была шаблонная переменная
	 * @return string Координата ячейки с учётом вставленных столбцов и строк
	 */
	public function getCurrentCellCoordinate($row_key, $col_key)
	{
		$col = $this->getCurrentCol($row_key, $col_key);
		$row_index = $this->getCurrentRowIndex($row_key, $col_key);
		return $col . $row_index;
	}
}
