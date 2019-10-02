<?php

namespace alhimik1986\PhpExcelTemplator\params;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Exception;

class SetterParam
{
	/**
	 * @var Worksheet Текущий лист таблицы
	 */
	public $sheet;

	/**
	 * @var string Имя шаблонной переменной в файле шаблона
	 */
	public $tpl_var_name;

	/**
	 * @var ExcelParam[] Параметры, передаваемые в сеттер
	 */
	public $params;

	/**
	 * @var string Индекс строки ячейки, где была шаблонная переменная
	 */
	public $row_key;

	/**
	 * @var string Индекс столбца ячейки, где была шаблонная переменная
	 */
	public $col_key;

	/**
	 * @var string Текущее содержимое ячейки таблицы
	 */
	public $col_content;

    /**
     * @param array $params
     * @throws Exception
     */
	public function __construct($params)
	{
		$fields = ['sheet', 'tpl_var_name', 'params', 'row_key', 'col_key', 'col_content'];
		foreach($fields as $field) {
			if ( ! array_key_exists($field, $params)) {
				throw new Exception('В конструкторе класса '.__CLASS__.' не был указан параметр '.$field.'.');
			}
			$this->$field = $params[$field];
		}
	}
}
