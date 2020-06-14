<?php

namespace alhimik1986\PhpExcelTemplator\params;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

class SetterParam
{
	/**
	 * @var Worksheet
	 */
	public $sheet;

	/**
	 * @var string The name of the template variable in the template file
	 */
	public $tpl_var_name;

	/**
	 * @var ExcelParam[]
	 */
	public $params;

	/**
	 * @var string The row index, where was template variable
	 */
	public $row_key;

	/**
	 * @var string The column index, where was template variable
	 */
	public $col_key;

	/**
	 * @var string The cell content
	 */
	public $col_content;

    /**
     * @param array $params
     * @throws RuntimeException
     */
	public function __construct($params)
	{
		$fields = ['sheet', 'tpl_var_name', 'params', 'row_key', 'col_key', 'col_content'];
		foreach($fields as $field) {
			if ( ! array_key_exists($field, $params)) {
                throw new RuntimeException('In the constructor of '.__CLASS__.' the parameter '.$field.' was not specified.');
			}
			$this->$field = $params[$field];
		}
	}
}
