<?php

namespace alhimik1986\PhpExcelTemplator\params;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

class CallbackParam
{
	/**
	 * @var Worksheet
	 */
	public $sheet;

	/**
	 * @var string Current cell coordinate
	 */
	public $coordinate;

	/**
	 * @var mixed The value of the parameter, passed to the setter
	 */
	public $param;

	/**
	 * @var string The template variable name of the template file
	 */
	public $tpl_var_name;

	/**
	 * @var int Current index of the parameter. Scheme: ICellSetter->value[$row_index][$col_index]
	 */
	public $row_index;

	/**
	 * @var int Current subindex of the parameter. Scheme: ICellSetter->value[$row_index][$col_index]
	 */
	public $col_index;


    /**
     * @param array $params
     * @throws RuntimeException
     */
	public function __construct($params)
	{
		$fields = ['sheet', 'coordinate', 'param', 'tpl_var_name', 'row_index', 'col_index'];
		foreach($fields as $field) {
			if ( ! array_key_exists($field, $params)) {
				throw new RuntimeException('In the constructor of '.__CLASS__.' the parameter '.$field.' was not specified.');
			}
			$this->$field = $params[$field];
		}
	}
}
