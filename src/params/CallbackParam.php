<?php

namespace alhimik1986\PhpExcelTemplator\params;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

class CallbackParam
{
    public Worksheet $sheet;

    /**
     * Current cell coordinate
     */
    public string $coordinate;

    /**
     * @var mixed
     * The value of the parameter, passed to the setter
     */
    public $param;

	/**
	 * The template variable name of the template file
	 */
	public string $tpl_var_name;

	/**
	 * Current index of the parameter. Scheme: ICellSetter->value[$row_index][$col_index]
	 */
	public int $row_index;

    /**
     * Current subindex of the parameter. Scheme: ICellSetter->value[$row_index][$col_index]
     */
    public int $col_index;

	public function __construct(array $params)
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
