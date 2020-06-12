<?php

namespace alhimik1986\PhpExcelTemplator\params;

class ExcelParam
{
	/**
	 * @var string The setter class name that will insert values into a table
	 */
	public $setterClass;

	/**
	 * @var mixed Value of the parameter
	 */
	public $value;

	/**
	 * @var callable Callback function to change style of table cell
	 */
	public $callback;

	/**
	 * @param string $setterClass
	 * @param mixed $value
	 * @param callable $callback
	 */
	public function __construct($setterClass, $value, callable $callback=null)
	{
		$this->setterClass = $setterClass;
		$this->value       = $value;
		$this->callback    = $callback;
	}
}
