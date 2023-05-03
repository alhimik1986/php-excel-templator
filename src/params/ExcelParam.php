<?php

namespace alhimik1986\PhpExcelTemplator\params;

class ExcelParam
{
	/**
	 * @var string The setter class name that will insert values into a table
	 */
	public string $setterClass;

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
	 * @param callable|null $callback
	 */
	public function __construct(string $setterClass, $value, callable $callback=null)
	{
		$this->setterClass = $setterClass;
		$this->value       = $value;
		$this->callback    = $callback;
	}
}
