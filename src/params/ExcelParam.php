<?php

namespace alhimik1986\PhpExcelTemplator\params;

class ExcelParam
{
	/**
	 * @var string Класс сеттера, который вставляет значения определённого типа в таблицу
	 */
	public $setterClass;

	/**
	 * @var mixed Значение для параметра
	 */
	public $value;

	/**
	 * @var callable Функция обратного вызова, необходимая для форматирования ячейки таблицы
	 */
	public $callback;

	/**
	 * @param string $setterClass Класс сеттера, который вставляет значения определённого типа в таблицу
	 * @param mixed $value Значение для параметра
	 * @param callable $callback Функция обратного вызова, необходимая для форматирования ячейки таблицы
	 */
	public function __construct($setterClass, $value, callable $callback=null)
	{
		$this->setterClass = $setterClass;
		$this->value       = $value;
		$this->callback    = $callback;
	}
}
