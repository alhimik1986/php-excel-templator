<?php

namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use alhimik1986\PhpExcelTemplator\setters\CellSetterStringValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArrayValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArray2DValue;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\SetterParam;

class PhpExcelTemplator
{
	/**
	 * @param string $templateFile Путь к файлу шаблона
	 * @param string $fileName Имя экспортируемого файла
	 * @param ExcelParam[] | array $params Параметры, передаваемые в сеттер
	 */
	public static function outputToFile($templateFile, $fileName, $params)
	{
		$spreadsheet = IOFactory::load($templateFile);
		$sheet = $spreadsheet->getActiveSheet();
		$templateVarsArr = $sheet->toArray();
		static::renderWorksheet($sheet, $templateVarsArr, $params);
		static::outputSpreadsheetToFile($spreadsheet, $fileName);
	}

	/**
	 * @param string $templateFile Путь к файлу шаблона
	 * @param string $fileName Имя экспортируемого файла
	 * @param ExcelParam[] | array $params Параметры, передаваемые в сеттер
	 */
	public static function saveToFile($templateFile, $fileName, $params)
	{
		$spreadsheet = IOFactory::load($templateFile);
		$sheet = $spreadsheet->getActiveSheet();
		$templateVarsArr = $sheet->toArray();
		static::renderWorksheet($sheet, $templateVarsArr, $params);
		static::saveSpreadsheetToFile($spreadsheet, $fileName);
	}

	/**
	 * @param Worksheet $sheet Лист, в котором хранятся шаблонные переменные
	 * @param array $templateVarsArr Массив ячеек, содержащийся в таблице шаблона
	 * @param ExcelParam[] | array $params Параметры, передаваемые в сеттер
	 * @return Worksheet
	 */
	public static function renderWorksheet(Worksheet $sheet, $templateVarsArr, $params)
	{
		$params = static::getCorrectedParams($params);
		static::clearTemplateVarsInSheet($sheet, $templateVarsArr, $params);
		static::insertParams($sheet, $templateVarsArr, $params);
		return $sheet;
	}

	/**
	 * Проверяю передаваемые параметры и, если они не являются объектами ExcelParam,
	 * а представляют собой просто массив, то создаю для них ExcelParam
	 * с соответствующим сеттером.
	 * @param ExcelParam[] | array $params Параметры, которым нужно присвоить
	 * соответствующий сеттер, если он не задан
	 * @return ExcelParam[] Скорректированные параметры
	 */
	protected static function getCorrectedParams($params)
	{
		foreach($params as $key=>$param) {
			if ( ! $param instanceof ExcelParam) {
				$setterClass = CellSetterStringValue::class;

				if (is_array($param)) {
					$valueArr = reset($param);
					$setterClass = is_array($valueArr)
						? CellSetterArray2DValue::class
						: CellSetterArrayValue::class;
				}

				$params[$key] = new ExcelParam($setterClass, $param);
			}
		}

		return $params;
	}

	/**
	 * Вывести файл для скачивания.
	 * @param Spreadsheet $spreadsheet
	 * @param string $fileName Имя файла
	 */
	public static function outputSpreadsheetToFile(Spreadsheet $spreadsheet, $fileName)
	{
		$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
		self::setHeaders($fileName);
		$writer->save('php://output');
		exit;
	}

	/**
	 * Сохранить в файл.
	 * @param Spreadsheet $spreadsheet
	 * @param string $fileName Имя файла
	 */
	public static function saveSpreadsheetToFile(Spreadsheet $spreadsheet, $fileName)
	{
		$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
		$writer->save(basename($fileName));
		exit;
	}

	/**
	 * Устанавливаю параметры header, необходимые для скачивания excel-файла.
	 * @param string $fileName - Имя файла 
	 */
	protected static function setHeaders($fileName)
	{
		header('Content-Disposition: attachment; filename="'.$fileName);
		header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Type: text/html; charset=windows-1251;');
		header('Pragma: public');

		header('Content-Transfer-Encoding: binary');
		header('Cache-Control: must-revalidate');
	}

	/**
	 * Очищает шаблонные переменные в файле шаблона
	 * @param Worksheet $sheet Лист в excel
	 * @param array $templateVars Содержимое файла шаблона
	 * @param ExcelParam[] $params Параметры, передаваемые в сеттер
	 */
	protected static function clearTemplateVarsInSheet(Worksheet $sheet, $templateVars, $params)
	{
		$paramKeys = array_keys($params);
		foreach($templateVars as $row_key=>$row) {
			foreach($row as $col_key=>$col_content) {
				if ($col_content) {
					foreach($paramKeys as $paramKey) {
						if (strpos($col_content, $paramKey) !== false) {
							$sheet->setCellValueExplicitByColumnAndRow(
								$col_key+1,
								$row_key+1,
								null,
								DataType::TYPE_NULL
							);
						}
					}
				}
			}
		}
	}

	/**
	 * Вставляет параметры в указанные шаблонные переменные
	 * @param Worksheet $sheet
	 * @param array $templateVarsArr Массив ячеек, содержащийся в таблице шаблона
	 * @param ExcelParam $params Параметры, передаваемые в сеттер
	 */
	public static function insertParams(Worksheet $sheet, $templateVarsArr, $params)
	{
		$insertedCells = new InsertedCells();
		foreach($templateVarsArr as $row_key=>$row) {
			foreach($row as $col_key=>$col_content) {
				$colVarNames = self::_getTemplateVarsFromString($col_content, $params);
				foreach($colVarNames as $tpl_var_name) {
					$setterClass = $params[$tpl_var_name]->setterClass;
					$setter = new $setterClass();
					$setterParam = new SetterParam(['sheet'=>$sheet, 'tpl_var_name'=>$tpl_var_name, 'params'=>$params, 'row_key'=>$row_key, 'col_key'=>$col_key, 'col_content'=>$col_content]);
					$insertedCells = $setter->setCellValue($setterParam, $insertedCells);

					// Чтобы можно было использовать несколько шаблонных переменных в одной ячейке таблицы
					if (count($colVarNames) > 1) {
						$coordinate = $insertedCells->getCurrentCellCoordinate($row_key, $col_key);
						$col_content = $sheet->getCell($coordinate)->getValue();
					}
				}
			}
		}
	}

	/**
	 * @param string $string Содержимое ячейки таблицы, в которой может находиться шаблонная переменная
	 * @param ExcelParam[] $params Параметры, передаваемые в сеттер
	 * @return String[] Шаблонные переменные в строке
	 */
	private static function _getTemplateVarsFromString($string, $params)
	{
		$result = [];
		$paramKeys = array_keys($params);

		foreach($paramKeys as $paramKey) {
			if (strpos($string, $paramKey) !== false) {
				$result[] = $paramKey;
			}
		}

		return $result;
	}
}