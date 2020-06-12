<?php

namespace alhimik1986\PhpExcelTemplator;

use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use alhimik1986\PhpExcelTemplator\setters\CellSetterStringValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArrayValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArray2DValue;
use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\SetterParam;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class PhpExcelTemplator
{
    public const BEFORE_INSERT_PARAMS = 'BeforeInsertParams';
    public const AFTER_INSERT_PARAMS  = 'AfterInsertParams';
    public const BEFORE_SAVE          = 'BeforeSave';

    /**
     * @param string $templateFile Path to *.xlsx template file
     * @param string $outputFile Exported file path
     * @param ExcelParam[] | array $params Parameters of the setter
     * @param callable[] $callbacks An associative array of callbacks to change cell styles without using setters
     * @param callable[] $events Events, applied for additional manipulations with the spreadsheet and the writer
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
	public static function outputToFile($templateFile, $outputFile, $params, $callbacks=[], $events=[]): void
	{
		$spreadsheet = static::getSpreadsheet($templateFile);
		$sheet = $spreadsheet->getActiveSheet();
		$templateVarsArr = $sheet->toArray();
		static::renderWorksheet($sheet, $templateVarsArr, $params, $callbacks, $events);
		static::outputSpreadsheetToFile($spreadsheet, $outputFile, $events);
	}

    /**
     * @param string $templateFile Path to *.xlsx template file
     * @param string $outputFile Exported file path
     * @param ExcelParam[] | array $params Parameters of the setter
     * @param array $callbacks An associative array of callbacks to change cell styles without using setters
     * @param callable[] $events Events, applied for additional manipulations with the spreadsheet and the writer
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
	public static function saveToFile($templateFile, $outputFile, $params, $callbacks=[], $events=[]): void
	{
        $spreadsheet = static::getSpreadsheet($templateFile);
		$sheet = $spreadsheet->getActiveSheet();
		$templateVarsArr = $sheet->toArray();
		static::renderWorksheet($sheet, $templateVarsArr, $params, $callbacks, $events);
		static::saveSpreadsheetToFile($spreadsheet, $outputFile, $events);
	}

    /**
     * @param string $templateFile Path to *.xlsx template file
     * @return Spreadsheet
     */
	protected static function getSpreadsheet($templateFile): Spreadsheet
    {
        return IOFactory::load($templateFile);
    }

    /**
     * @param Worksheet $sheet The sheet, which contains the template variables
     * @param array $templateVarsArr An array of cells contained in the template file
     * @param ExcelParam[] | array $params Parameters of the setter
     * @param array $callbacks An associative array of callbacks to change cell styles without using setters
     * @param callable[] $events Events, applied for additional manipulations with the spreadsheet and the writer
     * @return Worksheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
	public static function renderWorksheet(Worksheet $sheet, $templateVarsArr, $params, $callbacks=[], $events=[]): Worksheet
	{
		$params = static::getCorrectedParams($params, $callbacks);
		static::clearTemplateVarsInSheet($sheet, $templateVarsArr, $params);

		if (isset($events[self::BEFORE_INSERT_PARAMS]) && is_callable($events[self::BEFORE_INSERT_PARAMS])) {
            $events[self::BEFORE_INSERT_PARAMS]($sheet, $templateVarsArr);
        }

		static::insertParams($sheet, $templateVarsArr, $params);

        if (isset($events[self::AFTER_INSERT_PARAMS]) && is_callable($events[self::AFTER_INSERT_PARAMS])) {
            $events[self::AFTER_INSERT_PARAMS]($sheet, $templateVarsArr);
        }

		return $sheet;
	}

	/**
     * If the params are an array, it will be converted to ExcelParam with the corresponding setter.
	 * @param ExcelParam[] | array $params
	 * @param array $callbacks
	 * @return ExcelParam[]
	 */
	protected static function getCorrectedParams($params, $callbacks): array
	{
		foreach($params as $key=>$param) {
			if ( ! $param instanceof ExcelParam) {
				$setterClass = CellSetterStringValue::class;
				$callback = array_key_exists($key, $callbacks) ? $callbacks[$key] : static function(){};

				if (is_array($param)) {
					$valueArr = reset($param);
					$setterClass = is_array($valueArr)
						? CellSetterArray2DValue::class
						: CellSetterArrayValue::class;
				}

				$params[$key] = new ExcelParam($setterClass, $param, $callback);
			}
		}

		return $params;
	}

    /**
     * Exports the spreadsheet to download it as a file.
     * @param Spreadsheet $spreadsheet
     * @param string $outputFile
     * @param callable[] $events
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
	public static function outputSpreadsheetToFile(Spreadsheet $spreadsheet, $outputFile, $events=[]): void
	{
        $writer = static::getWriter($spreadsheet);
		Calculation::getInstance($spreadsheet)->clearCalculationCache();
		self::setHeaders(basename($outputFile));

        if (isset($events[self::BEFORE_SAVE]) && is_callable($events[self::BEFORE_SAVE])) {
            $events[self::BEFORE_SAVE]($spreadsheet, $writer);
        }

		$writer->save('php://output');
	}

    /**
     * Saves the spreadsheet as a file.
     * @param Spreadsheet $spreadsheet
     * @param string $outputFile
     * @param callable[] $events
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
	public static function saveSpreadsheetToFile(Spreadsheet $spreadsheet, $outputFile, $events=[]): void
	{
		$writer = static::getWriter($spreadsheet);
		Calculation::getInstance($spreadsheet)->clearCalculationCache();

        if (isset($events['beforeSave']) && is_callable($events['beforeSave'])) {
            $events['beforeSave']($spreadsheet, $writer);
        }

		$writer->save($outputFile);
	}

    /**
     * @param Spreadsheet $spreadsheet
     * @return IWriter
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
	protected static function getWriter(Spreadsheet $spreadsheet): IWriter
    {
        return IOFactory::createWriter($spreadsheet, 'Xlsx');
    }

	/**
	 * Sets the header parameters needed to download the excel file.
	 * @param string $fileName
	 */
	protected static function setHeaders($fileName): void
	{
		header('Content-Disposition: attachment; filename="'.$fileName);
		header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Type: text/html; charset=windows-1251;');
		header('Pragma: public');

		header('Content-Transfer-Encoding: binary');
		header('Cache-Control: must-revalidate');
	}

	/**
	 * Clears template variables in the sheet of the template file
	 * @param Worksheet $sheet
	 * @param array $templateVars Cells array of the template file
	 * @param ExcelParam[] $params
	 */
	protected static function clearTemplateVarsInSheet(Worksheet $sheet, $templateVars, $params): void
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
     * Inserts values to cells instead of template variables
     * @param Worksheet $sheet
     * @param array $templateVarsArr
     * @param ExcelParam[] $params
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     */
	public static function insertParams(Worksheet $sheet, $templateVarsArr, $params): void
	{
		$insertedCells = new InsertedCells();
		foreach($templateVarsArr as $row_key=>$row) {
			foreach($row as $col_key => $col_content) {
				$colVarNames = self::_getTemplateVarsFromString($col_content, $params);
				foreach($colVarNames as $tpl_var_name) {
					$setterClass = $params[$tpl_var_name]->setterClass;
					$setter = new $setterClass();
					$setterParam = new SetterParam(['sheet'=>$sheet, 'tpl_var_name'=>$tpl_var_name, 'params'=>$params, 'row_key'=>$row_key, 'col_key'=>$col_key, 'col_content'=>$col_content]);
					$insertedCells = $setter->setCellValue($setterParam, $insertedCells);

                    // After inserting value to the cell, I get the content of this cell to insert another value
                    // (when the cell has some template variables)
					if (count($colVarNames) > 1) {
						$coordinate = $insertedCells->getCurrentCellCoordinate($row_key, $col_key);
                        /** @noinspection NullPointerExceptionInspection */
                        $col_content = $sheet->getCell($coordinate)->getValue();
					}
				}
			}
		}
	}

	/**
	 * @param string $string The content of the cell, that may contain a template variable
	 * @param ExcelParam[] $params
	 * @return String[] Template variables in a string
	 */
	private static function _getTemplateVarsFromString($string, $params): array
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
