<?php

namespace alhimik1986\PhpExcelTemplator;

use alhimik1986\PhpExcelTemplator\params\ExcelParam;
use alhimik1986\PhpExcelTemplator\params\SetterParam;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArray2DValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterArrayValue;
use alhimik1986\PhpExcelTemplator\setters\CellSetterStringValue;
use alhimik1986\PhpExcelTemplator\setters\ICellSetter;
use Exception;
use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class PhpExcelTemplator
{
    public const BEFORE_INSERT_PARAMS = 'BeforeInsertParams';
    public const AFTER_INSERT_PARAMS  = 'AfterInsertParams';
    public const BEFORE_SAVE          = 'BeforeSave';

    /**
     * @param string                             $templateFile Path to *.xlsx template file
     * @param string                             $outputFile Exported file path
     * @param ExcelParam[]|array<string, string>|array<string, array<string, string>>|array<string, array<string, array<string, string>>> $params Parameters of the setter
     * @param array<string, callable>            $callbacks An associative array of callbacks to change cell styles without using setters
     * @param array<string, callable|mixed>      $events Events, applied for additional manipulations with the spreadsheet and the writer
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function outputToFile(string $templateFile, string $outputFile, array $params, array $callbacks = [], array $events = []): void
    {
        $spreadsheet     = static::getSpreadsheet($templateFile);
        $sheet           = $spreadsheet->getActiveSheet();
        $templateVarsArr = $sheet->toArray();
        static::renderWorksheet($sheet, $templateVarsArr, $params, $callbacks, $events);
        static::outputSpreadsheetToFile($spreadsheet, $outputFile, $events);
    }

    /**
     * @param string                        $templateFile Path to *.xlsx template file
     * @param string                        $outputFile Exported file path
     * @param ExcelParam[]|array<string, string>|array<string, array<string, string>>|array<string, array<string, array<string, string>>> $params Parameters of the setter
     * @param array<string, callable>       $callbacks An associative array of callbacks to change cell styles without using setters
     * @param array<string, callable|mixed> $events Events, applied for additional manipulations with the spreadsheet and the writer
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function saveToFile(string $templateFile, string $outputFile, array $params, array $callbacks = [], array $events = []): void
    {
        $spreadsheet     = static::getSpreadsheet($templateFile);
        $sheet           = $spreadsheet->getActiveSheet();
        $templateVarsArr = $sheet->toArray();
        static::renderWorksheet($sheet, $templateVarsArr, $params, $callbacks, $events);
        static::saveSpreadsheetToFile($spreadsheet, $outputFile, $events);
    }

    /**
     * @param string $templateFile Path to *.xlsx template file
     *
     * @return Spreadsheet
     */
    protected static function getSpreadsheet(string $templateFile): Spreadsheet
    {
        return IOFactory::load($templateFile);
    }

    /**
     * @param Worksheet                      $sheet The sheet, which contains the template variables
     * @param array<int, array<int, string>> $templateVarsArr An array of cells contained in the template file
     * @param ExcelParam[]|array<string, string>|array<string, array<string, string>>|array<string, array<string, array<string, string>>> $params Parameters of the setter
     * @param array<string, callable>        $callbacks An associative array of callbacks to change cell styles without using setters
     * @param array<string, callable|mixed>  $events Events, applied for additional manipulations with the spreadsheet and the writer
     *
     * @return Worksheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     */
    public static function renderWorksheet(Worksheet $sheet, array $templateVarsArr, array $params, array $callbacks = [], array $events = []): Worksheet
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
     *
     * @param ExcelParam[]|array<string, string>|array<string, array<string, string>>|array<string, array<string, array<string, string>>> $params Parameters of the setter
     * @param array<string, callable>                                                                                      $callbacks
     *
     * @return array<string, ExcelParam>
     */
    protected static function getCorrectedParams(array $params, array $callbacks): array
    {
        $result = [];
        foreach ($params as $key => $param) {
            if (!$param instanceof ExcelParam) {
                $setterClass = CellSetterStringValue::class;
                $callback    = array_key_exists($key, $callbacks) ? $callbacks[$key] : static function () {
                };

                if (is_array($param)) {
                    $valueArr    = reset($param);
                    $setterClass = is_array($valueArr)
                        ? CellSetterArray2DValue::class
                        : CellSetterArrayValue::class;
                }

                $result[$key] = new ExcelParam($setterClass, $param, $callback);
            } else {
                $result[$key] = $param;
            }
        }

        return $result;
    }

    /**
     * Exports the spreadsheet to download it as a file.
     *
     * @param Spreadsheet                   $spreadsheet
     * @param string                        $outputFile
     * @param array<string, callable|mixed> $events
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public static function outputSpreadsheetToFile(Spreadsheet $spreadsheet, string $outputFile, array $events = []): void
    {
        $writer = static::getWriter($spreadsheet);
        Calculation::getInstance($spreadsheet)->clearCalculationCache();
        static::setHeaders(basename($outputFile));

        if (isset($events[self::BEFORE_SAVE]) && is_callable($events[self::BEFORE_SAVE])) {
            $events[self::BEFORE_SAVE]($spreadsheet, $writer);
        }

        $writer->save('php://output');
    }

    /**
     * Saves the spreadsheet as a file.
     *
     * @param Spreadsheet                   $spreadsheet
     * @param string                        $outputFile
     * @param array<string, callable|mixed> $events
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public static function saveSpreadsheetToFile(Spreadsheet $spreadsheet, string $outputFile, array $events = []): void
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
     *
     * @return IWriter
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected static function getWriter(Spreadsheet $spreadsheet): IWriter
    {
        return IOFactory::createWriter($spreadsheet, 'Xlsx');
    }

    /**
     * Sets the header parameters needed to download the Excel file.
     *
     * @param string $fileName
     */
    protected static function setHeaders(string $fileName): void
    {
        header('Content-Disposition: attachment; filename="' . $fileName . '"');
        header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Pragma: public');

        header('Content-Transfer-Encoding: binary');
        header('Cache-Control: must-revalidate');
    }

    /**
     * Clears template variables in the sheet of the template file
     *
     * @param Worksheet                      $sheet
     * @param array<int, array<int, string>> $templateVars Cells array of the template file
     * @param array<string, ExcelParam>      $params
     */
    protected static function clearTemplateVarsInSheet(Worksheet $sheet, array $templateVars, array $params): void
    {
        $paramKeys = array_keys($params);
        foreach ($templateVars as $row_key => $row) {
            foreach ($row as $col_key => $col_content) {
                if ($col_content) {
                    foreach ($paramKeys as $paramKey) {
                        if (strpos($col_content, $paramKey) !== false) {
                            $sheet->setCellValueExplicitByColumnAndRow(
                                $col_key + 1,
                                $row_key + 1,
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
     *
     * @param Worksheet                      $sheet
     * @param array<int, array<int, string>> $templateVarsArr
     * @param array<string, ExcelParam>      $params
     *
     * @throws Exception
     */
    public static function insertParams(Worksheet $sheet, array $templateVarsArr, array $params): void
    {
        $insertedCells = new InsertedCells();
        foreach ($templateVarsArr as $rowKey => $row) {
            foreach ($row as $colKey => $colContent) {
                $colVarNames = self::_getTemplateVarsFromString($colContent, $params);
                foreach ($colVarNames as $tplVarName) {

                    $setterClass = $params[$tplVarName]->setterClass;
                    /** @var ICellSetter $setter */
                    $setter = new $setterClass();

                    $setterParam = new SetterParam([
                        'sheet'        => $sheet,
                        'tplVarName'   => $tplVarName,
                        'params'       => $params,
                        'rowKey'       => $rowKey,
                        'colKey'      => $colKey,
                        'colContent'  => $colContent
                    ]);

                    $insertedCells = $setter->setCellValue($setterParam, $insertedCells);

                    // After inserting value to the cell, I get the content of this cell to insert another value
                    // (when the cell has some template variables)
                    if (count($colVarNames) > 1) {
                        $coordinate  = $insertedCells->getCurrentCellCoordinate($rowKey, $colKey);
                        $colContent = $sheet->getCell($coordinate)->getValue();
                    }
                }
            }
        }
    }

    /**
     * @param ?string                   $string The content of the cell, that may contain a template variable
     * @param array<string, ExcelParam> $params
     *
     * @return String[] Template variables in a string
     */
    private static function _getTemplateVarsFromString(?string $string, array $params): array
    {
        $result    = [];
        $paramKeys = array_keys($params);

        foreach ($paramKeys as $paramKey) {
            if ($string !== null && strpos($string, $paramKey) !== false) {
                $result[] = $paramKey;
            }
        }

        return $result;
    }
}
