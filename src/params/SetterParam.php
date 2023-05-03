<?php

namespace alhimik1986\PhpExcelTemplator\params;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

class SetterParam
{
    public Worksheet $sheet;

    public string $tplVarName;

    /**
     * @var ExcelParam[]
     */
    public array $params;

    /**
     * The row index, where was template variable
     */
    public int $rowKey;

    /**
     * The column index, where was template variable
     */
    public int $colKey;

    /**
     * The cell content
     */
    public string $colContent;

    public function __construct(array $params)
    {
        $fields = ['sheet', 'tplVarName', 'params', 'rowKey', 'colKey', 'colContent'];
        foreach ($fields as $field) {
            if (!array_key_exists($field, $params)) {
                throw new RuntimeException('In the constructor of ' . __CLASS__ . ' the parameter ' . $field . ' was not specified.');
            }
            $this->$field = $params[$field];
        }
    }
}
