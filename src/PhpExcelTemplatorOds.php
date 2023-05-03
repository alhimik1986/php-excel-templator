<?php

namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception;
use PhpOffice\PhpSpreadsheet\Reader\Ods;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class PhpExcelTemplatorOds extends PhpExcelTemplator
{
    /**
     * @throws Exception
     */
    protected static function getSpreadsheet($templateFile): Spreadsheet
    {
        return (new Ods())->load($templateFile);
    }

    protected static function getWriter(Spreadsheet $spreadsheet): IWriter
    {
        return IOFactory::createWriter($spreadsheet, 'Ods');
    }
}
