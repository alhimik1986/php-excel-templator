<?php


namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Ods;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class PhpExcelTemplatorOds extends PhpExcelTemplator
{
    /**
     * {@inheritDoc}
     */
    protected static function getSpreadsheet($templateFile): Spreadsheet
    {
        $reader = new Ods();
        return $reader->load($templateFile);
    }

    /**
     * {@inheritDoc}
     */
    protected static function getWriter(Spreadsheet $spreadsheet): IWriter
    {
        return IOFactory::createWriter($spreadsheet, 'Ods');
    }
}
