<?php


namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Ods;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class PhpExcelTemplatorOds extends PhpExcelTemplator
{
    /**
     * {@inheritDoc}
     */
    protected static function getSpreadsheet($templateFile)
    {
        $reader = new Ods();
        return $reader->load($templateFile);
    }

    /**
     * {@inheritDoc}
     */
    protected static function getWriter(Spreadsheet $spreadsheet)
    {
        return IOFactory::createWriter($spreadsheet, 'Xlsx');
    }
}