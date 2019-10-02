<?php


namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Ods;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class PhpExcelTemplatorXls extends PhpExcelTemplator
{
    /**
     * {@inheritDoc}
     */
    protected static function getWriter(Spreadsheet $spreadsheet)
    {
        return IOFactory::createWriter($spreadsheet, 'Xls');
    }
}
