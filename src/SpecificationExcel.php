<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class SpecificationExcel
{
    protected $spreadsheet;

    private $output;

    public function __construct($output)
    {
        $this->output = $output;
    }

    public function generate($version)
    {
        $sheet = $this->createSheet();
        Header::create($sheet, $version);
        ProjectImage::create($sheet, $version);

        SpecTable::create($sheet);
        TotalTable::create($sheet);
        Note::create($sheet);
        Steps::create($sheet);
        Footer::create($sheet);

        $sheet->getProtection()->setSheet(true);

        $writer = new Xlsx($this->spreadsheet);
        $writer->setOffice2003Compatibility(true);
        $writer->save($this->output);
    }

    private function createSheet($title = "Specification Table")
    {
        $this->spreadsheet = new Spreadsheet();
        $sheet = $this->spreadsheet->getActiveSheet();
        $sheet->setTitle($title);

        return $sheet;
    }
}
