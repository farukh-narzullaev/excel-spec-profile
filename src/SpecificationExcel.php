<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Cell\Hyperlink;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\RichText\RichText;

class SpecificationExcel
{
    protected $spreadsheet;

    private $output;

    public function __construct($output)
    {
        $this->output = $output;
    }

    public function generate()
    {
        $sheet = $this->createSheet();
        Header::create($sheet);
        ProjectImage::create($sheet);
        SpecTable::create($sheet);

        $writer = new Xlsx($this->spreadsheet);
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
