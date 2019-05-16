<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Footer
{
    public static function create(Worksheet $sheet)
    {
        $sheet->mergeCells('A58:F59');
        $sheet->mergeCells('H58:M59');
        $sheet->mergeCells('O58:T59');

        $sheet->mergeCells('G58:G59');
        $sheet->mergeCells('N58:N59');

        static::separator($sheet, 'G58');
        static::separator($sheet, 'N58');

        static::contacts($sheet, 'A58');
        static::email($sheet, 'H58');
        static::website($sheet, 'O58');

        $sheet->getStyle('A58:T59')->applyFromArray([
            'font' => [
                'bold'  => true,
                'name'  => 'Lucida Sans',
                'color' => ['argb' => Color::COLOR_WHITE]
            ],
            'alignment' => ['horizontal' => 'center', 'vertical' => 'center'],
            'fill' => [
                'fillType'   => Fill::FILL_SOLID,
                'startColor' => ['argb' => '525453']
            ]
        ]);
    }

    private static function separator(Worksheet $sheet, $cell)
    {
        $sheet->getCell($cell)->setValue('|');
        $sheet->getStyle($cell)->applyFromArray([
            'alignment' => ['horizontal' => 'center', 'vertical' => 'center'],
        ]);
    }

    private static function contacts(Worksheet $sheet, $cell)
    {
        //Contact us on 1800 008 828
        $richText = new RichText();
        $richText->createTextRun("Contact us on ")->getFont()->setSize(18)->getColor()->setARGB('FFFFFF');
        $richText->createTextRun("1800 008 828")->getFont()->setBold(true)->setSize(18)->getColor()->setARGB('FFFFFF');

        $sheet->getCell($cell)->setValue($richText);
    }

    private static function email(Worksheet $sheet, $cell)
    {
        $richText = new RichText();
        $richText->createTextRun("support@sculptform.com.au")->getFont()->setSize(18)->getColor()->setARGB('FFFFFF');

        $sheet->getCell($cell)->setValue($richText);
    }

    private static function website(Worksheet $sheet, $cell)
    {
        $richText = new RichText();
        $richText->createTextRun("sculptform.com.au")->getFont()->setBold(true)->setSize(18)->getColor()->setARGB('FFFFFF');

        $sheet->getCell($cell)->setValue($richText);
    }
}
