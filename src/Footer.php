<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Footer
{
    private static $pos;

    public static function create(Worksheet $sheet, $pos)
    {
        static::$pos = ($pos + 9);
        $pos = static::$pos;
        $merge = $pos + 1;

        $sheet->mergeCells("A{$pos}:F{$merge}");
        $sheet->mergeCells("H{$pos}:M{$merge}");
        $sheet->mergeCells("O{$pos}:T{$merge}");

        $sheet->mergeCells("G{$pos}:G{$merge}");
        $sheet->mergeCells("N{$pos}:N{$merge}");

        static::separator($sheet, "G{$pos}");
        static::separator($sheet, "N{$pos}");

        static::contacts($sheet, "A{$pos}");
        static::email($sheet, "H{$pos}");
        static::website($sheet, "O{$pos}");

        $sheet->getStyle("A{$pos}:T{$merge}")->applyFromArray([
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
