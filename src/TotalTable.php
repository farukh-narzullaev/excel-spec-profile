<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class TotalTable
{
    private static $pos;

    public static function create(Worksheet $sheet, $pos)
    {
        static::$pos = ($pos + 2);
        static::createHeader($sheet);

        static::$pos++;
        static::createContent($sheet);
        static::$pos += 2;

        return static::$pos;
    }

    private static function createContent(Worksheet $sheet)
    {
        static::styleContent($sheet);
        static::fillContent($sheet);
    }

    private static function fillContent(Worksheet $sheet)
    {
        $pos = static::$pos;
        $richText = new RichText();
        $richText->createText('');

        $price = $richText->createTextRun('$408.08 AUD');
        $price->getFont()->setBold(true);
        $price->getFont()->setSize(20);
        $price
            ->getFont()
            ->getColor()->setARGB('696969');

        $richText->createText("\r\n");
        $based = $richText->createTextRun(" based on [300+ sqm]");
        $based->getFont()->setSize(9);
        $based
            ->getFont()
            ->getColor()->setARGB('696969');

        $sheet->getCell("K{$pos}")->setValue($richText);

        $sheet->getCell("O{$pos}")->setValue('0.961');
        $sheet->getCell("R{$pos}")->setValue('15.9kg');
    }

    private static function styleContent(Worksheet $sheet)
    {
        $pos = static::$pos;
        $merge = ($pos + 2);

        $sheet->mergeCells("K{$pos}:N{$merge}");
        $sheet->mergeCells("O{$pos}:Q{$merge}");
        $sheet->mergeCells("R{$pos}:T{$merge}");

        $sheet
            ->getStyle("K{$pos}:T{$merge}")
            ->applyFromArray([
                'font' => [
                    'color' => ['argb' => '696969'],
                    'size'  => 20,
                    'bold'  => true,
                ],
                'alignment' => [
                    'horizontal' => 'center', 
                    'vertical'   => 'center',
                    'wrapText'   => true
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['argb' => 'd4d4d4']
                    ]
                ],
            ]);
    }

    private static function createHeader(Worksheet $sheet)
    {
        $pos = static::$pos;
        $sheet->mergeCells("K{$pos}:N{$pos}");
        $sheet->getCell("K{$pos}")->setValue('Supply Cost per m²');

        $sheet->mergeCells("O{$pos}:Q{$pos}");
        $sheet->getCell("O{$pos}")->setValue('Acoustic Rating*');

        $sheet->mergeCells("R{$pos}:T{$pos}");
        $sheet->getCell("R{$pos}")->setValue('Total Weight');

        $sheet->getStyle("K{$pos}")->applyFromArray(static::headerStyles());
        $sheet->getStyle("O{$pos}")->applyFromArray(static::headerStyles());
        $sheet->getStyle("R{$pos}")->applyFromArray(static::headerStyles());
    }

    private static function headerStyles()
    {
        return [
            'font' => [
                'bold'  => true,
                'name'  => 'Lucida Sans',
                'color' => ['argb' => Color::COLOR_WHITE]
            ],
            'alignment' => ['horizontal' => 'center'],
            'fill' => [
                'fillType'   => Fill::FILL_SOLID,
                'startColor' => ['argb' => '525453']
            ],
        ];
    }
}
