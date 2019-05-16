<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class TotalTable
{
    public static function create(Worksheet $sheet)
    {
        static::createHeader($sheet);
        static::createContent($sheet);
    }

    private static function createContent(Worksheet $sheet)
    {
        static::styleContent($sheet);
        static::fillContent($sheet);
    }

    private static function fillContent(Worksheet $sheet)
    {
        $richText = new RichText();
        $richText->createText('');
        
        $price = $richText->createTextRun('$408.08 AUD');
        $price->getFont()->setBold(true);
        $price->getFont()->setSize(20);
        $price
            ->getFont()
            ->getColor()->setARGB('696969');

        $based = $richText->createTextRun("\n based on [300+ sqm]");
        $based->getFont()->setSize(9);
        $based
            ->getFont()
            ->getColor()->setARGB('696969');

        $sheet->getCell('K26')->setValue($richText);

        $sheet->getCell('O26')->setValue('0.961');
        $sheet->getCell('R26')->setValue('15.9kg');
    }

    private static function styleContent(Worksheet $sheet)
    {
        $sheet->mergeCells('K26:N28');
        $sheet->mergeCells('O26:Q28');
        $sheet->mergeCells('R26:T28');

        $sheet
            ->getStyle('K26:T28')
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
        $sheet->mergeCells('K25:N25');
        $sheet->getCell('K25')->setValue('Supply Cost per m²');

        $sheet->mergeCells('O25:Q25');
        $sheet->getCell('O25')->setValue('Acoustic Rating*');

        $sheet->mergeCells('R25:T25');
        $sheet->getCell('R25')->setValue('Total Weight');

        $sheet->getStyle('K25')->applyFromArray(static::headerStyles());
        $sheet->getStyle('O25')->applyFromArray(static::headerStyles());
        $sheet->getStyle('R25')->applyFromArray(static::headerStyles());
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
