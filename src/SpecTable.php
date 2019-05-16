<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class SpecTable
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
        static::contentNames($sheet);
        $sheet->getCell('P14')->setValue('Temporary Project_omIJ4');
        $sheet->getCell('P15')->setValue('Sculptform Click-on Battens');
        $sheet->getCell('P16')->setValue('Interior Only');
        $sheet->getCell('P17')->setValue("(32mm space), 42x42 Timber - Block, Spotted Gum,\n(32mm space), 42x32 Timber - Dome, Spotted Gum,\n(32mm space), 60x32 Timber - Dome, Spotted Gum");
        $sheet->getCell('P18')->setValue('32mm');
        $sheet->getCell('P19')->setValue('Spotted Gum');
        $sheet->getCell('P20')->setValue('Clear Oil');
        $sheet->getCell('P21')->setValue('Suspended Ceiling Track');
        $sheet->getCell('P22')->setValue('Matt black');
        $sheet->getCell('P23')->setValue('Yes');
    }

    private static function contentNames(Worksheet $sheet)
    {
        $sheet->getCell('K14')->setValue('PROJECT NAME');
        $sheet->getCell('K15')->setValue('PRODUCT');
        $sheet->getCell('K16')->setValue('APPLICATION TYPE');
        $sheet->getCell('K17')->setValue('SEQUENCE');
        $sheet->getCell('K18')->setValue('SPACING');
        $sheet->getCell('K19')->setValue('SPECIES');
        $sheet->getCell('K20')->setValue('COATING');
        $sheet->getCell('K21')->setValue('MOUNTING TRACK TYPE');
        $sheet->getCell('K22')->setValue('MOUNTING TRACK COLOR');
        $sheet->getCell('K23')->setValue('ACOUSTIC BACKING');
    }

    private static function styleContent(Worksheet $sheet)
    {
        $sheet->mergeCells('K14:O14'); $sheet->mergeCells('P14:T14');
        $sheet->mergeCells('K15:O15'); $sheet->mergeCells('P15:T15');
        $sheet->mergeCells('K16:O16'); $sheet->mergeCells('P16:T16');
        $sheet->mergeCells('K17:O17'); $sheet->mergeCells('P17:T17');
        $sheet->mergeCells('K18:O18'); $sheet->mergeCells('P18:T18');
        $sheet->mergeCells('K19:O19'); $sheet->mergeCells('P19:T19');
        $sheet->mergeCells('K20:O20'); $sheet->mergeCells('P20:T20');
        $sheet->mergeCells('K21:O21'); $sheet->mergeCells('P21:T21');
        $sheet->mergeCells('K22:O22'); $sheet->mergeCells('P22:T22');
        $sheet->mergeCells('K23:O23'); $sheet->mergeCells('P23:T23');

        $sheet
            ->getStyle('K14:T23')
            ->applyFromArray([
                'font' => [
                    'color' => ['argb' => '696969']
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['argb' => 'd4d4d4']
                    ]
                ],
            ]);

        $sheet
            ->getStyle('K14:O23')
            ->applyFromArray([
                'fill' => [
                    'fillType'   => Fill::FILL_SOLID,
                    'startColor' => ['argb' => 'f8f9f8']
                ]
            ]);

        $sheet->getStyle('K14')->getBorders()->getTop()->setBorderStyle(Border::BORDER_NONE);
        $sheet->getStyle('O14')->getBorders()->getTop()->setBorderStyle(Border::BORDER_NONE);
    }

    private static function createHeader(Worksheet $sheet)
    {
        $sheet->mergeCells('K13:T13');
        $sheet->getCell('K13')->setValue('Specification Table');
        $sheet->getStyle('K13')->applyFromArray(static::specTableHeaderStyles());
    }

    private static function specTableHeaderStyles()
    {
        return [
            'font' => [
                'bold'  => true,
                'name'  => 'Lucida Sans',
                // 'size'  => 15,
                'color' => ['argb' => Color::COLOR_WHITE]
            ],
            'alignment' => ['horizontal' => 'center'],
            // 'borders' => [
            //     'top'   => ['borderStyle' => Border::BORDER_THIN, 'color' => ['argb' => 'd4d4d4']],
            //     'right' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['argb' => 'd4d4d4']],
            //     'left'  => ['borderStyle' => Border::BORDER_THIN, 'color' => ['argb' => 'd4d4d4']],
            // ],
            'fill' => [
                'fillType'   => Fill::FILL_SOLID,
                'startColor' => ['argb' => '525453']
            ],
        ];
    }
}
