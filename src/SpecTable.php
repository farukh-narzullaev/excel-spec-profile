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
        $sheet->getCell('M13')->setValue('Temporary Project_omIJ4');
        $sheet->getCell('M14')->setValue('Sculptform Click-on Battens');
        $sheet->getCell('M15')->setValue('Interior Only');
        $sheet->getCell('M16')->setValue("(32mm space), 42x42 Timber - Block, Spotted Gum,\n(32mm space), 42x32 Timber - Dome, Spotted Gum,\n(32mm space), 60x32 Timber - Dome, Spotted Gum");
        $sheet->getCell('M17')->setValue('32mm');
        $sheet->getCell('M18')->setValue('Spotted Gum');
        $sheet->getCell('M19')->setValue('Clear Oil');
        $sheet->getCell('M20')->setValue('Suspended Ceiling Track');
        $sheet->getCell('M21')->setValue('Matt black');
        $sheet->getCell('M22')->setValue('Yes');
    }

    private static function contentNames(Worksheet $sheet)
    {
        $sheet->getCell('I13')->setValue('PROJECT NAME');
        $sheet->getCell('I14')->setValue('PRODUCT');
        $sheet->getCell('I15')->setValue('APPLICATION TYPE');
        $sheet->getCell('I16')->setValue('SEQUENCE');
        $sheet->getCell('I17')->setValue('SPACING');
        $sheet->getCell('I18')->setValue('SPECIES');
        $sheet->getCell('I19')->setValue('COATING');
        $sheet->getCell('I20')->setValue('MOUNTING TRACK TYPE');
        $sheet->getCell('I21')->setValue('MOUNTING TRACK COLOR');
        $sheet->getCell('I22')->setValue('ACOUSTIC BACKING');

    }

    private static function styleContent(Worksheet $sheet)
    {
        $sheet->mergeCells('I13:L13'); $sheet->mergeCells('M13:P13');
        $sheet->mergeCells('I14:L14'); $sheet->mergeCells('M14:P14');
        $sheet->mergeCells('I15:L15'); $sheet->mergeCells('M15:P15');
        $sheet->mergeCells('I16:L16'); $sheet->mergeCells('M16:P16');
        $sheet->mergeCells('I17:L17'); $sheet->mergeCells('M17:P17');
        $sheet->mergeCells('I18:L18'); $sheet->mergeCells('M18:P18');
        $sheet->mergeCells('I19:L19'); $sheet->mergeCells('M19:P19');
        $sheet->mergeCells('I20:L20'); $sheet->mergeCells('M20:P20');
        $sheet->mergeCells('I21:L21'); $sheet->mergeCells('M21:P21');
        $sheet->mergeCells('I22:L22'); $sheet->mergeCells('M22:P22');

        $sheet
            ->getStyle('I13:P22')
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
            ->getStyle('I13:I22')
            ->applyFromArray([
                'fill' => [
                    'fillType'   => Fill::FILL_SOLID,
                    'startColor' => ['argb' => 'f8f9f8']
                ]
            ]);

        $sheet->getStyle('I13')->getBorders()->getTop()->setBorderStyle(Border::BORDER_NONE);
        $sheet->getStyle('M13')->getBorders()->getTop()->setBorderStyle(Border::BORDER_NONE);
    }

    private static function createHeader(Worksheet $sheet)
    {
        $sheet->mergeCells('I12:P12');
        $sheet->getCell('I12')->setValue('Specification Table');
        $sheet->getStyle('I12')->applyFromArray(static::specTableHeaderStyles());
    }

    private static function specTableHeaderStyles()
    {
        return [
            'font' => [
                'bold'  => true,
                'name'  => 'Arial',
                // 'size'  => 15,
                'color' => ['argb' => Color::COLOR_WHITE]
            ],
            'alignment' => ['horizontal' => 'center'],
            'borders' => [
                'top'   => ['borderStyle' => Border::BORDER_THIN, 'color' => ['argb' => 'd4d4d4']],
                'right' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['argb' => 'd4d4d4']],
                'left'  => ['borderStyle' => Border::BORDER_THIN, 'color' => ['argb' => 'd4d4d4']],
            ],
            'fill' => [
                'fillType'   => Fill::FILL_SOLID,
                'startColor' => ['argb' => '525453']
            ],
        ];
    }
}
