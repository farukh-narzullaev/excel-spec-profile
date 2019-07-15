<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class SpecTable
{
    protected static $str = "(32mm space), 42x42 Timber - Block, Spotted Gum,
(32mm space), 42x32 Timber - Dome, Spotted Gum,
(32mm space), 60x32 Timber - Dome, Spotted Gum,
32mm space), 32x60 Timber - Block, White
Oak,(32mm space), 60x32 Timber - Flute, White
Oak,(32mm space), 42x42 Timber - Block, White
Oak,(32mm space), 60x32 Timber - Dome, White
Oak,(32mm space), 60x19 Timber - Block, White
Oak,(32mm space), 32x32 Timber - Flute, White Oak";

    protected static $sequenceContentLines = 0;


    protected static $cells = [
        'name'        => null,
        'product'     => null,
        'app_type'    => null,
        'sequence'    => null,
        'profile'     => null,
        'spacing'     => null,
        'species'     => null,
        'coating'     => null,
        'track_type'  => null,
        'track_color' => null,
        'backing'     => null,
    ];

    public static function create(Worksheet $sheet)
    {
        static::$sequenceContentLines = count(explode(PHP_EOL, static::$str));

        $pos = static::setCells();

        static::createHeader($sheet);
        static::createContent($sheet);

        return $pos;
    }

    private static function setCells()
    {
        static::$cells['name'] = [
            'key' => 'K14', 'value' => 'P14', 'keyMerge' => 'K14:O14', 'valueMerge' => 'P14:T14',
            'title' => 'project name', 'content' => 'Temporary Project_omIJ4',
        ];

        static::$cells['product'] = [
            'key' => 'K15', 'value' => 'P15', 'keyMerge' => 'K15:O15', 'valueMerge' => 'P15:T15',
            'title' => 'product', 'content' => 'Sculptform Click-on Battens',
        ];

        $pos = 15;
        $pos++;
        static::$cells['app_type'] = [
            'key' => "K{$pos}",
            'value' => "P{$pos}",
            'title' => 'application type',
            'content' => 'Interior Only',
            'keyMerge' => "K{$pos}:O{$pos}",
            'valueMerge' => "P{$pos}:T{$pos}",
        ];

        $pos++;
        $offset = ($pos + static::$sequenceContentLines * 2 + 3);
        static::$cells['sequence'] = [
            'key' => "K{$pos}",
            'value' => "P{$pos}",
            'title' => 'sequence',
            'content' => static::$str,
            'keyMerge' => "K{$pos}:O{$offset}",
            'valueMerge' => "P{$pos}:T{$offset}",
        ];
        $pos = $offset;

        $pos++;
        static::$cells['profile'] = [
            'key' => "K{$pos}",
            'value' => "P{$pos}",
            'title' => 'profile',
            'content' => 'Profile Content',
            'keyMerge' => "K{$pos}:O{$pos}",
            'valueMerge' => "P{$pos}:T{$pos}",
        ];

        $pos++;
        static::$cells['spacing'] = [
            'key' => "K{$pos}",
            'value' => "P{$pos}",
            'title' => 'spacing',
            'content' => '32mm',
            'keyMerge' => "K{$pos}:O{$pos}",
            'valueMerge' => "P{$pos}:T{$pos}",
        ];

        $pos++;
        static::$cells['species'] = [
            'key' => "K{$pos}",
            'value' => "P{$pos}",
            'title' => 'species',
            'content' => 'Spotted Gum',
            'keyMerge' => "K{$pos}:O{$pos}",
            'valueMerge' => "P{$pos}:T{$pos}",
        ];

        $pos++;
        static::$cells['coating'] = [
            'key' => "K{$pos}",
            'value' => "P{$pos}",
            'title' => 'coating',
            'content' => 'Clear Oil',
            'keyMerge' => "K{$pos}:O{$pos}",
            'valueMerge' => "P{$pos}:T{$pos}"
        ];

        $pos++;
        static::$cells['track_type'] = [
            'key' => "K{$pos}",
            'value' => "P{$pos}",
            'title' => 'mounting track type',
            'content' => 'Suspended Ceiling Track',
            'keyMerge' => "K{$pos}:O{$pos}",
            'valueMerge' => "P{$pos}:T{$pos}",
        ];

        $pos++;
        static::$cells['track_color'] = [
            'key' => "K{$pos}",
            'value' => "P{$pos}",
            'title' => 'mounting track color',
            'content' => 'Matt black',
            'keyMerge' => "K{$pos}:O{$pos}",
            'valueMerge' => "P{$pos}:T{$pos}"
        ];

        $pos++;
        static::$cells['backing'] = [
            'key' => "K{$pos}",
            'value' => "P{$pos}",
            'title' => 'acoustic backing',
            'content' => 'Yes',
            'keyMerge' => "K{$pos}:O{$pos}",
            'valueMerge' => "P{$pos}:T{$pos}"
        ];

        return $pos;
    }

    private static function createContent(Worksheet $sheet)
    {
        static::styleContent($sheet);
        static::fillContent($sheet);
    }

    private static function fillContent(Worksheet $sheet)
    {
        static::contentNames($sheet);
        foreach (static::$cells as $cell) {
            $sheet->getCell($cell['value'])->setValue($cell['content']);
        }
    }

    private static function contentNames(Worksheet $sheet)
    {
        foreach (static::$cells as $cell) {
            $sheet->getCell($cell['key'])->setValue(strtoupper($cell['title']));
        }
    }

    private static function styleContent(Worksheet $sheet)
    {
        foreach (static::$cells as $cell) {
            $sheet->mergeCells($cell['keyMerge']);
            $sheet->mergeCells($cell['valueMerge']);
        }

        $K = explode(":", reset(static::$cells)['keyMerge'])[0];
        $O = explode(":", end(static::$cells)['keyMerge'])[1];
        $T = explode(":", end(static::$cells)['valueMerge'])[1];

        $sheet
            ->getStyle("{$K}:{$T}")
            ->applyFromArray([
                'font' => [
                    'size' => 14,
                    'color' => ['argb' => '696969']
                ],
                'alignment' => ['vertical' => 'top', 'wrapText' => true],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['argb' => 'd4d4d4']
                    ]
                ],
            ]);

        $sheet
            ->getStyle("{$K}:{$O}")
            ->applyFromArray([
                'alignment' => ['vertical' => 'center'],
                'fill' => [
                    'fillType'   => Fill::FILL_SOLID,
                    'startColor' => ['argb' => 'f8f9f8']
                ]
            ]);

        //$sheet->getStyle('K14')->getBorders()->getTop()->setBorderStyle(Border::BORDER_NONE);
        //$sheet->getStyle('O14')->getBorders()->getTop()->setBorderStyle(Border::BORDER_NONE);
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
                'bold'  => false,
                'name'  => 'Lucida Sans',
                'size'  => 16,
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
