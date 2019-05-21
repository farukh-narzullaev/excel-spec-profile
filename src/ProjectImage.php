<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ProjectImage
{
    /**
     * @var array
     */
    private static $sizes = [
        'win' => [576, 258],
        'mac' => [636, 258],
        'nix' => [719, 245],
    ];

    /**
     * @param Worksheet $sheet
     * @param string    $version
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function create(Worksheet $sheet, $version)
    {
        $projectImage = new Drawing();
        $projectImage->setPath('images/1.png');
        $projectImage->setResizeProportional(false);

        list($width, $height) = static::$sizes[$version];

        $projectImage->setWidthAndHeight($width, $height);
        $projectImage->setCoordinates('A13');
        $projectImage->setOffsetY(1);
        $projectImage->setWorksheet($sheet);

        $sheet->getStyle('A13:I25')
            ->applyFromArray([
                'borders' => [
                    'outline' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['argb' => '757575']
                    ]
                ]
            ]);
    }
}
