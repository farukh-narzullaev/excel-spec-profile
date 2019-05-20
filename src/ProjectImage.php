<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ProjectImage
{
    public static function create(Worksheet $sheet)
    {
        $projectImage = new Drawing();
        $projectImage->setPath('images/1.png');
        $projectImage->setResizeProportional(false);

        // 576x258 Windows
        // 636x258 Mac
        // 719x245 Linux
        $projectImage->setWidthAndHeight(719, 245);
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
