<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Color;
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
        $projectImage->setWidthAndHeight(559, 268);
        $projectImage->setCoordinates('A12');

        $projectImage->setWorksheet($sheet);

        $sheet->getStyle('A12:G25')
            ->getBorders()
            ->getAllBorders()
            ->setBorderStyle(Border::BORDER_THICK)
            ->setColor(new Color('757575'));

    }
}
