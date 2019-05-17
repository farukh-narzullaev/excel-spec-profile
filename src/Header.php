<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Cell\Hyperlink;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Header
{
    public static function create(Worksheet $sheet)
    {
        $headerLink = new Hyperlink('http://sculptform.com.au', 'Go to the SculptForm');

        $header = new Drawing();
        $header->setName('Logo');
        $header->setDescription('Logo');
        $header->setPath('images/header.png');
        $header->setResizeProportional(false);

        // 1280 Windows
        // 1413 Mac
        $header->setWidth(1413);
        $header->setHeight(200);
        //$header->setWidthAndHeight(1450, 200);
        $header->setHyperlink($headerLink);
        $header->setWorksheet($sheet);
    }
}
