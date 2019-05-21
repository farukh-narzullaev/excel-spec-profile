<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Cell\Hyperlink;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Header
{
    /**
     * @var array
     */
    private static $sizes = [
        'win' => 1280,
        'mac' => 1413,
        'nix' => 1598,
    ];

    /**
     * @param Worksheet $sheet
     * @param string    $version
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function create(Worksheet $sheet, $version)
    {
        $headerLink = new Hyperlink('http://sculptform.com.au', 'Go to the SculptForm');

        $header = new Drawing();
        $header->setName('Logo');
        $header->setDescription('Logo');
        $header->setPath('images/header.png');
        $header->setResizeProportional(false);

        $header->setWidth(static::$sizes[$version]);
        $header->setHeight(210);
        $header->setHyperlink($headerLink);
        $header->setWorksheet($sheet);
    }
}
