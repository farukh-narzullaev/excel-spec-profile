<?php

namespace App;

use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Note
{
    private static $pos;

    public static function create(Worksheet $sheet, $pos)
    {
        static::$pos = ($pos + 2);
        static::title($sheet);

        static::$pos += 2;
        static::includes($sheet);

        static::$pos += 3;
        static::doesNotInclude($sheet);

        static::$pos += 3;
        static::notes($sheet);

        static::$pos += 3;
        static::based($sheet);

        return static::$pos;
    }

    private static function based(Worksheet $sheet)
    {
        $pos = static::$pos;
        $text = "*Based on professional opinion only. See disclaimer in Price & Spec for full details.";

        $sheet->mergeCells("K{$pos}:T{$pos}");
        $sheet->getCell("K{$pos}")->setValue($text);

        $styles = static::getStyles();
        $styles['font']['italic'] = true;

        $sheet->getStyle("K{$pos}")->applyFromArray($styles);
    }

    private static function notes(Worksheet $sheet)
    {
        $pos = static::$pos;
        $merge = $pos + 2;

        $text = "Please note: We have a minimum order of $15k due to manufacturing costs. The information provided in the Price & Spec Tool is a guideline for your convenience. Every care has been taken to ensure reasonable accuracy however variations may occur. See our Terms of Use for more details.";

        $sheet->mergeCells("K{$pos}:T{$merge}");
        $sheet->getCell("K{$pos}")->setValue($text);

        $styles = static::getStyles();
        $styles['font']['italic'] = true;

        $sheet->getStyle("K{$pos}")->applyFromArray($styles);
    }

    private static function doesNotInclude(Worksheet $sheet)
    {
        $pos = static::$pos;
        $merge = $pos + 2;

        $sheet->mergeCells("K{$pos}:T{$merge}");

        $richText = static::createRichText(
            "Price does not include: ",
            "any custom corner/edging treatment, flashing, installation, substrates, hangers and TCR for suspended ceilings, fixings for substrate or insulation batts. Pricing is based on the amount of product required (M2) as specified on the pricing panel.");

        $sheet->getCell("K{$pos}")->setValue($richText);

        $sheet->getStyle("K{$pos}")->applyFromArray(static::getStyles());
    }

    private static function includes(Worksheet $sheet)
    {
        $pos = static::$pos;
        $merge = $pos + 2;

        $sheet->mergeCells("K{$pos}:T{$merge}");

        $richText = static::createRichText(
            "Price includes: ",
            "freight to site within Australia, contact Sculptform direct for advice on shipping outside of Australia. Corner trims, Endcaps, L profiles and all other standard componentry see our website for more information.");

        $sheet->getCell("K{$pos}")->setValue($richText);

        $sheet->getStyle("K{$pos}")->applyFromArray(static::getStyles());
    }

    private static function getStyles()
    {
        return [
            'font' => [
                'color' => ['argb' => '696969'],
            ],
            'alignment' => [
                'horizontal' => 'left',
                'vertical'   => 'top',
                'wrapText'   => true
            ],
        ];
    }

    private static function createRichText($heading, $text)
    {
        $richText = new RichText();
        $richText->createText('');

        $price = $richText->createTextRun($heading);
        $price->getFont()->setBold(true);
        $price->getFont()->setSize(12);
        $price
            ->getFont()
            ->getColor()->setARGB('696969');

        $based = $richText->createTextRun($text);
        $based
            ->getFont()
            ->getColor()->setARGB('696969');

        return $richText;
    }

    private static function title(Worksheet $sheet)
    {
        $pos = static::$pos;
        $merge = $pos + 1;
        $sheet->mergeCells("K{$pos}:T{$merge}");

        $sheet->getCell("K{$pos}")->setValue('IMPORTANT PRICING NOTE');
        $sheet
            ->getStyle("K{$pos}")
            ->applyFromArray([
                'font' => [
                    'color' => ['argb' => '696969'],
                    'size'  => 20,
                    'bold'  => true,
                ],
                'alignment' => [
                    'horizontal' => 'left',
                    'vertical'   => 'center',
                    'wrapText'   => true
                ],
            ]);
    }
}
