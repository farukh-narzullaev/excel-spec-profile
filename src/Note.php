<?php

namespace App;

use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Note
{
    public static function create(Worksheet $sheet)
    {
        static::title($sheet);
        static::includes($sheet);
        static::doesNotInclude($sheet);
        static::notes($sheet);
        static::based($sheet);
    }

    private static function based(Worksheet $sheet)
    {
        $text = "*Based on professional opinion only. See disclaimer in Price & Spec for full details.";

        $sheet->mergeCells('K42:T42');
        $sheet->getCell('K42')->setValue($text);

        $styles = static::getStyles();
        $styles['font']['italic'] = true;

        $sheet->getStyle('K42')->applyFromArray($styles);
    }

    private static function notes(Worksheet $sheet)
    {
        $text = "Please note: We have a minimum order of $15k due to manufacturing costs. The information provided in the Price & Spec Tool is a guideline for your convenience. Every care has been taken to ensure reasonable accuracy however variations may occur. See our Terms of Use for more details.";

        $sheet->mergeCells('K39:T41');
        $sheet->getCell('K39')->setValue($text);

        $styles = static::getStyles();
        $styles['font']['italic'] = true;

        $sheet->getStyle('K39')->applyFromArray($styles);
    }

    private static function doesNotInclude(Worksheet $sheet)
    {
        $sheet->mergeCells('K36:T38');

        $richText = static::createRichText(
            "Price does not include: ",
            "any custom corner/edging treatment, flashing, installation, substrates, hangers and TCR for suspended ceilings, fixings for substrate or insulation batts. Pricing is based on the amount of product required (M2) as specified on the pricing panel.");

        $sheet->getCell('K36')->setValue($richText);

        $sheet->getStyle('K36')->applyFromArray(static::getStyles());
    }

    private static function includes(Worksheet $sheet)
    {
        $sheet->mergeCells('K33:T35');

        $richText = static::createRichText(
            "Price includes: ",
            "freight to site within Australia, contact Sculptform direct for advice on shipping outside of Australia. Corner trims, Endcaps, L profiles and all other standard componentry see our website for more information.");

        $sheet->getCell('K33')->setValue($richText);

        $sheet->getStyle('K33')->applyFromArray(static::getStyles());
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
        //$price->getFont()->setSize(20);
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
        $sheet->mergeCells('K31:T32');
        $sheet->getCell('K31')->setValue('IMPORTANT PRICING NOTE');
        $sheet
            ->getStyle('K31')
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
