<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Steps
{
    private static $pos;

    public static function create(Worksheet $sheet, $pos)
    {
        static::$pos = ($pos + 3);
        static::title($sheet);

        static::$pos += 3;
        static::steps($sheet);

        return static::$pos;
    }

    private static function steps(Worksheet $sheet)
    {
        static::step1($sheet);
        static::step2($sheet);
        static::step3($sheet);
    }

    private static function step1(Worksheet $sheet)
    {
        $pos = static::$pos;
        $merge = $pos + 6;

        $cell = "A{$pos}:F{$merge}";
        $bold = "Calculate your cost.";
        $text = "If the product specification and pricing is satisfactory, simply multiply this m 2 price by the m 2 of material required for your project.";

        static::styleStep($sheet, $cell);

        $richText = new RichText();
        $richText->createTextRun("Step 1")->getFont()->setBold(true)->setSize(20)->getColor()->setARGB('8b0000');
        $richText->createText("\r\n");
        $richText->createTextRun("{$bold} ")->getFont()->setBold(true)->setSize(15)->getColor()->setARGB('696969');
        $richText->createTextRun($text)->getFont()->setSize(14)->getColor()->setARGB('696969');

        $sheet->getCell(explode(":", $cell)[0])->setValue($richText);
    }

    private static function step2(Worksheet $sheet)
    {
        $pos = static::$pos;
        $merge = $pos + 6;

        $cell = "H{$pos}:M{$merge}";
        $bold = " BOQ estimation ";
        $text = "Once the tender has been approved, email us the final construction drawings to receive a formal";
        $text2 = "which includes all standard trims and componentry.";

        static::styleStep($sheet, $cell);

        $richText = new RichText();
        $richText->createTextRun("Step 2 ")->getFont()->setBold(true)->setSize(20)->getColor()->setARGB('8b0000');
        $richText->createText("\r\n");
        $richText->createTextRun($text)->getFont()->setSize(14)->getColor()->setARGB('696969');
        $richText->createTextRun("{$bold}")->getFont()->setBold(true)->setSize(15)->getColor()->setARGB('696969');
        $richText->createTextRun($text2)->getFont()->setSize(14)->getColor()->setARGB('696969');


        $sheet->getCell(explode(":", $cell)[0])->setValue($richText);
    }

    private static function step3(Worksheet $sheet)
    {
        $pos = static::$pos;
        $merge = $pos + 6;

        $cell = "O{$pos}:T{$merge}";
        $bold = "place your order.";
        $text = "Review itemised quote and follow the prompts to ";

        static::styleStep($sheet, $cell);

        $richText = new RichText();
        $richText->createTextRun("Step 3 ")->getFont()->setBold(true)->setSize(20)->getColor()->setARGB('8b0000');
        $richText->createText("\r\n");
        $richText->createTextRun($text)->getFont()->setSize(14)->getColor()->setARGB('696969');
        $richText->createTextRun("{$bold} ")->getFont()->setBold(true)->setSize(15)->getColor()->setARGB('696969');

        $sheet->getCell(explode(":", $cell)[0])->setValue($richText);
    }

    private static function styleStep(Worksheet $sheet, $cell)
    {
        $sheet->mergeCells($cell);
        $sheet->getStyle($cell)->applyFromArray([
            'font' => [
                'color' => ['argb' => '696969'],
                'size' => 18,
            ],
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => '757575']
                ]
            ],
            'alignment' => [
                'horizontal' => 'left',
                'vertical'   => 'top',
                'wrapText'   => true
            ],
        ]);
    }

    private static function title(Worksheet $sheet)
    {
        $pos = static::$pos;
        $merge = $pos + 1;

        $sheet->mergeCells("A{$pos}:T{$merge}");

        $richText = new RichText();
        $richText->createText('');

        $part1 = $richText->createTextRun("Happy with your specification? ");
        $part1
            ->getFont()
            ->setBold(true)
            ->setSize(25)
            ->getColor()->setARGB('8b0000');

        $part2 = $richText->createTextRun("Follow these 3 simple steps below to order.");
        $part2
            ->getFont()
            ->setSize(25)
            ->getColor()->setARGB('696969');

        $sheet->getCell("A{$pos}")->setValue($richText);

        $sheet
            ->getStyle("K{$pos}")
            ->applyFromArray([
                'alignment' => [
                    'horizontal' => 'left',
                    'vertical'   => 'center',
                    'wrapText'   => true
                ],
            ]);
    }
}
