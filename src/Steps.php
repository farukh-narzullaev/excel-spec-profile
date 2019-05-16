<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Steps
{
    public static function create(Worksheet $sheet)
    {
        static::title($sheet);
        static::steps($sheet);
    }

    private static function steps(Worksheet $sheet)
    {
        static::step1($sheet);
        static::step2($sheet);
        static::step3($sheet);
    }

    private static function step1(Worksheet $sheet)
    {
        $cell = "A49:F55";
        $bold = "Calculate your cost.";
        $text = "If the product specification and pricing is satisfactory, simply multiply this m 2 price by the m 2 of material required for your project.";

        static::styleStep($sheet, $cell);

        $richText = new RichText();
        $richText->createTextRun("Step 1 \n")->getFont()->setBold(true)->setSize(20)->getColor()->setARGB('8b0000');
        $richText->createTextRun("{$bold} ")->getFont()->setBold(true)->setSize(15)->getColor()->setARGB('696969');
        $richText->createTextRun($text)->getFont()->setSize(14)->getColor()->setARGB('696969');

        $sheet->getCell(explode(":", $cell)[0])->setValue($richText);
    }

    private static function step2(Worksheet $sheet)
    {
        $cell = "H49:M55";
        $bold = " BOQ estimation ";
        $text = "Once the tender has been approved, email us the final construction drawings to receive a formal";
        $text2 = "which includes all standard trims and componentry.";

        static::styleStep($sheet, $cell);

        $richText = new RichText();
        $richText->createTextRun("Step 2 \n")->getFont()->setBold(true)->setSize(20)->getColor()->setARGB('8b0000');
        $richText->createTextRun($text)->getFont()->setSize(14)->getColor()->setARGB('696969');
        $richText->createTextRun("{$bold}")->getFont()->setBold(true)->setSize(15)->getColor()->setARGB('696969');
        $richText->createTextRun($text2)->getFont()->setSize(14)->getColor()->setARGB('696969');


        $sheet->getCell(explode(":", $cell)[0])->setValue($richText);
    }

    private static function step3(Worksheet $sheet)
    {
        $cell = "O49:T55";
        $bold = "place your order.";
        $text = "Review itemised quote and follow the prompts to ";

        static::styleStep($sheet, $cell);

        $richText = new RichText();
        $richText->createTextRun("Step 3 \n")->getFont()->setBold(true)->setSize(20)->getColor()->setARGB('8b0000');
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
        $sheet->mergeCells('A46:T47');

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

        $sheet->getCell('A46')->setValue($richText);

        $sheet
            ->getStyle('K46')
            ->applyFromArray([
                'alignment' => [
                    'horizontal' => 'left',
                    'vertical'   => 'center',
                    'wrapText'   => true
                ],
            ]);
    }
}
