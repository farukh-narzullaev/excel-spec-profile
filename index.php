<?php

require_once 'vendor/autoload.php';

use App\SpecificationExcel;

$output = __DIR__."/output/file.xlsx";

$excel = new SpecificationExcel($output);
$excel->generate();
