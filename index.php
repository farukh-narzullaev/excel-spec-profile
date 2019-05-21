<?php

require_once 'vendor/autoload.php';

use App\SpecificationExcel;

$versions = [
    'win',
    'mac',
    'nix'
];

if (!isset($argv[1]) or !in_array($argv[1], $versions)) {
    echo "\e[1;37;42mNot enough parameters!\e[0m\n";
    return;
}

$output = __DIR__."/output/{$argv[1]}.xlsx";
$excel = new SpecificationExcel($output);
$excel->generate($argv[1]);

echo "\e[1;37;42mFile save in {$output}\e[0m\n";
