<?php

require_once __DIR__ . '/vendor/autoload.php';

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Common\Type;
use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Writer\Style\BorderBuilder;
use Box\Spout\Writer\Style\Border;

ini_set('max_execution_time', 30);
ini_set('memory_limit', '1G');

PHP_Timer::start();

$reader = ReaderFactory::create(Type::XLSX);
$reader->open(__DIR__ . '/big.xlsx');

foreach ($reader->getSheetIterator() as $sheet) {
    foreach ($sheet->getRowIterator() as $row) {}
}

$reader->close();

$time = PHP_Timer::stop();
var_dump($time);

print PHP_Timer::secondsToTimeString($time);
print PHP_Timer::resourceUsage();

die();

$bold = (new StyleBuilder())->setFontBold()->build();
$italic = (new StyleBuilder())->setFontItalic()->build();
$strike = (new StyleBuilder())->setFontStrikethrough()->build();
$redBackground = (new StyleBuilder())->setBackgroundColor(Color::RED)->build();
$greenBackground = (new StyleBuilder())->setBackgroundColor(Color::GREEN)->build();
$borderRight = (new StyleBuilder())->setBorder(
    (new BorderBuilder())->setBorderRight()->build()
)->build();
$borderTopOrange = (new StyleBuilder())->setBorder(
    (new BorderBuilder())->setBorderTop(Color::ORANGE)->build()
)->build();
$borderTopYellowThickDashed  = (new StyleBuilder())->setBorder(
    (new BorderBuilder())->setBorderTop(Color::YELLOW, Border::WIDTH_THICK, Border::STYLE_DASHED)->build()
)->build();

$styles = [
    'bold' => $bold,
    'italic' => $italic,
    'strike' => $strike,
    'green-background' => $greenBackground,
    'red-background' => $redBackground,
    'border-right' => $borderRight,
    'border-top-orange' => $borderTopOrange,
    'border-top-yellow-thick-dashed' => $borderTopYellowThickDashed,
];

$sheets = ['One', 'Two', 'Three'];
$maxRows = 1000;

/** @var \Box\Spout\Writer\XLSX\Writer $writer */
$writer = WriterFactory::create(Type::XLSX);
$writer->openToFile(__DIR__ . '/test.xlsx');
foreach($sheets as $sheetName) {
    $sheet = $writer->addNewSheetAndMakeItCurrent();
    $sheet->setName($sheetName);
    for($i = 1; $i < $maxRows; $i++) {
        /**
         * @var  $text
         * @var  $style Box\Spout\Writer\Style\Style;
         */
        foreach($styles as $text => $style) {
            try {
                $writer->addRowWithStyle([$text], $style);

                $key = array_rand($styles, 1);
                $merger = $styles[$key];
                $mergedStyle = $style->mergeWith($merger);
                $text = $text . '-' . $key;
                $writer->addRowWithStyle([$text], $mergedStyle);

            } catch(\Exception $e) {
                var_dump($text);
            }

        }
    }
}
$writer->close();

