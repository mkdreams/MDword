<?php
require_once(__DIR__ . '/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__ . '/temple.docx';
$rtemplate = __DIR__ . '/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);

//table
$TemplateProcessor->clones('people', 5);

$TemplateProcessor->setValue('name#0', 'colin0');
$TemplateProcessor->setValue('name#1', [['text' => 'colin1', 'table_style' => ['vMerge'=>3,'gridSpan'=>2], 'type' => MDWORD_TEXT]]);
$TemplateProcessor->setValue('name#2', 'colin2');

$TemplateProcessor->setValue('sex#1', 'woman');

$TemplateProcessor->setValue('age#0', [['text' => '280', 'table_style' => ['gridSpan'=>2], 'type' => MDWORD_TEXT]]);
$TemplateProcessor->setValue('age#2', [['text' => '282', 'table_style' => ['gridSpan'=>2], 'type' => MDWORD_TEXT]]);


$TemplateProcessor->saveAs($rtemplate);
