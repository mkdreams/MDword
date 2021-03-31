<?php 
require_once(__DIR__.'/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);

//simple set value
$TemplateProcessor->setValue('name', 'colin');
// $TemplateProcessor->setValue('header', 'HEADER');
// $TemplateProcessor->setValue('footer', 'FOOTER111');

$TemplateProcessor->saveAs($rtemplate,true);

