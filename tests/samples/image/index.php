<?php 
require_once(__DIR__.'/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);


// $TemplateProcessor->showMedies();//help method for get md5

$TemplateProcessor->setImageValue('image insert', dirname(__FILE__).'/logo.jpg');

$TemplateProcessor->setImageValue('image replace', dirname(__FILE__).'/logo.jpg');

$TemplateProcessor->setImageValue('2529be7711acbb60c7e4ac1693c680a0', dirname(__FILE__).'/logo.jpg');


$TemplateProcessor->saveAs($rtemplate);

