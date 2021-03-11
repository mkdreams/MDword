<?php 
require_once(__DIR__.'/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple-mdword.docx';
$rtemplatePHPWord = __DIR__.'/r-temple-phpword.docx';

$phpWord = new \PhpOffice\PhpWord\PhpWord();

$text = 'this is test content';
$n = 20000;

$b = microtime(true);
$section = $phpWord->addSection();
for($i = 0;$i<$n;$i++) {
    $section->addText(
        $text,
        array('name' => '黑体', 'size' => 15, 'color' => '000000', 'bold' => true, 'alignMent' => 'center'),
        array('alignment' => 'center', 'spaceAfter' => 200, 'spaceBefore' => 200)
    );
}
$phpWord->save($rtemplatePHPWord);
$e = microtime(true);
var_dump($e-$b);

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);

$b = microtime(true);
$TemplateProcessor->clones("test",$n);
$datas = [];
for($i = 0;$i<$n;$i++) {
    $TemplateProcessor->setValue('test#'.$i, $text);
}
$TemplateProcessor->saveAs($rtemplate);
$e = microtime(true);
var_dump($e-$b);
return ;



