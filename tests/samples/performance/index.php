<?php 
require_once(__DIR__.'/../../../Autoloader.php');
require_once(__DIR__.'/../../function.php');

use MDword\WordProcessor;
ini_set('memory_limit','1600M');
set_time_limit(1200);

$ns = [
100,
500,
1000,
10000,
// 50000
];

$markdownTableDatas = [["测试项","用时(S)"]];


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple.docx';
foreach($ns as $n) {
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
    $markdownTableDatas[] = ["1页母版赋值".$n."次",number_format($e-$b,2)];
}

$template = __DIR__.'/1750pages.docx';
$rtemplate = __DIR__.'/r-1750pages.docx';

foreach($ns as $n) {
    $b = microtime(true);
    $TemplateProcessor = new WordProcessor();
    $TemplateProcessor->load($template);
    $TemplateProcessor->cloneP('test', $n);
    for($i=0;$i<$n;$i++) {
        $TemplateProcessor->setValue('test#'.$i, 'MDword'.$i);
    }
    $TemplateProcessor->saveAs($rtemplate);
    $e = microtime(true);
    $markdownTableDatas[] = ["1750页母版赋值".$n."次",number_format($e-$b,2)];
}
    

echo markdownTable($markdownTableDatas);
return ;



