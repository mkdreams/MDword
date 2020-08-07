<?php 
require_once(__DIR__.'/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);

$datas = [
    ['price'=>100,'change'=>5,'changepercent'=>0.05],
    ['price'=>200,'change'=>-10,'changepercent'=>-0.05],
    ['price'=>500,'change'=>100,'changepercent'=>0.20],
];
$bind = $TemplateProcessor->getBind($datas);
$bind->bindValue('item',[])
->bindValue('stockprice',['price'],'item')
->bindValue('change',['change'],'item',function($value) {
    if($value > 0) {
        $style = 'red';
    }else{
        $style = 'green';
    }
    return [['type'=>MDWORD_TEXT,'text'=>$value,'style'=>$style]];
})
->bindValue('changepercent',['changepercent'],'item',function($value) {
    if($value > 0) {
        $style = 'red';
    }else{
        $style = 'green';
    }
    return [['type'=>MDWORD_TEXT,'text'=>($value*100).'%','style'=>$style]];
})
;

$TemplateProcessor->saveAs($rtemplate);

