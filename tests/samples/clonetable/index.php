<?php 
require_once(__DIR__.'/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);

$datas = [
    ['title'=>'有效成分1信息','A1'=>'中文通用名','B1'=>'中文通用名1','C1'=>'生物活性','D1'=>'生物活性1','F1'=>'结构式','F2'=> dirname(__FILE__).'/img.png'],
    ['title'=>'有效成分2信息','A1'=>'中文通用名','B1'=>'中文通用名2','C1'=>'生物活性','D1'=>'生物活性2','F1'=>'结构式','F2'=> dirname(__FILE__).'/img.png'],
];


$TemplateProcessor->clones('table',count($datas));
foreach($datas as $key => $data) {
    foreach($data as $vk => $v) {
        if($vk === 'F2') {
            $TemplateProcessor->setValue($vk.'#'.$key,[['type'=>MDWORD_IMG,'text'=>$v]]);
        }else{
            $TemplateProcessor->setValue($vk.'#'.$key,[['type'=>MDWORD_TEXT,'text'=>$v]]);
        }
    }
}

$TemplateProcessor->saveAs($rtemplate);

