<?php 
require_once(__DIR__.'/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);

$datas = [
    ['title'=>'MDword','date'=>date('Y-m-d'),'link'=>'https://github.com/mkdreams/MDword','content'=>'OFFICE WORD 动态数据 绑定数据 生成报告<br/>OFFICE WORD Dynamic data binding data generation report.'],
    ['title'=>'Wake up India, you\'re harming yourself','date'=>'2020-08-05','link'=>'http://epaper.chinadaily.com.cn/a/202008/06/WS5f2b56e4a3107831ec7540e6.html','content'=>'On Tuesday, reports said that the Indian government had announced a ban on Baidu and Weibo, two popular smartphone apps developed in China.<br/>Combined with the recent ban on short video sharing apps such as TikTok and Kwai, and social media app WeChat, India has now blocked its residents from using almost all popular Chinese apps.<br/>That apart, in the past few months, India has provoked border clashes with China, set limitations on Chinese enterprises and imposed higher tariffs on some products imported from China.'],
];
$bind = $TemplateProcessor->getBind($datas);
$bind->bindValue('item',[])
->bindValue('title',['title'],'item')
->bindValue('date',['date'],'item')
->bindValue('link',['link'],'item',function($value) {
    return [['type'=>MDWORD_LINK,'text'=>$value,'link'=>$value]];
})
->bindValue('content',['content'],'item',function($value) {
    $valueArr = explode('<br/>', $value);
    $texts = [];
    foreach($valueArr as $text) {
        $texts[] = ['type'=>MDWORD_TEXT,'text'=>$text];
        $texts[] = ['type'=>MDWORD_BREAK,'text'=>2];
    }
    
    return $texts;
})
;

$bind = $TemplateProcessor->getBind($datas);
$bind->bindValue('tableitem',[])
->bindValue('tabletitle',['title'],'tableitem',function($value,$row) {
    return [['type'=>MDWORD_LINK,'text'=>$value,'link'=>$row['link']]];
})
->bindValue('tabledate',['date'],'tableitem')
->bindValue('tablecontent',['content'],'tableitem',function($value) {
    $valueArr = explode('<br/>', $value);
    $texts = [];
    $textsCount = count($valueArr)-1;
    foreach($valueArr as $key => $text) {
        $texts[] = ['type'=>MDWORD_TEXT,'text'=>'　　'.$text];
        if($key < $textsCount) {
            $texts[] = ['type'=>MDWORD_BREAK,'text'=>1];
        }
    }
    
    return $texts;
})
;



$TemplateProcessor->cloneP('titletext',2);
$TemplateProcessor->setValue('titletext#0',[['type'=>MDWORD_TEXT,'text'=>'clone titletext#0','pstyle'=>'titlestyle-1','style'=>'titlestyle-1']]);
$TemplateProcessor->setValue('titletext#1',[['type'=>MDWORD_TEXT,'text'=>'clone titletext#1','pstyle'=>'titlestyle-1','style'=>'titlestyle-1']]);
$TemplateProcessor->setValue('titletext-sub',[['type'=>MDWORD_TEXT,'text'=>'titletext-sub','pstyle'=>'titlestyle-2']]);

$TemplateProcessor->cloneSection('section2',2);
$TemplateProcessor->setValue('section2#0','section2#012345');
$TemplateProcessor->setValue('section2#1','section2#1');
$TemplateProcessor->setValue('header2#1','header2#1');
$TemplateProcessor->deleteSection('section2#0');
$TemplateProcessor->saveAs($rtemplate);

