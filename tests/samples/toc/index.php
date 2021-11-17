<?php 
require_once(__DIR__.'/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);

//TOC and bind data
$redWords = 'WORD';
$datas = [
    ['title'=>'MDword Github','date'=>date('Y-m-d'),'link'=>'https://github.com/mkdreams/MDword','content'=>'OFFICE WORD 动态数据 绑定数据 生成报告<br/>OFFICE WORD Dynamic data binding data generation report.'],
    ['title'=>'Wake up India, you\'re harming yourself','date'=>'2020-08-05','link'=>'http://epaper.chinadaily.com.cn/a/202008/06/WS5f2b56e4a3107831ec7540e6.html','content'=>'On Tuesday, reports said that the Indian government had announced a ban on Baidu and Weibo, two popular smartphone apps developed in China.<br/>Combined with the recent ban on short video sharing apps such as TikTok and Kwai, and social media app WeChat, India has now blocked its residents from using almost all popular Chinese apps.<br/>That apart, in the past few months, India has provoked border clashes with China, set limitations on Chinese enterprises and imposed higher tariffs on some products imported from China.'],
    ['title'=>'MDword Github','date'=>date('Y-m-d'),'link'=>'https://github.com/mkdreams/MDword','content'=>'OFFICE WORD 动态数据 绑定数据 生成报告<br/>OFFICE WORD Dynamic data binding data generation report.'],
];
$bind = $TemplateProcessor->getBind($datas);
$bind->bindValue('news',[])
->bindValue('title',['title'],'news')
->bindValue('date',['date'],'news')
->bindValue('link',['link'],'news',function($value) {
    return [['type'=>MDWORD_LINK,'text'=>$value,'link'=>$value]];
})
->bindValue('content',['content'],'news',function($value) use($redWords) {
    $valueArr = explode('<br/>', $value);
    $texts = [];
    foreach($valueArr as $text) {
        $texts[] = ['type'=>MDWORD_TEXT,'text'=>$text];
        $tempTexts = explode($redWords,$text);
        foreach($tempTexts as $idx => $tempText) {
            
            if($idx != 0) {
                $texts[] = ['type'=>MDWORD_TEXT,'style'=>'red','text'=>$redWords];
            }
            $texts[] = ['type'=>MDWORD_TEXT,'text'=>$tempText];
        }

        $texts[] = ['type'=>MDWORD_BREAK,'text'=>2];
    }

    return $texts;
})
;


//toc table
$InnerVars = $TemplateProcessor->getInnerVars();

$TemplateProcessor->clones('table_item', 3);

$TemplateProcessor->setValue('table_title#0',['text'=>$InnerVars['levels'][1]['text'],'name'=>$InnerVars['levels'][1]['name']],MDWORD_REF);
$TemplateProcessor->setValue('table_title#1',['text'=>$InnerVars['levels'][2]['text'],'name'=>$InnerVars['levels'][2]['name']],MDWORD_REF);
$TemplateProcessor->setValue('table_title#2',['text'=>$InnerVars['levels'][3]['text'],'name'=>$InnerVars['levels'][3]['name']],MDWORD_REF);

$TemplateProcessor->setValue('table_page#0',['name'=>$InnerVars['levels'][1]['name']],MDWORD_PAGEREF);
$TemplateProcessor->setValue('table_page#1',['name'=>$InnerVars['levels'][2]['name']],MDWORD_PAGEREF);
$TemplateProcessor->setValue('table_page#2',['name'=>$InnerVars['levels'][3]['name']],MDWORD_PAGEREF);

//page
$TemplateProcessor->setValue('now_page',['name'=>'bookmarket_toc'],MDWORD_NOWPAGE);
$TemplateProcessor->setValue('total_page',['name'=>'bookmarket_toc'],MDWORD_TOTALPAGE);

$TemplateProcessor->setValue('toc_page',['name'=>'bookmarket_toc'],MDWORD_PAGEREF);


$TemplateProcessor->saveAs($rtemplate);

