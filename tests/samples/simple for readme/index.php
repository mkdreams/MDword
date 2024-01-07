<?php
require_once(__DIR__ . '/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__ . '/temple.docx';
$rtemplate = __DIR__ . '/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);

//simple set value
$TemplateProcessor->setValue('value', 'r-value');
$TemplateProcessor->setValue('value', 'r-value2');

//image
$TemplateProcessor->setImageValue('image', dirname(__FILE__) . '/logo.jpg');

//table
$TemplateProcessor->clones('people', 3);

$TemplateProcessor->setValue('name#0', 'colin0');
$TemplateProcessor->setValue('name#1', [['text' => 'colin1', 'style' => 'style', 'type' => MDWORD_TEXT]]);
$TemplateProcessor->setValue('name#2', 'colin2');

$TemplateProcessor->setValue('sex#1', 'woman');

$TemplateProcessor->setValue('age#0', '280');
$TemplateProcessor->setValue('age#1', '281');
$TemplateProcessor->setValue('age#2', '282');

//list
$TemplateProcessor->cloneP('item', 3);
$TemplateProcessor->setValue('item#0', 'ITEM1');
$TemplateProcessor->setValue('item#1', 'ITEM2');
$TemplateProcessor->setValue('item#2', 'ITEM3');

//TOC and bind data
$redWords = 'WORD';
$datas = [
    [
        'title' => 'MDword Github', 'date' => date('Y-m-d'),
        'link' => 'https://github.com/mkdreams/MDword',
        'images' => dirname(__FILE__) . '/logo.jpg,' . dirname(__FILE__) . '/img.png',
        'content' => 'OFFICE WORD 动态数据 绑定数据 生成报告<br/>OFFICE WORD Dynamic data binding data generation report.'
    ],
    [
        'title' => 'Wake up India, you\'re harming yourself',
        'date' => '2020-08-05',
        'link' => 'http://epaper.chinadaily.com.cn/a/202008/06/WS5f2b56e4a3107831ec7540e6.html',
        'images' => dirname(__FILE__) . '/logo.jpg',
        'content' => 'On Tuesday, reports said that the Indian government had announced a ban on Baidu and Weibo, two popular smartphone apps developed in China.<br/>Combined with the recent ban on short video sharing apps such as TikTok and Kwai, and social media app WeChat, India has now blocked its residents from using almost all popular Chinese apps.<br/>That apart, in the past few months, India has provoked border clashes with China, set limitations on Chinese enterprises and imposed higher tariffs on some products imported from China.'
    ],
];
$bind = $TemplateProcessor->getBind($datas);
$bind->bindValue('news', [])
    ->bindValue('title', ['title'], 'news')
    ->bindValue('date', ['date'], 'news')
    ->bindValue('link', ['link'], 'news', function ($value) {
        return [['type' => MDWORD_LINK, 'text' => $value, 'link' => $value]];
    })
    ->bindValue('images', ['images'], 'news', function ($value) {
        $images = explode(',', $value);
        $texts = [];
        foreach ($images as $key => $image) {
            if ($key === 0) {
                $texts[] = ['type' => MDWORD_IMG, 'text' => $image, 'style' => 'imgstyle'];
            } else {
                $texts[] = ['type' => MDWORD_IMG, 'text' => $image];
            }
        }
        return $texts;
    })
    ->bindValue('content', ['content'], 'news', function ($value) use ($redWords) {
        $valueArr = explode('<br/>', $value);
        $texts = [];
        foreach ($valueArr as $text) {
            $texts[] = ['type' => MDWORD_TEXT, 'text' => $text];
            $tempTexts = explode($redWords, $text);
            foreach ($tempTexts as $idx => $tempText) {

                if ($idx != 0) {
                    $texts[] = ['type' => MDWORD_TEXT, 'style' => 'red', 'text' => $redWords];
                }
                $texts[] = ['type' => MDWORD_TEXT, 'text' => $tempText];
            }

            $texts[] = ['type' => MDWORD_BREAK, 'text' => 2];
        }

        return $texts;
    });

$numDatas = [
    [
        'title'=>'title-1',
        'content'=>'content-1'
    ],
    [
        'title'=>'title-2',
        'sub'=>[
            [
                'title'=>'subTitle-2-1',
                'content'=>'content-2-1',
            ],
            [
                'title'=>'subTitle-2-2',
                'content'=>'content-2-2',
            ],
        ]
    ],
    [
        'title'=>'title-3',
        'sub'=>[
            [
                'title'=>'subTitle-3-1',
                'content'=>'content-3-1',
            ],
            [
                'title'=>'subTitle-3-2',
                'content'=>'content-3-2',
            ],
        ]
    ],
];

$TemplateProcessor->cloneP('num',count($numDatas));
foreach($numDatas as $idx => $numData) {
    $TemplateProcessor->cloneP('num'.'#'.$idx,3);

    $TemplateProcessor->setValue('num'.'#'.$idx.'#0',[['text' => $numData['title'], 'pstyle' => 'numstyle-level-1', 'style' => 'numstyle-level-1', 'type' => MDWORD_TEXT]]);

    if(isset($numData['content'])) {
        $TemplateProcessor->setValue('num'.'#'.$idx.'#1',[['text' => $numData['content'], 'pstyle' => 'numstyle-level-3', 'style' => 'numstyle-level-3', 'type' => MDWORD_TEXT]]);
    }else{
        $TemplateProcessor->deleteP('num'.'#'.$idx.'#1');
    }

    $subName = 'num'.'#'.$idx.'#2';
    if(isset($numData['sub'])) {
        $TemplateProcessor->cloneP($subName,count($numData['sub']));

        foreach($numData['sub'] as $subIdx => $subData) {
            $TemplateProcessor->cloneP($subName.'#'.$subIdx,2);

            $TemplateProcessor->setValue($subName.'#'.$subIdx.'#0',[['text' => $subData['title'], 'pstyle' => 'numstyle-level-2', 'style' => 'numstyle-level-2', 'type' => MDWORD_TEXT]]);
            $TemplateProcessor->setValue($subName.'#'.$subIdx.'#1',[['text' => $subData['content'], 'pstyle' => 'numstyle-level-3', 'style' => 'numstyle-level-3', 'type' => MDWORD_TEXT]]);
        }
    }else{
        $TemplateProcessor->deleteP($subName);
    }
}

//links
$TemplateProcessor->setValue('plink', [
    ['type'=>MDWORD_LINK,'text' => 'colin1','link'=>'https://baidu.com?v=1'],
    ['type' =>MDWORD_LINK,'text' => 'colin2', 'style' => 'style','link'=>'https://baidu.com?v=2'],
    ['type'=>MDWORD_LINK,'text' => 'colin3','link'=>'https://baidu.com?v=3'],
]);

$TemplateProcessor->deleteP('numstyle');


$TemplateProcessor->deleteP('style');
$TemplateProcessor->deleteP('red');
$TemplateProcessor->deleteP('imgstyle');

$TemplateProcessor->saveAs($rtemplate);
