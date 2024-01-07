<?php
require_once(__DIR__ . '/../../../Autoloader.php');

use MDword\WordProcessor;


//TOC and bind data
$redWords = 'WORD';
$datas = [
    [
        'title' => 'MDword Github', 'date' => date('Y-m-d'),
        'content' => 'OFFICE WORD 动态数据 绑定数据 生成报告<br/>OFFICE WORD Dynamic data binding data generation report.'
    ],
    [
        'title' => 'Wake up India, you\'re harming yourself',
        'content' => 'On Tuesday, reports said that the Indian government had announced a ban on Baidu and Weibo, two popular smartphone apps developed in China.<br/>Combined with the recent ban on short video sharing apps such as TikTok and Kwai, and social media app WeChat, India has now blocked its residents from using almost all popular Chinese apps.<br/>That apart, in the past few months, India has provoked border clashes with China, set limitations on Chinese enterprises and imposed higher tariffs on some products imported from China.'
    ],
];

$template = __DIR__ . '/wps-temple.docx';
$rtemplate = __DIR__ . '/r-wps-temple.docx';
$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);
$bind = $TemplateProcessor->getBind($datas);
$bind->bindValue('news', [])
    ->bindValue('title', ['title'], 'news')
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

        $texts[] = ['type' => MDWORD_PAGE_BREAK];
        return $texts;
    });
$TemplateProcessor->saveAs($rtemplate);

$template = __DIR__ . '/word-temple.docx';
$rtemplate = __DIR__ . '/r-word-temple.docx';
$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);
$bind = $TemplateProcessor->getBind($datas);
$bind->bindValue('news', [])
    ->bindValue('title', ['title'], 'news')
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

        $texts[] = ['type' => MDWORD_PAGE_BREAK];
        return $texts;
    });
$TemplateProcessor->saveAs($rtemplate);
