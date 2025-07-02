<?php 
require_once(__DIR__.'/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);


$TemplateProcessor->clones('row',3);
$TemplateProcessor->setValue('rowIndex#0',0);
$TemplateProcessor->setValue('rowIndex#1',1);
$TemplateProcessor->setValue('rowIndex#2',2);
$TemplateProcessor->setValue('rowIndex#2',3);
$TemplateProcessor->setValue('rowIndex#2',4);

$TemplateProcessor->setImageValue('rowImage#0',dirname(__FILE__).'/img.jpg');
$TemplateProcessor->setImageValue('rowImage#1',dirname(__FILE__).'/img2.bmp');

$rows = [
    ['index'=>2,'image'=> dirname(__FILE__).'/img.jpg'],
    ['index'=>3,'image'=> dirname(__FILE__).'/img3.png'],
    ['index'=>4,'image'=> dirname(__FILE__).'/img3.png'],
];

$bind = $TemplateProcessor->getBind($rows);
$bind->bindValue('row#2',[])
->bindValue('rowIndex#2',['index'],'row#2')
->bindValue('rowImage#2',['image'],'row#2',MDWORD_IMG)
;

$TemplateProcessor->setImageValue('rowImage#2#0',dirname(__FILE__).'/img.jpg');

$TemplateProcessor->setValue('image insert', [['text' => dirname(__FILE__).'/words.png','type' => MDWORD_IMG,'width'=>300]]);

$TemplateProcessor->setImageValue('image replace', dirname(__FILE__).'/img.jpg');

//replace image by md5
$Medies = $TemplateProcessor->getMedies();
foreach($Medies as $Medie) {
    if('a2a9f8b099d7c5764da685b628468052' === $Medie['md5']) {
        $TemplateProcessor->setImageValue($Medie['md5'], dirname(__FILE__).'/img.jpg');
    }
}

$TemplateProcessor->setValue("full", [['text' => dirname(__FILE__).'/full.png','temple'=>file_get_contents(dirname(__FILE__).'/temple/full.xml'),'type' => MDWORD_IMG,'width'=>'100%']]);
$TemplateProcessor->setValue("full2", [
    ['type' => MDWORD_TEXT,'text'=> "This is text."],
    ['type' => MDWORD_PAGE_BREAK,'replace'=>false],
    ['text' => dirname(__FILE__).'/full.png','temple'=>file_get_contents(dirname(__FILE__).'/temple/full.xml'),'type' => MDWORD_IMG,'width'=>'100%'],
    ['type' => MDWORD_PAGE_BREAK,'replace'=>false],
    ['text' => dirname(__FILE__).'/full.png','temple'=>file_get_contents(dirname(__FILE__).'/temple/full.xml'),'type' => MDWORD_IMG,'width'=>'100%'],
    ['type' => MDWORD_PAGE_BREAK,'replace'=>false],
    ['type' => MDWORD_TEXT,'text'=> "This is text2."],
    ['type' => MDWORD_PAGE_BREAK,'replace'=>false],
    ['type' => MDWORD_TEXT,'text'=> "This is below."],
    ['text' => dirname(__FILE__).'/full.png','temple'=>file_get_contents(dirname(__FILE__).'/temple/full.xml'),'type' => MDWORD_IMG,'width'=>'100%'],
]);


$TemplateProcessor->saveAs($rtemplate);

