<?php 
require_once(__DIR__.'/../../../Autoloader.php');

use MDword\WordProcessor;


$template = __DIR__.'/temple.docx';
$rtemplate = __DIR__.'/r-temple.docx';

$TemplateProcessor = new WordProcessor();
$TemplateProcessor->load($template);


// $TemplateProcessor->showMedies();//help method for get md5

$TemplateProcessor->clones('row',3);
// $TemplateProcessor->setValue('rowIndex#0',0);
// $TemplateProcessor->setValue('rowIndex#1',1);
// $TemplateProcessor->setValue('rowIndex#2',2);
// $TemplateProcessor->setValue('rowIndex#2',3);
$TemplateProcessor->setValue('rowIndex#2',4);

// $TemplateProcessor->setImageValue('rowImage#0',dirname(__FILE__).'/img.jpg');
// $TemplateProcessor->setImageValue('rowImage#1',dirname(__FILE__).'/img2.bmp');
// $TemplateProcessor->setImageValue('rowImage#2',dirname(__FILE__).'/img3.png');

$rows = [
    ['index'=>0,'image'=> dirname(__FILE__).'/img.jpg'],
    ['index'=>1,'image'=> dirname(__FILE__).'/img2.bmp'],
    ['index'=>2,'image'=> dirname(__FILE__).'/img3.png'],
];

$bind = $TemplateProcessor->getBind($rows);
$bind->bindValue('row#2',[])
// ->bindValue('rowIndex#2',['index'],'row#2')
// ->bindValue('rowImage#2',['image'],'row#2',MDWORD_IMG)
;

$TemplateProcessor->setValue('rowIndex#2#2',2);

$TemplateProcessor->setImageValue('rowImage#2#1',dirname(__FILE__).'/img.jpg');

$TemplateProcessor->setImageValue('image insert', dirname(__FILE__).'/img.jpg');

$TemplateProcessor->setImageValue('image replace', dirname(__FILE__).'/img.jpg');

$TemplateProcessor->setImageValue('2529be7711acbb60c7e4ac1693c680a0', dirname(__FILE__).'/img.jpg');


$TemplateProcessor->saveAs($rtemplate);

