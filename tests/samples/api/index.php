<?php 
require_once(__DIR__.'/../../../Autoloader.php');
use MDword\Api\GetWord;
use MDword\Api\ParseWord;

$rtemplate = __DIR__.'/r-temple.docx';

$parseWord = new ParseWord();
$parseWord->getBlockList();

// $getWord = new GetWord();
// $content = $getWord->build();
// file_put_contents($rtemplate, $content);