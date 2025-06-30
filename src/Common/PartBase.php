<?php
namespace MDword\Common;

use MDword\Read\Word;
use MDword\Edit\Part\Rels;

class PartBase
{
    /**
     * @var \DOMDocument
     */
    public $DOMDocument;
    
    /**
     * @var Rels
     */
    public $rels = null;
    
    protected $rootPath;
    
    private $id = 0;
    /**
     * @var Word
     */
    public $word;
    
    public $partName = null;
    
    protected $domList = [];
    protected $domIdxToName = [];
    protected $idxExtendIdxs = [];
    
    protected $xmlns = [];
    protected $rIdToNode = [];

    private static $controlCharacters = array();

    public function __construct($word=null) {
        $this->rootPath = dirname(__DIR__);
        
        $this->word = $word;
    }
    
    public function initNameSpaces() {
        $context = $this->DOMDocument->documentElement;
        $xpath = new \DOMXPath($this->DOMDocument);
        foreach( $xpath->query('namespace::*', $context) as $node ) {
            $this->xmlns[$node->localName] = $node->nodeValue;
        }
    }
    
    public function getExist($item,$name) {
        $itemTemps = $item->getElementsByTagName($name);

        if($itemTemps->length === 0) {
            return false;
        }else{
            return true;
        }
    }

    public function getVal($item,$name,$ns='w') {
        $itemTemps = $item->getElementsByTagName($name);

        if($itemTemps->length === 0) {
            return null;
        }elseif($itemTemps->length === 1) {
            return $this->getAttr($itemTemps->item(0),'val',$ns);
        }else{
            //todo
        }
    }

    public function getAttr($item,$name,$ns='w') {
        if(isset($this->xmlns[$ns])) {
            return $item->getAttributeNS($this->xmlns[$ns],$name);
        }else{
            return $item->getAttribute($name);
        }
    }
    
    public function setAttr($item,$name,$value,$ns='w') {
        if(isset($this->xmlns[$ns])) {
            return $item->setAttributeNS($this->xmlns[$ns],$name,$value);
        }else{
            return $item->setAttribute($name,$value);
        }
    }
    
    public function createElementNS($item,$name,$value,$ns='w') {
        if(isset($this->xmlns[$ns])) {
            return $this->DOMDocument->createElementNS($this->xmlns[$ns],$name,$value);
        }else{
            return $this->DOMDocument->createElement($name,$value);
        }
    }
    
    public function hasAttr($item,$name,$ns='w') {
        return $item->hasAttributeNS($this->xmlns[$ns],$name);
    }
    
    public function removeAttr($item,$name) {
        $item->removeAttribute($name);
    }
    
    
    public function __get($name) {
        return $this->$name;
    }
    
    public function markDelete($item) {
        if(!is_null($item)) {
            $item->setAttribute('md',(++$this->id));
        }
    }
    
    protected function removeMarkDelete($item) {
        if(!is_null($item)) {
            $item->removeAttribute('md');
        }
    }
    
    public function deleteMarked() {
        $this->deleteByXpath('//*[@md]');
    }
    
    
    public function deleteByXpath($xpath) {
        $DOMXPath = new \DOMXPath($this->DOMDocument);
        $context = $this->DOMDocument->documentElement;
        $DOMXPath->registerNamespace('w', $this->xmlns['w']);
        $nodes = $DOMXPath->query($xpath, $context);
        foreach( $nodes as $node ) {
            $node->parentNode->removeChild($node);
        }
    }
    
    public function insertAfter($copy,$targetNode) {
        if($nextSibling = $targetNode->nextSibling) {
            if($parentNode = $nextSibling->parentNode) {
                $parentNode->insertBefore($copy,$nextSibling);
            }
        }else{
            if($parentNode = $targetNode->parentNode) {
                $parentNode->appendChild($copy);
            }
        }
    }
    
    public function appendChild($parentNode,$targetNode) {
        $parentNode->appendChild($targetNode);
    }
    
    public function insertBefore($copy,$targetNode) {
        if($parentNode = $targetNode->parentNode) {
            $parentNode->insertBefore($copy,$targetNode);
        }
    }
    
    public function removeChild($item) {
        $parentNode = $item->parentNode;
        $parentNode->removeChild($item);
    }
    
    function pathRelToAbs($RelUrl, $PrefixUrl = '', $SuffixUrl = '')
    {
        $RelUrlRep = str_replace('\\', '/', $RelUrl);
        
        $UrlArr = explode('/', $RelUrlRep);
        
        $NewUrlArr = array();
        
        foreach ($UrlArr as $value) {
            
            if ($value == '..' && ! empty($NewUrlArr)) {
                
                array_pop($NewUrlArr);
            } else if ($value != '..' && $value != '.' && $value != '') {
                
                $NewUrlArr[] = $value;
            }
        }
        
        $UrlStr = ! empty($NewUrlArr) ? implode('/', $NewUrlArr) : '/';
        
        return $PrefixUrl . $UrlStr . $SuffixUrl;
    }
    
    public static function getTempDir()
    {
        $tempDir = sys_get_temp_dir();
        
        if (!empty(self::$tempDir)) {
            $tempDir = self::$tempDir;
        }
        
        return $tempDir;
    }
    
    protected function parserRange($range) {
        $rangeArr = preg_split('/\!|\:/', $range);
        $count = count($rangeArr);
        $rangInfo = [];
        $match = [];
        if($count == 2) {
            $rangInfo[0] = $rangeArr[0];
            preg_match('/([a-z]+?)[\$]*?([0-9]+?)/i', $rangeArr[1],$match);
            $rangInfo[1] = [];
            $rangInfo[2] = [];
            $rangInfo[1][0] = $match[1];
            $rangInfo[2][0] = $match[2];
        }elseif($count == 3) {
            $rangInfo[0] = $rangeArr[0];
            preg_match('/([a-z]+?)[\$]*?([0-9]+?)/i', $rangeArr[1],$match);
            $rangInfo[1] = [];
            $rangInfo[2] = [];
            $rangInfo[1][0] = $match[1];
            $rangInfo[2][0] = $match[2];
            
            preg_match('/([a-z]+?)[\$]*?([0-9]+?)/i', $rangeArr[2],$match);
            $rangInfo[1][1] = $match[1];
            $rangInfo[2][1] = $match[2];
        }else{
            var_dump('parserRange 数量不对！');
        }
        
        $rangInfo[3] = $range;
        return $rangInfo;
    }
    
    protected function initRels($xmlType=19) {
        $partInfo = pathinfo($this->partName);
        $partNameRel = $partInfo['dirname'].'/_rels/'.$partInfo['basename'].'.rels';
        $this->rels = new Rels($this->word, $this->word->getXmlDom($partNameRel));
        $this->rels->partName = $partNameRel;
        $this->rels->partInfo = $partInfo;
        $this->word->parts[$xmlType][] = ['PartName'=>$this->rels->partName,'DOMElement'=>$this->rels->DOMDocument];
    }
    
    public function createNodeByXml($xmlname,$callback=null)
    {
        $filename = MDWORD_SRC_DIRECTORY.'/XmlTemple/'.$xmlname.'.xml';
        if(is_file($filename)) {
            $xml = file_get_contents($filename);
        }else{
            $xml = $xmlname;
        }
        
        if(strpos($xml, '<?') !== 0) {
            $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">'
                .$xml.'</w:document>';
        }
        
        
        $dom = new \DOMDocument();
        $dom->loadXML($xml,LIBXML_NOBLANKS);
        
        if($callback === null) {
            return $this->DOMDocument->importNode($dom->documentElement->firstChild,true);
        }else{
            return $this->DOMDocument->importNode($callback($dom->documentElement),true);
        }
    }
    
    protected function initChartRels($relArr) {
        $partInfo = pathinfo($relArr['PartName']);
        $this->rels = new Rels($this->word, $relArr['dom']);
        $this->rels->partName = $relArr['relName'];
        $this->rels->partInfo = $partInfo;
        
        $this->word->parts[19][] = ['PartName'=>$this->rels->partName,'DOMElement'=>$this->rels->DOMDocument];
    }
    
    
    protected function initCommentRange() {
        if(isset($this->DOMDocument->documentElement->tagList['w:commentRangeStart'])) {
            $commentRangeStartItems = $this->DOMDocument->documentElement->tagList['w:commentRangeStart'];
        }else{
            $commentRangeStartItems = [];
        }
        $tempBlocks = [];
        foreach($commentRangeStartItems as $commentRangeStartItem) {
            $id = $this->getAttr($commentRangeStartItem, 'id');
            $commentRangeEndItem = $this->getCommentRangeEnd($this->DOMDocument,$id);
            $name = $this->commentsblocks[$id];
            $traces = $this->getRangeTrace($id,$commentRangeStartItem, $commentRangeEndItem);
            if(!isset($tempBlocks[$name])) {
                $tempBlocks[$name] = [];
                $size = 0;
            }else{
                $size = sizeof($tempBlocks[$name])+1;
            }

            $tempOneBlock = array_map(function($trace) use ($name,$size) {
                $this->domIdxToName[$trace->idxBegin][] = [$size,$name];
                return $trace->idxBegin;
            }, $traces);

            if(strpos($id,'r') === 0) {
                foreach($tempOneBlock as $nodeId) {
                    if(!isset($this->domList[$nodeId])) {
                        continue;
                    }
                    $ts = $this->domList[$nodeId]->getElementsByTagName('t');
                    foreach($ts as $t) {
                        $t->nodeValue = '';
                    }
                }
            }

            $tempBlocks[$name][] = $tempOneBlock;
        }
        
        return $tempBlocks;
    }
    
    protected function treeToList($node) {
        static $index = 0;
        
        if(is_null($node)) {
            return $index;
        }
        $node->idxBegin = $index;
        $node->tagList = [];
        $this->domList[$index++] = $node;
        if(($node->hasChildNodes())) {
            foreach($node->childNodes as $childNode) {
                if($childNode->nodeType !== 3) {
                    $node->tagList[$childNode->tagName][] = $childNode;
                    $tags = $this->treeToList($childNode);
                    foreach($tags as $tag => $vals) {
                        foreach($vals as $val) {
                            $node->tagList[$tag][] = $val;
                            if ($tag === 'a:blip') {
                                $rId = $this->getAttr($val,'r:embed');
                                $this->rIdToNode[$rId] = $val;
                            }
                        }
                    }
                }
            }
        }
        
        $node->idxEnd = $index-1;

        return $node->tagList;
    }
    
    public function treeToListCallback($node,$callback) {
        if(is_null($node)) {
            return ;
        }
        
        if(($node->hasChildNodes())) {
            foreach($node->childNodes as $childNode) {
                if($childNode->nodeType !== 3) {
                    $this->treeToListCallback($callback($childNode),$callback);
                }
            }
        }
    }

    public function getTextContent($nodes) {
        $text = '';
        $this->treeToListCallback($nodes,function($node) use (&$text) {
            //jump delete
            if($this->getAttr($node,'md',null)) {
                return null;
            }

            if($node->localName === 't') {
                $text .= $node->textContent;
                return null;
            }

            return $node;
        });

        return $text;
    }
    
    
    protected function replace($node, $targetNode) {
        if($parentNode = $targetNode->parentNode) {
            $parentNode->replaceChild($node, $targetNode);
        }
    }
    
    protected function getTreeToListBeginIdOldToNew($node,$beginId,$main=true) {
        static $beginIdOldToNew = [];
        static $index = 0;
        if($main) {
            $beginIdOldToNew = [];
            $index = $beginId;
        }
        
        if(!isset($beginIdOldToNew[$node->idxBegin])) {
            $beginIdOldToNew[$node->idxBegin] = $index;
        }
        
        $index++;
        if(($node->hasChildNodes())) {
            foreach($node->childNodes as $childNode) {
                if($childNode->nodeType !== 3) {
                    $this->getTreeToListBeginIdOldToNew($childNode,0,false);
                }
            }
            
        }
        
        return $beginIdOldToNew;
    }
    
    protected function getCommentRangeEnd($parentNode,$id) {
        if(isset($parentNode->tagList['w:commentRangeEnd'])) {
            $commentRangeEndItems = $parentNode->tagList['w:commentRangeEnd'];
        }else{
            $commentRangeEndItems = [];
        }
        foreach($commentRangeEndItems as $commentRangeEndItem) {
            $eid = $this->getAttr($commentRangeEndItem, 'id');
            
            if($id === $eid) {
                return $commentRangeEndItem;
            }
        }
        
        return null;
    }

    protected function getParentToTag($parentNode, $target) {
        while($parentNode = $parentNode->parentNode) {
            if($parentNode->localName === $target) {
                return $parentNode;
            }
        }
    }
    
    protected function getRangeTrace($id,$commentRangeStartItem,$commentRangeEndItem) {
        $delTags = ['commentRangeStart'=>0];
        $traces = [];
        $startParentNode = $parentNode = $commentRangeStartItem;
        $parentNodeCount = 0;
        while($parentNode = $parentNode->parentNode) {
            $commentRangeEndItem = $this->getCommentRangeEnd($parentNode,$id);
            if(is_null($commentRangeEndItem)) {
                $startParentNode = $parentNode;
            }else{
                break;
            }
            $parentNodeCount++;
        }
        
        if(!isset($delTags[$startParentNode->localName])) {
            $traces[] = $startParentNode;
        }
        
        $nextNodeCount = 0;
        $nextSibling = $startParentNode->nextSibling;
        $nextNodeCount++;
        while(true) {
            if(is_null($nextSibling)) {
                break;
            }
            
            if($nextSibling === $commentRangeEndItem) {
                break;
            }
            
            if($this->getCommentRangeEnd($nextSibling,$id) !== null) {
                $childNodes = $nextSibling->childNodes;
                $find = false;
                $preCount = 0;
                $endCount = 0;
                $childNodesTemp = [];
                foreach ($childNodes as $childNode) {
                    if($find === false) {
                        $childNodesTemp[] = $childNode;
                    }
                    
                    if($find === false && ($childNode === $commentRangeEndItem || $this->getCommentRangeEnd($childNode,$id) !== null)) {
                        $find = true;
                        continue;
                    }
                    
                    if($this->isNeedSpace($childNode)) {
                        if($find) {
                            $endCount++;
                        }else{
                            $preCount++;
                        }
                    }
                }
                
                if($endCount === 0) {
                    if(!isset($delTags[$nextSibling->localName])) {
                        $traces[] = $nextSibling;
                    }
                }elseif(count($traces) === 0) {
                    foreach($childNodesTemp as $trace) {
                        if(!isset($delTags[$trace->localName])) {
                            $traces[] = $trace;
                        }
                    }
                }
                
                break;
            }else{
                if(!isset($delTags[$nextSibling->localName])) {
                    $traces[] = $nextSibling;
                }
            }
            
            $nextSibling = $nextSibling->nextSibling;
            $nextNodeCount++;
        }
        
        return array_values($traces);
    }
    
    protected function isNeedSpace($node) {
        $ts = $node->getElementsByTagName('t');
        if($ts->length > 0) {
            return true;
        }else{
            return false;
        }
    }
    
    protected function showXml($node) {
        echo $this->DOMDocument->saveXML($node);exit;
    }
    
    protected function htmlspecialcharsBase($string,$needConvertSpecialFont = true) {
        $string = $this->my_html_entity_decode($string);
        $string = $this->filterUtf8($this->filterSpecailCodeForWord($string));

        $string = $this->controlCharacterPHP2OOXML($string);
        //preg string build by https://github.com/mkdreams/MDfontSubChar

        if($needConvertSpecialFont === true) {
            //SimSun-ExtB
            $string = preg_replace_callback('/(\{\%\{\d\}\%\}|)[\x{20000}-\x{2a6d6}\x{2a700}-\x{2b734}\x{2b740}-\x{2b81d}\x{2b8b8}-\x{2b8b9}\x{2bac7}-\x{2bac8}\x{2bb5f}-\x{2bb60}\x{2bb62}-\x{2bb63}\x{2bb7c}-\x{2bb7d}\x{2bb83}-\x{2bb84}\x{2bc1b}-\x{2bc1c}\x{2bd77}\x{2bd87}\x{2bdf7}\x{2be29}\x{2c029}-\x{2c02a}\x{2c0a9}\x{2c0ca}\x{2c1d5}\x{2c1d9}\x{2c1f9}\x{2c27c}\x{2c288}\x{2c2a4}\x{2c317}\x{2c35b}\x{2c361}\x{2c364}\x{2c488}\x{2c494}\x{2c497}\x{2c542}\x{2c613}\x{2c618}\x{2c621}\x{2c629}\x{2c62b}-\x{2c62d}\x{2c62f}\x{2c642}\x{2c64a}-\x{2c64b}\x{2c72c}\x{2c72f}\x{2c79f}\x{2c7c1}\x{2c7fd}\x{2c8d9}\x{2c8de}\x{2c8e1}\x{2c8f3}\x{2c907}\x{2c90a}\x{2c91d}\x{2ca02}\x{2ca0e}\x{2ca7d}\x{2caa9}\x{2cb29}\x{2cb2d}-\x{2cb2e}\x{2cb31}\x{2cb38}-\x{2cb39}\x{2cb3b}\x{2cb3f}\x{2cb41}\x{2cb4a}\x{2cb4e}\x{2cb5a}-\x{2cb5b}\x{2cb64}\x{2cb69}\x{2cb6c}\x{2cb6f}\x{2cb73}\x{2cb76}\x{2cb78}\x{2cb7c}\x{2cbb1}\x{2cbbf}-\x{2cbc0}\x{2cbce}\x{2cc56}\x{2cc5f}\x{2ccf5}-\x{2ccf6}\x{2ccfd}\x{2ccff}\x{2cd02}-\x{2cd03}\x{2cd0a}\x{2cd8b}\x{2cd8d}\x{2cd8f}-\x{2cd90}\x{2cd9f}-\x{2cda0}\x{2cda8}\x{2cdad}-\x{2cdae}\x{2cdd5}\x{2ce18}\x{2ce1a}\x{2ce23}\x{2ce26}\x{2ce2a}\x{2ce7c}\x{2ce88}\x{2ce93}]+/u', function($match){
                if($match[1] === '') {
                    return '{%{0}%}'.$match[0].'{%{/0}%}';
                }else{
                    return $match[0];
                }
            }, $string);
    
            //Cambria Math
            $string = preg_replace_callback('/(\{\%\{\d\}\%\}|)[\x{0}\x{d}\x{100}\x{102}-\x{112}\x{114}-\x{11a}\x{11c}-\x{12a}\x{12c}-\x{143}\x{145}-\x{147}\x{149}-\x{14c}\x{14e}-\x{151}\x{154}-\x{15f}\x{162}-\x{16a}\x{16c}-\x{177}\x{179}-\x{191}\x{193}-\x{1cd}\x{1cf}\x{1d1}\x{1d3}\x{1d5}\x{1d7}\x{1d9}\x{1db}\x{1dd}-\x{1f8}\x{1fa}-\x{250}\x{252}-\x{260}\x{262}-\x{2c5}\x{2c8}\x{2cc}-\x{2d8}\x{2da}-\x{2db}\x{2dd}-\x{377}\x{37a}-\x{37f}\x{384}-\x{38a}\x{38c}\x{38e}-\x{390}\x{3aa}-\x{3b0}\x{3c2}\x{3ca}-\x{400}\x{402}-\x{40f}\x{450}\x{452}-\x{52f}\x{531}-\x{556}\x{559}-\x{55f}\x{561}-\x{587}\x{589}-\x{58a}\x{58d}-\x{58f}\x{e3f}\x{1d00}-\x{1dca}\x{1dfe}-\x{1f15}\x{1f18}-\x{1f1d}\x{1f20}-\x{1f45}\x{1f48}-\x{1f4d}\x{1f50}-\x{1f57}\x{1f59}\x{1f5b}\x{1f5d}\x{1f5f}-\x{1f7d}\x{1f80}-\x{1fb4}\x{1fb6}-\x{1fc4}\x{1fc6}-\x{1fd3}\x{1fd6}-\x{1fdb}\x{1fdd}-\x{1fef}\x{1ff2}-\x{1ff4}\x{1ff6}-\x{1ffe}\x{2000}-\x{200f}\x{2011}-\x{2012}\x{2017}\x{201b}\x{201f}\x{2024}\x{202f}\x{2034}\x{203c}-\x{203e}\x{2044}-\x{2046}\x{2057}\x{205e}-\x{205f}\x{2061}-\x{2063}\x{2070}-\x{2071}\x{2074}-\x{208e}\x{2090}-\x{209c}\x{20a0}-\x{20ab}\x{20ad}-\x{20b5}\x{20b8}-\x{20ba}\x{20bc}-\x{20bf}\x{20d0}-\x{20df}\x{20e1}\x{20e5}-\x{20e6}\x{20e8}-\x{20ea}\x{2100}-\x{2102}\x{2104}\x{2106}-\x{2108}\x{210a}-\x{2115}\x{2117}-\x{2120}\x{2123}-\x{214f}\x{2153}-\x{215e}\x{2183}-\x{2184}\x{2194}-\x{2195}\x{219a}-\x{2207}\x{2209}-\x{220e}\x{2210}\x{2212}-\x{2214}\x{2216}-\x{2219}\x{221b}-\x{221c}\x{2221}-\x{2222}\x{2224}\x{2226}\x{222c}-\x{222d}\x{222f}-\x{2233}\x{2238}-\x{223c}\x{223e}-\x{2247}\x{2249}-\x{224b}\x{224d}-\x{2251}\x{2253}-\x{225f}\x{2262}-\x{2263}\x{2268}-\x{226d}\x{2270}-\x{2294}\x{2296}-\x{2298}\x{229a}-\x{22a4}\x{22a6}-\x{22be}\x{22c0}-\x{2311}\x{2313}-\x{232a}\x{2330}-\x{23cf}\x{23dc}-\x{23e0}\x{246a}-\x{2473}\x{24ea}-\x{24f4}\x{24ff}\x{2592}\x{25a2}-\x{25b1}\x{25b4}-\x{25bb}\x{25be}-\x{25c5}\x{25c8}-\x{25ca}\x{25cc}-\x{25cd}\x{25d0}-\x{25e1}\x{25e6}-\x{25ff}\x{2660}-\x{2663}\x{2666}\x{2720}\x{2776}-\x{277f}\x{27c0}-\x{27ff}\x{2900}-\x{2aff}\x{2b04}\x{2b06}-\x{2b07}\x{2b0c}-\x{2b0d}\x{2b1a}\x{2c60}-\x{2c7f}\x{2e17}\x{a64c}-\x{a64d}\x{a717}-\x{a71a}\x{a720}-\x{a721}\x{fb00}-\x{fb04}\x{fb13}-\x{fb17}\x{fe00}\x{1d400}-\x{1d454}\x{1d456}-\x{1d49c}\x{1d49e}-\x{1d49f}\x{1d4a2}\x{1d4a5}-\x{1d4a6}\x{1d4a9}-\x{1d4ac}\x{1d4ae}-\x{1d4b9}\x{1d4bb}\x{1d4bd}-\x{1d4c0}\x{1d4c2}-\x{1d4c3}\x{1d4c5}-\x{1d505}\x{1d507}-\x{1d50a}\x{1d50d}-\x{1d514}\x{1d516}-\x{1d51c}\x{1d51e}-\x{1d539}\x{1d53b}-\x{1d53e}\x{1d540}-\x{1d544}\x{1d546}\x{1d54a}-\x{1d550}\x{1d552}-\x{1d6a5}\x{1d6a8}-\x{1d7cb}\x{1d7ce}-\x{1d7ff}\x{1d4c1}]+/u', function($match){
                if($match[1] === '') {
                    return '{%{1}%}'.$match[0].'{%{/1}%}';
                }else{
                    return $match[0];
                }
            }, $string);
        }

        return htmlspecialchars($string,ENT_COMPAT, 'UTF-8');
    }
    
    protected function my_html_entity_decode($string) {
        static $stringMd5 = '';
        
        $md5 = md5($string);
        if($stringMd5 !== $md5) {
            $stringMd5 = $md5;
            $string = html_entity_decode($string);
            $string = $this->my_html_entity_decode($string);
        }else{
            $stringMd5 = '';
        }
        return $string;
    }
    
    private function filterUtf8($string)
    {
        if ($string) {
            preg_match_all('/[\x00-\x7F]|[\xC0-\xDF][\x80-\xBF]|[\xE0-\xEF][\x80-\xBF]{2}|[\xF0-\xF7][\x80-\xBF]{3}/s', $string, $matches);
            return implode('', $matches[0]);
        }
        return $string;
    }
    
    private function filterSpecailCodeForWord($string) {
        $string = str_replace(array('￾'),'', $string);
        $string = preg_replace('/[\x14-\x1F]|[\x08]/i', '', $string);
        return $string;
    }
    
     /**
     * Convert from PHP control character to OpenXML escaped control character
     *
     * @param string $value Value to escape
     * @return string
     */
    public static function controlCharacterPHP2OOXML($value = '')
    {
        if (empty(self::$controlCharacters)) {
            self::buildControlCharacters();
        }

        return str_replace(array_values(self::$controlCharacters), array_keys(self::$controlCharacters), $value);
    }

    /**
     * Build control characters array.
     *
     * @return void
     */
    private static function buildControlCharacters()
    {
        for ($i = 0; $i <= 19; ++$i) {
            if ($i != 9 && $i != 10 && $i != 13) {
                $find = '_x' . sprintf('%04s', strtoupper(dechex($i))) . '_';
                $replace = chr($i);
                self::$controlCharacters[$find] = $replace;
            }
        }
    }

    /**
     * @param \DOMDocument $rPr
     * @param \DOMDocument $default
     */
    public function mergerPr($rPr,$default) {
        if(is_null($rPr)) {
            return null;
        }
        if(is_null($default)) {
            return null;
        }
        
        foreach($default->childNodes as $node) {
            $localName = $node->localName;
            $items = $rPr->getElementsByTagName($localName);
            if($items->length > 0) {
                $item = $items->item(0);
                $rPrAttrs = $this->getAttributes($item);
                foreach($node->attributes as $attribute) {
                    if(!isset($rPrAttrs[$attribute->nodeName]) && !isset($rPrAttrs[$attribute->nodeName.'Theme'])) {
                        $this->setAttr($item, $attribute->localName, $attribute->value);
                    }
                }
            }else{
                $copyNode = $this->createElementNS($rPr,$node->nodeName,null);
                foreach($node->attributes as $attribute) {
                    $this->setAttr($copyNode, $attribute->localName, $attribute->value);
                }
                $this->appendChild($rPr, $copyNode);
            }
        }
        
        return $rPr;
    }
    
    private function getAttributes($node) {
        $attrs = [];
        foreach($node->attributes as $attribute) {
            $attrs[$attribute->nodeName] = [$attribute->prefix,$attribute->name,$attribute->value];
        }
        
        return $attrs;
    }

    public function free() {
        $bindValueFunc = new \ReflectionMethod($this,'treeToList');
        $statics = $bindValueFunc->getStaticVariables();
        $statics['index'] = 0;

        $bindValueFunc = new \ReflectionMethod($this,'getTreeToListBeginIdOldToNew');
        $statics = $bindValueFunc->getStaticVariables();
        $statics['beginIdOldToNew'] = [];
        $statics['index'] = [];
    }

    public function getIndex($pNode,$subNode) {
        $tag = $subNode->localName;
        $tagNodes = $pNode->getElementsByTagName($tag);
        $idx = -1;
        foreach($tagNodes as $index => $tagNode) {
            if($tagNode === $subNode) {
                $idx = $index;
                break;
            }
        }
        
        return $idx;
    }
}