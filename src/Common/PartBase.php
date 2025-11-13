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
            $pregUnicode = [
                'SimSun-ExtB' => '([\x{20000}-\x{2a6d6}\x{2a700}-\x{2b734}\x{2b740}-\x{2b81d}\x{2b8b8}-\x{2b8b9}\x{2bac7}-\x{2bac8}\x{2bb5f}-\x{2bb60}\x{2bb62}-\x{2bb63}\x{2bb7c}-\x{2bb7d}\x{2bb83}-\x{2bb84}\x{2bc1b}-\x{2bc1c}\x{2bd77}\x{2bd87}\x{2bdf7}\x{2be29}\x{2c029}-\x{2c02a}\x{2c0a9}\x{2c0ca}\x{2c1d5}\x{2c1d9}\x{2c1f9}\x{2c27c}\x{2c288}\x{2c2a4}\x{2c317}\x{2c35b}\x{2c361}\x{2c364}\x{2c488}\x{2c494}\x{2c497}\x{2c542}\x{2c613}\x{2c618}\x{2c621}\x{2c629}\x{2c62b}-\x{2c62d}\x{2c62f}\x{2c642}\x{2c64a}-\x{2c64b}\x{2c72c}\x{2c72f}\x{2c79f}\x{2c7c1}\x{2c7fd}\x{2c8d9}\x{2c8de}\x{2c8e1}\x{2c8f3}\x{2c907}\x{2c90a}\x{2c91d}\x{2ca02}\x{2ca0e}\x{2ca7d}\x{2caa9}\x{2cb29}\x{2cb2d}-\x{2cb2e}\x{2cb31}\x{2cb38}-\x{2cb39}\x{2cb3b}\x{2cb3f}\x{2cb41}\x{2cb4a}\x{2cb4e}\x{2cb5a}-\x{2cb5b}\x{2cb64}\x{2cb69}\x{2cb6c}\x{2cb6f}\x{2cb73}\x{2cb76}\x{2cb78}\x{2cb7c}\x{2cbb1}\x{2cbbf}-\x{2cbc0}\x{2cbce}\x{2cc56}\x{2cc5f}\x{2ccf5}-\x{2ccf6}\x{2ccfd}\x{2ccff}\x{2cd02}-\x{2cd03}\x{2cd0a}\x{2cd8b}\x{2cd8d}\x{2cd8f}-\x{2cd90}\x{2cd9f}-\x{2cda0}\x{2cda8}\x{2cdad}-\x{2cdae}\x{2cdd5}\x{2ce18}\x{2ce1a}\x{2ce23}\x{2ce26}\x{2ce2a}\x{2ce7c}\x{2ce88}\x{2ce93}]+)',
                'Arial Unicode MS' => '([\x{100}\x{102}-\x{112}\x{114}-\x{11a}\x{11c}-\x{12a}\x{12c}-\x{143}\x{145}-\x{147}\x{149}-\x{14c}\x{14e}-\x{151}\x{154}-\x{15f}\x{162}-\x{16a}\x{16c}-\x{177}\x{179}-\x{191}\x{193}-\x{1cd}\x{1cf}\x{1d1}\x{1d3}\x{1d5}\x{1d7}\x{1d9}\x{1db}\x{1dd}-\x{1f5}\x{1fa}-\x{217}\x{250}\x{252}-\x{260}\x{262}-\x{2a8}\x{2b0}-\x{2c5}\x{2c8}\x{2cc}-\x{2d8}\x{2da}-\x{2db}\x{2dd}-\x{2de}\x{2e0}-\x{2e9}\x{300}-\x{345}\x{360}-\x{361}\x{374}-\x{375}\x{37a}\x{37e}\x{384}-\x{38a}\x{38c}\x{38e}-\x{390}\x{3aa}-\x{3b0}\x{3c2}\x{3ca}-\x{3ce}\x{3d0}-\x{3d6}\x{3da}\x{3dc}\x{3de}\x{3e0}\x{3e2}-\x{3f3}\x{402}-\x{40c}\x{40e}-\x{40f}\x{452}-\x{45c}\x{45e}-\x{486}\x{490}-\x{4c4}\x{4c7}-\x{4c8}\x{4cb}-\x{4cc}\x{4d0}-\x{4eb}\x{4ee}-\x{4f5}\x{4f8}-\x{4f9}\x{531}-\x{556}\x{559}-\x{55f}\x{561}-\x{587}\x{589}\x{591}-\x{5a1}\x{5a3}-\x{5b9}\x{5bb}-\x{5c4}\x{5d0}-\x{5ea}\x{5f0}-\x{5f4}\x{60c}\x{61b}\x{61f}\x{621}-\x{63a}\x{640}-\x{652}\x{660}-\x{66d}\x{670}-\x{6b7}\x{6ba}-\x{6be}\x{6c0}-\x{6ce}\x{6d0}-\x{6ed}\x{6f0}-\x{6f9}\x{901}-\x{903}\x{905}-\x{939}\x{93c}-\x{94d}\x{950}-\x{954}\x{958}-\x{970}\x{981}-\x{983}\x{985}-\x{98c}\x{98f}-\x{990}\x{993}-\x{9a8}\x{9aa}-\x{9b0}\x{9b2}\x{9b6}-\x{9b9}\x{9bc}\x{9be}-\x{9c4}\x{9c7}-\x{9c8}\x{9cb}-\x{9cd}\x{9d7}\x{9dc}-\x{9dd}\x{9df}-\x{9e3}\x{9e6}-\x{9fa}\x{a02}\x{a05}-\x{a0a}\x{a0f}-\x{a10}\x{a13}-\x{a28}\x{a2a}-\x{a30}\x{a32}-\x{a33}\x{a35}-\x{a36}\x{a38}-\x{a39}\x{a3c}\x{a3e}-\x{a42}\x{a47}-\x{a48}\x{a4b}-\x{a4d}\x{a59}-\x{a5c}\x{a5e}\x{a66}-\x{a74}\x{a81}-\x{a83}\x{a85}-\x{a8b}\x{a8d}\x{a8f}-\x{a91}\x{a93}-\x{aa8}\x{aaa}-\x{ab0}\x{ab2}-\x{ab3}\x{ab5}-\x{ab9}\x{abc}-\x{ac5}\x{ac7}-\x{ac9}\x{acb}-\x{acd}\x{ad0}\x{ae0}\x{ae6}-\x{aef}\x{b01}-\x{b03}\x{b05}-\x{b0c}\x{b0f}-\x{b10}\x{b13}-\x{b28}\x{b2a}-\x{b30}\x{b32}-\x{b33}\x{b36}-\x{b39}\x{b3c}-\x{b43}\x{b47}-\x{b48}\x{b4b}-\x{b4d}\x{b56}-\x{b57}\x{b5c}-\x{b5d}\x{b5f}-\x{b61}\x{b66}-\x{b70}\x{b82}-\x{b83}\x{b85}-\x{b8a}\x{b8e}-\x{b90}\x{b92}-\x{b95}\x{b99}-\x{b9a}\x{b9c}\x{b9e}-\x{b9f}\x{ba3}-\x{ba4}\x{ba8}-\x{baa}\x{bae}-\x{bb5}\x{bb7}-\x{bb9}\x{bbe}-\x{bc2}\x{bc6}-\x{bc8}\x{bca}-\x{bcd}\x{bd7}\x{be7}-\x{bf2}\x{c01}-\x{c03}\x{c05}-\x{c0c}\x{c0e}-\x{c10}\x{c12}-\x{c28}\x{c2a}-\x{c33}\x{c35}-\x{c39}\x{c3e}-\x{c44}\x{c46}-\x{c48}\x{c4a}-\x{c4d}\x{c55}-\x{c56}\x{c60}-\x{c61}\x{c66}-\x{c6f}\x{c82}-\x{c83}\x{c85}-\x{c8c}\x{c8e}-\x{c90}\x{c92}-\x{ca8}\x{caa}-\x{cb3}\x{cb5}-\x{cb9}\x{cbe}-\x{cc4}\x{cc6}-\x{cc8}\x{cca}-\x{ccd}\x{cd5}-\x{cd6}\x{cde}\x{ce0}-\x{ce1}\x{ce6}-\x{cef}\x{d02}-\x{d03}\x{d05}-\x{d0c}\x{d0e}-\x{d10}\x{d12}-\x{d28}\x{d2a}-\x{d39}\x{d3e}-\x{d43}\x{d46}-\x{d48}\x{d4a}-\x{d4d}\x{d57}\x{d60}-\x{d61}\x{d66}-\x{d6f}\x{e01}-\x{e3a}\x{e3f}-\x{e5b}\x{e81}-\x{e82}\x{e84}\x{e87}-\x{e88}\x{e8a}\x{e8d}\x{e94}-\x{e97}\x{e99}-\x{e9f}\x{ea1}-\x{ea3}\x{ea5}\x{ea7}\x{eaa}-\x{eab}\x{ead}-\x{eb9}\x{ebb}-\x{ebd}\x{ec0}-\x{ec4}\x{ec6}\x{ec8}-\x{ecd}\x{ed0}-\x{ed9}\x{edc}-\x{edd}\x{f00}-\x{f47}\x{f49}-\x{f69}\x{f71}-\x{f8b}\x{f90}-\x{f95}\x{f97}\x{f99}-\x{fad}\x{fb1}-\x{fb7}\x{fb9}\x{10a0}-\x{10c5}\x{10d0}-\x{10f6}\x{10fb}\x{1100}-\x{1159}\x{115f}-\x{11a2}\x{11a8}-\x{11f9}\x{1e00}-\x{1e9b}\x{1ea0}-\x{1ef9}\x{1f00}-\x{1f15}\x{1f18}-\x{1f1d}\x{1f20}-\x{1f45}\x{1f48}-\x{1f4d}\x{1f50}-\x{1f57}\x{1f59}\x{1f5b}\x{1f5d}\x{1f5f}-\x{1f7d}\x{1f80}-\x{1fb4}\x{1fb6}-\x{1fc4}\x{1fc6}-\x{1fd3}\x{1fd6}-\x{1fdb}\x{1fdd}-\x{1fef}\x{1ff2}-\x{1ff4}\x{1ff6}-\x{1ffe}\x{2000}-\x{200d}\x{2011}-\x{2012}\x{2017}\x{201b}\x{201f}\x{2023}-\x{2024}\x{2027}-\x{2029}\x{2031}\x{2034}\x{2036}-\x{2038}\x{203c}-\x{2046}\x{2070}\x{2074}-\x{208e}\x{20a0}-\x{20ab}\x{20d0}-\x{20e1}\x{2100}-\x{2102}\x{2104}\x{2106}-\x{2108}\x{210a}-\x{2115}\x{2117}-\x{2120}\x{2123}-\x{2138}\x{2153}-\x{215f}\x{216c}-\x{216f}\x{217a}-\x{2182}\x{2194}-\x{2195}\x{219a}-\x{21ea}\x{2200}-\x{2207}\x{2209}-\x{220e}\x{2210}\x{2212}-\x{2214}\x{2216}-\x{2219}\x{221b}-\x{221c}\x{2221}-\x{2222}\x{2224}\x{2226}\x{222c}-\x{222d}\x{222f}-\x{2233}\x{2238}-\x{223c}\x{223e}-\x{2247}\x{2249}-\x{224b}\x{224d}-\x{2251}\x{2253}-\x{225f}\x{2262}-\x{2263}\x{2268}-\x{226d}\x{2270}-\x{2294}\x{2296}-\x{2298}\x{229a}-\x{22a4}\x{22a6}-\x{22be}\x{22c0}-\x{22f1}\x{2300}\x{2302}-\x{2311}\x{2313}-\x{237a}\x{2395}\x{2400}-\x{2424}\x{2440}-\x{244a}\x{246a}-\x{2473}\x{249c}-\x{24ea}\x{254c}-\x{254f}\x{2574}-\x{2580}\x{2590}-\x{2592}\x{25a2}-\x{25b1}\x{25b4}-\x{25bb}\x{25be}-\x{25c5}\x{25c8}-\x{25ca}\x{25cc}-\x{25cd}\x{25d0}-\x{25e1}\x{25e6}-\x{25ef}\x{2600}-\x{2604}\x{2607}-\x{2608}\x{260a}-\x{2613}\x{261a}-\x{263f}\x{2641}\x{2643}-\x{266f}\x{2701}-\x{2704}\x{2706}-\x{2709}\x{270c}-\x{2727}\x{2729}-\x{274b}\x{274d}\x{274f}-\x{2752}\x{2756}\x{2758}-\x{275e}\x{2761}-\x{2767}\x{2776}-\x{2794}\x{2798}-\x{27af}\x{27b1}-\x{27be}\x{3004}\x{3018}-\x{301c}\x{301f}-\x{3020}\x{302a}-\x{3037}\x{303f}\x{3094}\x{3099}-\x{309a}\x{30f7}-\x{30fb}\x{312a}-\x{312c}\x{3131}-\x{318e}\x{3190}-\x{319f}\x{3200}-\x{321c}\x{322a}-\x{3230}\x{3232}-\x{3243}\x{3260}-\x{327b}\x{327f}-\x{32a2}\x{32a4}-\x{32b0}\x{32c0}-\x{32cb}\x{32d0}-\x{32fe}\x{3300}-\x{3376}\x{337b}-\x{338d}\x{3390}-\x{339b}\x{339f}-\x{33a0}\x{33a2}-\x{33c3}\x{33c5}-\x{33cd}\x{33cf}-\x{33d0}\x{33d3}-\x{33d4}\x{33d6}-\x{33dd}\x{33e0}-\x{33fe}\x{ac00}-\x{d7a3}\x{e801}-\x{e805}\x{e867}\x{e890}\x{e8a5}-\x{e8a7}\x{f001}-\x{f002}\x{f700}-\x{f71a}\x{f71c}-\x{f71d}\x{f900}-\x{f92b}\x{f92d}-\x{f978}\x{f97a}-\x{f994}\x{f996}-\x{f9e6}\x{f9e8}-\x{f9f0}\x{f9f2}-\x{fa0b}\x{fa10}\x{fa12}\x{fa15}-\x{fa17}\x{fa19}-\x{fa1e}\x{fa22}\x{fa25}-\x{fa26}\x{fa2a}-\x{fa2d}\x{fb00}-\x{fb06}\x{fb13}-\x{fb17}\x{fb1e}-\x{fb36}\x{fb38}-\x{fb3c}\x{fb3e}\x{fb40}-\x{fb41}\x{fb43}-\x{fb44}\x{fb46}-\x{fbb1}\x{fbd3}-\x{fd3f}\x{fd50}-\x{fd8f}\x{fd92}-\x{fdc7}\x{fdf0}-\x{fdfb}\x{fe20}-\x{fe23}\x{fe32}\x{fe58}\x{fe70}-\x{fe72}\x{fe74}\x{fe76}-\x{fefc}\x{ff61}-\x{ffbe}\x{ffc2}-\x{ffc7}\x{ffca}-\x{ffcf}\x{ffd2}-\x{ffd7}\x{ffda}-\x{ffdc}\x{ffe6}\x{ffe8}-\x{ffee}\x{fffc}-\x{fffd}]+)',
                'Cambria Math' => '([\x{0}\x{d}\x{1f6}-\x{1f8}\x{218}-\x{24f}\x{2a9}-\x{2af}\x{2df}\x{2ea}-\x{2ff}\x{346}-\x{35f}\x{362}-\x{373}\x{376}-\x{377}\x{37b}-\x{37d}\x{37f}\x{3cf}\x{3d7}-\x{3d9}\x{3db}\x{3dd}\x{3df}\x{3e1}\x{3f4}-\x{400}\x{40d}\x{450}\x{45d}\x{487}-\x{48f}\x{4c5}-\x{4c6}\x{4c9}-\x{4ca}\x{4cd}-\x{4cf}\x{4ec}-\x{4ed}\x{4f6}-\x{4f7}\x{4fa}-\x{52f}\x{58a}\x{58d}-\x{58f}\x{1d00}-\x{1dca}\x{1dfe}-\x{1dff}\x{1e9c}-\x{1e9f}\x{1efa}-\x{1eff}\x{200e}-\x{200f}\x{202f}\x{2057}\x{205e}-\x{205f}\x{2061}-\x{2063}\x{2071}\x{2090}-\x{209c}\x{20ad}-\x{20b5}\x{20b8}-\x{20ba}\x{20bc}-\x{20bf}\x{20e5}-\x{20e6}\x{20e8}-\x{20ea}\x{2139}-\x{214f}\x{2183}-\x{2184}\x{21eb}-\x{21ff}\x{22f2}-\x{22ff}\x{2301}\x{237b}-\x{2394}\x{2396}-\x{23cf}\x{23dc}-\x{23e0}\x{24eb}-\x{24f4}\x{24ff}\x{25f0}-\x{25ff}\x{27c0}-\x{27ff}\x{2900}-\x{2aff}\x{2b04}\x{2b06}-\x{2b07}\x{2b0c}-\x{2b0d}\x{2b1a}\x{2c60}-\x{2c7f}\x{2e17}\x{a64c}-\x{a64d}\x{a717}-\x{a71a}\x{a720}-\x{a721}\x{fe00}\x{1d400}-\x{1d454}\x{1d456}-\x{1d49c}\x{1d49e}-\x{1d49f}\x{1d4a2}\x{1d4a5}-\x{1d4a6}\x{1d4a9}-\x{1d4ac}\x{1d4ae}-\x{1d4b9}\x{1d4bb}\x{1d4bd}-\x{1d4c0}\x{1d4c2}-\x{1d4c3}\x{1d4c5}-\x{1d505}\x{1d507}-\x{1d50a}\x{1d50d}-\x{1d514}\x{1d516}-\x{1d51c}\x{1d51e}-\x{1d539}\x{1d53b}-\x{1d53e}\x{1d540}-\x{1d544}\x{1d546}\x{1d54a}-\x{1d550}\x{1d552}-\x{1d6a5}\x{1d6a8}-\x{1d7cb}\x{1d7ce}-\x{1d7ff}\x{1d4c1}]+)',
                'Leelawadee UI' => '([\x{ede}-\x{edf}\x{1780}-\x{17dd}\x{17e0}-\x{17e9}\x{17f0}-\x{17f9}\x{19e0}-\x{1a1b}\x{1a1e}-\x{1a1f}\x{a9cf}]+)',
                'Myanmar Text' => '([\x{1000}-\x{109f}\x{2060}\x{a9e0}-\x{a9fe}\x{aa60}-\x{aa7f}]+)',
                'Segoe UI Emoji' => '([\x{2049}\x{20e3}\x{23e9}-\x{23f3}\x{23f8}-\x{23fa}\x{2614}-\x{2615}\x{2618}\x{2672}-\x{267f}\x{2692}-\x{2697}\x{2699}\x{269b}-\x{269c}\x{26a0}-\x{26ab}\x{26b0}-\x{26b1}\x{26bd}-\x{26be}\x{26c4}-\x{26c5}\x{26c7}-\x{26c8}\x{26ce}-\x{26cf}\x{26d1}\x{26d3}-\x{26d4}\x{26dd}\x{26e3}\x{26e9}-\x{26ea}\x{26f0}-\x{26f5}\x{26f7}-\x{26fa}\x{26fd}\x{2705}\x{270a}-\x{270b}\x{2728}\x{274c}\x{274e}\x{2753}-\x{2755}\x{2757}\x{2795}-\x{2797}\x{27b0}\x{27bf}\x{2b05}\x{2b12}-\x{2b19}\x{2b1b}-\x{2b1c}\x{2b50}-\x{2b52}\x{2b55}\x{303d}\x{3244}-\x{3247}\x{e008}-\x{e009}\x{f8ff}\x{fe0e}-\x{fe0f}\x{1f000}-\x{1f02b}\x{1f0cf}\x{1f170}-\x{1f19a}\x{1f1e6}-\x{1f1ff}\x{1f201}-\x{1f202}\x{1f210}-\x{1f23a}\x{1f250}-\x{1f251}\x{1f300}-\x{1f321}\x{1f324}-\x{1f393}\x{1f396}-\x{1f397}\x{1f399}-\x{1f39b}\x{1f39e}-\x{1f3f0}\x{1f3f3}-\x{1f3f5}\x{1f3f7}-\x{1f4fd}\x{1f4ff}-\x{1f53d}\x{1f549}-\x{1f54e}\x{1f550}-\x{1f567}\x{1f56f}-\x{1f570}\x{1f573}-\x{1f57a}\x{1f587}\x{1f58a}-\x{1f58d}\x{1f590}\x{1f594}-\x{1f596}\x{1f5a4}-\x{1f5a5}\x{1f5a8}\x{1f5b1}-\x{1f5b2}\x{1f5bc}\x{1f5c2}-\x{1f5c4}\x{1f5d1}-\x{1f5d3}\x{1f5dc}-\x{1f5de}\x{1f5e1}\x{1f5e3}\x{1f5e8}\x{1f5ef}\x{1f5f3}\x{1f5fa}-\x{1f64f}\x{1f680}-\x{1f6c5}\x{1f6cb}-\x{1f6d2}\x{1f6d5}\x{1f6e0}-\x{1f6e5}\x{1f6e9}\x{1f6eb}-\x{1f6ec}\x{1f6f0}\x{1f6f3}-\x{1f6fa}\x{1f7e0}-\x{1f7eb}\x{1f90d}-\x{1f93a}\x{1f93c}-\x{1f945}\x{1f947}-\x{1f971}\x{1f973}-\x{1f976}\x{1f97a}-\x{1f9a2}\x{1f9a5}-\x{1f9aa}\x{1f9ae}-\x{1f9ca}\x{1f9cd}-\x{1f9ff}\x{1fa70}-\x{1fa73}\x{1fa78}-\x{1fa7a}\x{1fa80}-\x{1fa82}\x{1fa90}-\x{1fa95}]+)',
                'Segoe UI Symbol' => '([\x{c}\x{202a}-\x{202e}\x{2047}-\x{2048}\x{204a}-\x{2056}\x{2058}-\x{205d}\x{2064}\x{206a}-\x{206f}\x{20e2}\x{20e4}\x{20e7}\x{20eb}-\x{20f0}\x{23d0}-\x{23db}\x{23e1}-\x{23e8}\x{23f4}-\x{23f7}\x{2425}-\x{2426}\x{24f5}-\x{24fe}\x{2596}-\x{259f}\x{2616}-\x{2617}\x{2619}\x{2670}-\x{2671}\x{2680}-\x{2691}\x{2698}\x{269a}\x{269d}-\x{269f}\x{26ac}-\x{26af}\x{26b2}-\x{26bc}\x{26bf}-\x{26c3}\x{26c6}\x{26c9}-\x{26cd}\x{26d0}\x{26d2}\x{26d5}-\x{26dc}\x{26de}-\x{26e2}\x{26e4}-\x{26e8}\x{26eb}-\x{26ef}\x{26f6}\x{26fb}-\x{26fc}\x{26fe}-\x{2700}\x{275f}-\x{2760}\x{2768}-\x{2775}\x{2800}-\x{28ff}\x{2b00}-\x{2b03}\x{2b08}-\x{2b0b}\x{2b0e}-\x{2b11}\x{2b1d}-\x{2b4f}\x{2b53}-\x{2b54}\x{2b56}-\x{2b73}\x{2b76}-\x{2b95}\x{2b98}-\x{2bb9}\x{2bbd}-\x{2bc8}\x{2bca}-\x{2bd1}\x{2bec}-\x{2bef}\x{2c80}-\x{2cf3}\x{2cf9}-\x{2cff}\x{2e00}-\x{2e16}\x{2e18}-\x{2e31}\x{2e3f}\x{3248}-\x{324f}\x{4dc0}-\x{4dff}\x{e000}-\x{e007}\x{e00a}-\x{e019}\x{e052}\x{e070}-\x{e072}\x{e07d}-\x{e0b6}\x{e0b8}-\x{e0bd}\x{e0bf}-\x{e0fc}\x{e100}-\x{e156}\x{e158}-\x{e17f}\x{e181}-\x{e1fe}\x{e200}-\x{e236}\x{e239}-\x{e23a}\x{e241}-\x{e26c}\x{e26e}-\x{e271}\x{e27e}-\x{e292}\x{e294}-\x{e299}\x{e29b}-\x{e2fe}\x{f006}-\x{f009}\x{f010}-\x{f013}\x{fe24}-\x{fe26}\x{10400}-\x{1044f}\x{1d100}-\x{1d126}\x{1d129}-\x{1d1e8}\x{1d300}-\x{1d356}\x{1d4a7}-\x{1d4a8}\x{1f030}-\x{1f093}\x{1f0a0}-\x{1f0ae}\x{1f0b1}-\x{1f0bf}\x{1f0c1}-\x{1f0ce}\x{1f0d1}-\x{1f0f5}\x{1f100}-\x{1f10c}\x{1f110}-\x{1f12e}\x{1f130}-\x{1f169}\x{1f200}\x{1f240}-\x{1f248}\x{1f322}-\x{1f323}\x{1f394}-\x{1f395}\x{1f398}\x{1f39c}-\x{1f39d}\x{1f3f1}-\x{1f3f2}\x{1f3f6}\x{1f4fe}\x{1f53e}-\x{1f548}\x{1f568}-\x{1f56e}\x{1f571}-\x{1f572}\x{1f57b}-\x{1f586}\x{1f588}-\x{1f589}\x{1f58e}-\x{1f58f}\x{1f591}-\x{1f593}\x{1f597}-\x{1f5a3}\x{1f5a6}-\x{1f5a7}\x{1f5a9}-\x{1f5b0}\x{1f5b3}-\x{1f5bb}\x{1f5bd}-\x{1f5c1}\x{1f5c5}-\x{1f5d0}\x{1f5d4}-\x{1f5db}\x{1f5df}-\x{1f5e0}\x{1f5e2}\x{1f5e4}-\x{1f5e7}\x{1f5e9}-\x{1f5ee}\x{1f5f0}-\x{1f5f2}\x{1f5f4}-\x{1f5f9}\x{1f650}-\x{1f67f}\x{1f6c6}-\x{1f6ca}\x{1f6e6}-\x{1f6e8}\x{1f6ea}\x{1f6f1}-\x{1f6f2}\x{1f700}-\x{1f773}\x{1f780}-\x{1f785}\x{1f787}\x{1f789}\x{1f78c}-\x{1f78d}\x{1f791}-\x{1f798}\x{1f79a}-\x{1f79e}\x{1f7a0}-\x{1f7c0}\x{1f7c2}-\x{1f7c4}\x{1f7c6}-\x{1f7ca}\x{1f7cc}-\x{1f7ce}\x{1f7d0}\x{1f7d2}\x{1f7d4}\x{1f800}-\x{1f80b}\x{1f810}-\x{1f847}\x{1f850}-\x{1f859}\x{1f860}-\x{1f887}\x{1f890}-\x{1f8ad}\x{1f93b}]+)',
                'Nirmala UI' => '([\x{900}\x{904}\x{93a}-\x{93b}\x{94e}-\x{94f}\x{955}-\x{957}\x{971}-\x{980}\x{9bd}\x{9ce}\x{9fb}-\x{9fe}\x{a01}\x{a03}\x{a51}\x{a75}-\x{a76}\x{a8c}\x{ae1}-\x{ae3}\x{af0}-\x{af1}\x{af9}-\x{aff}\x{b35}\x{b44}\x{b62}-\x{b63}\x{b71}-\x{b77}\x{bb6}\x{bd0}\x{be6}\x{bf3}-\x{bfa}\x{c00}\x{c04}\x{c34}\x{c3d}\x{c58}-\x{c5a}\x{c62}-\x{c63}\x{c78}-\x{c81}\x{c84}\x{cbc}-\x{cbd}\x{ce2}-\x{ce3}\x{cf1}-\x{cf2}\x{d00}-\x{d01}\x{d29}\x{d3a}-\x{d3d}\x{d44}\x{d4e}-\x{d4f}\x{d54}-\x{d56}\x{d58}-\x{d5f}\x{d62}-\x{d63}\x{d70}-\x{d7f}\x{d82}-\x{d83}\x{d85}-\x{d96}\x{d9a}-\x{db1}\x{db3}-\x{dbb}\x{dbd}\x{dc0}-\x{dc6}\x{dca}\x{dcf}-\x{dd4}\x{dd6}\x{dd8}-\x{ddf}\x{de6}-\x{def}\x{df2}-\x{df4}\x{fd5}-\x{fd8}\x{1c50}-\x{1c7f}\x{1cd0}-\x{1cf9}\x{a830}-\x{a839}\x{a8e0}-\x{a8ff}\x{abc0}-\x{abed}\x{abf0}-\x{abf9}\x{110d0}-\x{110e8}\x{110f0}-\x{110f9}\x{11100}-\x{11134}\x{11136}-\x{11146}\x{111e1}-\x{111f4}]+)',
            ];

            $pregStr = '/'.implode('|',$pregUnicode).'/u';
            $string = preg_replace_callback($pregStr, function($match){
                foreach($match as $idx => $v) {
                    if($idx === 0 || $v === '') {
                        continue;
                    }

                    return '{%{'.$idx.'}%}'.$match[0].'{%{/'.$idx.'}%}';
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