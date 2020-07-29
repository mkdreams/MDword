<?php
namespace MDword\Common;

use MDword\Read\Word;
use MDword\Edit\Part\Rels;

class PartBase
{
    /**
     * @var \DOMDocument
     */
    protected $DOMDocument;
    
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
    
    public function hasAttr($item,$name,$ns='w') {
        return $item->hasAttributeNS($this->xmlns[$ns],$name);
    }
    
    public function removeAttr($item,$name) {
        $item->removeAttribute($name);
    }
    
    
    public function __get($name) {
        return $this->$name;
    }
    
    protected function markDelete($item) {
        $item->setAttribute('md',(++$this->id));
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
    
    public function createNodeByXml($xmlname)
    {
        $filename = MDWORD_SRC_DIRECTORY.'/XmlTemple/'.$xmlname.'.xml';
        if(is_file($filename)) {
            $xml = file_get_contents($filename);
        }else{
            $xml = $xmlname;
        }
        
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">'
            .$xml.'</w:document>';
        
        $dom = new \DOMDocument();
        $dom->loadXML($xml);
        
        return $this->DOMDocument->importNode($dom->documentElement->firstChild,true);
    }
    
    protected function initChartRels($relArr) {
        $partInfo = pathinfo($relArr['PartName']);
        $this->rels = new Rels($this->word, $relArr['dom']);
        $this->rels->partName = $relArr['relName'];
        $this->rels->partInfo = $partInfo;
        
        $this->word->parts[19][] = ['PartName'=>$this->rels->partName,'DOMElement'=>$this->rels->DOMDocument];
    }
    
    
    protected function initCommentRange() {
        $this->treeToList($this->DOMDocument->documentElement);
        
        $commentRangeStartItems = $this->DOMDocument->getElementsByTagName('commentRangeStart');
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
            $tempBlocks[$name][] = array_map(function($trace) use ($name,$size) {
                $this->domIdxToName[$trace->idxBegin][] = [$size,$name];
                return $trace->idxBegin;
            }, $traces);
        }
        
        return $tempBlocks;
    }
    
    protected function treeToList($node) {
        static $index = 0;
        
        if(is_null($node)) {
            return $index;
        }
        $node->idxBegin = $index;
        $this->domList[$index++] = $node;
        if(($node->hasChildNodes())) {
            foreach($node->childNodes as $childNode) {
                if($childNode->nodeType !== 3) {
                    $this->treeToList($childNode);
                }
            }
            
        }
        
        $node->idxEnd = $index-1;
    }
    
    protected function getCommentRangeEnd($parentNode,$id) {
        $commentRangeEndItems = $parentNode->getElementsByTagName('commentRangeEnd');
        
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
        
        $traces[] = $startParentNode;
        
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
                foreach ($childNodes as $childNode) {
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
                    $traces[] = $nextSibling;
                }
                
                break;
            }else{
                $traces[] = $nextSibling;
            }
            
            $nextSibling = $nextSibling->nextSibling;
            $nextNodeCount++;
        }
        
        foreach($traces as $key => $trace) {
            $tagName = $trace->localName;
            if(isset($delTags[$tagName])) {
                unset($traces[$key]);
                continue;
            }
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
}