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
    
    protected function htmlspecialcharsBase($string) {
        $string = $this->my_html_entity_decode($string);
        $string = $this->filterUtf8($this->filterSpecailCodeForWord($string));

        $string = $this->controlCharacterPHP2OOXML($string);
        $string = preg_replace('/['.chr(240).'-'.chr(247).']/', '{%{}%}'."$0", $string);

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