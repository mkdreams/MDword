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
        $parentNode = $targetNode->parentNode;
        $parentNode->insertBefore($copy,$targetNode);
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
    
    public function creatNode($mixed=[], \DOMElement $domElement = null)
    {
        foreach ($mixed as $tagName => $nodeInfo) {
            $node = $this->DOMDocument->createElementNS($this->xmlns['w'],$tagName);
            foreach($nodeInfo as $attrName => $attr) {
                if($attrName === 'childs') {
                    $this->creatNode($attr,$node);
                }elseif($attrName === 'text'){
                    $node->nodeValue = $attr;
                }else{
                    if(strpos($attrName, ':') > 0) {
                        $attrNameArr = explode(':', $attrName,2);
                        $node->setAttributeNS($this->xmlns[$attrNameArr[0]], $attrNameArr[1], $attr);
                    }else{
                        $node->setAttributeNS($this->xmlns['w'], $attrName, $attr);
                    }
                }
            }
            
            if(!is_null($domElement)) {
                $domElement->appendChild($node);
            }
        }
        
        return $node;
    }

    protected function initChartRels($relArr) {
        $partInfo = pathinfo($relArr['PartName']);
        $this->rels = new Rels($this->word, $relArr['dom']);
        $this->rels->partName = $relArr['relName'];
        $this->rels->partInfo = $partInfo;

        $this->word->parts[19][] = ['PartName'=>$this->rels->partName,'DOMElement'=>$this->rels->DOMDocument];
    }
    
    
}
