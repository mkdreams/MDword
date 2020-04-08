<?php
namespace MDword\Common;

use MDword\Read\Word;

class PartBase
{
    protected $DOMDocument;
    /**
     * @var \DOMDocument
     */
    protected $refDOMDocument = null;
    
    protected $rootPath;
    
    private $id = 0;
    /**
     * @var Word
     */
    public $word;
    
    public $partName = null;
    
    public $partNameRel = null;
    
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
        return $item->getAttributeNS($this->xmlns[$ns],$name);
    }
    
    public function setAttr($item,$name,$value,$ns='w') {
        return $item->setAttributeNS($this->xmlns[$ns],$name,$value);
    }
    
    
    public function __get($name) {
        return $this->$name;
    }
    
    protected function markDelete($item) {
        $item->setAttribute('md',(++$this->id));
    }
    
    public function deleteMarked() {
        $xpath = new \DOMXPath($this->DOMDocument);
        $context = $this->DOMDocument->documentElement;
        foreach( $xpath->query('//*[@md]', $context) as $node ) {
            $node->parentNode->removeChild($node);
        }
    }
}
