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
    
    protected function removeMarkDelete($item) {
        $item->removeAttribute('md');
    }
    
    public function deleteMarked() {
        $xpath = new \DOMXPath($this->DOMDocument);
        $context = $this->DOMDocument->documentElement;
        foreach( $xpath->query('//*[@md]', $context) as $node ) {
            $node->parentNode->removeChild($node);
        }
    }
    
    protected function getRels($xmlType=19) {
        $partInfo = pathinfo($this->partName);
        $partNameRel = $partInfo['dirname'].'/_rels/'.$partInfo['basename'].'.rels';
        $this->rels = new Rels($this->word, $this->word->getXmlDom($partNameRel));
        $this->rels->partName = $partNameRel;
        $this->rels->partInfo = $partInfo;
        $this->word->parts[$xmlType][] = ['PartName'=>$this->rels->partName,'DOMElement'=>$this->rels->DOMDocument];
    }
}
