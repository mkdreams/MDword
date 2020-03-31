<?php
namespace MDword\Common;

class PartBase
{
    protected $DOMDocument;
    
    protected $rootPath;
    
    protected $xmlns = [];
    
    public function __construct() {
        $this->rootPath = dirname(__DIR__);
    }
    
    public function initNameSpaces() {
        $context = $this->DOMDocument->documentElement;
        $xpath = new \DOMXPath($this->DOMDocument);
        foreach( $xpath->query('namespace::*', $context) as $node ) {
            $this->xmlns[$node->localName] = $node->nodeValue;
        }
    }
    
    public function getAttr($item,$name,$ns='w') {
        return $item->getAttributeNS($this->xmlns[$ns],'id');
    }
    
    public function __get($name) {
        return $this->$name;
    }
}
