<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Comments extends PartBase
{
    public $blocks = [];
    public function __construct($word,\DOMDocument $DOMDocument) {
        parent::__construct($word);
        
        $this->DOMDocument = $DOMDocument;
        $this->initNameSpaces();
        $this->blocks = $this->getBlocks();
    }
    
    private function deleteBlock($block) {
        $beginNode = $block[0]['node'];
        $endNode = $block[1]['node'];
        
        $parentNode = $beginNode->parentNode;
        
        if(!is_null($beginNode)) {
            $parentNode->removeChild($beginNode);
        }
        if(!is_null($endNode)) {
            $parentNode->removeChild($endNode);
        }
    }
    
    private function getBlocks() {
        $items = $this->DOMDocument->getElementsByTagName('comment');
        
        $preIsDollar = 0;
        $beginNode = null;
        $endNode = null;
        $blockName = '';
        
        $blocks = [];
        foreach ($items as $item) {
            $text = $item->nodeValue;
            $textArr = [];
            preg_match_all("/./u",$text,$textArr);
            if(!is_array($textArr)) {
                continue;
            }
            
            foreach($textArr[0] as $word) {
                if($word == '') {
                    continue;
                }
                
                if($word === '$') {
                    $preIsDollar = 1;
                }elseif($preIsDollar == 1 && $word === '{') {
                    $beginNode = $item;
                    $preIsDollar = 0;
                }elseif(!is_null($beginNode) && $word === '}') {
                    $endNode = $item;
                }elseif(!is_null($beginNode)){
                    $blockName .= $word;
                    $preIsDollar = 0;
                }else{
                    $preIsDollar = 0;
                }
                
                if(!is_null($beginNode) && !is_null($endNode)) {
                    $blocks[$item->getAttributeNS($this->xmlns['w'],'id')] = $blockName;
                    $blockName = '';
                    $beginNode = null;
                    $endNode = null;
                }
            }
        }
        
        return $blocks;
    }
}
