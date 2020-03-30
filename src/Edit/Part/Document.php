<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Document extends PartBase
{
    public function __construct(\DOMDocument $DOMDocument) {
        parent::__construct();
        
        $this->DOMDocument = $DOMDocument;
    }
    
    public function setValue($name,$value) {
        $blocks = $this->getBlocks();
//         var_dump($blocks);exit;
        if(isset($blocks[$name])) {
            foreach($blocks[$name] as $block) {
                $this->update($block,$value);
//                 $this->deleteBlock($block);
            }
            
        }
        
//         echo $this->DOMDocument->saveXML();exit;
//         var_dump($name,$value,$this->DOMDocument);exit;
    }
    
    private function update($block,$value) {
        $beginNode = $block[0]['node'];
        $preStr = $block[0]['preStr'];
        
        $endNode = $block[1]['node'];
        $endStr = $block[1]['endStr'];
        
        $middleNodes = $block[2];
        
        
        //value
        $copy = clone $beginNode;
        $copy->getElementsByTagName('t')->item(0)->nodeValue= $preStr.$value.$endStr;
        $beginNode->parentNode->insertBefore($copy,$beginNode);
        
        $parentNode = $beginNode->parentNode;
        
        $parentNode->removeChild($beginNode);
        if($beginNode !== $endNode) {
            foreach ($middleNodes as $middleNode) {
                $parentNode->removeChild($middleNode);
            }
        }
        $parentNode->removeChild($endNode);
    }
    
    private function deleteBlock($block) {
        $beginNode = $block[0]['node'];
        $endNode = $block[1]['node'];
        
//         var_dump($beginNode,$endNode);exit;
        $parentNode = $beginNode->parentNode;
        
        
        
        if(!is_null($beginNode)) {
            $parentNode->removeChild($beginNode);
        }
        if(!is_null($endNode)) {
            $parentNode->removeChild($endNode);
        }
        
        
        
    }
    
    private function getBlocks() {
        $items = $this->DOMDocument->getElementsByTagName('r');
        
        $preIsDollar = 0;
        $beginNode = null;
        $endNode = null;
        $blockName = '';
        
        $blocks = [];
        $middleNodes = [];
        foreach ($items as $itemKey => $item) {
            $text = $item->getElementsByTagName('t')->item(0)->nodeValue;
            preg_match_all("/./u",$text,$textArr);
            if(!is_array($textArr)) {
                continue;
            }
            
            $endKey = -1;
            
            $endStr = null;
            
            foreach($textArr[0] as $key => $word) {
                if(!is_null($endStr)) {
                    $endStr .= $word;
                }
                if($word == '') {
                    continue;
                }
                
                if($word === '$') {
                    $preIsDollar = 1;
                    $preIsDollarItem = $item;
                    $preIsDollarKey = $key;
                }elseif($preIsDollar == 1 && $word === '{') {
                    $beginNode = ['node'=>$preIsDollarItem,'pos'=>$preIsDollarKey,
                        'preStr'=>implode('', array_slice($textArr[0],$endKey+1,$preIsDollarKey-$endKey-1))];
                    $preIsDollar = 0;
                    
                    if(!is_null($endStr)) {
                        $endStr = '';
                        unset($endStr);
                    }
                }elseif(!is_null($beginNode) && $word === '}') {
                    $endStr = '';
                    $endNode = ['node'=>$item,'pos'=>$key,
                        'endStr'=> &$endStr
                    ];
                    $endKey = $key;
                    $preStr = '';
                }elseif(!is_null($beginNode)){
                    $blockName .= $word;
                    $middleNodes[$itemKey] = $item;
                    $preIsDollar = 0;
                }else{
                    $preIsDollar = 0;
                }
                
                if(!is_null($beginNode) && !is_null($endNode)) {
                    $blocks[$blockName][] = [$beginNode,$endNode,$middleNodes];
                    $blockName = '';
                    $middleNodes = [];
                    $beginNode = null;
                    $endNode = null;
                }
            }
            
            unset($endStr);
        }
        
        return $blocks;
    }
}
