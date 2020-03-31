<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Document extends PartBase
{
    public $commentsblocks;
    public function __construct(\DOMDocument $DOMDocument,$blocks = []) {
        parent::__construct();
        
        $this->DOMDocument = $DOMDocument;
        $this->commentsblocks = $blocks;
        
        $this->initNameSpaces();
    }
    
    public function setValue($name,$value,$type='text') {
        $blocks = $this->getBlocks();
//         var_dump($name,$blocks);exit;
        if(isset($blocks[$name])) {
            foreach($blocks[$name] as $block) {
                $this->update($block,$value,$type);
            }
        }
        
//         echo $this->DOMDocument->saveXML();exit;
//         var_dump($name,$value,$this->DOMDocument);exit;
    }
    
    private function update($block,$value,$type) {
//         var_dump($block,$value);exit;
        
        
        $beginNode = $block[0];
        
        $endNode = $block[1];
        
        $middleNodes = $block[2];
//         var_dump($beginNode,$endNode,$middleNodes);exit;
        $node = null;
        $deleteNodes = [];
        switch ($type) {
            case 'text':
                foreach($middleNodes as $middleNode) {
                    if($middleNode->localName == 'r') {
                        $deleteNodes[] = $middleNode;
                        if(is_null($node)) {
                            $node = $middleNode;
                        }
                    }
                }
                
                if(!is_null($node)) {
                    $copy = clone $node;
                    $copy->getElementsByTagName('t')->item(0)->nodeValue= $value;
                    $parentNode = $node->parentNode;
                    $parentNode->insertBefore($copy,$node);
                }
                
                //remove comments and middle node
                foreach($deleteNodes as $deleteNode) {
                    $deleteNode->parentNode->removeChild($deleteNode);
                }
                $beginNode->parentNode->removeChild($beginNode);
                $endNode->parentNode->removeChild($endNode);
                
                break;
            default:
                break;
        }
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
//         var_dump($this->commentsblocks);exit;
        $commentRangeStartItems = $this->DOMDocument->getElementsByTagName('commentRangeStart');
        $commentRangeEndItems = $this->DOMDocument->getElementsByTagName('commentRangeEnd');
        
        $blocks = [];
        foreach($commentRangeStartItems as $key => $commentRangeStartItem) {
            $middleNodes = [];
            $nextSibling = $commentRangeStartItem->nextSibling;
            $commentRangeEndItem = $commentRangeEndItems->item($key);
//             var_dump($commentRangeEndItem);exit;
            while($nextSibling !== null && $nextSibling !== $commentRangeEndItem) {
                $middleNodes[] = $nextSibling;
                $nextSibling = $nextSibling->nextSibling;
            }
            $id = $this->getAttr($commentRangeStartItem, 'id');
            $blocks[$this->commentsblocks[$id]][] = [$commentRangeStartItem,$commentRangeEndItem,$middleNodes];
//             var_dump($id,$name);exit;
//             $blocks['']
//             var_dump($commentRangeStartItem,$commentRangeStartItem->nextSibling,$commentRangeEndItems->item($key));exit;
        }
        
//         var_dump($blocks);exit;
        
        return $blocks;
        var_dump($commentRangeStartItems,$commentRangeEndItems);exit;
        
        $preIsDollar = 0;
        $beginNode = null;
        $endNode = null;
        $blockName = '';
        
        $blocks = [];
        $middleNodes = [];
        foreach ($items as $itemKey => $item) {
            $t = $item->getElementsByTagName('t');
            if($t->length !== 1) {
//                 echo "skip one r\r\n";
                continue;
            }
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
