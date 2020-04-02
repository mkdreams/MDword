<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Document extends PartBase
{
    public $commentsblocks;
    public function __construct($word,\DOMDocument $DOMDocument,$blocks = []) {
        parent::__construct($word);
        
        $this->DOMDocument = $DOMDocument;
        $this->commentsblocks = $blocks;
        
        $this->initNameSpaces();
    }
    
    public function setValue($name,$value,$type='text') {
        $blocks = $this->getBlocks();
        if(isset($blocks[$name])) {
            foreach($blocks[$name] as $block) {
                $this->update($block,$value,$type);
            }
        }
        
//         echo $this->DOMDocument->saveXML();exit;
//         var_dump($name,$value,$this->DOMDocument);exit;
    }
    
    private function update($block,$value,$type) {
//         var_dump($block,$value,$type);exit;
        
        
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
            case 'image':
                foreach($middleNodes as $middleNode) {
                    $pictNode = $middleNode->getElementsByTagName('pict')->item(0);
                    if($pictNode) {
                        $deleteNodes[] = $middleNode;
                        if(is_null($node)) {
                            $node = $pictNode;
                        }
                    }
                }
                
                if(!is_null($node)) {
                    $rid = $this->getAttr($node->getElementsByTagName('imagedata')->item(0), 'id', 'r');
                    $this->updateRef($rid,$value);
//                     $node->getElementsByTagName('imagedata')->item(0)->nodeValue= $value;
//                     $parentNode = $node->parentNode;
//                     $parentNode->insertBefore($copy,$node);
                }
                
                //remove comments and middle node
//                 foreach($deleteNodes as $deleteNode) {
//                     $deleteNode->parentNode->removeChild($deleteNode);
//                 }
//                 $beginNode->parentNode->removeChild($beginNode);
//                 $endNode->parentNode->removeChild($endNode);
                
                break;
            default:
                break;
        }
    }
    
    private function updateRef($rid,$file) {
//         var_dump(is_file($file));exit;
        
        $partInfo = pathinfo($this->partName);
        if(is_null($this->refDOMDocument)) {
            $this->partNameRel = $partInfo['dirname'].'/_rels/'.$partInfo['basename'].'.rels';
            $this->refDOMDocument = $this->word->getXmlDom($this->partNameRel);
            $this->word->parts[19][] = ['PartName'=>$this->partNameRel,'DOMElement'=>$this->refDOMDocument];
        }
        
        
        $Relationships = $this->refDOMDocument->getElementsByTagName('Relationship');
        $length = $Relationships->length;
        foreach ($Relationships as $Relationship) {
            if($Relationship->getAttribute('Id') === $rid) {
                $oldValue = $partInfo['dirname'].'/'.$Relationship->getAttribute('Target');
                $target = 'media/image'.++$length.'.png';
                $Relationship->setAttribute('Target',$target);
                $target = $partInfo['dirname'].'/'.$target;
                $this->word->zip->addFromString($target, file_get_contents($file));
                $this->word->zip->deleteName($oldValue);
            }
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
        static $blocks = null;
        
        if(!is_null($blocks)) {
            return $blocks;
        }
        
        $commentRangeStartItems = $this->DOMDocument->getElementsByTagName('commentRangeStart');
        $commentRangeEndItems = $this->DOMDocument->getElementsByTagName('commentRangeEnd');
        
        $blocks = [];
        foreach($commentRangeStartItems as $key => $commentRangeStartItem) {
            $middleNodes = [];
            $nextSibling = $commentRangeStartItem->nextSibling;
            $commentRangeEndItem = $commentRangeEndItems->item($key);
            
            while($nextSibling !== null && $nextSibling !== $commentRangeEndItem) {
                $middleNodes[] = $nextSibling;
                $nextSibling = $nextSibling->nextSibling;
            }
            $id = $this->getAttr($commentRangeStartItem, 'id');
            $blocks[$this->commentsblocks[$id]][] = [$commentRangeStartItem,$commentRangeEndItem,$middleNodes];
        }
        
        
        return $blocks;
    }
}
