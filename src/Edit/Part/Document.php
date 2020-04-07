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
        $blocks = $this->getBlocks($name);
        foreach($blocks as $block) {
            $this->update($block,$value,$type);
        }
        
//         echo $this->DOMDocument->saveXML();exit;
//         var_dump($name,$value,$this->DOMDocument);exit;
    }
    
    public function clone($name,$value,$type="clone") {
        $blocks = $this->getBlocks();
//         var_dump($this->commentsblocks,$blocks);exit;
        if(isset($blocks[$name])) {
            foreach($blocks[$name] as $block) {
                $this->update($block,$value,$type);
            }
        }
    }
    
    private function update($block,$value,$type) {
        $beginNode = $block[0];
        
        $endNode = $block[1];
        
        $traces = $block[2];
        $parentNodeCount = $traces['parentNodeCount'];
        $nextNodeCount = $traces['nextNodeCount'];
//         var_dump($beginNode,$endNode,$middleNodes);exit;
        $targetNode = null;
        $deleteNodes = [];
        $middleNodes = [];
        switch ($type) {
            case 'text':
                $middleNodes = [];
                $parentNode = $beginNode;
                for($i=0;$i<=$parentNodeCount;$i++) {
                    $nextSibling = $parentNode;
                    while($nextSibling = $nextSibling->nextSibling) {
                        if(is_null($targetNode)) {
                            if($nextSibling->localName == 'r') {//is target
                                $targetNode = $nextSibling;
                            }else{//sub find target
                                $rs = $nextSibling->getElementsByTagName('r');
                                if($rs->length > 0) {
                                    $targetNode = $rs->item(0);
                                }
                            }
                        }else{
                            $this->markDelete($nextSibling);
                        }
                    }
                    
                    if($i === $parentNodeCount) {//top parent
                        
                    }else{
                        $parentNode = $parentNode->parentNode;
                    }
                    
                }
                
//                 var_dump($this->DOMDocument->getElementById('dm#2'));exit;
                $this->deleteMarked();
                echo $this->DOMDocument->saveXML();exit;
//                 var_dump($traces);exit;
                foreach($middleNodes as $middleNode) {
                    if($middleNode->localName == 'r') {
                        $deleteNodes[] = $middleNode;
                        if(is_null($node)) {
                            $node = $middleNode;
                        }
                    }
                }
                
//                 if(!is_null($node)) {
//                     $copy = clone $node;
//                     $copy->getElementsByTagName('t')->item(0)->nodeValue= $value;
//                     $parentNode = $node->parentNode;
//                     $parentNode->insertBefore($copy,$node);
//                 }
                
//                 //remove comments and middle node
//                 foreach($deleteNodes as $deleteNode) {
//                     $deleteNode->parentNode->removeChild($deleteNode);
//                 }
//                 $beginNode->parentNode->removeChild($beginNode);
//                 $endNode->parentNode->removeChild($endNode);
                
                break;
            case 'image':
//                 foreach($middleNodes as $middleNode) {
//                     $pictNode = $middleNode->getElementsByTagName('pict')->item(0);
//                     if($pictNode) {
//                         $deleteNodes[] = $middleNode;
//                         if(is_null($node)) {
//                             $node = $pictNode;
//                         }
//                     }
//                 }
                
//                 if(!is_null($node)) {
//                     $rid = $this->getAttr($node->getElementsByTagName('imagedata')->item(0), 'id', 'r');
//                     $this->updateRef($rid,$value);
//                 }
                break;
            case 'clone':
                
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
    
    private function getBlocks($name) {
        $commentRangeStartItems = $this->DOMDocument->getElementsByTagName('commentRangeStart');
        
        $blocks = [];
        foreach($commentRangeStartItems as $commentRangeStartItem) {
            $id = $this->getAttr($commentRangeStartItem, 'id');
            if($this->commentsblocks[$id] !== $name) {
                continue;
            }
            $commentRangeEndItem = $this->getCommentRangeEnd($this->DOMDocument,$id);
            
            $trace = $this->getRangeTrace($id,$commentRangeStartItem, $commentRangeEndItem);
//             exit;
            /*
            $middleNodes = [];
            $nextSibling = $commentRangeStartItem->nextSibling;
            
            while($nextSibling !== null && $nextSibling !== $commentRangeEndItem) {
                $middleNodes[] = $nextSibling;
                $nextSibling = $nextSibling->nextSibling;
            }
            
            //父级查找
            if($nextSibling === null) {
                $parentNode = $commentRangeStartItem;
                while($parentNode = $parentNode->parentNode) {
                    $commentRangeEndItem = $this->getCommentRangeEnd($parentNode,$id);
                    if(is_null($commentRangeEndItem)) {
                        $preParentNode = $parentNode;
                    }else{
                        break;
                    }
                }
                
                //middles
                $nextSibling = $preParentNode->nextSibling;
                while($this->getCommentRangeEnd($nextSibling,$id) !== null) {
                    $middleNodes[] = $nextSibling;
                    $nextSibling = $nextSibling->nextSibling;
                }
                
                
//                 var_dump($nextSibling,$preParentNode,$commentRangeEndItem);
//                 $parentNode->getElementsByTagName('commentRangeStart');
                
//                 var_dump($this->commentsblocks,$id = $this->getAttr($commentRangeStartItem, 'id'));exit;
            }
            */
            $blocks[] = [$commentRangeStartItem,$commentRangeEndItem,$trace];
        }
        
//         var_dump($blocks['table']);exit;
        
        return $blocks;
    }
    
    private function getRangeTrace($id,$commentRangeStartItem,$commentRangeEndItem) {
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
        
        $nextNodeCount = 0;
        if($parentNodeCount === 0) {
            $nextSibling = $commentRangeStartItem->nextSibling;
            $nextNodeCount++;
            while($nextSibling !== null && $nextSibling !== $commentRangeEndItem) {
                $nextSibling = $nextSibling->nextSibling;
                $nextNodeCount++;
            }
            
            $endParentNode = $commentRangeEndItem;
        }else{
            $nextSibling = $startParentNode->nextSibling;
            $nextNodeCount++;
            while(!is_null($nextSibling) && $this->getCommentRangeEnd($nextSibling,$id) === null && $nextSibling !== $commentRangeEndItem) {
                $nextSibling = $nextSibling->nextSibling;
                $nextNodeCount++;
            }
            $endParentNode = $nextSibling;
//             var_dump($startParentNode,$endParentNode);
        }
        
        return ['parentNodeCount'=>$parentNodeCount,'nextNodeCount'=>$nextNodeCount];
    }
    
    private function getCommentRangeEnd($parentNode,$id) {
        $commentRangeEndItems = $parentNode->getElementsByTagName('commentRangeEnd');
        
        foreach($commentRangeEndItems as $commentRangeEndItem) {
            $eid = $this->getAttr($commentRangeEndItem, 'id');
            
            if($id === $eid) {
                return $commentRangeEndItem;
            }
        }
        
        return null;
    }
}
